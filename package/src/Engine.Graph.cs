// =============================================================================
//  Mailbox Transfer - engine, part 1: Microsoft Graph transport
// -----------------------------------------------------------------------------
//  Author  : Nicolas Fabert
//  Version : 0.1.0 (phase 0)
//  The src\Engine.*.cs files are compiled together by MailboxTransfer.psm1
//  (Initialize-MbtEngine). Engine.Auth.cs is compiled separately (needs MSAL).
//
//  What this file does, and nothing more:
//    - one HttpClient, a bearer token read from a Func<string> at each request;
//    - retry of 429 / 503 / 504 (and network errors) honouring Retry-After;
//    - @odata.nextLink pagination, followed as returned when it is https on graph.microsoft.com
//      (the rule of the Microsoft Graph SDKs);
//    - 308 redirects (archive mailbox with auto-expanding folders) followed by hand; MailboxRedirected tells which mailbox they lead to.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer
{
    /// <summary>Error answered by Microsoft Graph (or by the Outlook import endpoint).</summary>
    public sealed class GraphException : Exception
    {
        public int StatusCode { get; }
        public string ErrorCode { get; }
        public string Url { get; }

        public GraphException(int statusCode, string errorCode, string message, string url)
            : base(message)
        {
            StatusCode = statusCode;
            ErrorCode = errorCode;
            Url = url;
        }
    }

    /// <summary>
    /// A request body that knows how far it was handed to the transport. When sending fails and <see cref="WriteFailed"/> is set, the write
    /// of the body itself threw before the last byte was handed over; when <see cref="AnyByteWritten"/> is false not a single byte was
    /// handed over (connection, TLS or DNS failure, timeout before the body was asked for). Either way the service cannot have acted on
    /// a request it did not receive whole.
    /// </summary>
    public interface IUploadProgress
    {
        long BytesWritten { get; }
        long TotalBytes { get; }
        bool WriteFailed { get; }
        /// <summary>True as soon as one piece of the body was handed to the transport (never reset, even if the transport asks for the body again).</summary>
        bool AnyByteWritten { get; }
    }

    /// <summary>A 429 or 503 received: when, which, and the time the client waits (Retry-After when the service gave one).</summary>
    public sealed class ThrottleEvent
    {
        public DateTime Utc;
        public int Status;
        public double WaitSeconds;
        public bool HadRetryAfter;
        public string Method;
        public string Host;
        public override string ToString()
        {
            return Utc.ToString("HH:mm:ss") + " UTC  " + Status + " on " + Method + " " + Host + ", " + (HadRetryAfter ? "Retry-After " : "waiting ") + WaitSeconds.ToString("0", System.Globalization.CultureInfo.InvariantCulture) + " s";
        }
    }

    public sealed class GraphOptions
    {
        public string BaseUrl = "https://graph.microsoft.com/v1.0";
        /// <summary>Total number of attempts of one request (first try included).</summary>
        public int MaxAttempts = 6;
        public double BaseBackoffSeconds = 2;
        public double MaxBackoffSeconds = 60;
        /// <summary>A Retry-After longer than this is shortened to it.</summary>
        public double MaxRetryAfterSeconds = 600;
        public bool UseJitter = true;
        /// <summary>Waiting time of one request until its response headers arrive.</summary>
        public TimeSpan RequestTimeout = TimeSpan.FromMinutes(15);
        /// <summary>
        /// Slowest upload the client accepts, in bytes per second: a request with a body may take <see cref="RequestTimeout"/> plus the time
        /// the body needs at this speed (a 15-minute limit alone cut off a large item on a slow link).
        /// </summary>
        public long MinUploadBytesPerSecond = 64 * 1024;
        /// <summary>A streamed response (exportItems, message $value) that delivers no byte for this long is abandoned.</summary>
        public TimeSpan ReadInactivityTimeout = TimeSpan.FromMinutes(2);
        /// <summary>Multiplies every waiting time (0 in the unit tests).</summary>
        public double DelayScale = 1.0;
    }

    public sealed class GraphClient : IDisposable
    {
        static readonly Random Rng = new Random();
        static readonly Regex AuthTokenParam = new Regex(@"authtoken=", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        readonly HttpClient _http;
        readonly Func<string> _token;
        readonly object _lock = new object();

        public GraphOptions Options { get; }
        /// <summary>Waiting times (before scaling) computed for the retries, for diagnostics and tests.</summary>
        public List<TimeSpan> Delays { get; } = new List<TimeSpan>();
        public int RequestCount { get; private set; }
        public int RetryCount { get; private set; }
        public int PagesRead { get; private set; }
        /// <summary>Answers 429 received (throttling), 503 and 504 received, over the whole life of the client.</summary>
        public int Throttled429 { get; private set; }
        public int Unavailable503 { get; private set; }
        public int Timeout504 { get; private set; }
        /// <summary>Requests wait here after a 429 / 503 with a delay, so that every worker obeys the Retry-After of one of them.</summary>
        DateTime _gateUntilUtc = DateTime.MinValue;
        /// <summary>Every 429 / 503 received (the last 2000), and an optional observer called for each one (the jobs lower their concurrency).</summary>
        public List<ThrottleEvent> ThrottleEvents { get; } = new List<ThrottleEvent>();
        public Action<ThrottleEvent> ThrottleHandler;
        /// <summary>Optional observer of the notable events of a request that are not throttling (an upload interrupted and sent again).</summary>
        public Action<string> Notice;
        /// <summary>
        /// Optional observer of a redirect to another mailbox (308 of an auto-expanding archive): called with the mailbox id the
        /// request was sent to, the mailbox id of the Location, and the Location without its query. It runs on the thread of the request.
        /// </summary>
        public Action<string, string, string> MailboxRedirected;

        public ThrottleEvent[] GetThrottleEvents() { lock (_lock) { return ThrottleEvents.ToArray(); } }

        public GraphClient(Func<string> tokenProvider, GraphOptions options = null, HttpMessageHandler handler = null)
        {
            _token = tokenProvider;
            Options = options ?? new GraphOptions();
            _http = new HttpClient(handler ?? new SocketsHttpHandler
            {
                AllowAutoRedirect = false,
                AutomaticDecompression = DecompressionMethods.All,
                PooledConnectionLifetime = TimeSpan.FromMinutes(10)
            });
            _http.Timeout = Timeout.InfiniteTimeSpan;   // the time limit is per request (Options.RequestTimeout)
        }

        public void Dispose() { _http.Dispose(); }

        public string Url(string path) { return Options.BaseUrl + path; }

        /// <summary>
        /// Removes the secret of a pre-authenticated URL (import URL) from a text meant for a log: the whole query string of a URL that is not
        /// on graph.microsoft.com or that carries an authtoken (whatever the other parameters, their order or their case).
        /// </summary>
        public static string Redact(string url)
        {
            if (url == null) return null;
            int q = url.IndexOf('?');
            if (q < 0) return url;
            return AuthTokenParam.IsMatch(url) || !IsGraphUrl(url) ? url.Substring(0, q) + "?***" : url;
        }

        /// <summary>Time a request with a body of this size may take: the request timeout plus the upload at <see cref="GraphOptions.MinUploadBytesPerSecond"/>.</summary>
        public TimeSpan TimeoutFor(long bodyBytes)
        {
            if (bodyBytes <= 0 || Options.MinUploadBytesPerSecond <= 0) return Options.RequestTimeout;
            return Options.RequestTimeout + TimeSpan.FromSeconds((double)bodyBytes / Options.MinUploadBytesPerSecond);
        }

        /// <summary>
        /// The body of a response read with HttpCompletionOption.ResponseHeadersRead, as a stream that obeys the cancellation token and gives
        /// up after <see cref="GraphOptions.ReadInactivityTimeout"/> without a byte (otherwise a connection that goes silent blocks forever).
        /// </summary>
        public Stream ReadBody(HttpResponseMessage resp, CancellationToken ct = default(CancellationToken))
        {
            Stream s = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult();
            return new GuardedReadStream(s, ct, Options.ReadInactivityTimeout);
        }

        /// <summary>The Microsoft Graph SDK rule: https on graph.microsoft.com, otherwise the link is not ours.</summary>
        public static bool IsGraphUrl(string url)
        {
            Uri u;
            return Uri.TryCreate(url, UriKind.Absolute, out u)
                && u.Scheme == Uri.UriSchemeHttps
                && string.Equals(u.Host, "graph.microsoft.com", StringComparison.OrdinalIgnoreCase);
        }

        // ---------------------------------------------------------------------------------------------
        //  Sending
        // ---------------------------------------------------------------------------------------------

        /// <summary>
        /// Sends a request with retry. <paramref name="content"/> builds a new body for every attempt.
        /// <paramref name="idempotent"/> false (an import creates an item): a 504 or a network error is not
        /// retried, because the item may already exist. 429 and 503 are always retried: the request was refused. The one exception is a
        /// body (<see cref="IUploadProgress"/>) whose write failed before its last byte: the service never received the whole request, so
        /// it is retried like a refusal, and when the attempts are used up the error is "UploadInterrupted" (nothing was created). The same holds for a body of
        /// which no byte was handed to the transport at all (the request never left: error code "NotSent"). A failure after the whole body was
        /// handed over, and any HTTP status received, are not touched.
        /// The caller disposes the response.
        /// </summary>
        public HttpResponseMessage Send(string method, string url, Func<HttpContent> content = null, bool authorize = true,
            bool idempotent = true, HttpCompletionOption completion = HttpCompletionOption.ResponseContentRead,
            CancellationToken ct = default(CancellationToken), Action<string> finalUrl = null, IDictionary<string, string> headers = null)
        {
            int redirects = 0;
            for (int attempt = 1; ; attempt++)
            {
                WaitGate(ct);
                lock (_lock) { RequestCount++; }
                HttpResponseMessage resp = null;
                Exception failure = null;
                HttpContent sentBody = null;
                try
                {
                    using (var request = new HttpRequestMessage(new HttpMethod(method), url))
                    using (var cts = CancellationTokenSource.CreateLinkedTokenSource(ct))
                    {
                        if (content != null) request.Content = sentBody = content();
                        if (headers != null) foreach (var h in headers) request.Headers.TryAddWithoutValidation(h.Key, h.Value);
                        if (authorize) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _token());
                        long bodyLength = 0;
                        if (request.Content != null) { try { bodyLength = request.Content.Headers.ContentLength ?? 0; } catch (Exception) { } }
                        cts.CancelAfter(TimeoutFor(bodyLength));
                        resp = _http.SendAsync(request, completion, cts.Token).GetAwaiter().GetResult();
                    }
                }
                catch (Exception ex) when (!ct.IsCancellationRequested &&
                    (ex is HttpRequestException || ex is IOException || ex is TaskCanceledException || ex is OperationCanceledException))
                {
                    failure = ex;
                }

                if (failure != null)
                {
                    var upload = sentBody as IUploadProgress;
                    if (upload != null && (upload.WriteFailed || !upload.AnyByteWritten))
                    {
                        bool notSent = !upload.WriteFailed;
                        string what = notSent
                            ? "request not sent (no byte of the body reached the transport), nothing created"
                            : "upload interrupted after " + upload.BytesWritten + " of " + upload.TotalBytes + " bytes, nothing created";
                        if (attempt >= Options.MaxAttempts)
                            throw new GraphException(0, notSent ? "NotSent" : "UploadInterrupted", method + " " + Redact(url) + " failed: " + what + " (" + attempt + " attempts): " + failure.Message, Redact(url));
                        var notice = Notice;
                        if (notice != null) { try { notice(what + ", retrying (attempt " + attempt + " of " + Options.MaxAttempts + ")"); } catch (Exception) { } }
                        Wait(null, attempt, ct: ct);
                        continue;
                    }
                    if (!idempotent || attempt >= Options.MaxAttempts)
                        throw new GraphException(0, "NetworkError", method + " " + Redact(url) + " failed: " + failure.Message, Redact(url));
                    Wait(null, attempt, ct: ct);
                    continue;
                }

                int status = (int)resp.StatusCode;
                if (status >= 200 && status < 300)
                {
                    if (finalUrl != null) finalUrl(url);
                    return resp;
                }

                if ((status == 301 || status == 302 || status == 307 || status == 308) && redirects < 3)
                {
                    string location = resp.Headers.Location == null ? null : resp.Headers.Location.ToString();
                    resp.Dispose();
                    if (!IsGraphUrl(location))
                        throw new GraphException(status, "UnexpectedRedirect", "Redirect to a URL that is not on graph.microsoft.com: " + Redact(location), Redact(url));
                    redirects++;
                    NoteRedirect(url, location);
                    url = location;
                    attempt--;
                    continue;
                }

                lock (_lock)
                {
                    if (status == 429) Throttled429++;
                    else if (status == 503) Unavailable503++;
                    else if (status == 504) Timeout504++;
                }
                bool retryable = status == 429 || status == 503 || (status == 504 && idempotent);
                if (retryable && attempt < Options.MaxAttempts)
                {
                    Wait(resp, attempt, status == 429 || status == 503, method, url, ct);
                    resp.Dispose();
                    continue;
                }

                string body = ReadErrorBody(resp);
                string code, message;
                ParseError(body, out code, out message);
                resp.Dispose();
                throw new GraphException(status, code,
                    method + " " + Redact(url) + " -> " + status + " " + (code ?? "") + (message == null ? "" : ": " + message), Redact(url));
            }
        }

        void NoteRedirect(string from, string to)
        {
            var observer = MailboxRedirected;
            if (observer == null) return;
            string a = MailboxApi.MailboxIdOfUrl(from), b = MailboxApi.MailboxIdOfUrl(to);
            if (b == null || string.Equals(a, b, StringComparison.OrdinalIgnoreCase)) return;
            int q = to.IndexOf('?');
            try { observer(a, b, Redact(q < 0 ? to : to.Substring(0, q))); } catch (Exception) { }
        }

        void Wait(HttpResponseMessage resp, int attempt, bool shared = false, string method = null, string url = null, CancellationToken ct = default(CancellationToken))
        {
            TimeSpan delay = ComputeDelay(resp, attempt);
            TimeSpan scaled = TimeSpan.FromTicks((long)(delay.Ticks * Options.DelayScale));
            ThrottleEvent ev = null;
            lock (_lock)
            {
                RetryCount++; Delays.Add(delay);
                if (shared && scaled > TimeSpan.Zero && DateTime.UtcNow + scaled > _gateUntilUtc) _gateUntilUtc = DateTime.UtcNow + scaled;
                if (shared && resp != null)
                {
                    Uri u;
                    ev = new ThrottleEvent
                    {
                        Utc = DateTime.UtcNow, Status = (int)resp.StatusCode, WaitSeconds = delay.TotalSeconds, HadRetryAfter = resp.Headers.RetryAfter != null,
                        Method = method, Host = url != null && Uri.TryCreate(url, UriKind.Absolute, out u) ? u.Host : null
                    };
                    if (ThrottleEvents.Count >= 2000) ThrottleEvents.RemoveAt(0);
                    ThrottleEvents.Add(ev);
                }
            }
            if (ev != null && ThrottleHandler != null) { try { ThrottleHandler(ev); } catch (Exception) { } }
            // the wait ends as soon as the caller cancels (a Retry-After can be 600 s)
            if (scaled > TimeSpan.Zero) ct.WaitHandle.WaitOne(scaled);
            ct.ThrowIfCancellationRequested();
        }

        /// <summary>Sleeps until the delay given by the last 429 / 503 is over (the other workers asked nothing yet: they would be refused too).</summary>
        void WaitGate(CancellationToken ct)
        {
            TimeSpan left;
            lock (_lock) { left = _gateUntilUtc - DateTime.UtcNow; }
            if (left > TimeSpan.Zero) ct.WaitHandle.WaitOne(left);
            ct.ThrowIfCancellationRequested();
        }

        TimeSpan ComputeDelay(HttpResponseMessage resp, int attempt)
        {
            RetryConditionHeaderValue retryAfter = resp == null ? null : resp.Headers.RetryAfter;
            double seconds;
            if (retryAfter != null && retryAfter.Delta.HasValue)
                seconds = retryAfter.Delta.Value.TotalSeconds;
            else if (retryAfter != null && retryAfter.Date.HasValue)
                seconds = Math.Max(0, (retryAfter.Date.Value - DateTimeOffset.UtcNow).TotalSeconds);
            else
            {
                seconds = Math.Min(Options.MaxBackoffSeconds, Options.BaseBackoffSeconds * Math.Pow(2, attempt - 1));
                if (Options.UseJitter) lock (Rng) { seconds += Rng.NextDouble(); }
            }
            return TimeSpan.FromSeconds(Math.Min(seconds, Options.MaxRetryAfterSeconds));
        }

        static string ReadErrorBody(HttpResponseMessage resp)
        {
            try
            {
                using (var s = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
                {
                    var buffer = new byte[8192];
                    int n = s.Read(buffer, 0, buffer.Length);
                    return n > 0 ? Encoding.UTF8.GetString(buffer, 0, n) : "";
                }
            }
            catch { return ""; }
        }

        /// <summary>Graph answers {"error":{"code","message"}}; the Outlook endpoint answers {"Message":"..."}.</summary>
        static void ParseError(string body, out string code, out string message)
        {
            code = null; message = null;
            if (string.IsNullOrWhiteSpace(body)) return;
            try
            {
                using (var doc = JsonDocument.Parse(body))
                {
                    JsonElement root = doc.RootElement, err;
                    if (root.ValueKind != JsonValueKind.Object) return;
                    if (Json.TryGet(root, "error", out err) && err.ValueKind == JsonValueKind.Object)
                    {
                        code = Json.Str(err, "code");
                        message = Json.Str(err, "message");
                        // the import-export API puts the real reason in innererror.internalexception (for instance the permissions a delegated call to another mailbox needs)
                        JsonElement inner, ex;
                        if (Json.TryGet(err, "innererror", out inner) && inner.ValueKind == JsonValueKind.Object
                            && Json.TryGet(inner, "internalexception", out ex) && ex.ValueKind == JsonValueKind.Object)
                        {
                            string detail = Json.Str(ex, "message");
                            if (!string.IsNullOrEmpty(detail) && !string.Equals(detail, message, StringComparison.Ordinal))
                                message = (message ?? "") + " [" + detail + "]";
                        }
                    }
                    else message = Json.Str(root, "message");
                }
            }
            catch (JsonException) { message = body.Length > 300 ? body.Substring(0, 300) : body; }
        }

        // ---------------------------------------------------------------------------------------------
        //  JSON and pagination
        // ---------------------------------------------------------------------------------------------

        /// <param name="finalUrl">Optional: receives the URL that answered, after the redirects (the mailbox of an auto-expanded folder is in it).</param>
        /// <param name="headers">Optional request headers (Prefer: odata.maxpagesize=200, Prefer: outlook.body-content-type="text").</param>
        public JsonDocument GetJson(string url, CancellationToken ct = default(CancellationToken), Action<string> finalUrl = null, IDictionary<string, string> headers = null)
        {
            using (var resp = Send("GET", url, ct: ct, finalUrl: finalUrl, headers: headers))
            using (var s = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
                return JsonDocument.Parse(s);
        }

        public JsonDocument PostJson(string url, string jsonBody, CancellationToken ct = default(CancellationToken))
        {
            Func<HttpContent> body = jsonBody == null ? (Func<HttpContent>)null
                : () => new StringContent(jsonBody, Encoding.UTF8, "application/json");
            using (var resp = Send("POST", url, body, ct: ct))
            using (var s = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
                return JsonDocument.Parse(s);
        }

        /// <summary>
        /// Every element of the "value" arrays of a collection, page after page. The next page is the
        /// @odata.nextLink as returned, provided it is https on graph.microsoft.com.
        /// </summary>
        /// <param name="pageUrl">Optional: called for each page, before its elements are returned, with the URL that answered (after the redirects).</param>
        /// <param name="headers">Optional request headers, sent with every page.</param>
        /// <param name="retryPage">Optional: asked when a page fails with a GraphException, with the exception and the attempt number (1 = the first failure).
        /// It returns true to ask the same page again (after waiting, if it wants to), false to let the exception go; it may throw another exception.</param>
        public IEnumerable<JsonElement> EnumerateValues(string url, CancellationToken ct = default(CancellationToken), Action<string> pageUrl = null, IDictionary<string, string> headers = null,
            Func<GraphException, int, bool> retryPage = null)
        {
            while (url != null)
            {
                string next = null;
                var page = new List<JsonElement>();
                string answered = url;
                JsonDocument fetched;
                for (int attempt = 1; ; attempt++)
                {
                    try { fetched = GetJson(url, ct, u => answered = u, headers); break; }
                    catch (GraphException ex) { if (retryPage == null || !retryPage(ex, attempt)) throw; }
                }
                using (JsonDocument doc = fetched)
                {
                    JsonElement root = doc.RootElement, value, link;
                    if (Json.TryGet(root, "value", out value) && value.ValueKind == JsonValueKind.Array)
                        foreach (JsonElement e in value.EnumerateArray()) page.Add(e.Clone());
                    if (Json.TryGet(root, "@odata.nextLink", out link) && link.ValueKind == JsonValueKind.String)
                        next = link.GetString();
                }
                PagesRead++;
                if (next != null && !IsGraphUrl(next))
                    throw new GraphException(0, "UnexpectedNextLink", "@odata.nextLink is not https on graph.microsoft.com: " + Redact(next), Redact(url));
                if (pageUrl != null) pageUrl(answered);
                foreach (JsonElement e in page) yield return e;
                url = next;
            }
        }
    }

    /// <summary>
    /// Read-only stream over a response body: every read obeys the cancellation token and is abandoned after the inactivity timeout (the inner
    /// stream is disposed to release a read that ignores its token).
    /// </summary>
    public sealed class GuardedReadStream : Stream
    {
        readonly Stream _inner;
        readonly CancellationToken _ct;
        readonly TimeSpan _inactivity;

        public GuardedReadStream(Stream inner, CancellationToken ct, TimeSpan inactivity)
        {
            _inner = inner; _ct = ct; _inactivity = inactivity;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            _ct.ThrowIfCancellationRequested();
            using (var cts = CancellationTokenSource.CreateLinkedTokenSource(_ct))
            {
                if (_inactivity > TimeSpan.Zero && _inactivity != Timeout.InfiniteTimeSpan) cts.CancelAfter(_inactivity);
                using (cts.Token.Register(() => { try { _inner.Dispose(); } catch (Exception) { } }))
                {
                    try { return _inner.ReadAsync(buffer, offset, count, cts.Token).GetAwaiter().GetResult(); }
                    catch (Exception ex) when (ex is OperationCanceledException || ex is ObjectDisposedException || ex is IOException || ex is NullReferenceException)
                    {
                        if (_ct.IsCancellationRequested) throw new OperationCanceledException(_ct);
                        if (cts.IsCancellationRequested)
                            throw new IOException("No data received for " + (int)_inactivity.TotalSeconds + " s: the response is abandoned.", ex);
                        throw;
                    }
                }
            }
        }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { throw new NotSupportedException(); } }
        public override long Position { get { throw new NotSupportedException(); } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }

    /// <summary>Small helpers on System.Text.Json (case-insensitive property lookup, tolerant reads).</summary>
    public static class Json
    {
        public static bool TryGet(JsonElement obj, string name, out JsonElement value)
        {
            if (obj.ValueKind == JsonValueKind.Object)
                foreach (JsonProperty p in obj.EnumerateObject())
                    if (string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase)) { value = p.Value; return true; }
            value = default(JsonElement);
            return false;
        }

        public static string Str(JsonElement obj, string name)
        {
            JsonElement v;
            return TryGet(obj, name, out v) && v.ValueKind == JsonValueKind.String ? v.GetString() : null;
        }

        /// <summary>Claims of an access token (payload only, no validation): who, which permissions, until when.</summary>
        public static JsonDocument JwtClaims(string accessToken)
        {
            string[] parts = accessToken.Split('.');
            if (parts.Length < 2) throw new ArgumentException("Not a JWT.");
            string b64 = parts[1].Replace('-', '+').Replace('_', '/');
            b64 = b64.PadRight(b64.Length + (4 - b64.Length % 4) % 4, '=');
            return JsonDocument.Parse(Convert.FromBase64String(b64));
        }

        public static long Long(JsonElement obj, string name, long fallback = 0)
        {
            JsonElement v; long n;
            return TryGet(obj, name, out v) && v.ValueKind == JsonValueKind.Number && v.TryGetInt64(out n) ? n : fallback;
        }

        public static DateTimeOffset? Date(JsonElement obj, string name)
        {
            DateTimeOffset d;
            string s = Str(obj, name);
            return s != null && DateTimeOffset.TryParse(s, System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.AssumeUniversal, out d) ? d : (DateTimeOffset?)null;
        }
    }
}
