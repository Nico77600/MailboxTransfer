// =============================================================================
//  Mailbox Transfer - fake mailbox import and export service (tests only, no network)
// -----------------------------------------------------------------------------
//  An HttpMessageHandler that behaves like the API for a few in-memory mailboxes: folders, items, paged listings,
//  exportItems, createImportSession, importItem, folder creation. Thread-safe (the engines call it from several workers).
//  It can throttle (429 with Retry-After), fail requests or items on demand, and measures the concurrency it sees.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer.Tests
{
    public sealed class FakeFolder
    {
        public string Id, Name, ParentId, Type = "IPF.Note", WellKnown;
        /// <summary>Not in the top-level listing of the service (Recoverable Items: reachable by id and by its well-known name only, measured live 2026-11-09).</summary>
        public bool Hidden;
    }

    public sealed class FakeItem
    {
        public string Id, ChangeKey, FolderId;
        public byte[] Data;
        /// <summary>createdDateTime / lastModifiedDateTime of the listing: items added by a test are old, items created by importItem carry the time of the import (as on the real service).</summary>
        public DateTime CreatedUtc = new DateTime(2026, 10, 28, 9, 0, 0, DateTimeKind.Utc);
    }

    /// <summary>A listing of the items of a folder that fails on purpose (see <see cref="FakeMailboxService.ListingFailures"/>).</summary>
    public sealed class FakeListingFailure
    {
        /// <summary>How many listings still fail; int.MaxValue: always.</summary>
        public int Left = int.MaxValue;
        public int Status = 500;
        public string Code = "ErrorInternalServerError";
        public string Message = "An internal server error occurred. The operation failed., Cannot query rows in a table.";
        /// <summary>How many listings it has failed so far.</summary>
        public int Hits;
    }

    /// <summary>Builds ids in the layout Exchange Online gives (see ExchangeEntryId): the mailbox guid as text, then the EntryID of a folder or of a message.</summary>
    public static class FakeIds
    {
        static void Counter(List<byte> b, long n) { for (int i = 5; i >= 0; i--) b.Add((byte)((n >> (8 * i)) & 0xFF)); }
        static byte[] Bytes(Guid g) { return g.ToByteArray(); }
        /// <summary>A provider uid that differs per mailbox (live: not the mailbox guid).</summary>
        static byte[] Provider(string mailboxGuid) { return Guid.Parse(mailboxGuid).ToByteArray().Select(x => (byte)(x ^ 0x5A)).Reverse().ToArray(); }

        static string Wrap(string mailboxGuid, List<byte> entry)
        {
            var b = new List<byte> { 0x00, 0x03, 0x24, 0x00 };
            b.AddRange(Encoding.ASCII.GetBytes(mailboxGuid));
            b.Add(0x00); b.Add((byte)(entry.Count & 0xFF)); b.Add((byte)(entry.Count >> 8));
            b.AddRange(entry);
            return Convert.ToBase64String(b.ToArray()).Replace('+', '-').Replace('/', '_');
        }

        /// <summary>The id of folder number <paramref name="folder"/> (global counter) of the database <paramref name="database"/> as the mailbox <paramref name="mailboxGuid"/> lists it. The provider uid is derived from the mailbox: it differs per mailbox, as live.</summary>
        public static string Folder(string mailboxGuid, long folder, string database = "1a2b3c07-5e7f-4a1b-9c2d-3e4f5a6b7c8d")
        {
            var e = new List<byte> { 0, 0, 0, 0 };
            e.AddRange(Provider(mailboxGuid));
            e.Add(0x01); e.Add(0x00);
            e.AddRange(Bytes(Guid.Parse(database)));
            Counter(e, folder);
            e.Add(0); e.Add(0);
            return Wrap(mailboxGuid, e);
        }

        public static string Message(string mailboxGuid, long folder, long message, string database = "1a2b3c07-5e7f-4a1b-9c2d-3e4f5a6b7c8d")
        {
            var e = new List<byte> { 0, 0, 0, 0 };
            e.AddRange(Provider(mailboxGuid));
            e.Add(0x07); e.Add(0x00);
            e.AddRange(Bytes(Guid.Parse(database)));
            Counter(e, folder);
            e.Add(0); e.Add(0);
            e.AddRange(Bytes(Guid.Parse(database)));
            Counter(e, message);
            e.Add(0); e.Add(0);
            return Wrap(mailboxGuid, e);
        }
    }

    public sealed class FakeExportCall
    {
        public int Items;
        public long Bytes;
        public string Mailbox;
    }

    public sealed class FakeMailboxService : HttpMessageHandler
    {
        readonly object _lock = new object();
        readonly Dictionary<string, List<FakeFolder>> _folders = new Dictionary<string, List<FakeFolder>>();
        readonly Dictionary<string, List<FakeItem>> _items = new Dictionary<string, List<FakeItem>>();
        int _seq, _inFlight;
        int _throttleLeft;
        double _throttleRetryAfter;
        string _throttleContains;

        public int PageSize = 3;
        /// <summary>Milliseconds each request takes (lets the workers overlap).</summary>
        public int LatencyMs = 5;
        public int MaxConcurrent { get; private set; }
        public readonly List<FakeExportCall> ExportCalls = new List<FakeExportCall>();
        public readonly List<string> Log = new List<string>();
        public readonly List<DateTime> Times = new List<DateTime>();
        public readonly List<DateTime> ThrottledAt = new List<DateTime>();
        /// <summary>exportItems answers 403 (permission missing).</summary>
        public bool ForbidExport;
        public int Throttled;
        public readonly HashSet<string> BadItemIds = new HashSet<string>();
        /// <summary>A request that exports more than this number of items fails with 500 (exercises the item by item retry).</summary>
        public int FailExportAbove = int.MaxValue;
        /// <summary>importItem answers this status for the next calls (a queue; 0 = succeed).</summary>
        public readonly Queue<int> ImportFailures = new Queue<int>();
        /// <summary>One entry per importItem call that succeeds: the item is created, then the 200 answer carries this body instead of {"itemId":...} ("" = an empty body, "garbage" = text that is not JSON). A null entry = the normal answer.</summary>
        public readonly Queue<string> ImportAnswers = new Queue<string>();
        /// <summary>Creating a folder of one of these names is refused with status <see cref="RefuseFolderStatus"/> (nothing is created).</summary>
        public readonly HashSet<string> RefuseFolderNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public int RefuseFolderStatus = 403;
        /// <summary>
        /// One entry per importItem call: n >= 0 the connection drops after n bytes of the body (nothing is created); -1 the whole body is read,
        /// then the answer is lost (the item IS created: see <see cref="CreateOnLostAnswer"/>); -2 the request fails before any byte of the body
        /// is sent (connection refused, TLS, DNS: nothing is created).
        /// </summary>
        public readonly Queue<long> UploadDrops = new Queue<long>();
        /// <summary>The answer lost by a -1 entry of <see cref="UploadDrops"/> comes after the service created the item (the worst case: it exists).</summary>
        public bool CreateOnLostAnswer;
        /// <summary>An injected 409 names this mailbox as the one where the session is expected.</summary>
        public string Force409Mailbox;
        public string ImportExpirationOffsetMinutes = "60";
        public int SessionsCreated;
        /// <summary>
        /// Auto-expanding archive: folder id -> the auxiliary archive that really holds it. importItem into such a folder through another mailbox answers 409
        /// "the target session is expected in mailbox X". The folder is still listed by the main archive (a placeholder, no item).
        /// </summary>
        public readonly Dictionary<string, string> MovedFolderToMailbox = new Dictionary<string, string>();
        /// <summary>
        /// With <see cref="MovedFolderToMailbox"/>: folder requests (folder, childFolders, items, folder creation) for a moved folder answer 308 with the Location
        /// in the auxiliary archive, and exportItems of an item that lives there answers ErrorArchiveFolderMovedPermanently with the URL to use (as documented).
        /// </summary>
        public bool RedirectFolders;
        /// <summary>The main archive lists the items of a moved folder itself (no 308), whatever mailbox they are in: the case where only exportItems tells (unverified live).</summary>
        public bool ListMovedFromMain;
        /// <summary>The users/{id}/settings/exchange answers (null: 403 as with a token without User.Read.All). Beta has inPlaceArchiveMailboxId, v1.0 only primaryMailboxId.</summary>
        public string ExchangeSettingsBetaJson, ExchangeSettingsV1Json;
        /// <summary>Mailboxes whose folder listing answers a bare 500 (no code, no message): a soft-deleted or inactive mailbox has no route (measured live 2026-11-09).</summary>
        public readonly HashSet<string> BareFailureMailboxes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>Mailboxes whose folder listing answers 404 ErrorInvalidUser.</summary>
        public readonly HashSet<string> InvalidUserMailboxes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        /// <summary>The entries of <see cref="UploadDrops"/> apply only to importItem calls of this mailbox (null: any).</summary>
        public string UploadDropMailbox;
        /// <summary>As the real service does (measured): GET .../folders lists the first level only, the children come from .../folders/{id}/childFolders. Off: every folder in one listing.</summary>
        public bool RootOnlyFolders;
        /// <summary>
        /// "mailbox|folder id" -> a listing of the items of that folder that fails: by default 500 ErrorInternalServerError "Cannot query rows in a table" as the
        /// auxiliary archive answered for a folder whose content was not moved yet (live 2026-10-31).
        /// </summary>
        public readonly Dictionary<string, FakeListingFailure> ListingFailures = new Dictionary<string, FakeListingFailure>();

        public string AddMailbox(string id)
        {
            lock (_lock) { _folders[id] = new List<FakeFolder>(); _items[id] = new List<FakeItem>(); }
            return id;
        }

        public FakeFolder AddFolder(string mailbox, string id, string name, string parentId = "ROOT", string type = "IPF.Note")
        {
            var f = new FakeFolder { Id = id, Name = name, ParentId = parentId, Type = type };
            lock (_lock) _folders[mailbox].Add(f);
            return f;
        }

        public FakeItem AddItem(string mailbox, string folderId, string id, byte[] data, string changeKey = "ck1", DateTime? createdUtc = null)
        {
            var i = new FakeItem { Id = id, FolderId = folderId, Data = data, ChangeKey = changeKey };
            if (createdUtc.HasValue) i.CreatedUtc = createdUtc.Value;
            lock (_lock) _items[mailbox].Add(i);
            return i;
        }

        public void ChangeItem(string mailbox, string id, byte[] data, string changeKey)
        {
            lock (_lock) { var i = _items[mailbox].First(x => x.Id == id); i.Data = data; i.ChangeKey = changeKey; }
        }

        public FakeItem[] ItemsOf(string mailbox, string folderId = null)
        {
            lock (_lock) return _items[mailbox].Where(i => folderId == null || i.FolderId == folderId).ToArray();
        }

        public FakeFolder[] FoldersOf(string mailbox) { lock (_lock) return _folders[mailbox].ToArray(); }

        /// <summary>The next <paramref name="count"/> requests answer 429 with this Retry-After (seconds).</summary>
        public void Throttle(int count, double retryAfterSeconds, string urlContains = null)
        {
            lock (_lock) { _throttleLeft = count; _throttleRetryAfter = retryAfterSeconds; _throttleContains = urlContains; }
        }

        public int RequestsOf(string contains) { lock (_lock) return Log.Count(l => l.Contains(contains)); }

        static HttpResponseMessage Json(HttpStatusCode code, string json)
        {
            return new HttpResponseMessage(code) { Content = new StringContent(json, Encoding.UTF8, "application/json") };
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            string method = request.Method.Method;
            string url = request.RequestUri.OriginalString;
            long drop = long.MinValue;
            lock (_lock)
            {
                if (method == "POST" && url.Contains("/importItem?") && UploadDrops.Count > 0 && (UploadDropMailbox == null || MailboxOf(url) == UploadDropMailbox)) drop = UploadDrops.Dequeue();
            }
            if (drop >= 0)   // the connection drops while the body is written: the service got nothing and creates nothing
            {
                try { await request.Content.CopyToAsync(new DroppingStream(drop)).ConfigureAwait(false); }
                catch (IOException ex) { throw new HttpRequestException("Error while copying content to a stream.", ex); }
                throw new InvalidOperationException("the body was smaller than the drop limit");
            }
            if (drop == -2) throw new HttpRequestException("No connection could be made because the target machine actively refused it.");   // nothing sent
            byte[] body = request.Content == null ? null : await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
            if (drop == -1)   // the whole body was received, the answer is lost
            {
                if (CreateOnLostAnswer) { lock (_lock) { Import(MailboxOf(url), body, false); } }
                throw new HttpRequestException("The response ended prematurely.");
            }
            int now = Interlocked.Increment(ref _inFlight);
            try
            {
                lock (_lock)
                {
                    if (now > MaxConcurrent) MaxConcurrent = now;
                    Log.Add(method + " " + url); Times.Add(DateTime.UtcNow);
                    if (_throttleLeft > 0 && (_throttleContains == null || url.Contains(_throttleContains)))
                    {
                        _throttleLeft--;
                        Throttled++; ThrottledAt.Add(DateTime.UtcNow);
                        var r = Json((HttpStatusCode)429, "{\"error\":{\"code\":\"TooManyRequests\",\"message\":\"slow down\"}}");
                        r.Headers.TryAddWithoutValidation("Retry-After", _throttleRetryAfter.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        return r;
                    }
                }
                if (LatencyMs > 0) await Task.Delay(LatencyMs, ct).ConfigureAwait(false);
                return Route(method, url, body);
            }
            finally { Interlocked.Decrement(ref _inFlight); }
        }

        string MailboxOf(string url)
        {
            int i = url.IndexOf("MBX:", StringComparison.Ordinal);
            if (i < 0) return null;
            int j = i;
            while (j < url.Length && url[j] != '/' && url[j] != '\'' && url[j] != '?') j++;
            return url.Substring(i, j - i);
        }

        HttpResponseMessage Route(string method, string url, byte[] body)
        {
            string mailbox = MailboxOf(url);
            string path = url.Split('?')[0];
            lock (_lock)
            {
                if (path.EndsWith("/settings/exchange"))
                {
                    string settings = url.Contains("/beta/") ? ExchangeSettingsBetaJson : ExchangeSettingsV1Json;
                    return string.IsNullOrEmpty(settings) ? Json(HttpStatusCode.Forbidden, "{\"error\":{\"code\":\"Authorization_RequestDenied\",\"message\":\"User.Read.All is missing\"}}") : Json(HttpStatusCode.OK, settings);
                }
                string moved308 = mailbox == null || !_folders.ContainsKey(mailbox) ? null : RedirectOf(mailbox, url, path);
                if (moved308 != null)
                {
                    var redirect = new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent("") };
                    redirect.Headers.Location = new Uri(moved308);
                    return redirect;
                }
                if (mailbox == null || !_folders.ContainsKey(mailbox)) return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorMailboxNotFound\",\"message\":\"unknown mailbox\"}}");
                if (method == "GET" && path.EndsWith("/folders") && BareFailureMailboxes.Contains(mailbox)) return new HttpResponseMessage(HttpStatusCode.InternalServerError) { Content = new StringContent("") };
                if (method == "GET" && path.EndsWith("/folders") && InvalidUserMailboxes.Contains(mailbox)) return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorInvalidUser\",\"message\":\"The requested user is invalid.\"}}");
                if (method == "GET" && path.EndsWith("/folders")) return ListFolders(mailbox);
                if (method == "GET" && path.EndsWith("/items")) return ListItems(mailbox, url, path);
                if (method == "GET" && path.EndsWith("/childFolders")) return ChildFolders(mailbox, path);
                if (method == "GET" && path.Contains("/folders/")) return GetFolder(mailbox, path);
                if (method == "POST" && path.EndsWith("/exportItems")) return Export(mailbox, body);
                if (method == "POST" && path.EndsWith("/createImportSession")) return CreateSession(mailbox);
                if (method == "POST" && path.EndsWith("/importItem")) return Import(mailbox, body);
                if (method == "POST" && (path.EndsWith("/folders") || path.EndsWith("/childFolders"))) return CreateFolder(mailbox, path, body);
            }
            return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"NoRoute\",\"message\":\"" + method + " " + path.Replace("\"", "'") + "\"}}");
        }

        string RedirectOf(string mailbox, string url, string path)
        {
            if (!RedirectFolders) return null;
            int at = path.IndexOf("/folders/", StringComparison.Ordinal);
            if (at < 0) return null;
            string rest = path.Substring(at + 9);
            int slash = rest.IndexOf('/');
            string folderId = Uri.UnescapeDataString(slash < 0 ? rest : rest.Substring(0, slash));
            string target;
            if (!MovedFolderToMailbox.TryGetValue(folderId, out target) || target == mailbox) return null;
            return url.Replace(mailbox, target);
        }

        string FolderJson(string mailbox, FakeFolder f)
        {
            int children = _folders[mailbox].Count(x => x.ParentId == f.Id);
            int count = _items[mailbox].Count(i => i.FolderId == f.Id);
            return JsonSerializer.Serialize(new Dictionary<string, object>
            {
                { "id", f.Id }, { "displayName", f.Name }, { "parentFolderId", f.ParentId }, { "type", f.Type }, { "childFolderCount", children }, { "totalItemCount", count }, { "wellKnownName", f.WellKnown }
            });
        }

        HttpResponseMessage ListFolders(string mailbox)
        {
            var folders = (RootOnlyFolders ? _folders[mailbox].Where(f => f.ParentId == "ROOT") : _folders[mailbox]).Where(f => !f.Hidden);
            return Json(HttpStatusCode.OK, "{\"value\":[" + string.Join(",", folders.Select(f => FolderJson(mailbox, f))) + "]}");
        }

        HttpResponseMessage ChildFolders(string mailbox, string path)
        {
            int at = path.LastIndexOf("/folders/", StringComparison.Ordinal);
            string id = Uri.UnescapeDataString(path.Substring(at + 9, path.Length - at - 9 - "/childFolders".Length));
            return Json(HttpStatusCode.OK, "{\"value\":[" + string.Join(",", _folders[mailbox].Where(f => f.ParentId == id).Select(f => FolderJson(mailbox, f))) + "]}");
        }

        HttpResponseMessage GetFolder(string mailbox, string path)
        {
            int at = path.LastIndexOf("/folders/", StringComparison.Ordinal);
            string id = Uri.UnescapeDataString(path.Substring(at + 9));
            FakeFolder f = _folders[mailbox].FirstOrDefault(x => x.Id == id) ?? (id.StartsWith("recoverableitems", StringComparison.Ordinal) ? _folders[mailbox].FirstOrDefault(x => x.WellKnown == id) : null);
            if (f == null) return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"The specified object was not found in the store.\"}}");
            return Json(HttpStatusCode.OK, FolderJson(mailbox, f));
        }

        HttpResponseMessage ListItems(string mailbox, string url, string path)
        {
            int at = path.LastIndexOf("/folders/", StringComparison.Ordinal);
            string folderId = Uri.UnescapeDataString(path.Substring(at + 9, path.Length - at - 9 - "/items".Length));
            FakeListingFailure failure;
            if (ListingFailures.TryGetValue(mailbox + "|" + folderId, out failure) && failure.Left > 0)
            {
                if (failure.Left != int.MaxValue) failure.Left--;
                failure.Hits++;
                return Json((HttpStatusCode)failure.Status, "{\"error\":{\"code\":" + JsonSerializer.Serialize(failure.Code) + ",\"message\":" + JsonSerializer.Serialize(failure.Message) + "}}");
            }
            int skip = 0;
            int q = url.IndexOf("skip=", StringComparison.Ordinal);
            if (q >= 0) skip = int.Parse(url.Substring(q + 5).Split('&')[0]);
            string holder;
            var source = ListMovedFromMain && MovedFolderToMailbox.TryGetValue(folderId, out holder) && holder != mailbox ? _items[holder] : _items[mailbox];
            var all = source.Where(i => i.FolderId == folderId).ToList();
            var page = all.Skip(skip).Take(PageSize).ToList();
            var sb = new StringBuilder("{\"value\":[");
            sb.Append(string.Join(",", page.Select(i => JsonSerializer.Serialize(new Dictionary<string, object>
            {
                { "id", i.Id }, { "type", "IPM.Note" }, { "changeKey", i.ChangeKey }, { "size", i.Data.Length },
                { "createdDateTime", i.CreatedUtc.ToString("o") }, { "lastModifiedDateTime", i.CreatedUtc.ToString("o") }
            }))));
            sb.Append("]");
            if (skip + PageSize < all.Count)
                sb.Append(",\"@odata.nextLink\":\"https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/" + mailbox + "/folders/" + Uri.EscapeDataString(folderId) + "/items?%24skip=" + (skip + PageSize) + "\"");
            sb.Append("}");
            return Json(HttpStatusCode.OK, sb.ToString());
        }

        HttpResponseMessage Export(string mailbox, byte[] body)
        {
            string[] ids;
            using (var doc = JsonDocument.Parse(body)) ids = doc.RootElement.GetProperty("itemIds").EnumerateArray().Select(e => e.GetString()).ToArray();
            long bytes = 0;
            var parts = new List<string>();
            foreach (string id in ids)
            {
                var item = _items[mailbox].FirstOrDefault(i => i.Id == id);
                if (item == null)
                {
                    // an item of an auxiliary archive asked from the main archive: the documented ErrorArchiveFolderMovedPermanently with the URL to use
                    string holder = null;
                    foreach (var kv in _items)
                    {
                        FakeItem other = kv.Key == mailbox ? null : kv.Value.FirstOrDefault(i => i.Id == id);
                        string target;
                        if (other != null && MovedFolderToMailbox.TryGetValue(other.FolderId, out target) && target == kv.Key) { holder = kv.Key; break; }
                    }
                    if (holder != null && (RedirectFolders || ListMovedFromMain))
                    {
                        parts.Add("{\"itemId\":" + JsonSerializer.Serialize(id) + ",\"Error\":{\"code\":\"ErrorArchiveFolderMovedPermanently\",\"message\":\"https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/" + holder + "/exportItems\"}}");
                        continue;
                    }
                }
                if (item == null || BadItemIds.Contains(id))
                {
                    parts.Add("{\"itemId\":" + JsonSerializer.Serialize(id) + ",\"Error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"not found\"}}");
                    continue;
                }
                bytes += item.Data.Length;
                parts.Add("{\"itemId\":" + JsonSerializer.Serialize(id) + ",\"changeKey\":" + JsonSerializer.Serialize(item.ChangeKey) + ",\"data\":\"" + Convert.ToBase64String(item.Data) + "\"}");
            }
            if (ForbidExport) return Json(HttpStatusCode.Forbidden, "{\"error\":{\"code\":\"ErrorAccessDenied\",\"message\":\"MailboxItem.Export is missing\"}}");
            ExportCalls.Add(new FakeExportCall { Items = ids.Length, Bytes = bytes, Mailbox = mailbox });
            if (ids.Length > FailExportAbove) return Json(HttpStatusCode.InternalServerError, "{\"error\":{\"code\":\"InternalServerError\",\"message\":\"too many items\"}}");
            return Json(HttpStatusCode.OK, "{\"value\":[" + string.Join(",", parts) + "]}");
        }

        HttpResponseMessage CreateSession(string mailbox)
        {
            SessionsCreated++;
            string exp = DateTime.UtcNow.AddMinutes(double.Parse(ImportExpirationOffsetMinutes, System.Globalization.CultureInfo.InvariantCulture)).ToString("o");
            return Json(HttpStatusCode.OK, "{\"importUrl\":\"https://outlook.office365.com/api/gv1.0/Mailboxes('" + mailbox + "')/importItem?authtoken=TOK" + SessionsCreated + "\",\"expirationDateTime\":\"" + exp + "\"}");
        }

        HttpResponseMessage Import(string mailbox, byte[] body, bool injectFailures = true)
        {
            if (injectFailures && ImportFailures.Count > 0)
            {
                int status = ImportFailures.Dequeue();
                if (status == 409 && Force409Mailbox != null) return Json(HttpStatusCode.Conflict, "{\"Message\":\"Invalid import session. The target session is expected in mailbox " + Force409Mailbox + ".\"}");
                if (status != 0) return Json((HttpStatusCode)status, "{\"Message\":\"injected failure " + status + "\"}");
            }
            string folderId, mode, data;
            using (var doc = JsonDocument.Parse(body))
            {
                folderId = doc.RootElement.GetProperty("FolderId").GetString();
                mode = doc.RootElement.GetProperty("Mode").GetString();
                data = doc.RootElement.GetProperty("Data").GetString();
            }
            string moved;
            if (MovedFolderToMailbox.TryGetValue(folderId, out moved) && moved != mailbox)
                return Json(HttpStatusCode.Conflict, "{\"Message\":\"Invalid import session. The target session is expected in mailbox " + moved + ".\"}");
            if (!_folders[mailbox].Any(f => f.Id == folderId)) return Json(HttpStatusCode.NotFound, "{\"Message\":\"folder not found\"}");
            if (mode != "create") return Json(HttpStatusCode.BadRequest, "{\"Message\":\"only create is supported\"}");
            string id = "NEW" + (++_seq).ToString("D6");
            _items[mailbox].Add(new FakeItem { Id = id, FolderId = folderId, Data = Convert.FromBase64String(data), ChangeKey = "ckN", CreatedUtc = DateTime.UtcNow });
            string answer = injectFailures && ImportAnswers.Count > 0 ? ImportAnswers.Dequeue() : null;
            if (answer != null) return Json(HttpStatusCode.OK, answer == "" ? "" : "<html>not json</html>");
            return Json(HttpStatusCode.OK, "{\"itemId\":\"" + id + "\",\"changeKey\":\"ckN\"}");
        }

        HttpResponseMessage CreateFolder(string mailbox, string path, byte[] body)
        {
            string name, type;
            using (var doc = JsonDocument.Parse(body))
            {
                name = doc.RootElement.GetProperty("displayName").GetString();
                type = doc.RootElement.TryGetProperty("type", out var t) ? t.GetString() : "IPF.Note";
            }
            if (RefuseFolderNames.Contains(name)) return Json((HttpStatusCode)RefuseFolderStatus, "{\"error\":{\"code\":\"ErrorAccessDenied\",\"message\":\"the folder " + name + " is refused\"}}");
            string parent = "ROOT";
            int at = path.LastIndexOf("/folders/", StringComparison.Ordinal);
            if (path.EndsWith("/childFolders") && at >= 0) parent = Uri.UnescapeDataString(path.Substring(at + 9, path.Length - at - 9 - "/childFolders".Length));
            string id = "F" + (++_seq).ToString("D6");
            var f = new FakeFolder { Id = id, Name = name, ParentId = parent, Type = type };
            _folders[mailbox].Add(f);
            string holds;
            if (MovedFolderToMailbox.TryGetValue(parent, out holds) && holds == mailbox) MovedFolderToMailbox[id] = mailbox;   // a child created in an auxiliary archive lives there
            return Json(HttpStatusCode.Created, FolderJson(mailbox, f));
        }
    }
}
