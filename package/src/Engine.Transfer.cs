// =============================================================================
//  Mailbox Transfer - engine, part 6: what the export and import engines share
// -----------------------------------------------------------------------------
//  - TransferMetrics / TransferSnapshot: bytes, items, rates (overall and recent), failures, throttling answers (429).
//  - TransferJob: a job runs on its own threads inside the engine (a PowerShell script block called from another thread
//    would run outside its runspace). The caller polls Snapshot() and TakeMessages() from its own thread, Wait()s, and may
//    Cancel() (Ctrl+C): the workers stop after the request in flight.
//  - ImportSessionPool: one import session per target mailbox, reused for every item until shortly before it expires.
//  - StorePreviews: what the browser of the store shows of an item (from the FastTransfer reader).
//
//  Context: this is the owner's own tooling to migrate or back up Exchange Online mailboxes that the operator is authorised to
//  access (delegated consent of the mailbox user); the lab scripts only target the owner's private lab tenant.
//  Throttling policy: compliance, not evasion. Retry-After is honoured exactly (GraphClient), one 429 / 503 stops every worker
//  until its delay is over, the concurrency is halved on 429 / 503 and grows back slowly, and it never exceeds the documented
//  4 concurrent requests per mailbox. One application and one identity only: nothing here multiplies the quota.
// =============================================================================
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer
{
    /// <summary>The API allows 4 concurrent requests per application and mailbox (Outlook service limits).</summary>
    public static class TransferLimits
    {
        public const int MaxConcurrency = 4;

        public static int CheckConcurrency(int concurrency)
        {
            if (concurrency < 1 || concurrency > MaxConcurrency)
                throw new ArgumentOutOfRangeException("concurrency", "Concurrency must be between 1 and " + MaxConcurrency + " (the service allows 4 concurrent requests per application and mailbox).");
            return concurrency;
        }
    }

    public sealed class TransferSnapshot
    {
        public string Phase;
        public double Seconds;
        public long Bytes, Items, Failed, TotalItems, TotalBytes;
        /// <summary>Since the start of the transfer phase, and over the last 30 seconds or so.</summary>
        public double MiBPerSecond, RecentMiBPerSecond, ItemsPerSecond, GiBPerHour;
        public int Requests, Retries, Throttled429, Active, Limit;

        static string Size(double b)
        {
            string[] u = { "B", "KiB", "MiB", "GiB", "TiB" };
            int i = 0;
            while (b >= 1024 && i < u.Length - 1) { b /= 1024; i++; }
            return i == 0 ? b.ToString("0", CultureInfo.InvariantCulture) + " B" : b.ToString("0.0", CultureInfo.InvariantCulture) + " " + u[i];
        }

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(Phase).Append("  ").Append(TimeSpan.FromSeconds(Math.Round(Seconds)).ToString("hh\\:mm\\:ss"));
            sb.Append("  items ").Append(Items).Append(TotalItems > 0 ? "/" + TotalItems : "");
            if (Failed > 0) sb.Append(" (").Append(Failed).Append(" failed)");
            sb.Append("  ").Append(Size(Bytes)).Append(TotalBytes > 0 ? "/" + Size(TotalBytes) : "");
            sb.Append("  ").Append(MiBPerSecond.ToString("0.0", CultureInfo.InvariantCulture)).Append(" MiB/s");
            sb.Append(" (now ").Append(RecentMiBPerSecond.ToString("0.0", CultureInfo.InvariantCulture)).Append(")");
            sb.Append("  ").Append(ItemsPerSecond.ToString("0.00", CultureInfo.InvariantCulture)).Append(" items/s");
            sb.Append("  ").Append(GiBPerHour.ToString("0.0", CultureInfo.InvariantCulture)).Append(" GiB/h");
            sb.Append("  429: ").Append(Throttled429).Append("  retries: ").Append(Retries).Append("  workers: ").Append(Active).Append("/").Append(Limit);
            return sb.ToString();
        }
    }

    public sealed class TransferMetrics
    {
        readonly Stopwatch _watch = Stopwatch.StartNew();
        readonly List<KeyValuePair<double, long>> _samples = new List<KeyValuePair<double, long>>();
        readonly object _lock = new object();
        long _bytes, _items, _failed;
        long _transferStartedMs = -1;
        public volatile string Phase = "starting";
        public long TotalItems, TotalBytes;
        /// <summary>Progress per folder (the console lists the folders of a running job from it).</summary>
        public readonly FolderTracker Folders = new FolderTracker();
        int _active;

        public void Active(int delta) { Interlocked.Add(ref _active, delta); }

        /// <summary>The transfer phase begins (rates are computed from here, not from the listing of the folders).</summary>
        public void BeginTransfer(long totalItems, long totalBytes)
        {
            TotalItems = totalItems; TotalBytes = totalBytes;
            lock (_lock)
            {
                _transferStartedMs = _watch.ElapsedMilliseconds;
                _samples.Clear();
                _samples.Add(new KeyValuePair<double, long>(0, Interlocked.Read(ref _bytes)));
            }
        }

        /// <summary>Seconds since BeginTransfer, -1 before.</summary>
        public double TransferSeconds
        {
            get { lock (_lock) { return _transferStartedMs < 0 ? -1 : (_watch.ElapsedMilliseconds - _transferStartedMs) / 1000.0; } }
        }

        public void ItemDone(long bytes) { Interlocked.Add(ref _bytes, bytes); Interlocked.Increment(ref _items); }
        public void ItemFailed() { Interlocked.Increment(ref _failed); }
        public long Failed { get { return Interlocked.Read(ref _failed); } }
        public long Items { get { return Interlocked.Read(ref _items); } }
        public long Bytes { get { return Interlocked.Read(ref _bytes); } }

        public TransferSnapshot Snapshot(GraphClient graph)
        {
            var s = new TransferSnapshot { Phase = Phase, Seconds = _watch.Elapsed.TotalSeconds, TotalItems = TotalItems, TotalBytes = TotalBytes };
            s.Bytes = Interlocked.Read(ref _bytes); s.Items = Interlocked.Read(ref _items); s.Failed = Interlocked.Read(ref _failed);
            s.Active = Volatile.Read(ref _active);
            if (graph != null) { s.Requests = graph.RequestCount; s.Retries = graph.RetryCount; s.Throttled429 = graph.Throttled429; }
            lock (_lock)
            {
                if (_transferStartedMs >= 0)
                {
                    double t = (_watch.ElapsedMilliseconds - _transferStartedMs) / 1000.0;
                    if (t > 0)
                    {
                        s.MiBPerSecond = s.Bytes / 1048576.0 / t;
                        s.ItemsPerSecond = s.Items / t;
                        s.GiBPerHour = s.Bytes / 1073741824.0 / t * 3600;
                        _samples.Add(new KeyValuePair<double, long>(t, s.Bytes));
                        var from = _samples[0];
                        foreach (var smp in _samples) { if (t - smp.Key >= 30) from = smp; else break; }
                        double dt = t - from.Key;
                        s.RecentMiBPerSecond = dt > 0 ? (s.Bytes - from.Value) / 1048576.0 / dt : s.MiBPerSecond;
                        while (_samples.Count > 2 && t - _samples[1].Key > 600) _samples.RemoveAt(0);
                    }
                }
            }
            return s;
        }
    }

    /// <summary>One window (5 minutes by default) of the transfer phase: what was done in it, and how many times the service said 429.</summary>
    public sealed class WindowStat
    {
        public int Index;
        public double StartSeconds, EndSeconds;
        public long Bytes, Items, Failed;
        public int Throttled429, Limit;
        public double MiBPerSecond { get { return EndSeconds > StartSeconds ? Bytes / 1048576.0 / (EndSeconds - StartSeconds) : 0; } }
        public double GiBPerHour { get { return EndSeconds > StartSeconds ? Bytes / 1073741824.0 / (EndSeconds - StartSeconds) * 3600 : 0; } }

        static string Clock(double s) { return TimeSpan.FromSeconds(Math.Round(s)).ToString("hh\\:mm\\:ss"); }

        public override string ToString()
        {
            return "window " + Index + " (" + Clock(StartSeconds) + "-" + Clock(EndSeconds) + "): " +
                MiBPerSecond.ToString("0.0", CultureInfo.InvariantCulture) + " MiB/s, " + GiBPerHour.ToString("0.0", CultureInfo.InvariantCulture) + " GiB/h, " +
                Items + " items" + (Failed > 0 ? " (" + Failed + " failed)" : "") + ", 429: " + Throttled429 + ", concurrency " + Limit;
        }
    }

    /// <summary>
    /// Number of requests in flight allowed to the workers: starts at the maximum, is halved when the service answers 429 / 503
    /// (at most once per ReduceCooldown: four workers refused together count once), and goes back up by one after RampAfter
    /// without any throttling answer.
    /// </summary>
    public sealed class AdaptiveLimiter
    {
        readonly object _lock = new object();
        readonly int _max;
        int _limit, _active;
        DateTime _lastReduce = DateTime.MinValue, _lastThrottle = DateTime.MinValue, _lastRaise = DateTime.MinValue;
        public TimeSpan ReduceCooldown = TimeSpan.FromSeconds(10);
        public TimeSpan RampAfter = TimeSpan.FromMinutes(2);

        public AdaptiveLimiter(int max) { _max = max; _limit = max; }
        public int Max { get { return _max; } }
        public int Limit { get { lock (_lock) { return _limit; } } }

        sealed class Slot : IDisposable
        {
            readonly AdaptiveLimiter _l;
            bool _done;
            public Slot(AdaptiveLimiter l) { _l = l; }
            public void Dispose() { if (_done) return; _done = true; _l.Exit(); }
        }

        /// <summary>Waits for a free slot; dispose the result when the request is done.</summary>
        public IDisposable Enter(CancellationToken ct)
        {
            lock (_lock)
            {
                while (true)
                {
                    ct.ThrowIfCancellationRequested();
                    Ramp();
                    if (_active < _limit) { _active++; return new Slot(this); }
                    Monitor.Wait(_lock, 500);
                }
            }
        }

        void Exit() { lock (_lock) { _active--; Monitor.PulseAll(_lock); } }

        void Ramp()
        {
            DateTime now = DateTime.UtcNow;
            DateTime last = _lastThrottle > _lastRaise ? _lastThrottle : _lastRaise;
            if (_limit < _max && now - last >= RampAfter) { _limit++; _lastRaise = now; }
        }

        /// <summary>A 429 / 503 was received. Returns a message when the limit was lowered, else null.</summary>
        public string OnThrottle()
        {
            lock (_lock)
            {
                DateTime now = DateTime.UtcNow;
                _lastThrottle = now;
                if (_limit > 1 && now - _lastReduce >= ReduceCooldown)
                {
                    int before = _limit;
                    _limit = Math.Max(1, _limit / 2);
                    _lastReduce = now;
                    return "concurrency lowered from " + before + " to " + _limit + " (it goes back up by one after " + (int)RampAfter.TotalSeconds + " s without throttling)";
                }
                return null;
            }
        }
    }

    /// <summary>A job that runs on its own threads. The caller polls it (Snapshot, TakeMessages), waits for it and may cancel it.</summary>
    public abstract class TransferJob
    {
        readonly ConcurrentQueue<string> _messages = new ConcurrentQueue<string>();
        readonly object _errorLock = new object();
        Task _task, _sampler;
        Exception _error;
        protected readonly CancellationTokenSource Cts = new CancellationTokenSource();
        public readonly TransferMetrics Metrics = new TransferMetrics();
        public GraphClient Graph { get; protected set; }
        /// <summary>Length of a window of the sustained-rate report (seconds); tests make it short.</summary>
        public static int DefaultWindowSeconds = 300;
        public int WindowSeconds = DefaultWindowSeconds;
        public readonly List<WindowStat> Windows = new List<WindowStat>();
        protected AdaptiveLimiter Limiter;
        double _winStart;
        long _winBytes, _winItems, _winFailed;
        int _win429;
        public bool Cancelled { get { return Cts.IsCancellationRequested && _error == null; } }
        /// <summary>The error that stopped the job (authorisation refused, too many consecutive failures, disk...), or null.</summary>
        public Exception Error { get { return _error; } }

        /// <summary>Creates the limiter of the workers and lets it follow the 429 / 503 answers of the client.</summary>
        protected void InitLimiter(int max)
        {
            Limiter = new AdaptiveLimiter(max);
            Graph.Notice = Log;
            Graph.ThrottleHandler = ev =>
            {
                Log("throttled: " + ev);
                string lowered = Limiter.OnThrottle();
                if (lowered != null) Log(lowered);
            };
        }

        protected void Begin(Action run)
        {
            _task = Task.Factory.StartNew(() =>
            {
                try { run(); }
                catch (OperationCanceledException) { }
                catch (Exception ex) { Fatal(ex); }
                finally
                {
                    Metrics.Phase = Cts.IsCancellationRequested ? "stopped" : "done";
                    if (Graph != null) { Graph.ThrottleHandler = null; Graph.Notice = null; }
                }
            }, TaskCreationOptions.LongRunning);
            _sampler = Task.Factory.StartNew(Sampler, TaskCreationOptions.LongRunning);
        }

        /// <summary>Closes the windows of the transfer phase one after the other (WindowSeconds each) and writes each one to the messages.</summary>
        void Sampler()
        {
            while (!_task.IsCompleted)
            {
                double t = Metrics.TransferSeconds;
                if (t >= 0 && t - _winStart >= WindowSeconds) CloseWindow(t);
                Thread.Sleep(100);
            }
            // the last, partial window: closed when it is long enough or when it holds anything (a window shorter than half a second must not lose the items it saw)
            double end = Metrics.TransferSeconds;
            int c = Graph == null ? 0 : Graph.Throttled429;
            if (end >= 0 && (end > _winStart + 0.5 || Metrics.Items != _winItems || Metrics.Bytes != _winBytes || Metrics.Failed != _winFailed || c != _win429)) CloseWindow(end);
        }

        void CloseWindow(double t)
        {
            long b = Metrics.Bytes, i = Metrics.Items, f = Metrics.Failed;
            int c = Graph == null ? 0 : Graph.Throttled429;
            var w = new WindowStat
            {
                Index = Windows.Count + 1, StartSeconds = _winStart, EndSeconds = t, Bytes = b - _winBytes, Items = i - _winItems, Failed = f - _winFailed,
                Throttled429 = c - _win429, Limit = Limiter == null ? 0 : Limiter.Limit
            };
            _winStart = t; _winBytes = b; _winItems = i; _winFailed = f; _win429 = c;
            lock (Windows) Windows.Add(w);
            Log(w.ToString());
        }

        /// <summary>The windows closed so far (a copy).</summary>
        public WindowStat[] GetWindows() { lock (Windows) return Windows.ToArray(); }

        /// <summary>The 429 / 503 received by the client, with the time waited each time.</summary>
        public ThrottleEvent[] GetThrottleEvents() { return Graph.GetThrottleEvents(); }

        /// <summary>Stops the job for good: the first error wins, the other workers stop after their request.</summary>
        protected void Fatal(Exception ex)
        {
            lock (_errorLock) { if (_error == null) _error = ex; }
            Cts.Cancel();
        }

        protected void Log(string message) { _messages.Enqueue(message); }

        /// <summary>The messages (warnings, items refused, redirects) written since the last call.</summary>
        public string[] TakeMessages()
        {
            var list = new List<string>();
            string m;
            while (_messages.TryDequeue(out m)) list.Add(m);
            return list.ToArray();
        }

        public bool Wait(int milliseconds) { return _task.Wait(milliseconds) && _sampler.Wait(2000); }
        public bool IsCompleted { get { return _task.IsCompleted && _sampler.IsCompleted; } }
        public void Cancel() { Cts.Cancel(); }
        public TransferSnapshot Snapshot()
        {
            TransferSnapshot s = Metrics.Snapshot(Graph);
            s.Limit = Limiter == null ? 0 : Limiter.Limit;
            return s;
        }

        /// <summary>Runs the workers and waits for them (called by the job on its own thread).</summary>
        protected void RunWorkers(int count, Action<int> work)
        {
            var tasks = new Task[count];
            for (int i = 0; i < count; i++)
            {
                int n = i;
                tasks[i] = Task.Factory.StartNew(() =>
                {
                    Metrics.Active(1);
                    try { work(n); }
                    catch (OperationCanceledException) { }
                    catch (Exception ex) { Fatal(ex); }
                    finally { Metrics.Active(-1); }
                }, TaskCreationOptions.LongRunning);
            }
            Task.WaitAll(tasks);
        }

        /// <summary>
        /// Stops the job after too many failures in a row (a disk that is full, a token that is refused, a service that is down):
        /// marking every item Failed one after the other would help nobody.
        /// </summary>
        int _consecutiveFailures;
        public const int MaxConsecutiveFailures = 25;

        protected void Succeeded() { Interlocked.Exchange(ref _consecutiveFailures, 0); }

        protected void FailedOnce(string lastError)
        {
            if (Interlocked.Increment(ref _consecutiveFailures) >= MaxConsecutiveFailures)
                Fatal(new InvalidOperationException("Stopped: " + MaxConsecutiveFailures + " items failed in a row, the last error was: " + lastError));
        }
    }

    /// <summary>
    /// Learns the auxiliary archives of an auto-expanding archive from the redirects the service answers (308 Location of a folder or item
    /// listing, ErrorArchiveFolderMovedPermanently of exportItems, 409 "the target session is expected in mailbox X" of importItem) and records
    /// them in the store: the mailbox (kind AuxArchive) and its location (how and when it was discovered). A mailbox the store already
    /// knows as primary or main archive from the operator or Graph, and the mailboxes the job works on, are never recorded as auxiliary.
    /// </summary>
    public sealed class AuxLocationLearner
    {
        readonly ExportStore _store;
        readonly string _upn;
        readonly HashSet<string> _own;
        readonly Action<string> _log;
        readonly ConcurrentDictionary<string, byte> _seen = new ConcurrentDictionary<string, byte>(StringComparer.OrdinalIgnoreCase);

        public AuxLocationLearner(ExportStore store, string upn, IEnumerable<string> ownMailboxes, Action<string> log)
        {
            _store = store; _upn = upn; _log = log;
            _own = new HashSet<string>(ownMailboxes.Where(m => !string.IsNullOrEmpty(m)), StringComparer.OrdinalIgnoreCase);
        }

        /// <summary>Observer for GraphClient.MailboxRedirected (the third argument starts with "ErrorArchiveFolderMovedPermanently" for an export answer).</summary>
        public void OnRedirect(string from, string to, string detail)
        {
            Learn(to, detail != null && detail.StartsWith("ErrorArchiveFolderMovedPermanently", StringComparison.Ordinal) ? "ExportMoved" : "Redirect308",
                "asked " + from + ": " + detail);
        }

        /// <summary>Records <paramref name="mailboxId"/> as an auxiliary archive (once). True when it was new.</summary>
        public bool Learn(string mailboxId, string source, string detail)
        {
            if (string.IsNullOrEmpty(mailboxId) || _own.Contains(mailboxId) || !_seen.TryAdd(mailboxId, 1)) return false;
            try
            {
                foreach (StoreMailboxLocation known in _store.GetLocations())
                    if (MailboxApi.SameMailbox(known.MailboxId, mailboxId) && known.Kind != StoreMailboxKind.AuxArchive && known.Rank >= StoreMailboxLocation.RankGiven) return false;
                bool isNew = _store.RecordLocation(new StoreMailboxLocation { MailboxId = mailboxId, Kind = StoreMailboxKind.AuxArchive, Upn = _upn, Source = source, Detail = detail, Rank = StoreMailboxLocation.RankLearned });
                _store.EnsureMailbox(new StoreMailbox { Id = mailboxId, Kind = StoreMailboxKind.AuxArchive, Upn = _upn, DisplayName = "AuxArchive mailbox" });
                if (isNew && _log != null) _log("auxiliary archive discovered: " + mailboxId + " (" + source + ")");
                return isNew;
            }
            catch (ObjectDisposedException) { return false; }
        }
    }

    /// <summary>One import session per target mailbox, reused until shortly before its token expires.</summary>
    public sealed class ImportSessionPool
    {
        readonly MailboxApi _api;
        readonly Dictionary<string, ImportSession> _sessions = new Dictionary<string, ImportSession>(StringComparer.Ordinal);
        readonly object _lock = new object();
        public TimeSpan RenewBefore = TimeSpan.FromMinutes(2);
        /// <summary>Sessions created, and the lifetime announced by the service for each (expirationDateTime minus the creation time).</summary>
        public int Created { get; private set; }
        public List<TimeSpan> Lifetimes { get; } = new List<TimeSpan>();

        public ImportSessionPool(MailboxApi api) { _api = api; }

        public ImportSession Get(string mailboxId, CancellationToken ct = default(CancellationToken))
        {
            lock (_lock)
            {
                ImportSession s;
                if (_sessions.TryGetValue(mailboxId, out s) && !Expiring(s)) return s;
                return Create(mailboxId, ct);
            }
        }

        /// <summary>The session was refused (token expired before its time): a new one replaces it, unless another worker already did.</summary>
        public ImportSession Renew(string mailboxId, ImportSession refused, CancellationToken ct = default(CancellationToken))
        {
            lock (_lock)
            {
                ImportSession s;
                if (_sessions.TryGetValue(mailboxId, out s) && !object.ReferenceEquals(s, refused) && !Expiring(s)) return s;
                return Create(mailboxId, ct);
            }
        }

        bool Expiring(ImportSession s) { return s.Expiration.HasValue && s.Expiration.Value - DateTimeOffset.UtcNow < RenewBefore; }

        ImportSession Create(string mailboxId, CancellationToken ct)
        {
            DateTimeOffset before = DateTimeOffset.UtcNow;
            ImportSession s = _api.CreateImportSession(mailboxId, ct);
            _sessions[mailboxId] = s;
            Created++;
            if (s.Expiration.HasValue) Lifetimes.Add(s.Expiration.Value - before);
            return s;
        }
    }

    /// <summary>What the browser of the store shows of an item, from the summary of the FastTransfer reader.</summary>
    public static class StorePreviews
    {
        public static StorePreview From(FxMessageSummary m, long itemKey)
        {
            string sender = !string.IsNullOrEmpty(m.SenderName) ? m.SenderName : m.SentRepresentingName;
            string smtp = !string.IsNullOrEmpty(m.SenderSmtp) ? m.SenderSmtp : m.SentRepresentingSmtp;
            string from = sender != null && smtp != null ? sender + " <" + smtp + ">" : (sender ?? smtp);
            DateTime? sent = m.ClientSubmitTime.HasValue ? m.ClientSubmitTime.Value.UtcDateTime : (DateTime?)null;
            DateTime? received = m.DeliveryTime.HasValue ? m.DeliveryTime.Value.UtcDateTime : sent;
            var names = m.Attachments.Select(a => !string.IsNullOrEmpty(a.FileName) ? a.FileName : a.DisplayName).Where(n => !string.IsNullOrEmpty(n));
            var to = m.Recipients.Where(r => r.Type == 1).Select(r => !string.IsNullOrEmpty(r.DisplayName) ? r.DisplayName : r.Smtp).Where(n => !string.IsNullOrEmpty(n));
            return new StorePreview
            {
                ItemKey = itemKey, Subject = m.Subject, From = from, To = string.Join("; ", to), MessageClass = m.MessageClass,
                HasAttachments = m.Attachments.Count > 0, AttachmentCount = m.Attachments.Count, AttachmentNames = string.Join("; ", names),
                BodyPreview = !string.IsNullOrEmpty(m.BodyTextPreview) ? m.BodyTextPreview : m.BodyHtmlPreview, SentUtc = sent, ReceivedUtc = received
            };
        }
    }
}
