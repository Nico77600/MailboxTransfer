// =============================================================================
//  Mailbox Transfer - engine, part 5i: events of a job for the console (JSON lines), stop request, keep awake
// -----------------------------------------------------------------------------
//  The console never parses text: a job (the entry script in a child process) appends one JSON object per line to events.jsonl in
//  its job folder and the console reads the new lines (EventTail) with a timer. The format is documented in
//  package\docs\MailboxTransfer-Events.md. Kinds: plan, step, object, item, wait, rate, warn, decision, error, result, slept, info.
//  A graceful stop is a file: StopWatcher cancels a token when the file appears (the effect of Ctrl+C). AwakeKeeper keeps Windows from
//  sleeping by idle while a long job runs (SetThreadExecutionState is per thread: a thread of its own holds it).
// =============================================================================
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace MailboxTransfer
{
    /// <summary>Thread-safe JSON-lines appender. One call = one complete line, flushed at once; a reader never sees half a line except while the OS writes it.</summary>
    public sealed class EventLog : IDisposable
    {
        public const int Version = 1;
        readonly object _lock = new object();
        readonly FileStream _fs;
        readonly Func<DateTime> _clock;
        readonly Timer _heartbeat;
        readonly double _sleptSeconds;
        DateTime _lastTick;
        bool _disposed;
        public string Path { get; private set; }
        /// <summary>Lines written so far.</summary>
        public long Count { get; private set; }
        /// <summary>The last time (UTC) a line was written.</summary>
        public DateTime LastEventUtc { get; private set; }

        /// <param name="clock">UTC clock (tests inject one); null = the system clock.</param>
        /// <param name="heartbeatMs">Period of the heartbeat that detects a frozen process (sleep); 0 = no timer, call Heartbeat() yourself.</param>
        /// <param name="sleptSeconds">A wall-clock jump larger than this between two heartbeats is reported as a "slept" event.</param>
        public EventLog(string path, Func<DateTime> clock = null, int heartbeatMs = 5000, double sleptSeconds = 120)
        {
            if (string.IsNullOrEmpty(path)) throw new ArgumentException("No events file.");
            Path = System.IO.Path.GetFullPath(path);
            string dir = System.IO.Path.GetDirectoryName(Path);
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            _fs = new FileStream(Path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite | FileShare.Delete, 4096, FileOptions.WriteThrough);
            _clock = clock ?? (() => DateTime.UtcNow);
            _sleptSeconds = sleptSeconds;
            _lastTick = _clock();
            LastEventUtc = _lastTick;
            if (heartbeatMs > 0) _heartbeat = new Timer(_ => { try { Heartbeat(); } catch (Exception) { } }, null, heartbeatMs, heartbeatMs);
        }

        public static string Iso(DateTime utc) { return utc.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture); }

        /// <summary>One heartbeat: when the wall clock moved by more than the threshold since the previous sign of life, the process was frozen (sleep, hibernation): a "slept" line.</summary>
        public void Heartbeat() { CheckSlept(_clock()); }

        void CheckSlept(DateTime now)
        {
            DateTime from;
            lock (_lock) { from = _lastTick; _lastTick = now; }
            double gap = (now - from).TotalSeconds;
            if (gap > _sleptSeconds)
                WriteCore("slept", new Dictionary<string, object> { { "from", Iso(from) }, { "to", Iso(now) }, { "seconds", Math.Round(gap) } }, now);
        }

        // ---------------------------------------------------------------- generic line
        /// <summary>Writes {"v":1,"t":...,"e":kind, fields...}. Values: null (skipped), string, bool, numbers, DateTime, IDictionary, IEnumerable.</summary>
        public void Write(string kind, IDictionary fields)
        {
            if (string.IsNullOrEmpty(kind)) throw new ArgumentException("No event kind.");
            DateTime now = _clock();
            if (kind != "slept") CheckSlept(now);
            WriteCore(kind, fields, now);
        }

        void WriteCore(string kind, IDictionary fields, DateTime now)
        {
            byte[] line;
            using (var ms = new MemoryStream())
            {
                using (var w = new Utf8JsonWriter(ms))
                {
                    w.WriteStartObject();
                    w.WriteNumber("v", Version);
                    w.WriteString("t", Iso(now));
                    w.WriteString("e", kind);
                    if (fields != null)
                        foreach (DictionaryEntry kv in fields)
                        {
                            string name = Convert.ToString(kv.Key, CultureInfo.InvariantCulture);
                            if (name == "v" || name == "t" || name == "e" || kv.Value == null) continue;
                            w.WritePropertyName(name);
                            Value(w, kv.Value);
                        }
                    w.WriteEndObject();
                }
                ms.WriteByte((byte)'\n');
                line = ms.ToArray();
            }
            lock (_lock)
            {
                if (_disposed) return;
                _fs.Write(line, 0, line.Length);
                _fs.Flush();
                Count++;
                LastEventUtc = now;
            }
        }

        static void Value(Utf8JsonWriter w, object v)
        {
            if (v == null) { w.WriteNullValue(); return; }
            if (v is string) { w.WriteStringValue((string)v); return; }
            if (v is bool) { w.WriteBooleanValue((bool)v); return; }
            if (v is DateTime) { w.WriteStringValue(Iso((DateTime)v)); return; }
            if (v is DateTimeOffset) { w.WriteStringValue(Iso(((DateTimeOffset)v).UtcDateTime)); return; }
            if (v is double || v is float || v is decimal)
            {
                double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
                if (double.IsNaN(d) || double.IsInfinity(d)) w.WriteNullValue(); else w.WriteNumberValue(d);
                return;
            }
            if (v is sbyte || v is byte || v is short || v is ushort || v is int || v is uint || v is long) { w.WriteNumberValue(Convert.ToInt64(v, CultureInfo.InvariantCulture)); return; }
            if (v is ulong) { w.WriteNumberValue((ulong)v); return; }
            if (v is Enum) { w.WriteStringValue(v.ToString()); return; }
            var dict = v as IDictionary;
            if (dict != null)
            {
                w.WriteStartObject();
                foreach (DictionaryEntry kv in dict)
                {
                    if (kv.Value == null) continue;
                    w.WritePropertyName(Convert.ToString(kv.Key, CultureInfo.InvariantCulture));
                    Value(w, kv.Value);
                }
                w.WriteEndObject();
                return;
            }
            var seq = v as IEnumerable;
            if (seq != null)
            {
                w.WriteStartArray();
                foreach (object o in seq) Value(w, o);
                w.WriteEndArray();
                return;
            }
            w.WriteStringValue(Convert.ToString(v, CultureInfo.InvariantCulture));
        }

        static Dictionary<string, object> D(params object[] kv)
        {
            var d = new Dictionary<string, object>();
            for (int i = 0; i + 1 < kv.Length; i += 2) d[(string)kv[i]] = kv[i + 1];
            return d;
        }

        // ---------------------------------------------------------------- the kinds
        /// <summary>steps: pairs (id, title).</summary>
        public void Plan(string job, IEnumerable<KeyValuePair<string, string>> steps)
        {
            var list = new List<object>();
            foreach (var s in steps) list.Add(D("id", s.Key, "title", s.Value));
            Write("plan", D("job", job, "steps", list));
        }

        /// <summary>state: waiting | running | done | failed. -1 = unknown (the field is left out).</summary>
        public void Step(string id, string state, long done = -1, long total = -1, long bytes = -1, long bytesTotal = -1, string msg = null)
        {
            Write("step", D("id", id, "state", state, "done", done >= 0 ? (object)done : null, "total", total >= 0 ? (object)total : null,
                "bytes", bytes >= 0 ? (object)bytes : null, "bytesTotal", bytesTotal >= 0 ? (object)bytesTotal : null, "msg", msg));
        }

        public void Object(string step, string key, string state, long done = -1, long total = -1, string msg = null, long bytes = -1)
        {
            Write("object", D("step", step, "key", key, "state", state, "done", done >= 0 ? (object)done : null, "total", total >= 0 ? (object)total : null,
                "bytes", bytes >= 0 ? (object)bytes : null, "msg", msg));
        }

        public void Item(string step, string key, string state, string msg = null, string detail = null)
        {
            Write("item", D("step", step, "key", key, "state", state, "msg", msg, "detail", detail));
        }

        public void Wait(string reason, double seconds, int count = -1, double waited = -1)
        {
            Write("wait", D("reason", reason, "seconds", seconds, "count", count >= 0 ? (object)count : null, "waited", waited >= 0 ? (object)waited : null));
        }

        public void Rate(string window, double gibPerHour) { Write("rate", D("window", window, "gibPerHour", Math.Round(gibPerHour, 2))); }

        public void Warn(string code, string step, string key, string msg) { Write("warn", D("code", code, "step", step, "key", key, "msg", msg)); }

        public void Decision(string key, string msg, IEnumerable<string> candidates = null) { Write("decision", D("key", key, "msg", msg, "candidates", candidates)); }

        public void Error(string code, string msg, string tech = null) { Write("error", D("code", code, "msg", msg, "tech", tech)); }

        public void Info(string msg) { Write("info", D("msg", msg)); }

        public void Result(int exit, bool complete, IDictionary counts = null, string next = null, string reason = null)
        {
            Write("result", D("exit", exit, "complete", complete, "counts", counts, "next", next, "reason", reason));
        }

        public void Dispose()
        {
            lock (_lock)
            {
                if (_disposed) return;
                _disposed = true;
            }
            if (_heartbeat != null) _heartbeat.Dispose();
            _fs.Dispose();
        }
    }

    /// <summary>The lines added to an events file since a byte offset (what the console reads with its timer).</summary>
    public sealed class EventTailResult
    {
        public string[] Lines = new string[0];
        /// <summary>Offset to give to the next call (just after the last complete line read).</summary>
        public long Offset;
    }

    public static class EventTail
    {
        /// <summary>
        /// Reads the complete lines written after <paramref name="offset"/>. A last line without its newline (still being written) is left for the next call.
        /// A file shorter than the offset (replaced) is read again from zero. A missing file gives no line.
        /// </summary>
        public static EventTailResult Read(string path, long offset)
        {
            var r = new EventTailResult { Offset = offset };
            if (!File.Exists(path)) return r;
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                if (fs.Length < offset) offset = 0;
                r.Offset = offset;
                long n = fs.Length - offset;
                if (n <= 0) return r;
                if (n > 64L * 1024 * 1024) n = 64L * 1024 * 1024;
                var buf = new byte[n];
                fs.Position = offset;
                int got = 0;
                while (got < n) { int k = fs.Read(buf, got, (int)(n - got)); if (k <= 0) break; got += k; }
                int last = got - 1;
                while (last >= 0 && buf[last] != (byte)'\n') last--;
                if (last < 0) return r;
                string text = new UTF8Encoding(false).GetString(buf, 0, last + 1);
                var lines = new List<string>();
                foreach (string l in text.Split('\n')) { string t = l.TrimEnd('\r'); if (t.Length > 0) lines.Add(t); }
                r.Lines = lines.ToArray();
                r.Offset = offset + last + 1;
            }
            return r;
        }
    }

    /// <summary>Polls a stop-request file; when it exists the token is cancelled (the same effect as Ctrl+C: the requests in flight finish, the state is saved).</summary>
    public sealed class StopWatcher : IDisposable
    {
        readonly CancellationTokenSource _cts = new CancellationTokenSource();
        readonly Timer _timer;
        public string Path { get; private set; }
        public CancellationToken Token { get { return _cts.Token; } }
        public bool IsRequested { get { return _cts.IsCancellationRequested; } }

        /// <summary>The file is looked at once at once, then every <paramref name="pollMs"/>.</summary>
        public StopWatcher(string path, int pollMs = 1000)
        {
            Path = path;
            Check();
            _timer = new Timer(_ => Check(), null, pollMs, pollMs);
        }

        void Check()
        {
            try { if (!_cts.IsCancellationRequested && File.Exists(Path)) _cts.Cancel(); } catch (Exception) { }
        }

        public void Dispose() { _timer.Dispose(); _cts.Dispose(); }
    }

    /// <summary>Keeps Windows from going to sleep by idle (not from a lid closed or a manual Sleep) until Dispose. The flag is per thread: a thread of its own holds it.</summary>
    public sealed class AwakeKeeper : IDisposable
    {
        const uint EsContinuous = 0x80000000, EsSystemRequired = 0x00000001;
        [DllImport("kernel32.dll")] static extern uint SetThreadExecutionState(uint flags);

        readonly ManualResetEventSlim _started = new ManualResetEventSlim(false);
        readonly ManualResetEventSlim _release = new ManualResetEventSlim(false);
        readonly Thread _thread;
        bool _disposed;
        /// <summary>Windows accepted the request.</summary>
        public bool Succeeded { get; private set; }
        /// <summary>The thread holds the request now.</summary>
        public bool IsActive { get; private set; }
        /// <summary>The state the thread had just before it released it (EsContinuous | EsSystemRequired when it was held): proof for the tests.</summary>
        public uint StateBeforeRelease { get; private set; }

        AwakeKeeper()
        {
            _thread = new Thread(Run) { IsBackground = true, Name = "MailboxTransfer keep awake" };
        }

        public static AwakeKeeper Begin()
        {
            var k = new AwakeKeeper();
            k._thread.Start();
            k._started.Wait(5000);
            return k;
        }

        void Run()
        {
            Succeeded = SetThreadExecutionState(EsContinuous | EsSystemRequired) != 0;
            IsActive = Succeeded;
            _started.Set();
            _release.Wait();
            StateBeforeRelease = SetThreadExecutionState(EsContinuous);
            IsActive = false;
        }

        public void Dispose()
        {
            lock (this) { if (_disposed) return; _disposed = true; }
            _release.Set();
            _thread.Join(5000);
            IsActive = false;
        }
    }
    // =============================================================================================================
    //  Progress per folder: what the Run page of the console lists. The engine jobs (export, import, PST import, PST export, EML export
    //  and import) report each item they finish to a FolderTracker; the entry script turns the folders that changed into "object" events.
    // =============================================================================================================

    /// <summary>The progress of one folder (a copy: it does not change after it was taken).</summary>
    public sealed class FolderStat
    {
        public string Key;
        /// <summary>Items to do in this run (-1: not known yet).</summary>
        public long Total = -1;
        public long Done, Failed, Unknown, Skipped, Bytes;
        /// <summary>planned, running, done, partial (some items failed or were skipped), check (some are Unknown), failed, or an explicit state (notready, refused).</summary>
        public string State;
        public string Msg;
    }

    /// <summary>
    /// Counts per folder, safe to call from the worker threads. Expect says how many items a folder has to do, Done / Fail / Unknown / Skip report one
    /// item each; TakeChanged gives the folders that changed since the last call (the console follows them with its events file).
    /// </summary>
    public sealed class FolderTracker
    {
        sealed class Entry
        {
            public string Key; public long Total = -1, Done, Failed, Unknown, Skipped, Bytes; public string Explicit, ExplicitMsg; public bool Dirty, Finished;
        }
        readonly object _lock = new object();
        readonly Dictionary<string, Entry> _map = new Dictionary<string, Entry>(StringComparer.Ordinal);
        readonly List<Entry> _order = new List<Entry>();

        Entry Get(string key)
        {
            if (key == null) key = "";
            Entry e;
            if (!_map.TryGetValue(key, out e)) { e = new Entry { Key = key }; _map[key] = e; _order.Add(e); }
            e.Dirty = true;
            return e;
        }

        public int Count { get { lock (_lock) return _order.Count; } }

        /// <summary>Adds items to do in a folder (a folder can be told about several times: the totals add up).</summary>
        public void Expect(string key, long items)
        {
            if (items <= 0) return;
            lock (_lock) { Entry e = Get(key); e.Total = (e.Total < 0 ? 0 : e.Total) + items; e.Finished = false; }
        }

        public void Done(string key, long bytes = 0) { lock (_lock) { Entry e = Get(key); e.Done++; e.Bytes += Math.Max(0, bytes); } }
        public void Fail(string key) { lock (_lock) Get(key).Failed++; }
        public void Skip(string key) { lock (_lock) Get(key).Skipped++; }
        public void Unknown(string key) { lock (_lock) Get(key).Unknown++; }

        /// <summary>An Unknown item was settled by a look-up: it exists (created), or it does not and will be sent again (not created).</summary>
        public void Settled(string key, bool created, long bytes = 0)
        {
            lock (_lock)
            {
                Entry e = Get(key);
                if (e.Unknown > 0) e.Unknown--; else if (created) e.Total = (e.Total < 0 ? 0 : e.Total) + 1;
                if (created) { e.Done++; e.Bytes += Math.Max(0, bytes); }
            }
        }

        /// <summary>An Unknown item that will not be sent again and was not created: it counts as failed.</summary>
        public void SettledFailed(string key)
        {
            lock (_lock) { Entry e = Get(key); if (e.Unknown > 0) e.Unknown--; else e.Total = (e.Total < 0 ? 0 : e.Total) + 1; e.Failed++; }
        }

        /// <summary>The folder has no more items to come (its total was not known): the total becomes what was counted.</summary>
        public void Finish(string key)
        {
            lock (_lock) { Entry e = Get(key); e.Finished = true; }
        }

        /// <summary>A state of the folder that the counts do not tell (notready, refused...), with the reason.</summary>
        public void Mark(string key, string state, string msg = null)
        {
            lock (_lock) { Entry e = Get(key); e.Explicit = state; e.ExplicitMsg = msg; }
        }

        static FolderStat Copy(Entry e)
        {
            long processed = e.Done + e.Failed + e.Unknown + e.Skipped;
            long total = e.Total;
            if (e.Finished || (total >= 0 && total < processed)) total = Math.Max(total, processed);
            var s = new FolderStat { Key = e.Key, Total = total, Done = e.Done, Failed = e.Failed, Unknown = e.Unknown, Skipped = e.Skipped, Bytes = e.Bytes };
            if (e.Explicit != null) { s.State = e.Explicit; s.Msg = e.ExplicitMsg; return s; }
            bool finished = e.Finished || (total > 0 && processed >= total);
            if (finished)
            {
                if (e.Failed > 0 && e.Done == 0) s.State = "failed";
                else if (e.Unknown > 0) s.State = "check";
                else if (e.Failed > 0 || e.Skipped > 0) s.State = "partial";
                else s.State = "done";
            }
            else s.State = processed > 0 ? "running" : "planned";
            if (s.State == "failed" || s.State == "check" || s.State == "partial")
            {
                var parts = new List<string>();
                if (e.Failed > 0) parts.Add(e.Failed + " failed");
                if (e.Unknown > 0) parts.Add(e.Unknown + " to check");
                if (e.Skipped > 0) parts.Add(e.Skipped + " not exported");
                s.Msg = e.Done.ToString(CultureInfo.InvariantCulture) + " / " + total.ToString(CultureInfo.InvariantCulture) + ", " + string.Join(", ", parts);
            }
            return s;
        }

        /// <summary>The folders that changed since the last call, in the order they were first seen.</summary>
        public List<FolderStat> TakeChanged()
        {
            var list = new List<FolderStat>();
            lock (_lock)
                foreach (Entry e in _order)
                    if (e.Dirty) { e.Dirty = false; list.Add(Copy(e)); }
            return list;
        }

        /// <summary>Every folder, changed or not.</summary>
        public List<FolderStat> Snapshot()
        {
            lock (_lock) return _order.Select(Copy).ToList();
        }

        /// <summary>The sums over all the folders (Key null): items to do (the folders whose total is not known count what was done), done, failed, unknown, skipped, bytes.</summary>
        public FolderStat Sum()
        {
            var sum = new FolderStat { Total = 0 };
            lock (_lock)
                foreach (Entry e in _order)
                {
                    FolderStat s = Copy(e);
                    sum.Total += Math.Max(0, s.Total); sum.Done += s.Done; sum.Failed += s.Failed; sum.Unknown += s.Unknown; sum.Skipped += s.Skipped; sum.Bytes += s.Bytes;
                }
            return sum;
        }
    }

    /// <summary>The label of a folder of an export store in the console: its path, with the mailbox in front when it is not the primary one.</summary>
    public sealed class StoreFolderLabels
    {
        readonly ExportStore _store;
        Dictionary<long, string> _labels;
        readonly object _lock = new object();
        public StoreFolderLabels(ExportStore store) { _store = store; }

        public string Label(long? folderKey)
        {
            if (!folderKey.HasValue || folderKey.Value < 0) return "(no folder)";
            lock (_lock)
            {
                string s;
                if (_labels != null && _labels.TryGetValue(folderKey.Value, out s)) return s;
                Load();
                return _labels.TryGetValue(folderKey.Value, out s) ? s : "(folder " + folderKey.Value + ")";
            }
        }

        void Load()
        {
            var d = new Dictionary<long, string>();
            foreach (StoreMailbox m in _store.GetMailboxes())
            {
                string prefix = m.Kind == StoreMailboxKind.Primary ? "" : (m.Kind == StoreMailboxKind.Archive ? "Archive/" : "Auxiliary archive/");
                foreach (StoreFolder f in _store.GetFolders(m.Id)) d[f.Key] = prefix + (f.Path ?? f.DisplayName);
            }
            _labels = d;
        }
    }

    /// <summary>Writes the folders of a tracker that changed as "object" events, every period and once more at Dispose: for the jobs that run in the calling thread (PST export, EML).</summary>
    public sealed class FolderEventPump : IDisposable
    {
        readonly EventLog _log;
        readonly FolderTracker _tracker;
        readonly string _step, _stepEvents;
        readonly Timer _timer;
        readonly object _lock = new object();
        bool _disposed;
        long _lastDone = -1, _lastTotal = -1;

        /// <param name="stepEvents">The id of a step that gets a "running" step event with the sums of the tracker (items done and to do, bytes) when they change; null: none.</param>
        public FolderEventPump(EventLog log, FolderTracker tracker, string step, int periodMs = 1000, string stepEvents = null)
        {
            _log = log; _tracker = tracker; _step = step; _stepEvents = stepEvents;
            if (periodMs > 0) _timer = new Timer(_ => { try { Flush(); } catch (Exception) { } }, null, periodMs, periodMs);
        }

        /// <summary>Writes the "object" events of the folders that changed (and the step event of the sums). Returns how many events.</summary>
        public int Flush()
        {
            lock (_lock)
            {
                int n = Write(_log, _tracker, _step);
                if (_stepEvents != null && _log != null)
                {
                    FolderStat sum = _tracker.Sum();
                    long done = sum.Done + sum.Failed + sum.Unknown + sum.Skipped;
                    if (done != _lastDone || sum.Total != _lastTotal)
                    {
                        _lastDone = done; _lastTotal = sum.Total;
                        _log.Step(_stepEvents, "running", done, sum.Total > 0 ? sum.Total : -1, sum.Bytes > 0 ? sum.Bytes : -1, -1, null);
                        n++;
                    }
                }
                return n;
            }
        }

        public static int Write(EventLog log, FolderTracker tracker, string step)
        {
            if (log == null || tracker == null) return 0;
            int n = 0;
            foreach (FolderStat f in tracker.TakeChanged())
            {
                log.Object(step, f.Key, f.State, f.Done, f.Total, f.Msg, f.Bytes > 0 ? f.Bytes : -1);
                n++;
            }
            return n;
        }

        public void Dispose()
        {
            lock (_lock) { if (_disposed) return; _disposed = true; }
            if (_timer != null) _timer.Dispose();
            lock (_lock) Flush();
        }
    }
}
