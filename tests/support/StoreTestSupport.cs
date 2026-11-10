// Test helpers for Store.Tests.ps1 (compiled together with package\src\Engine.Store.cs, never shipped).
// Data is generated on the fly (deterministic from a seed): no source file, and the expected SHA-256 of an item can be
// recomputed independently of the store.

using System;
using System.Buffers.Binary;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer.TestSupport
{
    /// <summary>Sequential pseudo-random stream (xorshift64*), deterministic from the seed.</summary>
    public class GenStream : Stream
    {
        private ulong _state;
        private long _remaining;
        private long _pos;

        public GenStream(long seed, long length)
        {
            _state = (ulong)seed * 0x9E3779B97F4A7C15UL + 0x1234567UL;
            if (_state == 0) _state = 1;
            _remaining = length;
        }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { throw new NotSupportedException(); } }
        public override long Position { get { return _pos; } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_remaining <= 0) return 0;
            int n = (int)Math.Min(count, _remaining);
            var span = new Span<byte>(buffer, offset, n);
            int i = 0;
            ulong x = _state;
            for (; i + 8 <= n; i += 8)
            {
                x ^= x >> 12; x ^= x << 25; x ^= x >> 27;
                BinaryPrimitives.WriteUInt64LittleEndian(span.Slice(i, 8), x * 0x2545F4914F6CDD1DUL);
            }
            if (i < n)
            {
                x ^= x >> 12; x ^= x << 25; x ^= x >> 27;
                Span<byte> tail = stackalloc byte[8];
                BinaryPrimitives.WriteUInt64LittleEndian(tail, x * 0x2545F4914F6CDD1DUL);
                tail.Slice(0, n - i).CopyTo(span.Slice(i));
            }
            _state = x;
            _remaining -= n;
            _pos += n;
            return n;
        }

        public static string ExpectedSha256(long seed, long length)
        {
            using (var s = new GenStream(seed, length))
            using (var h = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
            {
                byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(256 * 1024);
                try
                {
                    int n;
                    while ((n = s.Read(buf, 0, buf.Length)) > 0) h.AppendData(buf, 0, n);
                }
                finally { System.Buffers.ArrayPool<byte>.Shared.Return(buf); }
                return Convert.ToHexString(h.GetHashAndReset()).ToLowerInvariant();
            }
        }
    }

    /// <summary>Produces stallAfter bytes, signals (event and optional file), then blocks until Release is set, then fails.</summary>
    public sealed class StallStream : GenStream
    {
        private readonly long _stallAfter;
        private readonly string _signalFile;
        private bool _stalled;
        public readonly ManualResetEventSlim Reached = new ManualResetEventSlim(false);
        public readonly ManualResetEventSlim Release = new ManualResetEventSlim(false);

        public StallStream(long seed, long length, long stallAfter, string signalFile) : base(seed, length)
        {
            _stallAfter = stallAfter; _signalFile = signalFile;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            long left = _stallAfter - Position;
            if (left <= 0)
            {
                _stalled = true;
                if (!string.IsNullOrEmpty(_signalFile)) File.WriteAllText(_signalFile, "stalled at " + Position);
                Reached.Set();
                Release.Wait();
                throw new IOException("source aborted after stall");
            }
            return base.Read(buffer, offset, (int)Math.Min(count, left));
        }
    }

    public sealed class FailingStream : GenStream
    {
        private readonly long _failAfter;
        public FailingStream(long seed, long length, long failAfter) : base(seed, length) { _failAfter = failAfter; }
        public override int Read(byte[] buffer, int offset, int count)
        {
            if (Position >= _failAfter) throw new IOException("simulated source failure at " + Position);
            return base.Read(buffer, offset, (int)Math.Min(count, _failAfter - Position));
        }
    }

    public sealed class Job
    {
        public long Key;
        public long Seed;
        public long Length;
    }

    public sealed class RunResult
    {
        public long Count;
        public long Bytes;
        public double Seconds;
        public string[] Errors;
        public double MBps { get { return Seconds > 0 ? Bytes / 1048576.0 / Seconds : 0; } }
        public double ItemsPerSec { get { return Seconds > 0 ? Count / Seconds : 0; } }
    }

    public static class Workload
    {
        private static ulong Mix(ulong z)
        {
            z += 0x9E3779B97F4A7C15UL;
            z = (z ^ (z >> 30)) * 0xBF58476D1CE4E5B9UL;
            z = (z ^ (z >> 27)) * 0x94D049BB133111EBUL;
            return z ^ (z >> 31);
        }

        /// <summary>Deterministic size of item #key for the crash workloads: 1 to 16 MiB.</summary>
        public static long CrashSize(long key)
        {
            return 1048576L + (long)(Mix((ulong)key) % (15UL * 1048576UL));
        }

        public static string ItemId(long n)
        {
            // ~130 characters, like a real Graph/EWS item id
            byte[] b = new byte[96];
            ulong x = Mix((ulong)n * 7919UL + 13UL);
            for (int i = 0; i < b.Length; i += 8)
            {
                x = Mix(x);
                BinaryPrimitives.WriteUInt64LittleEndian(new Span<byte>(b, i, 8), x);
            }
            return "AAMkAD" + Convert.ToBase64String(b).Replace('/', '_').Replace('+', '-');
        }

        public static string[] FolderIds(int count)
        {
            return Enumerable.Range(0, count).Select(i => "folder-" + i.ToString("D3")).ToArray();
        }

        /// <summary>Folders F000.. : F000 = root, then 3 levels (parent of Fi = F(i-1)/4 for i&gt;0 style tree).</summary>
        public static List<StoreFolder> MakeFolders(string mailboxId, string[] ids)
        {
            var list = new List<StoreFolder>();
            for (int i = 0; i < ids.Length; i++)
            {
                list.Add(new StoreFolder
                {
                    MailboxId = mailboxId,
                    FolderId = ids[i],
                    ParentFolderId = i == 0 ? null : ids[(i - 1) / 4],
                    DisplayName = i == 0 ? "Root" : "Folder " + i,
                    FolderType = "IPF.Note"
                });
            }
            return list;
        }

        /// <summary>
        /// Builds the entries and the write jobs. Sizes: mostly 1 to 40 KiB, then midCount items of 1 to 8 MiB, then the
        /// explicit big items. Order of the jobs is shuffled (seeded) so that big items run concurrently with small ones.
        /// Seed of an item = its position + 1 (stable).
        /// </summary>
        public static void Build(string mailboxId, string[] folderIds, int count, int midCount, long[] bigSizes, int rngSeed,
            out List<StoreItemEntry> entries, out List<long> sizes)
        {
            var rnd = new Random(rngSeed);
            entries = new List<StoreItemEntry>(count);
            sizes = new List<long>(count);
            for (int i = 0; i < count; i++)
            {
                long size;
                if (i < bigSizes.Length) size = bigSizes[i];
                else if (i < bigSizes.Length + midCount) size = 1048576L + rnd.Next(7 * 1048576);
                else size = 1024L + rnd.Next(39 * 1024);
                entries.Add(new StoreItemEntry
                {
                    MailboxId = mailboxId,
                    ItemId = ItemId(i),
                    FolderId = folderIds[rnd.Next(folderIds.Length)],
                    ChangeKey = "CK" + i,
                    ItemType = "IPM.Note",
                    Size = size
                });
                sizes.Add(size);
            }
        }

        public static List<long> ShuffledOrder(int count, int rngSeed)
        {
            var rnd = new Random(rngSeed);
            var idx = Enumerable.Range(0, count).ToList();
            for (int i = idx.Count - 1; i > 0; i--) { int j = rnd.Next(i + 1); long t = idx[i]; idx[i] = idx[j]; idx[j] = (int)t; }
            return idx.Select(x => (long)x).ToList();
        }
    }

    public static class Runner
    {
        /// <summary>Writes all the jobs with N concurrent threads (sync or async API). Never throws: errors are returned.</summary>
        public static RunResult WriteAll(ExportStore store, IList<Job> jobs, int threads, bool useAsync)
        {
            var errors = new ConcurrentQueue<string>();
            long next = -1, bytes = 0, count = 0;
            var sw = Stopwatch.StartNew();
            var ts = new List<Thread>();
            for (int t = 0; t < threads; t++)
            {
                var th = new Thread(() =>
                {
                    while (true)
                    {
                        long i = Interlocked.Increment(ref next);
                        if (i >= jobs.Count) break;
                        var j = jobs[(int)i];
                        try
                        {
                            using (var s = new GenStream(j.Seed, j.Length))
                            {
                                StoreWriteResult r = useAsync
                                    ? store.WriteItemAsync(j.Key, s, CancellationToken.None).GetAwaiter().GetResult()
                                    : store.WriteItem(j.Key, s, CancellationToken.None);
                                if (r.Length != j.Length) errors.Enqueue("item " + j.Key + ": length " + r.Length + " != " + j.Length);
                                Interlocked.Add(ref bytes, r.Length);
                                Interlocked.Increment(ref count);
                            }
                        }
                        catch (Exception ex) { errors.Enqueue("item " + j.Key + ": " + ex.GetType().Name + ": " + ex.Message); }
                    }
                });
                th.IsBackground = true;
                ts.Add(th);
                th.Start();
            }
            foreach (var th in ts) th.Join();
            sw.Stop();
            return new RunResult { Count = count, Bytes = bytes, Seconds = sw.Elapsed.TotalSeconds, Errors = errors.ToArray() };
        }

        /// <summary>Starts one WriteItem on the thread pool; the task result is null or the error message.</summary>
        public static Task<string> StartWrite(ExportStore store, long key, Stream source)
        {
            return Task.Run(() =>
            {
                try { store.WriteItem(key, source, CancellationToken.None); return (string)null; }
                catch (Exception ex) { return ex.GetType().Name + ": " + ex.Message; }
            });
        }

        /// <summary>
        /// Reads every job back (streaming) with N threads and checks: length, SHA-256 against the independent generator,
        /// and against the catalog value.
        /// </summary>
        public static RunResult VerifyAll(ExportStore store, IList<Job> jobs, int threads)
        {
            var errors = new ConcurrentQueue<string>();
            long next = -1, bytes = 0, count = 0;
            var sw = Stopwatch.StartNew();
            var ts = new List<Thread>();
            for (int t = 0; t < threads; t++)
            {
                var th = new Thread(() =>
                {
                    byte[] buf = new byte[256 * 1024];
                    while (true)
                    {
                        long i = Interlocked.Increment(ref next);
                        if (i >= jobs.Count) break;
                        var j = jobs[(int)i];
                        try
                        {
                            var item = store.GetItem(j.Key);
                            if (item == null || item.State != StoreItemState.Exported) { errors.Enqueue("item " + j.Key + ": not exported"); continue; }
                            long total = 0;
                            string actual;
                            using (var s = store.OpenItem(j.Key))
                            using (var h = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
                            {
                                int n;
                                while ((n = s.Read(buf, 0, buf.Length)) > 0) { h.AppendData(buf, 0, n); total += n; }
                                actual = Convert.ToHexString(h.GetHashAndReset()).ToLowerInvariant();
                            }
                            Interlocked.Add(ref bytes, total);
                            Interlocked.Increment(ref count);
                            if (total != j.Length) errors.Enqueue("item " + j.Key + ": read " + total + " != " + j.Length);
                            else if (actual != item.Sha256) errors.Enqueue("item " + j.Key + ": sha differs from the catalog");
                            else if (actual != GenStream.ExpectedSha256(j.Seed, j.Length)) errors.Enqueue("item " + j.Key + ": sha differs from the generator");
                        }
                        catch (Exception ex) { errors.Enqueue("item " + j.Key + ": " + ex.GetType().Name + ": " + ex.Message); }
                    }
                });
                th.IsBackground = true;
                ts.Add(th);
                th.Start();
            }
            foreach (var th in ts) th.Join();
            sw.Stop();
            return new RunResult { Count = count, Bytes = bytes, Seconds = sw.Elapsed.TotalSeconds, Errors = errors.ToArray() };
        }

        /// <summary>Reads (without checking content) all the Exported items of a store: raw read throughput.</summary>
        public static RunResult ReadAllRaw(ExportStore store, IList<long> keys, int threads)
        {
            long next = -1, bytes = 0, count = 0;
            var errors = new ConcurrentQueue<string>();
            var sw = Stopwatch.StartNew();
            var ts = new List<Thread>();
            for (int t = 0; t < threads; t++)
            {
                var th = new Thread(() =>
                {
                    byte[] buf = new byte[1024 * 1024];
                    while (true)
                    {
                        long i = Interlocked.Increment(ref next);
                        if (i >= keys.Count) break;
                        try
                        {
                            using (var s = store.OpenItem(keys[(int)i]))
                            {
                                int n;
                                while ((n = s.Read(buf, 0, buf.Length)) > 0) Interlocked.Add(ref bytes, n);
                            }
                            Interlocked.Increment(ref count);
                        }
                        catch (Exception ex) { errors.Enqueue(ex.Message); }
                    }
                });
                th.IsBackground = true;
                ts.Add(th);
                th.Start();
            }
            foreach (var th in ts) th.Join();
            sw.Stop();
            return new RunResult { Count = count, Bytes = bytes, Seconds = sw.Elapsed.TotalSeconds, Errors = errors.ToArray() };
        }
    }

    /// <summary>Hammers the read APIs (tree, pages, stats, pending) while writers work: they must never fail nor block the writers.</summary>
    public sealed class BackgroundReader
    {
        private readonly ExportStore _store;
        private readonly string _mailbox;
        private readonly Thread _t;
        private volatile bool _stop;
        public long Loops;
        public readonly ConcurrentQueue<string> Errors = new ConcurrentQueue<string>();

        public BackgroundReader(ExportStore store, string mailboxId)
        {
            _store = store; _mailbox = mailboxId;
            _t = new Thread(Run) { IsBackground = true };
        }

        public void Start() { _t.Start(); }
        public void Stop() { _stop = true; _t.Join(); }

        private void Run()
        {
            var rnd = new Random(1);
            while (!_stop)
            {
                try
                {
                    var tree = _store.GetFolderTree(_mailbox);
                    var stats = _store.GetStats();
                    var pend = _store.GetPendingItems(_mailbox, 100);
                    if (tree.Count > 0)
                    {
                        var all = new List<StoreFolderNode>();
                        var st = new Stack<StoreFolderNode>(tree);
                        while (st.Count > 0) { var n = st.Pop(); all.Add(n); foreach (var c in n.Children) st.Push(c); }
                        var f = all[rnd.Next(all.Count)];
                        _store.GetItems(f.FolderKey, 0, 50, (StoreItemSort)rnd.Next(3), rnd.Next(2) == 0);
                    }
                    Interlocked.Increment(ref Loops);
                }
                catch (Exception ex) { Errors.Enqueue(ex.GetType().Name + ": " + ex.Message); }
                Thread.Sleep(20);
            }
        }
    }

    /// <summary>Samples the private memory and the working set of the current process.</summary>
    public sealed class MemorySampler
    {
        private Timer _timer;
        public long PeakPrivate;
        public long PeakWorkingSet;
        public long PeakManaged;
        public long Baseline;

        public void Start(int intervalMs)
        {
            var p = Process.GetCurrentProcess();
            p.Refresh();
            Baseline = p.PrivateMemorySize64;
            PeakPrivate = Baseline; PeakWorkingSet = p.WorkingSet64; PeakManaged = GC.GetTotalMemory(false);
            _timer = new Timer(_ => Sample(), null, 0, intervalMs);
        }

        private void Sample()
        {
            try
            {
                using (var p = Process.GetCurrentProcess())
                {
                    long pm = p.PrivateMemorySize64, ws = p.WorkingSet64, mg = GC.GetTotalMemory(false);
                    if (pm > PeakPrivate) PeakPrivate = pm;
                    if (ws > PeakWorkingSet) PeakWorkingSet = ws;
                    if (mg > PeakManaged) PeakManaged = mg;
                }
            }
            catch { }
        }

        public void Stop()
        {
            if (_timer != null) { _timer.Dispose(); _timer = null; }
            Sample();
        }
    }
}
