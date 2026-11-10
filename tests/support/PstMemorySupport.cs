// Test and lab helpers for the tests of bounded memory (tests\Pst.Memory.Tests.ps1, tools\Measure-PstMemory.ps1): a value of any size
// generated on the fly, a memory probe, a streaming comparison. Compiled in the same unit as the engine (see Import-StoreEngine.ps1).
using System;
using System.Buffers.Text;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;

namespace MailboxTransfer.TestSupport
{
    /// <summary>A value of any length whose bytes are a function of their position (nothing is stored): compared by regenerating it.</summary>
    public sealed class PatternBlob : FxBlob
    {
        readonly long _length;
        readonly ulong _seed;

        public PatternBlob(long length, ulong seed) { _length = length; _seed = seed; }

        public override long Length { get { return _length; } }
        public override Stream Open(long offset, long length) { return new PatternStream(_seed, offset, length); }

        static ulong Mix(ulong z)
        {
            z += 0x9E3779B97F4A7C15UL;
            z = (z ^ (z >> 30)) * 0xBF58476D1CE4E5B9UL;
            z = (z ^ (z >> 27)) * 0x94D049BB133111EBUL;
            return z ^ (z >> 31);
        }

        sealed class PatternStream : Stream
        {
            readonly ulong _seed;
            long _pos, _end;

            public PatternStream(ulong seed, long offset, long length) { _seed = seed; _pos = offset; _end = offset + length; }

            public override bool CanRead { get { return true; } }
            public override bool CanSeek { get { return false; } }
            public override bool CanWrite { get { return false; } }
            public override long Length { get { throw new NotSupportedException(); } }
            public override long Position { get { throw new NotSupportedException(); } set { throw new NotSupportedException(); } }
            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }
            public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

            public override int Read(byte[] buffer, int offset, int count)
            {
                int n = (int)Math.Min(count, _end - _pos);
                for (int i = 0; i < n; i++)
                {
                    long p = _pos + i;
                    buffer[offset + i] = (byte)(Mix(_seed + (ulong)(p >> 3)) >> (int)((p & 7) * 8));
                }
                _pos += n;
                return n;
            }
        }
    }

    /// <summary>A value that announces a length it does not have (the source of a value cut short).</summary>
    public sealed class LyingBlob : FxBlob
    {
        readonly long _announced, _actual;
        public LyingBlob(long announced, long actual) { _announced = announced; _actual = actual; }
        public override long Length { get { return _announced; } }
        public override Stream Open(long offset, long length) { return new PatternBlob(_actual, 3).Open(offset, Math.Max(0, Math.Min(length, _actual - offset))); }
    }

    /// <summary>An IProgress that keeps the reports and cancels its token source once BytesDone reaches CancelAtBytes (a UI would marshal the report instead).</summary>
    public sealed class ProgressRecorder : IProgress<PstProgress>
    {
        public readonly List<PstProgress> Reports = new List<PstProgress>();
        public readonly CancellationTokenSource Cts = new CancellationTokenSource();
        public long CancelAtBytes = long.MaxValue;
        public CancellationToken Token { get { return Cts.Token; } }

        public void Report(PstProgress p)
        {
            lock (Reports) Reports.Add(p);
            if (p.BytesDone >= CancelAtBytes) Cts.Cancel();
        }

        public int Count { get { lock (Reports) return Reports.Count; } }
        public PstProgress Last { get { lock (Reports) return Reports.Count == 0 ? null : Reports[Reports.Count - 1]; } }
    }

    public static class Streams
    {
        /// <summary>True if the two sources hold the same bytes (read in pieces).</summary>
        public static bool Equal(IFxStreamSource a, IFxStreamSource b)
        {
            if (a.Length != b.Length) return false;
            var x = new byte[64 * 1024];
            var y = new byte[64 * 1024];
            using (Stream sa = a.OpenRead())
            using (Stream sb = b.OpenRead())
            {
                long left = a.Length;
                while (left > 0)
                {
                    int n = (int)Math.Min(x.Length, left);
                    int ga = 0, gb = 0;
                    while (ga < n) { int r = sa.Read(x, ga, n - ga); if (r <= 0) return false; ga += r; }
                    while (gb < n) { int r = sb.Read(y, gb, n - gb); if (r <= 0) return false; gb += r; }
                    for (int i = 0; i < n; i++) if (x[i] != y[i]) return false;
                    left -= n;
                }
            }
            return true;
        }

        /// <summary>Copies a source into a new file (in pieces) and returns the number of bytes written.</summary>
        public static long CopyToFile(IFxStreamSource source, string path)
        {
            using (Stream s = source.OpenRead())
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16))
            {
                s.CopyTo(fs, 80 * 1024);
                return fs.Length;
            }
        }
    }

    /// <summary>
    /// The body of importItem ({"FolderId":...,"Mode":"create","Data":"base64"}) written by ImportContent from a stream source into a
    /// sink that decodes the base64 and hashes it: nothing is kept, the hash must be that of the source and the length the announced one.
    /// </summary>
    public static class ImportBody
    {
        public static string Check(IFxStreamSource source, string folderId)
        {
            var content = new ImportContent(folderId, () => source.OpenRead(), source.Length);
            long announced = content.TotalBytes;
            long start = announced - (source.Length + 2) / 3 * 4 - 2;
            using (var sink = new HashSink(start, announced - 2))
            {
                content.CopyToAsync(sink).GetAwaiter().GetResult();
                if (sink.Written != announced) return "the body has " + sink.Written + " bytes, " + announced + " announced";
                byte[] got = sink.Finish(), want;
                using (var sha = SHA256.Create())
                using (Stream s = source.OpenRead()) want = sha.ComputeHash(s);
                if (!got.SequenceEqual(want)) return "the data of the body is not the stream";
            }
            return "";
        }

        sealed class HashSink : Stream
        {
            readonly long _start, _end;
            readonly IncrementalHash _hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            byte[] _carry = new byte[0];
            long _pos;
            public long Written { get { return _pos; } }

            public HashSink(long start, long end) { _start = start; _end = end; }

            public byte[] Finish()
            {
                var rest = new byte[3];
                int consumed, written;
                if (_carry.Length > 0)
                {
                    if (Base64.DecodeFromUtf8(_carry, rest, out consumed, out written, true) != System.Buffers.OperationStatus.Done) throw new InvalidDataException("bad base64");
                    _hash.AppendData(rest, 0, written);
                }
                return _hash.GetHashAndReset();
            }

            public override bool CanRead { get { return false; } }
            public override bool CanSeek { get { return false; } }
            public override bool CanWrite { get { return true; } }
            public override long Length { get { return _pos; } }
            public override long Position { get { return _pos; } set { throw new NotSupportedException(); } }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }

            public override void Write(byte[] buffer, int offset, int count)
            {
                long from = Math.Max(_pos, _start), to = Math.Min(_pos + count, _end);
                if (to > from)
                {
                    int n = (int)(to - from);
                    var all = new byte[_carry.Length + n];
                    Buffer.BlockCopy(_carry, 0, all, 0, _carry.Length);
                    Buffer.BlockCopy(buffer, offset + (int)(from - _pos), all, _carry.Length, n);
                    int usable = all.Length / 4 * 4;
                    var decoded = new byte[usable / 4 * 3];
                    int consumed, written;
                    bool last = to == _end;      // the last block may end with padding
                    if (last) { usable = all.Length; decoded = new byte[(usable / 4 + 1) * 3]; }
                    if (Base64.DecodeFromUtf8(new ReadOnlySpan<byte>(all, 0, usable), decoded, out consumed, out written, last) != System.Buffers.OperationStatus.Done) throw new InvalidDataException("bad base64");
                    _hash.AppendData(decoded, 0, written);
                    _carry = new byte[all.Length - usable];
                    Buffer.BlockCopy(all, usable, _carry, 0, _carry.Length);
                }
                _pos += count;
            }

            protected override void Dispose(bool disposing) { if (disposing) _hash.Dispose(); base.Dispose(disposing); }
        }
    }

    /// <summary>Samples the working set and the private bytes of the process every 10 ms between its creation and Stop.</summary>
    public sealed class MemoryProbe : IDisposable
    {
        readonly Process _p = Process.GetCurrentProcess();
        readonly Thread _t;
        volatile bool _stop;
        public long BaseWorkingSet, BasePrivate, PeakWorkingSet, PeakPrivate;

        public MemoryProbe()
        {
            GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
            _p.Refresh();
            BaseWorkingSet = PeakWorkingSet = _p.WorkingSet64;
            BasePrivate = PeakPrivate = _p.PrivateMemorySize64;
            _t = new Thread(Run) { IsBackground = true };
            _t.Start();
        }

        void Sample()
        {
            _p.Refresh();
            if (_p.WorkingSet64 > PeakWorkingSet) PeakWorkingSet = _p.WorkingSet64;
            if (_p.PrivateMemorySize64 > PeakPrivate) PeakPrivate = _p.PrivateMemorySize64;
        }

        void Run() { while (!_stop) { Sample(); Thread.Sleep(10); } }

        public long WorkingSetGrowth { get { return PeakWorkingSet - BaseWorkingSet; } }
        public long PrivateGrowth { get { return PeakPrivate - BasePrivate; } }

        public void Stop()
        {
            if (_stop) return;
            _stop = true;
            _t.Join();
            Sample();
        }

        public void Dispose() { Stop(); }
    }
}
