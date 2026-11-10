// =============================================================================
//  Mailbox Transfer - engine, part 4c: streams of FastTransfer data (values read lazily, messages written as a stream)
// -----------------------------------------------------------------------------
//  The memory used to move an item must not depend on its size (a 1 GiB attachment must not need 1 GiB of RAM, several times).
//  Everything big is therefore handled as a STREAM that can be opened again, never as a byte[]:
//
//    IFxStreamSource   a stream of known length that can be opened (again): Length + OpenRead. A file, an item of the local store
//                      (a range of a pack), the FastTransfer stream of a message built from a PST (FxMessageSource), a value.
//    FxBlob            a VALUE source: length + Open(offset, length). The value of a property that is too big to keep:
//                        FxBytesBlob       bytes in memory (small values, terminators)
//                        FxSegmentBlob     [offset, offset + length) of the FastTransfer stream inside a source, found again when
//                                          needed (FxPayload below)
//                        FxSliceBlob       part of another blob (a string without its terminator)
//                        FxConcatBlob      several blobs one after the other (a string from a PST plus its terminator)
//                        PstDataBlob       a data tree of a PST (Engine.PstFile.cs): read block by block
//
//  FxCodec.Decode(IFxStreamSource) keeps a variable-size single value above FxDecodeOptions.LazyAbove (64 KiB) as an FxBlob in
//  FxProperty.Lazy: the value is read once to follow the stream and forgotten (the lexer skips it), and read again, in pieces, by
//  whoever consumes it (the PST writer, the encoder). FxProperty.Raw still works for such a property: it reads the whole value
//  (and fails above 2 GiB, the limit of a byte[]); code that must stay bounded uses FxProperty.Blob / OpenValue instead.
//
//  FxPayload is the reader's view of the de-enveloped FastTransfer stream: the Graph service cuts it into records of type 2 (31,680
//  bytes), a value may start in one record and end in another, so a value is (logical offset, length) in the concatenation of
//  the payloads and FxPayload turns that into positions in the file with the records seen while decoding (16 bytes per record).
//
//  FxMessageSource turns an FxMessage (whose big values are lazy) into the FastTransfer stream (bare, or inside an envelope), by
//  PULLING: Length is known without reading the values and OpenRead gives a stream that copies them to the reader in pieces
//  (ImportContent writes it as base64 into the request of importItem without holding more than a few buffers).
//
//  Limits: the length of a value in a FastTransfer stream is an unsigned 32-bit integer (less than 4 GiB); the model itself has no
//  other limit (lengths are 64-bit). Exchange limits the size of a message to far less (150 MB by default).
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace MailboxTransfer
{
    /// <summary>
    /// A snapshot of the progress of a long operation of the PST / FastTransfer engine (writing messages, converting a message to a
    /// stream, checking a file). The engine never writes to the console: callers pass an IProgress&lt;PstProgress&gt; (the report may
    /// come from any thread, a UI marshals it) and a CancellationToken. A field that does not apply is 0 or null.
    /// </summary>
    public sealed class PstProgress
    {
        public string Phase;
        /// <summary>Bytes processed so far, and the total when it is known (else 0).</summary>
        public long BytesDone, BytesTotal;
        /// <summary>Items (messages, nodes) processed so far, and the total when it is known (else 0).</summary>
        public long Items, ItemsTotal;
        public string CurrentFolder, CurrentItem;
    }

    /// <summary>A stream of known length that can be opened again, each call giving a new stream at position 0.</summary>
    public interface IFxStreamSource
    {
        long Length { get; }
        Stream OpenRead();
    }

    /// <summary>Ways to get an IFxStreamSource.</summary>
    public static class FxStreamSource
    {
        sealed class FuncSource : IFxStreamSource
        {
            readonly Func<Stream> _open;
            public long Length { get; private set; }
            public FuncSource(Func<Stream> open, long length) { _open = open; Length = length; }
            public Stream OpenRead() { return _open(); }
        }

        /// <summary>A file (opened again, read-only and shared, at each OpenRead).</summary>
        public static IFxStreamSource FromFile(string path)
        {
            string full = System.IO.Path.GetFullPath(path);
            long length = new FileInfo(full).Length;
            return new FuncSource(() => new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16, FileOptions.SequentialScan), length);
        }

        public static IFxStreamSource FromBytes(byte[] bytes)
        {
            return new FuncSource(() => new MemoryStream(bytes, false), bytes.Length);
        }

        /// <summary>Any seekable stream that open can produce again (the stream is disposed by the caller of OpenRead).</summary>
        public static IFxStreamSource FromFunc(Func<Stream> open, long length) { return new FuncSource(open, length); }

        /// <summary>
        /// Copies a source to a stream in pieces. Cancellable between two pieces (OperationCanceledException); progress is reported about
        /// every MiB (BytesDone / BytesTotal). Returns the number of bytes copied.
        /// </summary>
        public static long CopyTo(IFxStreamSource source, Stream dest, CancellationToken ct, IProgress<PstProgress> progress)
        {
            long done = 0, reported = 0, total = source.Length;
            using (Stream s = source.OpenRead())
                FxIo.Copy(s, dest, total, ct, n =>
                {
                    done += n;
                    if (progress != null && (done - reported >= (1 << 20) || done == total))
                    {
                        reported = done;
                        progress.Report(new PstProgress { Phase = "copy", BytesDone = done, BytesTotal = total });
                    }
                });
            return done;
        }

        /// <summary>An exported item of the local store: a range of a pack file.</summary>
        public static IFxStreamSource FromStoreItem(ExportStore store, long itemKey)
        {
            StoreItem item = store.GetItem(itemKey);
            if (item == null || item.State != StoreItemState.Exported || !item.Length.HasValue)
                throw new InvalidOperationException("Store item " + itemKey + " is not exported.");
            return new FuncSource(() => store.OpenItem(itemKey), item.Length.Value);
        }
    }

    /// <summary>The value of a property that is not kept in memory: its length and a way to read it (again).</summary>
    public abstract class FxBlob : IFxStreamSource
    {
        public abstract long Length { get; }

        /// <summary>A forward-only stream of the bytes [offset, offset + length) of the value.</summary>
        public abstract Stream Open(long offset, long length);

        public Stream OpenRead() { return Open(0, Length); }

        public virtual FxBlob Slice(long offset, long length)
        {
            if (offset < 0 || length < 0 || offset + length > Length) throw new ArgumentOutOfRangeException("length");
            if (offset == 0 && length == Length) return this;
            return new FxSliceBlob(this, offset, length);
        }

        /// <summary>The bytes [offset, offset + count) (fewer at the end of the value).</summary>
        public byte[] ReadRange(long offset, int count)
        {
            long n = Math.Max(0, Math.Min(count, Length - offset));
            var b = new byte[n];
            using (Stream s = Open(offset, n)) FxIo.ReadFull(s, b, 0, b.Length);
            return b;
        }

        /// <summary>The whole value in memory (not above 2 GiB, the limit of a byte[]).</summary>
        public byte[] ToArray()
        {
            if (Length > Array.MaxLength) throw new InvalidOperationException("A value of " + Length + " bytes does not fit in a byte[]; read it as a stream.");
            return ReadRange(0, (int)Length);
        }

        public void CopyTo(Stream dest) { CopyTo(dest, CancellationToken.None); }

        public void CopyTo(Stream dest, CancellationToken ct)
        {
            using (Stream s = OpenRead()) FxIo.Copy(s, dest, Length, ct, null);
        }
    }

    static class FxIo
    {
        public const int CopyBuffer = 80 * 1024;

        public static int ReadFull(Stream s, byte[] b, int offset, int count)
        {
            int total = 0;
            while (total < count)
            {
                int n = s.Read(b, offset + total, count - total);
                if (n <= 0) break;
                total += n;
            }
            return total;
        }

        /// <summary>Copies exactly <paramref name="count"/> bytes; the source ending sooner is an error.</summary>
        public static void Copy(Stream from, Stream to, long count) { Copy(from, to, count, CancellationToken.None, null); }

        /// <summary>Same, checking the token between two pieces and telling how many bytes each piece had.</summary>
        public static void Copy(Stream from, Stream to, long count, CancellationToken ct, Action<long> onBytes)
        {
            byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(CopyBuffer);
            try
            {
                long left = count;
                while (left > 0)
                {
                    ct.ThrowIfCancellationRequested();
                    int n = from.Read(buf, 0, (int)Math.Min(buf.Length, left));
                    if (n <= 0) throw new EndOfStreamException("The source ends " + left + " bytes before the announced length.");
                    to.Write(buf, 0, n);
                    left -= n;
                    if (onBytes != null) onBytes(n);
                }
            }
            finally { System.Buffers.ArrayPool<byte>.Shared.Return(buf); }
        }

        public static void Skip(Stream s, long count)
        {
            if (count <= 0) return;
            if (s.CanSeek) { s.Seek(count, SeekOrigin.Current); return; }
            byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(CopyBuffer);
            try
            {
                long left = count;
                while (left > 0)
                {
                    int n = s.Read(buf, 0, (int)Math.Min(buf.Length, left));
                    if (n <= 0) throw new EndOfStreamException("The source ends before the offset asked.");
                    left -= n;
                }
            }
            finally { System.Buffers.ArrayPool<byte>.Shared.Return(buf); }
        }
    }

    /// <summary>A read-only stream of a part of another stream (the other one is disposed with it).</summary>
    sealed class FxLimitedStream : Stream
    {
        readonly Stream _inner;
        long _left;

        public FxLimitedStream(Stream inner, long length) { _inner = inner; _left = length; }

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
            if (_left <= 0 || count == 0) return 0;
            int n = _inner.Read(buffer, offset, (int)Math.Min(count, _left));
            if (n > 0) _left -= n;
            return n;
        }

        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }

    public sealed class FxBytesBlob : FxBlob
    {
        readonly byte[] _b;
        readonly int _off, _len;

        public static readonly FxBytesBlob Terminator = new FxBytesBlob(new byte[2]);

        public FxBytesBlob(byte[] bytes) : this(bytes, 0, bytes.Length) { }
        public FxBytesBlob(byte[] bytes, int offset, int length) { _b = bytes; _off = offset; _len = length; }

        public override long Length { get { return _len; } }
        public override Stream Open(long offset, long length) { return new MemoryStream(_b, _off + (int)offset, (int)length, false); }
    }

    public sealed class FxSliceBlob : FxBlob
    {
        readonly FxBlob _parent;
        readonly long _off, _len;

        public FxSliceBlob(FxBlob parent, long offset, long length) { _parent = parent; _off = offset; _len = length; }

        public override long Length { get { return _len; } }
        public override Stream Open(long offset, long length)
        {
            if (offset < 0 || length < 0 || offset + length > _len) throw new ArgumentOutOfRangeException("length");
            return _parent.Open(_off + offset, length);
        }
    }

    public sealed class FxConcatBlob : FxBlob
    {
        readonly FxBlob[] _parts;
        readonly long _len;

        public FxConcatBlob(params FxBlob[] parts) { _parts = parts; _len = parts.Sum(p => p.Length); }

        public override long Length { get { return _len; } }

        public override Stream Open(long offset, long length)
        {
            if (offset < 0 || length < 0 || offset + length > _len) throw new ArgumentOutOfRangeException("length");
            return new ConcatStream(_parts, offset, length);
        }

        sealed class ConcatStream : Stream
        {
            readonly FxBlob[] _parts;
            int _index;
            long _skip, _left;
            Stream _cur;
            long _curLeft;

            public ConcatStream(FxBlob[] parts, long offset, long length) { _parts = parts; _skip = offset; _left = length; }

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
                while (_left > 0 && count > 0)
                {
                    if (_cur == null)
                    {
                        if (_index >= _parts.Length) return 0;
                        FxBlob p = _parts[_index++];
                        if (_skip >= p.Length) { _skip -= p.Length; continue; }
                        long take = Math.Min(p.Length - _skip, _left);
                        _cur = p.Open(_skip, take);
                        _curLeft = take;
                        _skip = 0;
                    }
                    int n = _cur.Read(buffer, offset, (int)Math.Min(count, _curLeft));
                    if (n <= 0) throw new EndOfStreamException("A part of the value ends before its announced length.");
                    _curLeft -= n; _left -= n;
                    if (_curLeft == 0) { _cur.Dispose(); _cur = null; }
                    return n;
                }
                return 0;
            }

            protected override void Dispose(bool disposing) { if (disposing && _cur != null) { _cur.Dispose(); _cur = null; } base.Dispose(disposing); }
        }
    }

    /// <summary>
    /// A value converted while it is read: a string in a code page to UTF-16LE, without the NULs that end it (the terminator).
    /// The source is read once to know the length of the result.
    /// </summary>
    public sealed class FxTranscodedBlob : FxBlob
    {
        readonly FxBlob _src;
        readonly Encoding _enc;
        readonly long _len;

        public FxTranscodedBlob(FxBlob source, Encoding encoding)
        {
            _src = source; _enc = encoding;
            long chars = 0, pending = 0;
            Decoder d = encoding.GetDecoder();
            var buf = new byte[16 * 1024];
            var cbuf = new char[encoding.GetMaxCharCount(buf.Length) + 1];
            using (Stream s = source.OpenRead())
            {
                int n;
                while ((n = s.Read(buf, 0, buf.Length)) > 0)
                {
                    int got = d.GetChars(buf, 0, n, cbuf, 0, false);
                    for (int i = 0; i < got; i++)
                    {
                        if (cbuf[i] == '\0') pending++;
                        else { chars += pending + 1; pending = 0; }
                    }
                }
            }
            _len = chars * 2;
        }

        public override long Length { get { return _len; } }

        public override Stream Open(long offset, long length)
        {
            Stream s = new TranscodeStream(_src.OpenRead(), _enc);
            if (offset > 0) FxIo.Skip(s, offset);
            return new FxLimitedStream(s, length);
        }

        sealed class TranscodeStream : Stream
        {
            readonly Stream _in;
            readonly Decoder _dec;
            readonly byte[] _buf = new byte[16 * 1024];
            readonly char[] _chars;
            readonly MemoryStream _res = new MemoryStream();
            byte[] _out = new byte[0];
            int _outPos, _outLen;
            long _nuls;
            bool _eof;

            public TranscodeStream(Stream input, Encoding enc)
            {
                _in = input; _dec = enc.GetDecoder();
                _chars = new char[enc.GetMaxCharCount(_buf.Length) + 1];
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

            bool Fill()
            {
                while (!_eof)
                {
                    int n = _in.Read(_buf, 0, _buf.Length);
                    if (n <= 0) { _eof = true; return false; }
                    int got = _dec.GetChars(_buf, 0, n, _chars, 0, false);
                    _res.SetLength(0);
                    for (int i = 0; i < got; i++)
                    {
                        if (_chars[i] == '\0') { _nuls++; continue; }
                        for (; _nuls > 0; _nuls--) { _res.WriteByte(0); _res.WriteByte(0); }
                        _res.WriteByte((byte)_chars[i]); _res.WriteByte((byte)(_chars[i] >> 8));
                    }
                    _out = _res.GetBuffer(); _outLen = (int)_res.Length; _outPos = 0;
                    if (_outLen > 0) return true;
                }
                return false;
            }

            public override int Read(byte[] buffer, int offset, int count)
            {
                if (_outPos >= _outLen && !Fill()) return 0;
                int k = Math.Min(count, _outLen - _outPos);
                Buffer.BlockCopy(_out, _outPos, buffer, offset, k);
                _outPos += k;
                return k;
            }

            protected override void Dispose(bool disposing) { if (disposing) _in.Dispose(); base.Dispose(disposing); }
        }
    }

    /// <summary>
    /// Random access to the FastTransfer stream inside a source: the source is the bare stream, or the Graph envelope (records of
    /// type 2 hold the stream, the other records are left out). A logical offset is an offset in the stream without the envelope.
    /// </summary>
    public sealed class FxPayload
    {
        readonly IFxStreamSource _source;
        readonly List<FxEnvelopeEntry> _entries;
        long[] _start, _raw, _len;
        long _total;
        readonly object _gate = new object();

        public FxPayload(IFxStreamSource source, List<FxEnvelopeEntry> envelopeEntries)
        {
            _source = source; _entries = envelopeEntries;
            if (envelopeEntries == null) _total = source.Length;
        }

        void Index()
        {
            lock (_gate)
            {
                if (_start != null || _entries == null) return;
                var start = new List<long>(); var raw = new List<long>(); var len = new List<long>();
                long at = 0;
                foreach (FxEnvelopeEntry e in _entries)
                {
                    if (!e.IsData || e.Length == 0) continue;
                    start.Add(at); raw.Add(e.RawOffset); len.Add(e.Length);
                    at += e.Length;
                }
                _raw = raw.ToArray(); _len = len.ToArray(); _total = at;
                _start = start.ToArray();
            }
        }

        /// <summary>Length of the FastTransfer stream (the envelope left out).</summary>
        public long Length { get { Index(); return _total; } }

        public Stream Open(long offset, long length)
        {
            Index();
            if (offset < 0 || length < 0 || offset + length > _total) throw new ArgumentOutOfRangeException("length");
            Stream raw = _source.OpenRead();
            if (_entries == null)
            {
                try { FxIo.Skip(raw, offset); }
                catch (Exception) { raw.Dispose(); throw; }
                return new FxLimitedStream(raw, length);
            }
            return new PayloadStream(raw, _start, _raw, _len, offset, length);
        }

        sealed class PayloadStream : Stream
        {
            readonly Stream _raw;
            readonly long[] _start, _rawOff, _len;
            long _pos, _end;
            int _rec;
            long _rawPos = -1;

            public PayloadStream(Stream raw, long[] start, long[] rawOff, long[] len, long offset, long length)
            {
                _raw = raw; _start = start; _rawOff = rawOff; _len = len; _pos = offset; _end = offset + length;
                int i = Array.BinarySearch(start, offset);
                _rec = i >= 0 ? i : ~i - 1;
                if (_rec < 0) _rec = 0;
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

            public override int Read(byte[] buffer, int offset, int count)
            {
                if (_pos >= _end || count == 0) return 0;
                while (_rec < _start.Length && _pos >= _start[_rec] + _len[_rec]) _rec++;
                if (_rec >= _start.Length) throw new EndOfStreamException("The stream ends before the value does.");
                long inRec = _pos - _start[_rec];
                long want = Math.Min(Math.Min(count, _end - _pos), _len[_rec] - inRec);
                long at = _rawOff[_rec] + inRec;
                if (at != _rawPos)
                {
                    if (_rawPos >= 0 && at > _rawPos && !_raw.CanSeek) FxIo.Skip(_raw, at - _rawPos);
                    else _raw.Seek(at, SeekOrigin.Begin);
                }
                int n = _raw.Read(buffer, offset, (int)want);
                if (n <= 0) throw new EndOfStreamException("The stream ends before the value does.");
                _rawPos = at + n;
                _pos += n;
                return n;
            }

            protected override void Dispose(bool disposing) { if (disposing) _raw.Dispose(); base.Dispose(disposing); }
        }
    }

    /// <summary>A value of a decoded stream: [offset, offset + length) of its payload.</summary>
    public sealed class FxSegmentBlob : FxBlob
    {
        readonly FxPayload _payload;
        readonly long _off, _len;

        public FxSegmentBlob(FxPayload payload, long offset, long length) { _payload = payload; _off = offset; _len = length; }

        public override long Length { get { return _len; } }

        public override Stream Open(long offset, long length)
        {
            if (offset < 0 || length < 0 || offset + length > _len) throw new ArgumentOutOfRangeException("length");
            return _payload.Open(_off + offset, length);
        }
    }

    // ---------------------------------------------------------------------------------------------------------------------------
    //  The encoder as a sequence of parts
    // ---------------------------------------------------------------------------------------------------------------------------

    /// <summary>A piece of an encoded stream: bytes in memory, or the value of a property to copy from its source.</summary>
    struct FxPart
    {
        public byte[] Bytes;
        public int Offset, Count;
        public FxBlob Blob;

        public long Length { get { return Blob != null ? Blob.Length : Count; } }

        public void WriteTo(Stream dest, CancellationToken ct, Action<long> onBytes)
        {
            if (Blob == null) { dest.Write(Bytes, Offset, Count); if (onBytes != null) onBytes(Count); return; }
            var counter = dest as FxCountingStream;
            if (counter != null) { counter.Count += Blob.Length; return; }
            using (Stream s = Blob.OpenRead()) FxIo.Copy(s, dest, Blob.Length, ct, onBytes);
        }
    }

    /// <summary>A read-only view of a stream that checks a token at each read (a long read is cancellable).</summary>
    sealed class FxCancellableStream : Stream
    {
        readonly Stream _inner;
        readonly CancellationToken _ct;

        public FxCancellableStream(Stream inner, CancellationToken ct) { _inner = inner; _ct = ct; }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { throw new NotSupportedException(); } }
        public override long Position { get { throw new NotSupportedException(); } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) { _ct.ThrowIfCancellationRequested(); return _inner.Read(buffer, offset, count); }
    }

    /// <summary>A write-only stream that only counts (the values of an FxMessage are not read when the encoder writes into it).</summary>
    sealed class FxCountingStream : Stream
    {
        public long Count;
        public override bool CanRead { get { return false; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return true; } }
        public override long Length { get { return Count; } }
        public override long Position { get { return Count; } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { Count += count; }
    }

    /// <summary>The stream of parts as a readable stream: values are copied from their source in the pieces the reader asks for.</summary>
    sealed class FxPartsStream : Stream
    {
        readonly IEnumerator<FxPart> _parts;
        FxPart _cur;
        bool _have;
        int _bytesPos;
        Stream _blob;
        long _blobLeft;

        public FxPartsStream(IEnumerable<FxPart> parts) { _parts = parts.GetEnumerator(); }

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
            if (count == 0) return 0;
            while (true)
            {
                if (!_have)
                {
                    if (!_parts.MoveNext()) return 0;
                    _cur = _parts.Current; _have = true; _bytesPos = 0;
                    if (_cur.Blob != null) { _blob = _cur.Blob.OpenRead(); _blobLeft = _cur.Blob.Length; }
                }
                if (_cur.Blob == null)
                {
                    int k = Math.Min(count, _cur.Count - _bytesPos);
                    if (k > 0) { Buffer.BlockCopy(_cur.Bytes, _cur.Offset + _bytesPos, buffer, offset, k); _bytesPos += k; }
                    if (_bytesPos >= _cur.Count) _have = false;
                    if (k > 0) return k;
                    continue;
                }
                if (_blobLeft <= 0) { _blob.Dispose(); _blob = null; _have = false; continue; }
                int n = _blob.Read(buffer, offset, (int)Math.Min(count, _blobLeft));
                if (n <= 0) throw new EndOfStreamException("The source of a value ends " + _blobLeft + " bytes before its announced length.");
                _blobLeft -= n;
                return n;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { if (_blob != null) { _blob.Dispose(); _blob = null; } _parts.Dispose(); }
            base.Dispose(disposing);
        }
    }

    /// <summary>The records of an envelope around a payload, produced while the payload is read.</summary>
    sealed class FxEnvelopeReadStream : Stream
    {
        readonly Stream _inner;
        readonly List<FxEnvelopeEntry> _layout;
        int _next;
        byte[] _chunk = new byte[0];
        int _chunkPos;
        long _left;

        public FxEnvelopeReadStream(Stream payload, List<FxEnvelopeEntry> layout) { _inner = payload; _layout = layout; }

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
            if (count == 0) return 0;
            while (true)
            {
                if (_chunkPos < _chunk.Length)
                {
                    int k = Math.Min(count, _chunk.Length - _chunkPos);
                    Buffer.BlockCopy(_chunk, _chunkPos, buffer, offset, k);
                    _chunkPos += k;
                    return k;
                }
                if (_left > 0)
                {
                    int n = _inner.Read(buffer, offset, (int)Math.Min(count, _left));
                    if (n <= 0) throw new EndOfStreamException("The stream is shorter than the data records of the envelope.");
                    _left -= n;
                    return n;
                }
                if (_next >= _layout.Count) return 0;
                FxEnvelopeEntry e = _layout[_next++];
                byte[] payload = e.IsData ? null : e.Payload;
                _chunk = new byte[8 + (payload == null ? 0 : payload.Length)];
                Buffer.BlockCopy(BitConverter.GetBytes(e.Type), 0, _chunk, 0, 4);
                Buffer.BlockCopy(BitConverter.GetBytes((uint)e.Length), 0, _chunk, 4, 4);
                if (payload != null) Buffer.BlockCopy(payload, 0, _chunk, 8, payload.Length);
                _chunkPos = 0;
                if (e.IsData) _left = e.Length;
            }
        }

        protected override void Dispose(bool disposing) { if (disposing) _inner.Dispose(); base.Dispose(disposing); }
    }

    /// <summary>
    /// The FastTransfer stream of a message as an IFxStreamSource, produced on demand: Length is computed without reading the
    /// values of the properties, OpenRead copies them from their sources in pieces. With an envelope (the records of a Graph
    /// export, for instance those of the item the message came from) the stream is cut in records of 31,680 bytes inside it.
    /// </summary>
    public sealed class FxMessageSource : IFxStreamSource
    {
        readonly FxMessage _message;
        readonly FxEnvelope _envelope;
        readonly long _payload;
        readonly List<FxEnvelopeEntry> _layout;
        readonly long _length;

        public FxMessageSource(FxMessage message) : this(message, null) { }

        public FxMessageSource(FxMessage message, FxEnvelope envelope)
        {
            if (message == null) throw new ArgumentNullException("message");
            _message = message; _envelope = envelope;
            _payload = FxCodec.PayloadLength(message);
            if (envelope == null) _length = _payload;
            else
            {
                _layout = FxCodec.BuildLayout(envelope, _payload);
                _length = _layout.Sum(e => 8 + (e.IsData ? e.Length : (e.Payload == null ? 0 : e.Payload.Length)));
            }
        }

        public long Length { get { return _length; } }
        public long PayloadLength { get { return _payload; } }

        public Stream OpenRead()
        {
            Stream payload = new FxPartsStream(FxCodec.Parts(_message));
            return _layout == null ? payload : new FxEnvelopeReadStream(payload, _layout);
        }
    }
}
