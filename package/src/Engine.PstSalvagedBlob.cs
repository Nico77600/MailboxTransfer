// =============================================================================
//  Mailbox Transfer - engine, part 5k: a data tree read as far as it goes (the readable part of a damaged attachment)
// -----------------------------------------------------------------------------
//  A big value (an attachment) is a data tree: an XBLOCK (or an XXBLOCK of XBLOCKs) listing data blocks of 8,176 bytes (the last one
//  holds the rest). One bad block (CRC, signature, missing) makes the ordinary reader refuse the whole value. PstSalvagedBlob keeps the
//  blocks that read and puts zeros where the others were, so that the value has its true length and every readable byte at its place.
//
//  The size of a block that cannot be read is not stored anywhere: it is the total of its tree block minus the blocks that read; when
//  several blocks of one tree block are lost the lost ones are taken to be full (8,176 bytes) except the last one, which gets the rest.
//  That is exact for the trees of this writer and of Outlook (all blocks full but the last). A tree block (XBLOCK / XXBLOCK) that cannot
//  be read loses everything under it; a root that cannot be read loses the value (no length), which is then dropped as before.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace MailboxTransfer
{
    /// <summary>A range of bytes of a value: [Offset, Offset + Length).</summary>
    public sealed class PstByteRange
    {
        public long Offset, Length;
        public override string ToString() { return Offset.ToString("N0") + "-" + (Offset + Length - 1).ToString("N0") + " (" + Length.ToString("N0") + " bytes)"; }
    }

    /// <summary>An attachment kept with unreadable parts (PstRepairOptions.KeepDamagedAttachments): where, and which byte ranges are zeros.</summary>
    public sealed class PstDamagedAttachment
    {
        /// <summary>Where it is: "attachment 2", "attachment 0 embedded message: attachment 1", ...</summary>
        public string Where;
        /// <summary>File name before the suffix, and the name it has now.</summary>
        public string OriginalName, NewName;
        public long Length, BadBytes;
        public readonly List<PstByteRange> BadRanges = new List<PstByteRange>();

        public override string ToString()
        {
            return Where + (OriginalName != null ? " '" + OriginalName + "'" : "") + ": " + (Length - BadBytes).ToString("N0") + " of " + Length.ToString("N0") + " bytes read, " + BadBytes.ToString("N0")
                + " bytes zero-filled in " + BadRanges.Count + " range(s): " + string.Join(", ", BadRanges.Take(8)) + (BadRanges.Count > 8 ? ", ... (" + (BadRanges.Count - 8) + " more)" : "");
        }
    }

    public sealed class PstSalvagedBlob : FxBlob
    {
        struct Leaf { public ulong Bid; public long Length; }      // Bid 0: not readable, zeros

        readonly PstFile _file;
        readonly Leaf[] _leaves;
        readonly long _length;
        public readonly List<PstByteRange> BadRanges = new List<PstByteRange>();
        public long BadBytes;

        PstSalvagedBlob(PstFile f, Leaf[] leaves, long length)
        {
            _file = f; _leaves = leaves; _length = length;
            long pos = 0;
            foreach (Leaf l in leaves)
            {
                if (l.Bid == 0 && l.Length > 0)
                {
                    PstByteRange last = BadRanges.Count > 0 ? BadRanges[BadRanges.Count - 1] : null;
                    if (last != null && last.Offset + last.Length == pos) last.Length += l.Length;
                    else BadRanges.Add(new PstByteRange { Offset = pos, Length = l.Length });
                    BadBytes += l.Length;
                }
                pos += l.Length;
            }
        }

        public override long Length { get { return _length; } }

        /// <summary>
        /// The data tree under rootBid as a blob with zeros in place of the blocks that do not read; null when the tree cannot be walked
        /// (unreadable root, inconsistent totals), when nothing is wrong with it, or when nothing of it can be read.
        /// </summary>
        public static PstSalvagedBlob TryCreate(PstFile f, ulong rootBid, CancellationToken ct)
        {
            var leaves = new List<Leaf>();
            long total;
            var buf = new byte[PstFormat.MaxBlockData];
            try { total = Walk(f, rootBid, 2, leaves, buf, ct); }
            catch (PstFormatException) { return null; }
            long sum = 0, bad = 0;
            foreach (Leaf l in leaves) { sum += l.Length; if (l.Bid == 0) bad += l.Length; }
            if (sum != total || bad == 0 || bad >= total) return null;
            return new PstSalvagedBlob(f, leaves.ToArray(), total);
        }

        static long Walk(PstFile f, ulong bid, int depthLeft, List<Leaf> into, byte[] buf, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            if ((bid & 2) == 0)
            {
                PstBbtEntry e;
                if (!f.TryFindBlock(bid, out e)) throw new PstFormatException("BID 0x" + bid.ToString("X") + " is not in the block B-tree");
                int n = f.ReadBlockInto(e, buf);
                into.Add(new Leaf { Bid = bid, Length = n });
                return n;
            }
            byte[] b = f.ReadBlock(bid);
            if (b.Length < 8 || b[0] != 1) throw new PstFormatException("Block 0x" + bid.ToString("X") + " is not an XBLOCK / XXBLOCK");
            int level = b[1], cEnt = PstFormat.G16(b, 2);
            if (level < 1 || level > 2 || level > depthLeft || 8 + 8 * cEnt > b.Length) throw new PstFormatException("Data tree block 0x" + bid.ToString("X") + " is not valid here");
            long t = PstFormat.G32(b, 4), known = 0;
            int first = into.Count;
            var gaps = new List<int>();
            for (int i = 0; i < cEnt; i++)
            {
                ulong child = PstFormat.G64(b, 8 + 8 * i);
                int mark = into.Count;
                try
                {
                    if ((level == 1) != ((child & 2) == 0)) throw new PstFormatException("Data tree block 0x" + bid.ToString("X") + " has a child of the wrong kind");
                    known += Walk(f, child, level - 1, into, buf, ct);
                }
                catch (PstFormatException)
                {
                    into.RemoveRange(mark, into.Count - mark);
                    gaps.Add(into.Count);
                    into.Add(new Leaf { Bid = 0, Length = level == 1 ? PstFormat.MaxBlockData : (long)PstFormat.XEntriesMax * PstFormat.MaxBlockData });
                }
            }
            long rest = t - known;
            if (rest < 0 || (gaps.Count == 0 && rest != 0)) throw new PstFormatException("Data tree block 0x" + bid.ToString("X") + ": total " + t + " differs from the data found (" + known + ")");
            for (int g = 0; g < gaps.Count; g++)
            {
                Leaf l = into[gaps[g]];
                l.Length = g == gaps.Count - 1 ? rest : Math.Min(l.Length, rest);
                rest -= l.Length;
                into[gaps[g]] = l;
            }
            return t;
        }

        public override Stream Open(long offset, long length)
        {
            if (offset < 0 || length < 0 || offset + length > _length) throw new ArgumentOutOfRangeException("length");
            return new SalvagedStream(this, offset, length);
        }

        sealed class SalvagedStream : Stream
        {
            readonly PstSalvagedBlob _b;
            readonly byte[] _buf = new byte[PstFormat.MaxBlockData];
            int _i = -1, _bufLeaf = -1;
            long _leafStart, _pos;
            readonly long _end;

            public SalvagedStream(PstSalvagedBlob b, long offset, long length) { _b = b; _pos = offset; _end = offset + length; }

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
                if (_i < 0) { _i = 0; _leafStart = 0; }
                Leaf[] leaves = _b._leaves;
                while (_i < leaves.Length && _pos >= _leafStart + leaves[_i].Length) { _leafStart += leaves[_i].Length; _i++; }
                if (_i >= leaves.Length) throw new EndOfStreamException("The data tree ends before the value does.");
                Leaf l = leaves[_i];
                int within = (int)(_pos - _leafStart);
                int n = (int)Math.Min(Math.Min(count, l.Length - within), _end - _pos);
                if (l.Bid == 0) Array.Clear(buffer, offset, n);
                else
                {
                    if (_bufLeaf != _i)
                    {
                        PstBbtEntry e;
                        if (!_b._file.TryFindBlock(l.Bid, out e)) throw new PstFormatException("BID 0x" + l.Bid.ToString("X") + " is not in the block B-tree");
                        _b._file.ReadBlockInto(e, _buf);
                        _bufLeaf = _i;
                    }
                    Buffer.BlockCopy(_buf, within, buffer, offset, n);
                }
                _pos += n;
                return n;
            }
        }
    }

    /// <summary>The name of a damaged attachment: "name.ext" becomes "name (damaged).ext".</summary>
    static class PstDamagedName
    {
        public const string Marker = " (damaged)";

        public static string Mark(string name)
        {
            if (string.IsNullOrEmpty(name)) return "attachment" + Marker;
            int dot = name.LastIndexOf('.');
            string stem = dot > 0 ? name.Substring(0, dot) : name, ext = dot > 0 ? name.Substring(dot) : "";
            if (stem.EndsWith(Marker, StringComparison.Ordinal)) return name;
            return stem + Marker + ext;
        }
    }
}
