// =============================================================================
//  Mailbox Transfer - engine, part 5c: MS-PST reader, NDB and LTP layers (Unicode PST; NDB_CRYPT_NONE, PERMUTE and CYCLIC)
// -----------------------------------------------------------------------------
//  PstFile         header, pages, BBT / NBT lookup and traversal, blocks (trailer, signature, CRC checked), data trees, subnode trees
//  PstNodeView     a node or a subnode: its data blocks, its subnodes
//  PstDataBlob     the data tree of a node as an FxBlob (Engine.FxStream.cs): read block by block, never whole, whatever its size
//  PstHeapReader   Heap-on-Node;  PstBthReader  BTree-on-Heap;  PstPcReader  Property Context;  PstTcReader  Table Context
//
//  Written from the specification, not from the writer: every offset, count and size read from the file is checked before use
//  and a violation raises PstFormatException (with the file offset, the BID or the NID when known); B-tree walks are bounded in
//  depth (no cycle can make them loop) and data trees are bounded in size.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public sealed class PstFormatException : Exception
    {
        public long Offset = -1;
        public ulong Bid;
        public uint Nid;
        public PstFormatException(string message) : base(message) { }
        public PstFormatException(string message, long offset) : base(message + " (file offset 0x" + offset.ToString("X") + ")") { Offset = offset; }
    }

    public static class PstCrypt
    {
        // mpbbR, mpbbS, mpbbI of MS-PST 5.1 (768 bytes)
        const string Table =
            "QTYTYqghbrv0FswEf2ToXR7yyyp0xV410pVHnpYtmohMfYQ/26wxtkhf9sTYOYvnIzs4jsjB3yWxIKVGYE6c+6rTVlFFfFUAB8kr" +
            "nYWbCaCPrbMPY6uJS9enFVpxZkK/JkprmPrqd1OycAUs/Vk6hn7OBuuCeFfHjUOvtBzUW83i6SdPwwhygM+w7/Uobb4wTTSS1Q48" +
            "IjLl5PmfwtEKgRLh7pGDduOX5mGKF3mkt9yQelyMAqbKad5QGhGTuVKHWPztHTdJG2rgKTOZvWzZlPNAVG/wxnO41j5lGEQf3WcQ" +
            "8QwZ7K4DoRR7qQv/+KPAogH3LrwkaHUN/rovtdDaPRRTD1azyHqc62VIFxYVnwLMVHyDAA0MC6JiqHbb2e3HxaTcrIV01tCnm66a" +
            "lnFmw2OZuN1zko6EfaVe0V2TsVdRUICJUpRPTgprvI1/bkdGQUBEARHLAz/39OGpjzw6+fvwGTCCCS7JnaCGSe5vTW3ELYE0JYcb" +
            "iKr8BqESOP1MQnJkEzckanV3Q//mtEs2XOTYNT1FuSzstzErKQdoow5pexieITm+KBpbePUjyiqwrz7+BIzn5ZgyldP2Suim6unz" +
            "1S9wIPIfBWetVRDOzeMnO9q618Im1JEd0hwiM/j68Vrvz5C2i7W9wL8Ilx5s4mHgxsFZq7tY3l/fYHl+sopH8bTmC2pySIVOnuvi" +
            "+JRT4LugAuhaCavb47rGfMMQ3TkFljD1N2CCjMkTSmsd8/uPJpfKkRcBxDItbjGV/9kj0QBeedxEOxooxWFXIJA9g7lDvmfSRkJ2" +
            "wG1bfrIPFik8qQNUDdpd3/a3x2LNjQbTaVyG1hT3pWZ1rLHpRSFwDIefdKQiTG+/H1aqLrN4M1Cwo5K8zxkcp2PLHk0+SxubT+fw" +
            "7q06tVkE6kBVJVHleok4aFJ7/Ceu1736B/TMjl/vNZyEKxXVdzRJthIKf3GI/Z0YQX2T2Fgszv4kr964NsihgKaZmKgvDoFlc+TC" +
            "oorU4RHQCIsq8u2aZD/BbPns";
        static byte[] _t;
        static byte[] T { get { if (_t == null) _t = Convert.FromBase64String(Table); return _t; } }

        /// <summary>Decodes (or encodes) the data of a block in place. method: 1 = permute, 2 = cyclic (key = low 32 bits of the BID).</summary>
        public static void Decode(int method, byte[] b, int off, int len, ulong bid, bool encode)
        {
            byte[] t = T;
            if (method == 1)
            {
                int baseOff = encode ? 0 : 512;
                for (int i = 0; i < len; i++) b[off + i] = t[baseOff + b[off + i]];
            }
            else if (method == 2)
            {
                uint key = (uint)bid;
                ushort w = (ushort)(key ^ (key >> 16));
                for (int i = 0; i < len; i++)
                {
                    byte x = b[off + i];
                    x = (byte)(x + (byte)w);
                    x = t[x];                       // mpbbR
                    x = (byte)(x + (byte)(w >> 8));
                    x = t[256 + x];                 // mpbbS
                    x = (byte)(x - (byte)(w >> 8));
                    x = t[512 + x];                 // mpbbI
                    x = (byte)(x - (byte)w);
                    b[off + i] = x;
                    w = (ushort)(w + 1);
                }
            }
        }
    }

    public struct PstBbtEntry { public ulong Bid, Ib; public int Cb, CRef; }
    public struct PstNbtEntry { public uint Nid; public ulong BidData, BidSub; public uint Parent; }
    public struct PstSubEntry { public uint Nid; public ulong BidData, BidSub; }

    /// <summary>A replacement for the block and node B-trees of a file (built by a scan of the file).</summary>
    public interface IPstIndex
    {
        bool TryFindBlock(ulong bid, out PstBbtEntry e);
        bool TryFindNode(uint nid, out PstNbtEntry e);
        /// <summary>All the nodes in increasing NID order.</summary>
        IEnumerable<PstNbtEntry> Nodes();
    }

    public sealed class PstFile : IDisposable
    {
        readonly FileStream _fs;
        public readonly string Path;
        public long Length;
        public int Version, VersionClient, CryptMethod;
        public ulong BidNextB, BidNextP;
        public uint DwUnique;
        public uint[] Rgnid = new uint[32];
        public long FileEof, AMapLast, AMapFree, PMapFree;
        public PstBref NbtRoot, BbtRoot;
        public byte AMapValid;
        public bool HeaderCrcPartialOk, HeaderCrcFullOk;
        readonly Dictionary<long, byte[]> _pageCache = new Dictionary<long, byte[]>();
        // one FileStream (position + page cache): reads are serialized so that several streams over the file can be used in turn
        readonly object _io = new object();

        PstFile(FileStream fs, string path) { _fs = fs; Path = path; }

        /// <summary>
        /// Replaces the B-trees: when set, blocks and nodes are looked up here (a scan of the file, see PstScan) and the header
        /// roots are not used. Set by the salvage repair only.
        /// </summary>
        public IPstIndex Index;

        /// <summary>
        /// Repair of the name-to-id map only: a block whose CRC does not match is read when exactly one of its bytes is wrong and the CRC
        /// says which byte and what it was (PstCrcRepair); counted in BlocksCorrected. Any other block with a wrong CRC is refused as ever.
        /// </summary>
        public bool CorrectBlocks;
        public int BlocksCorrected;

        /// <summary>
        /// Opens a file for salvage: the header is read as far as it is valid and never refuses the file (a damaged header, a missing
        /// root, a file shorter than a header are allowed); the caller sets Index from a scan. Only the encryption method is taken
        /// from the header (when it looks valid); the scan may change it (CryptMethod is public).
        /// </summary>
        public static PstFile OpenSalvage(string path)
        {
            var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 16, FileOptions.RandomAccess);
            try
            {
                var f = new PstFile(fs, path);
                f.Length = fs.Length;
                f.Version = 23;
                if (f.Length >= PstFormat.HeaderSize)
                {
                    var h = f.ReadAt(0, PstFormat.HeaderSize);
                    if (h[513] <= 2) f.CryptMethod = h[513];
                    f.DwUnique = PstFormat.G32(h, 40);
                    for (int i = 0; i < 32; i++) f.Rgnid[i] = PstFormat.G32(h, 44 + 4 * i);
                }
                return f;
            }
            catch (Exception) { fs.Dispose(); throw; }
        }

        public static PstFile Open(string path)
        {
            var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 16, FileOptions.RandomAccess);
            try
            {
                var f = new PstFile(fs, path);
                f.ReadHeader();
                return f;
            }
            catch (Exception) { fs.Dispose(); throw; }
        }

        void ReadHeader()
        {
            Length = _fs.Length;
            if (Length < PstFormat.HeaderSize) throw new PstFormatException("The file is shorter than a PST header (" + Length + " bytes).");
            var h = ReadAt(0, PstFormat.HeaderSize);
            if (h[0] != 0x21 || h[1] != 0x42 || h[2] != 0x44 || h[3] != 0x4E) throw new PstFormatException("Not a PST file (bad magic).", 0);
            if (h[8] != 0x53 || h[9] != 0x4D) throw new PstFormatException("Bad wMagicClient.", 8);
            Version = PstFormat.G16(h, 10);
            VersionClient = PstFormat.G16(h, 12);
            if (Version == 14 || Version == 15) throw new PstFormatException("ANSI PST files (wVer " + Version + ") are not supported.", 10);
            if (Version < 23) throw new PstFormatException("Unknown PST version " + Version + ".", 10);
            BidNextP = PstFormat.G64(h, 32);
            DwUnique = PstFormat.G32(h, 40);
            for (int i = 0; i < 32; i++) Rgnid[i] = PstFormat.G32(h, 44 + 4 * i);
            FileEof = (long)PstFormat.G64(h, 184);
            AMapLast = (long)PstFormat.G64(h, 192);
            AMapFree = (long)PstFormat.G64(h, 200);
            PMapFree = (long)PstFormat.G64(h, 208);
            NbtRoot = new PstBref { Bid = PstFormat.G64(h, 216), Ib = PstFormat.G64(h, 224) };
            BbtRoot = new PstBref { Bid = PstFormat.G64(h, 232), Ib = PstFormat.G64(h, 240) };
            AMapValid = h[248];
            CryptMethod = h[513];
            BidNextB = PstFormat.G64(h, 516);
            HeaderCrcPartialOk = PstFormat.Crc(h, 8, 471) == PstFormat.G32(h, 4);
            HeaderCrcFullOk = PstFormat.Crc(h, 8, 516) == PstFormat.G32(h, 524);
            if (!HeaderCrcPartialOk) throw new PstFormatException("Header CRC (dwCRCPartial) mismatch.", 4);
            if (!HeaderCrcFullOk) throw new PstFormatException("Header CRC (dwCRCFull) mismatch.", 524);
            if (h[512] != 0x80) throw new PstFormatException("Bad bSentinel.", 512);
            if (CryptMethod != 0 && CryptMethod != 1 && CryptMethod != 2) throw new PstFormatException("Unsupported bCryptMethod " + CryptMethod + ".", 513);
        }

        public byte[] ReadAt(long ib, int len)
        {
            if (ib < 0 || len < 0 || ib + len > Length) throw new PstFormatException("Read of " + len + " bytes beyond the end of the file", ib);
            var b = new byte[len];
            ReadAtInto(ib, b, len);
            return b;
        }

        /// <summary>Reads len bytes at ib into the start of dst (no allocation).</summary>
        public void ReadAtInto(long ib, byte[] dst, int len)
        {
            if (ib < 0 || len < 0 || ib + len > Length) throw new PstFormatException("Read of " + len + " bytes beyond the end of the file", ib);
            lock (_io)
            {
                _fs.Position = ib;
                int got = 0;
                while (got < len)
                {
                    int n = _fs.Read(dst, got, len - got);
                    if (n <= 0) throw new PstFormatException("Unexpected end of file", ib + got);
                    got += n;
                }
            }
        }

        // ---------------------------------------------------------------- pages and B-trees
        /// <summary>Reads a 512-byte page, checks ptype, signature, CRC and (when given) the BID of the trailer. Returns the 512 bytes.</summary>
        public byte[] ReadPage(long ib, byte ptype, ulong expectedBid)
        {
            byte[] p;
            lock (_io)
            if (!_pageCache.TryGetValue(ib, out p))
            {
                if (ib % PstFormat.PageSize != 0) throw new PstFormatException("Page offset is not aligned on 512 bytes", ib);
                p = ReadAt(ib, PstFormat.PageSize);
                if (p[496] != ptype || p[497] != ptype) throw new PstFormatException("Unexpected page type 0x" + p[496].ToString("X2") + " (wanted 0x" + ptype.ToString("X2") + ")", ib);
                if (PstFormat.Crc(p, 0, 496) != PstFormat.G32(p, 500)) throw new PstFormatException("Page CRC mismatch", ib);
                ulong bid = PstFormat.G64(p, 504);
                if (PstFormat.G16(p, 498) != PstFormat.Sig((ulong)ib, bid)) throw new PstFormatException("Page signature mismatch", ib);
                if (_pageCache.Count > 2048) _pageCache.Clear();
                _pageCache[ib] = p;
            }
            if (expectedBid != 0 && PstFormat.G64(p, 504) != expectedBid) throw new PstFormatException("Page BID 0x" + PstFormat.G64(p, 504).ToString("X") + " does not match the reference 0x" + expectedBid.ToString("X"), ib);
            return p;
        }

        static void CheckBtPage(byte[] p, long ib, int expectedLevel, int leafEntrySize)
        {
            int cEnt = p[488], cEntMax = p[489], cbEnt = p[490], level = p[491];
            if (expectedLevel >= 0 && level != expectedLevel) throw new PstFormatException("B-tree page level " + level + " (wanted " + expectedLevel + ")", ib);
            int want = level == 0 ? leafEntrySize : 24;
            if (cbEnt < want) throw new PstFormatException("B-tree page entry size " + cbEnt + " (at least " + want + ")", ib);
            if (cEnt > cEntMax || cEnt * cbEnt > 488) throw new PstFormatException("B-tree page entry count " + cEnt + " does not fit", ib);
        }

        T Search<T>(PstBref root, byte ptype, int leafSize, ulong key, Func<byte[], int, T> leaf, out bool found)
        {
            found = false;
            PstBref cur = root;
            int expect = -1;
            for (int depth = 0; depth < 9; depth++)
            {
                byte[] p = ReadPage((long)cur.Ib, ptype, cur.Bid);
                CheckBtPage(p, (long)cur.Ib, expect, leafSize);
                int level = p[491], cEnt = p[488], cb = p[490];
                if (level == 0)
                {
                    int lo = 0, hi = cEnt - 1;
                    while (lo <= hi)
                    {
                        int mid = (lo + hi) >> 1;
                        ulong k = PstFormat.G64(p, mid * cb);
                        if (k == key) { found = true; return leaf(p, mid * cb); }
                        if (k < key) lo = mid + 1; else hi = mid - 1;
                    }
                    return default(T);
                }
                if (cEnt == 0) return default(T);
                int idx = -1;
                for (int i = 0; i < cEnt; i++) { if (PstFormat.G64(p, i * cb) <= key) idx = i; else break; }
                if (idx < 0) return default(T);
                cur = new PstBref { Bid = PstFormat.G64(p, idx * cb + 8), Ib = PstFormat.G64(p, idx * cb + 16) };
                expect = level - 1;
            }
            throw new PstFormatException("B-tree deeper than 8 levels (cycle?)");
        }

        public bool TryFindBlock(ulong bid, out PstBbtEntry e)
        {
            bid &= ~1UL;           // bit 0 of a BID is reserved: readers ignore it (2.2.2.2)
            if (Index != null) return Index.TryFindBlock(bid, out e);
            bool found;
            e = Search<PstBbtEntry>(BbtRoot, PstFormat.PtypeBBT, 24, bid,
                (p, o) => new PstBbtEntry { Bid = PstFormat.G64(p, o), Ib = PstFormat.G64(p, o + 8), Cb = PstFormat.G16(p, o + 16), CRef = PstFormat.G16(p, o + 18) }, out found);
            return found;
        }

        public bool TryFindNode(uint nid, out PstNbtEntry e)
        {
            if (Index != null) return Index.TryFindNode(nid, out e);
            bool found;
            e = Search<PstNbtEntry>(NbtRoot, PstFormat.PtypeNBT, 32, nid,
                (p, o) => new PstNbtEntry { Nid = (uint)PstFormat.G64(p, o), BidData = PstFormat.G64(p, o + 8), BidSub = PstFormat.G64(p, o + 16), Parent = PstFormat.G32(p, o + 24) }, out found);
            return found;
        }

        /// <summary>All the nodes of the NBT in key order (the order of the keys, the levels and the BREFs are checked).</summary>
        public IEnumerable<PstNbtEntry> EnumerateNodes()
        {
            if (Index != null) { foreach (PstNbtEntry n in Index.Nodes()) yield return n; yield break; }
            var stack = new Stack<KeyValuePair<PstBref, int>>();
            stack.Push(new KeyValuePair<PstBref, int>(NbtRoot, -1));
            long last = -1;
            while (stack.Count > 0)
            {
                var top = stack.Pop();
                byte[] p = ReadPage((long)top.Key.Ib, PstFormat.PtypeNBT, top.Key.Bid);
                CheckBtPage(p, (long)top.Key.Ib, top.Value, 32);
                int level = p[491], cEnt = p[488], cb = p[490];
                if (level == 0)
                {
                    for (int i = 0; i < cEnt; i++)
                    {
                        int o = i * cb;
                        ulong k = PstFormat.G64(p, o);
                        if (k > uint.MaxValue || (long)k <= last) throw new PstFormatException("NBT keys are not strictly increasing near NID 0x" + k.ToString("X"), (long)top.Key.Ib);
                        last = (long)k;
                        yield return new PstNbtEntry { Nid = (uint)k, BidData = PstFormat.G64(p, o + 8), BidSub = PstFormat.G64(p, o + 16), Parent = PstFormat.G32(p, o + 24) };
                    }
                }
                else
                {
                    if (level > 8) throw new PstFormatException("NBT level " + level + " is too deep", (long)top.Key.Ib);
                    for (int i = cEnt - 1; i >= 0; i--)
                        stack.Push(new KeyValuePair<PstBref, int>(new PstBref { Bid = PstFormat.G64(p, i * cb + 8), Ib = PstFormat.G64(p, i * cb + 16) }, level - 1));
                }
            }
        }

        // ---------------------------------------------------------------- blocks
        public byte[] ReadBlock(ulong bid)
        {
            PstBbtEntry e;
            if (!TryFindBlock(bid, out e)) { var x = new PstFormatException("BID 0x" + bid.ToString("X") + " is not in the block B-tree"); x.Bid = bid; throw x; }
            return ReadBlock(e);
        }

        static int BlockTotal(PstBbtEntry e)
        {
            int total = PstFormat.BlockSize(e.Cb);
            if (e.Cb > PstFormat.MaxBlockData || total > PstFormat.MaxBlock) throw new PstFormatException("Block 0x" + e.Bid.ToString("X") + " has an impossible size " + e.Cb, (long)e.Ib);
            if ((long)e.Ib % 64 != 0) throw new PstFormatException("Block 0x" + e.Bid.ToString("X") + " is not aligned on 64 bytes", (long)e.Ib);
            return total;
        }

        public byte[] ReadBlock(PstBbtEntry e)
        {
            BlockTotal(e);
            var data = new byte[e.Cb];
            ReadBlockInto(e, data);
            return data;
        }

        /// <summary>Reads and checks a block, puts its data (decoded) at the start of dest (at least e.Cb bytes) and returns the data length. No allocation.</summary>
        public int ReadBlockInto(PstBbtEntry e, byte[] dest)
        {
            int total = BlockTotal(e);
            byte[] raw = System.Buffers.ArrayPool<byte>.Shared.Rent(total);
            try
            {
                ReadAtInto((long)e.Ib, raw, total);
                int t = total - PstFormat.BlockTrailerSize;
                int cb = PstFormat.G16(raw, t);
                if (cb != e.Cb) { var x = new PstFormatException("Block 0x" + e.Bid.ToString("X") + ": trailer cb " + cb + " differs from the BBT (" + e.Cb + ")", (long)e.Ib); x.Bid = e.Bid; throw x; }
                if ((PstFormat.G64(raw, t + 8) & ~1UL) != (e.Bid & ~1UL)) { var x = new PstFormatException("Block 0x" + e.Bid.ToString("X") + ": trailer BID is 0x" + PstFormat.G64(raw, t + 8).ToString("X"), (long)e.Ib); x.Bid = e.Bid; throw x; }
                if (PstFormat.G16(raw, t + 2) != PstFormat.Sig(e.Ib, PstFormat.G64(raw, t + 8))) { var x = new PstFormatException("Block 0x" + e.Bid.ToString("X") + ": signature mismatch", (long)e.Ib); x.Bid = e.Bid; throw x; }
                if (PstFormat.Crc(raw, 0, cb) != PstFormat.G32(raw, t + 4))
                {
                    if (!CorrectBlocks || !PstCrcRepair.TrySingleByte(raw, cb, PstFormat.G32(raw, t + 4))) { var x = new PstFormatException("Block 0x" + e.Bid.ToString("X") + ": CRC mismatch", (long)e.Ib); x.Bid = e.Bid; throw x; }
                    BlocksCorrected++;
                }
                Buffer.BlockCopy(raw, 0, dest, 0, cb);
                if (CryptMethod != 0 && (e.Bid & 2) == 0) PstCrypt.Decode(CryptMethod, dest, 0, cb, e.Bid, false);
                return cb;
            }
            finally { System.Buffers.ArrayPool<byte>.Shared.Return(raw); }
        }

        /// <summary>Length of the data of a data tree (the total of its root, or the size of its single block) without reading the data blocks.</summary>
        public long DataTreeLength(ulong bid)
        {
            if (bid == 0) return 0;
            if ((bid & 2) == 0)
            {
                PstBbtEntry e;
                if (!TryFindBlock(bid, out e)) { var x = new PstFormatException("BID 0x" + bid.ToString("X") + " is not in the block B-tree"); x.Bid = bid; throw x; }
                return e.Cb;
            }
            byte[] b = ReadBlock(bid);
            if (b.Length < 8 || b[0] != 1) throw new PstFormatException("Block 0x" + bid.ToString("X") + " is not an XBLOCK / XXBLOCK");
            return PstFormat.G32(b, 4);
        }

        /// <summary>Leaf data blocks of a data tree, in order (a single data block gives one entry).</summary>
        public List<byte[]> ReadDataTreeBlocks(ulong bid)
        {
            var res = new List<byte[]>();
            if (bid == 0) return res;
            CollectLeaves(bid, 2, res);
            return res;
        }

        void CollectLeaves(ulong bid, int depthLeft, List<byte[]> into)
        {
            byte[] b = ReadBlock(bid);
            if ((bid & 2) == 0) { into.Add(b); return; }
            if (b.Length < 8 || b[0] != 1) throw new PstFormatException("Block 0x" + bid.ToString("X") + " is not an XBLOCK / XXBLOCK");
            int level = b[1], cEnt = PstFormat.G16(b, 2);
            if (level < 1 || level > 2 || level > depthLeft) throw new PstFormatException("Data tree level " + level + " of block 0x" + bid.ToString("X") + " is not valid here");
            if (8 + 8 * cEnt > b.Length) throw new PstFormatException("XBLOCK 0x" + bid.ToString("X") + " has " + cEnt + " entries but " + b.Length + " bytes");
            int before = into.Count;
            for (int i = 0; i < cEnt; i++)
            {
                ulong child = PstFormat.G64(b, 8 + 8 * i);
                if (level == 1)
                {
                    if ((child & 2) != 0) throw new PstFormatException("XBLOCK 0x" + bid.ToString("X") + " refers to an internal block");
                    into.Add(ReadBlock(child));
                }
                else
                {
                    if ((child & 2) == 0) throw new PstFormatException("XXBLOCK 0x" + bid.ToString("X") + " refers to a data block");
                    CollectLeaves(child, 1, into);
                }
            }
            long total = 0;
            for (int k = before; k < into.Count; k++) total += into[k].Length;
            if (total != PstFormat.G32(b, 4)) throw new PstFormatException("XBLOCK 0x" + bid.ToString("X") + ": lcbTotal " + PstFormat.G32(b, 4) + " differs from the data found (" + total + ")");
        }

        public byte[] ReadData(ulong bid)
        {
            List<byte[]> blocks = ReadDataTreeBlocks(bid);
            if (blocks.Count == 1) return blocks[0];
            long total = 0;
            foreach (byte[] b in blocks) total += b.Length;
            if (total > int.MaxValue) throw new PstFormatException("Data of more than 2 GB");
            var r = new byte[total];
            int pos = 0;
            foreach (byte[] b in blocks) { Buffer.BlockCopy(b, 0, r, pos, b.Length); pos += b.Length; }
            return r;
        }

        /// <summary>Subnodes of a subnode tree (SLBLOCK / SIBLOCK), sorted by NID.</summary>
        public List<PstSubEntry> ReadSubnodes(ulong bidSub)
        {
            var res = new List<PstSubEntry>();
            if (bidSub == 0) return res;
            ReadSl(bidSub, 1, res);
            for (int i = 1; i < res.Count; i++) if (res[i].Nid <= res[i - 1].Nid) throw new PstFormatException("Subnode NIDs are not strictly increasing in block 0x" + bidSub.ToString("X"));
            return res;
        }

        void ReadSl(ulong bid, int depthLeft, List<PstSubEntry> into)
        {
            if ((bid & 2) == 0) throw new PstFormatException("Subnode block 0x" + bid.ToString("X") + " is not internal");
            byte[] b = ReadBlock(bid);
            if (b.Length < 8 || b[0] != 2) throw new PstFormatException("Block 0x" + bid.ToString("X") + " is not an SLBLOCK / SIBLOCK");
            int level = b[1], cEnt = PstFormat.G16(b, 2);
            if (level == 0)
            {
                if (8 + 24 * cEnt > b.Length) throw new PstFormatException("SLBLOCK 0x" + bid.ToString("X") + " has " + cEnt + " entries but " + b.Length + " bytes");
                for (int i = 0; i < cEnt; i++)
                {
                    ulong nid = PstFormat.G64(b, 8 + 24 * i);
                    if (nid > uint.MaxValue) throw new PstFormatException("Subnode NID out of range in block 0x" + bid.ToString("X"));
                    into.Add(new PstSubEntry { Nid = (uint)nid, BidData = PstFormat.G64(b, 16 + 24 * i), BidSub = PstFormat.G64(b, 24 + 24 * i) });
                }
            }
            else if (level == 1 && depthLeft > 0)
            {
                if (8 + 16 * cEnt > b.Length) throw new PstFormatException("SIBLOCK 0x" + bid.ToString("X") + " has " + cEnt + " entries but " + b.Length + " bytes");
                for (int i = 0; i < cEnt; i++) ReadSl(PstFormat.G64(b, 16 + 16 * i), 0, into);
            }
            else throw new PstFormatException("Subnode block 0x" + bid.ToString("X") + " has level " + level);
        }

        public PstNodeView OpenNode(uint nid)
        {
            PstNbtEntry e;
            if (!TryFindNode(nid, out e)) throw new PstFormatException("Node 0x" + nid.ToString("X") + " does not exist") { Nid = nid };
            return new PstNodeView(this, e.Nid, e.BidData, e.BidSub);
        }

        public bool HasNode(uint nid) { PstNbtEntry e; return TryFindNode(nid, out e); }

        public void Dispose() { _fs.Dispose(); }
    }

    /// <summary>A node (top level) or a subnode: data tree and subnode tree.</summary>
    public sealed class PstNodeView
    {
        public readonly PstFile File;
        public readonly uint Nid;
        public readonly ulong BidData, BidSub;
        Dictionary<uint, PstSubEntry> _subs;
        List<byte[]> _blocks;

        public PstNodeView(PstFile f, uint nid, ulong bidData, ulong bidSub) { File = f; Nid = nid; BidData = bidData; BidSub = bidSub; }

        public List<byte[]> Blocks { get { if (_blocks == null) _blocks = File.ReadDataTreeBlocks(BidData); return _blocks; } }
        public byte[] ReadAll() { return File.ReadData(BidData); }
        /// <summary>Bytes of the data of the node, from the root of its data tree (nothing is read from the data blocks).</summary>
        public long DataLength { get { return File.DataTreeLength(BidData); } }
        /// <summary>The data of the node as a blob: a stream that reads the data tree block by block, whatever its size.</summary>
        public FxBlob OpenData() { return new PstDataBlob(File, BidData, DataLength); }

        Dictionary<uint, PstSubEntry> Subs
        {
            get
            {
                if (_subs == null)
                {
                    _subs = new Dictionary<uint, PstSubEntry>();
                    foreach (PstSubEntry s in File.ReadSubnodes(BidSub)) _subs[s.Nid] = s;
                }
                return _subs;
            }
        }

        public IEnumerable<uint> SubnodeNids { get { return Subs.Keys; } }
        public bool HasSubnode(uint nid) { return Subs.ContainsKey(nid); }

        public PstNodeView OpenSubnode(uint nid)
        {
            PstSubEntry s;
            if (!Subs.TryGetValue(nid, out s)) throw new PstFormatException("Subnode 0x" + nid.ToString("X") + " of node 0x" + Nid.ToString("X") + " does not exist") { Nid = nid };
            return new PstNodeView(File, s.Nid, s.BidData, s.BidSub);
        }
    }

    public sealed class PstHeapReader
    {
        readonly List<byte[]> _blocks;
        readonly int[][] _offs;
        public readonly byte ClientSig;
        public readonly uint UserRoot;

        public PstHeapReader(List<byte[]> blocks)
        {
            if (blocks == null || blocks.Count == 0 || blocks[0].Length < 12) throw new PstFormatException("The heap has no data");
            _blocks = blocks;
            _offs = new int[blocks.Count][];
            byte[] b = blocks[0];
            if (b[2] != 0xEC) throw new PstFormatException("Heap signature is 0x" + b[2].ToString("X2") + " (wanted 0xEC)");
            ClientSig = b[3];
            UserRoot = PstFormat.G32(b, 4);
        }

        public int BlockCount { get { return _blocks.Count; } }

        static int HeaderSize(int bi) { return bi == 0 ? 12 : (bi >= 8 && (bi - 8) % 128 == 0) ? 66 : 2; }

        int[] Offsets(int bi)
        {
            if (_offs[bi] != null) return _offs[bi];
            byte[] b = _blocks[bi];
            if (b.Length < 4) throw new PstFormatException("Heap block " + bi + " is too short");
            int ibHnpm = PstFormat.G16(b, 0);
            if (ibHnpm < 2 || ibHnpm + 4 > b.Length) throw new PstFormatException("Heap block " + bi + ": page map offset " + ibHnpm + " is outside the block");
            int cAlloc = PstFormat.G16(b, ibHnpm);
            if (ibHnpm + 4 + 2 * (cAlloc + 1) > b.Length) throw new PstFormatException("Heap block " + bi + ": page map with " + cAlloc + " allocations does not fit");
            var o = new int[cAlloc + 1];
            int hdr = HeaderSize(bi);
            for (int i = 0; i <= cAlloc; i++)
            {
                o[i] = PstFormat.G16(b, ibHnpm + 4 + 2 * i);
                if (o[i] > ibHnpm || o[i] < hdr || (i > 0 && o[i] < o[i - 1])) throw new PstFormatException("Heap block " + bi + ": allocation offset " + o[i] + " is out of order or out of range");
            }
            _offs[bi] = o;
            return o;
        }

        public int AllocCount(int block) { return Offsets(block).Length - 1; }

        public byte[] Get(uint hid)
        {
            if ((hid & 0x1F) != 0) throw new PstFormatException("HID 0x" + hid.ToString("X") + " has a non-zero type");
            int idx = (int)((hid >> 5) & 0x7FF), bi = (int)(hid >> 16);
            if (idx == 0) throw new PstFormatException("HID 0x" + hid.ToString("X") + " has index zero");
            if (bi >= _blocks.Count) throw new PstFormatException("HID 0x" + hid.ToString("X") + " is in block " + bi + " of " + _blocks.Count);
            int[] o = Offsets(bi);
            if (idx > o.Length - 1) throw new PstFormatException("HID 0x" + hid.ToString("X") + ": allocation " + idx + " of " + (o.Length - 1));
            int start = o[idx - 1], end = o[idx];
            var r = new byte[end - start];
            Buffer.BlockCopy(_blocks[bi], start, r, 0, r.Length);
            return r;
        }
    }

    public static class PstBthReader
    {
        /// <summary>All the leaf records (key + data) of a BTH in order.</summary>
        public static List<byte[]> Records(PstHeapReader heap, uint hidHeader, out int cbKey, out int cbEnt)
        {
            byte[] h = heap.Get(hidHeader);
            if (h.Length < 8 || h[0] != 0xB5) throw new PstFormatException("BTHHEADER of HID 0x" + hidHeader.ToString("X") + " is not valid");
            cbKey = h[1]; cbEnt = h[2];
            int levels = h[3];
            uint root = PstFormat.G32(h, 4);
            if (cbKey == 0 || cbEnt == 0) throw new PstFormatException("BTH key / entry sizes are zero");
            if (levels > 8) throw new PstFormatException("BTH with " + levels + " index levels");
            var res = new List<byte[]>();
            if (root == 0) return res;
            Walk(heap, root, levels, cbKey, cbEnt, res);
            return res;
        }

        static void Walk(PstHeapReader heap, uint hid, int level, int cbKey, int cbEnt, List<byte[]> res)
        {
            byte[] n = heap.Get(hid);
            if (level == 0)
            {
                int rec = cbKey + cbEnt;
                if (n.Length % rec != 0) throw new PstFormatException("BTH leaf of " + n.Length + " bytes is not a multiple of the record size " + rec);
                for (int i = 0; i < n.Length; i += rec)
                {
                    var r = new byte[rec];
                    Buffer.BlockCopy(n, i, r, 0, rec);
                    res.Add(r);
                }
            }
            else
            {
                int rec = cbKey + 4;
                if (n.Length % rec != 0) throw new PstFormatException("BTH index node of " + n.Length + " bytes is not a multiple of " + rec);
                for (int i = 0; i < n.Length; i += rec) Walk(heap, PstFormat.G32(n, i + cbKey), level - 1, cbKey, cbEnt, res);
            }
        }
    }

    /// <summary>Resolves HNIDs of a node (heap or subnode).</summary>
    public sealed class PstLtpNode
    {
        public readonly PstNodeView Node;
        public readonly PstHeapReader Heap;
        public PstLtpNode(PstNodeView node) { Node = node; Heap = new PstHeapReader(node.Blocks); }

        public byte[] ReadHnid(uint hnid)
        {
            if (hnid == 0) return new byte[0];
            if ((hnid & 0x1F) == 0) return Heap.Get(hnid);
            return Node.OpenSubnode(hnid).ReadAll();
        }
    }

    public sealed class PstPcReader
    {
        public sealed class Entry { public ushort Id, Type; public uint Hnid; }

        public readonly PstLtpNode Ltp;
        public readonly List<Entry> Entries = new List<Entry>();

        public PstPcReader(PstNodeView node)
        {
            Ltp = new PstLtpNode(node);
            if (Ltp.Heap.ClientSig != 0xBC) throw new PstFormatException("Node 0x" + node.Nid.ToString("X") + " is not a Property Context (client signature 0x" + Ltp.Heap.ClientSig.ToString("X2") + ")") { Nid = node.Nid };
            int cbKey, cbEnt;
            List<byte[]> recs = PstBthReader.Records(Ltp.Heap, Ltp.Heap.UserRoot, out cbKey, out cbEnt);
            if (cbKey != 2 || cbEnt != 6) throw new PstFormatException("PC of node 0x" + node.Nid.ToString("X") + " has a BTH with key " + cbKey + " / entry " + cbEnt + " bytes");
            int prev = -1;
            foreach (byte[] r in recs)
            {
                int id = PstFormat.G16(r, 0);
                if (id <= prev) throw new PstFormatException("PC of node 0x" + node.Nid.ToString("X") + ": property ids are not strictly increasing near 0x" + id.ToString("X4"));
                prev = id;
                Entries.Add(new Entry { Id = (ushort)id, Type = PstFormat.G16(r, 2), Hnid = PstFormat.G32(r, 4) });
            }
        }

        public Entry Find(ushort id)
        {
            int lo = 0, hi = Entries.Count - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                if (Entries[mid].Id == id) return Entries[mid];
                if (Entries[mid].Id < id) lo = mid + 1; else hi = mid - 1;
            }
            return null;
        }

        public PstProp Get(ushort id) { Entry e = Find(id); return e == null ? null : Decode(e); }

        public IEnumerable<PstProp> All() { foreach (Entry e in Entries) yield return Decode(e); }

        /// <summary>
        /// A single variable-size value stored in a subnode and longer than this is not read: the property gets Source (a PstDataBlob)
        /// and no Data. Default: no limit (Data is always filled).
        /// </summary>
        public long LazyThreshold = long.MaxValue;

        public PstProp Decode(Entry e) { return PstValueReader.Decode(Ltp, e.Id, e.Type, e.Hnid, true, LazyThreshold); }
    }

    public static class PstValueReader
    {
        /// <summary>Decodes one property from its type and HNID (inline in a PC; a cell of a TC when pcStyle is false).</summary>
        public static PstProp Decode(PstLtpNode ltp, ushort id, ushort type, uint hnid, bool pcStyle) { return Decode(ltp, id, type, hnid, pcStyle, long.MaxValue); }

        public static PstProp Decode(PstLtpNode ltp, ushort id, ushort type, uint hnid, bool pcStyle, long lazyAbove)
        {
            var p = new PstProp { Id = id, Type = type };
            ushort bt = (ushort)(type & 0x0FFF);
            bool multi = (type & 0x1000) != 0;
            if (!multi && bt == 0x000D)
            {
                byte[] o = ltp.Heap.Get(hnid);
                if (o.Length < 8) throw new PstFormatException("PtypObject record of " + o.Length + " bytes");
                uint nid = PstFormat.G32(o, 0);
                p.ObjectSize = PstFormat.G32(o, 4);
                p.ObjectView = ltp.Node.OpenSubnode(nid);
                return p;
            }
            if (!multi && PstTypes.IsFixedInTc(type))
            {
                int size = PstTypes.FixedSize(bt);
                if (PstTypes.InlineSize(type) > 0)
                {
                    p.Data = new byte[PstTypes.InlineSize(type)];
                    Buffer.BlockCopy(BitConverter.GetBytes(hnid), 0, p.Data, 0, p.Data.Length);
                }
                else
                {
                    byte[] d = ltp.ReadHnid(hnid);
                    if (d.Length < size) throw new PstFormatException("Fixed property 0x" + id.ToString("X4") + type.ToString("X4") + " has " + d.Length + " bytes (wanted " + size + ")");
                    p.Data = d;
                }
                return p;
            }
            if (!multi && lazyAbove < long.MaxValue && (hnid & 0x1F) != 0 && bt != 0x0048)
            {
                PstNodeView sn = ltp.Node.OpenSubnode(hnid);
                long length = sn.DataLength;
                if (length > lazyAbove) { p.Source = new PstDataBlob(ltp.Node.File, sn.BidData, length); return p; }
            }
            byte[] data = ltp.ReadHnid(hnid);
            if (!multi)
            {
                if (bt == 0x0048 && data.Length != 16 && data.Length != 0) throw new PstFormatException("PtypGuid property of " + data.Length + " bytes");
                p.Data = data;
                return p;
            }
            p.Items = UnpackMulti(type, data);
            return p;
        }

        public static List<byte[]> UnpackMulti(ushort type, byte[] d)
        {
            int fs = PstTypes.FixedSize((ushort)(type & 0x0FFF));
            var items = new List<byte[]>();
            if (fs > 0)
            {
                if (d.Length % fs != 0) throw new PstFormatException("Multi-valued property of " + d.Length + " bytes is not a multiple of " + fs);
                for (int i = 0; i < d.Length; i += fs) { var b = new byte[fs]; Buffer.BlockCopy(d, i, b, 0, fs); items.Add(b); }
                return items;
            }
            if (d.Length == 0) return items;
            if (d.Length < 4) throw new PstFormatException("Multi-valued property too short");
            uint count = PstFormat.G32(d, 0);
            if (count > (d.Length - 4) / 4) throw new PstFormatException("Multi-valued property: " + count + " values do not fit in " + d.Length + " bytes");
            for (int i = 0; i < count; i++)
            {
                uint s = PstFormat.G32(d, 4 + 4 * i);
                uint e = i + 1 < count ? PstFormat.G32(d, 8 + 4 * i) : (uint)d.Length;
                if (s > e || e > d.Length || s < 4 + 4 * count) throw new PstFormatException("Multi-valued property: value " + i + " has offsets " + s + ".." + e);
                var b = new byte[e - s];
                Buffer.BlockCopy(d, (int)s, b, 0, b.Length);
                items.Add(b);
            }
            return items;
        }
    }

    /// <summary>The data tree of a node as a value: read block by block (8 KB at a time), from any offset.</summary>
    public sealed class PstDataBlob : FxBlob
    {
        readonly PstFile _file;
        readonly ulong _bid;
        readonly long _length;

        public PstDataBlob(PstFile file, ulong bid, long length) { _file = file; _bid = bid; _length = length; }

        /// <summary>The root of the data tree and the file it is in (the repair reads a damaged tree as far as it goes from them).</summary>
        public ulong Bid { get { return _bid; } }
        public PstFile File { get { return _file; } }

        public override long Length { get { return _length; } }

        public override Stream Open(long offset, long length)
        {
            if (offset < 0 || length < 0 || offset + length > _length) throw new ArgumentOutOfRangeException("length");
            return new PstDataStream(_file, _bid, offset, length);
        }
    }

    /// <summary>
    /// Walks the XBLOCK / XXBLOCK of a data tree and reads one data block at a time into a buffer of 8 KB. Blocks before the offset
    /// are not read (their sizes come from the block B-tree and the totals of the tree blocks); the totals of the tree blocks that
    /// are walked are checked against the sizes of the blocks under them.
    /// </summary>
    sealed class PstDataStream : Stream
    {
        readonly PstFile _f;
        readonly IEnumerator<PstBbtEntry> _leaves;
        readonly byte[] _buf = new byte[PstFormat.MaxBlockData];
        long _skip, _left, _subTotal;
        int _pos, _len;

        public PstDataStream(PstFile f, ulong bid, long offset, long length)
        {
            _f = f; _skip = offset; _left = length;
            _leaves = Leaves(bid, 2).GetEnumerator();
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

        PstBbtEntry Find(ulong bid)
        {
            PstBbtEntry e;
            if (!_f.TryFindBlock(bid, out e)) { var x = new PstFormatException("BID 0x" + bid.ToString("X") + " is not in the block B-tree"); x.Bid = bid; throw x; }
            return e;
        }

        IEnumerable<PstBbtEntry> Leaves(ulong bid, int depthLeft)
        {
            if ((bid & 2) == 0)
            {
                PstBbtEntry single = Find(bid);
                _subTotal = single.Cb;
                if (_skip >= single.Cb) { _skip -= single.Cb; yield break; }
                yield return single;
                yield break;
            }
            byte[] b = _f.ReadBlock(bid);
            if (b.Length < 8 || b[0] != 1) throw new PstFormatException("Block 0x" + bid.ToString("X") + " is not an XBLOCK / XXBLOCK");
            int level = b[1], cEnt = PstFormat.G16(b, 2);
            if (level < 1 || level > 2 || level > depthLeft) throw new PstFormatException("Data tree level " + level + " of block 0x" + bid.ToString("X") + " is not valid here");
            if (8 + 8 * cEnt > b.Length) throw new PstFormatException("XBLOCK 0x" + bid.ToString("X") + " has " + cEnt + " entries but " + b.Length + " bytes");
            long total = PstFormat.G32(b, 4);
            if (_skip >= total) { _skip -= total; _subTotal = total; yield break; }
            long sum = 0;
            for (int i = 0; i < cEnt; i++)
            {
                ulong child = PstFormat.G64(b, 8 + 8 * i);
                if (level == 1)
                {
                    if ((child & 2) != 0) throw new PstFormatException("XBLOCK 0x" + bid.ToString("X") + " refers to an internal block");
                    PstBbtEntry e = Find(child);
                    sum += e.Cb;
                    if (_skip >= e.Cb) { _skip -= e.Cb; continue; }
                    yield return e;
                }
                else
                {
                    if ((child & 2) == 0) throw new PstFormatException("XXBLOCK 0x" + bid.ToString("X") + " refers to a data block");
                    foreach (PstBbtEntry e in Leaves(child, 1)) yield return e;
                    sum += _subTotal;
                }
            }
            if (sum != total) throw new PstFormatException("XBLOCK 0x" + bid.ToString("X") + ": lcbTotal " + total + " differs from the data found (" + sum + ")");
            _subTotal = total;
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            if (_left <= 0 || count == 0) return 0;
            while (_pos >= _len)
            {
                if (!_leaves.MoveNext()) throw new EndOfStreamException("The data tree ends before the value does.");
                _len = _f.ReadBlockInto(_leaves.Current, _buf);
                _pos = (int)_skip;
                _skip = 0;
            }
            int n = (int)Math.Min(Math.Min(count, _len - _pos), _left);
            Buffer.BlockCopy(_buf, _pos, buffer, offset, n);
            _pos += n; _left -= n;
            return n;
        }

        protected override void Dispose(bool disposing) { if (disposing) _leaves.Dispose(); base.Dispose(disposing); }
    }

    public sealed class PstTcReader
    {
        public sealed class Col { public uint Tag; public int IbData, CbData, IBit; public ushort Id { get { return (ushort)(Tag >> 16); } } public ushort Type { get { return (ushort)(Tag & 0xFFFF); } } }

        public readonly PstLtpNode Ltp;
        public readonly List<Col> Columns = new List<Col>();
        public readonly int RowSize;
        public readonly int[] Rgib = new int[4];
        public readonly int RowCount;
        readonly uint[] _rowIds;
        readonly uint _hnidRows;
        List<byte[]> _matrixBlocks;
        byte[] _matrixHeap;
        readonly int _rowsPerBlock;

        public PstTcReader(PstNodeView node)
        {
            Ltp = new PstLtpNode(node);
            if (Ltp.Heap.ClientSig != 0x7C) throw new PstFormatException("Node 0x" + node.Nid.ToString("X") + " is not a Table Context (client signature 0x" + Ltp.Heap.ClientSig.ToString("X2") + ")") { Nid = node.Nid };
            byte[] ti = Ltp.Heap.Get(Ltp.Heap.UserRoot);
            if (ti.Length < 22 || ti[0] != 0x7C) throw new PstFormatException("TCINFO of node 0x" + node.Nid.ToString("X") + " is not valid");
            int cCols = ti[1];
            if (ti.Length < 22 + 8 * cCols) throw new PstFormatException("TCINFO of node 0x" + node.Nid.ToString("X") + ": " + cCols + " column descriptors do not fit");
            for (int i = 0; i < 4; i++) Rgib[i] = PstFormat.G16(ti, 2 + 2 * i);
            RowSize = Rgib[3];
            uint hidIndex = PstFormat.G32(ti, 10);
            _hnidRows = PstFormat.G32(ti, 14);
            if (Rgib[0] > Rgib[1] || Rgib[1] > Rgib[2] || Rgib[2] > Rgib[3]) throw new PstFormatException("TCINFO of node 0x" + node.Nid.ToString("X") + ": rgib is not increasing");
            int ceb = (cCols + 7) / 8;
            if (RowSize < Rgib[2] + ceb) throw new PstFormatException("TCINFO: row size " + RowSize + " is smaller than the data and the CEB");
            for (int i = 0; i < cCols; i++)
            {
                int o = 22 + 8 * i;
                var c = new Col { Tag = PstFormat.G32(ti, o), IbData = PstFormat.G16(ti, o + 4), CbData = ti[o + 6], IBit = ti[o + 7] };
                if (c.IbData + c.CbData > RowSize || c.IBit >= cCols) throw new PstFormatException("TCINFO: column 0x" + c.Tag.ToString("X8") + " lies outside the row");
                Columns.Add(c);
            }
            int cbKey, cbEnt;
            List<byte[]> idx = PstBthReader.Records(Ltp.Heap, hidIndex, out cbKey, out cbEnt);
            if (cbKey != 4 || cbEnt != 4) throw new PstFormatException("TC row index BTH has key " + cbKey + " / entry " + cbEnt + " bytes");
            RowCount = idx.Count;
            _rowIds = new uint[RowCount];
            var seen = new bool[RowCount];
            foreach (byte[] r in idx)
            {
                uint rowIndex = PstFormat.G32(r, 4);
                if (rowIndex >= RowCount || seen[rowIndex]) throw new PstFormatException("TC row index: row number " + rowIndex + " is out of range or repeated");
                seen[rowIndex] = true;
                _rowIds[rowIndex] = PstFormat.G32(r, 0);
            }
            if (RowCount > 0 && RowSize == 0) throw new PstFormatException("TC with rows of size zero");
            _rowsPerBlock = RowSize == 0 ? 1 : PstFormat.MaxBlockData / RowSize;
            if (RowCount > 0 && _rowsPerBlock == 0) throw new PstFormatException("TC row of " + RowSize + " bytes does not fit in a block");
            if (RowCount > 0 && _hnidRows == 0) throw new PstFormatException("TC has rows but no Row Matrix");
        }

        public uint RowId(int row) { return _rowIds[row]; }
        public Col FindColumn(ushort id) { return Columns.FirstOrDefault(c => c.Id == id); }

        byte[] RowBytes(int row, out int offset)
        {
            if (row < 0 || row >= RowCount) throw new ArgumentOutOfRangeException("row");
            if ((_hnidRows & 0x1F) == 0)
            {
                if (_matrixHeap == null) _matrixHeap = Ltp.Heap.Get(_hnidRows);
                offset = row * RowSize;
                if (offset + RowSize > _matrixHeap.Length) throw new PstFormatException("TC Row Matrix of " + _matrixHeap.Length + " bytes is too short for row " + row);
                return _matrixHeap;
            }
            if (_matrixBlocks == null) _matrixBlocks = Ltp.Node.OpenSubnode(_hnidRows).Blocks;
            int bi = row / _rowsPerBlock, ri = row % _rowsPerBlock;
            if (bi >= _matrixBlocks.Count) throw new PstFormatException("TC Row Matrix has " + _matrixBlocks.Count + " blocks, row " + row + " would be in block " + bi);
            offset = ri * RowSize;
            if (offset + RowSize > _matrixBlocks[bi].Length) throw new PstFormatException("TC Row Matrix block " + bi + " is too short for row " + row);
            return _matrixBlocks[bi];
        }

        public bool CellExists(byte[] rowBuf, int rowOff, Col c)
        {
            int cebOff = rowOff + Rgib[2];
            int bit = c.IBit;
            return (rowBuf[cebOff + bit / 8] & (0x80 >> (bit % 8))) != 0;
        }

        /// <summary>The cell of a column in a row, null if the Cell Existence Block says the column does not exist.</summary>
        public PstProp GetCell(int row, Col c)
        {
            int off;
            byte[] buf = RowBytes(row, out off);
            if (!CellExists(buf, off, c)) return null;
            ushort type = c.Type;
            if (PstTypes.IsFixedInTc(type))
            {
                if (c.CbData != PstTypes.CellSize(type)) throw new PstFormatException("TC column 0x" + c.Tag.ToString("X8") + " has cbData " + c.CbData);
                var p = new PstProp { Id = c.Id, Type = type, Data = new byte[c.CbData] };
                Buffer.BlockCopy(buf, off + c.IbData, p.Data, 0, c.CbData);
                return p;
            }
            if (c.CbData != 4) throw new PstFormatException("TC column 0x" + c.Tag.ToString("X8") + " is variable but has cbData " + c.CbData);
            return PstValueReader.Decode(Ltp, c.Id, type, PstFormat.G32(buf, off + c.IbData), false);
        }

        public PstProp GetCell(int row, ushort id)
        {
            Col c = FindColumn(id);
            return c == null ? null : GetCell(row, c);
        }

        public List<PstProp> GetRow(int row)
        {
            var l = new List<PstProp>();
            foreach (Col c in Columns) { if (c.Id == 0x67F2 || c.Id == 0x67F3) continue; PstProp p = GetCell(row, c); if (p != null) l.Add(p); }
            return l;
        }
    }
}
