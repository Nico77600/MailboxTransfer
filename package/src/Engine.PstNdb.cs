// =============================================================================
//  Mailbox Transfer - engine, part 5a: MS-PST, NDB layer (our own implementation, Unicode PST, bCryptMethod NDB_CRYPT_NONE)
// -----------------------------------------------------------------------------
//  PstFormat       constants, NID / BID helpers, the CRC of MS-PST 5.3, the signature of 5.5
//  PstNode         a node being assembled: data (raw stream or leaf blocks of a heap) + subnodes, written by PstNdbWriter
//  PstNdbWriter    streaming writer of the file: header, AMap / PMap / FMap / FPMap, DList, data blocks, data trees (XBLOCK /
//                  XXBLOCK), subnode trees (SLBLOCK / SIBLOCK), BBT (built while the blocks are written) and NBT (built at Close)
//
//  Memory bound: the file is produced sequentially, one allocation section (253,952 bytes, the span of one AMap) at a time.
//  A block is copied in the section buffer, the section is flushed when it is full. The BBT is built incrementally (BIDs are
//  assigned in increasing order and each block is written right away, so the leaf entries come in key order): one pending page
//  per tree level. The NBT entries (32 bytes per top-level node) are kept in memory and sorted at Close.
//  The only data kept for a whole run is one small record per allocation section (free count, longest free run).
//
//  Layout (MS-PST 1.3.2 / 2.2.2.7): header (0..0x4200), DList page 0x4200, then sections of 253,952 bytes from 0x4400. The first
//  page of each section is its AMap; the next pages hold the PMap (every 8 sections), the FMap (every 496) and the FPMap (every
//  31,744), those three are written when the file is closed because they describe sections that did not exist yet.
//  The file always ends at the end of a section (free space of the last section is marked free in its AMap).
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public static class PstFormat
    {
        public const int PageSize = 512;
        public const int MaxBlock = 8192;
        public const int BlockTrailerSize = 16;
        public const int MaxBlockData = 8176;
        public const long FirstAMapOffset = 0x4400;
        public const long DListOffset = 0x4200;
        public const int SectionSize = 253952;            // 496 bytes * 8 bits * 64 bytes
        public const int HeaderSize = 564;
        public const int MaxHeapAlloc = 3580;
        public const int XEntriesMax = (MaxBlockData - 8) / 8;     // 1021 BIDs in an XBLOCK / XXBLOCK
        public const int SlEntriesMax = (MaxBlockData - 8) / 24;   // 340
        public const int SiEntriesMax = (MaxBlockData - 8) / 16;   // 510

        // page types
        public const byte PtypeBBT = 0x80, PtypeNBT = 0x81, PtypeFMap = 0x82, PtypePMap = 0x83, PtypeAMap = 0x84, PtypeFPMap = 0x85, PtypeDL = 0x86;

        // NID types (2.2.2.1)
        public const uint NtHid = 0, NtInternal = 1, NtNormalFolder = 2, NtSearchFolder = 3, NtNormalMessage = 4, NtAttachment = 5,
            NtSearchUpdateQueue = 6, NtSearchCriteria = 7, NtAssocMessage = 8, NtHierarchyTable = 0x0D, NtContentsTable = 0x0E,
            NtReceiveFolderTable = 0x0B, NtOutgoingQueueTable = 0x0C, NtAssocContentsTable = 0x0F, NtSearchContentsTable = 0x10, NtAttachmentTable = 0x11, NtRecipientTable = 0x12, NtLtp = 0x1F;

        // special NIDs (2.4.1, 2.7.1)
        public const uint NidMessageStore = 0x21, NidNameToIdMap = 0x61, NidRootFolder = 0x122, NidSearchManagementQueue = 0x1E1,
            NidSearchActivityList = 0x201, NidHierarchyTemplate = 0x60D, NidContentsTemplate = 0x60E, NidAssocContentsTemplate = 0x60F,
            NidSearchContentsTemplate = 0x610, NidAttachmentTable = 0x671, NidRecipientTable = 0x692;

        public static uint Nid(uint type, uint index) { return (index << 5) | type; }
        public static uint NidType(uint nid) { return nid & 0x1F; }
        public static uint NidIndex(uint nid) { return nid >> 5; }

        public static ushort Sig(ulong ib, ulong bid)
        {
            uint x = (uint)(ib ^ bid);
            return (ushort)((x >> 16) ^ x);
        }

        // ---- CRC of MS-PST 5.3: reflected polynomial 0xEDB88320, seed 0, no final complement (verified on the sample header of the
        // specification: dwCRCPartial 0x379AA90E, dwCRCFull 0x1FD283D6). Slicing by 8.
        static readonly uint[][] CrcT = BuildCrc();

        static uint[][] BuildCrc()
        {
            var t = new uint[8][];
            for (int k = 0; k < 8; k++) t[k] = new uint[256];
            for (uint i = 0; i < 256; i++)
            {
                uint c = i;
                for (int b = 0; b < 8; b++) c = (c & 1) != 0 ? (c >> 1) ^ 0xEDB88320u : c >> 1;
                t[0][i] = c;
            }
            for (int k = 1; k < 8; k++)
                for (int i = 0; i < 256; i++) t[k][i] = t[0][t[k - 1][i] & 0xFF] ^ (t[k - 1][i] >> 8);
            return t;
        }

        public static uint Crc(byte[] b, int off, int len) { return Crc(b, off, len, 0); }

        public static uint Crc(byte[] b, int off, int len, uint seed)
        {
            uint c = seed;
            uint[] t0 = CrcT[0], t1 = CrcT[1], t2 = CrcT[2], t3 = CrcT[3], t4 = CrcT[4], t5 = CrcT[5], t6 = CrcT[6], t7 = CrcT[7];
            while (len >= 8)
            {
                uint lo = BitConverter.ToUInt32(b, off) ^ c;
                uint hi = BitConverter.ToUInt32(b, off + 4);
                c = t7[lo & 0xFF] ^ t6[(lo >> 8) & 0xFF] ^ t5[(lo >> 16) & 0xFF] ^ t4[lo >> 24]
                  ^ t3[hi & 0xFF] ^ t2[(hi >> 8) & 0xFF] ^ t1[(hi >> 16) & 0xFF] ^ t0[hi >> 24];
                off += 8; len -= 8;
            }
            while (len-- > 0) c = t0[(c ^ b[off++]) & 0xFF] ^ (c >> 8);
            return c;
        }

        // ---- little-endian helpers
        public static void P16(byte[] b, int o, int v) { b[o] = (byte)v; b[o + 1] = (byte)(v >> 8); }
        public static void P32(byte[] b, int o, uint v) { b[o] = (byte)v; b[o + 1] = (byte)(v >> 8); b[o + 2] = (byte)(v >> 16); b[o + 3] = (byte)(v >> 24); }
        public static void P64(byte[] b, int o, ulong v) { P32(b, o, (uint)v); P32(b, o + 4, (uint)(v >> 32)); }
        public static ushort G16(byte[] b, int o) { return (ushort)(b[o] | (b[o + 1] << 8)); }
        public static uint G32(byte[] b, int o) { return (uint)(b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)); }
        public static ulong G64(byte[] b, int o) { return G32(b, o) | ((ulong)G32(b, o + 4) << 32); }

        /// <summary>Size of a block holding cb bytes of data: the data and the 16-byte trailer rounded up to 64.</summary>
        public static int BlockSize(int cb) { return (cb + BlockTrailerSize + 63) / 64 * 64; }

        public static int FillLevel(int freeBytes)
        {
            int[] th = { 3584, 2560, 2048, 1792, 1536, 1280, 1024, 768, 512, 256, 128, 64, 32, 16, 8 };
            for (int i = 0; i < th.Length; i++) if (freeBytes >= th[i]) return i;
            return 15;
        }
    }

    /// <summary>
    /// A node (or subnode) being assembled in memory: its data stream, either raw bytes or the leaf blocks of a heap (each at most
    /// 8,176 bytes), and its subnodes keyed by local NID.
    /// </summary>
    public sealed class PstNode
    {
        public byte[] Raw;
        public List<byte[]> Blocks;
        /// <summary>The data stream read from its source while the node is written (a big value): never held in memory.</summary>
        public FxBlob Source;
        public readonly SortedDictionary<uint, PstNode> Subnodes = new SortedDictionary<uint, PstNode>();
        uint _nextLtp = 1;

        public uint NewLtpNid() { return PstFormat.Nid(PstFormat.NtLtp, _nextLtp++); }

        public long DataLength
        {
            get
            {
                if (Raw != null) return Raw.Length;
                if (Source != null) return Source.Length;
                long n = 0;
                if (Blocks != null) foreach (byte[] b in Blocks) n += b.Length;
                return n;
            }
        }

        /// <summary>Data of the node and of all its subnodes (the "size" of an object).</summary>
        public long TotalLength
        {
            get
            {
                long n = DataLength;
                foreach (PstNode s in Subnodes.Values) n += s.TotalLength;
                return n;
            }
        }
    }

    public struct PstBref { public ulong Bid; public ulong Ib; }

    public sealed class PstNdbWriter : IDisposable
    {
        sealed class Level
        {
            public byte[] Buf = new byte[488];
            public int Count, Max, EntrySize, PagesWritten;
            public ulong FirstKey;
            public byte[] LastEntry;
        }

        // ---- B-tree built bottom-up, one pending page per level ----
        sealed class BTreeBuilder
        {
            readonly PstNdbWriter _w;
            readonly byte _ptype;
            readonly int _leafSize;
            readonly List<Level> _lv = new List<Level>();
            public long Entries;

            public BTreeBuilder(PstNdbWriter w, byte ptype, int leafSize) { _w = w; _ptype = ptype; _leafSize = leafSize; }

            Level Get(int level)
            {
                while (_lv.Count <= level)
                {
                    var l = new Level { EntrySize = _lv.Count == 0 ? _leafSize : 24 };
                    l.Max = 488 / l.EntrySize;
                    _lv.Add(l);
                }
                return _lv[level];
            }

            public void Add(int level, byte[] entry, ulong key)
            {
                Level l = Get(level);
                if (level == 0) Entries++;
                if (l.Count == 0) l.FirstKey = key;
                Buffer.BlockCopy(entry, 0, l.Buf, l.Count * l.EntrySize, l.EntrySize);
                l.LastEntry = entry;
                l.Count++;
                if (l.Count == l.Max) Emit(level);
            }

            PstBref Emit(int level)
            {
                Level l = _lv[level];
                var page = new byte[PstFormat.PageSize];
                Buffer.BlockCopy(l.Buf, 0, page, 0, l.Count * l.EntrySize);
                page[488] = (byte)l.Count; page[489] = (byte)l.Max; page[490] = (byte)l.EntrySize; page[491] = (byte)level;
                PstBref r = _w.WritePage(page, _ptype);
                l.PagesWritten++;
                ulong key = l.FirstKey;
                l.Count = 0;
                Array.Clear(l.Buf, 0, l.Buf.Length);
                var e = new byte[24];
                PstFormat.P64(e, 0, key); PstFormat.P64(e, 8, r.Bid); PstFormat.P64(e, 16, r.Ib);
                if (level + 1 < _lv.Count || true) Add(level + 1, e, key);
                return r;
            }

            public PstBref Finish()
            {
                for (int level = 0; ; level++)
                {
                    Level l = Get(level);
                    bool top = level == _lv.Count - 1;
                    if (top && l.PagesWritten == 0)
                    {
                        if (level > 0 && l.Count == 1)
                            return new PstBref { Bid = PstFormat.G64(l.Buf, 8), Ib = PstFormat.G64(l.Buf, 16) };   // a single child: it is the root
                        var page = new byte[PstFormat.PageSize];
                        Buffer.BlockCopy(l.Buf, 0, page, 0, l.Count * l.EntrySize);
                        page[488] = (byte)l.Count; page[489] = (byte)l.Max; page[490] = (byte)l.EntrySize; page[491] = (byte)level;
                        return _w.WritePage(page, _ptype);
                    }
                    if (l.Count > 0) Emit(level);
                }
            }
        }

        readonly FileStream _fs;
        readonly string _path;
        ulong _bidNextB = 4, _bidNextP = 1;
        uint _dwUnique = 1;
        public readonly uint[] Rgnid = new uint[32];

        int _sec = -1;
        long _secBase;
        readonly byte[] _buf = new byte[PstFormat.SectionSize];
        readonly byte[] _amap = new byte[496];
        int _cursor;
        readonly List<ushort> _secFree = new List<ushort>();
        readonly List<byte> _secRun = new List<byte>();
        readonly List<long> _pmapPages = new List<long>(), _fmapPages = new List<long>(), _fpmapPages = new List<long>();
        long _freeBytes;

        readonly BTreeBuilder _bbt;
        struct NbtRec { public uint Nid, Parent; public ulong BidData, BidSub; }
        readonly List<NbtRec> _nbt = new List<NbtRec>();
        public long BlockCount, PageCount;
        bool _closed;

        public PstNdbWriter(string path)
        {
            _path = path;
            _fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.Read, 1 << 20, FileOptions.None);
            _bbt = new BTreeBuilder(this, PstFormat.PtypeBBT, 24);
            for (int i = 0; i < 32; i++) Rgnid[i] = 0x400;
            Rgnid[(int)PstFormat.NtSearchFolder] = 0x4000;
            Rgnid[(int)PstFormat.NtNormalMessage] = 0x10000;
            Rgnid[(int)PstFormat.NtAssocMessage] = 0x8000;
            _fs.Write(new byte[PstFormat.FirstAMapOffset], 0, (int)PstFormat.FirstAMapOffset);
        }

        public string Path { get { return _path; } }

        /// <summary>Bytes of the file produced so far (complete sections plus the part of the current one in use).</summary>
        public long BytesUsed { get { return _sec < 0 ? PstFormat.FirstAMapOffset : _secBase + _cursor; } }

        /// <summary>Size of the file as it would be at Close, without the B-tree pages still to be written (estimated).</summary>
        public long ProjectedSize
        {
            get
            {
                long pendingNbt = (long)_nbt.Count * 32 * 11 / 10 + 4096;
                long pendingBbt = BlockCount > 0 ? 512 * 2 : 0;
                return BytesUsed + pendingNbt + pendingBbt;
            }
        }

        public uint NextNid(uint type)
        {
            return PstFormat.Nid(type, ++Rgnid[(int)type]);
        }

        // ---------------------------------------------------------------- allocation sections
        void StartSection()
        {
            _sec++;
            _secBase = PstFormat.FirstAMapOffset + (long)_sec * PstFormat.SectionSize;
            Array.Clear(_buf, 0, _buf.Length);
            Array.Clear(_amap, 0, _amap.Length);
            // pages of a section after its AMap (positions observed in Outlook's own validation of a file of 1.3 GB): a PMap every 8 sections; an FMap
            // every 496 sections starting at section 128 (the header holds the FMap of the first 128); an FPMap every 31,744 sections
            // starting at section 8192 (the header holds the FPMap of the first 1,024 PMaps)
            int pages = 1;
            if (_sec % 8 == 0) { _pmapPages.Add(_secBase + 512); pages = 2; }
            if (_sec >= 128 && (_sec - 128) % 496 == 0) { _fmapPages.Add(_secBase + pages * 512L); pages++; }
            if (_sec >= 8192 && (_sec - 8192) % 31744 == 0) { _fpmapPages.Add(_secBase + pages * 512L); pages++; }
            for (int bit = 0; bit < pages * 8; bit++) _amap[bit >> 3] |= (byte)(0x80 >> (bit & 7));
            _cursor = pages * 512;
        }

        void FlushSection()
        {
            int free = 0, run = 0, best = 0;
            for (int bit = 0; bit < 3968; bit++)
            {
                if ((_amap[bit >> 3] & (0x80 >> (bit & 7))) == 0) { free++; run++; if (run > best) best = run; }
                else run = 0;
            }
            _secFree.Add((ushort)free);
            _secRun.Add((byte)Math.Min(best, 255));
            _freeBytes += (long)free * 64;
            Buffer.BlockCopy(_amap, 0, _buf, 0, 496);
            WriteTrailer(_buf, 496, PstFormat.PtypeAMap, 0, PstFormat.Crc(_amap, 0, 496), (ulong)_secBase);
            _fs.Position = _secBase;
            _fs.Write(_buf, 0, _buf.Length);
        }

        static void WriteTrailer(byte[] b, int o, byte ptype, ushort sig, uint crc, ulong bid)
        {
            b[o] = ptype; b[o + 1] = ptype; PstFormat.P16(b, o + 2, sig); PstFormat.P32(b, o + 4, crc); PstFormat.P64(b, o + 8, bid);
        }

        /// <summary>Reserves size bytes (a multiple of 64) in the current section, at an offset aligned on align, and returns the file offset.</summary>
        long Alloc(int size, int align)
        {
            if (_sec < 0) StartSection();
            int c = (_cursor + align - 1) / align * align;
            if (c + size > PstFormat.SectionSize)
            {
                FlushSection();
                StartSection();
                c = _cursor;
            }
            for (int bit = c / 64; bit < (c + size) / 64; bit++) _amap[bit >> 3] |= (byte)(0x80 >> (bit & 7));
            _cursor = c + size;
            return _secBase + c;
        }

        // ---------------------------------------------------------------- pages
        PstBref WritePage(byte[] page512, byte ptype)
        {
            long ib = Alloc(PstFormat.PageSize, PstFormat.PageSize);
            ulong bid = _bidNextP++;
            WriteTrailer(page512, 496, ptype, PstFormat.Sig((ulong)ib, bid), PstFormat.Crc(page512, 0, 496), bid);
            Buffer.BlockCopy(page512, 0, _buf, (int)(ib - _secBase), PstFormat.PageSize);
            PageCount++;
            return new PstBref { Bid = bid, Ib = (ulong)ib };
        }

        // ---------------------------------------------------------------- blocks
        ulong NextBid(bool isInternal)
        {
            ulong v = _bidNextB;
            _bidNextB += 4;
            return isInternal ? (v | 2) : v;
        }

        /// <summary>Writes one block (cb data bytes, at most 8,176) and records it in the BBT. Returns its BID.</summary>
        public ulong WriteBlock(byte[] data, int off, int cb, bool isInternal)
        {
            if (cb < 0 || cb > PstFormat.MaxBlockData) throw new ArgumentOutOfRangeException("cb");
            int total = PstFormat.BlockSize(cb);
            long ib = Alloc(total, 64);
            ulong bid = NextBid(isInternal);
            int pos = (int)(ib - _secBase);
            if (cb > 0) Buffer.BlockCopy(data, off, _buf, pos, cb);
            int t = pos + total - PstFormat.BlockTrailerSize;
            PstFormat.P16(_buf, t, cb);
            PstFormat.P16(_buf, t + 2, PstFormat.Sig((ulong)ib, bid));
            PstFormat.P32(_buf, t + 4, PstFormat.Crc(_buf, pos, cb));
            PstFormat.P64(_buf, t + 8, bid);
            var e = new byte[24];
            PstFormat.P64(e, 0, bid); PstFormat.P64(e, 8, (ulong)ib); PstFormat.P16(e, 16, cb); PstFormat.P16(e, 18, 2);
            _bbt.Add(0, e, bid);
            BlockCount++;
            return bid;
        }

        /// <summary>Data tree of a byte array: one data block, or an XBLOCK / XXBLOCK over data blocks of 8,176 bytes.</summary>
        public ulong WriteStream(byte[] data, int off, long length)
        {
            if (length == 0) return 0;
            var bids = new List<ulong>();
            var cbs = new List<int>();
            long pos = 0;
            while (pos < length)
            {
                int n = (int)Math.Min(PstFormat.MaxBlockData, length - pos);
                bids.Add(WriteBlock(data, off + (int)pos, n, false));
                cbs.Add(n);
                pos += n;
            }
            return WriteTree(bids, cbs);
        }

        readonly byte[] _blockBuf = new byte[PstFormat.MaxBlockData];

        /// <summary>Checked between two blocks of a streamed value (set by PstWriter.AddMessage for the duration of a call).</summary>
        public System.Threading.CancellationToken Cancellation;
        /// <summary>Told about every MiB of a streamed value written (and the rest at its end).</summary>
        public Action<long> OnBytes;

        /// <summary>
        /// Data tree of a stream read from its source in blocks of 8,176 bytes, each written as soon as it is read (a single block, or an
        /// XBLOCK / XXBLOCK over them). Memory: one block, plus 12 bytes per block for the tree (1.5 MiB per GiB of data).
        /// A data tree holds less than 4 GiB: the total of an XBLOCK / XXBLOCK is a 32-bit field (MS-PST 2.2.2.8.3.2).
        /// </summary>
        public ulong WriteStream(FxBlob source)
        {
            long length = source.Length;
            if (length == 0) return 0;
            if (length > uint.MaxValue) throw new NotSupportedException("A value of " + length + " bytes does not fit in a PST data tree: the total of an XBLOCK is a 32-bit field (less than 4 GiB).");
            int count = (int)((length + PstFormat.MaxBlockData - 1) / PstFormat.MaxBlockData);
            var bids = new List<ulong>(count);
            var cbs = new List<int>(count);
            using (Stream s = source.OpenRead())
            {
                long pos = 0, unreported = 0;
                while (pos < length)
                {
                    Cancellation.ThrowIfCancellationRequested();
                    int n = (int)Math.Min(PstFormat.MaxBlockData, length - pos);
                    if (FxIo.ReadFull(s, _blockBuf, 0, n) < n) throw new EndOfStreamException("The source of a value ends " + (length - pos) + " bytes before its announced length.");
                    bids.Add(WriteBlock(_blockBuf, 0, n, false));
                    cbs.Add(n);
                    pos += n; unreported += n;
                    if (OnBytes != null && unreported >= (1 << 20)) { OnBytes(unreported); unreported = 0; }
                }
                if (OnBytes != null && unreported > 0) OnBytes(unreported);
            }
            return WriteTree(bids, cbs);
        }

        /// <summary>Data tree over blocks given one by one (the leaf blocks of a heap, the blocks of a row matrix).</summary>
        public ulong WriteBlocks(List<byte[]> blocks)
        {
            if (blocks == null || blocks.Count == 0) return 0;
            var bids = new List<ulong>(blocks.Count);
            var cbs = new List<int>(blocks.Count);
            foreach (byte[] b in blocks) { bids.Add(WriteBlock(b, 0, b.Length, false)); cbs.Add(b.Length); }
            return WriteTree(bids, cbs);
        }

        ulong WriteTree(List<ulong> bids, List<int> cbs)
        {
            if (bids.Count == 1) return bids[0];
            var xb = new List<ulong>();
            var xt = new List<long>();
            for (int i = 0; i < bids.Count; i += PstFormat.XEntriesMax)
            {
                int n = Math.Min(PstFormat.XEntriesMax, bids.Count - i);
                long total = 0;
                var b = new byte[8 + 8 * n];
                b[0] = 1; b[1] = 1; PstFormat.P16(b, 2, n);
                for (int k = 0; k < n; k++) { PstFormat.P64(b, 8 + 8 * k, bids[i + k]); total += cbs[i + k]; }
                PstFormat.P32(b, 4, (uint)total);
                xb.Add(WriteBlock(b, 0, b.Length, true));
                xt.Add(total);
            }
            if (xb.Count == 1) return xb[0];
            if (xb.Count > PstFormat.XEntriesMax) throw new NotSupportedException("A value too large for a data tree of two levels.");
            var x = new byte[8 + 8 * xb.Count];
            x[0] = 1; x[1] = 2; PstFormat.P16(x, 2, xb.Count);
            long all = 0;
            for (int k = 0; k < xb.Count; k++) { PstFormat.P64(x, 8 + 8 * k, xb[k]); all += xt[k]; }
            PstFormat.P32(x, 4, (uint)all);
            return WriteBlock(x, 0, x.Length, true);
        }

        struct SubEntry { public uint Nid; public ulong BidData, BidSub; }

        /// <summary>Writes the data and the subnodes of a node; returns the BID of the data (0 if none) and of the subnode tree (0 if none).</summary>
        public ulong WriteNode(PstNode n, out ulong bidSub)
        {
            ulong data = 0;
            if (n.Raw != null) data = WriteStream(n.Raw, 0, n.Raw.Length);
            else if (n.Source != null) data = WriteStream(n.Source);
            else if (n.Blocks != null) data = WriteBlocks(n.Blocks);
            bidSub = 0;
            if (n.Subnodes.Count > 0)
            {
                var entries = new List<SubEntry>(n.Subnodes.Count);
                foreach (KeyValuePair<uint, PstNode> kv in n.Subnodes)
                {
                    ulong ss;
                    ulong sd = WriteNode(kv.Value, out ss);
                    entries.Add(new SubEntry { Nid = kv.Key, BidData = sd, BidSub = ss });
                }
                bidSub = WriteSubnodeTree(entries);
            }
            return data;
        }

        ulong WriteSubnodeTree(List<SubEntry> entries)
        {
            var sl = new List<ulong>();
            var first = new List<uint>();
            for (int i = 0; i < entries.Count; i += PstFormat.SlEntriesMax)
            {
                int n = Math.Min(PstFormat.SlEntriesMax, entries.Count - i);
                var b = new byte[8 + 24 * n];
                b[0] = 2; b[1] = 0; PstFormat.P16(b, 2, n);
                for (int k = 0; k < n; k++)
                {
                    SubEntry e = entries[i + k];
                    PstFormat.P64(b, 8 + 24 * k, e.Nid); PstFormat.P64(b, 16 + 24 * k, e.BidData); PstFormat.P64(b, 24 + 24 * k, e.BidSub);
                }
                sl.Add(WriteBlock(b, 0, b.Length, true));
                first.Add(entries[i].Nid);
            }
            if (sl.Count == 1) return sl[0];
            if (sl.Count > PstFormat.SiEntriesMax) throw new NotSupportedException("Too many subnodes in one node.");
            var si = new byte[8 + 16 * sl.Count];
            si[0] = 2; si[1] = 1; PstFormat.P16(si, 2, sl.Count);
            for (int k = 0; k < sl.Count; k++) { PstFormat.P64(si, 8 + 16 * k, first[k]); PstFormat.P64(si, 16 + 16 * k, sl[k]); }
            return WriteBlock(si, 0, si.Length, true);
        }

        /// <summary>Registers a top-level node in the NBT (written at Close).</summary>
        public void AddNode(uint nid, ulong bidData, ulong bidSub, uint nidParent)
        {
            _nbt.Add(new NbtRec { Nid = nid, BidData = bidData, BidSub = bidSub, Parent = nidParent });
            // the high-water mark of the header for each NID type is the highest index in use (Outlook validates it)
            uint type = nid & 0x1F, index = nid >> 5;
            if (type != 0 && type != PstFormat.NtInternal && Rgnid[type] < index) Rgnid[type] = index;
        }

        public int NodeCount { get { return _nbt.Count; } }

        /// <summary>Writes a node and registers it in the NBT.</summary>
        public void WriteTopLevelNode(uint nid, PstNode n, uint nidParent)
        {
            ulong sub;
            ulong data = WriteNode(n, out sub);
            AddNode(nid, data, sub, nidParent);
        }

        // ---------------------------------------------------------------- Close
        public void Close()
        {
            if (_closed) return;
            _closed = true;
            // NBT
            _nbt.Sort((a, b) => a.Nid.CompareTo(b.Nid));
            var nbt = new BTreeBuilder(this, PstFormat.PtypeNBT, 32);
            uint prev = 0; bool havePrev = false;
            foreach (NbtRec r in _nbt)
            {
                if (havePrev && r.Nid == prev) throw new InvalidOperationException("Duplicate NID 0x" + r.Nid.ToString("X"));
                prev = r.Nid; havePrev = true;
                var e = new byte[32];
                PstFormat.P64(e, 0, r.Nid); PstFormat.P64(e, 8, r.BidData); PstFormat.P64(e, 16, r.BidSub); PstFormat.P32(e, 24, r.Parent);
                nbt.Add(0, e, r.Nid);
            }
            PstBref nbtRoot = nbt.Finish();
            PstBref bbtRoot = _bbt.Finish();
            if (_sec < 0) StartSection();
            FlushSection();
            int sections = _sec + 1;
            long eof = PstFormat.FirstAMapOffset + (long)sections * PstFormat.SectionSize;
            long lastAMap = _secBase;

            // PMap, FMap, FPMap pages (the PMap is deprecated: all pages "taken", so that cbPMapFree = 0 as the header says;
            // the FMap holds the longest free run of each AMap, 0xFF for AMaps that do not exist)
            var pm = new byte[PstFormat.PageSize];
            for (int i = 0; i < 496; i++) pm[i] = 0xFF;
            foreach (long ib in _pmapPages) PatchPage(ib, pm, PstFormat.PtypePMap);
            for (int j = 0; j < _fmapPages.Count; j++)
            {
                var fm = new byte[PstFormat.PageSize];
                for (int i = 0; i < 496; i++)
                {
                    int s = 128 + j * 496 + i;
                    fm[i] = s < sections ? _secRun[s] : (byte)0xFF;
                }
                PatchPage(_fmapPages[j], fm, PstFormat.PtypeFMap);
            }
            var fp = new byte[PstFormat.PageSize];
            for (int i = 0; i < 496; i++) fp[i] = 0xFF;
            foreach (long ib in _fpmapPages) PatchPage(ib, fp, PstFormat.PtypeFPMap);

            WriteDList(sections);

            // header
            var h = new byte[PstFormat.HeaderSize];
            h[0] = 0x21; h[1] = 0x42; h[2] = 0x44; h[3] = 0x4E;
            h[8] = 0x53; h[9] = 0x4D;
            PstFormat.P16(h, 10, 23);          // wVer: Unicode
            PstFormat.P16(h, 12, 19);          // wVerClient
            h[14] = 1; h[15] = 1;
            PstFormat.P64(h, 32, _bidNextP);
            PstFormat.P32(h, 40, ++_dwUnique);
            for (int i = 0; i < 32; i++) PstFormat.P32(h, 44 + 4 * i, Rgnid[i]);
            // root at 180
            PstFormat.P64(h, 184, (ulong)eof);
            PstFormat.P64(h, 192, (ulong)lastAMap);
            PstFormat.P64(h, 200, (ulong)_freeBytes);
            PstFormat.P64(h, 208, 0);
            PstFormat.P64(h, 216, nbtRoot.Bid); PstFormat.P64(h, 224, nbtRoot.Ib);
            PstFormat.P64(h, 232, bbtRoot.Bid); PstFormat.P64(h, 240, bbtRoot.Ib);
            h[248] = 2;                         // fAMapValid
            for (int i = 256; i < 512; i++) h[i] = 0xFF;      // rgbFM, rgbFP: deprecated, "MUST be filled with 0xFF" (MS-PST 2.2.2.6); the FMap pages hold the free runs
            h[512] = 0x80;                      // bSentinel
            h[513] = 0;                         // NDB_CRYPT_NONE
            PstFormat.P64(h, 516, _bidNextB);
            PstFormat.P32(h, 4, PstFormat.Crc(h, 8, 471));
            PstFormat.P32(h, 524, PstFormat.Crc(h, 8, 516));
            _fs.Position = 0;
            _fs.Write(h, 0, h.Length);
            _fs.Flush(true);
            _fs.Dispose();
        }

        void PatchPage(long ib, byte[] page512, byte ptype)
        {
            var p = (byte[])page512.Clone();
            WriteTrailer(p, 496, ptype, 0, PstFormat.Crc(p, 0, 496), (ulong)ib);
            _fs.Position = ib;
            _fs.Write(p, 0, p.Length);
        }

        void WriteDList(int sections)
        {
            var order = new List<int>();
            for (int s = 0; s < sections; s++) if (_secFree[s] > 0) order.Add(s);
            order.Sort((a, b) => { int c = _secFree[b].CompareTo(_secFree[a]); return c != 0 ? c : a.CompareTo(b); });
            int n = Math.Min(order.Count, 119);
            var p = new byte[PstFormat.PageSize];
            p[0] = 1; p[1] = (byte)n;
            PstFormat.P32(p, 4, (uint)(sections - 1));
            for (int i = 0; i < n; i++) PstFormat.P32(p, 8 + 4 * i, (uint)order[i] | ((uint)_secFree[order[i]] << 20));
            // dwPageNum (20 bits) is the low part, dwFreeSlots (12 bits) the high part
            WriteTrailer(p, 496, PstFormat.PtypeDL, PstFormat.Sig((ulong)PstFormat.DListOffset, 0), PstFormat.Crc(p, 0, 496), 0);
            _fs.Position = PstFormat.DListOffset;
            _fs.Write(p, 0, p.Length);
        }

        public void Dispose()
        {
            if (!_closed) { _closed = true; try { _fs.Dispose(); } catch (Exception) { } }
        }
    }
}
