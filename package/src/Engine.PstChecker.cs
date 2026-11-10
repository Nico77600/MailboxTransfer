// =============================================================================
//  Mailbox Transfer - engine, part 5g: MS-PST integrity checker (report only)
// -----------------------------------------------------------------------------
//  PstChecker.Check(path, options) returns a PstCheckReport: findings (Info / Warning / Error) by layer
//    NDB        header and CRCs, AMaps (and the allocation of every block and page against them), DList, BBT and NBT (page
//               trailers, key order, levels, cycles, dangling BIDs), blocks (trailer, signature, CRC), reference counts, orphan
//               blocks, data trees (XBLOCK / XXBLOCK sizes), subnode trees (SLBLOCK / SIBLOCK)
//    LTP        heaps (page maps, HIDs), BTHs, property contexts (types, sizes, subnodes), table contexts (columns, row index
//               against the Row Matrix, Cell Existence Block, HNIDs)
//    Messaging  message store, name-to-id map, mandatory nodes, folder tree (parents, hierarchy tables), contents and FAI tables
//               against the messages of the NBT, counts, messages (recipient and attachment tables against the subnodes,
//               embedded messages, named property ids)
//
//  The file is never trusted: every offset, size and count is checked before it is used, every B-tree and tree walk is bounded
//  (visited pages, depth limits), every phase is guarded so that a corrupt file gives findings, not an exception. Memory: the
//  header-sized tables per allocation section (1 bit per 64 bytes of file, twice) and about 30 bytes per block / node: a few
//  hundred MB for tens of GB. Findings carry a code (stable, usable by a future repair), the file offset, BID, NID and, for the
//  messaging layer, the folder path and subject. Repair is NOT done here.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public enum PstSeverity { Info = 0, Warning = 1, Error = 2 }

    public sealed class PstFinding
    {
        public PstSeverity Severity;
        public string Layer, Code, Message, Where;
        public long Offset = -1;
        public ulong Bid;
        public uint Nid;

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(Severity.ToString().ToUpperInvariant().PadRight(7)).Append(' ').Append(Layer).Append(' ').Append(Code).Append(": ").Append(Message);
            if (Offset >= 0) sb.Append(" [offset 0x").Append(Offset.ToString("X")).Append(']');
            if (Bid != 0) sb.Append(" [BID 0x").Append(Bid.ToString("X")).Append(']');
            if (Nid != 0) sb.Append(" [NID 0x").Append(Nid.ToString("X")).Append(']');
            if (!string.IsNullOrEmpty(Where)) sb.Append(" [").Append(Where).Append(']');
            return sb.ToString();
        }
    }

    public sealed class PstCheckOptions
    {
        /// <summary>Read every block and check its trailer, signature and CRC (the whole file is read).</summary>
        public bool VerifyBlocks = true;
        /// <summary>Check the LTP and Messaging layers (needs the NDB layer to be readable).</summary>
        public bool Messaging = true;
        /// <summary>Open every message (property context, recipient and attachment tables, embedded messages).</summary>
        public bool Deep = true;
        /// <summary>Findings listed per code; the others are only counted.</summary>
        public int MaxPerCode = 50;
        /// <summary>Rows of a table whose cells are all decoded; larger tables are sampled with a stride.</summary>
        public int MaxCellRows = 20000;
        public Action<string> Progress;
        /// <summary>Typed progress (phase, bytes of blocks verified / file length, messages checked / total); may be called from the thread running the check.</summary>
        public IProgress<PstProgress> ProgressReport;
        /// <summary>Stops the check between two blocks / two messages: the report then has Cancelled = true and holds the findings so far.</summary>
        public System.Threading.CancellationToken Cancellation;
    }

    public sealed class PstCheckReport
    {
        public string Path;
        public long FileLength;
        public readonly List<PstFinding> Findings = new List<PstFinding>();
        public readonly Dictionary<string, int> CountByCode = new Dictionary<string, int>();
        public readonly Dictionary<string, long> Stats = new Dictionary<string, long>();
        public int Errors, Warnings, Infos;
        /// <summary>The check was stopped by PstCheckOptions.Cancellation: the findings are those found until then.</summary>
        public bool Cancelled;
        public TimeSpan Elapsed;
        public bool Ok { get { return Errors == 0; } }
        /// <summary>0 no finding above Info, 2 warnings only, 1 errors (same convention as the other modes of the tool).</summary>
        public int ExitCode { get { return Errors > 0 ? 1 : Warnings > 0 ? 2 : 0; } }

        public string Summary()
        {
            var sb = new StringBuilder();
            sb.AppendLine(Path + " (" + FileLength.ToString("N0") + " bytes)");
            sb.Append("  ").Append(string.Join(", ", Stats.OrderBy(k => k.Key).Select(k => k.Key + " " + k.Value.ToString("N0"))));
            sb.AppendLine();
            sb.Append("  result: ").Append(Errors).Append(" error(s), ").Append(Warnings).Append(" warning(s), ").Append(Infos).Append(" information(s) in ").Append(Elapsed.TotalSeconds.ToString("N1")).Append(" s");
            return sb.ToString();
        }
    }

    public static class PstChecker
    {
        public static PstCheckReport Check(string path) { return Check(path, new PstCheckOptions()); }

        public static PstCheckReport Check(string path, PstCheckOptions options)
        {
            var c = new PstCheck(path, options ?? new PstCheckOptions());
            c.Run();
            return c.Report;
        }
    }

    sealed class PstCheck
    {
        const int Sec = PstFormat.SectionSize;
        const long First = PstFormat.FirstAMapOffset;

        readonly string _path;
        readonly PstCheckOptions _o;
        public readonly PstCheckReport Report = new PstCheckReport();
        FileStream _fs;
        long _len;
        readonly Stopwatch _sw = Stopwatch.StartNew();

        // header
        int _crypt;
        long _eof, _amapLast, _amapFree, _pmapFree;
        uint _unique;
        PstBref _nbtRoot, _bbtRoot;
        ulong _bidNextB, _bidNextP;
        byte _amapValid;
        uint[] _rgnid = new uint[32];

        // allocation
        int _nSec;
        byte[] _amap, _used;
        bool[] _amapOk;

        // BBT
        readonly List<ulong> _bBid = new List<ulong>();
        readonly List<long> _bIb = new List<long>();
        readonly List<ushort> _bCb = new List<ushort>(), _bCref = new List<ushort>(), _bRefs = new List<ushort>();
        bool _bbtSorted = true;

        // NBT
        readonly List<uint> _nNid = new List<uint>(), _nParent = new List<uint>();
        readonly List<ulong> _nData = new List<ulong>(), _nSub = new List<ulong>();
        bool _nbtSorted = true;

        PstFile _pf;
        int _namedCount = -1;
        // replication GUID -> NID map of node 0xC01 (null: missing or unreadable), values of the folder property contexts
        Dictionary<string, uint> _hmp;
        readonly Dictionary<uint, int[]> _folderVals = new Dictionary<uint, int[]>();
        readonly uint[] _subMax = new uint[32];

        public PstCheck(string path, PstCheckOptions o) { _path = path; _o = o; Report.Path = path; }

        // ---------------------------------------------------------------- findings
        void Add(PstSeverity sev, string layer, string code, string msg, long offset, ulong bid, uint nid, string where)
        {
            switch (sev) { case PstSeverity.Error: Report.Errors++; break; case PstSeverity.Warning: Report.Warnings++; break; default: Report.Infos++; break; }
            if (HeuristicCodes.Contains(code)) msg += Heuristic;
            int n;
            Report.CountByCode.TryGetValue(code, out n);
            Report.CountByCode[code] = ++n;
            if (n > _o.MaxPerCode) return;
            Report.Findings.Add(new PstFinding { Severity = sev, Layer = layer, Code = code, Message = msg, Offset = offset, Bid = bid, Nid = nid, Where = where });
        }

        void Err(string layer, string code, string msg, long off = -1, ulong bid = 0, uint nid = 0, string where = null) { Add(PstSeverity.Error, layer, code, msg, off, bid, nid, where); }
        void Warn(string layer, string code, string msg, long off = -1, ulong bid = 0, uint nid = 0, string where = null) { Add(PstSeverity.Warning, layer, code, msg, off, bid, nid, where); }
        void Info(string layer, string code, string msg, long off = -1, ulong bid = 0, uint nid = 0, string where = null) { Add(PstSeverity.Info, layer, code, msg, off, bid, nid, where); }

        // Rules reported as information: they come from the verdict of an old scan tool during development, they are neither in the
        // specification nor needed by Outlook to open a file (the other rules come from the specification or from the behaviour of Outlook)
        static readonly HashSet<string> HeuristicCodes = new HashSet<string> { "MSG-OUTGOING-QUEUE", "MSG-TEMPLATE-MISSING", "MSG-NODE-EC1", "HMP-MISSING", "HMP-ENTRY", "RECEIVE-TABLE-ROW", "TEMPLATE-COLUMN" };
        const string Heuristic = " [heuristic of an old scan tool: not in the specification, Outlook opens files without it]";

        void Progress(string s)
        {
            if (_o.Progress != null) _o.Progress(s);
            if (_o.ProgressReport != null) _o.ProgressReport.Report(new PstProgress { Phase = s, BytesTotal = _len });
        }

        bool ReadAt(long ib, byte[] buf, int len)
        {
            if (ib < 0 || ib + len > _len) return false;
            try
            {
                _fs.Position = ib;
                int got = 0;
                while (got < len) { int n = _fs.Read(buf, got, len - got); if (n <= 0) return false; got += n; }
                return true;
            }
            catch (IOException) { return false; }
        }

        // ---------------------------------------------------------------- run
        public void Run()
        {
            try
            {
                _fs = new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, 1 << 16, FileOptions.RandomAccess);
                _len = _fs.Length;
                Report.FileLength = _len;
            }
            catch (Exception e) { Err("NDB", "FILE-OPEN", "The file cannot be opened: " + e.Message); Finish(); return; }
            string reserved = FileNaming.ReservedDeviceName(System.IO.Path.GetFileName(_path));
            if (reserved != null)
                Warn("File", "FILE-RESERVED-NAME", "The file name '" + System.IO.Path.GetFileName(_path) + "' is the reserved Windows device name " + reserved + " (with any extension): this file opens in our tools, but Outlook crashes when it opens a PST named like this (observed with aux.pst; the content is fine). Rename it, for example to '" + FileNaming.SafeFileName(System.IO.Path.GetFileName(_path)) + "'.");
            try
            {
                if (Header())
                {
                    Phase("allocation maps", Sections);
                    Phase("block B-tree", WalkBbt);
                    Phase("node B-tree", WalkNbt);
                    Phase("header NID high-water marks", CheckHighWater);
                    Phase("data trees and subnode trees", Trees);
                    Phase("allocation against the maps", CompareAllocation);
                    if (_o.Messaging) Phase("LTP and messaging layers", Messaging);
                }
            }
            catch (OperationCanceledException) { Report.Cancelled = true; }
            catch (Exception e) { Err("NDB", "CHECK-ABORT", "The check stopped on an unexpected failure (" + e.GetType().Name + "): " + e.Message); }
            finally { _fs.Dispose(); if (_pf != null) _pf.Dispose(); }
            Finish();
        }

        void Phase(string name, Action a)
        {
            Progress(name);
            _o.Cancellation.ThrowIfCancellationRequested();
            try { a(); }
            catch (OperationCanceledException) { throw; }
            catch (Exception e) { Err("NDB", "CHECK-ABORT", "Phase '" + name + "' stopped on an unexpected failure (" + e.GetType().Name + "): " + e.Message); }
        }

        void Finish()
        {
            foreach (KeyValuePair<string, int> kv in Report.CountByCode)
                if (kv.Value > _o.MaxPerCode)
                    Report.Findings.Add(new PstFinding { Severity = PstSeverity.Info, Layer = "Check", Code = kv.Key, Message = (kv.Value - _o.MaxPerCode) + " more finding(s) of this code are counted but not listed", });
            Report.Stats["blocks"] = _bBid.Count;
            Report.Stats["nodes"] = _nNid.Count;
            Report.Elapsed = _sw.Elapsed;
        }

        // ---------------------------------------------------------------- header
        bool Header()
        {
            if (_len < PstFormat.HeaderSize) { Err("NDB", "HDR-SHORT", "The file is shorter than a PST header (" + _len + " bytes)."); return false; }
            var h = new byte[PstFormat.HeaderSize];
            ReadAt(0, h, h.Length);
            if (h[0] != 0x21 || h[1] != 0x42 || h[2] != 0x44 || h[3] != 0x4E) { Err("NDB", "HDR-MAGIC", "dwMagic is not '!BDN': this is not a PST file.", 0); return false; }
            if (h[8] != 0x53 || h[9] != 0x4D) Err("NDB", "HDR-MAGIC-CLIENT", "wMagicClient is not 'SM'.", 8);
            int ver = PstFormat.G16(h, 10);
            if (ver == 14 || ver == 15) { Err("NDB", "HDR-ANSI", "ANSI PST (wVer " + ver + "): the checker handles Unicode PST files only.", 10); return false; }
            if (ver < 23) { Err("NDB", "HDR-VERSION", "Unknown wVer " + ver + ".", 10); return false; }
            if (PstFormat.G16(h, 12) != 19) Info("NDB", "HDR-CLIENTVER", "wVerClient is " + PstFormat.G16(h, 12) + " (19 is the version of the specification).", 12);
            if (h[14] != 1 || h[15] != 1) Warn("NDB", "HDR-PLATFORM", "bPlatformCreate / bPlatformAccess are not 1.", 14);
            if (PstFormat.Crc(h, 8, 471) != PstFormat.G32(h, 4)) Err("NDB", "HDR-CRC-PARTIAL", "dwCRCPartial does not match the 471 bytes it covers.", 4);
            if (PstFormat.Crc(h, 8, 516) != PstFormat.G32(h, 524)) Err("NDB", "HDR-CRC-FULL", "dwCRCFull does not match the 516 bytes it covers.", 524);
            if (h[512] != 0x80) Err("NDB", "HDR-SENTINEL", "bSentinel is 0x" + h[512].ToString("X2") + " (0x80 expected).", 512);
            _crypt = h[513];
            if (_crypt != 0 && _crypt != 1 && _crypt != 2) { Err("NDB", "HDR-CRYPT", "bCryptMethod " + _crypt + " is not supported (0, 1 or 2).", 513); }
            _bidNextP = PstFormat.G64(h, 32);
            for (int i = 0; i < 32; i++) _rgnid[i] = PstFormat.G32(h, 44 + 4 * i);
            _eof = (long)PstFormat.G64(h, 184);
            _amapLast = (long)PstFormat.G64(h, 192);
            _amapFree = (long)PstFormat.G64(h, 200);
            _pmapFree = (long)PstFormat.G64(h, 208);
            _unique = PstFormat.G32(h, 40);
            _nbtRoot = new PstBref { Bid = PstFormat.G64(h, 216), Ib = PstFormat.G64(h, 224) };
            _bbtRoot = new PstBref { Bid = PstFormat.G64(h, 232), Ib = PstFormat.G64(h, 240) };
            _amapValid = h[248];
            _bidNextB = PstFormat.G64(h, 516);
            if (_eof > _len) Err("NDB", "HDR-EOF", "ibFileEof (" + _eof + ") is beyond the end of the file (" + _len + "): the file is truncated.", 184);
            else if (_eof != _len) Warn("NDB", "HDR-EOF-DIFF", "ibFileEof (" + _eof + ") differs from the size of the file (" + _len + ").", 184);
            for (int i = 256; i < 512; i++)
                if (h[i] != 0xFF) { Warn("NDB", "HDR-RGBFM-RGBFP", "rgbFM / rgbFP (the deprecated initial FMap and FPMap of the header) are not all 0xFF; the specification says they MUST be (MS-PST 2.2.2.6).", i); break; }
            if (_amapValid == 0) Warn("NDB", "HDR-AMAP-INVALID", "fAMapValid says the AMaps are INVALID: Outlook rebuilds them at the next open.", 248);
            else if (_amapValid > 2) Warn("NDB", "HDR-AMAP-VALUE", "fAMapValid has an undefined value " + _amapValid + ".", 248);
            if (_unique == 0) Warn("NDB", "HDR-UNIQUE", "dwUnique is zero: it is incremented at every change of the header, a header that was ever written has a positive value.", 40);
            if (_pmapFree != 0) Warn("NDB", "HDR-PMAPFREE", "cbPMapFree is " + _pmapFree + ": the PMap (deprecated) of a file without free page is all ones and the header says 0.", 208);
            if (_bidNextB == 0 || (_bidNextB & 3) != 0) Warn("NDB", "HDR-BIDNEXT", "bidNextB 0x" + _bidNextB.ToString("X") + " is not a positive multiple of 4.", 516);
            Report.Stats["crypt"] = _crypt;
            return _crypt <= 2;
        }

        // ---------------------------------------------------------------- allocation sections
        void Sections()
        {
            if (_len <= First) { Err("NDB", "AMAP-NONE", "The file has no allocation section."); _nSec = 0; return; }
            _nSec = (int)((_len - First + Sec - 1) / Sec);
            _amap = new byte[_nSec * 496];
            _used = new byte[_nSec * 496];
            _amapOk = new bool[_nSec];
            var p = new byte[512];
            long free = 0;
            long lastAmap = -1;
            for (int k = 0; k < _nSec; k++)
            {
                long at = First + (long)k * Sec;
                if (!ReadAt(at, p, 512)) { Err("NDB", "AMAP-MISSING", "The AMap of section " + k + " lies beyond the end of the file.", at); continue; }
                string why = PageProblem(p, at, PstFormat.PtypeAMap, true);
                if (why != null) { Err("NDB", "AMAP-BAD", "AMap " + k + ": " + why, at); continue; }
                _amapOk[k] = true;
                lastAmap = at;
                Buffer.BlockCopy(p, 0, _amap, k * 496, 496);
                // the AMap maps itself: its own page is allocated
                for (int bit = 0; bit < 8; bit++) _used[k * 496 + (bit >> 3)] |= (byte)(0x80 >> (bit & 7));
                int zeros = 0;
                for (int i = 0; i < 496; i++) zeros += 8 - PopCount(p[i]);
                free += (long)zeros * 64;
                // the three following pages may be a PMap, FMap or FPMap: a page that is valid is taken for what it says it is wherever it
                // is; a page that is not valid is reported only where a map is expected (PMap every 8 sections, FMap every 496 from 128,
                // FPMap every 31,744 from 8,192), else it is data (any block can start with bytes that look like a map type)
                var q = new byte[512];
                var expected = new List<byte>();
                if (k % 8 == 0) expected.Add(PstFormat.PtypePMap);
                if (k >= 128 && (k - 128) % 496 == 0) expected.Add(PstFormat.PtypeFMap);
                if (k >= 8192 && (k - 8192) % 31744 == 0) expected.Add(PstFormat.PtypeFPMap);
                for (int pg = 1; pg <= 3; pg++)
                {
                    long qa = at + pg * 512L;
                    if (!ReadAt(qa, q, 512)) continue;
                    if ((q[496] == PstFormat.PtypePMap || q[496] == PstFormat.PtypeFMap || q[496] == PstFormat.PtypeFPMap) && q[497] == q[496])
                    {
                        string w = PageProblem(q, qa, q[496], true);
                        if (w != null && !(pg - 1 < expected.Count && expected[pg - 1] == q[496])) continue;
                        if (w != null) Warn("NDB", "MAP-BAD", "Allocation page of type 0x" + q[496].ToString("X2") + ": " + w + " (deprecated map: Outlook rewrites it).", qa);
                        else if (q[496] == PstFormat.PtypePMap)
                        {
                            bool full = true;
                            for (int i = 0; i < 496 && full; i++) full = q[i] == 0xFF;
                            if (!full) Warn("NDB", "PMAP-NOT-FULL", "The PMap (deprecated) is not all ones: every page is taken in a file written without free page, Outlook rewrites it.", qa);
                        }
                        for (int bit = pg * 8; bit < pg * 8 + 8; bit++) _used[k * 496 + (bit >> 3)] |= (byte)(0x80 >> (bit & 7));
                    }
                }
            }
            if (_amapValid != 0 && free != _amapFree) Warn("NDB", "HDR-AMAPFREE", "cbAMapFree in the header (" + _amapFree + ") differs from the free space of the AMaps (" + free + ").", 200);
            if (lastAmap >= 0 && _amapLast != lastAmap) Warn("NDB", "HDR-AMAPLAST", "ibAMapLast (0x" + _amapLast.ToString("X") + ") is not the last AMap of the file (0x" + lastAmap.ToString("X") + ").", 192);
            Report.Stats["sections"] = _nSec;
            Report.Stats["freeBytes"] = free;
            // density list
            if (_len >= PstFormat.DListOffset + 512)
            {
                ReadAt(PstFormat.DListOffset, p, 512);
                if (p[496] == PstFormat.PtypeDL)
                {
                    string w = PageProblem(p, PstFormat.DListOffset, PstFormat.PtypeDL, false);
                    if (w != null) Warn("NDB", "DLIST-BAD", "Density list: " + w + " (it is recreated from the AMaps by Outlook).", PstFormat.DListOffset);
                    else
                    {
                        int cEnt = p[1];
                        if (cEnt > 119) Warn("NDB", "DLIST-COUNT", "Density list with " + cEnt + " entries (119 at most).", PstFormat.DListOffset);
                        for (int i = 0; i < Math.Min(cEnt, 119); i++)
                        {
                            uint e = PstFormat.G32(p, 8 + 4 * i);
                            int page = (int)(e & 0xFFFFF), slots = (int)(e >> 20);
                            if (page >= _nSec) Warn("NDB", "DLIST-ENTRY", "Density list entry " + i + " points to AMap " + page + " but there are " + _nSec + ".", PstFormat.DListOffset);
                            else if (_amapOk[page])
                            {
                                int z = 0; for (int b = 0; b < 496; b++) z += 8 - PopCount(_amap[page * 496 + b]);
                                if (z != slots) Info("NDB", "DLIST-FREE", "Density list entry " + i + ": " + slots + " free slots, the AMap has " + z + ".", PstFormat.DListOffset);
                            }
                        }
                    }
                }
            }
        }

        static byte ReverseBits(byte b) { int r = 0; for (int i = 0; i < 8; i++) if ((b & (1 << i)) != 0) r |= 0x80 >> i; return (byte)r; }

        static int PopCount(byte b) { int n = 0; while (b != 0) { n += b & 1; b >>= 1; } return n; }

        /// <summary>Null if the 512-byte page has a valid trailer (type, repeated type, CRC, and for non-map pages the signature).</summary>
        string PageProblem(byte[] p, long ib, byte ptype, bool mapPage)
        {
            if (p[496] != ptype || p[497] != ptype) return "page type 0x" + p[496].ToString("X2") + "/0x" + p[497].ToString("X2") + " (0x" + ptype.ToString("X2") + " expected)";
            if (PstFormat.Crc(p, 0, 496) != PstFormat.G32(p, 500)) return "CRC mismatch";
            ulong bid = PstFormat.G64(p, 504);
            if (mapPage)
            {
                if (PstFormat.G16(p, 498) != 0) return "wSig is not zero";
                if (bid != (ulong)ib) return "bid 0x" + bid.ToString("X") + " is not the page offset";
            }
            else if (ptype != PstFormat.PtypeDL && PstFormat.G16(p, 498) != PstFormat.Sig((ulong)ib, bid)) return "signature mismatch";
            return null;
        }

        // ---------------------------------------------------------------- allocation bookkeeping
        void MarkUsed(long ib, int size, string what, ulong bid)
        {
            if (ib < First) { Err("NDB", "ALLOC-RANGE", what + " lies before the first allocation section.", ib, bid); return; }
            long rel = ib - First;
            int k = (int)(rel / Sec);
            long inSec = rel % Sec;
            if (inSec + size > Sec) { Err("NDB", "ALLOC-CROSS", what + " crosses the end of its allocation section.", ib, bid); return; }
            if (k >= _nSec) return;
            bool overlap = false;
            for (long u = inSec / 64; u < (inSec + size) / 64; u++)
            {
                int idx = k * 496 + (int)(u >> 3);
                byte m = (byte)(0x80 >> (int)(u & 7));
                if ((_used[idx] & m) != 0) overlap = true;
                _used[idx] |= m;
            }
            if (overlap) Err("NDB", "ALLOC-OVERLAP", what + " overlaps another block, page or map.", ib, bid);
        }

        void CompareAllocation()
        {
            long unmarked = 0, unusedBytes = 0, firstUnmarked = -1;
            int reversedSections = 0;
            for (int k = 0; k < _nSec; k++)
            {
                if (!_amapOk[k]) continue;
                // the first 64 bytes of a section are bit 7 of byte 0 (most significant bit first): a map whose bits come least
                // significant first marks other blocks; if reading it backwards covers every block, say that instead
                bool direct = true, reversed = true;
                for (int i = 0; i < 496 && (direct || reversed); i++)
                {
                    byte a = _amap[k * 496 + i], u = _used[k * 496 + i];
                    if ((u & ~a & 0xFF) != 0) direct = false;
                    if ((u & ~ReverseBits(a) & 0xFF) != 0) reversed = false;
                }
                bool bitOrder = !direct && reversed;
                if (bitOrder) { reversedSections++; Err("NDB", "AMAP-BIT-ORDER", "AMap " + k + " numbers its bits least significant first: the first 64 bytes of the section are the most significant bit of byte 0.", First + (long)k * Sec); }
                for (int i = 0; i < 496; i++)
                {
                    byte a = _amap[k * 496 + i], u = _used[k * 496 + i];
                    if (bitOrder) continue;
                    byte miss = (byte)(u & ~a);
                    if (miss != 0)
                    {
                        unmarked += PopCount(miss);
                        if (firstUnmarked < 0) firstUnmarked = First + (long)k * Sec + (long)i * 8 * 64;
                    }
                    unusedBytes += PopCount((byte)(a & ~u)) * 64L;
                }
            }
            if (unmarked > 0) Err("NDB", "AMAP-UNMARKED", unmarked + " unit(s) of 64 bytes hold blocks or pages but the AMap says they are free (first near this offset): the next write would overwrite data.", firstUnmarked);
            if (unusedBytes > 0) Info("NDB", "AMAP-UNUSED", unusedBytes.ToString("N0") + " byte(s) are marked allocated in the AMaps but belong to no block, page or map of the trees (reclaimable space, or deprecated map pages).");
            Report.Stats["unusedBytes"] = unusedBytes;
        }

        // ---------------------------------------------------------------- B-trees
        sealed class Frame { public PstBref Ref; public int Level; public ulong Min, Max; }

        /// <summary>Walks a B-tree in key order. leaf(page, offset) is called for each leaf entry. Returns false if the root cannot be read.</summary>
        void WalkTree(PstBref root, byte ptype, string tree, int leafSize, Action<byte[], int> leaf)
        {
            var visited = new HashSet<long>();
            var stack = new Stack<Frame>();
            stack.Push(new Frame { Ref = root, Level = -1, Min = 0, Max = ulong.MaxValue });
            var page = new byte[512];
            bool anyLeaf = false;
            ulong last = 0; bool haveLast = false;
            while (stack.Count > 0)
            {
                Frame f = stack.Pop();
                long ib = (long)f.Ref.Ib;
                if (ib % 512 != 0 || ib < First || ib + 512 > _len) { Err("NDB", tree + "-PAGE-RANGE", tree + " page reference to offset 0x" + ib.ToString("X") + " is misaligned or outside the file.", ib, f.Ref.Bid); continue; }
                if (!visited.Add(ib)) { Err("NDB", tree + "-CYCLE", tree + " page at 0x" + ib.ToString("X") + " is reachable twice (cycle or shared page).", ib, f.Ref.Bid); continue; }
                if (!ReadAt(ib, page, 512)) { Err("NDB", tree + "-PAGE-READ", tree + " page cannot be read.", ib, f.Ref.Bid); continue; }
                Report.Stats["btreePages"] = (Report.Stats.ContainsKey("btreePages") ? Report.Stats["btreePages"] : 0) + 1;
                MarkUsed(ib, 512, tree + " page", f.Ref.Bid);
                string why = PageProblem(page, ib, ptype, false);
                if (why != null) { Err("NDB", tree + "-PAGE", tree + " page: " + why + ".", ib, f.Ref.Bid); continue; }
                if (_bidNextP != 0 && f.Ref.Bid >= _bidNextP) Warn("NDB", "HDR-BIDNEXTP", tree + " page BID 0x" + f.Ref.Bid.ToString("X") + " is not below bidNextP 0x" + _bidNextP.ToString("X") + " in the header: the next page allocated could collide.", ib, f.Ref.Bid);
                if (PstFormat.G64(page, 504) != f.Ref.Bid) Err("NDB", tree + "-PAGE-BID", tree + " page trailer BID 0x" + PstFormat.G64(page, 504).ToString("X") + " differs from the BID of the reference 0x" + f.Ref.Bid.ToString("X") + ".", ib, f.Ref.Bid);
                int cEnt = page[488], cEntMax = page[489], cb = page[490], level = page[491];
                if (level > 8) { Err("NDB", tree + "-LEVEL", tree + " page level " + level + " is impossible (8 at most).", ib); continue; }
                if (f.Level >= 0 && level != f.Level) { Err("NDB", tree + "-LEVEL", tree + " page level " + level + " (the parent expects " + f.Level + ").", ib); continue; }
                int want = level == 0 ? leafSize : 24;
                if (cb < want || cEnt > cEntMax || cEnt * cb > 488 || cEntMax * cb > 488) { Err("NDB", tree + "-PAGE-SHAPE", tree + " page counts are inconsistent (cEnt " + cEnt + ", cEntMax " + cEntMax + ", cbEnt " + cb + ", level " + level + ").", ib); continue; }
                if (cEnt == 0 && !(level == 0 && f.Level < 0)) Warn("NDB", tree + "-PAGE-EMPTY", tree + " page without entry.", ib);
                if (level == 0)
                {
                    for (int i = 0; i < cEnt; i++)
                    {
                        ulong key = PstFormat.G64(page, i * cb);
                        if (key < f.Min || key >= f.Max) Err("NDB", tree + "-KEY-RANGE", "Key 0x" + key.ToString("X") + " is outside the range its parent allows (0x" + f.Min.ToString("X") + "..0x" + f.Max.ToString("X") + ").", ib);
                        if (haveLast && key <= last) { Err("NDB", tree + "-KEY-ORDER", tree + " keys are not strictly increasing at 0x" + key.ToString("X") + ".", ib); }
                        last = key; haveLast = true;
                        anyLeaf = true;
                        leaf(page, i * cb);
                    }
                }
                else
                {
                    var keys = new ulong[cEnt];
                    for (int i = 0; i < cEnt; i++)
                    {
                        keys[i] = PstFormat.G64(page, i * cb);
                        if (i > 0 && keys[i] <= keys[i - 1]) Err("NDB", tree + "-KEY-ORDER", tree + " index keys are not strictly increasing at 0x" + keys[i].ToString("X") + ".", ib);
                    }
                    for (int i = cEnt - 1; i >= 0; i--)
                    {
                        ulong max = i + 1 < cEnt ? keys[i + 1] : f.Max;
                        stack.Push(new Frame { Ref = new PstBref { Bid = PstFormat.G64(page, i * cb + 8), Ib = PstFormat.G64(page, i * cb + 16) }, Level = level - 1, Min = keys[i], Max = max });
                    }
                }
            }
            if (!anyLeaf) Info("NDB", tree + "-EMPTY", tree + " has no entry.");
        }

        void WalkBbt()
        {
            var buf = new byte[PstFormat.MaxBlock];
            long verified = 0, blocksSeen = 0;
            WalkTree(_bbtRoot, PstFormat.PtypeBBT, "BBT", 24, (p, o) =>
            {
                ulong bid = PstFormat.G64(p, o);
                long ib = (long)PstFormat.G64(p, o + 8);
                int cb = PstFormat.G16(p, o + 16), cref = PstFormat.G16(p, o + 18);
                if (bid == 0) Err("NDB", "BBT-BID0", "A block has BID 0.", ib);
                if ((bid & 1) != 0) Warn("NDB", "BID-RESERVED-BIT", "BBT entry with BID 0x" + bid.ToString("X") + ": the reserved bit 0 is set (readers ignore it, a writer leaves it zero).", ib, bid);
                if (_bBid.Count > 0 && bid <= _bBid[_bBid.Count - 1]) _bbtSorted = false;
                _bBid.Add(bid); _bIb.Add(ib); _bCb.Add((ushort)cb); _bCref.Add((ushort)cref); _bRefs.Add(0);
                if (bid >= _bidNextB && _bidNextB != 0 && (bid & ~3UL) >= _bidNextB) Warn("NDB", "BBT-BID-NEXT", "BID 0x" + bid.ToString("X") + " is not below bidNextB 0x" + _bidNextB.ToString("X") + ".", ib, bid);
                if (cb > PstFormat.MaxBlockData) { Err("NDB", "BLOCK-SIZE", "Block size " + cb + " exceeds 8176 bytes.", ib, bid); return; }
                int total = PstFormat.BlockSize(cb);
                if (ib % 64 != 0) { Err("NDB", "BLOCK-ALIGN", "Block is not aligned on 64 bytes.", ib, bid); return; }
                if (ib < First || ib + total > _len) { Err("NDB", "BLOCK-RANGE", "Block lies outside the file.", ib, bid); return; }
                MarkUsed(ib, total, "Block", bid);
                if (!_o.VerifyBlocks) return;
                _o.Cancellation.ThrowIfCancellationRequested();
                if (!ReadAt(ib, buf, total)) { Err("NDB", "BLOCK-READ", "Block cannot be read.", ib, bid); return; }
                verified += total;
                if (_o.ProgressReport != null && (++blocksSeen & 0x3FFF) == 0) _o.ProgressReport.Report(new PstProgress { Phase = "block B-tree", BytesDone = verified, BytesTotal = _len, Items = blocksSeen });
                int t = total - PstFormat.BlockTrailerSize;
                if (PstFormat.G16(buf, t) != cb) Err("NDB", "BLOCK-CB", "Trailer cb " + PstFormat.G16(buf, t) + " differs from the BBT (" + cb + ").", ib, bid);
                if (PstFormat.G64(buf, t + 8) != bid) Err("NDB", "BLOCK-BID", "Trailer BID 0x" + PstFormat.G64(buf, t + 8).ToString("X") + " differs from the BBT.", ib, bid);
                if (PstFormat.G16(buf, t + 2) != PstFormat.Sig((ulong)ib, bid)) Err("NDB", "BLOCK-SIG", "Block signature mismatch.", ib, bid);
                if (PstFormat.Crc(buf, 0, Math.Min(cb, t)) != PstFormat.G32(buf, t + 4)) Err("NDB", "BLOCK-CRC", "Block CRC mismatch: the data is corrupt.", ib, bid);
            });
            Report.Stats["bytesVerified"] = verified;
            if (!_bbtSorted) Err("NDB", "BBT-ORDER", "The BBT leaf entries are not in increasing BID order.");
        }

        int FindBlock(ulong bid)
        {
            if (!_bbtSorted) return -1;
            bid &= ~1UL;       // bit 0 of a BID is reserved and ignored by readers
            int i = _bBid.BinarySearch(bid);
            return i >= 0 ? i : -1;
        }

        void WalkNbt()
        {
            WalkTree(_nbtRoot, PstFormat.PtypeNBT, "NBT", 32, (p, o) =>
            {
                ulong nid = PstFormat.G64(p, o);
                if (nid > uint.MaxValue) { Err("NDB", "NBT-NID", "NID 0x" + nid.ToString("X") + " does not fit in 32 bits."); nid &= 0xFFFFFFFF; }
                if (_nNid.Count > 0 && (uint)nid <= _nNid[_nNid.Count - 1]) _nbtSorted = false;
                _nNid.Add((uint)nid); _nData.Add(PstFormat.G64(p, o + 8)); _nSub.Add(PstFormat.G64(p, o + 16)); _nParent.Add(PstFormat.G32(p, o + 24));
            });
        }

        // The header keeps, for each NID type, the highest index in use (Outlook validates it, "header NID high-water marks")
        void CheckHighWater()
        {
            var max = new uint[32];
            for (int i = 0; i < _nNid.Count; i++) { uint t = _nNid[i] & 0x1F, ix = _nNid[i] >> 5; if (ix > max[t]) max[t] = ix; }
            foreach (uint t in new uint[] { 2, 3, 4, 8, 0x0D, 0x0E, 0x0F, 0x10 })
                if (max[t] > _rgnid[t] && max[t] >= 0x400) Warn("NDB", "HDR-NID-HIGHWATER", "The header says the highest NID index of type 0x" + t.ToString("X") + " is 0x" + _rgnid[t].ToString("X") + " but 0x" + max[t].ToString("X") + " is in use: the next NID allocated could collide.", 44 + 4 * (int)t);
        }

        // ---------------------------------------------------------------- data trees and subnode trees
        void Ref(ulong bid, uint nid, string what)
        {
            if ((bid & 1) != 0) Warn("NDB", "BID-RESERVED-BIT", what + " refers to BID 0x" + bid.ToString("X") + " whose reserved bit 0 is set (readers ignore it, a writer leaves it zero).", -1, bid, nid);
            int i = FindBlock(bid);
            if (i < 0) { Err("NDB", "BID-DANGLING", what + " refers to BID 0x" + bid.ToString("X") + " which is not in the BBT.", -1, bid, nid); return; }
            if (_bRefs[i] < ushort.MaxValue) _bRefs[i]++;
        }

        byte[] ReadInternal(int idx)
        {
            int cb = _bCb[idx];
            var b = new byte[PstFormat.BlockSize(cb)];
            if (_bIb[idx] < 0 || !ReadAt(_bIb[idx], b, b.Length)) return null;
            var d = new byte[cb];
            Buffer.BlockCopy(b, 0, d, 0, cb);
            return d;
        }

        long CheckDataTree(ulong bid, uint nid, string what)
        {
            if (bid == 0) return 0;
            int i = FindBlock(bid);
            Ref(bid, nid, what);
            if (i < 0) return 0;
            if ((bid & 2) == 0) return _bCb[i];
            var leaves = new List<int>();
            byte[] b = ReadInternal(i);
            if (b == null || b.Length < 8) { Err("NDB", "XBLOCK-READ", "Internal block cannot be read.", _bIb[i], bid, nid); return 0; }
            if (b[0] != 1) { Err("NDB", "XBLOCK-TYPE", "Internal block of a data tree has btype " + b[0] + " (1 expected).", _bIb[i], bid, nid); return 0; }
            int level = b[1], cEnt = PstFormat.G16(b, 2);
            long total = 0;
            if ((level != 1 && level != 2) || 8 + 8 * cEnt > b.Length) { Err("NDB", "XBLOCK-SHAPE", "XBLOCK level " + level + " with " + cEnt + " entries in " + b.Length + " bytes.", _bIb[i], bid, nid); return 0; }
            if (cEnt == 0) Err("NDB", "XBLOCK-EMPTY", "XBLOCK without entry.", _bIb[i], bid, nid);
            for (int k = 0; k < cEnt; k++)
            {
                ulong child = PstFormat.G64(b, 8 + 8 * k);
                if (level == 1)
                {
                    int ci = FindBlock(child);
                    Ref(child, nid, "XBLOCK");
                    if (ci < 0) continue;
                    if ((child & 2) != 0) Err("NDB", "XBLOCK-CHILD", "XBLOCK refers to an internal block.", _bIb[ci], child, nid);
                    total += _bCb[ci]; leaves.Add(_bCb[ci]);
                }
                else
                {
                    int ci = FindBlock(child);
                    Ref(child, nid, "XXBLOCK");
                    if (ci < 0) continue;
                    if ((child & 2) == 0) { Err("NDB", "XBLOCK-CHILD", "XXBLOCK refers to a data block.", _bIb[ci], child, nid); continue; }
                    byte[] cx = ReadInternal(ci);
                    if (cx == null || cx.Length < 8 || cx[0] != 1 || cx[1] != 1 || 8 + 8 * PstFormat.G16(cx, 2) > cx.Length) { Err("NDB", "XBLOCK-SHAPE", "XXBLOCK child is not a valid XBLOCK.", _bIb[ci], child, nid); continue; }
                    long sub = 0;
                    int ce = PstFormat.G16(cx, 2);
                    for (int m = 0; m < ce; m++)
                    {
                        ulong leafBid = PstFormat.G64(cx, 8 + 8 * m);
                        int li = FindBlock(leafBid);
                        Ref(leafBid, nid, "XBLOCK");
                        if (li >= 0) { if ((leafBid & 2) != 0) Err("NDB", "XBLOCK-CHILD", "XBLOCK refers to an internal block.", _bIb[li], leafBid, nid); sub += _bCb[li]; leaves.Add(_bCb[li]); }
                    }
                    if (sub != PstFormat.G32(cx, 4)) Err("NDB", "XBLOCK-TOTAL", "XBLOCK lcbTotal " + PstFormat.G32(cx, 4) + " differs from the sum of its data blocks (" + sub + ").", _bIb[ci], child, nid);
                    total += sub;
                }
            }
            if (total != PstFormat.G32(b, 4)) Err("NDB", "XBLOCK-TOTAL", "lcbTotal " + PstFormat.G32(b, 4) + " differs from the sum of the data blocks (" + total + ").", _bIb[i], bid, nid);
            for (int k = 0; k + 1 < leaves.Count; k++)
                if (leaves[k] != PstFormat.MaxBlockData) { Warn("NDB", "DATATREE-NOT-FULL", "Data block " + k + " of " + leaves.Count + " holds " + leaves[k] + " bytes: every block of a data tree but the last is full (8176) in the files of Outlook (the specification does not say it; a validation tool of Outlook rejects the node).", _bIb[i], bid, nid); break; }
            return total;
        }

        void CheckSubnodes(ulong bid, uint nid, int depth, HashSet<ulong> path)
        {
            if (bid == 0) return;
            if (depth > 8 || !path.Add(bid)) { Err("NDB", "SUBNODE-CYCLE", "Subnode trees nest too deeply or loop.", -1, bid, nid); return; }
            int i = FindBlock(bid);
            Ref(bid, nid, "Subnode tree");
            if (i < 0) { path.Remove(bid); return; }
            if ((bid & 2) == 0) { Err("NDB", "SUBNODE-INTERNAL", "A subnode block must be internal.", _bIb[i], bid, nid); path.Remove(bid); return; }
            byte[] b = ReadInternal(i);
            if (b == null || b.Length < 8 || b[0] != 2) { Err("NDB", "SUBNODE-TYPE", "Block is not an SLBLOCK / SIBLOCK.", _bIb[i], bid, nid); path.Remove(bid); return; }
            int level = b[1], cEnt = PstFormat.G16(b, 2);
            if (cEnt == 0) Err("NDB", "SUBNODE-EMPTY", "Subnode block without entry.", _bIb[i], bid, nid);
            if (level == 0)
            {
                if (8 + 24 * cEnt > b.Length) { Err("NDB", "SUBNODE-SHAPE", "SLBLOCK with " + cEnt + " entries in " + b.Length + " bytes.", _bIb[i], bid, nid); path.Remove(bid); return; }
                ulong prev = 0; bool have = false;
                for (int k = 0; k < cEnt; k++)
                {
                    ulong sn = PstFormat.G64(b, 8 + 24 * k);
                    if (have && sn <= prev) Err("NDB", "SUBNODE-ORDER", "Subnode NIDs are not strictly increasing.", _bIb[i], bid, nid);
                    prev = sn; have = true;
                    ulong sd = PstFormat.G64(b, 16 + 24 * k), ss = PstFormat.G64(b, 24 + 24 * k);
                    CheckDataTree(sd, nid, "Subnode 0x" + sn.ToString("X"));
                    CheckSubnodes(ss, nid, depth + 1, path);
                }
            }
            else if (level == 1)
            {
                if (8 + 16 * cEnt > b.Length) { Err("NDB", "SUBNODE-SHAPE", "SIBLOCK with " + cEnt + " entries in " + b.Length + " bytes.", _bIb[i], bid, nid); path.Remove(bid); return; }
                for (int k = 0; k < cEnt; k++)
                {
                    ulong child = PstFormat.G64(b, 16 + 16 * k);
                    int ci = FindBlock(child);
                    if (ci < 0) { Ref(child, nid, "SIBLOCK"); continue; }
                    // a child of an SIBLOCK is an SLBLOCK (level 0)
                    byte[] cb = ReadInternal(ci);
                    if (cb == null || cb.Length < 8 || cb[0] != 2 || cb[1] != 0) { Err("NDB", "SUBNODE-SHAPE", "SIBLOCK child is not an SLBLOCK.", _bIb[ci], child, nid); Ref(child, nid, "SIBLOCK"); continue; }
                    CheckSubnodes(child, nid, depth + 1, path);
                }
            }
            else Err("NDB", "SUBNODE-LEVEL", "Subnode block level " + level + ".", _bIb[i], bid, nid);
            path.Remove(bid);
        }

        void Trees()
        {
            if (!_bbtSorted || !_nbtSorted) { Err("NDB", "TREE-SKIP", "Data trees were not checked: the B-trees are not in key order."); return; }
            var path = new HashSet<ulong>();
            for (int n = 0; n < _nNid.Count; n++)
            {
                uint nid = _nNid[n];
                CheckDataTree(_nData[n], nid, "Node data");
                if (_nSub[n] != 0) { path.Clear(); CheckSubnodes(_nSub[n], nid, 0, path); }
            }
            int orphans = 0; long cref = 0;
            for (int i = 0; i < _bBid.Count; i++)
            {
                if (_bRefs[i] == 0) { orphans++; Warn("NDB", "BLOCK-ORPHAN", "Block is referenced by no node, subnode or data tree (wasted space).", _bIb[i], _bBid[i]); }
                else if (_bCref[i] < _bRefs[i] + 1) { cref++; Warn("NDB", "BLOCK-CREF", "Reference count " + _bCref[i] + " is lower than the number of references (" + _bRefs[i] + " + the BBT entry).", _bIb[i], _bBid[i]); }
            }
            Report.Stats["orphanBlocks"] = orphans;
        }

        // ---------------------------------------------------------------- LTP and messaging
        int FindNode(uint nid)
        {
            int i = _nNid.BinarySearch(nid);
            return i >= 0 ? i : -1;
        }

        string Hex(uint nid) { return "0x" + nid.ToString("X"); }

        void Messaging()
        {
            if (!_bbtSorted || !_nbtSorted) { Err("Messaging", "MSG-SKIP", "The LTP and messaging layers were not checked: the B-trees are not in key order."); return; }
            try { _pf = PstFile.Open(_path); }
            catch (Exception e) { Err("Messaging", "MSG-OPEN", "The file cannot be opened by the reader (" + e.Message + "); the LTP and messaging layers were not checked."); return; }
            var listed = new bool[_nNid.Count];

            // mandatory nodes
            uint[] mandatory = { 0x21, 0x61, 0x122, 0x12D, 0x12E, 0x12F, 0x60D, 0x60E, 0x60F, 0x610, 0x671, 0x692 };
            foreach (uint m in mandatory)
                if (FindNode(m) < 0) Err("Messaging", "MSG-MANDATORY", "Mandatory node " + Hex(m) + " is missing.", -1, 0, m);
            if (FindNode(0x1E1) < 0) Warn("Messaging", "MSG-MANDATORY", "Search management queue 0x1E1 is missing.", -1, 0, 0x1E1);
            if (FindNode(0x201) < 0) Warn("Messaging", "MSG-MANDATORY", "Search activity list 0x201 is missing.", -1, 0, 0x201);

            if (!_nNid.Exists(n => (n & 0x1F) == PstFormat.NtReceiveFolderTable)) Warn("Messaging", "MSG-RECEIVE-TABLE", "The receive folder table (type 0x0B) is missing.");
            else CheckReceiveTable();
            if (!_nNid.Exists(n => (n & 0x1F) == PstFormat.NtOutgoingQueueTable)) Info("Messaging", "MSG-OUTGOING-QUEUE", "The outgoing queue table (type 0x0C) is missing.");
            foreach (uint t in new uint[] { 0x6B6, 0x6D7, 0x6F8 })
                if (FindNode(t) < 0) Info("Messaging", "MSG-TEMPLATE-MISSING", "Template table " + Hex(t) + " is missing.", -1, 0, t);
            if (FindNode(0xEC1) < 0) Info("Messaging", "MSG-NODE-EC1", "Node 0xEC1 (empty internal node) is missing.", -1, 0, 0xEC1);
            LoadHmp();
            // required columns of the templates (beyond the specification table: Outlook rejects a template without them)
            var req = new Dictionary<uint, uint[]> { { 0x60D, new uint[] { 0x0E300102 } }, { 0x60E, new uint[] { 0x0E300102, 0x30130102 } }, { 0x60F, new uint[] { 0x682F001F } }, { 0x610, new uint[] { 0x0058000B } } };
            foreach (KeyValuePair<uint, uint[]> kv in req)
            {
                if (FindNode(kv.Key) < 0) continue;
                PstTcReader rt0 = Guard(() => new PstTcReader(_pf.OpenNode(kv.Key)), kv.Key, "template table");
                if (rt0 == null) continue;
                foreach (uint tag in kv.Value) if (!rt0.Columns.Exists(c => c.Tag == tag)) Info("Messaging", "TEMPLATE-COLUMN", "Template table " + Hex(kv.Key) + " lacks the column 0x" + tag.ToString("X8") + " that the template requires in Outlook.", -1, 0, kv.Key);
            }
            // the template tables carry no row
            foreach (uint t in new uint[] { 0x60D, 0x60E, 0x60F, 0x610, 0x671, 0x692 })
            {
                if (FindNode(t) < 0) continue;
                PstTcReader tc = Guard(() => new PstTcReader(_pf.OpenNode(t)), t, "template table");
                if (tc != null) { CheckTc(tc, t, "template " + Hex(t)); if (tc.RowCount != 0) Warn("Messaging", "TEMPLATE-ROWS", "Template table " + Hex(t) + " has " + tc.RowCount + " row(s) (none allowed).", -1, 0, t); }
            }

            string storeProblem = null;
            byte[] recordKey = null;
            if (FindNode(0x21) >= 0) recordKey = CheckStore(out storeProblem);
            if (FindNode(0x61) >= 0) CheckNameMap();

            // folders
            var folders = new List<int>();
            for (int i = 0; i < _nNid.Count; i++)
            {
                uint type = PstFormat.NidType(_nNid[i]);
                if (type == PstFormat.NtNormalFolder || type == PstFormat.NtSearchFolder) folders.Add(i);
            }
            Report.Stats["folders"] = folders.Count;
            var hierarchyChildren = new Dictionary<uint, uint>();
            var folderInfo = new Dictionary<uint, string>();
            foreach (int fi in folders)
            {
                uint fnid = _nNid[fi];
                folderInfo[fnid] = FolderName(fnid);
                _folderVals[fnid] = FolderValues(fnid);
            }
            var checkedMsgs = new List<KeyValuePair<int, string>>();
            foreach (int fi in folders)
            {
                uint fnid = _nNid[fi];
                string path = PathOf(fnid, folderInfo);
                CheckFolder(fi, path, folderInfo, listed, hierarchyChildren, checkedMsgs);
            }
            // folders: listed in a hierarchy table of their parent, parent chain ends at the root without loop
            foreach (int fi in folders)
            {
                uint fnid = _nNid[fi];
                if (fnid == PstFormat.NidRootFolder)
                {
                    if (_nParent[fi] != fnid) Warn("Messaging", "ROOT-PARENT", "The root folder's nidParent is not itself.", -1, 0, fnid);
                    continue;
                }
                uint parent = _nParent[fi];
                if (parent == 0 || FindNode(parent) < 0) { Err("Messaging", "FOLDER-PARENT", "Folder '" + folderInfo[fnid] + "' has parent " + Hex(parent) + " which is not a folder of the file.", -1, 0, fnid, folderInfo[fnid]); continue; }
                uint listedBy;
                if (!hierarchyChildren.TryGetValue(fnid, out listedBy)) Err("Messaging", "FOLDER-NOT-LISTED", "Folder '" + folderInfo[fnid] + "' is not listed in the hierarchy table of any folder (unreachable).", -1, 0, fnid, folderInfo[fnid]);
                else if (listedBy != parent) Err("Messaging", "FOLDER-PARENT-MISMATCH", "Folder '" + folderInfo[fnid] + "' is listed by " + Hex(listedBy) + " but its nidParent is " + Hex(parent) + ".", -1, 0, fnid, folderInfo[fnid]);
                uint cur = fnid; int hops = 0;
                while (cur != PstFormat.NidRootFolder && hops++ < 80)
                {
                    int ci = FindNode(cur);
                    if (ci < 0) break;
                    cur = _nParent[ci];
                    if (cur == 0) break;
                }
                if (hops >= 80) Err("Messaging", "FOLDER-LOOP", "The parent chain of folder '" + folderInfo[fnid] + "' loops or is deeper than 80 levels.", -1, 0, fnid, folderInfo[fnid]);
            }
            // messages not in any contents table
            long messages = 0, assoc = 0;
            for (int i = 0; i < _nNid.Count; i++)
            {
                uint type = PstFormat.NidType(_nNid[i]);
                if (type == PstFormat.NtNormalMessage || type == PstFormat.NtAssocMessage)
                {
                    if (type == PstFormat.NtNormalMessage) messages++; else assoc++;
                    if (!listed[i]) Err("Messaging", "MSG-NOT-LISTED", "Message " + Hex(_nNid[i]) + " (parent " + Hex(_nParent[i]) + ") is not in the contents table of its folder: it is invisible.", -1, 0, _nNid[i]);
                }
                else if (type == PstFormat.NtAttachment || type == PstFormat.NtRecipientTable || type == PstFormat.NtAttachmentTable)
                {
                    if (_nNid[i] != 0x671 && _nNid[i] != 0x692) Warn("Messaging", "NODE-STRAY", "Node " + Hex(_nNid[i]) + " of a type that only exists as a subnode is at the top level.", -1, 0, _nNid[i]);
                }
            }
            Report.Stats["messages"] = messages;
            Report.Stats["associatedMessages"] = assoc;
            if (_o.Deep)
            {
                long n = 0;
                foreach (KeyValuePair<int, string> kv in checkedMsgs)
                {
                    _o.Cancellation.ThrowIfCancellationRequested();
                    CheckMessageNode(kv.Key, kv.Value);
                    if ((++n % 5000) == 0) Progress("messages " + n + " / " + checkedMsgs.Count);
                    if (_o.ProgressReport != null && (n % 500) == 0) _o.ProgressReport.Report(new PstProgress { Phase = "messages", Items = n, ItemsTotal = checkedMsgs.Count, BytesTotal = _len });
                }
                CheckSubnodeHighWater();
            }
        }

        // receive folder table: a row for the empty message class (the default folder) whose 0x6605 is the NID of a folder
        void CheckReceiveTable()
        {
            foreach (uint rn in _nNid.Where(n => (n & 0x1F) == PstFormat.NtReceiveFolderTable))
            {
                PstTcReader tc = Guard(() => new PstTcReader(_pf.OpenNode(rn)), rn, "receive folder table");
                if (tc == null) continue;
                CheckTc(tc, rn, "receive folder table");
                bool dflt = false;
                for (int r = 0; r < tc.RowCount; r++)
                {
                    int row = r;
                    PstProp cls = Guard(() => tc.GetCell(row, 0x001A), rn, "receive folder row"), fo = Guard(() => tc.GetCell(row, 0x6605), rn, "receive folder row");
                    uint target = fo == null || fo.Data == null || fo.Data.Length < 4 ? 0 : BitConverter.ToUInt32(fo.Data, 0);
                    if (target == 0 || FindNode(target) < 0 || PstFormat.NidType(target) != PstFormat.NtNormalFolder) Info("Messaging", "RECEIVE-TABLE-ROW", "Receive folder row " + r + " does not refer to a folder of the file (0x6605 = " + Hex(target) + ").", -1, 0, rn);
                    if (cls != null && cls.Data != null && cls.Data.Length == 0) dflt = true;
                }
                if (!dflt) Info("Messaging", "RECEIVE-TABLE-ROW", "The receive folder table has no row for the empty message class (the default receive folder).", -1, 0, rn);
            }
        }

        // node 0xC01: BTH (key 16 bytes, entry 4 bytes) from the replication GUID of an object to its NID
        void LoadHmp()
        {
            if (FindNode(0xC01) < 0) { Info("Messaging", "HMP-MISSING", "Node 0xC01 (GUID to NID map of the objects) is missing.", -1, 0, 0xC01); return; }
            var map = Guard(() =>
            {
                var ltp = new PstLtpNode(_pf.OpenNode(0xC01));
                if (ltp.Heap.ClientSig != 0x9C) throw new PstFormatException("client signature 0x" + ltp.Heap.ClientSig.ToString("X2") + " (0x9C expected)");
                int cbKey, cbEnt;
                List<byte[]> recs = PstBthReader.Records(ltp.Heap, PstFormat.G32(ltp.Heap.Get(ltp.Heap.UserRoot), 0), out cbKey, out cbEnt);
                if (cbKey != 16 || cbEnt != 4) throw new PstFormatException("BTH with key " + cbKey + " / entry " + cbEnt + " bytes (16 / 4 expected)");
                var d = new Dictionary<string, uint>();
                foreach (byte[] r in recs) d[BitConverter.ToString(r, 0, 16)] = PstFormat.G32(r, 16);
                return d;
            }, 0xC01, "GUID to NID map");
            if (map == null) return;
            _hmp = map;
            foreach (KeyValuePair<string, uint> kv in map)
            {
                uint t = PstFormat.NidType(kv.Value);
                if (FindNode(kv.Value) < 0 || (t != PstFormat.NtNormalMessage && t != PstFormat.NtNormalFolder && t != PstFormat.NtSearchFolder)) { Info("Messaging", "HMP-ENTRY", "The GUID to NID map refers to " + Hex(kv.Value) + " which is not a message or folder of the file.", -1, 0, 0xC01); break; }
            }
        }

        void CheckHmpRow(PstTcReader tc, int r, uint nid, uint tn, string path)
        {
            if (_hmp == null) return;
            PstTcReader.Col c = tc.FindColumn(0x0E30);
            if (c == null) return;
            PstProp g = Guard(() => tc.GetCell(r, c), tn, "replication id cell");
            if (g == null || g.Data == null || g.Data.Length != 16) { Info("Messaging", "HMP-ENTRY", "Row of " + Hex(nid) + " has no replication GUID (PidTagReplItemid).", -1, 0, tn, path); return; }
            uint to;
            if (!_hmp.TryGetValue(BitConverter.ToString(g.Data), out to)) Info("Messaging", "HMP-ENTRY", "The replication GUID of " + Hex(nid) + " is not in the GUID to NID map (0xC01).", -1, 0, tn, path);
            else if (to != nid) Info("Messaging", "HMP-ENTRY", "The replication GUID of " + Hex(nid) + " maps to " + Hex(to) + " in 0xC01.", -1, 0, tn, path);
        }

        static bool SameValue(PstProp a, PstProp b)
        {
            if (a.Items != null || b.Items != null)
            {
                if (a.Items == null || b.Items == null || a.Items.Count != b.Items.Count) return false;
                for (int i = 0; i < a.Items.Count; i++) if (!a.Items[i].SequenceEqual(b.Items[i])) return false;
                return true;
            }
            if (a.Data == null || b.Data == null) return a.Data == b.Data;
            return a.Data.SequenceEqual(b.Data);
        }

        // the cells of a contents row repeat properties of the message: they must be the same values
        void CompareRow(PstTcReader tc, int r, int mi, uint tn, string path)
        {
            try
            {
                var pc = new PstPcReader(_pf.OpenNode(_nNid[mi]));
                foreach (PstTcReader.Col c in tc.Columns)
                {
                    if (c.Id == 0x67F2 || c.Id == 0x67F3) continue;
                    PstPcReader.Entry e = pc.Find(c.Id);
                    if (e == null) continue;
                    PstProp cell = tc.GetCell(r, c);
                    if (e.Type != c.Type) { Warn("Messaging", "CONTENTS-CELL", "Column 0x" + c.Tag.ToString("X8") + " of the row of " + Hex(_nNid[mi]) + " has the type 0x" + c.Type.ToString("X4") + " but the message has 0x" + e.Type.ToString("X4") + ".", -1, 0, tn, path); continue; }
                    if (cell == null) { Warn("Messaging", "CONTENTS-CELL", "The row of " + Hex(_nNid[mi]) + " has no cell for the column 0x" + c.Tag.ToString("X8") + " that the message has.", -1, 0, tn, path); continue; }
                    if (!SameValue(cell, pc.Decode(e))) Warn("Messaging", "CONTENTS-CELL", "The cell 0x" + c.Tag.ToString("X8") + " of the row of " + Hex(_nNid[mi]) + " differs from the property of the message.", -1, 0, tn, path);
                }
            }
            catch (PstFormatException ex) { Err("Messaging", "MSG-BROKEN", "Message cannot be read to compare its contents row: " + ex.Message, ex.Offset, ex.Bid, _nNid[mi], path); }
            catch (Exception ex) { Err("Messaging", "MSG-BROKEN", "Message cannot be read to compare its contents row (" + ex.GetType().Name + "): " + ex.Message, -1, 0, _nNid[mi], path); }
        }

        void TrackSubnode(uint sn)
        {
            uint t = sn & 0x1F, ix = sn >> 5;
            if ((t == PstFormat.NtAttachment || t == PstFormat.NtNormalMessage || t == PstFormat.NtAssocMessage) && ix > _subMax[t]) _subMax[t] = ix;
        }

        // attachment and embedded message NIDs live in subnode trees, outside the NBT: they come from the same counters of the header
        void CheckSubnodeHighWater()
        {
            foreach (uint t in new uint[] { PstFormat.NtAttachment, PstFormat.NtNormalMessage, PstFormat.NtAssocMessage })
                if (_subMax[t] > _rgnid[t] && _subMax[t] >= 0x400) Warn("NDB", "HDR-NID-HIGHWATER", "The header says the highest NID index of type 0x" + t.ToString("X") + " is 0x" + _rgnid[t].ToString("X") + " but 0x" + _subMax[t].ToString("X") + " is in use as a subnode (attachment or embedded message): the next NID allocated could collide.", 44 + 4 * (int)t);
        }

        T Guard<T>(Func<T> f, uint nid, string what) where T : class
        {
            try { return f(); }
            catch (PstFormatException e) { Err("LTP", "LTP-BROKEN", "The " + what + " cannot be read: " + e.Message, e.Offset, e.Bid, nid); return null; }
            catch (Exception e) { Err("LTP", "LTP-BROKEN", "The " + what + " cannot be read (" + e.GetType().Name + "): " + e.Message, -1, 0, nid); return null; }
        }

        string FolderName(uint nid)
        {
            try
            {
                var pc = new PstPcReader(_pf.OpenNode(nid));
                return PstConvert.StringOf(pc.Get(0x3001)) ?? "";
            }
            catch (Exception) { return "?"; }
        }

        // content count, unread count, has-subfolders (1/0), hidden count, hidden unread, associated count (-1: absent) of a folder
        int[] FolderValues(uint nid)
        {
            var v = new[] { -1, -1, -1, -1, -1, -1 };
            try
            {
                var pc = new PstPcReader(_pf.OpenNode(nid));
                ushort[] ids = { 0x3602, 0x3603, 0x360A, 0x6635, 0x6636, 0x3617 };
                for (int i = 0; i < ids.Length; i++)
                {
                    PstProp p = pc.Get(ids[i]);
                    if (p == null || p.Data == null || p.Data.Length == 0) continue;
                    v[i] = ids[i] == 0x360A ? (p.Data[0] != 0 ? 1 : 0) : PstConvert.Int32Of(p, -1);
                }
            }
            catch (Exception) { }
            return v;
        }

        string PathOf(uint nid, Dictionary<uint, string> names)
        {
            var parts = new List<string>();
            uint cur = nid; int hops = 0;
            while (hops++ < 40)
            {
                string n;
                if (!names.TryGetValue(cur, out n)) break;
                if (cur != PstFormat.NidRootFolder) parts.Add(n);
                int i = FindNode(cur);
                if (i < 0 || _nParent[i] == cur || _nParent[i] == 0) break;
                cur = _nParent[i];
            }
            parts.Reverse();
            return "/" + string.Join("/", parts);
        }

        // ---- store, name-to-id map
        byte[] CheckStore(out string problem)
        {
            problem = null;
            PstPcReader pc = Guard(() => new PstPcReader(_pf.OpenNode(0x21)), 0x21, "message store");
            if (pc == null) return null;
            CheckPc(pc, 0x21, "message store");
            byte[] rk = null;
            Func<ushort, int, string, byte[]> need = (id, size, name) =>
            {
                PstProp p = Guard(() => pc.Get(id), 0x21, "store property " + name);
                if (p == null || p.Data == null) { Err("Messaging", "STORE-PROP", "The message store lacks the required property " + name + " (0x" + id.ToString("X4") + ").", -1, 0, 0x21); return null; }
                if (size > 0 && p.Data.Length != size) Err("Messaging", "STORE-PROP-SIZE", "Store property " + name + " has " + p.Data.Length + " bytes (" + size + " expected).", -1, 0, 0x21);
                return p.Data;
            };
            rk = need(0x0FF9, 16, "PidTagRecordKey");
            need(0x3001, 0, "PidTagDisplayName");
            var eids = new[] { new KeyValuePair<ushort, string>(0x35E0, "PidTagIpmSubTreeEntryId"), new KeyValuePair<ushort, string>(0x35E3, "PidTagIpmWastebasketEntryId"), new KeyValuePair<ushort, string>(0x35E7, "PidTagFinderEntryId") };
            foreach (var e in eids)
            {
                byte[] d = need(e.Key, 24, e.Value);
                if (d == null || d.Length != 24) continue;
                uint nid = PstFormat.G32(d, 20);
                if (FindNode(nid) < 0) Err("Messaging", "STORE-ENTRYID", e.Value + " refers to " + Hex(nid) + " which does not exist.", -1, 0, 0x21);
                else if (PstFormat.NidType(nid) != PstFormat.NtNormalFolder) Err("Messaging", "STORE-ENTRYID", e.Value + " refers to " + Hex(nid) + " which is not a folder.", -1, 0, 0x21);
                if (rk != null && rk.Length == 16)
                    for (int i = 0; i < 16; i++) if (d[4 + i] != rk[i]) { Err("Messaging", "STORE-ENTRYID-UID", e.Value + " does not carry the record key of the store.", -1, 0, 0x21); break; }
            }
            PstProp pw = Guard(() => pc.Get(0x67FF), 0x21, "PidTagPstPassword");
            if (pw != null && pw.Data != null && BitConverter.ToUInt32(pw.Data, 0) != 0) Info("Messaging", "STORE-PASSWORD", "The store is password protected (CRC 0x" + BitConverter.ToUInt32(pw.Data, 0).ToString("X") + ").", -1, 0, 0x21);
            return rk;
        }

        void CheckNameMap()
        {
            PstPcReader pc = Guard(() => new PstPcReader(_pf.OpenNode(0x61)), 0x61, "name-to-id map");
            if (pc == null) return;
            CheckPc(pc, 0x61, "name-to-id map");
            PstProp bc = Guard(() => pc.Get(0x0001), 0x61, "bucket count");
            if (bc == null || bc.Data == null || bc.Data.Length < 4) { Err("Messaging", "NAMEMAP-BUCKETS", "PidTagNameidBucketCount is missing.", -1, 0, 0x61); return; }
            int buckets = BitConverter.ToInt32(bc.Data, 0);
            if (buckets <= 0 || buckets > 0x1000) { Err("Messaging", "NAMEMAP-BUCKETS", "PidTagNameidBucketCount " + buckets + " is not valid.", -1, 0, 0x61); return; }
            PstProp g = Guard(() => pc.Get(0x0002), 0x61, "GUID stream"), e = Guard(() => pc.Get(0x0003), 0x61, "entry stream"), s = Guard(() => pc.Get(0x0004), 0x61, "string stream");
            int entries = 0;
            if (e != null && e.Data != null)
            {
                if (e.Data.Length % 8 != 0) Err("Messaging", "NAMEMAP-ENTRIES", "The entry stream has " + e.Data.Length + " bytes (not a multiple of 8).", -1, 0, 0x61);
                entries = e.Data.Length / 8;
                int gcount = g == null || g.Data == null ? 0 : g.Data.Length / 16;
                if (g != null && g.Data != null && g.Data.Length % 16 != 0) Err("Messaging", "NAMEMAP-GUIDS", "The GUID stream has " + g.Data.Length + " bytes (not a multiple of 16).", -1, 0, 0x61);
                var inBuckets = new Dictionary<ulong, int>();
                for (int b = 0; b < buckets; b++)
                {
                    PstProp bp = Guard(() => pc.Get((ushort)(0x1000 + b)), 0x61, "hash bucket");
                    if (bp == null || bp.Data == null) continue;
                    for (int k = 0; k + 8 <= bp.Data.Length; k += 8)
                    {
                        uint dw0 = PstFormat.G32(bp.Data, k), dw1 = PstFormat.G32(bp.Data, k + 4);
                        if ((int)((dw0 ^ (dw1 & 0xFFFF)) % (uint)buckets) != b) Err("Messaging", "NAMEMAP-HASH", "A name is in hash bucket " + b + " but hashes to " + (int)((dw0 ^ (dw1 & 0xFFFF)) % (uint)buckets) + ".", -1, 0, 0x61);
                        inBuckets[dw1 >> 16 | 0] = 1;
                    }
                }
                for (int i = 0; i < entries; i++)
                {
                    uint dw = PstFormat.G32(e.Data, 8 * i);
                    int ng = PstFormat.G16(e.Data, 8 * i + 4), idx = PstFormat.G16(e.Data, 8 * i + 6);
                    int gi = ng >> 1;
                    if (idx != i) Warn("Messaging", "NAMEMAP-INDEX", "Entry " + i + " has wPropIdx " + idx + ".", -1, 0, 0x61);
                    if (gi >= 3 && gi - 3 >= gcount) Err("Messaging", "NAMEMAP-GUIDREF", "Entry " + i + " refers to GUID " + gi + " beyond the GUID stream.", -1, 0, 0x61);
                    if ((ng & 1) != 0)
                    {
                        if (s == null || s.Data == null || dw + 4 > s.Data.Length || dw + 4 + PstFormat.G32(s.Data, (int)dw) > s.Data.Length) Err("Messaging", "NAMEMAP-STRING", "Entry " + i + " refers to a string outside the string stream.", -1, 0, 0x61);
                    }
                    if (!inBuckets.ContainsKey((ulong)idx)) Err("Messaging", "NAMEMAP-BUCKET-MISSING", "Entry " + i + " is in no hash bucket.", -1, 0, 0x61);
                }
            }
            _namedCount = entries;
            Report.Stats["namedProperties"] = entries;
        }

        // ---- folders
        void CheckFolder(int fi, string path, Dictionary<uint, string> names, bool[] listed, Dictionary<uint, uint> hierarchyChildren, List<KeyValuePair<int, string>> msgs)
        {
            uint fnid = _nNid[fi];
            uint idx = PstFormat.NidIndex(fnid);
            bool search = PstFormat.NidType(fnid) == PstFormat.NtSearchFolder;
            PstPcReader pc = Guard(() => new PstPcReader(_pf.OpenNode(fnid)), fnid, "folder property context '" + path + "'");
            int contentCount = 0, unreadCount = 0; bool subFlag = false; bool havePc = pc != null;
            if (pc != null)
            {
                CheckPc(pc, fnid, path);
                PstProp name = Guard(() => pc.Get(0x3001), fnid, "folder name"), cc = Guard(() => pc.Get(0x3602), fnid, "content count"), uc = Guard(() => pc.Get(0x3603), fnid, "unread count"), sf = Guard(() => pc.Get(0x360A), fnid, "subfolders flag");
                if (name == null) Err("Messaging", "FOLDER-PROP", "Folder lacks PidTagDisplayName.", -1, 0, fnid, path);
                if (cc == null) Err("Messaging", "FOLDER-PROP", "Folder lacks PidTagContentCount.", -1, 0, fnid, path); else contentCount = PstConvert.Int32Of(cc, 0);
                if (uc == null) Err("Messaging", "FOLDER-PROP", "Folder lacks PidTagContentUnreadCount.", -1, 0, fnid, path); else unreadCount = PstConvert.Int32Of(uc, 0);
                if (sf == null) Err("Messaging", "FOLDER-PROP", "Folder lacks PidTagSubfolders.", -1, 0, fnid, path); else subFlag = sf.Data != null && sf.Data.Length > 0 && sf.Data[0] != 0;
            }
            // hierarchy table
            uint hn = PstFormat.Nid(PstFormat.NtHierarchyTable, idx);
            int children = 0;
            if (FindNode(hn) < 0) Err("Messaging", "FOLDER-TABLE", "Hierarchy table " + Hex(hn) + " of the folder is missing.", -1, 0, hn, path);
            else
            {
                PstTcReader ht = Guard(() => new PstTcReader(_pf.OpenNode(hn)), hn, "hierarchy table of '" + path + "'");
                if (ht != null)
                {
                    CheckTc(ht, hn, "hierarchy " + path);
                    children = ht.RowCount;
                    for (int r = 0; r < ht.RowCount; r++)
                    {
                        uint child = ht.RowId(r);
                        int ci = FindNode(child);
                        if (ci < 0 || (PstFormat.NidType(child) != PstFormat.NtNormalFolder && PstFormat.NidType(child) != PstFormat.NtSearchFolder)) { Err("Messaging", "HIERARCHY-ROW", "Hierarchy row " + r + " refers to " + Hex(child) + " which is not a folder of the file.", -1, 0, hn, path); continue; }
                        uint prev;
                        if (hierarchyChildren.TryGetValue(child, out prev)) Err("Messaging", "HIERARCHY-DUP", "Folder " + Hex(child) + " is listed by two parents (" + Hex(prev) + " and " + Hex(fnid) + ").", -1, 0, hn, path);
                        hierarchyChildren[child] = fnid;
                        if (_nParent[ci] != fnid) Err("Messaging", "HIERARCHY-PARENT", "Folder " + Hex(child) + " is listed here but its nidParent is " + Hex(_nParent[ci]) + ".", -1, 0, hn, path);
                        // the row repeats properties of the child folder
                        int[] cv;
                        if (_folderVals.TryGetValue(child, out cv))
                        {
                            ushort[] ids = { 0x3602, 0x3603, 0x360A, 0x6635, 0x6636, 0x3617 };
                            for (int k = 0; k < ids.Length; k++)
                            {
                                int kk = k, row = r;
                                PstProp cell = Guard(() => ht.GetCell(row, ids[kk]), hn, "hierarchy row");
                                if (cell == null || cell.Data == null || cell.Data.Length == 0) continue;
                                int rv = ids[k] == 0x360A ? (cell.Data[0] != 0 ? 1 : 0) : PstConvert.Int32Of(cell, -1);
                                if (cv[k] >= 0 && rv != cv[k]) Warn("Messaging", "HIERARCHY-ROW-VALUE", "The hierarchy row of folder " + Hex(child) + " has 0x" + ids[k].ToString("X4") + " = " + rv + " but the folder has " + cv[k] + ".", -1, 0, hn, path);
                            }
                        }
                        CheckHmpRow(ht, r, child, hn, path);
                        string nm; if (names.TryGetValue(child, out nm))
                        {
                            PstProp rn = Guard(() => ht.GetCell(r, 0x3001), hn, "hierarchy row");
                            if (rn != null && PstConvert.StringOf(rn) != nm) Warn("Messaging", "HIERARCHY-STALE", "The hierarchy row of folder " + Hex(child) + " has the name '" + PstConvert.StringOf(rn) + "', the folder is named '" + nm + "'.", -1, 0, hn, path);
                        }
                    }
                }
            }
            if (havePc && subFlag != (children > 0)) Warn("Messaging", "FOLDER-SUBFOLDERS", "PidTagSubfolders is " + subFlag + " but the folder has " + children + " subfolder(s).", -1, 0, fnid, path);
            // contents and FAI
            int faiRows, faiUnread;
            ContentsTable(search ? PstFormat.Nid(PstFormat.NtSearchContentsTable, idx) : PstFormat.Nid(PstFormat.NtContentsTable, idx), fnid, path, false, search, listed, msgs, havePc ? contentCount : -1, havePc ? unreadCount : -1, out faiRows, out faiUnread);
            ContentsTable(PstFormat.Nid(PstFormat.NtAssocContentsTable, idx), fnid, path, true, search, listed, msgs, -1, -1, out faiRows, out faiUnread);
            // the hidden (associated) counts of the folder, when it carries them, agree with the FAI table
            int[] fv;
            if (faiRows >= 0 && _folderVals.TryGetValue(fnid, out fv))
            {
                if (fv[3] >= 0 && fv[3] != faiRows) Warn("Messaging", "FOLDER-HIDDEN-COUNT", "PidTagPstHiddenCount is " + fv[3] + " but the FAI table has " + faiRows + " row(s).", -1, 0, fnid, path);
                if (fv[4] >= 0 && fv[4] != faiUnread) Warn("Messaging", "FOLDER-HIDDEN-COUNT", "PidTagPstHiddenUnread is " + fv[4] + " but " + faiUnread + " hidden message(s) are unread.", -1, 0, fnid, path);
                if (fv[5] >= 0 && fv[5] != faiRows) Warn("Messaging", "FOLDER-HIDDEN-COUNT", "PidTagAssocContentCount is " + fv[5] + " but the FAI table has " + faiRows + " row(s).", -1, 0, fnid, path);
            }
        }

        void ContentsTable(uint tn, uint folder, string path, bool assoc, bool search, bool[] listed, List<KeyValuePair<int, string>> msgs, int contentCount, int unreadCount, out int rowCount, out int unreadRows)
        {
            rowCount = -1; unreadRows = 0;
            if (FindNode(tn) < 0)
            {
                if (assoc || !search) Err("Messaging", "FOLDER-TABLE", (assoc ? "FAI contents" : "Contents") + " table " + Hex(tn) + " of the folder is missing.", -1, 0, tn, path);
                return;
            }
            PstTcReader tc = Guard(() => new PstTcReader(_pf.OpenNode(tn)), tn, (assoc ? "FAI contents" : "contents") + " table of '" + path + "'");
            if (tc == null) return;
            CheckTc(tc, tn, (assoc ? "FAI " : "contents ") + path);
            rowCount = tc.RowCount;
            int unread = 0;
            int cellStride = tc.RowCount <= _o.MaxCellRows ? 1 : (tc.RowCount + _o.MaxCellRows - 1) / _o.MaxCellRows;
            PstTcReader.Col flagsCol = tc.FindColumn(0x0E07);
            uint want = assoc ? PstFormat.NtAssocMessage : PstFormat.NtNormalMessage;
            for (int r = 0; r < tc.RowCount; r++)
            {
                uint mid = tc.RowId(r);
                int mi = FindNode(mid);
                if (mi < 0 || PstFormat.NidType(mid) != want)
                {
                    if (!search) Err("Messaging", "CONTENTS-ROW", "Contents row " + r + " refers to " + Hex(mid) + " which is not a " + (assoc ? "FAI " : "") + "message of the file.", -1, 0, tn, path);
                    continue;
                }
                if (search) continue;
                if (listed[mi]) Err("Messaging", "CONTENTS-DUP", "Message " + Hex(mid) + " is listed twice.", -1, 0, tn, path);
                listed[mi] = true;
                if (_nParent[mi] != folder) Err("Messaging", "CONTENTS-PARENT", "Message " + Hex(mid) + " is listed here but its nidParent is " + Hex(_nParent[mi]) + ".", -1, 0, tn, path);
                if (flagsCol != null)
                {
                    PstProp fl = null;
                    try { fl = tc.GetCell(r, flagsCol); } catch (Exception) { }
                    if (fl != null && fl.Data != null && (BitConverter.ToInt32(fl.Data, 0) & 1) == 0) unread++;
                }
                if (!assoc) CheckHmpRow(tc, r, mid, tn, path);
                if (_o.Deep && r % cellStride == 0) CompareRow(tc, r, mi, tn, path);
                msgs.Add(new KeyValuePair<int, string>(mi, path));
            }
            unreadRows = unread;
            if (!assoc && !search)
            {
                if (contentCount >= 0 && contentCount != tc.RowCount) Warn("Messaging", "FOLDER-COUNT", "PidTagContentCount is " + contentCount + " but the contents table has " + tc.RowCount + " row(s).", -1, 0, folder, path);
                if (unreadCount >= 0 && flagsCol != null && unreadCount != unread) Warn("Messaging", "FOLDER-UNREAD", "PidTagContentUnreadCount is " + unreadCount + " but " + unread + " message(s) are unread.", -1, 0, folder, path);
            }
        }

        // ---- LTP checks
        // A big value is not loaded to be checked: it is read through (every block is checked: trailer, signature, CRC, totals of the tree).
        static void ReadThrough(FxBlob value)
        {
            byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(64 * 1024);
            try
            {
                using (System.IO.Stream s = value.OpenRead())
                {
                    long left = value.Length;
                    while (left > 0)
                    {
                        int n = s.Read(buf, 0, (int)Math.Min(buf.Length, left));
                        if (n <= 0) throw new PstFormatException("The data tree ends " + left + " bytes before the value does.");
                        left -= n;
                    }
                }
            }
            finally { System.Buffers.ArrayPool<byte>.Shared.Return(buf); }
        }

        void CheckPc(PstPcReader pc, uint nid, string where)
        {
            long savedLazy = pc.LazyThreshold;
            pc.LazyThreshold = 64 * 1024;
            try { CheckPcEntries(pc, nid, where); }
            finally { pc.LazyThreshold = savedLazy; }
        }

        void CheckPcEntries(PstPcReader pc, uint nid, string where)
        {
            var node = pc.Ltp.Node;
            foreach (PstPcReader.Entry en in pc.Entries)
            {
                try
                {
                    ushort bt = (ushort)(en.Type & 0x0FFF);
                    bool multi = (en.Type & 0x1000) != 0;
                    if (!IsKnownType(bt)) { Warn("LTP", "PC-TYPE", "Property 0x" + en.Id.ToString("X4") + " has the unknown type 0x" + en.Type.ToString("X4") + ".", -1, 0, nid, where); continue; }
                    if (multi && (bt == 0x000D || bt == 0x000B && false)) { Err("LTP", "PC-TYPE", "Property 0x" + en.Id.ToString("X4") + " is multi-valued PtypObject.", -1, 0, nid, where); continue; }
                    if (!multi && bt == 0x000D)
                    {
                        PstProp o = pc.Decode(en);
                        if (o.ObjectView == null) Err("LTP", "PC-OBJECT", "Object property 0x" + en.Id.ToString("X4") + " has no subnode.", -1, 0, nid, where);
                        continue;
                    }
                    PstProp p = pc.Decode(en);
                    if (!multi)
                    {
                        int fs = PstTypes.FixedSize(bt);
                        if (fs > 0 && (p.Data == null || p.Data.Length != fs && !(PstTypes.InlineSize(bt) > 0))) Err("LTP", "PC-SIZE", "Property 0x" + en.Id.ToString("X4") + en.Type.ToString("X4") + " has " + (p.Data == null ? 0 : p.Data.Length) + " bytes (" + fs + " expected).", -1, 0, nid, where);
                        if (bt == 0x001F && p.Data != null && p.Data.Length % 2 != 0) Err("LTP", "PC-STRING", "String property 0x" + en.Id.ToString("X4") + " has an odd length (" + p.Data.Length + ").", -1, 0, nid, where);
                        if (p.Source != null)
                        {
                            if (bt == 0x001F && p.Source.Length % 2 != 0) Err("LTP", "PC-STRING", "String property 0x" + en.Id.ToString("X4") + " has an odd length (" + p.Source.Length + ").", -1, 0, nid, where);
                            ReadThrough(p.Source);
                        }
                    }
                    else if (bt == 0x001F && p.Items != null) foreach (byte[] it in p.Items) if (it.Length % 2 != 0) { Err("LTP", "PC-STRING", "Multi-valued string property 0x" + en.Id.ToString("X4") + " has a value of odd length.", -1, 0, nid, where); break; }
                    if (en.Id >= 0x8000 && _namedCount >= 0 && en.Id >= 0x8000 + _namedCount) Err("Messaging", "NAMEDPROP-UNMAPPED", "Property id 0x" + en.Id.ToString("X4") + " is a named property id but the name-to-id map has only " + _namedCount + " entries.", -1, 0, nid, where);
                }
                catch (PstFormatException e) { Err("LTP", "PC-VALUE", "Property 0x" + en.Id.ToString("X4") + en.Type.ToString("X4") + " cannot be read: " + e.Message, e.Offset, e.Bid, nid, where); }
                catch (Exception e) { Err("LTP", "PC-VALUE", "Property 0x" + en.Id.ToString("X4") + " cannot be read (" + e.GetType().Name + "): " + e.Message, -1, 0, nid, where); }
            }
        }

        static bool IsKnownType(ushort bt)
        {
            switch (bt)
            {
                case 0x0001: case 0x0002: case 0x0003: case 0x0004: case 0x0005: case 0x0006: case 0x0007: case 0x000A: case 0x000B: case 0x000D:
                case 0x0014: case 0x001E: case 0x001F: case 0x0040: case 0x0048: case 0x00FB: case 0x0102: return true;
            }
            return false;
        }

        void CheckTc(PstTcReader tc, uint nid, string where)
        {
            try
            {
                var seenBits = new HashSet<int>();
                uint prev = 0; bool have = false;
                foreach (PstTcReader.Col c in tc.Columns)
                {
                    if (have && c.Tag <= prev) Err("LTP", "TC-COLUMN-ORDER", "Table columns are not sorted by tag (0x" + c.Tag.ToString("X8") + ").", -1, 0, nid, where);
                    prev = c.Tag; have = true;
                    if (!seenBits.Add(c.IBit)) Err("LTP", "TC-IBIT", "Two columns share the existence bit " + c.IBit + ".", -1, 0, nid, where);
                    ushort type = c.Type;
                    int cell = (type & 0x1000) != 0 ? 4 : PstTypes.CellSize(type);
                    if (c.CbData != cell) Err("LTP", "TC-CBDATA", "Column 0x" + c.Tag.ToString("X8") + " has cbData " + c.CbData + " (" + cell + " expected for its type).", -1, 0, nid, where);
                    if (c.Id == 0x67F2 && (c.IbData != 0 || c.IBit != 0)) Warn("LTP", "TC-ROWID", "PidTagLtpRowId must be at offset 0 with existence bit 0.", -1, 0, nid, where);
                    if (c.Id == 0x67F3 && (c.IbData != 4 || c.IBit != 1)) Warn("LTP", "TC-ROWVER", "PidTagLtpRowVer must be at offset 4 with existence bit 1.", -1, 0, nid, where);
                }
                if (tc.FindColumn(0x67F2) == null) Err("LTP", "TC-ROWID", "The table has no PidTagLtpRowId column.", -1, 0, nid, where);
                // row index: unique, increasing keys
                // (the reader already verified that the row numbers are a permutation of 0..RowCount-1)
                var ids = new HashSet<uint>();
                for (int r = 0; r < tc.RowCount; r++) if (!ids.Add(tc.RowId(r))) { Err("LTP", "TC-ROWID-DUP", "Row id 0x" + tc.RowId(r).ToString("X") + " appears twice in the row index.", -1, 0, nid, where); break; }
                if (tc.RowCount > 0)
                {
                    PstTcReader.Col idc = tc.FindColumn(0x67F2);
                    int stride = tc.RowCount <= _o.MaxCellRows ? 1 : (tc.RowCount + _o.MaxCellRows - 1) / _o.MaxCellRows;
                    for (int r = 0; r < tc.RowCount; r += stride)
                    {
                        if (idc != null)
                        {
                            PstProp rid = tc.GetCell(r, idc);
                            if (rid == null || BitConverter.ToUInt32(rid.Data, 0) != tc.RowId(r)) { Err("LTP", "TC-ROWID-MISMATCH", "Row " + r + ": PidTagLtpRowId in the matrix differs from the row index (0x" + tc.RowId(r).ToString("X") + ").", -1, 0, nid, where); }
                        }
                        foreach (PstTcReader.Col c in tc.Columns)
                        {
                            if (c.Id == 0x67F2 || c.Id == 0x67F3) continue;
                            PstProp p = tc.GetCell(r, c);
                            if (p != null && p.Data != null && c.Type == 0x001F && p.Data.Length % 2 != 0) { Err("LTP", "PC-STRING", "Row " + r + ": string cell 0x" + c.Tag.ToString("X8") + " has an odd length.", -1, 0, nid, where); }
                        }
                    }
                    if (stride > 1) Info("LTP", "TC-SAMPLED", "Table of " + tc.RowCount + " rows: cells checked one row in " + stride + ".", -1, 0, nid, where);
                }
            }
            catch (PstFormatException e) { Err("LTP", "TC-BROKEN", "Table is inconsistent: " + e.Message, e.Offset, e.Bid, nid, where); }
            catch (Exception e) { Err("LTP", "TC-BROKEN", "Table cannot be checked (" + e.GetType().Name + "): " + e.Message, -1, 0, nid, where); }
        }

        // ---- messages
        void CheckMessageNode(int ni, string folderPath)
        {
            uint nid = _nNid[ni];
            string where = folderPath;
            try
            {
                PstNodeView node = _pf.OpenNode(nid);
                CheckMessage(node, nid, where, 0);
            }
            catch (PstFormatException e) { Err("Messaging", "MSG-BROKEN", "Message cannot be read: " + e.Message, e.Offset, e.Bid, nid, where); }
            catch (Exception e) { Err("Messaging", "MSG-BROKEN", "Message cannot be read (" + e.GetType().Name + "): " + e.Message, -1, 0, nid, where); }
        }

        void CheckMessage(PstNodeView node, uint nid, string where, int depth)
        {
            if (depth > 32) { Err("Messaging", "MSG-DEPTH", "Embedded messages nested deeper than 32 levels.", -1, 0, nid, where); return; }
            PstPcReader pc = new PstPcReader(node);
            PstProp subj = null;
            try { subj = pc.Get(0x0037); } catch (Exception) { }
            string w = where + " | " + (subj != null && subj.Data != null ? Trunc(PstConvert.StringOf(subj)) : "(no subject)") + (depth > 0 ? " (embedded)" : "");
            CheckPc(pc, nid, w);
            foreach (ushort need in new ushort[] { 0x001A, 0x0E07, 0x0E08, 0x0E17, 0x3007, 0x3008, 0x300B })
                if (pc.Find(need) == null) Warn("Messaging", "MSG-PROP", "Message lacks the required property 0x" + need.ToString("X4") + ".", -1, 0, nid, w);
            PstProp flags = pc.Get(0x0E07);
            bool hasAttach = flags != null && flags.Data != null && (BitConverter.ToInt32(flags.Data, 0) & 0x10) != 0;
            PstProp size = pc.Get(0x0E08);
            long total = 0;
            // recipients
            if (!node.HasSubnode(PstFormat.NidRecipientTable)) Err("Messaging", "MSG-RECIPIENTS", "The recipient table (subnode 0x692) is missing.", -1, 0, nid, w);
            else
            {
                PstTcReader rt = new PstTcReader(node.OpenSubnode(PstFormat.NidRecipientTable));
                CheckTc(rt, nid, w + " recipients");
                Report.Stats["recipients"] = (Report.Stats.ContainsKey("recipients") ? Report.Stats["recipients"] : 0) + rt.RowCount;
                foreach (PstTcReader.Col c in rt.Columns) if (c.Id >= 0x8000 && _namedCount >= 0 && c.Id >= 0x8000 + _namedCount) Err("Messaging", "NAMEDPROP-UNMAPPED", "Recipient column 0x" + c.Id.ToString("X4") + " is not in the name-to-id map.", -1, 0, nid, w);
            }
            // attachments
            var attNids = new List<uint>();
            foreach (uint sn in node.SubnodeNids) { TrackSubnode(sn); if (PstFormat.NidType(sn) == PstFormat.NtAttachment) attNids.Add(sn); }
            if (node.HasSubnode(PstFormat.NidAttachmentTable))
            {
                if (!hasAttach) Warn("Messaging", "MSG-HASATTACH", "There is an attachment table but mfHasAttach is not set in PidTagMessageFlags.", -1, 0, nid, w);
                PstTcReader at = new PstTcReader(node.OpenSubnode(PstFormat.NidAttachmentTable));
                CheckTc(at, nid, w + " attachments");
                var rows = new HashSet<uint>();
                for (int r = 0; r < at.RowCount; r++) rows.Add(at.RowId(r));
                if (rows.Count != attNids.Count || !attNids.All(rows.Contains)) Err("Messaging", "MSG-ATTACH-MISMATCH", "The attachment table lists " + rows.Count + " attachment(s) but the message holds " + attNids.Count + " attachment subnode(s).", -1, 0, nid, w);
                Report.Stats["attachments"] = (Report.Stats.ContainsKey("attachments") ? Report.Stats["attachments"] : 0) + at.RowCount;
                for (int r = 0; r < at.RowCount; r++)
                {
                    uint an = at.RowId(r);
                    if (!node.HasSubnode(an)) { Err("Messaging", "ATTACH-MISSING", "Attachment " + Hex(an) + " listed in the attachment table has no subnode.", -1, 0, nid, w); continue; }
                    try { total += CheckAttachment(node.OpenSubnode(an), nid, w, depth); }
                    catch (PstFormatException e) { Err("Messaging", "ATTACH-BROKEN", "Attachment " + Hex(an) + " cannot be read: " + e.Message, e.Offset, e.Bid, nid, w); }
                }
            }
            else
            {
                if (hasAttach) Err("Messaging", "MSG-ATTACH-TABLE", "mfHasAttach is set but the attachment table (subnode 0x671) is missing.", -1, 0, nid, w);
                if (attNids.Count > 0) Err("Messaging", "MSG-ATTACH-MISMATCH", "The message holds " + attNids.Count + " attachment subnode(s) but no attachment table.", -1, 0, nid, w);
            }
            if (size != null && size.Data != null)
            {
                int ms = BitConverter.ToInt32(size.Data, 0);
                long computed = NodeBytes(node.BidData, node.BidSub, 0);
                if (ms != computed) Warn("Messaging", "MSG-SIZE-COMPUTED", "PidTagMessageSize is " + ms + " but the node holds " + computed + " bytes (data and subnodes): Outlook recomputes it.", -1, 0, nid, w);
                if (ms <= 0 || ms < total) Warn("Messaging", "MSG-SIZE", "PidTagMessageSize " + ms + " is not plausible (attachments alone hold " + total + " bytes).", -1, 0, nid, w);
            }
        }

        // size of the data of a block tree without reading the data blocks
        long DataBytes(ulong bid)
        {
            if (bid == 0) return 0;
            int i = FindBlock(bid);
            if (i < 0) return 0;
            if ((bid & 2) == 0) return _bCb[i];
            byte[] b = ReadInternal(i);
            return b != null && b.Length >= 8 ? PstFormat.G32(b, 4) : 0;
        }

        // data of a node and of all its subnodes (what PidTagMessageSize adds up)
        long NodeBytes(ulong bidData, ulong bidSub, int depth)
        {
            long n = DataBytes(bidData);
            if (bidSub == 0 || depth > 8) return n;
            foreach (PstSubEntry se in _pf.ReadSubnodes(bidSub)) n += NodeBytes(se.BidData, se.BidSub, depth + 1);
            return n;
        }

        // size of one property value as PidTagAttachSize adds it up
        long ValueBytes(PstPcReader pc, PstPcReader.Entry en, PstNodeView node)
        {
            ushort bt = (ushort)(en.Type & 0x0FFF);
            bool multi = (en.Type & 0x1000) != 0;
            if (!multi && PstTypes.InlineSize(en.Type) > 0) return bt == 0x000B ? 1 : PstTypes.InlineSize(en.Type);
            if (!multi && bt == 0x000D) { var o = pc.Ltp.Heap.Get(en.Hnid); return o.Length >= 8 ? PstFormat.G32(o, 4) : 0; }
            if (en.Hnid == 0) return 0;
            if ((en.Hnid & 0x1F) == 0) return pc.Ltp.Heap.Get(en.Hnid).Length;
            PstSubEntry? found = null;
            foreach (PstSubEntry se in _pf.ReadSubnodes(node.BidSub)) if (se.Nid == en.Hnid) found = se;
            return found.HasValue ? NodeBytes(found.Value.BidData, found.Value.BidSub, 1) : 0;
        }

        long CheckAttachment(PstNodeView an, uint msgNid, string where, int depth)
        {
            PstPcReader pc = new PstPcReader(an);
            CheckPc(pc, msgNid, where + " attachment");
            foreach (ushort need in new ushort[] { 0x0E20, 0x3705, 0x370B })
                if (pc.Find(need) == null) Warn("Messaging", "ATTACH-PROP", "Attachment lacks the required property 0x" + need.ToString("X4") + ".", -1, 0, msgNid, where);
            PstPcReader.Entry asz = pc.Find(0x0E20);
            if (asz != null)
            {
                long sum = 0;
                foreach (PstPcReader.Entry en in pc.Entries) sum += ValueBytes(pc, en, an);
                PstProp ap = pc.Decode(asz);
                if (ap.Data != null && BitConverter.ToInt32(ap.Data, 0) != sum) Warn("Messaging", "ATTACH-SIZE", "PidTagAttachSize is " + BitConverter.ToInt32(ap.Data, 0) + " but the values of the properties add up to " + sum + " bytes: Outlook considers it invalid.", -1, 0, msgNid, where);
            }
            PstProp method = pc.Get(0x3705);
            int m = method == null ? 1 : PstConvert.Int32Of(method, 1);
            PstPcReader.Entry data = pc.Find(0x3701);
            long bytes = 0;
            if (m == 5)
            {
                if (data == null || (data.Type & 0x0FFF) != 0x000D) Err("Messaging", "ATTACH-EMBEDDED", "Attachment method 5 (embedded message) without a PtypObject PidTagAttachDataObject.", -1, 0, msgNid, where);
                else
                {
                    PstProp o = pc.Decode(data);
                    uint en = o.ObjectView.Nid, et = PstFormat.NidType(en);
                    if (et != PstFormat.NtNormalMessage && et != PstFormat.NtAssocMessage) Warn("Messaging", "ATTACH-EMBEDDED-NID", "The embedded message of the attachment is the subnode " + Hex(en) + " whose NID type is 0x" + et.ToString("X") + ", not a message type.", -1, 0, msgNid, where);
                    TrackSubnode(en);
                    CheckMessage(o.ObjectView, msgNid, where, depth + 1);
                    bytes = o.ObjectSize;
                }
            }
            else if (m == 1)
            {
                if (data == null) Warn("Messaging", "ATTACH-DATA", "Attachment by value without PidTagAttachDataBinary.", -1, 0, msgNid, where);
                else if ((data.Type & 0x0FFF) != 0x0102) Err("Messaging", "ATTACH-DATA", "PidTagAttachDataBinary has type 0x" + data.Type.ToString("X4") + ".", -1, 0, msgNid, where);
                else if (data.Hnid == 0) bytes = 0;            // an empty value has no storage (HNID 0)
                else if ((data.Hnid & 0x1F) != 0) { bytes = an.OpenSubnode(data.Hnid).DataLength; }
                else bytes = pc.Ltp.Heap.Get(data.Hnid).Length;
            }
            return bytes;
        }

        static string Trunc(string s) { return s == null ? "" : s.Length > 60 ? s.Substring(0, 60) + "..." : s; }
    }
}
