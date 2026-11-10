// =============================================================================
//  Mailbox Transfer - engine, part 5i: scan of a damaged PST for valid blocks and pages (salvage)
// -----------------------------------------------------------------------------
//  PstScan.Scan(file) reads the whole file once, 4 MiB at a time, and finds
//    * every block: a block ends on a 64-byte boundary with a 16-byte trailer (cb, signature, CRC of the data, BID) and starts
//      cb + 16 rounded up to 64 bytes earlier; it is valid when the signature (offset xor BID) and the CRC of its data agree, which a
//      random sequence of bytes does with a probability of about 2^-48. The block B-tree is not read at all.
//    * every node B-tree leaf page (512 bytes, ptype 0x81, level 0, valid CRC and signature): the entries give NID -> data BID,
//      subnode BID, parent. Index pages above the leaves are not needed.
//  The result replaces the two B-trees (PstFile.Index): block lookup by binary search on the sorted blocks, node lookup on the sorted
//  nodes. When the same NID appears in several pages (the old versions that a file keeps in free space) the entry whose blocks exist
//  wins, then the one of the newest page (page BIDs increase with time). Cost: one pass over the file; memory 24 bytes per block and
//  about 40 bytes per node entry found, a few hundred MB for tens of GB.
//
//  Also kept: the roots of subnode trees and what refers to them, the property context heads that nothing refers to (message property contexts of lost\r\n//  entries, the name-to-id map: see PstRepairRun.ResolveNameMap) and, apart, the blocks whose signature is right and CRC wrong (a block with one wrong byte can be\r\n//  corrected by its CRC, PstCrcRepair; they are only looked up when UseDamagedBlocks is set).\r\n//\r\n//  Limits: a node whose leaf page is lost is lost (a block has no NID); an entry of a node deleted later can come back from an old
//  page; the allocation maps are not used to tell old pages from live ones.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public sealed class PstScan : IPstIndex
    {
        struct Blk { public ulong Bid; public long Ib; public int Cb; }
        struct Cand { public uint Nid, Parent; public ulong BidData, BidSub, PageBid; }

        struct XRoot { public ulong X, First; }
        // what the blocks say about each other, to find the roots of nodes whose entry is lost (see UnreferencedRoots)
        List<ulong> _heads = new List<ulong>(), _xchild = new List<ulong>(), _subRef = new List<ulong>();
        List<XRoot> _xroots = new List<XRoot>();
        // subnode tree blocks (SLBLOCK level 0, SIBLOCK) and what refers to them (SIBLOCK children, BidSub of subnode entries)
        List<ulong> _subRoots = new List<ulong>(), _subChild = new List<ulong>();
        ulong[] _headsS, _xchildS, _subRefS, _subRootsS, _subChildS;
        // blocks whose signature is right and CRC wrong (damaged data), and the ones among them that start like a property context
        List<Blk> _damaged = new List<Blk>();
        Blk[] _damagedS = new Blk[0];
        List<ulong> _headsDamaged = new List<ulong>();
        /// <summary>Block lookups also find the damaged blocks (they fail the CRC check unless PstFile.CorrectBlocks is set and one byte is wrong).</summary>
        public bool UseDamagedBlocks;
        public long DamagedBlocks { get { return _damagedS.Length; } }
        Blk[] _blocks = new Blk[0];
        PstNbtEntry[] _nodes = new PstNbtEntry[0];
        public long BytesScanned, BlocksFound, DuplicateBlocks, NbtLeafPages, BbtLeafPages, OtherPages, NodeEntries, NodesKept, StaleEntries, DamagedNodes;

        const int Chunk = 4 * 1024 * 1024;
        const int Overlap = 8192;

        public static PstScan Scan(PstFile f, System.Threading.CancellationToken ct, IProgress<PstProgress> progress)
        {
            var s = new PstScan();
            long len = f.Length;
            var chunks = new List<Blk[]>();
            Blk[] cur = new Blk[1 << 20];
            int n = 0;
            long total = 0;
            var leafPages = new List<long>();
            byte[] buf = new byte[Chunk];
            long pos = 0, doneEnd = 0, pageNext = 0;
            while (pos < len)
            {
                ct.ThrowIfCancellationRequested();
                int got = (int)Math.Min(Chunk, len - pos);
                f.ReadAtInto(pos, buf, got);
                long end = pos + got;
                // blocks: trailer at e - 16, e is a multiple of 64
                long e = Math.Max(pos + 64, (doneEnd / 64 + 1) * 64);
                for (; e <= end; e += 64)
                {
                    int t = (int)(e - 16 - pos);
                    int cb = PstFormat.G16(buf, t);
                    if (cb == 0 || cb > PstFormat.MaxBlockData) continue;
                    int blockTotal = PstFormat.BlockSize(cb);
                    long ib = e - blockTotal;
                    if (ib < pos) continue;
                    ulong bid = PstFormat.G64(buf, t + 8);
                    if (bid == 0 || PstFormat.G16(buf, t + 2) != PstFormat.Sig((ulong)ib, bid)) continue;
                    if (PstFormat.Crc(buf, (int)(ib - pos), cb) != PstFormat.G32(buf, t + 4))
                    {
                        // signature right, data wrong: a block hit by damage (kept apart, only looked up when asked for, see UseDamagedBlocks)
                        s.AddDamaged(f.CryptMethod, buf, (int)(ib - pos), cb, bid, ib);
                        continue;
                    }
                    s.Classify(f.CryptMethod, buf, (int)(ib - pos), cb, bid);
                    if (n == cur.Length) { chunks.Add(cur); cur = new Blk[1 << 20]; n = 0; }
                    cur[n++] = new Blk { Bid = bid & ~1UL, Ib = ib, Cb = cb };
                    total++;
                }
                doneEnd = end;
                // pages: 512-byte aligned
                long q = Math.Max(pageNext, (pos + 511) / 512 * 512);
                for (; q + 512 <= end; q += 512)
                {
                    int o = (int)(q - pos);
                    byte pt = buf[o + 496];
                    if (pt < 0x80 || pt > 0x86 || buf[o + 497] != pt) continue;
                    ulong pbid = PstFormat.G64(buf, o + 504);
                    if (PstFormat.G16(buf, o + 498) != PstFormat.Sig((ulong)q, pbid)) continue;
                    if (PstFormat.Crc(buf, o, 496) != PstFormat.G32(buf, o + 500)) continue;
                    if (pt == PstFormat.PtypeNBT && buf[o + 491] == 0 && buf[o + 490] >= 32 && buf[o + 488] <= buf[o + 489] && buf[o + 488] * buf[o + 490] <= 488) { leafPages.Add(q); s.NbtLeafPages++; }
                    else if (pt == PstFormat.PtypeBBT && buf[o + 491] == 0) s.BbtLeafPages++;
                    else s.OtherPages++;
                }
                pageNext = q;
                s.BytesScanned = end;
                if (progress != null) progress.Report(new PstProgress { Phase = "scan", BytesDone = end, BytesTotal = len, Items = total });
                if (end >= len) break;
                pos = end - Overlap;
                pos = pos / 512 * 512;
            }
            chunks.Add(cur);
            // the blocks, sorted by BID (a BID found twice: the first one)
            var all = new Blk[total];
            long k = 0;
            for (int c = 0; c < chunks.Count; c++)
            {
                int cnt = c == chunks.Count - 1 ? n : chunks[c].Length;
                Array.Copy(chunks[c], 0, all, k, cnt);
                k += cnt;
            }
            chunks = null; cur = null;
            Array.Sort(all, (a, b) => a.Bid != b.Bid ? a.Bid.CompareTo(b.Bid) : a.Ib.CompareTo(b.Ib));
            int u = 0;
            for (int i = 0; i < all.Length; i++)
            {
                if (u > 0 && all[u - 1].Bid == all[i].Bid) { s.DuplicateBlocks++; continue; }
                all[u++] = all[i];
            }
            if (u != all.Length) Array.Resize(ref all, u);
            s._blocks = all;
            s.BlocksFound = u;
            s._headsS = s._heads.ToArray(); Array.Sort(s._headsS);
            s._xchildS = s._xchild.ToArray(); Array.Sort(s._xchildS);
            s._subRefS = s._subRef.ToArray(); Array.Sort(s._subRefS);
            s._subRootsS = s._subRoots.ToArray(); Array.Sort(s._subRootsS);
            s._subChildS = s._subChild.ToArray(); Array.Sort(s._subChildS);
            s._heads = null; s._xchild = null; s._subRef = null; s._subRoots = null; s._subChild = null;
            s._damagedS = s._damaged.OrderBy(b => b.Bid).ThenBy(b => b.Ib).ToArray();
            s._damaged = null;

            // the node entries of the leaf pages found
            long entries = 0;
            var pg = new byte[512];
            foreach (long ib in leafPages) { f.ReadAtInto(ib, pg, 512); entries += pg[488]; }
            var cands = new Cand[entries];
            long ci = 0;
            foreach (long ib in leafPages)
            {
                ct.ThrowIfCancellationRequested();
                f.ReadAtInto(ib, pg, 512);
                ulong pbid = PstFormat.G64(pg, 504);
                int cEnt = pg[488], cbEnt = pg[490];
                for (int i = 0; i < cEnt; i++)
                {
                    int o = i * cbEnt;
                    ulong nid = PstFormat.G64(pg, o);
                    if (nid == 0 || nid > uint.MaxValue) continue;
                    cands[ci++] = new Cand { Nid = (uint)nid, BidData = PstFormat.G64(pg, o + 8), BidSub = PstFormat.G64(pg, o + 16), Parent = PstFormat.G32(pg, o + 24), PageBid = pbid };
                }
            }
            s.NodeEntries = ci;
            if (ci != cands.Length) Array.Resize(ref cands, (int)ci);
            // sort by NID, then best entry first: blocks present, then newest page
            var rank = new Func<Cand, int>(c => (c.BidData == 0 || s.HasBlock(c.BidData) ? 2 : 0) + (c.BidSub == 0 || s.HasBlock(c.BidSub) ? 1 : 0));
            var ranks = new int[cands.Length];
            for (int i = 0; i < cands.Length; i++) ranks[i] = rank(cands[i]);
            var order = new int[cands.Length];
            for (int i = 0; i < order.Length; i++) order[i] = i;
            Array.Sort(order, (a, b) =>
            {
                int c = cands[a].Nid.CompareTo(cands[b].Nid);
                if (c != 0) return c;
                c = ranks[b].CompareTo(ranks[a]);
                if (c != 0) return c;
                return cands[b].PageBid.CompareTo(cands[a].PageBid);
            });
            var nodes = new List<PstNbtEntry>();
            for (int i = 0; i < order.Length; i++)
            {
                Cand c = cands[order[i]];
                if (i > 0 && cands[order[i - 1]].Nid == c.Nid) { s.StaleEntries++; continue; }
                nodes.Add(new PstNbtEntry { Nid = c.Nid, BidData = c.BidData, BidSub = c.BidSub, Parent = c.Parent });
                if (ranks[order[i]] != 3) s.DamagedNodes++;
            }
            s._nodes = nodes.ToArray();
            s.NodesKept = s._nodes.Length;
            return s;
        }

        /// <summary>What a valid block is: the start of a heap-on-node holding a property context, a tree block (children), a subnode block (data BIDs of the subnodes).</summary>
        void Classify(int crypt, byte[] buf, int rel, int cb, ulong bid)
        {
            if ((bid & 2) == 0)
            {
                if (cb < 12) return;
                var tmp = new byte[] { buf[rel], buf[rel + 1], buf[rel + 2], buf[rel + 3] };
                if (crypt != 0) PstCrypt.Decode(crypt, tmp, 0, 4, bid, false);
                if (tmp[2] == 0xEC && tmp[3] == 0xBC) _heads.Add(bid & ~1UL);
            }
            else if (cb >= 8)
            {
                byte b0 = buf[rel];
                int cEnt = PstFormat.G16(buf, rel + 2);
                if (b0 == 1)
                {
                    for (int i = 0; i < cEnt && 8 + 8 * i + 8 <= cb; i++)
                    {
                        ulong ch = PstFormat.G64(buf, rel + 8 + 8 * i) & ~1UL;
                        _xchild.Add(ch);
                        if (i == 0 && buf[rel + 1] == 1) _xroots.Add(new XRoot { X = bid & ~1UL, First = ch });
                    }
                }
                else if (b0 == 2 && buf[rel + 1] == 0)
                {
                    _subRoots.Add(bid & ~1UL);
                    for (int i = 0; i < cEnt && 8 + 24 * i + 24 <= cb; i++)
                    {
                        _subRef.Add(PstFormat.G64(buf, rel + 16 + 24 * i) & ~1UL);
                        ulong inner = PstFormat.G64(buf, rel + 24 + 24 * i) & ~1UL;
                        if (inner != 0) _subChild.Add(inner);
                    }
                }
                else if (b0 == 2 && buf[rel + 1] == 1)
                {
                    _subRoots.Add(bid & ~1UL);
                    for (int i = 0; i < cEnt && 8 + 16 * i + 16 <= cb; i++) _subChild.Add(PstFormat.G64(buf, rel + 16 + 16 * i) & ~1UL);
                }
            }
        }

        static bool Has(ulong[] sorted, ulong v) { return Array.BinarySearch(sorted, v) >= 0; }

        /// <summary>
        /// Roots of data that look like a property context (a heap with client signature 0xBC, alone or as the first block of a tree)
        /// and that no node entry, no subnode entry and no tree refers to: what is left when the entry of a node is lost (a node B-tree
        /// destroyed or cut off). Old versions of a deleted or modified message in free space look the same: use it when the node B-tree
        /// is gone, not otherwise.
        /// </summary>
        public List<ulong> UnreferencedRoots()
        {
            var nodeData = new HashSet<ulong>();
            foreach (PstNbtEntry n in _nodes) nodeData.Add(n.BidData & ~1UL);
            var res = new List<ulong>();
            foreach (ulong h in _headsS)
                if (!Has(_xchildS, h) && !Has(_subRefS, h) && !nodeData.Contains(h)) res.Add(h);
            foreach (XRoot x in _xroots)
                if (Has(_headsS, x.First) && !Has(_xchildS, x.X) && !Has(_subRefS, x.X) && !nodeData.Contains(x.X)) res.Add(x.X);
            return res;
        }

        /// <summary>
        /// Roots of subnode trees (SLBLOCK / SIBLOCK) that no node entry, no subnode entry and no SIBLOCK refers to: what is left of the
        /// subnode tree of a node whose entry is lost (the long values of a property context are in its subnodes).
        /// </summary>
        public List<ulong> UnreferencedSubRoots()
        {
            var nodeSub = new HashSet<ulong>();
            foreach (PstNbtEntry n in _nodes) if (n.BidSub != 0) nodeSub.Add(n.BidSub & ~1UL);
            var res = new List<ulong>();
            foreach (ulong b in _subRootsS)
                if (!Has(_subChildS, b) && !nodeSub.Contains(b)) res.Add(b);
            return res;
        }

        void AddDamaged(int crypt, byte[] buf, int rel, int cb, ulong bid, long ib)
        {
            _damaged.Add(new Blk { Bid = bid & ~1UL, Ib = ib, Cb = cb });
            if ((bid & 2) == 0 && cb >= 12)
            {
                var tmp = new byte[] { buf[rel], buf[rel + 1], buf[rel + 2], buf[rel + 3] };
                if (crypt != 0) PstCrypt.Decode(crypt, tmp, 0, 4, bid, false);
                if (tmp[2] == 0xEC && tmp[3] == 0xBC) _headsDamaged.Add(bid & ~1UL);
            }
        }

        /// <summary>Damaged blocks that start like a property context and that no node entry, subnode entry or tree refers to.</summary>
        public List<ulong> UnreferencedDamagedHeads()
        {
            var nodeData = new HashSet<ulong>();
            foreach (PstNbtEntry n in _nodes) nodeData.Add(n.BidData & ~1UL);
            var res = new List<ulong>();
            foreach (ulong h in _headsDamaged)
                if (!Has(_xchildS, h) && !Has(_subRefS, h) && !nodeData.Contains(h)) res.Add(h);
            return res;
        }

        public const uint FirstUnlinkedIndex = 0x7000000;

        /// <summary>Gives each root a made-up NID of the message type (parent 0) so that the rest of the repair reads it like a node.</summary>
        public List<uint> AddUnlinked(IList<ulong> dataBids)
        {
            uint next = FirstUnlinkedIndex;
            if (_nodes.Length > 0 && PstFormat.NidIndex(_nodes[_nodes.Length - 1].Nid) >= next) next = PstFormat.NidIndex(_nodes[_nodes.Length - 1].Nid) + 1;
            var res = new List<uint>();
            var all = new List<PstNbtEntry>(_nodes);
            foreach (ulong b in dataBids)
            {
                uint nid = PstFormat.Nid(PstFormat.NtNormalMessage, next++);
                all.Add(new PstNbtEntry { Nid = nid, BidData = b, BidSub = 0, Parent = 0 });
                res.Add(nid);
            }
            _nodes = all.ToArray();
            return res;
        }

        bool HasBlock(ulong bid)
        {
            PstBbtEntry e;
            return TryFindBlock(bid, out e);
        }

        /// <summary>
        /// The encryption method of the blocks: the one under which the first data blocks of the nodes found start like a heap (0xEC at
        /// offset 2). The header's value is kept when it fits as well as another.
        /// </summary>
        public int DetectCrypt(PstFile f, int fromHeader)
        {
            var hits = new int[3];
            int sampled = 0;
            foreach (PstNbtEntry n in _nodes)
            {
                if (n.BidData == 0 || (n.BidData & 2) != 0) continue;
                PstBbtEntry e;
                if (!TryFindBlock(n.BidData, out e) || e.Cb < 12) continue;
                byte[] raw;
                try { raw = f.ReadAt((long)e.Ib, e.Cb); } catch (PstFormatException) { continue; }
                for (int m = 0; m < 3; m++)
                {
                    var c = (byte[])raw.Clone();
                    if (m != 0) PstCrypt.Decode(m, c, 0, c.Length, e.Bid, false);
                    if (c[2] == 0xEC) hits[m]++;
                }
                if (++sampled >= 200) break;
            }
            int best = fromHeader >= 0 && fromHeader <= 2 ? fromHeader : 0;
            for (int m = 0; m < 3; m++) if (hits[m] > hits[best]) best = m;
            return best;
        }

        public string Describe()
        {
            return "scan of " + BytesScanned.ToString("N0") + " bytes: " + BlocksFound.ToString("N0") + " valid blocks, " + NbtLeafPages.ToString("N0") + " node B-tree leaf pages (" + NodeEntries.ToString("N0") + " entries, "
                + NodesKept.ToString("N0") + " nodes kept, " + StaleEntries.ToString("N0") + " older entries ignored, " + DamagedNodes.ToString("N0") + " nodes with a missing block)";
        }

        public bool TryFindBlock(ulong bid, out PstBbtEntry e)
        {
            bid &= ~1UL;
            int lo = 0, hi = _blocks.Length - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                ulong b = _blocks[mid].Bid;
                if (b == bid) { e = new PstBbtEntry { Bid = bid, Ib = (ulong)_blocks[mid].Ib, Cb = _blocks[mid].Cb, CRef = 2 }; return true; }
                if (b < bid) lo = mid + 1; else hi = mid - 1;
            }
            if (UseDamagedBlocks)
            {
                lo = 0; hi = _damagedS.Length - 1;
                while (lo <= hi)
                {
                    int mid = (lo + hi) >> 1;
                    ulong b = _damagedS[mid].Bid;
                    if (b == bid) { e = new PstBbtEntry { Bid = bid, Ib = (ulong)_damagedS[mid].Ib, Cb = _damagedS[mid].Cb, CRef = 2 }; return true; }
                    if (b < bid) lo = mid + 1; else hi = mid - 1;
                }
            }
            e = default(PstBbtEntry);
            return false;
        }

        public bool TryFindNode(uint nid, out PstNbtEntry e)
        {
            int lo = 0, hi = _nodes.Length - 1;
            while (lo <= hi)
            {
                int mid = (lo + hi) >> 1;
                uint k = _nodes[mid].Nid;
                if (k == nid) { e = _nodes[mid]; return true; }
                if (k < nid) lo = mid + 1; else hi = mid - 1;
            }
            e = default(PstNbtEntry);
            return false;
        }

        public IEnumerable<PstNbtEntry> Nodes() { return _nodes; }
    }
}
