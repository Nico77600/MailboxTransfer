// =============================================================================
//  Mailbox Transfer - engine, part 5j: the name-to-id map in a damaged PST (reading it as far as it goes, finding it again without its node)
// -----------------------------------------------------------------------------
//  Named properties are stored in a message only by id (0x8000 and up); the GUID and the LID or the name of each id are in one place, the
//  name-to-id map (NID 0x61), a property context: PidTagNameidBucketCount (0x0001), GUID stream (0x0002), entry stream (0x0003),
//  string stream (0x0004) and the hash buckets (0x1000 and up). When the node entry of the map is lost (a leaf page of the node B-tree
//  damaged) every message loses its named properties (calendar times, recurrence, contact fields, categories, ...). The map itself is
//  usually intact in the file: this part finds it from the blocks (PstScan) with no entry pointing at it.
//
//    PstNameMapReader.Read       the map of a property context, entry by entry; what cannot be read is listed
//    PstRepairRun.ResolveNameMap the repair's choice: the original map when it reads without error, else the best map found by the scan
//                                (the one that resolves the most named-property ids used by the messages), else what is left of the
//                                original; the result is always written again by the writer, from what the messages use
// =============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace MailboxTransfer
{
    /// <summary>Where the name-to-id map of a repaired file comes from.</summary>
    public enum PstNameMapSource
    {
        /// <summary>Not determined (plan only, or the repair did not get that far).</summary>
        None = 0,
        /// <summary>Node 0x61 of the source, read without any error.</summary>
        Original = 1,
        /// <summary>Node 0x61 lost or unreadable; a map found in the blocks of the file by the scan, with nothing pointing at it.</summary>
        RecoveredByScan = 2,
        /// <summary>No complete map: the map of the result is built from the ids the messages use that the damaged map still resolves; the others are dropped.</summary>
        RebuiltFromUsage = 3,
    }

    public sealed class PstNameMapInfo
    {
        public int Entries, Good, Buckets, BucketRecords, BucketMismatch, Unconfirmed;
        public bool HasBucketCount, HasGuids, HasStrings;
        public long BucketCount;
    }

    /// <summary>
    /// A block with one wrong byte: the CRC is linear (seed 0, no final xor), so the CRC of the damaged data and the stored one differ by the
    /// CRC of an error vector that is zero everywhere but one byte; the byte and its error are found by trying the 255 error values at every
    /// position (cb x 255 steps). Accepted only when exactly one (position, error) fits and the corrected block has the stored CRC.
    /// </summary>
    public static class PstCrcRepair
    {
        static readonly uint[] T0 = BuildT0();

        static uint[] BuildT0()
        {
            var t = new uint[256];
            for (int i = 0; i < 256; i++) t[i] = PstFormat.Crc(new[] { (byte)i }, 0, 1);
            return t;
        }

        /// <summary>Corrects raw[0..cb) in place when one byte is wrong; false (raw unchanged) when the CRC says nothing or says it ambiguously.</summary>
        public static bool TrySingleByte(byte[] raw, int cb, uint stored)
        {
            uint delta = PstFormat.Crc(raw, 0, cb) ^ stored;
            if (delta == 0 || cb <= 0) return false;
            var v = new uint[256];
            for (int e = 1; e < 256; e++) v[e] = T0[e];
            int pos = -1, err = 0, fits = 0;
            for (int k = 0; k < cb; k++)            // k zero bytes follow the wrong byte
            {
                for (int e = 1; e < 256; e++)
                {
                    if (v[e] == delta) { fits++; pos = cb - 1 - k; err = e; }
                    v[e] = T0[v[e] & 0xFF] ^ (v[e] >> 8);
                }
            }
            if (fits != 1) return false;
            raw[pos] ^= (byte)err;
            if (PstFormat.Crc(raw, 0, cb) == stored) return true;
            raw[pos] ^= (byte)err;
            return false;
        }
    }

    public static class PstNameMapReader
    {
        struct Rec { public uint Hash; public int Ng, Idx; }

        /// <summary>
        /// The entries of a name-to-id map property context, by id. errors == null: the first problem raises a PstFormatException;
        /// otherwise a bad entry is skipped and listed, and each stream is read on its own: with the string stream lost the entries
        /// named by a number are still resolved; with the entry stream lost the entries are rebuilt from the hash buckets (a record is
        /// hash, GUID word and index: a number is its own hash, a name is found in the string stream by its hash). info (optional)
        /// receives the figures that tell a coherent map from a mix of blocks of different versions.
        /// </summary>
        public static Dictionary<ushort, PstNameMap.Entry> Read(PstPcReader pc, List<string> errors, PstNameMapInfo info) { return Read(pc, errors, info, false); }

        /// <summary>
        /// confirm (needs errors): the blocks were corrected by their CRC (one byte each, see PstCrcRepair), a second check against a chance fit. An entry
        /// is kept only when the hash buckets, a second copy of it, agree (same GUID word and index; the hash of a number is the number, the
        /// hash of a name is the hash of that name); the others are dropped and listed.
        /// </summary>
        public static Dictionary<ushort, PstNameMap.Entry> Read(PstPcReader pc, List<string> errors, PstNameMapInfo info, bool confirm)
        {
            var names = new Dictionary<ushort, PstNameMap.Entry>();
            if (info == null) info = new PstNameMapInfo();
            PstPcReader.Entry bc = pc.Find(0x0001);
            if (bc != null && (bc.Type & 0x0FFF) == 0x0003) { info.HasBucketCount = true; info.BucketCount = bc.Hnid; }
            PstProp g = Stream(pc, 0x0002, "GUID stream", errors), e = Stream(pc, 0x0003, "entry stream", errors), s = Stream(pc, 0x0004, "string stream", errors);
            bool entriesLost = e == null && pc.Find(0x0003) != null;
            info.HasGuids = g != null && g.Data != null;
            info.HasStrings = s != null && s.Data != null;
            if (g != null && g.Data != null && g.Data.Length % 16 != 0) { Fail(errors, "Name-to-ID GUID stream of " + g.Data.Length + " bytes"); g = null; }
            // the hash buckets (0x1000 + n): (hash, GUID word, index)
            var recs = new List<Rec>();
            foreach (PstPcReader.Entry be in pc.Entries)
            {
                if (be.Id < 0x1000) continue;
                if (info.HasBucketCount && be.Id >= 0x1000 + info.BucketCount) continue;
                PstProp b;
                try { b = pc.Decode(be); }
                catch (PstFormatException x) { Fail(errors, "bucket 0x" + be.Id.ToString("X4") + ": " + x.Message); info.BucketMismatch++; continue; }
                if (b.Data == null) continue;
                info.Buckets++;
                for (int o = 0; o + 8 <= b.Data.Length; o += 8)
                    recs.Add(new Rec { Hash = PstFormat.G32(b.Data, o), Ng = PstFormat.G16(b.Data, o + 4), Idx = PstFormat.G16(b.Data, o + 6) });
            }
            info.BucketRecords = recs.Count;
            var byIndex = new Dictionary<int, int>();       // property index -> GUID word, to check the buckets against the entries
            if (e != null && e.Data != null)
            {
                byte[] ent = e.Data;
                if (ent.Length % 8 != 0) { Fail(errors, "Name-to-ID entry stream of " + ent.Length + " bytes"); return names; }
                info.Entries = ent.Length / 8;
                for (int i = 0; i < ent.Length; i += 8)
                {
                    try
                    {
                        uint dw = PstFormat.G32(ent, i);
                        int ng = PstFormat.G16(ent, i + 4), idx = PstFormat.G16(ent, i + 6);
                        PstNameMap.Entry en = MakeEntry(dw, ng, idx, g, s);
                        if (errors != null && names.ContainsKey(en.Id)) throw new PstFormatException("Name-to-ID property index " + idx + " appears twice");
                        names[en.Id] = en;
                        byIndex[idx] = ng;
                        info.Good++;
                    }
                    catch (PstFormatException x) { if (errors == null) throw; errors.Add("entry " + (i / 8) + ": " + x.Message); }
                }
                foreach (Rec r in recs)
                {
                    int word;
                    if (!byIndex.TryGetValue(r.Idx, out word) || word != r.Ng) info.BucketMismatch++;
                }
                if (confirm && errors != null)
                {
                    ILookup<int, Rec> byRec = recs.ToLookup(r => r.Idx);
                    foreach (PstNameMap.Entry en in names.Values.ToList())
                    {
                        int idx = en.Id - 0x8000;
                        uint hash = en.Name != null ? PstNameMap.NameHash(en.Name) : en.Lid.GetValueOrDefault();
                        if (byRec[idx].Any(r => r.Ng == byIndex[idx] && r.Hash == hash)) continue;
                        names.Remove(en.Id);
                        info.Unconfirmed++; info.Good--;
                        errors.Add("entry for index " + idx + " dropped: the hash buckets do not confirm it");
                    }
                }
            }
            else if (entriesLost && errors != null)
            {
                // the entry stream is gone: every entry is also in a bucket
                var strings = new Dictionary<uint, uint>();      // hash of a name -> its offset in the string stream
                if (s != null && s.Data != null)
                {
                    int off = 0;
                    while (off + 4 <= s.Data.Length)
                    {
                        uint len = PstFormat.G32(s.Data, off);
                        if (len > s.Data.Length - off - 4 || (len & 1) != 0) break;
                        strings[PstNameMap.NameHash(Encoding.Unicode.GetString(s.Data, off + 4, (int)len))] = (uint)off;
                        off += 4 + (int)len + (int)((4 - (len % 4)) % 4);
                    }
                }
                int rebuilt = 0, bad = 0;
                foreach (Rec r in recs)
                {
                    try
                    {
                        uint dw = r.Hash;
                        if ((r.Ng & 1) != 0 && !strings.TryGetValue(r.Hash, out dw)) throw new PstFormatException("no name of the string stream has hash 0x" + r.Hash.ToString("X8"));
                        PstNameMap.Entry en = MakeEntry(dw, r.Ng, r.Idx, g, s);
                        if (names.ContainsKey(en.Id)) continue;
                        names[en.Id] = en;
                        rebuilt++;
                    }
                    catch (PstFormatException x) { bad++; errors.Add("bucket record for index " + r.Idx + ": " + x.Message); }
                }
                info.Entries = rebuilt + bad; info.Good = rebuilt;
                errors.Add("entry stream unreadable: " + rebuilt + " entr" + (rebuilt == 1 ? "y" : "ies") + " rebuilt from the hash buckets" + (bad > 0 ? ", " + bad + " not recoverable" : ""));
            }
            return names;
        }

        static PstProp Stream(PstPcReader pc, ushort id, string what, List<string> errors)
        {
            try { return pc.Get(id); }
            catch (PstFormatException x) { if (errors == null) throw; errors.Add(what + ": " + x.Message); return null; }
        }

        static void Fail(List<string> errors, string message)
        {
            if (errors == null) throw new PstFormatException(message);
            errors.Add(message);
        }

        static PstNameMap.Entry MakeEntry(uint dw, int ng, int idx, PstProp g, PstProp s)
        {
            if (idx >= 0x8000) throw new PstFormatException("Name-to-ID entry with property index 0x" + idx.ToString("X"));
            bool isName = (ng & 1) != 0;
            int gi = ng >> 1;
            var en = new PstNameMap.Entry { Id = (ushort)(0x8000 + idx), GuidIndex = gi };
            if (gi == 1) en.Set = PstConvert.PsMapi;
            else if (gi == 2) en.Set = PstConvert.PsPublicStrings;
            else if (gi >= 3)
            {
                int off = (gi - 3) * 16;
                if (g == null || g.Data == null || off + 16 > g.Data.Length) throw new PstFormatException("Name-to-ID entry refers to GUID " + gi + " which is not in the GUID stream");
                var gb = new byte[16];
                Buffer.BlockCopy(g.Data, off, gb, 0, 16);
                en.Set = new Guid(gb);
            }
            if (isName)
            {
                if (s == null || s.Data == null || dw + 4 > s.Data.Length) throw new PstFormatException("Name-to-ID entry refers to a string outside the string stream");
                uint len = PstFormat.G32(s.Data, (int)dw);
                if (len > s.Data.Length - dw - 4 || (len & 1) != 0) throw new PstFormatException("Name-to-ID string length out of range");
                en.Name = Encoding.Unicode.GetString(s.Data, (int)dw + 4, (int)len);
            }
            else en.Lid = dw;
            return en;
        }
    }

    sealed partial class PstRepairRun
    {
        sealed class NameMapCandidate
        {
            public ulong Root, Sub;
            public Dictionary<ushort, PstNameMap.Entry> Names;
            public PstNameMapInfo Info;
            public int Errors, Resolves, Blocks;
            public bool Original, Corrected;
        }

        /// <summary>Maximum number of messages whose property ids are read to choose among several maps.</summary>
        const int UsageSample = 4000;

        void SetNameMap(Dictionary<ushort, PstNameMap.Entry> names, PstNameMapSource source, string detail, int candidates)
        {
            _reader.SetNameMap(names);
            _rep.NameMapSource = source;
            _rep.NameMapDetail = detail;
            _rep.NameMapEntries = names.Count;
            _rep.NameMapCandidates = candidates;
        }

        /// <summary>
        /// Decides which name-to-id map the repair reads the named properties with (see the top of this file) and tells the report.
        /// allowScan (levels Rebuild and Salvage): the map may be searched in the blocks of the file; at Salvage the scan of the repair is used,
        /// at Rebuild the file is scanned for this only. False (level Recompute): never searched.
        /// </summary>
        void ResolveNameMap(bool allowScan)
        {
            var errs = new List<string>();
            Dictionary<ushort, PstNameMap.Entry> orig = null;
            var origInfo = new PstNameMapInfo();
            bool hasNode;
            try { hasNode = _file.HasNode(PstFormat.NidNameToIdMap); }
            catch (Exception x) when (!(x is OperationCanceledException))
            {
                // a damaged page on the way to the node: Rebuild cannot walk the node B-tree at all (escalation, no scan for nothing),
                // Recompute goes on and the named properties are dropped as ever
                if (_level == PstRepairLevel.Rebuild) throw new PstRepairDiscoveryException("node B-tree: " + x.Message, x);
                hasNode = false; allowScan = false; errs.Add(x.Message);
            }
            if (hasNode)
            {
                try { orig = PstNameMapReader.Read(new PstPcReader(_file.OpenNode(PstFormat.NidNameToIdMap)), errs, origInfo); }
                catch (Exception x) when (!(x is OperationCanceledException)) { errs.Add(x.Message); orig = null; }
            }
            if (orig != null && errs.Count == 0)
            {
                SetNameMap(orig, PstNameMapSource.Original, "node 0x61 read without error: " + orig.Count + " entr" + (orig.Count == 1 ? "y" : "ies"), 0);
                return;
            }
            string why = !hasNode ? (errs.Count > 0 ? "node B-tree unreadable on the way to node 0x61 (" + errs[0] + ")" : "node 0x61 has no entry in the node B-tree") : "node 0x61 is damaged (" + (errs.Count > 0 ? errs[0] : "unreadable") + (errs.Count > 1 ? " and " + (errs.Count - 1) + " more" : "") + ")";
            if (hasNode) Event("NAMEMAP-UNREADABLE", "name-to-id map: " + string.Join("; ", errs.Take(5)) + (errs.Count > 5 ? " (+" + (errs.Count - 5) + ")" : ""), null, PstFormat.NidNameToIdMap);

            var cands = new List<NameMapCandidate>();
            if (allowScan)
            {
                try { cands = FindNameMapCandidates(); }
                catch (Exception x) when (!(x is OperationCanceledException)) { Event("NAMEMAP-SCAN-FAILED", "search for the name-to-id map failed (" + x.Message + ")", null, PstFormat.NidNameToIdMap); }
                if (hasNode)
                {
                    NameMapCandidate len = null;
                    try { len = ReadOriginalCorrected(); }
                    catch (Exception x) when (!(x is OperationCanceledException)) { }
                    if (len != null) cands.Add(len);
                }
            }
            int found = cands.Count;
            if (orig != null && orig.Count > 0) cands.Add(new NameMapCandidate { Names = orig, Info = origInfo, Errors = errs.Count, Original = true });
            if (cands.Count == 0)
            {
                SetNameMap(orig ?? new Dictionary<ushort, PstNameMap.Entry>(), PstNameMapSource.RebuiltFromUsage,
                    why + (allowScan ? "; no other map found in the blocks of the file" : "; not searched at this level (use Rebuild or Salvage)") + ": the named properties cannot be resolved and are dropped", 0);
                Event("NAMEMAP-REBUILT", _rep.NameMapDetail, null, PstFormat.NidNameToIdMap);
                return;
            }
            if (cands.Count > 1)
            {
                try { RankByUsage(cands); }
                catch (Exception x) when (!(x is OperationCanceledException)) { Event("NAMEMAP-RANK-FAILED", "the messages could not be read to choose among the maps (" + x.Message + "): the newest one with the most entries is used", null, PstFormat.NidNameToIdMap); }
            }
            NameMapCandidate best = cands.OrderByDescending(c => c.Resolves).ThenBy(c => c.Original ? 1 : 0).ThenByDescending(c => c.Names.Count).ThenByDescending(c => c.Root).First();
            string others = cands.Count > 1 ? "; " + cands.Count + " candidate(s)" + (_usageIds != null ? " (ids resolved: " + string.Join(", ", cands.OrderByDescending(c => c.Resolves).Select(c => c.Resolves.ToString())) + " of " + _usageIds.Count + " used)" : "") : "";
            if (best.Corrected && best.Original)
            {
                SetNameMap(best.Names, PstNameMapSource.RecoveredByScan, why + "; node 0x61 read after correcting a wrong byte in " + best.Blocks + " block(s) (the CRC gives the byte): " + best.Names.Count + " entr" + (best.Names.Count == 1 ? "y" : "ies") + (best.Info.Unconfirmed > 0 ? ", " + best.Info.Unconfirmed + " not confirmed by the hash buckets dropped" : "") + others, found);
                Event("NAMEMAP-RECOVERED", _rep.NameMapDetail, null, PstFormat.NidNameToIdMap);
            }
            else if (best.Original)
            {
                SetNameMap(best.Names, PstNameMapSource.RebuiltFromUsage, why + "; no better map found: " + best.Names.Count + " readable entr" + (best.Names.Count == 1 ? "y" : "ies") + ", the ids without entry are dropped" + others, found);
                Event("NAMEMAP-REBUILT", _rep.NameMapDetail, null, PstFormat.NidNameToIdMap);
            }
            else
            {
                SetNameMap(best.Names, PstNameMapSource.RecoveredByScan, why + "; map found in the blocks of the file (root block 0x" + best.Root.ToString("X") + (best.Sub != 0 ? ", subnodes 0x" + best.Sub.ToString("X") : "") + ", " + best.Names.Count + " entr" + (best.Names.Count == 1 ? "y" : "ies")
                    + (best.Corrected ? "; read after correcting a wrong byte in " + best.Blocks + " block(s) (the CRC gives the byte)" + (best.Info.Unconfirmed > 0 ? ", " + best.Info.Unconfirmed + " not confirmed by the hash buckets dropped" : "") : "") + ")" + others, found);
                Event("NAMEMAP-RECOVERED", _rep.NameMapDetail, null, PstFormat.NidNameToIdMap);
            }
        }

        /// <summary>
        /// Node 0x61 read after correcting the blocks that have exactly one wrong byte (PstCrcRepair; the scan keeps the damaged blocks aside,
        /// see PstScan.UseDamagedBlocks); an entry is kept only when the hash buckets confirm it. Null when no block needed correcting or nothing is left.
        /// </summary>
        NameMapCandidate ReadOriginalCorrected()
        {
            var scan = _file.Index as PstScan;
            _file.CorrectBlocks = true; _file.BlocksCorrected = 0;
            if (scan != null) scan.UseDamagedBlocks = true;
            try
            {
                var errors = new List<string>();
                var info = new PstNameMapInfo();
                var names = PstNameMapReader.Read(new PstPcReader(_file.OpenNode(PstFormat.NidNameToIdMap)), errors, info, true);
                if (_file.BlocksCorrected == 0 || names.Count == 0) return null;
                return new NameMapCandidate { Names = names, Info = info, Errors = errors.Count, Original = true, Corrected = true, Blocks = _file.BlocksCorrected };
            }
            finally { _file.CorrectBlocks = false; if (scan != null) scan.UseDamagedBlocks = false; }
        }

        HashSet<ushort> _usageIds;

        /// <summary>Named-property ids (0x8000 and up) of the properties of the messages (a sample of them), for choosing among several maps.</summary>
        void RankByUsage(List<NameMapCandidate> cands)
        {
            var nids = new List<uint>();
            foreach (PstNbtEntry n in _file.EnumerateNodes())
            {
                uint t = PstFormat.NidType(n.Nid);
                if (t == PstFormat.NtNormalMessage || t == PstFormat.NtAssocMessage) nids.Add(n.Nid);
            }
            int step = Math.Max(1, nids.Count / UsageSample);
            var ids = new HashSet<ushort>();
            for (int i = 0; i < nids.Count; i += step)
            {
                _ct.ThrowIfCancellationRequested();
                try
                {
                    var pc = new PstPcReader(_file.OpenNode(nids[i]));
                    foreach (PstPcReader.Entry e in pc.Entries) if (e.Id >= 0x8000) ids.Add(e.Id);
                }
                catch (Exception x) when (!(x is OperationCanceledException)) { }
            }
            _usageIds = ids;
            foreach (NameMapCandidate c in cands) c.Resolves = ids.Count(id => c.Names.ContainsKey(id));
        }

        /// <summary>
        /// Property contexts of the file that nothing refers to and that look like a name-to-id map (bucket count and entry stream present,
        /// every stream coherent, every bucket record an entry), with the subnode tree their long streams need.
        /// </summary>
        List<NameMapCandidate> FindNameMapCandidates()
        {
            PstFile sf = _file;
            PstScan scan = _file.Index as PstScan;
            bool own = false;
            if (scan == null)
            {
                sf = PstFile.OpenSalvage(_src);
                own = true;
                scan = PstScan.Scan(sf, _ct, _progress);
                sf.Index = scan;
                sf.CryptMethod = scan.DetectCrypt(sf, sf.CryptMethod);
                Event("NAMEMAP-SCAN", "name-to-id map to be found: " + scan.Describe(), null, PstFormat.NidNameToIdMap);
            }
            try
            {
                var res = new List<NameMapCandidate>();
                List<KeyValuePair<ulong, HashSet<uint>>> subs = null;
                for (int pass = 0; pass < 2; pass++)
                {
                bool lenient = pass == 1;        // second pass: the damaged blocks that start like a property context, corrected when one byte is wrong
                sf.CorrectBlocks = lenient; scan.UseDamagedBlocks = lenient;
                try
                {
                foreach (ulong root in lenient ? scan.UnreferencedDamagedHeads() : scan.UnreferencedRoots())
                {
                    _ct.ThrowIfCancellationRequested();
                    PstPcReader pc;
                    try { pc = new PstPcReader(new PstNodeView(sf, PstFormat.NidNameToIdMap, root, 0)); }
                    catch (Exception x) when (!(x is OperationCanceledException)) { continue; }
                    PstPcReader.Entry bc = pc.Find(0x0001), en = pc.Find(0x0003);
                    if (bc == null || en == null || bc.Hnid == 0 || bc.Hnid > 4096) continue;
                    var need = new List<uint>();
                    foreach (PstPcReader.Entry e in pc.Entries)
                        if (e.Type == 0x0102 && (e.Hnid & 0x1F) != 0 && (e.Id == 0x0002 || e.Id == 0x0003 || e.Id == 0x0004 || e.Id >= 0x1000)) need.Add(e.Hnid);
                    var subList = new List<ulong>();
                    if (need.Count == 0) subList.Add(0);
                    else
                    {
                        if (subs == null)
                        {
                            subs = new List<KeyValuePair<ulong, HashSet<uint>>>();
                            foreach (ulong sb in scan.UnreferencedSubRoots())
                            {
                                _ct.ThrowIfCancellationRequested();
                                try { subs.Add(new KeyValuePair<ulong, HashSet<uint>>(sb, new HashSet<uint>(sf.ReadSubnodes(sb).Select(x => x.Nid)))); }
                                catch (Exception x) when (!(x is OperationCanceledException)) { }
                            }
                        }
                        foreach (var kv in subs) if (need.All(kv.Value.Contains)) subList.Add(kv.Key);
                    }
                    foreach (ulong sub in subList)
                    {
                        var errors = new List<string>();
                        var info = new PstNameMapInfo();
                        Dictionary<ushort, PstNameMap.Entry> names;
                        sf.BlocksCorrected = 0;
                        try { names = PstNameMapReader.Read(new PstPcReader(new PstNodeView(sf, PstFormat.NidNameToIdMap, root, sub)), errors, info, lenient); }
                        catch (Exception x) when (!(x is OperationCanceledException)) { continue; }
                        if (names.Count == 0 || info.Buckets == 0) continue;
                        if (lenient ? sf.BlocksCorrected == 0 : (errors.Count > 0 || info.BucketMismatch > 0)) continue;
                        res.Add(new NameMapCandidate { Root = root, Sub = sub, Names = names, Info = info, Errors = errors.Count, Corrected = lenient, Blocks = sf.BlocksCorrected });
                    }
                }
                }
                finally { sf.CorrectBlocks = false; scan.UseDamagedBlocks = false; }
                }
                return res;
            }
            finally { if (own) sf.Dispose(); }
        }
    }
}
