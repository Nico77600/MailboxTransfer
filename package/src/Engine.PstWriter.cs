// =============================================================================
//  Mailbox Transfer - engine, part 5d: MS-PST writer, Messaging layer (message store, folders, messages, named properties)
// -----------------------------------------------------------------------------
//    PstConvert      FxProperty <-> PstProp (strings in UTF-16 without terminator, Boolean on one byte, string8 turned into Unicode)
//    PstNameMap      the Name-to-ID map (2.4.7): GUID stream, entry stream, string stream, hash buckets; ids 0x8000+ per PST
//    PstWriter       Create / CreateFolder / AddMessage / Close on top of PstNdbWriter and the LTP builders
//
//  What is written when (so that memory stays bounded):
//    - a message is written to the file as soon as AddMessage is called (its node, its subnodes, its large values);
//      only one small row for the contents table of its folder stays in memory (about 250 bytes);
//    - the folder nodes (property context, hierarchy, contents and FAI tables), the message store, the name-to-id map, the
//      template tables and the search objects are written at Close, when the counts are final;
//    - the B-trees: see PstNdbWriter.
//
//  Properties that the PST format computes or does not carry (so they are not stored; the reader gives them back when it can):
//    PidTagMessageRecipients / PidTagMessageAttachments (0x0E12 / 0x0E13: the markers of the sub-objects), PidTagParentEntryId,
//    PidTagEntryId, PidTagRecordKey, PidTagStoreRecordKey, PidTagStoreEntryId, PidTagObjectType, PidTagHasAttachments (derived from
//    PidTagMessageFlags), the Exchange identifiers PidTagFolderId / PidTagParentFolderId / PidTagMid / PidTagInstID..., PidTagAssociated
//    (derived from the node type), PidTagAttachNumber (derived from the position). PidTagMessageSize and PidTagAttachSize are
//    recomputed.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public static class PstConvert
    {
        /// <summary>Properties that are not stored in a PST (computed from the node, or Exchange identifiers).</summary>
        public static readonly HashSet<ushort> NotStored = new HashSet<ushort>
        {
            0x0E12, 0x0E13, 0x0E09, 0x0E1B, 0x0FF9, 0x0FFA, 0x0FFB, 0x0FFE, 0x0FFF,
            0x6748, 0x6749, 0x674A, 0x674D, 0x67A4, 0x67AA, 0x6740, 0x6741, 0x6743, 0x6744
        };

        public static readonly Guid PsMapi = new Guid("00020328-0000-0000-C000-000000000046");
        public static readonly Guid PsPublicStrings = new Guid("00020329-0000-0000-C000-000000000046");

        static byte[] StripNul(byte[] raw)
        {
            int n = raw.Length;
            if (n >= 2 && raw[n - 1] == 0 && raw[n - 2] == 0) n -= 2;
            if (n == raw.Length) return raw;
            var r = new byte[n];
            Buffer.BlockCopy(raw, 0, r, 0, n);
            return r;
        }

        static byte[] Utf16(FxTypeInfo t, byte[] raw)
        {
            if (t.CodePageString && t.CodePage == 1200 || (!t.CodePageString && t.BaseType == 0x001F)) return StripNul(raw);
            string s = FxValue.DecodeString(raw, t.BaseType, t.CodePageString ? t.CodePage : 1252);
            if (s.EndsWith("\0")) s = s.TrimEnd('\0');
            return Encoding.Unicode.GetBytes(s);
        }

        /// <summary>The PST form of a FastTransfer property (id already mapped); null for a type that has no PST form.</summary>
        public static PstProp FromFx(FxProperty p, ushort id)
        {
            FxTypeInfo t = p.TypeInfo;
            if (!t.Valid || t.IsObject) return null;
            int bt = t.BaseType;
            if (bt == 0x0001) return null;
            ushort type;
            bool isString = bt == 0x001E || bt == 0x001F;
            if (isString) type = 0x001F;
            else if (bt == 0x00FB) type = 0x0102;
            else type = (ushort)bt;
            if (t.MultiValued) type |= 0x1000;
            var r = new PstProp { Id = id, Type = type };
            if (!t.MultiValued)
            {
                if (p.IsLazy)
                {
                    r.Source = LazyOne(t, p.Lazy);
                    return r.Source == null ? null : r;
                }
                if (p.Raw == null) return null;
                r.Data = One(t, p.Raw);
                if (r.Data == null) return null;
            }
            else
            {
                r.Items = new List<byte[]>();
                foreach (byte[] it in p.Items) { byte[] d = One(t, it); if (d == null) return null; r.Items.Add(d); }
            }
            return r;
        }

        /// <summary>The PST form of a lazy value, still read from its source when it is written (null: not a type that can be lazy).</summary>
        static FxBlob LazyOne(FxTypeInfo t, FxBlob b)
        {
            int bt = t.BaseType;
            if (bt == 0x0102 || bt == 0x00FB) return b;
            if (bt != 0x001E && bt != 0x001F) return null;
            if (t.CodePageString && t.CodePage == 1200 || (!t.CodePageString && bt == 0x001F))
            {
                long n = b.Length;
                if (n >= 2) { byte[] tail = b.ReadRange(n - 2, 2); if (tail[0] == 0 && tail[1] == 0) return b.Slice(0, n - 2); }
                return b;
            }
            Encoding enc;
            try { enc = Encoding.GetEncoding(t.CodePageString ? t.CodePage : 1252); } catch (Exception) { enc = Encoding.Latin1; }
            return new FxTranscodedBlob(b, enc);
        }

        static byte[] One(FxTypeInfo t, byte[] raw)
        {
            int bt = t.BaseType;
            if (bt == 0x001E || bt == 0x001F) return Utf16(t, raw);
            if (bt == 0x000B) return new byte[] { (byte)((raw.Length > 0 && raw[0] != 0) || (raw.Length > 1 && raw[1] != 0) ? 1 : 0) };
            int fs = PstTypes.FixedSize((ushort)bt);
            if (fs > 0) { if (raw.Length < fs) return null; if (raw.Length == fs) return raw; var c = new byte[fs]; Buffer.BlockCopy(raw, 0, c, 0, fs); return c; }
            return raw;     // binary, server id
        }

        /// <summary>FastTransfer form of a PST property: string with its terminator, Boolean on 2 bytes.</summary>
        public static FxProperty ToFx(PstProp p)
        {
            ushort bt = (ushort)(p.Type & 0x0FFF);
            bool multi = (p.Type & 0x1000) != 0;
            if (!multi && p.Source != null)
                return FxProperty.CreateLazy(p.Id, p.Type, bt == 0x001F ? new FxConcatBlob(p.Source, FxBytesBlob.Terminator) : p.Source);
            if (!multi) return FxProperty.Create(p.Id, p.Type, OneToFx(bt, p.Data));
            return FxProperty.CreateMulti(p.Id, p.Type, p.Items.Select(i => OneToFx(bt, i)));
        }

        static byte[] OneToFx(ushort bt, byte[] d)
        {
            if (bt == 0x001F)
            {
                var r = new byte[d.Length + 2];
                Buffer.BlockCopy(d, 0, r, 0, d.Length);
                return r;
            }
            if (bt == 0x000B) return new byte[] { (byte)(d.Length > 0 && d[0] != 0 ? 1 : 0), 0 };
            return d;
        }

        public static string StringOf(PstProp p) { return p == null || p.Data == null ? null : Encoding.Unicode.GetString(p.Data); }
        public static int Int32Of(PstProp p, int dflt) { return p == null || p.Data == null || p.Data.Length < 4 ? dflt : BitConverter.ToInt32(p.Data, 0); }
    }

    /// <summary>The Name-to-ID map (2.4.7). Ids are assigned in order of first use, from 0x8000.</summary>
    public sealed class PstNameMap
    {
        public sealed class Entry { public Guid Set; public uint? Lid; public string Name; public ushort Id; public int GuidIndex; }

        public readonly List<Entry> Entries = new List<Entry>();
        readonly Dictionary<string, Entry> _byKey = new Dictionary<string, Entry>();
        readonly List<Guid> _guids = new List<Guid>();
        public const int BucketCount = 251;

        static string Key(Guid g, uint? lid, string name)
        {
            return g.ToString("N") + (lid.HasValue ? "|L" + lid.Value : "|S" + name);
        }

        public ushort Resolve(Guid set, uint? lid, string name)
        {
            string k = Key(set, lid, name);
            Entry e;
            if (_byKey.TryGetValue(k, out e)) return e.Id;
            if (Entries.Count >= 0x7FFF) throw new InvalidOperationException("Too many named properties for one PST.");
            int gi;
            if (set == PstConvert.PsMapi) gi = 1;
            else if (set == PstConvert.PsPublicStrings) gi = 2;
            else
            {
                int idx = _guids.IndexOf(set);
                if (idx < 0) { _guids.Add(set); idx = _guids.Count - 1; }
                gi = idx + 3;
            }
            e = new Entry { Set = set, Lid = lid, Name = name, Id = (ushort)(0x8000 + Entries.Count), GuidIndex = gi };
            Entries.Add(e);
            _byKey[k] = e;
            return e.Id;
        }

        public static uint NameHash(string name) { return PstFormat.Crc(Encoding.Unicode.GetBytes(name), 0, Encoding.Unicode.GetByteCount(name)); }

        public PstNode Build()
        {
            var props = new List<PstProp>();
            props.Add(PstProp.Int32(0x0001, BucketCount));
            if (Entries.Count > 0)
            {
                var guids = new byte[_guids.Count * 16];
                for (int i = 0; i < _guids.Count; i++) Buffer.BlockCopy(_guids[i].ToByteArray(), 0, guids, 16 * i, 16);
                var strs = new MemoryStream();
                var entries = new byte[Entries.Count * 8];
                var buckets = new List<byte[]>[BucketCount];
                for (int i = 0; i < Entries.Count; i++)
                {
                    Entry e = Entries[i];
                    uint dwProp, dwHash;
                    ushort nguid;
                    if (e.Lid.HasValue) { dwProp = dwHash = e.Lid.Value; nguid = (ushort)(e.GuidIndex << 1); }
                    else
                    {
                        byte[] nb = Encoding.Unicode.GetBytes(e.Name);
                        dwProp = (uint)strs.Position;
                        strs.Write(BitConverter.GetBytes((uint)nb.Length), 0, 4);
                        strs.Write(nb, 0, nb.Length);
                        int pad = (4 - (nb.Length % 4)) % 4;
                        for (int k = 0; k < pad; k++) strs.WriteByte(0);
                        dwHash = NameHash(e.Name);
                        nguid = (ushort)((e.GuidIndex << 1) | 1);
                    }
                    PstFormat.P32(entries, 8 * i, dwProp); PstFormat.P16(entries, 8 * i + 4, nguid); PstFormat.P16(entries, 8 * i + 6, i);
                    var h = new byte[8];
                    PstFormat.P32(h, 0, dwHash); PstFormat.P16(h, 4, nguid); PstFormat.P16(h, 6, i);
                    int bucket = (int)((dwHash ^ (uint)nguid) % BucketCount);
                    if (buckets[bucket] == null) buckets[bucket] = new List<byte[]>();
                    buckets[bucket].Add(h);
                }
                if (guids.Length > 0) props.Add(PstProp.Bin(0x0002, guids));
                props.Add(PstProp.Bin(0x0003, entries));
                if (strs.Length > 0) props.Add(PstProp.Bin(0x0004, strs.ToArray()));
                for (int b = 0; b < BucketCount; b++)
                {
                    if (buckets[b] == null) continue;
                    props.Add(PstProp.Bin((ushort)(0x1000 + b), buckets[b].SelectMany(x => x).ToArray()));
                }
            }
            return PstPc.Build(props);
        }
    }

    public sealed class PstFolder
    {
        internal uint Nid;
        internal PstFolder Parent;
        internal readonly List<PstFolder> Children = new List<PstFolder>();
        internal PstTable Contents, Fai;
        internal int ContentCount, UnreadCount, FaiCount;
        internal bool IsSearchFolder;
        internal byte[] ReplId = Guid.NewGuid().ToByteArray();
        public string Name { get; internal set; }
        public string ContainerClass { get; internal set; }
        public int MessageCount { get { return ContentCount; } }
        public uint FolderNid { get { return Nid; } }
    }

    public sealed class PstWriterOptions
    {
        /// <summary>Write the "SPAM Search Folder" of the minimal PST (2.7.3.4.4); it needs the search machinery (update queue, criteria object).</summary>
        public bool SpamSearchFolder = false;
        /// <summary>Name of the IPM subtree folder (default: the display name of the store, like Outlook-less PSTs of this tool).</summary>
        public string IpmSubtreeName;
        /// <summary>
        /// Deliberate defects, for the tests of the integrity checker only: contents-row-removed, dangling-nid, wrong-count,
        /// hierarchy-missing, no-recipient-table, attach-table-short, template-column-missing, no-outgoing-queue, no-receive-row,
        /// no-template-6b6, no-ec1, no-hmp, hmp-wrong-entry, message-size-wrong, attach-size-wrong, embedded-nid-type,
        /// contents-cell-stale, hierarchy-row-value, hidden-count-wrong, attach-highwater, namemap-empty, decoy-namemap. Never set in production.
        /// </summary>
        public string[] Faults = new string[0];
    }

    public sealed class PstWriter : IDisposable
    {
        // ---- template columns (2.4.4.4.1, 2.4.4.5.1, 2.4.4.6.1, 2.4.5.3.1, 2.4.6.1.1, 2.4.8.6.2.1)
        static readonly uint[] HierarchyCols = { 0x0E300102, 0x0E330014, 0x0E340102, 0x0E380003, 0x3001001F, 0x36020003, 0x36030003, 0x360A000B, 0x3613001F, 0x66350003, 0x66360003 };
        static readonly uint[] ContentsCols =
        {
            0x00170003, 0x001A001F, 0x00360003, 0x0037001F, 0x00390040, 0x0042001F, 0x0057000B, 0x0058000B, 0x0070001F, 0x00710102,
            0x0E03001F, 0x0E04001F, 0x0E060040, 0x0E070003, 0x0E080003, 0x0E170003, 0x0E300102, 0x0E330014, 0x0E340102, 0x0E380003,
            0x0E3C0102, 0x0E3D0102, 0x10970003, 0x30080040, 0x30130102, 0x65C60003
        };
        static readonly uint[] FaiCols = { 0x001A001F, 0x0E070003, 0x0E170003, 0x3001001F, 0x6800001F, 0x6803000B, 0x68051003, 0x70030003, 0x70040102, 0x70050102, 0x7006001F, 0x70070003, 0x682F001F };
        static readonly uint[] SearchContentsCols =
        {
            0x00170003, 0x001A001F, 0x00360003, 0x0037001F, 0x0042001F, 0x0057000B, 0x0058000B, 0x0E03001F, 0x0E04001F, 0x0E05001F, 0x0E060040,
            0x0E070003, 0x0E080003, 0x0E170003, 0x0E2A000B, 0x30080040, 0x67F10003
        };
        static readonly uint[] OutgoingQueueCols =
        {
            0x00170003, 0x001A001F, 0x00360003, 0x0037001F, 0x00390040, 0x0042001F, 0x0057000B, 0x0058000B, 0x0E03001F, 0x0E04001F, 0x0E05001F,
            0x0E060040, 0x0E070003, 0x0E080003, 0x0E100003, 0x0E140003, 0x0E170003, 0x0E2A000B, 0x30080040, 0x67F10003
        };
        static readonly uint[] RecipientCols = { 0x0C150003, 0x0E0F000B, 0x0FF90102, 0x0FFE0003, 0x0FFF0102, 0x3001001F, 0x3002001F, 0x3003001F, 0x300B0102, 0x39000003, 0x39FF001F, 0x3A40000B };
        static readonly uint[] AttachmentCols = { 0x0E200003, 0x3704001F, 0x37050003, 0x370B0003 };

        readonly PstNdbWriter _ndb;
        readonly PstWriterOptions _opts;
        bool Fault(string name) { return _opts.Faults != null && Array.IndexOf(_opts.Faults, name) >= 0; }
        readonly string _displayName;
        readonly byte[] _recordKey = Guid.NewGuid().ToByteArray();
        readonly PstNameMap _names = new PstNameMap();
        readonly List<PstFolder> _folders = new List<PstFolder>();
        readonly Random _rnd = new Random();
        public readonly PstFolder Root, IpmSubtree, DeletedItems, SearchRoot;
        PstFolder _spam;
        readonly List<KeyValuePair<byte[], uint>> _hmp = new List<KeyValuePair<byte[], uint>>();
        long _pendingBytes;
        bool _closed;
        public long MessagesWritten, AttachmentsWritten, RecipientsWritten, EmbeddedMessagesWritten;
        public readonly Dictionary<uint, int> DroppedProperties = new Dictionary<uint, int>();
        public string Path { get { return _ndb.Path; } }

        PstWriter(string path, string displayName, PstWriterOptions opts)
        {
            _opts = opts ?? new PstWriterOptions();
            _displayName = displayName ?? "";
            _ndb = new PstNdbWriter(path);
            PstFormat.P32(_replHistory, 0, 1); Buffer.BlockCopy(Guid.NewGuid().ToByteArray(), 0, _replHistory, 4, 16); PstFormat.P32(_replHistory, 20, 1);
            // folder NIDs: the root and the search objects have fixed NIDs, the other folders take 0x401.. of the folder type
            Root = NewFolder(null, "", null, PstFormat.NidRootFolder, false);
            IpmSubtree = NewFolder(Root, !string.IsNullOrEmpty(_opts.IpmSubtreeName) ? _opts.IpmSubtreeName : _displayName.Length > 0 ? _displayName : "Top of Personal Folders", null, 0, false);
            SearchRoot = NewFolder(Root, "Search Root", null, 0, false);
            DeletedItems = NewFolder(IpmSubtree, "Deleted Items", "IPF.Note", 0, false);
            if (_opts.SpamSearchFolder) _spam = NewFolder(Root, "SPAM Search Folder 2", null, PstFormat.Nid(PstFormat.NtSearchFolder, 0x111), true);
        }

        public static PstWriter Create(string path, string displayName) { return new PstWriter(path, displayName, null); }
        public static PstWriter Create(string path, string displayName, PstWriterOptions options) { return new PstWriter(path, displayName, options); }

        PstFolder NewFolder(PstFolder parent, string name, string cls, uint fixedNid, bool search)
        {
            var f = new PstFolder { Parent = parent, Name = name, ContainerClass = cls, IsSearchFolder = search };
            f.Nid = fixedNid != 0 ? fixedNid : PstFormat.Nid(PstFormat.NtNormalFolder, ++_ndb.Rgnid[PstFormat.NtNormalFolder]);
            f.Contents = new PstTable(ContentsCols);
            f.Fai = new PstTable(FaiCols);
            if (parent != null) parent.Children.Add(f);
            _folders.Add(f);
            _pendingBytes += 6000;
            return f;
        }

        /// <summary>Creates a folder under parent (the IPM subtree when null). The name is kept as given; folders of the same name are allowed.</summary>
        public PstFolder CreateFolder(PstFolder parent, string name, string containerClass)
        {
            if (_closed) throw new InvalidOperationException("The PST is closed.");
            PstFolder p = parent ?? IpmSubtree;
            if (p.IsSearchFolder) throw new ArgumentException("A search folder cannot have subfolders.");
            return NewFolder(p, name ?? "", containerClass ?? "IPF.Note", 0, false);
        }

        /// <summary>Size the file would have if it were closed now (blocks written + estimated tables and trees still to write).</summary>
        public long ProjectedSize { get { return _ndb.ProjectedSize + _pendingBytes + _pendingBytes / 8; } }

        /// <summary>Bytes of the FastTransfer message that count for the size of the PST (values + overhead per object), to decide when to roll.</summary>
        public static long EstimateSize(FxMessage m)
        {
            long n = 4096;
            foreach (FxProperty p in m.Properties) n += Size(p);
            foreach (FxRecipient r in m.Recipients) { n += 600; foreach (FxProperty p in r.Properties) n += Size(p); }
            foreach (FxAttachment a in m.Attachments)
            {
                n += 1500;
                foreach (FxProperty p in a.Properties) n += Size(p);
                if (a.Embedded != null) n += EstimateSize(a.Embedded);
            }
            return n;
        }

        static long Size(FxProperty p)
        {
            long n = 24;
            n += p.RawLength;
            if (p.Items != null) foreach (byte[] i in p.Items) n += i.Length + 4;
            return n;
        }

        // ---------------------------------------------------------------- messages
        sealed class Built { public PstNode Node; public List<PstProp> Props; public long Size; public bool Unread; public bool Fai; }

        ushort MapId(FxProperty p)
        {
            if (!p.IsNamed) return p.Id;
            if (p.NamedGuid == null || (p.NamedLid == null && p.NamedName == null)) return 0;
            return _names.Resolve(p.NamedGuid.Value, p.NamedLid, p.NamedName);
        }

        void Drop(FxProperty p)
        {
            uint tag = p.Tag;
            int n;
            DroppedProperties.TryGetValue(tag, out n);
            DroppedProperties[tag] = n + 1;
        }

        List<PstProp> ConvertProps(IEnumerable<FxProperty> src, bool attachment)
        {
            var res = new List<PstProp>();
            var seen = new HashSet<ushort>();
            foreach (FxProperty p in src)
            {
                if (!p.IsNamed && PstConvert.NotStored.Contains(p.Id) && !(attachment && false)) continue;
                if (attachment && !p.IsNamed && (p.Id == 0x0E21 || p.Id == 0x0E20)) continue;
                ushort id = MapId(p);
                if (id == 0 || !seen.Add(id)) { Drop(p); continue; }
                PstProp q = PstConvert.FromFx(p, id);
                if (q == null) { seen.Remove(id); if (p.Tag >> 16 != 0x3701 && (p.Type & 0x0FFF) != 0x0001) Drop(p); continue; }
                res.Add(q);
            }
            return res;
        }

        static PstProp Find(List<PstProp> l, ushort id) { return l.FirstOrDefault(p => p.Id == id); }

        static void Set(List<PstProp> l, PstProp p)
        {
            l.RemoveAll(x => x.Id == p.Id);
            l.Add(p);
        }

        static long Bytes(PstProp p)
        {
            long n = 16;
            n += p.DataLength;
            if (p.Items != null) foreach (byte[] i in p.Items) n += i.Length + 4;
            return n;
        }

        Built BuildMessage(FxMessage m, bool embedded)
        {
            List<PstProp> props = ConvertProps(m.Properties, false);
            long size = 0;

            // recipients
            var recipProps = new List<List<PstProp>>();
            foreach (FxRecipient r in m.Recipients) recipProps.Add(ConvertProps(r.Properties, false));
            string to = "", cc = "", bcc = "";
            for (int i = 0; i < recipProps.Count; i++)
            {
                PstProp rt = Find(recipProps[i], 0x0C15), dn = Find(recipProps[i], 0x3001);
                string name = PstConvert.StringOf(dn) ?? PstConvert.StringOf(Find(recipProps[i], 0x39FE)) ?? PstConvert.StringOf(Find(recipProps[i], 0x3003)) ?? "";
                int type = PstConvert.Int32Of(rt, 1) & 0xF;
                if (type == 1) to += (to.Length > 0 ? "; " : "") + name;
                else if (type == 2) cc += (cc.Length > 0 ? "; " : "") + name;
                else if (type == 3) bcc += (bcc.Length > 0 ? "; " : "") + name;
            }
            if (Find(props, 0x0E04) == null && to.Length > 0) props.Add(PstProp.Str(0x0E04, to));
            if (Find(props, 0x0E03) == null && cc.Length > 0) props.Add(PstProp.Str(0x0E03, cc));
            if (Find(props, 0x0E02) == null && bcc.Length > 0) props.Add(PstProp.Str(0x0E02, bcc));

            // PidTagConversationTopic is a calculated property of the PST (the subject without its prefix): the contents table row
            // and the property context must agree on it
            if (Find(props, 0x0070) == null) { PstProp sj = Find(props, 0x0037); if (sj != null && sj.Data != null) props.Add(PstProp.Make(0x0070, 0x001F, sj.Data)); }
            // defaults required by 2.4.5.1.1
            if (Find(props, 0x001A) == null) props.Add(PstProp.Str(0x001A, "IPM.Note"));
            PstProp flags = Find(props, 0x0E07);
            int mf = flags == null ? 1 : PstConvert.Int32Of(flags, 1);
            if (m.Attachments.Count > 0) mf |= 0x10; else mf &= ~0x10;
            Set(props, PstProp.Int32(0x0E07, mf));
            if (Find(props, 0x0E17) == null) props.Add(PstProp.Int32(0x0E17, 0));
            DateTime now = DateTime.UtcNow;
            PstProp when = Find(props, 0x0039) ?? Find(props, 0x0E06);
            DateTime created = when != null && when.Data != null && when.Data.Length == 8 ? SafeTime(BitConverter.ToInt64(when.Data, 0), now) : now;
            if (Find(props, 0x3007) == null) props.Add(PstProp.Time(0x3007, created));
            if (Find(props, 0x3008) == null) props.Add(PstProp.Time(0x3008, created));
            if (Find(props, 0x300B) == null) { var k = new byte[16]; _rnd.NextBytes(k); props.Add(PstProp.Bin(0x300B, k)); }

            // recipient table
            PstNode node;
            var node0 = new PstNode();
            PstNode recNode = BuildRecipientTable(recipProps, ref size);
            if (!Fault("no-recipient-table")) node0.Subnodes[PstFormat.NidRecipientTable] = recNode;
            RecipientsWritten += m.Recipients.Count;

            // attachments
            if (m.Attachments.Count > 0)
            {
                var attProps = new List<List<PstProp>>();
                var attNids = new List<uint>();
                for (int i = 0; i < m.Attachments.Count; i++)
                {
                    FxAttachment a = m.Attachments[i];
                    List<PstProp> ap = ConvertProps(a.Properties, true);
                    uint anid = PstFormat.Nid(PstFormat.NtAttachment, ++_ndb.Rgnid[PstFormat.NtAttachment]);
                    long asize = 0;
                    if (a.Embedded != null)
                    {
                        Built eb = BuildMessage(a.Embedded, true);
                        ap.RemoveAll(x => x.Id == 0x3701);
                        ap.Add(new PstProp { Id = 0x3701, Type = 0x000D, Object = eb.Node, ObjectSize = eb.Node.TotalLength, ObjectNid = Fault("embedded-nid-type") ? 0 : PstFormat.Nid(PstFormat.NtNormalMessage, ++_ndb.Rgnid[PstFormat.NtNormalMessage]) });
                        asize = eb.Node.TotalLength;
                        if (Find(ap, 0x3705) == null) ap.Add(PstProp.Int32(0x3705, 5));
                        EmbeddedMessagesWritten++;
                    }
                    else
                    {
                        PstProp data = Find(ap, 0x3701);
                        if (data != null) data.ForceSubnode = true;
                        if (data != null) asize = data.DataLength;
                        if (Find(ap, 0x3705) == null) ap.Add(PstProp.Int32(0x3705, 1));
                    }
                    if (Find(ap, 0x370B) == null) ap.Add(PstProp.Int32(0x370B, -1));
                    foreach (PstProp p in ap) if (p.Id != 0x3701) asize += Bytes(p);
                    // PidTagAttachSize is the sum of the sizes of the values of the attachment properties (the data, or the embedded
                    // message, included): that is what Outlook recomputes
                    asize = 0;
                    ap.Add(PstProp.Int32(0x0E20, 0));
                    foreach (PstProp p in ap)
                    {
                        if (p.Object != null) asize += p.ObjectSize;
                        else if (p.Data != null) asize += p.Data.Length;
                        else if (p.Source != null) asize += p.Source.Length;
                        else if (p.Items != null) foreach (byte[] it in p.Items) asize += it.Length;
                    }
                                        Set(ap, PstProp.Int32(0x0E20, (int)Math.Min(int.MaxValue, asize) + (Fault("attach-size-wrong") ? 3 : 0)));
                    size += asize;
                    attProps.Add(ap);
                    attNids.Add(anid);
                    AttachmentsWritten++;
                }
                // attachment table: the template columns plus the small properties found on the attachments
                var cols = new List<uint>(AttachmentCols);
                foreach (List<PstProp> ap in attProps)
                    foreach (PstProp p in ap)
                        if (p.Id != 0x3701 && !p.IsMulti && p.Object == null && Bytes(p) <= 300 && !cols.Contains(p.Tag) && cols.Count < 100) cols.Add(p.Tag);
                var at = new PstTable(cols);
                for (int i = 0; i < attProps.Count; i++)
                {
                    at.AddRow(attNids[i], at.Pick(attProps[i]));
                    PstNode an = PstPc.Build(attProps[i]);
                    node0.Subnodes[attNids[i]] = an;
                }
                if (Fault("attach-table-short")) at.RemoveLastRow();
                node0.Subnodes[PstFormat.NidAttachmentTable] = at.Build();
            }

            // PidTagMessageSize is the total length of the data of the node and of all its subnodes (what Outlook recomputes and
            // compares with the contents table row): build the property context with a placeholder, measure, build it again
            Set(props, PstProp.Int32(0x0E08, 0));
            node = PstPc.Build(props);
            long subTotal = 0;
            foreach (PstNode sn in node0.Subnodes.Values) subTotal += sn.TotalLength;
            long total = node.TotalLength + subTotal;
            Set(props, PstProp.Int32(0x0E08, (int)Math.Min(int.MaxValue, total) + (Fault("message-size-wrong") ? 7 : 0)));
            node = PstPc.Build(props);
            foreach (KeyValuePair<uint, PstNode> kv in node0.Subnodes) node.Subnodes[kv.Key] = kv.Value;
            bool fai = m.StartMarker == FxMessage.MarkerStartFAIMsg;
            FxProperty assoc = m.Find(0x67AA);
            if (assoc != null && assoc.Raw != null && assoc.Raw.Length > 0 && assoc.Raw[0] != 0) fai = true;
            return new Built { Node = node, Props = props, Size = total, Unread = (mf & 1) == 0, Fai = fai };
        }

        static DateTime SafeTime(long ft, DateTime dflt)
        {
            try { if (ft > 0 && ft < 2650467743999999999L) return DateTime.FromFileTimeUtc(ft); } catch (Exception) { }
            return dflt;
        }

        PstNode BuildRecipientTable(List<List<PstProp>> rows, ref long size)
        {
            var cols = new List<uint>(RecipientCols);
            foreach (List<PstProp> rp in rows)
                foreach (PstProp p in rp)
                    if (p.Object == null && !cols.Contains(p.Tag) && cols.Count < 120) cols.Add(p.Tag);
            var t = new PstTable(cols);
            for (int i = 0; i < rows.Count; i++)
            {
                t.AddRow((uint)(i + 1), t.Pick(rows[i]));
                foreach (PstProp p in rows[i]) size += Bytes(p);
            }
            return t.Build();
        }

        long _valueBytes;

        /// <summary>Writes a message in a folder and returns its NID.</summary>
        public uint AddMessage(PstFolder folder, FxMessage message)
        {
            return AddMessage(folder, message, System.Threading.CancellationToken.None, null);
        }

        /// <summary>
        /// Same, cancellable while a big value is written (OperationCanceledException between two blocks) and reporting the bytes of
        /// values streamed (BytesDone, cumulative for this writer), the messages written and the folder. After a cancellation in
        /// the middle of a message the file holds blocks that no node refers to: dispose the writer without Close and delete the file.
        /// </summary>
        public uint AddMessage(PstFolder folder, FxMessage message, System.Threading.CancellationToken ct, IProgress<PstProgress> progress)
        {
            ct.ThrowIfCancellationRequested();
            _ndb.Cancellation = ct;
            _ndb.OnBytes = n =>
            {
                _valueBytes += n;
                if (progress != null) progress.Report(new PstProgress { Phase = "write", BytesDone = _valueBytes, Items = MessagesWritten, CurrentFolder = folder.Name });
            };
            try
            {
                uint nid = AddMessageCore(folder, message);
                if (progress != null) progress.Report(new PstProgress { Phase = "write", BytesDone = _valueBytes, Items = MessagesWritten, CurrentFolder = folder.Name });
                return nid;
            }
            finally { _ndb.Cancellation = default(System.Threading.CancellationToken); _ndb.OnBytes = null; }
        }

        uint AddMessageCore(PstFolder folder, FxMessage message)
        {
            if (_closed) throw new InvalidOperationException("The PST is closed.");
            if (folder == null) throw new ArgumentNullException("folder");
            if (folder.IsSearchFolder) throw new ArgumentException("Messages cannot be added to a search folder.");
            Built b = BuildMessage(message, false);
            uint nid = b.Fai
                ? PstFormat.Nid(PstFormat.NtAssocMessage, ++_ndb.Rgnid[PstFormat.NtAssocMessage])
                : PstFormat.Nid(PstFormat.NtNormalMessage, ++_ndb.Rgnid[PstFormat.NtNormalMessage]);
            _ndb.WriteTopLevelNode(nid, b.Node, folder.Nid);
            byte[] replId = Guid.NewGuid().ToByteArray();
            if (!b.Fai) _hmp.Add(new KeyValuePair<byte[], uint>(replId, nid));
            if (b.Fai) { folder.Fai.AddRow(nid, folder.Fai.Pick(RowProps(b.Props, replId)), NextRowVer()); folder.FaiCount++; }
            else
            {
                folder.Contents.AddRow(nid, folder.Contents.Pick(RowProps(b.Props, replId)), NextRowVer());
                folder.ContentCount++;
                if (b.Unread) folder.UnreadCount++;
            }
            _pendingBytes += 260;
            MessagesWritten++;
            return nid;
        }

        // ---------------------------------------------------------------- replication columns of the table rows
        // Outlook regenerates these when they are missing ("minor inconsistencies"): a GUID per object (PidTagReplItemid), a change
        // number that grows, the version history of the store, and a row version that grows.
        long _changeNum = 0x400;
        uint _rowVer = 8;
        byte[] _replHistory = new byte[24];

        uint NextRowVer() { _rowVer += 2; return _rowVer; }

        List<PstProp> RowProps(List<PstProp> props, byte[] replId)
        {
            var l = new List<PstProp>(props);
            if (Fault("contents-cell-stale")) { l.RemoveAll(x => x.Id == 0x0037); l.Add(PstProp.Str(0x0037, "stale subject")); }
            l.Add(PstProp.Bin(0x0E30, replId));
            _changeNum += 8;
            l.Add(PstProp.Int64(0x0E33, _changeNum));
            l.Add(PstProp.Bin(0x0E34, _replHistory));
            return l;
        }

        // ---------------------------------------------------------------- Close
        // Node 0xC01 (internal, heap client signature 0x9C): a BTH of 16-byte keys to message NIDs that Outlook creates when it is
        // missing ("HMP not found - creating"). Its key is not documented; one random GUID per message is used.
        PstNode BuildMessageMap()
        {
            var recs = new List<byte[]>();
            _hmp.Sort((a, b) => { for (int i = 0; i < 16; i++) { int c = a.Key[i].CompareTo(b.Key[i]); if (c != 0) return c; } return 0; });
            if (Fault("hmp-wrong-entry") && _hmp.Count > 0) _hmp[0] = new KeyValuePair<byte[], uint>(_hmp[0].Key, PstFormat.NidRootFolder);
            foreach (KeyValuePair<byte[], uint> kv in _hmp) { var r = new byte[20]; Buffer.BlockCopy(kv.Key, 0, r, 0, 16); PstFormat.P32(r, 16, kv.Value); recs.Add(r); }
            var heap = new PstHeap(0x9C);
            uint root = heap.Reserve(4);
            uint hdr = PstBth.Build(heap, 16, 4, recs);
            heap.Set(root, BitConverter.GetBytes(hdr));
            return new PstNode { Blocks = heap.Finish(root) };
        }

        static byte[] EntryId(byte[] recordKey, uint nid)
        {
            var e = new byte[24];
            Buffer.BlockCopy(recordKey, 0, e, 4, 16);
            PstFormat.P32(e, 20, nid);
            return e;
        }

        PstTable EmptyTable(uint[] cols) { return new PstTable(cols); }

        public void Close()
        {
            if (_closed) return;
            _closed = true;
            // template and search objects
            _ndb.WriteTopLevelNode(PstFormat.NidHierarchyTemplate, EmptyTable(Fault("template-column-missing") ? HierarchyCols.Where(c => c != 0x0E300102).ToArray() : HierarchyCols).Build(), 0);
            _ndb.WriteTopLevelNode(PstFormat.NidContentsTemplate, EmptyTable(ContentsCols).Build(), 0);
            _ndb.WriteTopLevelNode(PstFormat.NidAssocContentsTemplate, EmptyTable(FaiCols).Build(), 0);
            _ndb.WriteTopLevelNode(PstFormat.NidSearchContentsTemplate, EmptyTable(SearchContentsCols).Build(), 0);
            _ndb.WriteTopLevelNode(PstFormat.NidRecipientTable, EmptyTable(RecipientCols).Build(), 0);
            _ndb.WriteTopLevelNode(PstFormat.NidAttachmentTable, EmptyTable(AttachmentCols).Build(), 0);
            // receive folder table (message class -> folder), outgoing queue and the search templates Outlook expects
            var rf = new PstTable(new uint[] { 0x001A001F, 0x0FFF0102, 0x30080040, 0x66050003 });
            var rfRow = new List<PstProp>();
            rfRow.Add(PstProp.Str(0x001A, "")); rfRow.Add(PstProp.Int32(0x6605, (int)PstFormat.NidRootFolder));
            if (!Fault("no-receive-row")) rf.AddRow(1, rf.Pick(rfRow));
            _ndb.WriteTopLevelNode(PstFormat.Nid(PstFormat.NtReceiveFolderTable, 0x31), rf.Build(), 0);
            if (!Fault("no-outgoing-queue")) _ndb.WriteTopLevelNode(PstFormat.Nid(PstFormat.NtOutgoingQueueTable, 0x32), EmptyTable(OutgoingQueueCols).Build(), 0);
            if (!Fault("no-template-6b6")) _ndb.WriteTopLevelNode(0x6B6, EmptyTable(new uint[] { 0x0E330014, 0x0E370102, 0x0E380003 }).Build(), 0);
            _ndb.WriteTopLevelNode(0x6D7, EmptyTable(new uint[] { 0x001A001F, 0x0E300102, 0x0E310102, 0x0E330014, 0x0E340102, 0x0E380003, 0x0E3E0102 }).Build(), 0);
            _ndb.WriteTopLevelNode(0x6F8, EmptyTable(new uint[] { 0x0E330014, 0x30070040 }).Build(), 0);
            if (!Fault("no-ec1")) _ndb.AddNode(0xEC1, 0, 0, 0);
            foreach (PstFolder hf in _folders) if (hf.Parent != null && !hf.IsSearchFolder) _hmp.Add(new KeyValuePair<byte[], uint>(hf.ReplId, hf.Nid));
            if (!Fault("no-hmp")) _ndb.WriteTopLevelNode(0xC01, BuildMessageMap(), 0);
            _ndb.AddNode(PstFormat.NidSearchManagementQueue, 0, 0, 0);
            _ndb.AddNode(PstFormat.NidSearchActivityList, 0, 0, 0);

            // folders
            foreach (PstFolder f in _folders) WriteFolder(f);

            if (Fault("attach-highwater")) _ndb.Rgnid[PstFormat.NtAttachment] = 0x400;

            // name-to-id map
            _ndb.WriteTopLevelNode(PstFormat.NidNameToIdMap, Fault("namemap-empty") ? new PstNameMap().Build() : _names.Build(), 0);
            if (Fault("decoy-namemap"))     // a second, different map in the file (the old version a real file keeps in free space)
            {
                var decoy = new PstNameMap();
                decoy.Resolve(new Guid("deadbeef-0000-4000-8000-00000000d00d"), 0x1234, null);
                _ndb.WriteTopLevelNode(PstFormat.Nid(PstFormat.NtInternal, 0x777), decoy.Build(), 0);
            }

            // message store
            var sp = new List<PstProp>();
            // replication version history (a count, a GUID and a number, as in the sample of the specification), flags, and the two
            // properties Outlook writes in the store (6633 / 66FA, meaning not documented)
            sp.Add(PstProp.Bin(0x0E34, _replHistory));
            sp.Add(PstProp.Int32(0x0E38, 0));
            sp.Add(PstProp.Bool(0x6633, true));
            sp.Add(PstProp.Int32(0x66FA, 0x000E000D));
            sp.Add(PstProp.Bin(0x0FF9, _recordKey));
            sp.Add(PstProp.Str(0x3001, _displayName));
            sp.Add(PstProp.Int32(0x35DF, 0x89));
            sp.Add(PstProp.Bin(0x35E0, EntryId(_recordKey, IpmSubtree.Nid)));
            sp.Add(PstProp.Bin(0x35E3, EntryId(_recordKey, DeletedItems.Nid)));
            sp.Add(PstProp.Bin(0x35E7, EntryId(_recordKey, SearchRoot.Nid)));
            sp.Add(PstProp.Int32(0x67FF, 0));
            _ndb.WriteTopLevelNode(PstFormat.NidMessageStore, PstPc.Build(sp), 0);
            _ndb.Close();
        }

        void WriteFolder(PstFolder f)
        {
            uint idx = PstFormat.NidIndex(f.Nid);
            // property context
            var props = new List<PstProp>();
            props.Add(PstProp.Str(0x3001, f.Name ?? ""));
            props.Add(PstProp.Int32(0x3602, f.ContentCount + (Fault("wrong-count") && f.ContentCount > 0 ? 5 : 0)));
            props.Add(PstProp.Int32(0x3603, f.UnreadCount));
            props.Add(PstProp.Bool(0x360A, f.Children.Count > 0));
            if (!string.IsNullOrEmpty(f.ContainerClass)) props.Add(PstProp.Str(0x3613, f.ContainerClass));
            if (f.FaiCount > 0) props.Add(PstProp.Int32(0x3617, f.FaiCount));
            if (Fault("hidden-count-wrong")) props.Add(PstProp.Int32(0x6635, f.FaiCount + 3));
            uint parentNid = f.Parent == null ? f.Nid : f.Parent.Nid;
            _ndb.WriteTopLevelNode(f.Nid, PstPc.Build(props), parentNid);

            // hierarchy table: one row per child
            var ht = new PstTable(HierarchyCols);
            bool skipFirst = Fault("hierarchy-missing");
            foreach (PstFolder c in f.Children)
            {
                if (skipFirst) { skipFirst = false; continue; }
                var cp = new List<PstProp>();
                cp.Add(PstProp.Str(0x3001, c.Name ?? ""));
                cp.Add(PstProp.Int32(0x3602, c.ContentCount + (Fault("hierarchy-row-value") ? 1 : 0)));
                cp.Add(PstProp.Int32(0x3603, c.UnreadCount));
                cp.Add(PstProp.Bool(0x360A, c.Children.Count > 0));
                if (!string.IsNullOrEmpty(c.ContainerClass)) cp.Add(PstProp.Str(0x3613, c.ContainerClass));
                ht.AddRow(c.Nid, ht.Pick(RowProps(cp, c.ReplId)), NextRowVer());
            }
            _ndb.WriteTopLevelNode(PstFormat.Nid(PstFormat.NtHierarchyTable, idx), ht.Build(), 0);
            if (f.IsSearchFolder)
                _ndb.WriteTopLevelNode(PstFormat.Nid(PstFormat.NtSearchContentsTable, idx), EmptyTable(SearchContentsCols).Build(), 0);
            else
            {
                if (Fault("contents-row-removed")) f.Contents.RemoveLastRow();
                if (Fault("dangling-nid") && f.ContentCount > 0) f.Contents.AddRow(PstFormat.Nid(PstFormat.NtNormalMessage, 0x7FFF000), new PstProp[f.Contents.Columns.Count]);
                _ndb.WriteTopLevelNode(PstFormat.Nid(PstFormat.NtContentsTable, idx), f.Contents.Build(), 0);
            }
            _ndb.WriteTopLevelNode(PstFormat.Nid(PstFormat.NtAssocContentsTable, idx), f.Fai.Build(), 0);
        }

        public void Dispose()
        {
            if (!_closed) { _closed = true; _ndb.Dispose(); }
        }
    }

    /// <summary>
    /// A PST split by size: files base.pst, base-0002.pst, ... each at most maxBytes (50 GiB = 51,200 MiB by default, the default
    /// MaxLargeFileSize of Outlook), the folder tree being recreated in each file. A message that alone exceeds the limit is written
    /// alone in its own file. The size is the projected one (PstWriter.ProjectedSize), which does not underestimate the file.
    /// </summary>
    public sealed class PstSplitWriter : IDisposable
    {
        public const long DefaultMaxBytes = 51200L * 1024 * 1024;
        readonly string _dir, _stem, _displayName;
        readonly long _max;
        PstWriter _cur;
        Dictionary<string, PstFolder> _folders = new Dictionary<string, PstFolder>(StringComparer.Ordinal);
        public readonly List<string> Files = new List<string>();
        public long MessagesWritten;

        public PstSplitWriter(string path, string displayName, long maxBytes)
        {
            _dir = System.IO.Path.GetDirectoryName(System.IO.Path.GetFullPath(path));
            _stem = System.IO.Path.GetFileNameWithoutExtension(path);
            _displayName = displayName;
            _max = maxBytes > 0 ? maxBytes : DefaultMaxBytes;
        }

        public PstWriter Current { get { return _cur; } }

        void Next()
        {
            if (_cur != null) { _cur.Close(); _cur = null; }
            string name = Files.Count == 0 ? _stem + ".pst" : _stem + "-" + (Files.Count + 1).ToString("D4") + ".pst";
            string full = System.IO.Path.Combine(_dir, name);
            _cur = PstWriter.Create(full, _displayName);
            Files.Add(full);
            _folders = new Dictionary<string, PstFolder>(StringComparer.Ordinal);
        }

        PstFolder Folder(string path, string cls)
        {
            string key = path ?? "";
            PstFolder f;
            if (_folders.TryGetValue(key, out f)) return f;
            PstFolder parent = null;
            string acc = "";
            foreach (string part in key.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
            {
                acc = acc.Length == 0 ? part : acc + "/" + part;
                if (!_folders.TryGetValue(acc, out f))
                {
                    f = _cur.CreateFolder(parent, part, acc == key ? cls : "IPF.Note");
                    _folders[acc] = f;
                }
                parent = f;
            }
            if (parent == null) parent = _cur.IpmSubtree;
            _folders[key] = parent;
            return parent;
        }

        /// <summary>Adds a message to the folder path ("Mail/2024/Invoices"), rolling to a new file when the limit would be exceeded.</summary>
        public uint AddMessage(string folderPath, string containerClass, FxMessage message)
        {
            return AddMessage(folderPath, containerClass, message, System.Threading.CancellationToken.None, null);
        }

        /// <summary>Same, cancellable and reporting progress (see PstWriter.AddMessage).</summary>
        public uint AddMessage(string folderPath, string containerClass, FxMessage message, System.Threading.CancellationToken ct, IProgress<PstProgress> progress)
        {
            ct.ThrowIfCancellationRequested();
            long est = PstWriter.EstimateSize(message);
            if (_cur == null) Next();
            else if (_cur.MessagesWritten > 0 && _cur.ProjectedSize + est > _max) Next();
            uint nid = _cur.AddMessage(Folder(folderPath, containerClass), message, ct, progress);
            MessagesWritten++;
            return nid;
        }

        /// <summary>The file that received the last message and the file list; call Close first to have them complete.</summary>
        public void Close()
        {
            if (_cur != null) { _cur.Close(); _cur = null; }
        }

        public void Dispose() { if (_cur != null) { _cur.Dispose(); _cur = null; } }
    }
}
