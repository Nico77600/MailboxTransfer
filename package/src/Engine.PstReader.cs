// =============================================================================
//  Mailbox Transfer - engine, part 5e: MS-PST reader, Messaging layer
// -----------------------------------------------------------------------------
//  PstReader.Open(path)                       the message store, the folder tree, the messages of a folder
//  PstReader.ReadMessage(nid)                 a message as an FxMessage (properties, recipients, attachments, embedded messages):
//                                             FxCodec.EncodePayload turns it into a FastTransfer stream (import from PST)
//  Named properties come back with their GUID and LID or name (the PST ids 0x8000+ are local to the file).
//  Values the PST computes and does not store are not rebuilt, except PidTagAttachNumber (the position of the attachment).
// =============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public sealed class PstReaderFolder
    {
        public uint Nid, ParentNid;
        public string Name, ContainerClass, Path;
        public int ContentCount, UnreadCount, AssocCount;
        public bool HasSubfolders;
        public readonly List<PstReaderFolder> Children = new List<PstReaderFolder>();
    }

    public sealed partial class PstReader : IDisposable
    {
        public readonly PstFile File;
        /// <summary>
        /// A single value longer than this stays in the PST: the property of ReadMessage is lazy (FxProperty.Lazy, read block by block
        /// when it is used). Default 64 KiB; long.MaxValue reads every value into memory.
        /// </summary>
        public long LazyThreshold = 64 * 1024;
        public string DisplayName;
        public byte[] RecordKey;
        Dictionary<ushort, PstNameMap.Entry> _names;
        PstReaderFolder _root;

        PstReader(PstFile f) { File = f; }

        public static PstReader Open(string path)
        {
            PstFile f = PstFile.Open(path);
            try
            {
                var r = new PstReader(f);
                var store = new PstPcReader(f.OpenNode(PstFormat.NidMessageStore));
                r.DisplayName = PstConvert.StringOf(store.Get(0x3001));
                PstProp rk = store.Get(0x0FF9);
                r.RecordKey = rk == null ? null : rk.Data;
                return r;
            }
            catch (Exception) { f.Dispose(); throw; }
        }

        public void Dispose() { File.Dispose(); }

        // ---------------------------------------------------------------- folders
        public PstReaderFolder Root
        {
            get
            {
                if (_root == null) _root = LoadFolder(PstFormat.NidRootFolder, null, "", 0);
                return _root;
            }
        }

        PstReaderFolder LoadFolder(uint nid, PstReaderFolder parent, string path, int depth)
        {
            if (depth > 64) throw new PstFormatException("Folder hierarchy deeper than 64 levels");
            PstNbtEntry nbt;
            if (!File.TryFindNode(nid, out nbt)) throw new PstFormatException("Folder node 0x" + nid.ToString("X") + " does not exist") { Nid = nid };
            var pc = new PstPcReader(new PstNodeView(File, nbt.Nid, nbt.BidData, nbt.BidSub));
            var f = new PstReaderFolder { Nid = nid, ParentNid = nbt.Parent };
            f.Name = PstConvert.StringOf(pc.Get(0x3001)) ?? "";
            f.ContainerClass = PstConvert.StringOf(pc.Get(0x3613));
            f.ContentCount = PstConvert.Int32Of(pc.Get(0x3602), 0);
            f.UnreadCount = PstConvert.Int32Of(pc.Get(0x3603), 0);
            f.AssocCount = PstConvert.Int32Of(pc.Get(0x3617), 0);
            PstProp sub = pc.Get(0x360A);
            f.HasSubfolders = sub != null && sub.Data != null && sub.Data.Length > 0 && sub.Data[0] != 0;
            f.Path = parent == null ? "" : (path.Length == 0 ? f.Name : path + "/" + f.Name);
            uint idx = PstFormat.NidIndex(nid);
            uint hnid = PstFormat.Nid(PstFormat.NtHierarchyTable, idx);
            if (File.HasNode(hnid))
            {
                var tc = new PstTcReader(File.OpenNode(hnid));
                for (int i = 0; i < tc.RowCount; i++)
                {
                    uint child = tc.RowId(i);
                    if (child == nid) throw new PstFormatException("Folder 0x" + nid.ToString("X") + " lists itself as a subfolder");
                    f.Children.Add(LoadFolder(child, f, f.Path, depth + 1));
                }
            }
            return f;
        }

        public IEnumerable<PstReaderFolder> AllFolders()
        {
            var st = new Stack<PstReaderFolder>();
            st.Push(Root);
            while (st.Count > 0)
            {
                PstReaderFolder f = st.Pop();
                yield return f;
                for (int i = f.Children.Count - 1; i >= 0; i--) st.Push(f.Children[i]);
            }
        }

        /// <summary>NIDs of the messages of a folder, from its contents table (or its FAI contents table).</summary>
        public List<uint> MessageNids(PstReaderFolder folder, bool associated)
        {
            uint tn = PstFormat.Nid(associated ? PstFormat.NtAssocContentsTable : PstFormat.NtContentsTable, PstFormat.NidIndex(folder.Nid));
            var res = new List<uint>();
            if (!File.HasNode(tn)) return res;
            var tc = new PstTcReader(File.OpenNode(tn));
            for (int i = 0; i < tc.RowCount; i++) res.Add(tc.RowId(i));
            return res;
        }

        // ---------------------------------------------------------------- named properties
        /// <summary>Repair: a damaged name-to-id map is read entry by entry; the entries that cannot be read are listed here instead of raising.</summary>
        public List<string> NameMapErrors;

        Dictionary<ushort, PstNameMap.Entry> Names
        {
            get
            {
                if (_names != null) return _names;
                var names = new Dictionary<ushort, PstNameMap.Entry>();
                _names = names;
                bool tol = NameMapErrors != null;
                try
                {
                    if (!File.HasNode(PstFormat.NidNameToIdMap)) return names;
                    names = PstNameMapReader.Read(new PstPcReader(File.OpenNode(PstFormat.NidNameToIdMap)), tol ? NameMapErrors : null, null);
                    _names = names;
                }
                catch (PstFormatException x) { if (!tol) { _names = null; throw; } NameMapErrors.Add(x.Message); }
                return names;
            }
        }

        /// <summary>Repair: the name-to-id map to use instead of node 0x61 (found elsewhere in the file, see PstRepairRun.ResolveNameMap).</summary>
        internal void SetNameMap(Dictionary<ushort, PstNameMap.Entry> names) { _names = names; }

        /// <summary>Repair: the named-property ids the properties read so far use, the ones the map does not resolve, and the number of properties dropped for that.</summary>
        internal readonly HashSet<ushort> NamedIdsUsed = new HashSet<ushort>(), NamedIdsUnresolved = new HashSet<ushort>();
        internal long NamedPropsDropped;
        public PstNameMap.Entry NameOf(ushort id) { PstNameMap.Entry e; return Names.TryGetValue(id, out e) ? e : null; }

        FxProperty ToFxNamed(PstProp p)
        {
            FxProperty f = PstConvert.ToFx(p);
            if (p.Id >= 0x8000)
            {
                PstNameMap.Entry e;
                if (Names.TryGetValue(p.Id, out e))
                {
                    f.NamedGuid = e.Set;
                    if (e.Lid.HasValue) f.NamedLid = e.Lid; else f.NamedName = e.Name;
                }
            }
            return f;
        }

        // ---------------------------------------------------------------- messages
        public FxMessage ReadMessage(uint nid)
        {
            return ReadMessage(File.OpenNode(nid), 0);
        }

        /// <summary>
        /// The FastTransfer stream of a message, produced on demand (Length known, values copied from the PST while it is read), bare or
        /// inside an envelope: what ImportContent / GraphClient.ImportItem send. The reader must stay open while the stream is read.
        /// </summary>
        public FxMessageSource OpenMessageSource(uint nid, FxEnvelope envelope)
        {
            return new FxMessageSource(ReadMessage(nid), envelope);
        }

        public FxMessageSource OpenMessageSource(uint nid) { return OpenMessageSource(nid, null); }

        /// <summary>
        /// Writes the FastTransfer stream of a message to a stream (a file, a network stream) without holding it: cancellable between
        /// two pieces, progress reported about every MiB (BytesDone / BytesTotal). Returns the number of bytes written.
        /// </summary>
        public long CopyMessageTo(uint nid, FxEnvelope envelope, System.IO.Stream dest, System.Threading.CancellationToken ct, IProgress<PstProgress> progress)
        {
            return FxStreamSource.CopyTo(OpenMessageSource(nid, envelope), dest, ct, progress);
        }

        FxMessage ReadMessage(PstNodeView node, int depth)
        {
            if (depth > 32) throw new PstFormatException("Embedded messages nested deeper than 32 levels");
            var m = new FxMessage();
            var pc = new PstPcReader(node) { LazyThreshold = LazyThreshold };
            foreach (PstProp p in pc.All()) m.Properties.Add(ToFxNamed(p));
            if (node.HasSubnode(PstFormat.NidRecipientTable))
            {
                var tc = new PstTcReader(node.OpenSubnode(PstFormat.NidRecipientTable));
                for (int i = 0; i < tc.RowCount; i++)
                {
                    var r = new FxRecipient();
                    foreach (PstProp p in tc.GetRow(i)) r.Properties.Add(ToFxNamed(p));
                    m.Recipients.Add(r);
                }
            }
            if (node.HasSubnode(PstFormat.NidAttachmentTable))
            {
                var tc = new PstTcReader(node.OpenSubnode(PstFormat.NidAttachmentTable));
                for (int i = 0; i < tc.RowCount; i++)
                {
                    uint anid = tc.RowId(i);
                    PstNodeView an = node.OpenSubnode(anid);
                    var apc = new PstPcReader(an) { LazyThreshold = LazyThreshold };
                    var a = new FxAttachment();
                    a.Properties.Add(FxProperty.FromInt32(0x0E21, i));
                    foreach (PstProp p in apc.All())
                    {
                        if (p.ObjectView != null)
                        {
                            a.Embedded = ReadMessage(p.ObjectView, depth + 1);
                            a.Properties.Add(FxProperty.Create(0x3701, 0x000D, new byte[0]));
                        }
                        else a.Properties.Add(ToFxNamed(p));
                    }
                    m.Attachments.Add(a);
                }
            }
            return m;
        }
    }
}
