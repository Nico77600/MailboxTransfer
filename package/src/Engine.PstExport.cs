// =============================================================================
//  Mailbox Transfer - engine, part 5f: export store -> PST files (the real folder tree, split by size, resumable)
// -----------------------------------------------------------------------------
//    PstExportOptions     what to convert (mailboxes, selection), where, how big a PST may be, merge of the auxiliary archives
//    PstExporter.Plan     read only: locations, folders, item counts, announced sizes, estimated number of PST files
//    PstExporter.Run      writes <name>.pst, <name>-002.pst, ... and <name>.manifest.json; resumes an interrupted run
//    PstExportSelection   a browse selection (BrowseSelection) turned into folder keys / item keys of a store
//    PstWellKnown         the standard folders of a mailbox (Inbox, Sent Items, Calendar...) recognised by name, for their container class and order
//
//  The engine writes nothing to the console: progress goes through IProgress<PstExportProgress>, cancellation through a
//  CancellationToken, the outcome is a PstExportResult. The memory used does not depend on the size of the items (FxCodec keeps the
//  values above 64 KiB in the pack of the store and PstWriter streams them); it grows only with the folder tree and with the item
//  records of the manifest (about 150 bytes per item).
//
//  Layout of the PST: IPM subtree > one top folder per location (Primary, Archive, Auxiliary archive n (guid8)) > the folders of the
//  mailbox, with their names, their hierarchy and their container class. The auxiliary archives of an auto-expanding archive are
//  merged into the Archive tree by default: a folder of an auxiliary archive is the same logical folder as the one of the main
//  archive that has the same folder entry key (ExchangeEntryId.FolderKey); only a folder whose id is not in the Exchange format is
//  matched by its path.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;

namespace MailboxTransfer
{
    public sealed class PstExportOptions
    {
        /// <summary>The first PST file (D:\out\mail.pst); the next ones are mail-002.pst, mail-003.pst... and the manifest mail.manifest.json.</summary>
        public string Destination;
        /// <summary>Largest PST in bytes (default 50 GiB = 51,200 MiB). A single item larger than this is written alone in its own PST and reported.</summary>
        public long MaxPartBytes = PstSplitWriter.DefaultMaxBytes;
        /// <summary>Merge the folders of the auxiliary archives into the Archive tree (default), else one top folder per auxiliary archive.</summary>
        public bool MergeAuxIntoArchive = true;
        /// <summary>Only the mailboxes of these kinds (Primary, Archive, AuxArchive); null: all. With the merge on, Archive also takes the auxiliary archives.</summary>
        public StoreMailboxKind[] Kinds;
        /// <summary>Only these mailboxes (ids of the store); null: all.</summary>
        public string[] MailboxIds;
        /// <summary>Only these folders (store folder keys, descendants already included by the caller); null with ItemKeys null: everything.</summary>
        public HashSet<long> FolderKeys;
        /// <summary>Only these items (store item keys), in addition to the folders of FolderKeys.</summary>
        public HashSet<long> ItemKeys;
        /// <summary>Run the integrity checker on each finished PST.</summary>
        public bool Check = true;
        /// <summary>Start again from zero: delete the manifest and the PST files of an earlier run.</summary>
        public bool Force;
        /// <summary>Display name of the PST store (default: "MailboxTransfer" and the user of the store).</summary>
        public string DisplayName;
        public string PrimaryName = "Primary";
        public string ArchiveName = "Archive";
        /// <summary>{0} = number (1, 2...), {1} = the first 8 characters of the mailbox guid.</summary>
        public string AuxArchiveNameFormat = "Auxiliary archive {0} ({1})";
        /// <summary>
        /// The name of that folder in versions 1.0.x (French). Only used to recognise it: an export of 1.0.x that is continued keeps writing the folder under the
        /// name its first parts already have, so that the parts of one export agree and no second folder appears. A new export never writes it.
        /// </summary>
        public const string LegacyAuxArchiveNameFormat = "Archive auxiliaire {0} ({1})";
        /// <summary>Items read from the catalog at a time.</summary>
        public int PageSize = 1000;
        /// <summary>Told how many items each folder of the PST has to get and each item written, failed or kept from an earlier run (progress per folder for the console); null: not followed.</summary>
        public FolderTracker Folders;
    }

    public sealed class PstExportLocation
    {
        public int Index;
        public string Name;
        /// <summary>Primary, Archive or AuxArchive (the kind of the main mailbox of the location).</summary>
        public string Kind;
        public List<string> MailboxIds = new List<string>();
        public int Folders;
        public long Items, AnnouncedBytes, StoredBytes;
        internal HashSet<string> Members = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        internal PstExportFolder Top;
        internal Dictionary<string, PstExportFolder> ByKey = new Dictionary<string, PstExportFolder>(StringComparer.Ordinal);
    }

    public sealed class PstExportFolder
    {
        /// <summary>Order of creation: the top folder of a location, then its folders depth first (standard folders first, then by name).</summary>
        public int Index;
        public int Parent = -1;
        public int Location;
        public int Depth;
        public string Name;
        /// <summary>Path in the PST below the IPM subtree: "Archive/Inbox/Reports".</summary>
        public string Path;
        public string ContainerClass;
        /// <summary>Inbox, Sent Items, Drafts, Deleted Items, Outbox, Junk Email, Calendar, Contacts, Tasks, Notes or Journal for a standard folder at the top of a location; else null.</summary>
        public string WellKnown;
        public bool IsTop;
        /// <summary>False: not in the selection (and no selected folder below it); it is not created.</summary>
        public bool Included = true;
        /// <summary>Items of this folder (not of its subfolders) that are Exported and selected, their announced size and the bytes stored.</summary>
        public long Items, AnnouncedBytes, StoredBytes;
        public List<long> SourceFolderKeys = new List<long>();
        internal List<PstExportFolder> Children = new List<PstExportFolder>();
        internal int Rank = 100;
        internal bool Explicit;
    }

    public sealed class PstExportPlan
    {
        public long MaxPartBytes;
        public bool MergeAuxIntoArchive;
        public List<PstExportLocation> Locations = new List<PstExportLocation>();
        public List<PstExportFolder> Folders = new List<PstExportFolder>();
        public long Items, AnnouncedBytes, StoredBytes;
        /// <summary>Estimate from the stored sizes (the PST is about as big as the FastTransfer streams, plus a few KiB per item and folder).</summary>
        public int EstimatedParts;
        /// <summary>Items that are in the store but not Exported (Pending, Failed, Skipped): they are not in the PST.</summary>
        public long ItemsNotExported;
        public List<string> Notes = new List<string>();
        /// <summary>Items (store keys) whose estimated size alone exceeds the limit.</summary>
        public List<long> OversizeItems = new List<long>();
        public IEnumerable<PstExportFolder> IncludedFolders { get { return Folders.Where(f => f.Included); } }
    }

    public sealed class PstExportProgress
    {
        /// <summary>plan, write, check, manifest or done.</summary>
        public string Phase;
        public int PartIndex;
        public string PartPath;
        public string Folder;
        public long ItemsDone, ItemsTotal;
        /// <summary>Stored bytes of the items processed (the items of the parts reused included) and of the whole export.</summary>
        public long BytesDone, BytesTotal;
        /// <summary>Projected size of the PST being written.</summary>
        public long PartBytes;
        public string Message;
    }

    public sealed class PstExportProgressSink : IProgress<PstExportProgress>
    {
        readonly Action<PstExportProgress> _handler;
        public PstExportProgressSink(Action<PstExportProgress> handler) { _handler = handler; }
        public void Report(PstExportProgress value) { if (_handler != null) _handler(value); }
    }

    public sealed class PstExportItemRecord
    {
        public long Key { get; set; }
        public int Folder { get; set; }
        public long Length { get; set; }
        public string Sha256 { get; set; }
    }

    public sealed class PstExportProblem
    {
        public long ItemKey { get; set; }
        public string Subject { get; set; }
        public string Folder { get; set; }
        public string Reason { get; set; }
    }

    public sealed class PstExportOversize
    {
        public long ItemKey { get; set; }
        public string Subject { get; set; }
        public long EstimatedBytes { get; set; }
        public int Part { get; set; }
    }

    public sealed class PstExportPartInfo
    {
        public int Index { get; set; }
        public string File { get; set; }
        public long Bytes { get; set; }
        public int ItemCount { get; set; }
        public int FolderCount { get; set; }
        /// <summary>Bytes of the FastTransfer streams of the items.</summary>
        public long SourceBytes { get; set; }
        /// <summary>The part holds an item larger than the limit.</summary>
        public bool Oversize { get; set; }
        public bool Checked { get; set; }
        public int CheckErrors { get; set; }
        public int CheckWarnings { get; set; }
        /// <summary>No error found by the checker (or the check was not asked for).</summary>
        public bool Verified { get; set; }
        public DateTime CompletedUtc { get; set; }
        public List<PstExportItemRecord> Items { get; set; } = new List<PstExportItemRecord>();
        /// <summary>This run reused the part (it was complete and its header matched the manifest).</summary>
        [System.Text.Json.Serialization.JsonIgnore] public bool Reused { get; set; }
        [System.Text.Json.Serialization.JsonIgnore] public string Path { get; set; }
        [System.Text.Json.Serialization.JsonIgnore] public double Seconds { get; set; }
        [System.Text.Json.Serialization.JsonIgnore] public double CheckSeconds { get; set; }
    }

    public sealed class PstExportFolderInfo
    {
        public int Index { get; set; }
        public int Parent { get; set; }
        public string Path { get; set; }
        public string ContainerClass { get; set; }
        public string WellKnown { get; set; }
        public long Items { get; set; }
        public long AnnouncedBytes { get; set; }
        public long StoredBytes { get; set; }
    }

    public sealed class PstExportLocationInfo
    {
        public string Name { get; set; }
        public string Kind { get; set; }
        public List<string> MailboxIds { get; set; } = new List<string>();
        public long Items { get; set; }
        public long AnnouncedBytes { get; set; }
        public long StoredBytes { get; set; }
    }

    /// <summary>The manifest of an export to PST (JSON): the parts, the folders and, for each part, the items (store key, folder, length and SHA-256 of the source stream).</summary>
    public sealed class PstExportManifest
    {
        public int Version { get; set; } = 1;
        public string Tool { get; set; } = "MailboxTransfer";
        public string Store { get; set; }
        public DateTime CreatedUtc { get; set; }
        public DateTime UpdatedUtc { get; set; }
        /// <summary>True when every exported item of the selection is in a part.</summary>
        public bool Complete { get; set; }
        public long MaxPartBytes { get; set; }
        public bool MergeAuxIntoArchive { get; set; }
        public string DisplayName { get; set; }
        public long PlannedItems { get; set; }
        public List<PstExportLocationInfo> Locations { get; set; } = new List<PstExportLocationInfo>();
        public List<PstExportFolderInfo> Folders { get; set; } = new List<PstExportFolderInfo>();
        public List<PstExportPartInfo> Parts { get; set; } = new List<PstExportPartInfo>();
        public List<PstExportProblem> Failed { get; set; } = new List<PstExportProblem>();
        public List<PstExportOversize> Oversize { get; set; } = new List<PstExportOversize>();

        public static string PathFor(string destination)
        {
            string full = System.IO.Path.GetFullPath(destination);
            return System.IO.Path.Combine(System.IO.Path.GetDirectoryName(full), System.IO.Path.GetFileNameWithoutExtension(full) + ".manifest.json");
        }

        public static PstExportManifest Load(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 16, FileOptions.SequentialScan))
            {
                var m = JsonSerializer.Deserialize<PstExportManifest>(fs, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                if (m == null) throw new InvalidDataException("The manifest is empty.");
                if (m.Version != 1) throw new InvalidDataException("Manifest version " + m.Version + " is not supported (this build reads version 1).");
                return m;
            }
        }

        /// <summary>Written to a temporary file, then moved over the manifest: an interrupted write never leaves a half manifest.</summary>
        public void Save(string path)
        {
            UpdatedUtc = DateTime.UtcNow;
            string tmp = path + ".tmp";
            using (var fs = new FileStream(tmp, FileMode.Create, FileAccess.Write, FileShare.None, 1 << 16))
            using (var w = new Utf8JsonWriter(fs, new JsonWriterOptions { Indented = true }))
            {
                w.WriteStartObject();
                w.WriteNumber("Version", Version);
                w.WriteString("Tool", Tool);
                w.WriteString("Store", Store);
                w.WriteString("CreatedUtc", CreatedUtc);
                w.WriteString("UpdatedUtc", UpdatedUtc);
                w.WriteBoolean("Complete", Complete);
                w.WriteNumber("MaxPartBytes", MaxPartBytes);
                w.WriteBoolean("MergeAuxIntoArchive", MergeAuxIntoArchive);
                w.WriteString("DisplayName", DisplayName);
                w.WriteNumber("PlannedItems", PlannedItems);
                w.WritePropertyName("Locations"); JsonSerializer.Serialize(w, Locations);
                w.WritePropertyName("Folders"); JsonSerializer.Serialize(w, Folders);
                w.WritePropertyName("Parts");
                w.WriteStartArray();
                foreach (PstExportPartInfo p in Parts)
                {
                    w.WriteStartObject();
                    w.WriteNumber("Index", p.Index);
                    w.WriteString("File", p.File);
                    w.WriteNumber("Bytes", p.Bytes);
                    w.WriteNumber("ItemCount", p.ItemCount);
                    w.WriteNumber("FolderCount", p.FolderCount);
                    w.WriteNumber("SourceBytes", p.SourceBytes);
                    w.WriteBoolean("Oversize", p.Oversize);
                    w.WriteBoolean("Checked", p.Checked);
                    w.WriteNumber("CheckErrors", p.CheckErrors);
                    w.WriteNumber("CheckWarnings", p.CheckWarnings);
                    w.WriteBoolean("Verified", p.Verified);
                    w.WriteString("CompletedUtc", p.CompletedUtc);
                    w.WritePropertyName("Items");
                    w.WriteStartArray();
                    foreach (PstExportItemRecord r in p.Items)
                        // numbers and hex only: nothing to escape, one compact line per item
                        w.WriteRawValue("{\"Key\":" + r.Key + ",\"Folder\":" + r.Folder + ",\"Length\":" + r.Length + ",\"Sha256\":\"" + (r.Sha256 ?? "") + "\"}", true);
                    w.WriteEndArray();
                    w.WriteEndObject();
                }
                w.WriteEndArray();
                w.WritePropertyName("Failed"); JsonSerializer.Serialize(w, Failed);
                w.WritePropertyName("Oversize"); JsonSerializer.Serialize(w, Oversize);
                w.WriteEndObject();
            }
            if (File.Exists(path)) File.Replace(tmp, path, null); else File.Move(tmp, path);
        }
    }

    public sealed class PstExportResult
    {
        public PstExportPlan Plan;
        public string ManifestPath;
        public List<PstExportPartInfo> Parts = new List<PstExportPartInfo>();
        /// <summary>Items written by this run, and items already in the parts reused.</summary>
        public long ItemsWritten, ItemsReused;
        public long BytesWritten;
        public int PartsWritten, PartsReused, PartsRebuilt;
        public List<PstExportProblem> Failed = new List<PstExportProblem>();
        public List<PstExportOversize> Oversize = new List<PstExportOversize>();
        public List<string> Notes = new List<string>();
        /// <summary>Every selected item is in a part and no item failed.</summary>
        public bool Complete;
        public bool Cancelled;
        /// <summary>Parts whose check found errors.</summary>
        public int PartsWithErrors;
        public string Error;
        public TimeSpan Elapsed;
        public bool Ok { get { return Complete && !Cancelled && Error == null && PartsWithErrors == 0; } }
    }

    // ---------------------------------------------------------------------------------------------
    //  Standard folders
    // ---------------------------------------------------------------------------------------------

    public static class PstWellKnown
    {
        sealed class Def { public string Key; public int Rank; public string Class; }
        static readonly Dictionary<string, Def> ByName = new Dictionary<string, Def>(StringComparer.OrdinalIgnoreCase);

        static void Add(string key, int rank, string cls, params string[] names)
        {
            var d = new Def { Key = key, Rank = rank, Class = cls };
            ByName[key] = d;
            foreach (string n in names) ByName[n] = d;
        }

        static PstWellKnown()
        {
            Add("Inbox", 1, "IPF.Note", "Bo\u00EEte de r\u00E9ception", "Posteingang", "Bandeja de entrada", "Posta in arrivo", "Postvak IN");
            Add("Drafts", 2, "IPF.Note", "Brouillons", "Entw\u00FCrfe", "Borradores", "Bozze", "Concepten");
            Add("Sent Items", 3, "IPF.Note", "\u00C9l\u00E9ments envoy\u00E9s", "Gesendete Elemente", "Elementos enviados", "Posta inviata", "Verzonden items");
            Add("Deleted Items", 4, "IPF.Note", "\u00C9l\u00E9ments supprim\u00E9s", "Gel\u00F6schte Elemente", "Elementos eliminados", "Posta eliminata", "Verwijderde items");
            Add("Outbox", 5, "IPF.Note", "Bo\u00EEte d'envoi", "Postausgang", "Bandeja de salida", "Posta in uscita", "Postvak UIT");
            Add("Junk Email", 6, "IPF.Note", "Courrier ind\u00E9sirable", "Junk-E-Mail", "Correo no deseado", "Posta indesiderata", "Ongewenste e-mail");
            Add("Calendar", 7, "IPF.Appointment", "Calendrier", "Kalender", "Calendario", "Agenda");
            Add("Contacts", 8, "IPF.Contact", "Contactos", "Kontakte", "Contatti", "Contactpersonen");
            Add("Tasks", 9, "IPF.Task", "T\u00E2ches", "Aufgaben", "Tareas", "Attivit\u00E0", "Taken");
            Add("Notes", 10, "IPF.StickyNote", "Notizen", "Notas", "Note", "Notities");
            Add("Journal", 11, "IPF.Journal", "Journal", "Diario", "Dagboek");
        }

        /// <summary>The standard folder a top-level folder name stands for (English names and the usual translations), else null.</summary>
        public static string Detect(string name, out int rank, out string containerClass)
        {
            rank = 100; containerClass = null;
            Def d;
            if (name != null && ByName.TryGetValue(name.Trim(), out d)) { rank = d.Rank; containerClass = d.Class; return d.Key; }
            return null;
        }

        /// <summary>The container class of a folder: the type the source gave, else the class of the standard folder, else IPF.Note.</summary>
        public static string ContainerClass(string folderType, string wellKnownClass)
        {
            if (!string.IsNullOrEmpty(folderType) && folderType.StartsWith("IPF", StringComparison.OrdinalIgnoreCase)) return folderType;
            return string.IsNullOrEmpty(wellKnownClass) ? "IPF.Note" : wellKnownClass;
        }
    }

    // ---------------------------------------------------------------------------------------------
    //  A browse selection -> keys of the store
    // ---------------------------------------------------------------------------------------------

    public sealed class PstExportSelection
    {
        public HashSet<long> FolderKeys = new HashSet<long>();
        public HashSet<long> ItemKeys = new HashSet<long>();
        public List<string> Problems = new List<string>();
        public int EntriesMatched;

        /// <summary>
        /// Looks each entry of a selection (written by the browse view) up in the store: a folder by its id, else by its folder entry key (the same in
        /// the main archive and in the auxiliary archives), else by its path in a mailbox of the same kind; its subfolders by path when asked; the
        /// selected items by id. An entry the store does not know is a problem, not an error.
        /// </summary>
        public static PstExportSelection Resolve(ExportStore store, BrowseSelection selection)
        {
            var r = new PstExportSelection();
            var boxes = store.GetMailboxes();
            var folders = new Dictionary<string, List<StoreFolder>>(StringComparer.OrdinalIgnoreCase);
            foreach (StoreMailbox m in boxes) folders[m.Id] = store.GetFolders(m.Id);
            foreach (SelectionEntry e in selection.Entries)
            {
                var matched = new List<StoreFolder>();
                string key = ExchangeEntryId.Parse(e.FolderId) == null ? null : ExchangeEntryId.FolderKeyOf(e.FolderId);
                foreach (StoreMailbox m in boxes)
                    foreach (StoreFolder f in folders[m.Id])
                    {
                        if (string.Equals(f.FolderId, e.FolderId, StringComparison.Ordinal) || (e.ExportFolderId != null && string.Equals(f.FolderId, e.ExportFolderId, StringComparison.Ordinal))
                            || (key != null && ExchangeEntryId.Parse(f.FolderId) != null && ExchangeEntryId.FolderKeyOf(f.FolderId) == key))
                            matched.Add(f);
                    }
                if (matched.Count == 0 && e.FolderPath != null)
                    foreach (StoreMailbox m in boxes)
                    {
                        if (!string.Equals(m.Kind.ToString(), selection.Kind, StringComparison.OrdinalIgnoreCase)) continue;
                        matched.AddRange(folders[m.Id].Where(f => string.Equals(f.Path, e.FolderPath, StringComparison.OrdinalIgnoreCase)));
                    }
                if (matched.Count == 0) { r.Problems.Add("'" + (e.FolderPath ?? e.FolderId) + "' is not in the store."); continue; }
                r.EntriesMatched++;
                bool itemsOnly = e.Items != null && e.Items.Count > 0;
                foreach (StoreFolder f in matched)
                {
                    if (!itemsOnly)
                    {
                        r.FolderKeys.Add(f.Key);
                        if (e.IncludeSubfolders)
                        {
                            string prefix = f.Path + "/";
                            foreach (StoreFolder d in folders[f.MailboxId])
                                if (d.Path != null && d.Path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) r.FolderKeys.Add(d.Key);
                        }
                    }
                }
                if (itemsOnly)
                {
                    int found = 0;
                    foreach (SelectionItem it in e.Items)
                    {
                        bool hit = false;
                        foreach (StoreMailbox m in boxes)
                        {
                            long? k = store.FindItemKey(m.Id, it.Id);
                            if (k.HasValue) { r.ItemKeys.Add(k.Value); hit = true; }
                        }
                        if (hit) found++;
                    }
                    if (found < e.Items.Count) r.Problems.Add((e.Items.Count - found) + " selected item(s) of '" + (e.FolderPath ?? e.FolderId) + "' are not in the store.");
                }
            }
            return r;
        }
    }

    // ---------------------------------------------------------------------------------------------
    //  The exporter
    // ---------------------------------------------------------------------------------------------

    /// <summary>
    /// File names the product creates (PST parts, manifest, repaired PST, selection files, stores). A Windows device name (CON, PRN, AUX, NUL,
    /// CONIN$, CONOUT$, COM0-9, LPT0-9 and COM and LPT followed by a superscript 1, 2 or 3), alone or followed by any extension, in any case,
    /// is not a file: Windows opens the device, and Outlook crashes on a PST named "aux.pst" (the content is fine). Policy: a name given by the
    /// operator that is not usable is refused before any work, with the name to use instead (<see cref="Require"/>); a name the product derives
    /// (user, kind of mailbox) is made safe (<see cref="SafeStem"/>) and reported by the caller.
    /// </summary>
    public static class FileNaming
    {
        static readonly HashSet<string> Devices = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$",
            "COM0", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9", "COM\u00B9", "COM\u00B2", "COM\u00B3",
            "LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9", "LPT\u00B9", "LPT\u00B2", "LPT\u00B3"
        };

        /// <summary>The part of a file name before the first dot, without trailing spaces: what Windows compares with the device names.</summary>
        static string Head(string fileName)
        {
            string s = (fileName ?? "").TrimEnd(' ', '.');
            int dot = s.IndexOf('.');
            return (dot >= 0 ? s.Substring(0, dot) : s).TrimEnd(' ');
        }

        /// <summary>True when the file name (with or without extension, in any case) is a Windows device name.</summary>
        public static bool IsReservedDeviceName(string fileName) { return Devices.Contains(Head(fileName)); }

        /// <summary>The device name a file name stands for ("AUX" for "aux.v2.pst"), or null.</summary>
        public static string ReservedDeviceName(string fileName) { string h = Head(fileName); return Devices.Contains(h) ? h.ToUpperInvariant() : null; }

        /// <summary>
        /// A name that can be created from a name that cannot: invalid characters become "_", trailing dots and spaces go, a device name gets "_"
        /// after its head ("aux.pst" -> "aux_.pst", "CON.v2.pst" -> "CON_.v2.pst"), an empty name becomes "mailbox".
        /// </summary>
        public static string SafeFileName(string fileName)
        {
            var sb = new StringBuilder();
            char[] bad = System.IO.Path.GetInvalidFileNameChars();
            foreach (char c in (fileName ?? "").Trim())
                sb.Append(Array.IndexOf(bad, c) >= 0 ? '_' : c);
            string s = sb.ToString().TrimEnd(' ', '.').TrimStart(' ');
            if (s.Length == 0) return "mailbox";
            string head = Head(s);
            if (Devices.Contains(head)) s = head + "_" + s.Substring(head.Length);
            return s;
        }

        /// <summary>A file name stem (no extension) that is safe, see <see cref="SafeFileName"/>.</summary>
        public static string SafeStem(string stem) { return SafeFileName(stem); }

        /// <summary>Null when the name can be created, else why not.</summary>
        public static string NameProblem(string name)
        {
            if (string.IsNullOrEmpty(name)) return "the name is empty";
            string dev = ReservedDeviceName(name);
            if (dev != null) return "it is the reserved Windows device name " + dev + " (with or without extension), which Windows cannot use as a file name and Outlook crashes on";
            if (name.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0) return "it has a character that Windows does not allow in a file name";
            if (name.EndsWith(".") || name.EndsWith(" ")) return "it ends with a dot or a space, which Windows silently removes";
            return null;
        }

        /// <summary>
        /// Null when the path can be created, else the reason with the name to use instead. The last name is always checked (<paramref name="isFile"/>:
        /// a file or a folder); the folders above it only when they do not exist yet, since the product would create them.
        /// </summary>
        public static string PathProblem(string path, bool isFile, string what)
        {
            if (string.IsNullOrWhiteSpace(path)) return what + ": no path given.";
            string root = System.IO.Path.GetPathRoot(path) ?? "";
            string rest = path.Substring(Math.Min(root.Length, path.Length));
            string[] parts = rest.Split(new[] { '\\', '/' }, StringSplitOptions.RemoveEmptyEntries);
            string current = root;
            for (int i = 0; i < parts.Length; i++)
            {
                current = current.Length == 0 ? parts[i] : System.IO.Path.Combine(current, parts[i]);
                if (parts[i] == "." || parts[i] == "..") continue;
                bool isLeafFile = isFile && i == parts.Length - 1;
                if (!isLeafFile && Directory.Exists(current)) continue;
                string why = NameProblem(parts[i]);
                if (why != null)
                    return what + ": '" + parts[i] + "' cannot be used (" + why + "). Use another name, for example '" + SafeFileName(parts[i]) + "'.";
            }
            return null;
        }

        /// <summary>Throws an ArgumentException with the reason and the name to use instead when the path cannot be created.</summary>
        public static void Require(string path, bool isFile, string what)
        {
            string why = PathProblem(path, isFile, what);
            if (why != null) throw new ArgumentException(why);
        }
    }

    public static class PstExporter
    {
        const long MinPartBytes = 1L << 20;
        /// <summary>What a PST adds to the FastTransfer stream of an item, on average (properties, tables, tree pages): used for the estimate of the plan.</summary>
        const long ItemOverhead = 8192;

        static readonly Regex GuidEight = new Regex(@"[0-9a-fA-F]{8}", RegexOptions.Compiled);

        /// <summary>File of a part: <c>name.pst</c> for the first, <c>name-002.pst</c>, <c>name-003.pst</c>... for the next ones.</summary>
        public static string PartPath(string destination, int index)
        {
            string full = System.IO.Path.GetFullPath(destination);
            string dir = System.IO.Path.GetDirectoryName(full);
            string stem = System.IO.Path.GetFileNameWithoutExtension(full);
            return System.IO.Path.Combine(dir, index <= 1 ? stem + ".pst" : stem + "-" + index.ToString("D3") + ".pst");
        }

        // ---------------------------------------------------------------- the tree

        static string Guid8(string mailboxId)
        {
            Match m = GuidEight.Match(mailboxId ?? "");
            return m.Success ? m.Value.ToLowerInvariant() : (mailboxId ?? "");
        }

        static string FolderIdentity(StoreFolder f)
        {
            return ExchangeEntryId.Parse(f.FolderId) != null ? "k:" + ExchangeEntryId.FolderKeyOf(f.FolderId) : "p:" + (f.Path ?? f.DisplayName ?? "").ToLowerInvariant();
        }

        sealed class Tree
        {
            public List<PstExportLocation> Locations = new List<PstExportLocation>();
            public List<PstExportFolder> Folders = new List<PstExportFolder>();
            public List<string> Notes = new List<string>();
        }

        static bool Selected(StoreMailbox m, PstExportOptions o)
        {
            if (o.MailboxIds != null && o.MailboxIds.Length > 0 && !o.MailboxIds.Any(id => string.Equals(id, m.Id, StringComparison.OrdinalIgnoreCase))) return false;
            if (o.Kinds != null && o.Kinds.Length > 0)
            {
                // with the merge on, "Archive" is the whole archive: the main archive and its auxiliary archives
                if (o.Kinds.Contains(m.Kind)) return true;
                return false;
            }
            return true;
        }

        static List<StoreMailbox> Chosen(List<StoreMailbox> all, PstExportOptions o)
        {
            List<StoreMailbox> chosen = all.Where(m => Selected(m, o)).ToList();
            // with the merge on and Archive asked, the auxiliary archives follow (they are the same logical archive)
            if (o.MergeAuxIntoArchive && o.Kinds != null && o.Kinds.Contains(StoreMailboxKind.Archive))
                foreach (StoreMailbox m in all)
                    if (m.Kind == StoreMailboxKind.AuxArchive && !chosen.Contains(m) && (o.MailboxIds == null || o.MailboxIds.Length == 0)) chosen.Add(m);
            return chosen;
        }

        /// <summary>
        /// The default name (no extension) of the first PST when the operator gives a folder: the local part of the user ("adele" for adele@contoso.com),
        /// and for an export of auxiliary archives only "-AuxArchive1" (the number of the auxiliary archive in the store; "-AuxArchives" for several),
        /// made safe for the file system (<see cref="FileNaming.SafeStem"/>). Never a bare Windows device name.
        /// </summary>
        public static string DefaultStem(ExportStore store, PstExportOptions options)
        {
            if (store == null) throw new ArgumentNullException("store");
            PstExportOptions o = options ?? new PstExportOptions();
            List<StoreMailbox> all = store.GetMailboxes().OrderBy(m => m.Key).ToList();
            List<StoreMailbox> chosen = Chosen(all, o);
            string user = chosen.Concat(all).Select(m => m.Upn).FirstOrDefault(u => !string.IsNullOrEmpty(u));
            string local = user == null ? "mailbox" : user.Split('@')[0];
            local = Regex.Replace(local, @"[^\w\.\-]", "_");
            string suffix = "";
            if (chosen.Count > 0 && chosen.All(m => m.Kind == StoreMailboxKind.AuxArchive))
                suffix = chosen.Count == 1 ? "-AuxArchive" + (all.Where(m => m.Kind == StoreMailboxKind.AuxArchive).ToList().IndexOf(chosen[0]) + 1) : "-AuxArchives";
            return FileNaming.SafeStem(local + suffix);
        }

        /// <summary>The manifest of an earlier run of the destination that this run would continue (null: a new export, Force, or no readable manifest; Run reports an unreadable one).</summary>
        static PstExportManifest PeekManifest(PstExportOptions o)
        {
            if (o.Force || string.IsNullOrWhiteSpace(o.Destination)) return null;
            try
            {
                string path = PstExportManifest.PathFor(System.IO.Path.GetFullPath(o.Destination));
                return File.Exists(path) ? PstExportManifest.Load(path) : null;
            }
            catch (Exception ex) when (ex is JsonException || ex is InvalidDataException || ex is IOException || ex is UnauthorizedAccessException) { return null; }
        }

        /// <summary>
        /// The top folder of an auxiliary archive: the current name, except when the export being continued already wrote that archive under the 1.0.x name
        /// (<see cref="PstExportOptions.LegacyAuxArchiveNameFormat"/>): the same name is kept for the rest of that export.
        /// </summary>
        static string AuxArchiveName(PstExportOptions o, PstExportManifest resumed, StoreMailbox m, int number)
        {
            string legacy = string.Format(PstExportOptions.LegacyAuxArchiveNameFormat, number, Guid8(m.Id));
            if (resumed != null)
                foreach (PstExportLocationInfo l in resumed.Locations)
                    if (l.Name != null && string.Equals(l.Kind, "AuxArchive", StringComparison.OrdinalIgnoreCase) && l.MailboxIds.Any(id => string.Equals(id, m.Id, StringComparison.OrdinalIgnoreCase))
                        && (string.Equals(l.Name, legacy, StringComparison.OrdinalIgnoreCase) || l.Name.StartsWith(legacy + " (", StringComparison.OrdinalIgnoreCase)))
                        return legacy;
            return string.Format(o.AuxArchiveNameFormat, number, Guid8(m.Id));
        }

        static Tree BuildTree(ExportStore store, PstExportOptions o, CancellationToken ct, PstExportManifest resumed = null)
        {
            var t = new Tree();
            List<StoreMailbox> all = store.GetMailboxes().OrderBy(m => m.Key).ToList();
            List<StoreMailbox> chosen = Chosen(all, o);

            var groups = new List<PstExportLocation>();
            var archiveGroups = new List<PstExportLocation>();
            var names = new Dictionary<PstExportLocation, StoreMailbox>();
            foreach (StoreMailbox m in chosen.Where(x => x.Kind == StoreMailboxKind.Primary))
                groups.Add(NewGroup(m, "Primary", o.PrimaryName, names));
            foreach (StoreMailbox m in chosen.Where(x => x.Kind == StoreMailboxKind.Archive))
            {
                PstExportLocation g = NewGroup(m, "Archive", o.ArchiveName, names);
                groups.Add(g); archiveGroups.Add(g);
            }
            int auxNumber = 0;
            foreach (StoreMailbox m in all.Where(x => x.Kind == StoreMailboxKind.AuxArchive))
            {
                auxNumber++;
                if (!chosen.Contains(m)) continue;
                if (o.MergeAuxIntoArchive)
                {
                    PstExportLocation g = archiveGroups.FirstOrDefault(a => names[a].Upn != null && string.Equals(names[a].Upn, m.Upn, StringComparison.OrdinalIgnoreCase)) ?? archiveGroups.FirstOrDefault();
                    if (g == null)
                    {
                        g = NewGroup(m, "Archive", o.ArchiveName, names);
                        groups.Add(g); archiveGroups.Add(g);
                    }
                    else { g.MailboxIds.Add(m.Id); g.Members.Add(m.Id); }
                }
                else
                {
                    PstExportLocation g = NewGroup(m, "AuxArchive", AuxArchiveName(o, resumed, m, auxNumber), names);
                    groups.Add(g);
                }
            }
            // two locations with the same name (two users in one store): the user tells them apart
            foreach (var same in groups.GroupBy(g => g.Name, StringComparer.OrdinalIgnoreCase).Where(x => x.Count() > 1))
                foreach (PstExportLocation g in same)
                {
                    StoreMailbox m = names[g];
                    g.Name = g.Name + " (" + (string.IsNullOrEmpty(m.Upn) ? Guid8(m.Id) : m.Upn) + ")";
                }
            for (int i = 0; i < groups.Count; i++) groups[i].Index = i;
            t.Locations = groups;

            // one flat list of folder rows per mailbox of the store
            var rows = new Dictionary<string, List<StoreFolder>>(StringComparer.OrdinalIgnoreCase);
            foreach (StoreMailbox m in all) rows[m.Id] = store.GetFolders(m.Id);

            foreach (PstExportLocation g in groups)
            {
                ct.ThrowIfCancellationRequested();
                var top = new PstExportFolder { Location = g.Index, Name = g.Name, ContainerClass = "IPF.Note", IsTop = true, Depth = 0 };
                g.Top = top;
                var nodeOfRow = new Dictionary<string, PstExportFolder>(StringComparer.Ordinal);
                foreach (string member in g.MailboxIds.ToList())
                {
                    List<StoreFolder> list = rows[member];
                    var byId = new Dictionary<string, StoreFolder>(StringComparer.Ordinal);
                    foreach (StoreFolder f in list) byId[f.FolderId] = f;
                    foreach (StoreFolder f in list) Place(g, top, member, f, byId, nodeOfRow, new HashSet<string>(StringComparer.Ordinal));
                }
                // the rows of the other mailboxes of the store that name a folder of this location (the items of an expanded folder are
                // recorded under the folder of the main archive while an auxiliary archive holds them)
                foreach (StoreMailbox m in all)
                {
                    if (g.Members.Contains(m.Id)) continue;
                    foreach (StoreFolder f in rows[m.Id])
                    {
                        PstExportFolder n;
                        if (g.ByKey.TryGetValue(FolderIdentity(f), out n) && !n.SourceFolderKeys.Contains(f.Key)) n.SourceFolderKeys.Add(f.Key);
                    }
                }
            }

            // standard folders first, then by name; depth first numbering
            foreach (PstExportLocation g in groups)
            {
                foreach (PstExportFolder child in g.Top.Children)
                {
                    int rank; string cls;
                    string wk = PstWellKnown.Detect(child.Name, out rank, out cls);
                    child.WellKnown = wk;
                    child.Rank = wk == null ? 100 : rank;
                    child.ContainerClass = PstWellKnown.ContainerClass(child.ContainerClass == "" ? null : child.ContainerClass, cls);
                }
                Number(g.Top, null, t.Folders);
                g.Folders = t.Folders.Count(f => f.Location == g.Index);
            }
            return t;
        }

        static PstExportLocation NewGroup(StoreMailbox m, string kind, string name, Dictionary<PstExportLocation, StoreMailbox> names)
        {
            var g = new PstExportLocation { Name = name, Kind = kind };
            g.MailboxIds.Add(m.Id); g.Members.Add(m.Id);
            names[g] = m;
            return g;
        }

        static PstExportFolder Place(PstExportLocation g, PstExportFolder top, string mailbox, StoreFolder f, Dictionary<string, StoreFolder> byId, Dictionary<string, PstExportFolder> nodeOfRow, HashSet<string> visiting)
        {
            string rowKey = mailbox + "|" + f.FolderId;
            PstExportFolder node;
            if (nodeOfRow.TryGetValue(rowKey, out node)) return node;
            string identity = FolderIdentity(f);
            if (g.ByKey.TryGetValue(identity, out node))
            {
                if (!node.SourceFolderKeys.Contains(f.Key)) node.SourceFolderKeys.Add(f.Key);
                if (string.IsNullOrEmpty(node.ContainerClass) && !string.IsNullOrEmpty(f.FolderType)) node.ContainerClass = f.FolderType;
                nodeOfRow[rowKey] = node;
                return node;
            }
            PstExportFolder parent = top;
            StoreFolder pf;
            if (f.ParentFolderId != null && f.ParentFolderId != f.FolderId && byId.TryGetValue(f.ParentFolderId, out pf) && visiting.Add(f.FolderId))
                parent = Place(g, top, mailbox, pf, byId, nodeOfRow, visiting);
            node = new PstExportFolder { Location = g.Index, Name = string.IsNullOrEmpty(f.DisplayName) ? "(no name)" : f.DisplayName, ContainerClass = f.FolderType ?? "" };
            node.SourceFolderKeys.Add(f.Key);
            parent.Children.Add(node);
            g.ByKey[identity] = node;
            nodeOfRow[rowKey] = node;
            node.Depth = parent.Depth + 1;
            return node;
        }

        static void Number(PstExportFolder n, PstExportFolder parent, List<PstExportFolder> list)
        {
            n.Index = list.Count;
            n.Parent = parent == null ? -1 : parent.Index;
            n.Path = parent == null ? n.Name : parent.Path + "/" + n.Name;
            n.Depth = parent == null ? 0 : parent.Depth + 1;
            if (n.Depth >= 2) n.ContainerClass = PstWellKnown.ContainerClass(n.ContainerClass == "" ? null : n.ContainerClass, null);
            list.Add(n);
            n.Children.Sort((a, b) =>
            {
                int c = a.Rank.CompareTo(b.Rank);
                return c != 0 ? c : string.Compare(a.Name, b.Name, StringComparison.OrdinalIgnoreCase);
            });
            foreach (PstExportFolder c in n.Children) Number(c, n, list);
        }

        // ---------------------------------------------------------------- enumeration of the items

        sealed class Entry { public PstExportFolder Folder; public StoreItem Item; }

        static IEnumerable<Entry> Enumerate(ExportStore store, PstExportOptions o, Tree t, CancellationToken ct)
        {
            bool all = o.FolderKeys == null && o.ItemKeys == null;
            int page = Math.Max(1, o.PageSize);
            foreach (PstExportFolder f in t.Folders)
            {
                if (f.IsTop) continue;
                PstExportLocation g = t.Locations[f.Location];
                foreach (long src in f.SourceFolderKeys)
                {
                    bool whole = all || (o.FolderKeys != null && o.FolderKeys.Contains(src));
                    if (!whole && (o.ItemKeys == null || o.ItemKeys.Count == 0)) continue;
                    int p = 0;
                    StoreItemPage pg;
                    do
                    {
                        ct.ThrowIfCancellationRequested();
                        pg = store.GetItems(src, p, page, StoreItemSort.ReceivedDate, false, StoreItemState.Exported);
                        foreach (StoreItem it in pg.Items)
                        {
                            if (!g.Members.Contains(it.MailboxId)) continue;
                            if (!whole && !o.ItemKeys.Contains(it.Key)) continue;
                            yield return new Entry { Folder = f, Item = it };
                        }
                        p++;
                    } while (p < pg.PageCount);
                }
            }
        }

        // ---------------------------------------------------------------- plan

        /// <summary>Read only: the locations, folders, item counts, announced sizes and the estimated number of PST files. Nothing is written.</summary>
        public static PstExportPlan Plan(ExportStore store, PstExportOptions options, CancellationToken ct = default(CancellationToken))
        {
            if (store == null) throw new ArgumentNullException("store");
            PstExportOptions o = options ?? new PstExportOptions();
            if (!string.IsNullOrWhiteSpace(o.Destination)) FileNaming.Require(o.Destination, true, "PST destination");
            return PlanCore(store, o, BuildTree(store, o, ct, PeekManifest(o)), ct);
        }

        static PstExportPlan PlanCore(ExportStore store, PstExportOptions o, Tree t, CancellationToken ct)
        {
            var plan = new PstExportPlan { MaxPartBytes = o.MaxPartBytes, MergeAuxIntoArchive = o.MergeAuxIntoArchive, Locations = t.Locations, Folders = t.Folders };
            long cur = 0; int parts = 0, curItems = 0;
            foreach (Entry e in Enumerate(store, o, t, ct))
            {
                PstExportFolder f = e.Folder;
                long len = e.Item.Length ?? 0;
                f.Items++; f.AnnouncedBytes += e.Item.Size; f.StoredBytes += len;
                f.Explicit = true;
                long est = len + ItemOverhead;
                if (parts == 0 || (curItems > 0 && cur + est > o.MaxPartBytes)) { parts++; cur = 0; curItems = 0; }
                cur += est; curItems++;
                if (est > o.MaxPartBytes) plan.OversizeItems.Add(e.Item.Key);
            }
            plan.EstimatedParts = parts;
            // folders that stay: the selected ones, those with items, their ancestors; with no selection, all of them
            bool all = o.FolderKeys == null && o.ItemKeys == null;
            foreach (PstExportFolder f in t.Folders)
            {
                if (f.IsTop) continue;
                f.Included = all || f.Items > 0 || (o.FolderKeys != null && f.SourceFolderKeys.Any(k => o.FolderKeys.Contains(k)));
            }
            for (int i = t.Folders.Count - 1; i >= 0; i--)
            {
                PstExportFolder f = t.Folders[i];
                if (f.IsTop) continue;
                if (f.Included && f.Parent >= 0) t.Folders[f.Parent].Included = true;
            }
            foreach (PstExportFolder f in t.Folders)
            {
                if (!f.IsTop) continue;
                // a location with nothing in it is left out
                f.Included = t.Folders.Any(x => x.Location == f.Location && !x.IsTop && x.Included);
            }
            foreach (PstExportLocation g in t.Locations)
            {
                g.Folders = t.Folders.Count(f => f.Location == g.Index && f.Included);
                foreach (PstExportFolder f in t.Folders.Where(x => x.Location == g.Index)) { g.Items += f.Items; g.AnnouncedBytes += f.AnnouncedBytes; g.StoredBytes += f.StoredBytes; }
                plan.Items += g.Items; plan.AnnouncedBytes += g.AnnouncedBytes; plan.StoredBytes += g.StoredBytes;
            }
            foreach (StoreMailbox m in store.GetMailboxes().Where(x => t.Locations.Any(l => l.Members.Contains(x.Id))))
                foreach (StoreStateStat s in store.GetItemStats(m.Id))
                    if (s.State != StoreItemState.Exported.ToString() && s.Count > 0)
                    {
                        plan.ItemsNotExported += s.Count;
                        plan.Notes.Add(s.Count + " item(s) of " + m.Id + " are " + s.State + " in the store: they are not in the PST (run the export again to complete the store).");
                    }
            if (t.Locations.Count == 0) plan.Notes.Add("No mailbox of the store matches the selection.");
            if (plan.Items == 0 && t.Locations.Count > 0) plan.Notes.Add("No exported item in the selection: no PST would be written.");
            if (plan.OversizeItems.Count > 0) plan.Notes.Add(plan.OversizeItems.Count + " item(s) are larger than the PST limit: each one is written alone in its own PST.");
            return plan;
        }

        // ---------------------------------------------------------------- run

        static bool VerifyPartFile(string path, long expectedBytes, out string why)
        {
            why = null;
            try
            {
                if (!File.Exists(path)) { why = "file missing"; return false; }
                long len = new FileInfo(path).Length;
                if (len != expectedBytes) { why = "size " + len + " but the manifest says " + expectedBytes; return false; }
                using (PstFile f = PstFile.Open(path))
                    if (f.FileEof != len) { why = "header says " + f.FileEof + " bytes"; return false; }
                return true;
            }
            catch (Exception ex) { why = ex.Message; return false; }
        }

        /// <summary>
        /// Existing files that have the name of a part of this destination: <c>name.pst</c> and <c>name-NNN.pst</c> (exactly three digits, the form
        /// <see cref="PartPath"/> writes). A file such as <c>name-2019.pst</c> or <c>name-old.pst</c> is not ours and is never listed.
        /// </summary>
        static IEnumerable<string> LeftoverFiles(string destination)
        {
            string full = System.IO.Path.GetFullPath(destination);
            string dir = System.IO.Path.GetDirectoryName(full);
            string stem = System.IO.Path.GetFileNameWithoutExtension(full);
            if (!Directory.Exists(dir)) yield break;
            var rx = new Regex("^" + Regex.Escape(stem) + @"(-\d{3})?\.pst$", RegexOptions.IgnoreCase);
            foreach (string f in Directory.GetFiles(dir, "*.pst"))
                if (rx.IsMatch(System.IO.Path.GetFileName(f))) yield return f;
        }

        /// <summary>
        /// Deletes the files of the parts <paramref name="from"/> to <paramref name="to"/> of this destination, by their exact names. Only the
        /// parts a manifest lists (and the one a killed run was writing) are removed this way: no other file of the folder is looked at.
        /// </summary>
        static void DeleteParts(string destination, int from, int to)
        {
            for (int i = from; i <= to; i++) TryDelete(PartPath(destination, i));
        }

        static void TryDelete(string path) { try { if (File.Exists(path)) File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { } }

        static void Report(IProgress<PstExportProgress> progress, PstExportProgress p) { if (progress != null) progress.Report(p); }

        /// <summary>
        /// Writes the PST files. An earlier run of the same destination (manifest present) is resumed: the parts it completed are verified (file
        /// present, size of the manifest, PST header) and kept, the part it left unfinished is deleted and rebuilt, and the items already in a
        /// kept part are not written again. Cancelling (token) stops after the item being written, deletes the unfinished part and returns a
        /// result with Cancelled set; the next run continues. Needs a store opened for reading; it does not modify it.
        /// </summary>
        public static PstExportResult Run(ExportStore store, PstExportOptions options, IProgress<PstExportProgress> progress, CancellationToken ct)
        {
            if (store == null) throw new ArgumentNullException("store");
            if (options == null || string.IsNullOrWhiteSpace(options.Destination)) throw new ArgumentException("Destination is needed.");
            PstExportOptions o = options;
            FileNaming.Require(o.Destination, true, "PST destination");
            if (o.MaxPartBytes < MinPartBytes) throw new ArgumentOutOfRangeException("options", "MaxPartBytes must be at least 1 MiB.");
            var sw = Stopwatch.StartNew();
            var res = new PstExportResult();
            string dest = System.IO.Path.GetFullPath(o.Destination);
            string dir = System.IO.Path.GetDirectoryName(dest);
            Directory.CreateDirectory(dir);
            string manifestPath = PstExportManifest.PathFor(dest);
            res.ManifestPath = manifestPath;

            Report(progress, new PstExportProgress { Phase = "plan", Message = "reading the folders and the items of the store" });
            Tree tree = BuildTree(store, o, ct, PeekManifest(o));
            PstExportPlan plan = PlanCore(store, o, tree, ct);
            res.Plan = plan;
            if (o.Folders != null)
                foreach (PstExportFolder pf in plan.IncludedFolders) o.Folders.Expect(pf.Path, pf.Items);
            if (plan.Items == 0) { res.Notes.Add("Nothing to write: " + (plan.Locations.Count == 0 ? "no mailbox matches" : "no exported item in the selection") + "."); res.Complete = plan.ItemsNotExported == 0; res.Elapsed = sw.Elapsed; return res; }

            using (new FileStream(System.IO.Path.Combine(dir, System.IO.Path.GetFileNameWithoutExtension(dest) + ".pstexport.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.DeleteOnClose))
            {
                string user = store.GetMailboxes().Select(m => m.Upn).FirstOrDefault(u => !string.IsNullOrEmpty(u));
                string displayName = !string.IsNullOrEmpty(o.DisplayName) ? o.DisplayName : "MailboxTransfer" + (user != null ? " " + user : "");
                // ---- resume
                PstExportManifest manifest = null;
                var done = new List<PstExportPartInfo>();
                var doneKeys = new HashSet<long>();
                if (o.Force)
                {
                    // only what this destination wrote: the destination itself, the parts of the manifest and the one after them (a run killed while writing it)
                    int last = 1;
                    try { if (File.Exists(manifestPath)) last = Math.Max(1, PstExportManifest.Load(manifestPath).Parts.Select(x => x.Index).DefaultIfEmpty(0).Max()); }
                    catch (Exception ex) when (ex is JsonException || ex is InvalidDataException || ex is IOException) { }
                    TryDelete(manifestPath);
                    DeleteParts(dest, 1, last + 1);
                }
                else if (File.Exists(manifestPath))
                {
                    try { manifest = PstExportManifest.Load(manifestPath); }
                    catch (Exception ex) when (ex is JsonException || ex is InvalidDataException)
                    { throw new InvalidOperationException("The manifest " + manifestPath + " cannot be read (" + ex.Message + "). Use Force to start again.", ex); }
                    if (manifest.MaxPartBytes != o.MaxPartBytes || manifest.MergeAuxIntoArchive != o.MergeAuxIntoArchive)
                        throw new InvalidOperationException("The manifest " + manifestPath + " was written with other options (part size " + manifest.MaxPartBytes + ", merge " + manifest.MergeAuxIntoArchive +
                            "): the parts would not line up. Use the same options, or Force to start again.");
                    foreach (PstExportPartInfo p in manifest.Parts.OrderBy(x => x.Index))
                    {
                        string path = PartPath(dest, p.Index);
                        string why = null;
                        if (p.Index != done.Count + 1) why = "out of sequence";
                        else if (p.Checked && !p.Verified) why = "the check found errors";
                        else VerifyPartFile(path, p.Bytes, out why);
                        if (why != null)
                        {
                            res.Notes.Add("Part " + p.Index + " (" + System.IO.Path.GetFileName(path) + ") is not usable (" + why + "), it is written again.");
                            break;
                        }
                        p.Reused = true; p.Path = path;
                        done.Add(p);
                        foreach (PstExportItemRecord r in p.Items) doneKeys.Add(r.Key);
                    }
                    res.PartsReused = done.Count;
                    res.PartsRebuilt = manifest.Parts.Count - done.Count;
                    int lastListed = Math.Max(done.Count, manifest.Parts.Select(x => x.Index).DefaultIfEmpty(0).Max());
                    manifest.Parts = done.ToList();
                    DeleteParts(dest, done.Count + 1, lastListed + 1);   // the parts the manifest listed and could not be reused, and the one a killed run was writing
                }
                else
                {
                    var existing = LeftoverFiles(dest).ToList();
                    if (existing.Count > 0)
                        throw new IOException("Files already exist at the destination without a manifest (" + System.IO.Path.GetFileName(existing[0]) + "): use Force to replace them, or another destination.");
                }
                if (manifest == null)
                    manifest = new PstExportManifest { CreatedUtc = DateTime.UtcNow, MaxPartBytes = o.MaxPartBytes, MergeAuxIntoArchive = o.MergeAuxIntoArchive };
                manifest.Store = store.Name + (string.IsNullOrEmpty(store.Source) ? "" : " (" + store.Source + ")");
                manifest.DisplayName = displayName;
                manifest.PlannedItems = plan.Items;
                manifest.Complete = false;
                manifest.Locations = plan.Locations.Select(l => new PstExportLocationInfo { Name = l.Name, Kind = l.Kind, MailboxIds = l.MailboxIds.ToList(), Items = l.Items, AnnouncedBytes = l.AnnouncedBytes, StoredBytes = l.StoredBytes }).ToList();
                manifest.Folders = plan.IncludedFolders.Select(f => new PstExportFolderInfo { Index = f.Index, Parent = f.Parent, Path = f.Path, ContainerClass = f.ContainerClass, WellKnown = f.WellKnown, Items = f.Items, AnnouncedBytes = f.AnnouncedBytes, StoredBytes = f.StoredBytes }).ToList();
                manifest.Failed = new List<PstExportProblem>();
                manifest.Oversize = manifest.Oversize.Where(x => x.Part <= done.Count).ToList();
                res.Oversize.AddRange(manifest.Oversize);
                foreach (PstExportPartInfo p in done) { res.ItemsReused += p.ItemCount; }
                res.Parts.AddRange(done);
                long itemsDone = res.ItemsReused;
                long bytesDone = done.Sum(p => p.SourceBytes);

                // ---- write
                PstWriter w = null;
                string wPath = null;
                int wIndex = done.Count;
                var fmap = new Dictionary<int, PstFolder>();
                PstExportPartInfo cur = null;
                var sink = new PstProgressSink(pp => Report(progress, new PstExportProgress
                {
                    Phase = "write", PartIndex = wIndex, PartPath = wPath, Folder = pp.CurrentFolder, ItemsDone = itemsDone, ItemsTotal = plan.Items,
                    BytesDone = bytesDone, BytesTotal = plan.StoredBytes, PartBytes = w == null ? 0 : w.ProjectedSize
                }));
                bool fatalIo = false;
                try
                {
                    Func<int, PstFolder> ensure = null;
                    ensure = idx =>
                    {
                        PstFolder pf;
                        if (fmap.TryGetValue(idx, out pf)) return pf;
                        PstExportFolder nf = tree.Folders[idx];
                        PstFolder parent = nf.Parent < 0 ? null : ensure(nf.Parent);
                        pf = w.CreateFolder(parent, nf.Name, nf.ContainerClass);
                        fmap[idx] = pf;
                        cur.FolderCount++;
                        return pf;
                    };
                    Action start = () =>
                    {
                        wIndex++;
                        wPath = PartPath(dest, wIndex);
                        TryDelete(wPath);
                        w = PstWriter.Create(wPath, displayName);
                        fmap = new Dictionary<int, PstFolder>();
                        cur = new PstExportPartInfo { Index = wIndex, File = System.IO.Path.GetFileName(wPath), Path = wPath };
                        // the first part holds the whole tree, empty folders included
                        if (wIndex == 1) foreach (PstExportFolder f in plan.IncludedFolders) ensure(f.Index);
                    };
                    Action<Stopwatch> finish = clock =>
                    {
                        Report(progress, new PstExportProgress { Phase = "write", PartIndex = wIndex, PartPath = wPath, Message = "closing the PST", ItemsDone = itemsDone, ItemsTotal = plan.Items, BytesDone = bytesDone, BytesTotal = plan.StoredBytes });
                        w.Close();
                        w = null;
                        cur.Bytes = new FileInfo(wPath).Length;
                        cur.ItemCount = cur.Items.Count;
                        cur.CompletedUtc = DateTime.UtcNow;
                        cur.Seconds = clock.Elapsed.TotalSeconds;
                        cur.Verified = true;
                        if (cur.Bytes > o.MaxPartBytes && !cur.Oversize) res.Notes.Add("Part " + cur.Index + " is " + cur.Bytes + " bytes, above the limit of " + o.MaxPartBytes + " (the estimate of the size of the items was low).");
                        if (o.Check)
                        {
                            var sw2 = Stopwatch.StartNew();
                            var co = new PstCheckOptions { Cancellation = ct };
                            var ps = cur.Index;
                            co.ProgressReport = new PstProgressSink(pp => Report(progress, new PstExportProgress { Phase = "check", PartIndex = ps, PartPath = wPath, BytesDone = pp.BytesDone, BytesTotal = pp.BytesTotal, ItemsDone = itemsDone, ItemsTotal = plan.Items }));
                            PstCheckReport rep = PstChecker.Check(cur.Path, co);
                            if (rep.Cancelled) throw new OperationCanceledException(ct);
                            cur.Checked = true; cur.CheckErrors = rep.Errors; cur.CheckWarnings = rep.Warnings; cur.Verified = rep.Errors == 0; cur.CheckSeconds = sw2.Elapsed.TotalSeconds;
                            if (rep.Errors > 0) res.PartsWithErrors++;
                        }
                        manifest.Parts.Add(cur);
                        res.Parts.Add(cur);
                        res.PartsWritten++;
                        res.BytesWritten += cur.Bytes;
                        manifest.Save(manifestPath);
                        Report(progress, new PstExportProgress { Phase = "manifest", PartIndex = cur.Index, PartPath = cur.Path, ItemsDone = itemsDone, ItemsTotal = plan.Items, BytesDone = bytesDone, BytesTotal = plan.StoredBytes, Message = "part " + cur.Index + " done" });
                        cur = null;
                    };

                    Stopwatch partClock = null;
                    foreach (Entry e in Enumerate(store, o, tree, ct))
                    {
                        ct.ThrowIfCancellationRequested();
                        StoreItem it = e.Item;
                        if (doneKeys.Contains(it.Key)) { if (o.Folders != null) o.Folders.Done(e.Folder.Path, it.Length ?? 0); continue; }
                        long itemKey = it.Key;
                        FxDocument doc;
                        try
                        {
                            doc = FxCodec.Decode(FxStreamSource.FromFunc(() => store.OpenItem(itemKey), it.Length ?? 0));
                        }
                        catch (Exception ex) when (!(ex is OperationCanceledException))
                        {
                            res.Failed.Add(new PstExportProblem { ItemKey = it.Key, Subject = it.Subject, Folder = e.Folder.Path, Reason = "cannot read the item: " + ex.Message });
                            if (o.Folders != null) o.Folders.Fail(e.Folder.Path);
                            continue;
                        }
                        if (doc.Error != null || doc.Truncated)
                        {
                            res.Failed.Add(new PstExportProblem { ItemKey = it.Key, Subject = it.Subject, Folder = e.Folder.Path, Reason = doc.Error ?? "the FastTransfer stream is truncated" });
                            if (o.Folders != null) o.Folders.Fail(e.Folder.Path);
                            itemsDone++; bytesDone += it.Length ?? 0;
                            continue;
                        }
                        long est = PstWriter.EstimateSize(doc.Message);
                        if (w != null && w.MessagesWritten > 0 && w.ProjectedSize + est > o.MaxPartBytes) { finish(partClock); partClock = null; }
                        if (w == null) { partClock = Stopwatch.StartNew(); start(); }
                        PstFolder target = ensure(e.Folder.Index);
                        try
                        {
                            w.AddMessage(target, doc.Message, ct, sink);
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (IOException) { fatalIo = true; throw; }
                        catch (Exception ex)
                        {
                            res.Failed.Add(new PstExportProblem { ItemKey = it.Key, Subject = it.Subject, Folder = e.Folder.Path, Reason = "not written: " + ex.Message });
                            if (o.Folders != null) o.Folders.Fail(e.Folder.Path);
                            itemsDone++; bytesDone += it.Length ?? 0;
                            continue;
                        }
                        cur.Items.Add(new PstExportItemRecord { Key = it.Key, Folder = e.Folder.Index, Length = it.Length ?? 0, Sha256 = it.Sha256 });
                        cur.SourceBytes += it.Length ?? 0;
                        if (est > o.MaxPartBytes)
                        {
                            cur.Oversize = true;
                            var ov = new PstExportOversize { ItemKey = it.Key, Subject = it.Subject, EstimatedBytes = est, Part = cur.Index };
                            res.Oversize.Add(ov); manifest.Oversize.Add(ov);
                        }
                        itemsDone++; bytesDone += it.Length ?? 0;
                        res.ItemsWritten++;
                        if (o.Folders != null) o.Folders.Done(e.Folder.Path, it.Length ?? 0);
                        Report(progress, new PstExportProgress { Phase = "write", PartIndex = wIndex, PartPath = wPath, Folder = e.Folder.Path, ItemsDone = itemsDone, ItemsTotal = plan.Items, BytesDone = bytesDone, BytesTotal = plan.StoredBytes, PartBytes = w.ProjectedSize });
                    }
                    if (w != null) { finish(partClock); partClock = null; }
                    manifest.Failed = res.Failed.ToList();
                    // items of the store that are not Exported yet (Pending, Failed there) are not in the PST: the conversion is not complete until they are exported and it is run again
                    res.Complete = res.Failed.Count == 0 && plan.ItemsNotExported == 0;
                    manifest.Complete = res.Complete;
                    if (plan.ItemsNotExported > 0) res.Notes.Add(plan.ItemsNotExported + " item(s) of the store are not exported yet (Pending or Failed in the store): they are not in the PST. Finish the export, then run the conversion again.");
                    manifest.Save(manifestPath);
                    Report(progress, new PstExportProgress { Phase = "done", ItemsDone = itemsDone, ItemsTotal = plan.Items, BytesDone = bytesDone, BytesTotal = plan.StoredBytes });
                }
                catch (OperationCanceledException)
                {
                    res.Cancelled = true;
                    if (w != null) { try { w.Dispose(); } catch (Exception) { } w = null; TryDelete(wPath); }
                    manifest.Failed = res.Failed.ToList();
                    try { manifest.Save(manifestPath); } catch (IOException) { }
                }
                catch (Exception ex)
                {
                    if (w != null) { try { w.Dispose(); } catch (Exception) { } w = null; TryDelete(wPath); }
                    res.Error = (fatalIo ? "I/O error: " : "") + ex.Message;
                    manifest.Failed = res.Failed.ToList();
                    try { manifest.Save(manifestPath); } catch (IOException) { }
                }
            }
            res.Elapsed = sw.Elapsed;
            return res;
        }
    }
}
