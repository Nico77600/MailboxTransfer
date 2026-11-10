// =============================================================================
//  Mailbox Transfer - engine, part 5i: import from PST - sources, folder plan and scan
// -----------------------------------------------------------------------------
//  PstImportFiles       the PST files of an import: files and folders, split parts (name.pst, name-0002.pst, ...) in order
//  PstImportPrepare     integrity check of each file with our checker; a damaged file is refused unless a repair is asked for
//                       (then it is repaired into a temporary copy with PstRepair and the import reads that copy)
//  PstImportWellKnown         which folders of the PST are the Inbox, Sent Items, Deleted Items, ... (store entry ids, then names)
//  PstImportPlanner     the folders of the PST mapped to folders of the target, the items and their size, the estimated duration
//  PstImportScanner     the messages of a folder: class, subject, size of the FastTransfer stream and a fingerprint (journal key)
//
//  The engine never writes to the console: long operations take a CancellationToken and an IProgress<PstProgress>.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace MailboxTransfer
{
    public sealed class PstImportOptions
    {
        /// <summary>PST files and/or folders holding PST files (split parts are taken in order).</summary>
        public List<string> PstPaths = new List<string>();
        public string TargetMailboxId;
        /// <summary>Path in the target below which the folder tree of the PST is recreated ("Imported/2019"). Required unless MergeWellKnown covers everything.</summary>
        public string TargetFolder;
        /// <summary>PST path prefix (below the top folder of the PST) -> target path prefix (below TargetFolder). The longest prefix wins.</summary>
        public IDictionary<string, string> FolderMap;
        /// <summary>The well-known folders of the PST (Inbox, Sent Items, ...) are imported into the matching folders of the mailbox. Default: they are subfolders of TargetFolder like any other.</summary>
        public bool MergeWellKnown;
        /// <summary>Each PST file gets its own subfolder (named after its store) below TargetFolder. Default: all files are one tree (the parts of a split PST).</summary>
        public bool PerFileSubfolder;
        /// <summary>Check each file with the integrity checker first (default). False: the items that cannot be read are refused one by one.</summary>
        public bool Check = true;
        /// <summary>A damaged file is repaired into a temporary copy (Level) and the import reads the copy; else a damaged file is refused.</summary>
        public bool Repair;
        public PstRepairLevel RepairLevel = PstRepairLevel.Rebuild;
        public bool KeepDamagedAttachments;
        /// <summary>Folder for the repaired copies. Null: next to the journal.</summary>
        public string WorkPath;
        public string JobName = "pstimport";
        public int Concurrency = TransferLimits.MaxConcurrency;
        public int MaxAttempts = 3;
        public bool AutoReconcile = true;
        public double ReconcileDelaySeconds = 300;
        public int ReconcileRounds = 3;
        /// <summary>Observed rate of an import into one mailbox, for the estimated duration (about 4 GiB per hour).</summary>
        public double AssumedGiBPerHour = 4;
        /// <summary>Only the folders of the PST these patterns choose (paths below the top folder, see <see cref="FolderFilter"/>); empty: every folder.</summary>
        public List<string> IncludeFolders = new List<string>();
        /// <summary>The folders (and what is below them) these patterns choose are left out; an exclusion wins over an inclusion.</summary>
        public List<string> ExcludeFolders = new List<string>();
    }

    /// <summary>A PST file of the import, as it was checked.</summary>
    public sealed class PstImportSource
    {
        public string SourcePath;
        /// <summary>The repaired copy the import reads, or null when the source itself is read.</summary>
        public string WorkPath;
        /// <summary>Identity in the journal: file name and length (the same file again is the same source).</summary>
        public string PstId;
        public long Length, LastWriteTicks;
        public bool Checked, Repaired, Refused;
        public int CheckErrors, CheckWarnings;
        /// <summary>What was found or done (the first findings, the repair result).</summary>
        public string Note;
        public string OpenPath { get { return WorkPath ?? SourcePath; } }
        /// <summary>The check found errors and no repair was asked: the file is not imported.</summary>
        public bool Damaged { get { return Checked && CheckErrors > 0; } }
    }

    public sealed class PstFolderPlan
    {
        public string PstId;
        public uint Nid;
        /// <summary>Path below the top folder of the PST ("" for the top folder itself).</summary>
        public string PstPath;
        public string TargetPath;
        public string ContainerClass;
        /// <summary>inbox, sentitems, deleteditems, drafts, outbox, junkemail, calendar, contacts, tasks, notes, journal, or null.</summary>
        public string WellKnown;
        public long Items, Bytes, Associated;
        public bool Exists;
        public string TargetFolderId;
        public string TargetMailbox;
    }

    public sealed class PstImportPlan
    {
        public List<PstImportSource> Sources = new List<PstImportSource>();
        public List<PstFolderPlan> Folders = new List<PstFolderPlan>();
        /// <summary>Target paths that do not exist yet (parents before children).</summary>
        public List<string> PathsToCreate = new List<string>();
        public int FoldersExisting;
        public long Items, Bytes, Associated, Unreadable;
        public double EstimatedSeconds;
        /// <summary>Things the operator should know: folders left out, files that are damaged, ...</summary>
        public List<string> Notes = new List<string>();
        /// <summary>The import would not start: a damaged file without repair.</summary>
        public bool Blocked { get { return Sources.Any(s => s.Damaged && !s.Repaired); } }
    }

    public sealed class PstScanItem
    {
        public uint Nid;
        public string Class, Subject;
        /// <summary>Length of the FastTransfer stream (without the envelope), 0 when the message could not be read.</summary>
        public long Size;
        public string Fingerprint;
        /// <summary>The message cannot be read (or expressed): why.</summary>
        public string Error;
    }

    public static class PstImportFiles
    {
        static readonly Regex Part = new Regex(@"^(?<base>.*?)(?:-(?<n>\d{2,6}))?$", RegexOptions.Compiled);

        /// <summary>The files in import order: a folder gives its *.pst files (a split PST: name.pst, name-0002.pst, name-0003.pst, ...), a file stays where it was given.</summary>
        public static List<string> Resolve(IEnumerable<string> paths)
        {
            var list = new List<string>();
            foreach (string p in paths)
            {
                if (string.IsNullOrWhiteSpace(p)) continue;
                if (Directory.Exists(p))
                {
                    var files = Directory.GetFiles(p, "*.pst", SearchOption.TopDirectoryOnly).ToList();
                    if (files.Count == 0) throw new FileNotFoundException("No .pst file in the folder " + p);
                    files.Sort(CompareParts);
                    list.AddRange(files.Select(Path.GetFullPath));
                }
                else if (File.Exists(p)) list.Add(Path.GetFullPath(p));
                else throw new FileNotFoundException("PST file or folder not found: " + p);
            }
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return list.Where(seen.Add).ToList();
        }

        static int CompareParts(string a, string b)
        {
            Match ma = Part.Match(Path.GetFileNameWithoutExtension(a)), mb = Part.Match(Path.GetFileNameWithoutExtension(b));
            int c = string.Compare(ma.Groups["base"].Value, mb.Groups["base"].Value, StringComparison.OrdinalIgnoreCase);
            if (c != 0) return c;
            int na = ma.Groups["n"].Success ? int.Parse(ma.Groups["n"].Value) : 0, nb = mb.Groups["n"].Success ? int.Parse(mb.Groups["n"].Value) : 0;
            return na.CompareTo(nb);
        }

        /// <summary>
        /// Identity of a PST file in the journal: name, length and a hash of its first megabyte (the header with its CRCs, the root B-tree references and
        /// the first blocks). Two different files with the same name and size do not collide; a file that was changed (by Outlook, for instance)
        /// gets another identity, which the job reports instead of importing it again silently.
        /// </summary>
        public static string IdOf(string path)
        {
            return LegacyIdOf(path) + "|" + ContentTag(path);
        }

        /// <summary>The identity used before the content tag existed (name and length): a journal that knows it keeps it.</summary>
        public static string LegacyIdOf(string path)
        {
            var fi = new FileInfo(path);
            return fi.Name.ToLowerInvariant() + "|" + fi.Length;
        }

        const int TagBytes = 1024 * 1024;

        static string ContentTag(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 4096))
                {
                    var buf = new byte[(int)Math.Min(fs.Length, TagBytes)];
                    int got = 0, n;
                    while (got < buf.Length && (n = fs.Read(buf, got, buf.Length - got)) > 0) got += n;
                    return BitConverter.ToString(SHA256.HashData(new ReadOnlySpan<byte>(buf, 0, got))).Replace("-", "").ToLowerInvariant().Substring(0, 16);
                }
            }
            catch (IOException) { return "unread"; }
            catch (UnauthorizedAccessException) { return "unread"; }
        }
    }

    public static class PstImportPrepare
    {
        /// <summary>
        /// Checks each file with our integrity checker (unless Check is off). A file with errors is repaired into a temporary copy when the options ask for it
        /// (the copy is kept: a second run reads the same copy, so the node ids stay the ones the journal knows), else it is marked Damaged and the import must not start.
        /// <paramref name="known"/> returns what an earlier run of the job recorded for a file (check result and repaired copy), so that a resumed run does not check again.
        /// </summary>
        public static List<PstImportSource> Run(PstImportOptions o, Func<string, PstImportSource> known, string defaultWorkPath, bool plan,
            IProgress<PstProgress> progress, CancellationToken ct)
        {
            var result = new List<PstImportSource>();
            foreach (string path in PstImportFiles.Resolve(o.PstPaths))
            {
                ct.ThrowIfCancellationRequested();
                var fi = new FileInfo(path);
                var s = new PstImportSource { SourcePath = path, PstId = PstImportFiles.IdOf(path), Length = fi.Length, LastWriteTicks = fi.LastWriteTimeUtc.Ticks };
                PstImportSource before = known == null ? null : known(s.PstId);
                if (before == null && known != null)
                {
                    // a journal written before the content tag knows the file by name and length: keep that key, so that the job goes on
                    string legacy = PstImportFiles.LegacyIdOf(path);
                    PstImportSource old = known(legacy);
                    if (old != null) { s.PstId = legacy; before = old; }
                }
                if (before != null && before.Checked && before.LastWriteTicks == s.LastWriteTicks && (before.CheckErrors == 0 || before.Repaired))
                {
                    s.Checked = true; s.CheckErrors = before.CheckErrors; s.CheckWarnings = before.CheckWarnings; s.Note = before.Note;
                    s.Repaired = before.Repaired;
                    if (before.WorkPath != null && File.Exists(before.WorkPath)) s.WorkPath = before.WorkPath;
                    if (!s.Repaired || s.WorkPath != null) { result.Add(s); continue; }
                }
                if (o.Check)
                {
                    var co = new PstCheckOptions { MaxPerCode = 5, Cancellation = ct, ProgressReport = progress };
                    PstCheckReport rep = PstChecker.Check(path, co);
                    ct.ThrowIfCancellationRequested();
                    s.Checked = true; s.CheckErrors = rep.Errors; s.CheckWarnings = rep.Warnings;
                    s.Note = rep.Errors == 0 ? (rep.Warnings == 0 ? "clean" : rep.Warnings + " warning(s)")
                        : rep.Errors + " error(s): " + string.Join("; ", rep.Findings.Where(f => f.Severity == PstSeverity.Error).Take(3).Select(f => f.Code + " " + f.Message));
                }
                if (s.Damaged && o.Repair && !plan) Repair(o, s, defaultWorkPath, progress, ct);
                else if (s.Damaged && o.Repair && plan) s.Note += " (would be repaired into a temporary copy)";
                result.Add(s);
            }
            return result;
        }

        static void Repair(PstImportOptions o, PstImportSource s, string defaultWorkPath, IProgress<PstProgress> progress, CancellationToken ct)
        {
            string dir = o.WorkPath ?? defaultWorkPath;
            if (string.IsNullOrEmpty(dir)) dir = Path.Combine(Path.GetTempPath(), "mbt-pstimport");
            Directory.CreateDirectory(dir);
            string target = Path.Combine(dir, Path.GetFileNameWithoutExtension(s.SourcePath) + "-" + s.Length + ".repaired.pst");
            var ro = new PstRepairOptions { Level = o.RepairLevel, Overwrite = true, KeepDamagedAttachments = o.KeepDamagedAttachments, CheckBefore = false, CheckAfter = true };
            PstRepairReport r = PstRepair.Repair(s.SourcePath, target, ro, progress, ct);
            if (!r.Success)
            {
                s.Note += "; repair failed: " + r.Error;
                return;
            }
            s.WorkPath = target; s.Repaired = true;
            s.Note += "; repaired into " + target + " (" + r.ItemsRecovered + " item(s) recovered, " + r.ItemsPartial + " partial, " + r.ItemsLost + " lost)";
            s.CheckErrors = r.After == null ? 0 : r.After.Errors;
        }
    }

    public static class PstImportWellKnown
    {
        static readonly Dictionary<string, string[]> Names = new Dictionary<string, string[]>
        {
            { "inbox", new[] { "Inbox", "Boîte de réception", "Posteingang", "Bandeja de entrada", "Posta in arrivo", "Postvak IN", "Caixa de Entrada" } },
            { "sentitems", new[] { "Sent Items", "Sent", "Éléments envoyés", "Gesendete Elemente", "Elementos enviados", "Posta inviata", "Verzonden items", "Itens Enviados" } },
            { "deleteditems", new[] { "Deleted Items", "Éléments supprimés", "Gelöschte Elemente", "Elementos eliminados", "Posta eliminata", "Verwijderde items", "Itens Excluídos" } },
            { "drafts", new[] { "Drafts", "Brouillons", "Entwürfe", "Borradores", "Bozze", "Concepten", "Rascunhos" } },
            { "outbox", new[] { "Outbox", "Boîte d'envoi", "Postausgang", "Bandeja de salida", "Posta in uscita", "Postvak UIT", "Caixa de Saída" } },
            { "junkemail", new[] { "Junk Email", "Junk E-mail", "Courrier indésirable", "Junk-E-Mail", "Correo no deseado", "Posta indesiderata", "Ongewenste e-mail", "Lixo Eletrónico" } },
            { "calendar", new[] { "Calendar", "Calendrier", "Kalender", "Calendario", "Agenda" } },
            { "contacts", new[] { "Contacts", "Contactos", "Kontakte", "Contatti", "Contactpersonen", "Contatos" } },
            { "tasks", new[] { "Tasks", "Tâches", "Aufgaben", "Tareas", "Attività", "Taken", "Tarefas" } },
            { "notes", new[] { "Notes", "Notizen", "Notas", "Note", "Notities" } },
            { "journal", new[] { "Journal", "Journal", "Tagebuch", "Diario", "Dagboek" } },
        };
        static readonly Dictionary<string, string> ByClass = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "IPF.Appointment", "calendar" }, { "IPF.Contact", "contacts" }, { "IPF.Task", "tasks" }, { "IPF.StickyNote", "notes" }, { "IPF.Journal", "journal" }
        };

        /// <summary>The well-known folders among the folders directly below the top folder: NID of the folder -> name (see the keys of this class).</summary>
        public static Dictionary<uint, string> Classify(PstReader r, PstReaderFolder top)
        {
            var result = new Dictionary<uint, string>();
            // 1. the entry ids of the message store (Deleted Items, Sent Items, Outbox)
            try
            {
                var store = new PstPcReader(r.File.OpenNode(PstFormat.NidMessageStore));
                foreach (var kv in new[] { new KeyValuePair<ushort, string>(0x35E3, "deleteditems"), new KeyValuePair<ushort, string>(0x35E4, "sentitems"), new KeyValuePair<ushort, string>(0x35E2, "outbox") })
                {
                    PstProp p = store.Get(kv.Key);
                    if (p == null || p.Data == null || p.Data.Length < 24) continue;
                    uint nid = BitConverter.ToUInt32(p.Data, p.Data.Length - 4);
                    if (top.Children.Any(c => c.Nid == nid) && !result.ContainsKey(nid) && !result.ContainsValue(kv.Value)) result[nid] = kv.Value;
                }
            }
            catch (Exception) { /* a store without these properties: the names decide */ }
            // 2. the names, then the container class (the first folder of the class)
            foreach (PstReaderFolder c in top.Children)
            {
                if (result.ContainsKey(c.Nid)) continue;
                foreach (var kv in Names)
                {
                    if (result.ContainsValue(kv.Key)) continue;
                    if (!kv.Value.Any(n => string.Equals(n, c.Name, StringComparison.OrdinalIgnoreCase))) continue;
                    if (ByClass.ContainsValue(kv.Key) && !ByClass.Any(b => b.Value == kv.Key && string.Equals(b.Key, c.ContainerClass, StringComparison.OrdinalIgnoreCase))) continue;
                    result[c.Nid] = kv.Key;
                    break;
                }
            }
            foreach (PstReaderFolder c in top.Children)
            {
                string wk;
                if (result.ContainsKey(c.Nid) || c.ContainerClass == null || !ByClass.TryGetValue(c.ContainerClass, out wk) || result.ContainsValue(wk)) continue;
                result[c.Nid] = wk;
            }
            return result;
        }
    }

    public static class PstImportPlanner
    {
        const uint NidIpmSubtree = 0x8022;

        /// <summary>The folder of the PST that holds the user's folders (the IPM subtree, "Top of Personal Folders"), or null.</summary>
        public static PstReaderFolder TopFolder(PstReader r)
        {
            PstReaderFolder root = r.Root;
            PstReaderFolder top = root.Children.FirstOrDefault(c => c.Nid == NidIpmSubtree);
            if (top != null) return top;
            return root.Children.FirstOrDefault(c => (c.Nid & 0x1F) == 2 && !string.Equals(c.Name, "Search Root", StringComparison.OrdinalIgnoreCase));
        }

        static string Norm(string p) { return (p ?? "").Trim('/'); }

        public static string Clean(string name)
        {
            string n = (name ?? "").Replace('/', '_').Replace('\\', '_').Trim();
            return n.Length == 0 ? "(no name)" : n;
        }

        /// <summary>Target path of a PST path (below the top folder): the map (longest prefix), then the target folder in front.</summary>
        public static string MapPath(PstImportOptions o, string root, string pstPath)
        {
            string src = Norm(pstPath), mapped = src;
            if (o.FolderMap != null)
            {
                string best = null, bestValue = null;
                foreach (var kv in o.FolderMap)
                {
                    string k = Norm(kv.Key);
                    bool hit = string.Equals(src, k, StringComparison.OrdinalIgnoreCase) || src.StartsWith(k + "/", StringComparison.OrdinalIgnoreCase);
                    if (hit && (best == null || k.Length > best.Length)) { best = k; bestValue = Norm(kv.Value); }
                }
                if (best != null)
                {
                    string rest = src.Substring(best.Length).TrimStart('/');
                    mapped = rest.Length == 0 ? bestValue : (bestValue.Length == 0 ? rest : bestValue + "/" + rest);
                }
                else mapped = RecoverableItems.Restore(src);
            }
            else mapped = RecoverableItems.Restore(src);
            root = Norm(root);
            if (root.Length == 0) return mapped;
            return mapped.Length == 0 ? root : root + "/" + mapped;
        }

        static IEnumerable<KeyValuePair<string, PstReaderFolder>> Walk(PstReaderFolder top)
        {
            var stack = new Stack<KeyValuePair<string, PstReaderFolder>>();
            stack.Push(new KeyValuePair<string, PstReaderFolder>("", top));
            while (stack.Count > 0)
            {
                var kv = stack.Pop();
                yield return kv;
                for (int i = kv.Value.Children.Count - 1; i >= 0; i--)
                {
                    PstReaderFolder c = kv.Value.Children[i];
                    stack.Push(new KeyValuePair<string, PstReaderFolder>(kv.Key.Length == 0 ? Clean(c.Name) : kv.Key + "/" + Clean(c.Name), c));
                }
            }
        }

        /// <summary>
        /// The folders of one source mapped to the target. <paramref name="tree"/> is the tree of the target (null: nothing is known about it, every folder is to create).
        /// Reads the folder tree of the PST only (no message).
        /// </summary>
        public static List<PstFolderPlan> MapFolders(PstImportOptions o, PstImportSource src, IEnumerable<FolderInfo> tree, string storeName, out string note)
        {
            return MapFolders(o, src, tree, storeName, FolderFilter.Create(o.IncludeFolders, o.ExcludeFolders), out note);
        }

        /// <summary>Same, with the filter given (a caller that maps several files shares one filter, to learn which patterns matched nothing at all).</summary>
        public static List<PstFolderPlan> MapFolders(PstImportOptions o, PstImportSource src, IEnumerable<FolderInfo> tree, string storeName, FolderFilter filter, out string note)
        {
            note = null;
            var plans = new List<PstFolderPlan>();
            var existing = new Dictionary<string, FolderInfo>(StringComparer.OrdinalIgnoreCase);
            var wellKnown = new Dictionary<string, FolderInfo>(StringComparer.OrdinalIgnoreCase);
            if (tree != null)
                foreach (FolderInfo f in tree)
                {
                    if (!existing.ContainsKey(f.Path)) existing[f.Path] = f;
                    if (!string.IsNullOrEmpty(f.WellKnownName) && !wellKnown.ContainsKey(f.WellKnownName)) wellKnown[f.WellKnownName] = f;
                }
            using (PstReader r = PstReader.Open(src.OpenPath))
            {
                PstReaderFolder top = TopFolder(r);
                if (top == null) { note = src.SourcePath + ": no folder tree (no 'Top of Personal Folders'): nothing to import"; return plans; }
                var others = r.Root.Children.Where(c => c.Nid != top.Nid && (c.Nid & 0x1F) == 2 && !string.Equals(c.Name, "Search Root", StringComparison.OrdinalIgnoreCase)).Select(c => c.Name).ToList();
                if (others.Count > 0) note = src.SourcePath + ": not imported (search folders and other folders outside the user's tree): " + string.Join(", ", others);
                Dictionary<uint, string> wk = o.MergeWellKnown ? PstImportWellKnown.Classify(r, top) : new Dictionary<uint, string>();
                Dictionary<uint, string> wkAll = PstImportWellKnown.Classify(r, top);
                string baseRoot = o.TargetFolder;
                if (o.PerFileSubfolder)
                {
                    string store = Clean(string.IsNullOrEmpty(storeName) ? top.Name : storeName);
                    baseRoot = Norm(baseRoot).Length == 0 ? store : Norm(baseRoot) + "/" + store;
                }
                // the target path of the well-known folders that are merged: the folder of the mailbox; below it, the subfolders follow
                var mergedRoot = new Dictionary<string, string>();   // PST path of a merged folder -> target path
                foreach (var kv in Walk(top))
                {
                    PstReaderFolder f = kv.Value;
                    string wkName;
                    wkAll.TryGetValue(f.Nid, out wkName);
                    string target = null;
                    if (o.MergeWellKnown && wk.ContainsKey(f.Nid) && wellKnown.ContainsKey(wk[f.Nid]))
                    {
                        target = wellKnown[wk[f.Nid]].Path;
                        mergedRoot[kv.Key] = target;
                    }
                    else
                    {
                        string parent = kv.Key.Contains('/') ? kv.Key.Substring(0, kv.Key.LastIndexOf('/')) : "";
                        string leaf = kv.Key.Substring(kv.Key.LastIndexOf('/') + 1);
                        string mergedParent;
                        if (kv.Key.Length > 0 && mergedRoot.TryGetValue(parent, out mergedParent)) { target = mergedParent + "/" + leaf; mergedRoot[kv.Key] = target; }
                        else target = MapPath(o, baseRoot, kv.Key);
                    }
                    // a folder that is not chosen is left out, its subfolders are still looked at (they can be chosen by their own pattern)
                    if (filter != null && !filter.Allows(kv.Key)) continue;
                    var fp = new PstFolderPlan
                    {
                        PstId = src.PstId, Nid = f.Nid, PstPath = kv.Key, TargetPath = target, ContainerClass = f.ContainerClass, WellKnown = wkName,
                        Associated = f.AssocCount, Items = f.ContentCount
                    };
                    FolderInfo found;
                    if (existing.TryGetValue(target, out found)) { fp.Exists = true; fp.TargetFolderId = found.Id; }
                    plans.Add(fp);
                }
            }
            return plans;
        }

        /// <summary>
        /// The plan of the whole import, read only: the files (checked), the folders mapped to the target, the number and size of the items (every message is read once,
        /// its big values stay in the file), the folders to create and the estimated duration. Nothing is written and nothing is repaired.
        /// </summary>
        public static PstImportPlan Build(PstImportOptions o, IEnumerable<FolderInfo> tree, IProgress<PstProgress> progress, CancellationToken ct)
        {
            var plan = new PstImportPlan();
            List<FolderInfo> treeList = tree == null ? null : tree.ToList();
            plan.Sources = PstImportPrepare.Run(o, null, null, true, progress, ct);
            var existing = new Dictionary<string, FolderInfo>(StringComparer.OrdinalIgnoreCase);
            if (treeList != null) foreach (FolderInfo f in treeList) if (!existing.ContainsKey(f.Path)) existing[f.Path] = f;
            var needed = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            FolderFilter filter = FolderFilter.Create(o.IncludeFolders, o.ExcludeFolders);
            foreach (PstImportSource src in plan.Sources)
            {
                if (src.Damaged && !src.Repaired)
                {
                    plan.Notes.Add(src.SourcePath + " is damaged (" + src.Note + "): " + (o.Repair ? "it would be repaired into a temporary copy by the import; the items of that copy are not counted here" : "the import is refused unless a repair is asked for (-Repair)"));
                    continue;
                }
                string note;
                string storeName = null;
                using (PstReader r = PstReader.Open(src.OpenPath)) storeName = r.DisplayName;
                List<PstFolderPlan> folders = MapFolders(o, src, treeList, storeName, filter, out note);
                if (note != null) plan.Notes.Add(note);
                using (PstReader r = PstReader.Open(src.OpenPath))
                {
                    var byNid = AllFolders(r).ToDictionary(f => f.Nid);
                    foreach (PstFolderPlan fp in folders)
                    {
                        ct.ThrowIfCancellationRequested();
                        long items = 0, bytes = 0, unreadable = 0;
                        PstReaderFolder rf = byNid[fp.Nid];
                        foreach (PstScanItem it in PstImportScanner.Scan(r, src, rf, ct))
                        {
                            if (it.Error != null) { unreadable++; continue; }
                            items++; bytes += it.Size;
                        }
                        fp.Items = items; fp.Bytes = bytes;
                        plan.Items += items; plan.Bytes += bytes; plan.Unreadable += unreadable; plan.Associated += fp.Associated;
                        if (progress != null) progress.Report(new PstProgress { Phase = "plan", Items = plan.Items, CurrentFolder = fp.PstPath });
                        plan.Folders.Add(fp);
                        var chain = new List<string>();
                        for (string p = fp.TargetPath; p.Length > 0; p = p.Contains('/') ? p.Substring(0, p.LastIndexOf('/')) : "") chain.Add(p);
                        chain.Reverse();
                        foreach (string p in chain) if (!existing.ContainsKey(p) && seen.Add(p)) needed.Add(p);
                        if (fp.Exists) plan.FoldersExisting++;
                    }
                }
            }
            if (filter != null)
            {
                foreach (string u in filter.Unmatched()) plan.Notes.Add(u + " matches no folder of the PST file(s)");
                if (plan.Folders.Count == 0) plan.Notes.Add("no folder is chosen (" + filter + "): nothing would be imported");
            }
            plan.PathsToCreate = needed.OrderBy(p => p.Count(c => c == '/')).ThenBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
            plan.EstimatedSeconds = o.AssumedGiBPerHour <= 0 ? 0 : plan.Bytes / (o.AssumedGiBPerHour * 1073741824.0) * 3600.0;
            return plan;
        }

        public static IEnumerable<PstReaderFolder> AllFolders(PstReader r) { return r.AllFolders(); }
    }

    public static class PstImportScanner
    {
        /// <summary>
        /// The messages of a folder (not the folder-associated ones), one at a time: class, subject, size and fingerprint. A message that cannot be read comes back with Error.
        /// The fingerprint identifies the content in the journal: the file, the node and its data blocks (a modified message has other blocks), the identifying properties and the size.
        /// </summary>
        public static IEnumerable<PstScanItem> Scan(PstReader r, PstImportSource src, PstReaderFolder folder, CancellationToken ct)
        {
            foreach (uint nid in r.MessageNids(folder, false))
            {
                ct.ThrowIfCancellationRequested();
                yield return ScanOne(r, src, nid);
            }
        }

        public static PstScanItem ScanOne(PstReader r, PstImportSource src, uint nid)
        {
            var it = new PstScanItem { Nid = nid };
            PstNbtEntry nbt;
            ulong bidData = 0, bidSub = 0;
            if (r.File.TryFindNode(nid, out nbt)) { bidData = nbt.BidData; bidSub = nbt.BidSub; }
            try
            {
                FxMessage m = r.ReadMessage(nid);
                FxImportStream.Prepare(m);
                it.Class = m.GetString(0x001A);
                it.Subject = Printable(m.GetString(0x0037));
                it.Size = FxCodec.PayloadLength(m);
                it.Fingerprint = Fingerprint(src.PstId, nid, bidData, bidSub, m, it.Size);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                it.Error = (ex is ArgumentException ? "cannot be expressed as a FastTransfer stream: " : "cannot be read from the PST: ") + ex.Message;
                it.Fingerprint = Fingerprint(src.PstId, nid, bidData, bidSub, null, 0);
            }
            return it;
        }

        public static string Printable(string s)
        {
            if (s == null) return null;
            var sb = new StringBuilder(s.Length);
            foreach (char c in s) if (c >= ' ') sb.Append(c);
            return sb.ToString();
        }

        static readonly ushort[] Identifying = { 0x001A, 0x0037, 0x0039, 0x0E06, 0x1035, 0x300B, 0x3001 };

        static string Fingerprint(string pstId, uint nid, ulong bidData, ulong bidSub, FxMessage m, long size)
        {
            using (var h = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
            {
                Action<string> add = s => h.AppendData(Encoding.UTF8.GetBytes(s + "\n"));
                add(pstId); add(nid.ToString()); add(bidData.ToString()); add(bidSub.ToString()); add(size.ToString());
                if (m != null)
                    foreach (ushort id in Identifying)
                    {
                        FxProperty p = m.Find(id);
                        if (p == null || p.IsLazy || p.RawLength > 4096) { add(id.ToString("X4") + ":-"); continue; }
                        add(id.ToString("X4") + ":" + Convert.ToBase64String(p.Raw ?? new byte[0]));
                    }
                return BitConverter.ToString(h.GetHashAndReset()).Replace("-", "").ToLowerInvariant();
            }
        }
    }
}
