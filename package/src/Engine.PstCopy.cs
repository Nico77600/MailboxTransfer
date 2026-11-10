// =============================================================================
//  Mailbox Transfer - engine, part 19: copy folders and messages of a PST into new PST files (PST -> PST)
// -----------------------------------------------------------------------------
//  PstCopier.Plan   what would be copied: folders, items, an estimate of the size and of the number of files (reads the folder tree and the contents tables only)
//  PstCopier.Run    reads each chosen message with PstReader and writes it with PstWriter (our own writer), split on item boundaries at MaxPartBytes
//                   (name.pst, name-002.pst, ...), every part recreates the folders its items need, and each finished part is checked with the integrity checker.
//
//  What is chosen: the folders the filter takes (FolderFilter: paths below the top folder of the PST, a path takes the folder and what is below it) and/or the
//  messages whose node ids are listed in ItemNids (any folder: the folders above them are created, holding only those messages). Neither: everything.
//  The source is only read. Folder-associated items (rules, views, settings) and the search folders are not copied.
//  No resume: a copy that is stopped or fails leaves no file (the parts written so far are deleted), the destination must not exist (or Force).
// =============================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace MailboxTransfer
{
    public sealed class PstCopyOptions
    {
        /// <summary>The PST files to read: files and/or folders of PST files (split parts in order, see <see cref="PstImportFiles"/>).</summary>
        public List<string> SourcePaths = new List<string>();
        /// <summary>The first new file (D:\out\copy.pst); the next ones are copy-002.pst, copy-003.pst...</summary>
        public string Destination;
        public long MaxPartBytes = PstSplitWriter.DefaultMaxBytes;
        public List<string> IncludeFolders = new List<string>();
        public List<string> ExcludeFolders = new List<string>();
        /// <summary>Messages chosen one by one (node ids of the PST), in addition to the chosen folders; only for a single source file.</summary>
        public HashSet<uint> ItemNids;
        public bool Check = true;
        /// <summary>Delete the files of an earlier copy of the same destination.</summary>
        public bool Force;
        /// <summary>Display name of the new PST (default: the one of the source).</summary>
        public string DisplayName;
        /// <summary>Progress per folder for the console (items to copy, copied, failed); null: not followed.</summary>
        public FolderTracker Folders;
    }

    public sealed class PstCopyFolder
    {
        public string Source;
        public string Path, Name, ContainerClass;
        public uint Nid;
        /// <summary>Messages of the folder that will be copied.</summary>
        public long Items;
        /// <summary>The folder is chosen (as a whole); false: it is created only because a chosen message, or a chosen folder, is below it.</summary>
        public bool Chosen;
        internal PstReaderFolder Folder;
        internal List<uint> Nids;
        internal PstCopyFolder Parent;
    }

    public sealed class PstCopyPlan
    {
        public List<PstCopyFolder> Folders = new List<PstCopyFolder>();
        public long Items, SourceBytes;
        /// <summary>The size of the new files is about the share of the items that are copied, of the size of the source (the messages are not measured one by one).</summary>
        public long EstimatedBytes;
        public int EstimatedParts;
        public List<string> Notes = new List<string>();
        public int FoldersChosen { get { return Folders.Count; } }
    }

    public sealed class PstCopyResult
    {
        public PstCopyPlan Plan;
        public List<PstExportPartInfo> Parts = new List<PstExportPartInfo>();
        public long ItemsWritten, BytesWritten;
        public List<PstExportProblem> Failed = new List<PstExportProblem>();
        public List<PstExportOversize> Oversize = new List<PstExportOversize>();
        public List<string> Notes = new List<string>();
        public int PartsWithErrors;
        public bool Cancelled;
        public string Error;
        public TimeSpan Elapsed;
        /// <summary>Every chosen message is in a part, and no part has check errors.</summary>
        public bool Complete { get { return !Cancelled && Error == null && Failed.Count == 0; } }
        public bool Ok { get { return Complete && PartsWithErrors == 0; } }
    }

    public static class PstCopier
    {
        const long MinPartBytes = 1L << 20;

        static List<PstCopyFolder> Walk(PstReaderFolder top, string source)
        {
            var list = new List<PstCopyFolder>();
            var stack = new Stack<Tuple<string, PstReaderFolder, PstCopyFolder>>();
            stack.Push(Tuple.Create("", top, (PstCopyFolder)null));
            while (stack.Count > 0)
            {
                var t = stack.Pop();
                var pf = new PstCopyFolder { Source = source, Path = t.Item1, Name = t.Item2.Name, ContainerClass = t.Item2.ContainerClass, Nid = t.Item2.Nid, Folder = t.Item2, Parent = t.Item3 };
                list.Add(pf);
                for (int i = t.Item2.Children.Count - 1; i >= 0; i--)
                {
                    PstReaderFolder c = t.Item2.Children[i];
                    stack.Push(Tuple.Create(t.Item1.Length == 0 ? PstImportPlanner.Clean(c.Name) : t.Item1 + "/" + PstImportPlanner.Clean(c.Name), c, pf));
                }
            }
            return list;
        }

        /// <summary>The folders of the sources, with what is chosen. Folders that are neither chosen nor above a chosen message or folder are left out.</summary>
        static List<PstCopyFolder> Choose(PstCopyOptions o, List<string> files, PstCopyPlan plan, CancellationToken ct)
        {
            FolderFilter filter = FolderFilter.Create(o.IncludeFolders, o.ExcludeFolders);
            bool everything = filter == null && (o.ItemNids == null || o.ItemNids.Count == 0);
            var chosen = new List<PstCopyFolder>();
            foreach (string file in files)
            {
                using (PstReader r = PstReader.Open(file))
                {
                    PstReaderFolder top = PstImportPlanner.TopFolder(r);
                    if (top == null) { plan.Notes.Add(file + ": no folder tree (no 'Top of Personal Folders'): nothing to copy"); continue; }
                    List<PstCopyFolder> all = Walk(top, file);
                    var keep = new HashSet<PstCopyFolder>();
                    foreach (PstCopyFolder f in all)
                    {
                        ct.ThrowIfCancellationRequested();
                        bool whole = everything || (filter != null && filter.Allows(f.Path));
                        if (whole) { f.Chosen = true; f.Nids = r.MessageNids(f.Folder, false); f.Items = f.Nids.Count; keep.Add(f); }
                        else if (o.ItemNids != null && o.ItemNids.Count > 0)
                        {
                            List<uint> mine = r.MessageNids(f.Folder, false).Where(o.ItemNids.Contains).ToList();
                            if (mine.Count > 0) { f.Nids = mine; f.Items = mine.Count; keep.Add(f); }
                        }
                    }
                    // the folders above a kept folder are needed to hold it
                    foreach (PstCopyFolder f in all.Where(keep.Contains).ToList())
                        for (PstCopyFolder p = f.Parent; p != null; p = p.Parent) if (keep.Add(p)) { p.Nids = new List<uint>(); p.Items = 0; }
                    chosen.AddRange(all.Where(keep.Contains));
                }
            }
            if (filter != null)
                foreach (string u in filter.Unmatched()) plan.Notes.Add(u + " matches no folder of the PST file(s)");
            return chosen;
        }

        static List<string> Sources(PstCopyOptions o)
        {
            if (o.SourcePaths == null || o.SourcePaths.Count == 0) throw new ArgumentException("SourcePaths is needed: the PST file(s) to copy from.");
            List<string> files = PstImportFiles.Resolve(o.SourcePaths);
            if (o.ItemNids != null && o.ItemNids.Count > 0 && files.Count != 1) throw new ArgumentException("Messages chosen by node id belong to one PST file: give exactly one source file.");
            return files;
        }

        public static PstCopyPlan Plan(PstCopyOptions o, CancellationToken ct = default(CancellationToken))
        {
            List<string> files = Sources(o);
            var plan = new PstCopyPlan();
            plan.Folders = Choose(o, files, plan, ct);
            plan.Items = plan.Folders.Sum(f => f.Items);
            long total = 0, totalItems = 0;
            foreach (string file in files)
            {
                total += new FileInfo(file).Length;
                using (PstReader r = PstReader.Open(file)) foreach (PstReaderFolder f in r.AllFolders()) totalItems += f.ContentCount;
            }
            plan.SourceBytes = total;
            plan.EstimatedBytes = totalItems <= 0 ? 0 : (long)(total * (double)Math.Min(plan.Items, totalItems) / totalItems);
            plan.EstimatedParts = plan.Items == 0 ? 0 : (int)Math.Max(1, (plan.EstimatedBytes + o.MaxPartBytes - 1) / Math.Max(1, o.MaxPartBytes));
            if (plan.Items == 0 && plan.Folders.Count == 0) plan.Notes.Add("Nothing is chosen: no folder of the PST file(s) is taken.");
            return plan;
        }

        static void TryDelete(string path) { try { if (File.Exists(path)) File.Delete(path); } catch (IOException) { } catch (UnauthorizedAccessException) { } }

        static IEnumerable<string> Leftover(string destination)
        {
            string full = Path.GetFullPath(destination);
            string dir = Path.GetDirectoryName(full), stem = Path.GetFileNameWithoutExtension(full);
            if (!Directory.Exists(dir)) yield break;
            var rx = new Regex("^" + Regex.Escape(stem) + @"(-\d{3,})?\.pst$", RegexOptions.IgnoreCase);
            foreach (string f in Directory.GetFiles(dir, "*.pst")) if (rx.IsMatch(Path.GetFileName(f))) yield return f;
        }

        static void Report(IProgress<PstExportProgress> progress, PstExportProgress p) { if (progress != null) progress.Report(p); }

        public static PstCopyResult Run(PstCopyOptions o, IProgress<PstExportProgress> progress, CancellationToken ct)
        {
            if (o == null || string.IsNullOrWhiteSpace(o.Destination)) throw new ArgumentException("Destination is needed.");
            FileNaming.Require(o.Destination, true, "PST destination");
            if (o.MaxPartBytes < MinPartBytes) throw new ArgumentOutOfRangeException("options", "MaxPartBytes must be at least 1 MiB.");
            var sw = Stopwatch.StartNew();
            var res = new PstCopyResult();
            string dest = Path.GetFullPath(o.Destination);
            Directory.CreateDirectory(Path.GetDirectoryName(dest));
            Report(progress, new PstExportProgress { Phase = "plan", Message = "reading the folders of the PST" });
            List<string> files = Sources(o);
            foreach (string f in files)
                if (string.Equals(Path.GetFullPath(f), dest, StringComparison.OrdinalIgnoreCase) || Leftover(dest).Any(x => string.Equals(Path.GetFullPath(x), Path.GetFullPath(f), StringComparison.OrdinalIgnoreCase)))
                    throw new IOException("The destination " + dest + " is the source file or one of its parts: choose another name.");
            PstCopyPlan plan = new PstCopyPlan();
            plan.Folders = Choose(o, files, plan, ct);
            plan.Items = plan.Folders.Sum(f => f.Items);
            res.Plan = plan;
            if (o.Folders != null) foreach (PstCopyFolder pf in plan.Folders) o.Folders.Expect(pf.Path.Length == 0 ? "(top)" : pf.Path, pf.Items);
            if (plan.Items == 0) { res.Notes.Add("Nothing to write: no message is in the chosen folders."); res.Elapsed = sw.Elapsed; return res; }

            var existing = Leftover(dest).ToList();
            if (existing.Count > 0)
            {
                if (!o.Force) throw new IOException("Files already exist at the destination (" + Path.GetFileName(existing[0]) + "): the copy never overwrites a file. Choose another name, or use Force to replace them.");
                foreach (string f in existing) TryDelete(f);
            }

            string displayName = o.DisplayName;
            long itemsDone = 0;
            PstWriter w = null;
            string wPath = null;
            int wIndex = 0;
            PstExportPartInfo cur = null;
            Stopwatch partClock = null;
            var written = new List<string>();
            var fmap = new Dictionary<PstCopyFolder, PstFolder>();
            var sink = new PstProgressSink(pp => Report(progress, new PstExportProgress { Phase = "write", PartIndex = wIndex, PartPath = wPath, Folder = pp.CurrentFolder, ItemsDone = itemsDone, ItemsTotal = plan.Items, PartBytes = w == null ? 0 : w.ProjectedSize }));
            bool fatalIo = false;
            PstReader reader = null;
            string readerFile = null;
            try
            {
                Func<PstCopyFolder, PstFolder> ensure = null;
                ensure = f =>
                {
                    PstFolder pf;
                    if (fmap.TryGetValue(f, out pf)) return pf;
                    if (f.Parent == null) { pf = w.IpmSubtree; fmap[f] = pf; return pf; }
                    PstFolder parent = ensure(f.Parent);
                    // the writer makes a Deleted Items of its own: the one of the source takes its place
                    if (parent == w.IpmSubtree && string.Equals(f.Name, "Deleted Items", StringComparison.OrdinalIgnoreCase)) pf = w.DeletedItems;
                    else { pf = w.CreateFolder(parent, f.Name, string.IsNullOrEmpty(f.ContainerClass) ? "IPF.Note" : f.ContainerClass); cur.FolderCount++; }
                    fmap[f] = pf;
                    return pf;
                };
                Action start = () =>
                {
                    wIndex++;
                    wPath = PstExporter.PartPath(dest, wIndex);
                    TryDelete(wPath);
                    if (string.IsNullOrEmpty(displayName))
                        using (PstReader r0 = PstReader.Open(files[0])) displayName = string.IsNullOrEmpty(r0.DisplayName) ? "MailboxTransfer copy" : r0.DisplayName;
                    w = PstWriter.Create(wPath, displayName);
                    written.Add(wPath);
                    fmap = new Dictionary<PstCopyFolder, PstFolder>();
                    cur = new PstExportPartInfo { Index = wIndex, File = Path.GetFileName(wPath), Path = wPath };
                    // the first file holds the whole tree, empty folders included
                    if (wIndex == 1) foreach (PstCopyFolder f in plan.Folders) ensure(f);
                };
                Action finish = () =>
                {
                    Report(progress, new PstExportProgress { Phase = "write", PartIndex = wIndex, PartPath = wPath, Message = "closing the PST", ItemsDone = itemsDone, ItemsTotal = plan.Items });
                    w.Close();
                    w = null;
                    cur.Bytes = new FileInfo(wPath).Length;
                    cur.ItemCount = cur.Items.Count;
                    cur.CompletedUtc = DateTime.UtcNow;
                    cur.Seconds = partClock.Elapsed.TotalSeconds;
                    cur.Verified = true;
                    if (o.Check)
                    {
                        var sw2 = Stopwatch.StartNew();
                        var co = new PstCheckOptions { Cancellation = ct };
                        int ps = cur.Index;
                        co.ProgressReport = new PstProgressSink(pp => Report(progress, new PstExportProgress { Phase = "check", PartIndex = ps, PartPath = wPath, BytesDone = pp.BytesDone, BytesTotal = pp.BytesTotal, ItemsDone = itemsDone, ItemsTotal = plan.Items }));
                        PstCheckReport rep = PstChecker.Check(cur.Path, co);
                        if (rep.Cancelled) throw new OperationCanceledException(ct);
                        cur.Checked = true; cur.CheckErrors = rep.Errors; cur.CheckWarnings = rep.Warnings; cur.Verified = rep.Errors == 0; cur.CheckSeconds = sw2.Elapsed.TotalSeconds;
                        if (rep.Errors > 0) res.PartsWithErrors++;
                    }
                    res.Parts.Add(cur);
                    res.BytesWritten += cur.Bytes;
                    cur = null;
                };

                foreach (PstCopyFolder f in plan.Folders)
                {
                    if (f.Nids == null || f.Nids.Count == 0) continue;
                    if (reader == null || readerFile != f.Source) { if (reader != null) reader.Dispose(); reader = PstReader.Open(f.Source); readerFile = f.Source; }
                    string key = f.Path.Length == 0 ? "(top)" : f.Path;
                    foreach (uint nid in f.Nids)
                    {
                        ct.ThrowIfCancellationRequested();
                        FxMessage m;
                        string subject = null;
                        try { m = reader.ReadMessage(nid); subject = m.GetString(0x0037); }
                        catch (Exception ex) when (!(ex is OperationCanceledException))
                        {
                            res.Failed.Add(new PstExportProblem { ItemKey = nid, Subject = subject, Folder = key, Reason = "cannot read the message: " + ex.Message });
                            if (o.Folders != null) o.Folders.Fail(key);
                            itemsDone++;
                            continue;
                        }
                        long est = PstWriter.EstimateSize(m);
                        if (w != null && w.MessagesWritten > 0 && w.ProjectedSize + est > o.MaxPartBytes) { finish(); partClock = null; }
                        if (w == null) { partClock = Stopwatch.StartNew(); start(); }
                        PstFolder target = ensure(f);
                        try { w.AddMessage(target, m, ct, sink); }
                        catch (OperationCanceledException) { throw; }
                        catch (IOException) { fatalIo = true; throw; }
                        catch (Exception ex)
                        {
                            res.Failed.Add(new PstExportProblem { ItemKey = nid, Subject = subject, Folder = key, Reason = "not written: " + ex.Message });
                            if (o.Folders != null) o.Folders.Fail(key);
                            itemsDone++;
                            continue;
                        }
                        cur.Items.Add(new PstExportItemRecord { Key = nid, Folder = 0, Length = 0 });
                        if (est > o.MaxPartBytes)
                        {
                            cur.Oversize = true;
                            res.Oversize.Add(new PstExportOversize { ItemKey = nid, Subject = subject, EstimatedBytes = est, Part = cur.Index });
                        }
                        itemsDone++;
                        res.ItemsWritten++;
                        if (o.Folders != null) o.Folders.Done(key, 0);
                        Report(progress, new PstExportProgress { Phase = "write", PartIndex = wIndex, PartPath = wPath, Folder = key, ItemsDone = itemsDone, ItemsTotal = plan.Items, PartBytes = w.ProjectedSize });
                    }
                }
                if (w == null && wIndex == 0) { partClock = Stopwatch.StartNew(); start(); }   // every message failed: the folders are still written
                if (w != null) { finish(); partClock = null; }
                Report(progress, new PstExportProgress { Phase = "done", ItemsDone = itemsDone, ItemsTotal = plan.Items });
            }
            catch (OperationCanceledException)
            {
                res.Cancelled = true;
                if (w != null) { try { w.Dispose(); } catch (Exception) { } w = null; }
                foreach (string f in written) TryDelete(f);
                res.Parts.Clear();
            }
            catch (Exception ex)
            {
                if (w != null) { try { w.Dispose(); } catch (Exception) { } w = null; }
                foreach (string f in written) TryDelete(f);
                res.Parts.Clear();
                res.Error = (fatalIo ? "I/O error: " : "") + ex.Message;
            }
            finally { if (reader != null) reader.Dispose(); }
            res.Elapsed = sw.Elapsed;
            return res;
        }
    }
}
