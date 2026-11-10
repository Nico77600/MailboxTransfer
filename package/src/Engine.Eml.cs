// =============================================================================
//  Mailbox Transfer - engine, part 15: EML export (live mailbox, export store, PST) and comparison of two MIME messages
// -----------------------------------------------------------------------------
//  EmlExporter.Run(source, options)   one pipeline for every source: folder tree -> directory tree, one .eml per message, a manifest
//                                     (manifest.jsonl, one JSON line per message: id, folder, file, subject, size, SHA-256, dates, state),
//                                     resumable (a message already exported with its file in place is not read again), streaming
//                                     (the message goes to a .part file, hashed on the way, then renamed), cancellable, -WhatIf.
//  Sources
//    LiveEmlSource    Mail API GET /users/{mailbox}/messages/{id}/$value (Primary and the main archive, delegated Mail.Read): Exchange's own MIME
//    StoreEmlSource   an export store (FastTransfer streams)  -> MapiToMime
//    PstEmlSource     a PST (any location, auxiliary archives included, exported to PST or store)  -> MapiToMime
//  Not exported (reported, never silently dropped): items that are not messages (appointments, contacts, tasks, notes...), folders whose
//  class is not IPF.Note, folders the Mail API cannot list. Limitation of v1: no .ics / .vcf conversion.
//  EmlCompare.Compare   semantic comparison of two messages (headers, structure, text bodies, attachment hashes).
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer
{
    // ---------------------------------------------------------------------------------------------------------------------
    //  File names
    // ---------------------------------------------------------------------------------------------------------------------

    public static class EmlNames
    {
        static readonly HashSet<string> Reserved = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "CON", "PRN", "AUX", "NUL", "COM0", "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT0", "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9", "CONIN$", "CONOUT$", "CLOCK$"
        };

        /// <summary>A file or directory name that Windows (and the other systems) accept: invalid characters, control characters, reserved device names (AUX, CON, NUL, COM1... also with an extension), trailing dots and spaces, length.</summary>
        public static string Component(string name, int max)
        {
            if (string.IsNullOrWhiteSpace(name)) return "_";
            var sb = new StringBuilder(name.Length);
            foreach (char c in name)
            {
                if (c < 32 || c == '<' || c == '>' || c == ':' || c == '"' || c == '/' || c == '\\' || c == '|' || c == '?' || c == '*' || c == '\u007f') sb.Append('_');
                else if (char.IsSurrogate(c)) sb.Append(c);
                else sb.Append(c);
            }
            string s = sb.ToString().Trim().TrimEnd('.', ' ');
            if (s.Length == 0) s = "_";
            if (s.Length > max)
            {
                s = s.Substring(0, max);
                if (char.IsHighSurrogate(s[s.Length - 1])) s = s.Substring(0, s.Length - 1);
                s = s.TrimEnd('.', ' ');
                if (s.Length == 0) s = "_";
            }
            int dot = s.IndexOf('.');
            string stem = dot < 0 ? s : s.Substring(0, dot);
            if (Reserved.Contains(stem.TrimEnd(' '))) s = "_" + s;
            return s;
        }

        public static string ShortHash(string text)
        {
            using (var sha = SHA1.Create())
            {
                byte[] h = sha.ComputeHash(Encoding.UTF8.GetBytes(text ?? ""));
                return BitConverter.ToString(h, 0, 4).Replace("-", "").ToLowerInvariant();
            }
        }

        /// <summary>
        /// Relative directory of every folder path ("Inbox/Projects"). Two folders whose names are the same once made safe and compared without
        /// case (Windows) both get a short hash of their path, so the result does not depend on the order of the walk. A path too long for
        /// <paramref name="budget"/> characters has its components shortened.
        /// </summary>
        public static Dictionary<string, string> FolderDirectories(IEnumerable<string> folderPaths, int budget)
        {
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            var paths = folderPaths.Distinct(StringComparer.Ordinal).ToList();
            int max = 80;
            while (true)
            {
                var siblings = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);   // parent raw path|safe lower name -> raw names
                var parts = new Dictionary<string, string[]>(StringComparer.Ordinal);
                foreach (string p in paths)
                {
                    string[] comps = p.Length == 0 ? new string[0] : p.Split('/');
                    parts[p] = comps;
                    for (int i = 0; i < comps.Length; i++)
                    {
                        string parent = string.Join("/", comps.Take(i));
                        string key = parent + "|" + Component(comps[i], max).ToLowerInvariant();
                        HashSet<string> set;
                        if (!siblings.TryGetValue(key, out set)) siblings[key] = set = new HashSet<string>(StringComparer.Ordinal);
                        set.Add(comps[i]);
                    }
                }
                bool tooLong = false;
                foreach (string p in paths)
                {
                    string[] comps = parts[p];
                    var safe = new List<string>();
                    for (int i = 0; i < comps.Length; i++)
                    {
                        string parent = string.Join("/", comps.Take(i));
                        string c = Component(comps[i], max);
                        if (siblings[parent + "|" + c.ToLowerInvariant()].Count > 1)
                        {
                            string h = ShortHash(parent + "/" + comps[i]);
                            c = Component(comps[i], Math.Max(1, max - 9)) + "~" + h;
                        }
                        safe.Add(c);
                    }
                    string dir = string.Join("\\", safe);
                    result[p] = dir;
                    if (dir.Length > budget) tooLong = true;
                }
                if (!tooLong || max <= 8) break;
                max = Math.Max(8, max / 2);
                result.Clear();
            }
            return result;
        }

        /// <summary>"20261028-220436 Subject [a1b2c3d4].eml" cut so that directory + name fit in maxPath.</summary>
        public static string FileName(DateTimeOffset? received, string subject, string id, int dirLength, int maxPath)
        {
            string stamp = received.HasValue ? received.Value.UtcDateTime.ToString("yyyyMMdd-HHmmss", CultureInfo.InvariantCulture) : "00000000-000000";
            string tail = " [" + ShortHash(id) + "].eml";
            int room = maxPath - dirLength - 1 - stamp.Length - 1 - tail.Length;
            int subjectMax = Math.Max(0, Math.Min(80, room));
            string subj = subjectMax == 0 ? "" : Component(subject ?? "", subjectMax);
            if (subject == null || subject.Trim().Length == 0) subj = subjectMax > 0 ? "(no subject)".Substring(0, Math.Min(12, subjectMax)) : "";
            return Component(stamp + (subj.Length > 0 ? " " + subj : "") + tail, 250).Replace("..", "._");
        }
    }

    // ---------------------------------------------------------------------------------------------------------------------
    //  Model of an export
    // ---------------------------------------------------------------------------------------------------------------------

    public sealed class EmlProgress
    {
        public string Stage, Folder, Item;
        public long Exported, Resumed, Skipped, Failed, Bytes;
        public long FolderItems = -1;
        public override string ToString() { return Stage + " " + Folder + " exported " + Exported + " resumed " + Resumed + " skipped " + Skipped + " failed " + Failed; }
    }

    /// <summary>Writes a progress line to the console at most every few seconds (callbacks come from worker threads: PowerShell script blocks cannot be used).</summary>
    public sealed class ConsoleProgress : IProgress<EmlProgress>, IProgress<EmlImportProgress>
    {
        readonly double _seconds;
        DateTime _last = DateTime.MinValue;
        readonly object _lock = new object();
        public ConsoleProgress(double seconds = 5) { _seconds = seconds; }
        void Line(string text)
        {
            lock (_lock)
            {
                if ((DateTime.UtcNow - _last).TotalSeconds < _seconds) return;
                _last = DateTime.UtcNow;
                Console.WriteLine("  " + text);
            }
        }
        public void Report(EmlProgress p)
        {
            Line(string.Format("{0}: {1} exported, {2} resumed, {3} skipped, {4} failed, {5:N1} MiB{6}", p.Folder ?? p.Stage, p.Exported, p.Resumed, p.Skipped, p.Failed, p.Bytes / 1048576.0, p.FolderItems >= 0 ? " (folder: " + p.FolderItems + " item(s))" : ""));
        }
        public void Report(EmlImportProgress p)
        {
            Line(string.Format("{0}/{1}: {2} imported, {3} resumed, {4} failed, {5} unknown", p.Done + p.Resumed + p.Failed + p.Unknown, p.Total, p.Done, p.Resumed, p.Failed, p.Unknown));
        }
    }

    public sealed class EmlExportOptions
    {
        public string OutputDirectory;
        public bool WhatIf;
        /// <summary>A message already in the manifest with its file in place is not read again (default).</summary>
        public bool Resume = true;
        /// <summary>Messages read at the same time (live source only; the service allows 4 concurrent requests per mailbox).</summary>
        public int Concurrency = 1;
        public int MaxPath = 240;
        /// <summary>Only the folders for which it returns true (full path "Inbox/Projects"); null: all.</summary>
        public Func<string, bool> FolderFilter;
        public MimeBuildOptions Mime = new MimeBuildOptions();
        public CancellationToken Cancellation;
        public IProgress<EmlProgress> Progress;
        /// <summary>Progress per folder for the console (items expected, exported, kept from an earlier run, skipped, failed); null: not followed.</summary>
        public FolderTracker Folders;
    }

    /// <summary>One line of manifest.jsonl.</summary>
    public sealed class EmlManifestEntry
    {
        public string State { get; set; }          // Exported, Skipped, Failed
        public string Id { get; set; }
        public string Source { get; set; }
        public string Folder { get; set; }
        public string File { get; set; }           // relative to the output directory
        public string Subject { get; set; }
        public string ItemClass { get; set; }
        public long Size { get; set; }             // bytes of the .eml
        public string Sha256 { get; set; }
        public string Received { get; set; }
        public bool? IsRead { get; set; }
        public string Reason { get; set; }
        public string[] Notes { get; set; }
        public string ExportedUtc { get; set; }
    }

    public sealed class EmlExportResult
    {
        public long Exported, Resumed, Skipped, Failed, Bytes;
        public int Folders, FoldersSkipped;
        public bool Cancelled, WhatIf;
        public string ManifestPath, OutputDirectory;
        public TimeSpan Elapsed;
        /// <summary>Skipped and failed items, folders skipped.</summary>
        public List<EmlManifestEntry> Problems = new List<EmlManifestEntry>();
        public List<string> FolderNotes = new List<string>();
        public bool Complete { get { return Failed == 0 && !Cancelled; } }
    }

    public sealed class EmlItemSource
    {
        public string Id, Subject, ItemClass;
        public long Size = -1;
        public DateTimeOffset? Received;
        public bool? IsRead;
        /// <summary>When set the item is reported and not exported.</summary>
        public string SkipReason;
        /// <summary>Writes the .eml; returns notes (what could not be carried over). Called once, never for a skipped item.</summary>
        public Func<Stream, CancellationToken, List<string>> Write;
    }

    public sealed class EmlFolderSource
    {
        public string Path;
        public long Count = -1;
        public string SkipReason;
        public IEnumerable<EmlItemSource> Items = new EmlItemSource[0];
    }

    public abstract class EmlSource
    {
        public abstract string Description { get; }
        /// <summary>A short label stored in the manifest.</summary>
        public abstract string Label { get; }
        public abstract List<EmlFolderSource> Folders(CancellationToken ct);
        public virtual bool AllowsParallel { get { return false; } }
    }

    // ---------------------------------------------------------------------------------------------------------------------
    //  Exporter
    // ---------------------------------------------------------------------------------------------------------------------

    sealed class HashingStream : Stream
    {
        readonly Stream _inner;
        readonly IncrementalHash _hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        public long Length2;
        public HashingStream(Stream inner) { _inner = inner; }
        public string Hex() { return BitConverter.ToString(_hash.GetHashAndReset()).Replace("-", ""); }
        public override bool CanRead { get { return false; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return true; } }
        public override long Length { get { return Length2; } }
        public override long Position { get { return Length2; } set { throw new NotSupportedException(); } }
        public override void Flush() { _inner.Flush(); }
        public override int Read(byte[] b, int o, int c) { throw new NotSupportedException(); }
        public override long Seek(long o, SeekOrigin so) { throw new NotSupportedException(); }
        public override void SetLength(long v) { throw new NotSupportedException(); }
        public override void Write(byte[] b, int o, int c) { _hash.AppendData(b, o, c); Length2 += c; _inner.Write(b, o, c); }
        protected override void Dispose(bool disposing) { if (disposing) _hash.Dispose(); base.Dispose(disposing); }
    }

    public static class EmlExporter
    {
        public const string ManifestName = "manifest.jsonl";
        static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions { DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull };

        public static List<EmlManifestEntry> ReadManifest(string directory)
        {
            var list = new List<EmlManifestEntry>();
            string path = Path.Combine(directory, ManifestName);
            if (!File.Exists(path)) return list;
            foreach (string line in File.ReadLines(path, Encoding.UTF8))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                try { var e = JsonSerializer.Deserialize<EmlManifestEntry>(line); if (e != null) list.Add(e); }
                catch (JsonException) { /* a line cut by a crash: the item is simply done again */ }
            }
            return list;
        }

        public static EmlExportResult Run(EmlSource source, EmlExportOptions o)
        {
            if (string.IsNullOrWhiteSpace(o.OutputDirectory)) throw new ArgumentException("OutputDirectory is required.");
            var started = System.Diagnostics.Stopwatch.StartNew();
            string root = Path.GetFullPath(o.OutputDirectory);
            var result = new EmlExportResult { OutputDirectory = root, WhatIf = o.WhatIf, ManifestPath = Path.Combine(root, ManifestName) };
            if (!o.WhatIf) Directory.CreateDirectory(root);

            var done = new Dictionary<string, EmlManifestEntry>(StringComparer.Ordinal);
            if (o.Resume && !o.WhatIf)
                foreach (EmlManifestEntry e in ReadManifest(root))
                    if (e.State == "Exported" && e.Id != null) done[e.Source + "|" + e.Id] = e;
            if (!o.WhatIf && Directory.Exists(root))
                foreach (string part in Directory.EnumerateFiles(root, "*.part", SearchOption.AllDirectories)) { try { File.Delete(part); } catch (IOException) { } }

            List<EmlFolderSource> folders = source.Folders(o.Cancellation);
            var wanted = folders.Where(f => o.FolderFilter == null || o.FolderFilter(f.Path)).ToList();
            Dictionary<string, string> dirs = EmlNames.FolderDirectories(wanted.Select(f => f.Path), Math.Max(20, o.MaxPath - root.Length - 60));
            var gate = new object();
            StreamWriter manifest = null;
            if (!o.WhatIf) manifest = new StreamWriter(new FileStream(result.ManifestPath, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false));
            try
            {
                foreach (EmlFolderSource folder in wanted)
                {
                    if (o.Cancellation.IsCancellationRequested) { result.Cancelled = true; break; }
                    if (folder.SkipReason != null)
                    {
                        result.FoldersSkipped++;
                        result.FolderNotes.Add(folder.Path + ": " + folder.SkipReason);
                        if (o.Folders != null) o.Folders.Mark(folder.Path, "refused", folder.SkipReason);
                        continue;
                    }
                    result.Folders++;
                    string dir = dirs[folder.Path];
                    string dirFull = dir.Length == 0 ? root : Path.Combine(root, dir);
                    if (!o.WhatIf) Directory.CreateDirectory(dirFull);
                    Report(o, "folder", folder.Path, null, result, folder.Count);
                    if (o.Folders != null) o.Folders.Expect(folder.Path, folder.Count);

                    Action<EmlItemSource> one = item =>
                    {
                        if (o.Cancellation.IsCancellationRequested) return;
                        var entry = new EmlManifestEntry
                        {
                            Id = item.Id, Source = source.Label, Folder = folder.Path, Subject = item.Subject, ItemClass = item.ItemClass,
                            Received = item.Received.HasValue ? item.Received.Value.UtcDateTime.ToString("o", CultureInfo.InvariantCulture) : null, IsRead = item.IsRead
                        };
                        EmlManifestEntry prior;
                        if (item.SkipReason != null)
                        {
                            entry.State = "Skipped"; entry.Reason = item.SkipReason;
                            lock (gate) { result.Skipped++; result.Problems.Add(entry); if (manifest != null) Append(manifest, entry); }
                            if (o.Folders != null) o.Folders.Skip(folder.Path);
                            return;
                        }
                        if (done.TryGetValue(source.Label + "|" + item.Id, out prior) && prior.File != null && File.Exists(Path.Combine(root, prior.File)))
                        {
                            lock (gate) { result.Resumed++; }
                            if (o.Folders != null) o.Folders.Done(folder.Path, prior.Size);
                            return;
                        }
                        string name = EmlNames.FileName(item.Received, item.Subject, item.Id, dirFull.Length, o.MaxPath);
                        string full = Path.Combine(dirFull, name);
                        entry.File = (dir.Length == 0 ? name : dir.Replace((char)92, (char)47) + "/" + name);
                        if (o.WhatIf)
                        {
                            lock (gate) { result.Exported++; result.Bytes += Math.Max(0, item.Size); }
                            if (o.Folders != null) o.Folders.Done(folder.Path, Math.Max(0, item.Size));
                            return;
                        }
                        string part = full + ".part";
                        try
                        {
                            List<string> notes;
                            string sha; long length;
                            using (var fs = new FileStream(part, FileMode.Create, FileAccess.Write, FileShare.None, 81920))
                            using (var hs = new HashingStream(fs))
                            {
                                notes = item.Write(hs, o.Cancellation) ?? new List<string>();
                                hs.Flush();
                                sha = hs.Hex(); length = hs.Length2;
                            }
                            if (File.Exists(full)) File.Delete(full);
                            File.Move(part, full);
                            entry.State = "Exported"; entry.Size = length; entry.Sha256 = sha; entry.Notes = notes.Count > 0 ? notes.ToArray() : null;
                            entry.ExportedUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
                            lock (gate) { result.Exported++; result.Bytes += length; Append(manifest, entry); }
                            if (o.Folders != null) o.Folders.Done(folder.Path, length);
                        }
                        catch (OperationCanceledException)
                        {
                            try { File.Delete(part); } catch (IOException) { }
                            lock (gate) { result.Cancelled = true; }
                        }
                        catch (Exception ex)
                        {
                            try { File.Delete(part); } catch (IOException) { }
                            entry.State = "Failed"; entry.File = null; entry.Reason = ex.GetType().Name + ": " + ex.Message;
                            lock (gate) { result.Failed++; result.Problems.Add(entry); Append(manifest, entry); }
                            if (o.Folders != null) o.Folders.Fail(folder.Path);
                        }
                        lock (gate)
                        {
                            if ((result.Exported + result.Resumed + result.Skipped + result.Failed) % 25 == 0) Report(o, "items", folder.Path, item.Subject, result, folder.Count);
                        }
                    };

                    int conc = source.AllowsParallel ? Math.Max(1, Math.Min(4, o.Concurrency)) : 1;
                    if (conc > 1)
                        {
                        var slots = new SemaphoreSlim(conc);
                        var tasks = new List<Task>();
                        foreach (EmlItemSource item in folder.Items)
                        {
                            if (o.Cancellation.IsCancellationRequested) break;
                            slots.Wait();
                            EmlItemSource captured = item;
                            tasks.Add(Task.Run(() => { try { one(captured); } finally { slots.Release(); } }));
                            tasks.RemoveAll(t => t.IsCompleted && !t.IsFaulted);
                        }
                        Task.WaitAll(tasks.ToArray());
                    }
                    else
                        foreach (EmlItemSource item in folder.Items) { one(item); if (o.Cancellation.IsCancellationRequested) break; }
                    if (o.Cancellation.IsCancellationRequested) result.Cancelled = true;
                    else if (o.Folders != null) o.Folders.Finish(folder.Path);
                    if (manifest != null) manifest.Flush();
                }
            }
            finally { if (manifest != null) manifest.Dispose(); }
            started.Stop();
            result.Elapsed = started.Elapsed;
            Report(o, "done", null, null, result, -1);
            return result;
        }

        static void Append(StreamWriter w, EmlManifestEntry e)
        {
            if (w == null) return;
            w.WriteLine(JsonSerializer.Serialize(e, JsonOpts));
        }

        static void Report(EmlExportOptions o, string stage, string folder, string item, EmlExportResult r, long count)
        {
            if (o.Progress == null) return;
            o.Progress.Report(new EmlProgress { Stage = stage, Folder = folder, Item = item, Exported = r.Exported, Resumed = r.Resumed, Skipped = r.Skipped, Failed = r.Failed, Bytes = r.Bytes, FolderItems = count });
        }
    }

    // ---------------------------------------------------------------------------------------------------------------------
    //  Sources
    // ---------------------------------------------------------------------------------------------------------------------

    /// <summary>A live mailbox through the Mail API: Primary or the main archive. The auxiliary archives are not reachable (export them to a store or a PST first).</summary>
    public sealed class LiveEmlSource : EmlSource
    {
        readonly MailboxBrowser _browser;
        readonly BrowseLocation _location;
        public bool IncludeNonMailFolders;

        public LiveEmlSource(MailboxBrowser browser, BrowseLocation location)
        {
            if (!location.HasMailApi)
                throw new InvalidOperationException("The Mail API does not reach an auxiliary archive: export it with -Mode Export (store) or to a PST, then export the EML from the store or the PST.");
            _browser = browser; _location = location;
        }

        public override string Description { get { return "mailbox " + _location.Kind + " " + _location.MailboxId + " (Mail API $value)"; } }
        public override string Label { get { return _location.MailboxId; } }
        public override bool AllowsParallel { get { return true; } }

        static bool IsMailFolderClass(string c)
        {
            return string.IsNullOrEmpty(c) || c.StartsWith("IPF.Note", StringComparison.OrdinalIgnoreCase);
        }

        public override List<EmlFolderSource> Folders(CancellationToken ct)
        {
            BrowseTree tree = _browser.GetFolderTree(_location, ct);
            var list = new List<EmlFolderSource>();
            foreach (BrowseFolder f in tree.Folders)
            {
                var fs = new EmlFolderSource { Path = f.Path, Count = f.TotalItemCount };
                if (f.Recoverable) fs.SkipReason = "Recoverable Items are not exported";
                else if (f.Source != MailboxBrowser.MailApi) fs.SkipReason = "folder only the import-export API lists" + (f.FolderClass != null ? " (" + f.FolderClass + ")" : "") + ": items are not messages (v1 does not convert them to .ics / .vcf)";
                else if (!IsMailFolderClass(f.FolderClass)) fs.SkipReason = "folder class " + f.FolderClass + " is not mail";
                else if (f.TotalItemCount == 0) fs.SkipReason = null;
                if (fs.SkipReason == null && f.TotalItemCount > 0) { BrowseFolder bf = f; fs.Items = Items(bf, ct); }
                list.Add(fs);
            }
            return list;
        }

        IEnumerable<EmlItemSource> Items(BrowseFolder folder, CancellationToken ct)
        {
            var q = new BrowseItemsQuery { PageSize = 100, Recipients = false, Sort = BrowseSort.ReceivedAscending };
            int listed = 0;
            foreach (BrowseItem row in _browser.EnumerateItems(_location, folder, q, ct))
            {
                BrowseItem r = row;
                listed++;
                var item = new EmlItemSource { Id = r.Id, Subject = r.Subject, ItemClass = r.ItemClass, Size = r.Size, Received = r.Received, IsRead = r.IsRead };
                if (!string.IsNullOrEmpty(r.ItemClass) && !MapiToMime.IsMailClass(r.ItemClass)) item.SkipReason = "item class " + r.ItemClass + " is not a message";
                else item.Write = (dest, token) => Fetch(r.Id, dest, token, r.HasAttachments);
                yield return item;
            }
            if (folder.State != BrowseFolderState.ContentNotAvailable && listed < folder.TotalItemCount)
                yield return new EmlItemSource
                {
                    Id = "non-messages:" + folder.Id, Subject = folder.Path, ItemClass = "(not messages)",
                    SkipReason = (folder.TotalItemCount - listed) + " item(s) of this folder are not messages (calendar items, contacts, tasks...): not exported (v1 does not convert them to .ics / .vcf)"
                };
            if (folder.State == BrowseFolderState.ContentNotAvailable)
                yield return new EmlItemSource { Id = "folder:" + folder.Id, Subject = folder.Path, SkipReason = "folder content not available now (" + folder.StateMessage + "): run again later" };
        }

        /// <summary>
        /// Copies the MIME. The Mail API row says whether the message has attachments: a MIME with attachment parts for a message that has none is the signature of an item whose
        /// attachments were lost in the mailbox while its MIME skeleton survived (the parts are empty); it is reported in the notes of the manifest.
        /// </summary>
        List<string> Fetch(string id, Stream dest, CancellationToken ct, bool? hasAttachments)
        {
            bool sawAttachment = false;
            byte[] needle = Encoding.ASCII.GetBytes("ontent-disposition: attachment");
            int matched = 0;
            GraphClient g = _browser.Graph;
            string url = g.Url("/users/" + _location.MailboxId + "/messages/" + Uri.EscapeDataString(id) + "/$value");
            using (HttpResponseMessage resp = g.Send("GET", url, completion: HttpCompletionOption.ResponseHeadersRead, ct: ct))
            using (Stream s = g.ReadBody(resp, ct))
            {
                var buf = new byte[81920];
                int n;
                while ((n = s.Read(buf, 0, buf.Length)) > 0)
                {
                    ct.ThrowIfCancellationRequested();
                    dest.Write(buf, 0, n);
                    if (!sawAttachment)
                        for (int i = 0; i < n; i++)
                        {
                            byte c = buf[i] >= 'A' && buf[i] <= 'Z' ? (byte)(buf[i] + 32) : buf[i];
                            if (c == needle[matched]) { if (++matched == needle.Length) { sawAttachment = true; break; } }
                            else matched = c == needle[0] ? 1 : 0;
                        }
                }
            }
            if (sawAttachment && hasAttachments == false)
                return new List<string> { "the Mail API says the message has no attachment but its MIME has attachment parts: the attachment content was lost in the mailbox (empty parts built from the MIME skeleton)" };
            return null;
        }
    }

    /// <summary>The common part of the sources built on FxMessage.</summary>
    public abstract class MapiEmlSource : EmlSource
    {
        public MimeBuildOptions Mime = new MimeBuildOptions();

        protected EmlItemSource ItemFor(string id, FxMessage m, long size)
        {
            string cls = m.GetString(Pid.MessageClass) ?? "IPM.Note";
            DateTimeOffset? recv = MapiToMime.Time(m.Properties, Pid.MessageDeliveryTime) ?? MapiToMime.Time(m.Properties, Pid.ClientSubmitTime);
            int? flags = MapiToMime.Int(m.Properties, Pid.MessageFlags);
            var item = new EmlItemSource
            {
                Id = id, Subject = m.GetString(Pid.Subject), ItemClass = cls, Size = size, Received = recv,
                IsRead = flags.HasValue ? (bool?)((flags.Value & 1) != 0) : null
            };
            if (!MapiToMime.IsMailClass(cls)) item.SkipReason = "item class " + cls + " is not a message (v1 does not convert it to .ics / .vcf)";
            else item.Write = (dest, ct) =>
            {
                var notes = new List<string>();
                MapiToMime.Write(m, dest, Mime, notes, ct);
                return notes;
            };
            return item;
        }
    }

    public sealed class StoreEmlSource : MapiEmlSource
    {
        readonly ExportStore _store;
        readonly string _mailboxId;

        public StoreEmlSource(ExportStore store, string mailboxId = null) { _store = store; _mailboxId = string.IsNullOrEmpty(mailboxId) ? null : mailboxId; }
        public override string Description { get { return "export store" + (_mailboxId != null ? " (mailbox " + _mailboxId + ")" : "") + " (FastTransfer streams -> MIME)"; } }
        public override string Label { get { return _mailboxId ?? "store"; } }

        public override List<EmlFolderSource> Folders(CancellationToken ct)
        {
            var list = new List<EmlFolderSource>();
            IEnumerable<string> ids = _mailboxId != null ? new[] { _mailboxId } : _store.GetMailboxes().Select(x => x.Id).ToArray();
            bool several = _mailboxId == null && ids.Count() > 1;
            foreach (string mb in ids)
            {
                var stack = new Stack<StoreFolderNode>();
                foreach (StoreFolderNode r in _store.GetFolderTree(mb).AsEnumerable().Reverse()) stack.Push(r);
                while (stack.Count > 0)
                {
                    StoreFolderNode node = stack.Pop();
                    foreach (StoreFolderNode c in node.Children.AsEnumerable().Reverse()) stack.Push(c);
                    long key = node.FolderKey;
                    string path = several ? mb.Replace(':', '_') + "/" + node.Path : node.Path;
                    var fs = new EmlFolderSource { Path = path, Count = node.Own == null ? -1 : node.Own.Count };
                    if (!string.IsNullOrEmpty(node.FolderType) && !node.FolderType.StartsWith("IPF.Note", StringComparison.OrdinalIgnoreCase) && node.FolderType != "Mail")
                    {
                        if (fs.Count != 0) fs.SkipReason = "folder type " + node.FolderType + " is not mail";
                    }
                    else fs.Items = Items(key, ct);
                    list.Add(fs);
                }
            }
            return list;
        }

        IEnumerable<EmlItemSource> Items(long folderKey, CancellationToken ct)
        {
            int page = 0;
            while (true)
            {
                StoreItemPage p = _store.GetItems(folderKey, page, 200, StoreItemSort.ReceivedDate, false, StoreItemState.Exported);
                foreach (StoreItem si in p.Items)
                {
                    ct.ThrowIfCancellationRequested();
                    FxDocument doc = null; string error = null;
                    try { doc = FxCodec.Decode(FxStreamSource.FromStoreItem(_store, si.Key)); }
                    catch (Exception ex) { error = ex.Message; }
                    if (doc == null || doc.Error != null)
                    {
                        yield return new EmlItemSource { Id = si.ItemId, Subject = si.Subject, Received = si.ReceivedUtc.HasValue ? (DateTimeOffset?)new DateTimeOffset(DateTime.SpecifyKind(si.ReceivedUtc.Value, DateTimeKind.Utc)) : null,
                            SkipReason = "stream cannot be decoded: " + (error ?? doc.Error) };
                        continue;
                    }
                    yield return ItemFor(si.ItemId, doc.Message, si.Size);
                }
                page++;
                if (page >= p.PageCount) break;
            }
        }
    }

    public sealed class PstEmlSource : MapiEmlSource
    {
        readonly PstReader _reader;
        readonly string _name;
        public PstEmlSource(PstReader reader, string name = null) { _reader = reader; _name = name ?? "pst"; }
        public override string Description { get { return "PST " + _reader.File.Path + " (MAPI -> MIME)"; } }
        public override string Label { get { return _name; } }

        public override List<EmlFolderSource> Folders(CancellationToken ct)
        {
            var list = new List<EmlFolderSource>();
            foreach (PstReaderFolder f in _reader.AllFolders())
            {
                if (f.ParentNid == 0 && f.Path == null) continue;
                string path = f.Path ?? f.Name;
                if (string.IsNullOrEmpty(path)) continue;
                var fs = new EmlFolderSource { Path = path, Count = f.ContentCount };
                if (!string.IsNullOrEmpty(f.ContainerClass) && !f.ContainerClass.StartsWith("IPF.Note", StringComparison.OrdinalIgnoreCase))
                { if (f.ContentCount > 0) fs.SkipReason = "folder class " + f.ContainerClass + " is not mail"; }
                else fs.Items = Items(f, ct);
                list.Add(fs);
            }
            return list;
        }

        IEnumerable<EmlItemSource> Items(PstReaderFolder f, CancellationToken ct)
        {
            foreach (uint nid in _reader.MessageNids(f, false))
            {
                ct.ThrowIfCancellationRequested();
                FxMessage m = null; string error = null;
                try { m = _reader.ReadMessage(nid); } catch (Exception ex) { error = ex.Message; }
                if (m == null) { yield return new EmlItemSource { Id = nid.ToString("X"), SkipReason = "message cannot be read: " + error }; continue; }
                yield return ItemFor(nid.ToString("X"), m, -1);
            }
        }
    }

    // ---------------------------------------------------------------------------------------------------------------------
    //  Comparison of two messages
    // ---------------------------------------------------------------------------------------------------------------------

    public sealed class EmlDifference
    {
        /// <summary>Header, Structure, Body, Attachment.</summary>
        public string Kind;
        public string Name;
        /// <summary>Same (equivalent), Changed, OnlyLeft, OnlyRight.</summary>
        public string Result;
        public string Left, Right;
        public override string ToString() { return Kind + " " + Name + ": " + Result + (Left != null ? " | " + Left : "") + (Right != null ? " | " + Right : ""); }
    }

    public sealed class EmlComparison
    {
        public List<EmlDifference> Items = new List<EmlDifference>();
        public string LeftStructure, RightStructure;
        public IEnumerable<EmlDifference> Differences { get { return Items.Where(i => i.Result != "Same"); } }
        public int Count(string kind, string result) { return Items.Count(i => i.Kind == kind && i.Result == result); }
        /// <summary>No difference in the text bodies and the attachments (the headers and the structure may differ).</summary>
        public bool ContentEqual { get { return !Items.Any(i => (i.Kind == "Body" || i.Kind == "Attachment") && i.Result != "Same"); } }
        public bool Identical { get { return !Differences.Any(); } }
    }

    public static class EmlCompare
    {
        public static EmlComparison Compare(byte[] left, byte[] right)
        {
            return Compare(MimeParser.Parse(left), MimeParser.Parse(right));
        }

        static string Norm(string v) { return Regex.Replace(MimeText.DecodeWords(v ?? ""), @"\s+", " ").Trim(); }

        public static EmlComparison Compare(MimeEntity a, MimeEntity b)
        {
            var r = new EmlComparison();
            var names = new List<string>();
            foreach (MimeHeader h in a.Headers.Concat(b.Headers)) if (!names.Contains(h.Name, StringComparer.OrdinalIgnoreCase)) names.Add(h.Name);
            foreach (string n in names)
            {
                if (n.StartsWith("Content-", StringComparison.OrdinalIgnoreCase) || n.Equals("MIME-Version", StringComparison.OrdinalIgnoreCase)) continue;
                List<string> la = a.HeaderAll(n).Select(Norm).ToList(), lb = b.HeaderAll(n).Select(Norm).ToList();
                if (la.Count == 0) r.Items.Add(new EmlDifference { Kind = "Header", Name = n, Result = "OnlyRight", Right = Short(lb[0]) });
                else if (lb.Count == 0) r.Items.Add(new EmlDifference { Kind = "Header", Name = n, Result = "OnlyLeft", Left = Short(la[0]) });
                else if (la.SequenceEqual(lb, StringComparer.Ordinal) || SameAddresses(n, la, lb) || SameDate(n, la, lb)) r.Items.Add(new EmlDifference { Kind = "Header", Name = n, Result = "Same" });
                else r.Items.Add(new EmlDifference { Kind = "Header", Name = n, Result = "Changed", Left = Short(string.Join(" || ", la)), Right = Short(string.Join(" || ", lb)) });
            }
            r.LeftStructure = Structure(a); r.RightStructure = Structure(b);
            r.Items.Add(new EmlDifference { Kind = "Structure", Name = "tree", Result = r.LeftStructure == r.RightStructure ? "Same" : "Changed", Left = r.LeftStructure, Right = r.RightStructure });

            List<MimeEntity> la2 = Leaves(a), lb2 = Leaves(b);
            var used = new HashSet<int>();
            foreach (MimeEntity x in la2)
            {
                int idx = -1;
                for (int i = 0; i < lb2.Count; i++)
                {
                    if (used.Contains(i)) continue;
                    if (Key(lb2[i]) == Key(x)) { idx = i; break; }
                }
                string kind = IsBody(x) ? "Body" : "Attachment";
                if (idx < 0)
                {
                    // same kind of part with another name?
                    for (int i = 0; i < lb2.Count; i++) if (!used.Contains(i) && IsBody(lb2[i]) == IsBody(x) && lb2[i].ContentType.MediaType == x.ContentType.MediaType) { idx = i; break; }
                }
                if (idx < 0) { r.Items.Add(new EmlDifference { Kind = kind, Name = Key(x), Result = "OnlyLeft", Left = Describe(x) }); continue; }
                used.Add(idx);
                MimeEntity y = lb2[idx];
                string da = Digest(x), db = Digest(y);
                bool same = da == db;
                string nm = Key(x) + (Key(x) != Key(y) ? " / " + Key(y) : "");
                r.Items.Add(new EmlDifference { Kind = kind, Name = nm, Result = same ? "Same" : "Changed", Left = same ? null : Describe(x), Right = same ? null : Describe(y) });
            }
            for (int i = 0; i < lb2.Count; i++)
                if (!used.Contains(i)) r.Items.Add(new EmlDifference { Kind = IsBody(lb2[i]) ? "Body" : "Attachment", Name = Key(lb2[i]), Result = "OnlyRight", Right = Describe(lb2[i]) });
            return r;
        }

        static string Short(string s) { return s.Length > 160 ? s.Substring(0, 157) + "..." : s; }

        static bool SameAddresses(string name, List<string> a, List<string> b)
        {
            string n = name.ToLowerInvariant();
            if (n != "from" && n != "to" && n != "cc" && n != "bcc" && n != "reply-to" && n != "sender") return false;
            Func<List<string>, string> f = l => string.Join(";", l.SelectMany(x => MimeText.ParseAddressList(x)).Select(x => (x.Address ?? "").ToLowerInvariant()));
            return f(a) == f(b);
        }

        static bool SameDate(string name, List<string> a, List<string> b)
        {
            if (!name.Equals("Date", StringComparison.OrdinalIgnoreCase) || a.Count != 1 || b.Count != 1) return false;
            DateTimeOffset? x = MimeText.ParseDate(a[0]), y = MimeText.ParseDate(b[0]);
            return x.HasValue && y.HasValue && x.Value.UtcDateTime == y.Value.UtcDateTime;
        }

        public static string Structure(MimeEntity e)
        {
            if (e.IsMultipart) return e.ContentType.SubType + "(" + string.Join(",", e.Parts.Select(Structure)) + ")";
            if (e.Message != null) return "rfc822[" + Structure(e.Message) + "]";
            return e.ContentType.SubType == "plain" || e.ContentType.SubType == "html" ? e.ContentType.SubType : e.ContentType.MediaType;
        }

        static List<MimeEntity> Leaves(MimeEntity e)
        {
            var l = new List<MimeEntity>();
            if (e.IsMultipart) foreach (MimeEntity p in e.Parts) l.AddRange(Leaves(p));
            else l.Add(e);
            return l;
        }

        static bool IsBody(MimeEntity e)
        {
            return e.Message == null && e.ContentType.Type == "text" && e.FileName == null && e.Disposition != "attachment" && (e.ContentType.SubType == "plain" || e.ContentType.SubType == "html");
        }

        static string Key(MimeEntity e)
        {
            if (IsBody(e)) return "text/" + e.ContentType.SubType;
            return (e.FileName ?? e.ContentId ?? (e.Message != null ? "message" : e.ContentType.MediaType)) + (e.ContentId != null ? " cid:" + e.ContentId : "");
        }

        static string Describe(MimeEntity e)
        {
            return e.ContentType.MediaType + (e.Message != null ? " (message)" : " " + e.ContentLength + " bytes");
        }

        static string Digest(MimeEntity e)
        {
            if (e.Message != null) return "msg:" + Digest2(Encoding.UTF8.GetBytes(string.Join(",", Leaves(e.Message).Select(Digest))));
            if (IsBody(e))
            {
                string t = e.Text().Replace("\r\n", "\n").Replace("\r", "\n");
                if (e.ContentType.SubType == "html")
                {
                    // Exchange re-serializes HTML (attribute quotes, tag case, line breaks): compared without them
                    t = Regex.Replace(t, @"(?is)<head\b.*?</head>", "");
                    t = Regex.Replace(t.ToLowerInvariant().Replace("\"", "").Replace("'", ""), @"\s+", "");
                }
                else t = Regex.Replace(t, @"\s+", " ");   // Exchange re-wraps plain text
                return "txt:" + Digest2(Encoding.UTF8.GetBytes(t.Trim()));
            }
            return "bin:" + Digest2(e.Content == null ? new byte[0] : e.Content.ToArray());
        }

        static string Digest2(byte[] b)
        {
            using (var sha = SHA256.Create()) return BitConverter.ToString(sha.ComputeHash(b)).Replace("-", "");
        }
    }
}
