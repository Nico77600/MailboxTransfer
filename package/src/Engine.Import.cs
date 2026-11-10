// =============================================================================
//  Mailbox Transfer - engine, part 8: import of the items of a store into a mailbox
// -----------------------------------------------------------------------------
//  Context: the owner's own tooling to migrate or back up Exchange Online mailboxes the operator is authorised to access
//  (delegated consent of the mailbox user). Creation only: nothing is ever replaced, moved or deleted in the target.
//
//  Phases of an ImportJob:
//    folders   the folders of the source (paths of the store) are mapped to target paths (below TargetRoot, or at the same path,
//              or through an explicit mapping by path prefix) and the missing ones are created, parents first. A folder is found
//              by its name under its parent and never renamed, moved or deleted, so a second run creates nothing.
//    plan      every Exported item of the source gets a row in the import job of the store (INSERT OR IGNORE: idempotent).
//    import    workers take the Pending rows one by one (Pending -> InFlight), import them with one session per target mailbox
//              (reused until it is about to expire) and mark them Done, Failed or Unknown:
//                Failed   the service refused the request (4xx, 429, 503) or never received the whole body (nothing handed to the
//                         transport, or an upload interrupted; sent again in place first): nothing was created, a new run tries again;
//                Unknown  timeout, network error after the whole body was sent, 5xx, interruption: the item may exist. Never replayed by itself.
//  A crash leaves rows InFlight; the store turns them into Unknown when it is opened again.
//  Automatic reconciliation (on by default, ImportOptions.AutoReconcile): an Unknown row does not stop the run. After a settle delay
//  (the service may still be finishing the creation) the target folder is listed and the item is looked for (ImportReconciler):
//    one candidate     the item was created: the row is Done with its id;
//    no candidate      nothing was created: the row goes back to Pending and the item is imported again in the same run (one more attempt,
//                      MaxAttempts respected; ReconcileRounds lookups at most per row);
//    several candidates  genuinely ambiguous: the row stays Unknown and is reported; only this case (or the rounds used up) needs a person.
//  The Unknown rows left by an earlier run (crash, kill) are reconciled the same way at the start of the run, once their delay is over.
//  The result of a run counts what THIS run did; JobTotal / JobDone / JobFailed / JobUnknown / JobPending count the whole job (all runs
//  together, Unknown rows left by an earlier run included): the job is complete only when JobComplete is true.
//
//  Auto-expanding archive: a folder of the main archive can live in an auxiliary archive. The service then answers importItem 409 "the target
//  session is expected in mailbox X" (and folder requests 308 to X). The import records X as an AuxArchive location, creates (or reuses) an import
//  session for X and asks again there, with the folder id as X lists it: the documentation shows the same folder id in X, and when X does not know
//  that id the folder is found in the tree of X by its path (same names under the same parent path). A folder is never created in an auxiliary
//  archive by this routing: the folder of the main archive is the logical owner, a second tree in X would diverge from it; an item whose folder
//  cannot be found in X is Failed with nothing created. What was learned (folder -> mailbox, folder id) is kept in the store for the job, the row
//  targets the mailbox that serves it (so the reconciliation of an Unknown row looks there) and the next items of the folder go there directly.
//
//  Throttling is obeyed, not worked around (see Engine.Transfer.cs): Retry-After exactly, concurrency halved on 429 / 503, 4 at most.
// =============================================================================
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;

namespace MailboxTransfer
{
    public sealed class ImportOptions
    {
        public string SourceMailboxId;
        public string TargetMailboxId;
        /// <summary>Path in the target below which everything is created ("MBT-Phase1-A" or "A/B"). Null: the same paths as the source.</summary>
        public string TargetRoot;
        /// <summary>Source path prefix -> target path prefix (relative to TargetRoot). The longest prefix wins.</summary>
        public IDictionary<string, string> Mapping;
        public int Concurrency = TransferLimits.MaxConcurrency;
        /// <summary>A run with the name of an existing job resumes it.</summary>
        public string JobName = "import";
        /// <summary>A Failed row goes back to Pending at the start of a run while it was tried fewer times than this.</summary>
        public int MaxAttempts = 3;
        /// <summary>Only the source folder of this path and the folders below it. Null: all.</summary>
        public string SourceFolderPath;
        /// <summary>An Unknown row is looked up in the target folder and settled by the run itself (see the header). False: Unknown rows stay Unknown until a person settles them.</summary>
        public bool AutoReconcile = true;
        /// <summary>Seconds between the end of the request that left a row Unknown and the lookup (the service may still be creating the item).</summary>
        public double ReconcileDelaySeconds = 300;
        /// <summary>How many times a row may be looked up (and imported again when nothing was found) before it is left Unknown.</summary>
        public int ReconcileRounds = 3;
    }

    public sealed class ImportFolderPlan
    {
        public long SourceFolderKey;
        public string SourcePath, SourceType, TargetPath, TargetFolderId;
        public bool Exists;
        public long Items, Bytes;
        /// <summary>Items of the folder that are in the store but not Exported (Pending or Failed there): the import does not take them.</summary>
        public long NotExported;
    }

    public sealed class ImportPlan
    {
        public List<ImportFolderPlan> Folders = new List<ImportFolderPlan>();
        /// <summary>Target paths that do not exist yet (parents before children).</summary>
        public List<string> PathsToCreate = new List<string>();
        public int FoldersExisting;
        public long Items, Bytes;
        /// <summary>Items of the store in the folders of the plan that are not Exported yet (Pending or Failed in the store).</summary>
        public long NotExported;
    }

    public sealed class ImportRunResult
    {
        /// <summary>What THIS run did.</summary>
        public long JobId, Done, Failed, Unknown, Bytes;
        public int FoldersCreated, FoldersFailed, SessionsCreated;
        /// <summary>Folders of the target that live in an auxiliary archive (known to the job), and those mailboxes.</summary>
        public int FoldersInAux;
        public List<string> AuxArchives = new List<string>();
        /// <summary>The whole job (all runs together) at the end of this run. JobPending includes the rows InFlight.</summary>
        public string JobName;
        public bool JobKnown;
        public long JobTotal, JobDone, JobFailed, JobUnknown, JobPending;
        /// <summary>Unknown rows of the job that this run did not produce (an earlier run, a crash).</summary>
        public long JobUnknownEarlier;
        /// <summary>Rows of the job that were InFlight when an earlier run stopped and became Unknown when the store was opened.</summary>
        public int RecoveredToUnknown;
        /// <summary>Automatic reconciliation (this run): rows found created in the target (now Done), rows imported again because nothing was found,
        /// rows left Unknown because several items could be them, rows left Unknown after ReconcileRounds lookups.</summary>
        public bool AutoReconcile;
        public long ReconciledCreated, ReconciledReimported, ReconcileAmbiguous, ReconcileExhausted;
        public List<double> SessionLifetimeSeconds = new List<double>();
        public List<string> Problems = new List<string>();
        public TransferSnapshot Final;
        public bool Completed;

        /// <summary>
        /// Items this run left out of the job: the items of a folder that could not be created in the target, and the items of the store that are not Exported yet.
        /// They are not rows of the journal, so they are counted here: the job is not complete while there are some (run again once the folder exists / the export is finished).
        /// </summary>
        public long NotQueued;

        /// <summary>Every row of the job is Done and nothing was left out (NotQueued). A run that completed (not cancelled) can still leave a job incomplete.</summary>
        public bool JobComplete { get { return JobKnown && JobDone == JobTotal && NotQueued == 0; } }

        /// <summary>One sentence about the automatic reconciliation of this run, or null when it did nothing.</summary>
        public string ReconcileSummary()
        {
            if (!AutoReconcile) return null;
            if (ReconciledCreated + ReconciledReimported + ReconcileAmbiguous + ReconcileExhausted == 0) return null;
            return "reconciliation: " + ReconciledCreated + " Unknown row(s) found created in the target (marked Done), " + ReconciledReimported +
                " found absent and imported again, " + ReconcileAmbiguous + " left Unknown because several items could be them" +
                (ReconcileExhausted > 0 ? ", " + ReconcileExhausted + " left Unknown after the allowed rounds" : "");
        }

        /// <summary>One sentence about the job as a whole, for the end of a run.</summary>
        public string JobSummary()
        {
            string name = "job '" + JobName + "'";
            if (!JobKnown) return name + " state not read: the run stopped before the job was opened";
            if (JobComplete) return name + " complete: " + JobDone + " of " + JobTotal + " done";
            var parts = new List<string>();
            if (JobUnknown > 0)
                parts.Add(JobUnknown + " unknown (" + (JobUnknownEarlier > 0 ? JobUnknownEarlier + " from an earlier run, " : "") + "the item may exist, never replayed by itself)");
            if (JobFailed > 0) parts.Add(JobFailed + " failed (nothing created, the next run tries again while attempts remain)");
            if (JobPending > 0) parts.Add(JobPending + " pending");
            if (NotQueued > 0) parts.Add(NotQueued + " left out of the job (the target folder could not be created, or the items are not exported yet in the store: run again when that is settled)");
            return name + " incomplete: " + JobDone + " of " + JobTotal + " done, " + string.Join(", ", parts);
        }
    }

    /// <summary>
    /// What an import knows about the folders that live in an auxiliary archive: (planned mailbox, folder) -> (mailbox that serves it, folder id as
    /// that mailbox lists it). Shared by the workers of a job; kept in the store (table importRoute) so that a resumed run starts with it.
    /// </summary>
    public sealed class ImportRouter
    {
        readonly MailboxApi _api;
        readonly ExportStore _store;
        readonly ConcurrentDictionary<string, KeyValuePair<string, string>> _routes = new ConcurrentDictionary<string, KeyValuePair<string, string>>(StringComparer.OrdinalIgnoreCase);
        readonly ConcurrentDictionary<string, string> _paths = new ConcurrentDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        readonly Dictionary<string, List<FolderInfo>> _trees = new Dictionary<string, List<FolderInfo>>(StringComparer.OrdinalIgnoreCase);
        readonly object _treeLock = new object();

        public long JobId;
        public AuxLocationLearner Learner;

        public ImportRouter(MailboxApi api, ExportStore store) { _api = api; _store = store; }

        static string Key(string mailbox, string folder) { return mailbox + "|" + folder; }

        public int Count { get { return _routes.Count; } }
        public List<string> Mailboxes { get { return _routes.Values.Select(v => v.Key).Distinct(StringComparer.OrdinalIgnoreCase).ToList(); } }

        public void Load(long jobId)
        {
            JobId = jobId;
            foreach (var kv in _store.GetImportRoutes(jobId)) _routes[kv.Key] = kv.Value;
        }

        /// <summary>The path of a folder of the target as the plan knows it (used to find the folder in an auxiliary archive by its names).</summary>
        public void SetPath(string mailbox, string folderId, string path) { _paths[Key(mailbox, folderId)] = path; }

        public string PathOf(string mailbox, string folderId) { string p; return _paths.TryGetValue(Key(mailbox, folderId), out p) ? p : null; }

        public bool TryRoute(string mailbox, string folderId, out string toMailbox, out string toFolder)
        {
            KeyValuePair<string, string> v;
            if (_routes.TryGetValue(Key(mailbox, folderId), out v)) { toMailbox = v.Key; toFolder = v.Value; return true; }
            toMailbox = null; toFolder = null;
            return false;
        }

        public void Learn(string mailbox, string folderId, string toMailbox, string toFolder)
        {
            var v = new KeyValuePair<string, string>(toMailbox, toFolder);
            KeyValuePair<string, string> old;
            if (_routes.TryGetValue(Key(mailbox, folderId), out old) && old.Key == toMailbox && old.Value == toFolder) return;
            _routes[Key(mailbox, folderId)] = v;
            if (JobId > 0) _store.SaveImportRoute(JobId, mailbox, folderId, toMailbox, toFolder);
        }

        /// <summary>The id, in <paramref name="mailbox"/>, of the folder with this path (same names under the same parent path, case ignored), or null. The tree of a mailbox is read once.</summary>
        public string FindFolderByPath(string mailbox, string path, CancellationToken ct)
        {
            List<FolderInfo> tree;
            lock (_treeLock)
            {
                if (!_trees.TryGetValue(mailbox, out tree)) { tree = _api.GetFolderTree(mailbox, ct); _trees[mailbox] = tree; }
            }
            FolderInfo f = tree.FirstOrDefault(x => string.Equals(x.Path, path, StringComparison.OrdinalIgnoreCase));
            return f == null ? null : f.Id;
        }
    }

    /// <summary>Shared by the import and by the replication of the lab: one item, one session, the refusals that can be repaired.</summary>
    public static class ImportHelper
    {
        static readonly Regex ExpectedMailbox = new Regex(@"expected in mailbox\s+(?<id>MBX:[0-9A-Fa-f-]+@[0-9A-Fa-f-]+)", RegexOptions.Compiled);

        /// <summary>The mailbox named by a 409 "Invalid import session. The target session is expected in mailbox MBX:..." or null.</summary>
        public static string ExpectedMailboxOf(GraphException ex)
        {
            if (ex == null || ex.StatusCode != 409 || ex.Message == null) return null;
            Match m = ExpectedMailbox.Match(ex.Message);
            return m.Success ? m.Groups["id"].Value : null;
        }

        /// <summary>Imports an item of the store without auxiliary archive routing beyond the 409 retry (see the other overload).</summary>
        public static ImportResult ImportWithSession(MailboxApi api, ImportSessionPool pool, ExportStore store, string mailboxId, string folderId,
            long itemKey, long size, CancellationToken ct, Action<string> log)
        {
            return ImportWithSession(api, pool, store, mailboxId, folderId, itemKey, size, ct, log, null, null);
        }

        /// <summary>
        /// Imports an item of the store. 409 "the target session is expected in mailbox X" (the folder lives in an auxiliary archive
        /// of an auto-expanding archive): X is recorded as an AuxArchive location, a session is created (or reused) for X and the item is asked
        /// again there with the folder id as X lists it: the same id (as the documentation shows), else the folder of the same path in the tree of X.
        /// A folder is never created in X. 401: the token of the session expired before its time, a new session replaces it.
        /// All are refusals: nothing was created. <paramref name="onRoute"/> is called with the (mailbox, folder) the item is about to be sent to
        /// when it is not the planned one (the caller updates its row before the request, so that an Unknown outcome is looked up there).
        /// </summary>
        public static ImportResult ImportWithSession(MailboxApi api, ImportSessionPool pool, ExportStore store, string mailboxId, string folderId,
            long itemKey, long size, CancellationToken ct, Action<string> log, ImportRouter router, Action<string, string> onRoute)
        {
            return ImportWithSession(api, pool, () => store.OpenItem(itemKey), mailboxId, folderId, size, ct, log, router, onRoute);
        }

        /// <summary>The same for any source of the stream (<paramref name="open"/> gives a new stream at position 0 each time the body is written): the items of a PST.</summary>
        public static ImportResult ImportWithSession(MailboxApi api, ImportSessionPool pool, Func<Stream> open, string mailboxId, string folderId,
            long size, CancellationToken ct, Action<string> log, ImportRouter router, Action<string, string> onRoute)
        {
            string mailbox = mailboxId, folder = folderId;
            string routedMailbox, routedFolder;
            if (router != null && router.TryRoute(mailboxId, folderId, out routedMailbox, out routedFolder))
            {
                mailbox = routedMailbox; folder = routedFolder;
                if (onRoute != null) onRoute(mailbox, folder);
            }
            ImportSession session = pool.Get(mailbox, ct);
            int redirects = 0, renewals = 0;
            bool pathTried = false;
            while (true)
            {
                try
                {
                    ImportResult done = api.ImportItem(session, folder, open, size, ct);
                    if (router != null && (!MailboxApi.SameMailbox(mailbox, mailboxId) || folder != folderId)) router.Learn(mailboxId, folderId, mailbox, folder);
                    return done;
                }
                catch (GraphException ex) when (ExpectedMailboxOf(ex) != null)
                {
                    string other = ExpectedMailboxOf(ex);
                    if (redirects >= 1 || MailboxApi.SameMailbox(other, mailbox))
                        throw new GraphException(409, "AuxArchiveRefused",
                            "The folder " + folder + " lives in an auxiliary archive (auto-expanding archive) and the service still refuses the import after a session was created for " +
                            other + " (MailboxLocationType AuxArchive in Get-MailboxLocation). The folder id must be the one that mailbox lists. Nothing was created. " + ex.Message, ex.Url);
                    redirects++;
                    if (log != null) log("the folder lives in the auxiliary archive " + other + ": creating an import session for it");
                    if (router != null && router.Learner != null) router.Learner.Learn(other, "Import409", "importItem into folder " + folderId + " of " + mailbox + " answered 409 expecting this mailbox");
                    mailbox = other;
                    session = pool.Get(mailbox, ct);
                    if (onRoute != null) onRoute(mailbox, folder);
                }
                catch (GraphException ex) when (router != null && redirects > 0 && !pathTried && (ex.StatusCode == 404 || ex.StatusCode == 400))
                {
                    // the auxiliary archive does not know the folder id of the main archive: look for the folder of the same path there. Never create it.
                    pathTried = true;
                    string path = router.PathOf(mailboxId, folderId);
                    string id = path == null ? null : router.FindFolderByPath(mailbox, path, ct);
                    if (id == null || id == folder)
                        throw new GraphException(404, "AuxFolderNotFound",
                            "The folder " + (path ?? folderId) + " lives in the auxiliary archive " + mailbox + " but that mailbox does not list a folder " + (id == null ? "of that path" : "with that id") +
                            "; folders are not created in an auxiliary archive. Nothing was created. " + ex.Message, ex.Url);
                    if (log != null) log("the auxiliary archive " + mailbox + " lists the folder '" + path + "' with another id: " + id);
                    folder = id;
                    if (onRoute != null) onRoute(mailbox, folder);
                }
                catch (GraphException ex) when (ex.StatusCode == 401 && renewals < 2)
                {
                    renewals++;
                    if (log != null) log("the import session was refused (401), creating a new one");
                    session = pool.Renew(mailbox, session, ct);
                }
            }
        }

        /// <summary>
        /// The request was refused or never fully sent, so nothing was created: 4xx, 429 and 503 after the retries, and an upload interrupted
        /// before the last byte of the body (UploadInterrupted) or before its first byte (NotSent). Anything else may have created the item.
        /// </summary>
        public static bool NothingCreated(Exception ex)
        {
            var ge = ex as GraphException;
            // a 2xx answer that could not be read means the item exists: never "nothing created"
            if (ge == null) return !(ex is OperationCanceledException) && !(ex is System.Text.Json.JsonException);
            return (ge.StatusCode >= 400 && ge.StatusCode < 500) || ge.StatusCode == 503 || ge.ErrorCode == "UploadInterrupted" || ge.ErrorCode == "NotSent";
        }
    }

    /// <summary>
    /// Looks for the item of an Unknown import row in the listing of its target folder. What the listing of the items returns is
    /// id, type, changeKey, size, createdDateTime and lastModifiedDateTime: no subject, no InternetMessageId, no search key. A candidate is an item
    ///   - of exactly the size of the source item (the announced size of the export listing is the size the target listing gives back);
    ///   - that no import row of the store is Done with (an item the store already accounts for is never the lost one);
    ///   - that was created or modified after the request began. On the lab tenant an imported item carries the time of the import
    ///     (created 23:18:18Z, modified 23:18:22Z for an import started at 23:18:15Z), not the dates of the source: an item that was already in the
    ///     folder before the attempt (an earlier import, a copy of the same mail) is therefore not mistaken for the one the attempt created.
    /// The time bound is the start of the attempt minus ClockSkew when this run knows it, else the end of the request minus LookBack (the longest
    /// attempt: request timeout 15 minutes and the in-place retries). The caller still has to make sure that two rows do not claim the same item.
    /// </summary>
    public static class ImportReconciler
    {
        /// <summary>Difference tolerated between the clock of this computer and the one of the service.</summary>
        public static readonly TimeSpan ClockSkew = TimeSpan.FromMinutes(5);
        /// <summary>How long before the end of its request (the time recorded in the row) an attempt may have begun.</summary>
        public static readonly TimeSpan LookBack = TimeSpan.FromMinutes(30);

        /// <summary>The oldest creation / modification time an item created by the attempt can have.</summary>
        public static DateTime NotBefore(DateTime? attemptStartUtc, DateTime? requestEndedUtc)
        {
            if (attemptStartUtc.HasValue) return attemptStartUtc.Value - ClockSkew;
            return (requestEndedUtc ?? DateTime.UtcNow) - LookBack;
        }

        public static List<string> Candidates(long size, IEnumerable<ItemInfo> listing, ISet<string> known, DateTime notBeforeUtc)
        {
            var found = new List<string>();
            foreach (ItemInfo i in listing)
            {
                if (i.Size != size || (known != null && known.Contains(i.Id))) continue;
                DateTimeOffset? latest = i.Modified.HasValue && (!i.Created.HasValue || i.Modified.Value > i.Created.Value) ? i.Modified : i.Created;
                if (latest.HasValue && latest.Value.UtcDateTime < notBeforeUtc) continue;   // already there before the attempt
                found.Add(i.Id);
            }
            return found;
        }
    }

    public static class ImportPlanner
    {
        static string Norm(string p) { return (p ?? "").Trim('/'); }

        /// <summary>Target path of a source path: explicit mapping (longest prefix), then TargetRoot in front.</summary>
        public static string MapPath(ImportOptions o, string sourcePath)
        {
            string src = Norm(sourcePath);
            string mapped = src;
            if (o.Mapping != null)
            {
                string best = null, bestValue = null;
                foreach (var kv in o.Mapping)
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
            string root = Norm(o.TargetRoot);
            if (root.Length == 0) return mapped;
            return mapped.Length == 0 ? root : root + "/" + mapped;
        }

        static IEnumerable<StoreFolderNode> Flatten(IEnumerable<StoreFolderNode> roots)
        {
            var stack = new Stack<StoreFolderNode>(roots);
            while (stack.Count > 0)
            {
                var n = stack.Pop();
                yield return n;
                foreach (var c in n.Children) stack.Push(c);
            }
        }

        static bool InFilter(string filter, string path)
        {
            return filter == null || string.Equals(path, filter, StringComparison.OrdinalIgnoreCase) || path.StartsWith(filter + "/", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>What the import would do, from the store and the tree of the target (read only, nothing is created).</summary>
        public static ImportPlan Build(ExportStore store, ImportOptions o, IEnumerable<FolderInfo> targetTree)
        {
            var existing = new Dictionary<string, FolderInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (FolderInfo f in targetTree) if (!existing.ContainsKey(f.Path)) existing[f.Path] = f;

            var plan = new ImportPlan();
            var needed = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string filter = string.IsNullOrEmpty(o.SourceFolderPath) ? null : Norm(o.SourceFolderPath);
            foreach (StoreFolderNode n in Flatten(store.GetFolderTree(o.SourceMailboxId)).OrderBy(x => x.Path, StringComparer.OrdinalIgnoreCase))
            {
                if (!InFilter(filter, n.Path)) continue;
                string target = MapPath(o, n.Path);
                FolderInfo found;
                var fp = new ImportFolderPlan
                {
                    SourceFolderKey = n.FolderKey, SourcePath = n.Path, SourceType = n.FolderType, TargetPath = target,
                    Items = n.Own.Exported, Bytes = n.Own.StoredBytes, NotExported = n.Own.Pending + n.Own.Failed
                };
                if (existing.TryGetValue(target, out found)) { fp.Exists = true; fp.TargetFolderId = found.Id; plan.FoldersExisting++; }
                plan.Folders.Add(fp);
                plan.Items += fp.Items; plan.Bytes += fp.Bytes; plan.NotExported += fp.NotExported;
                // the folder and its ancestors, parents first
                var chain = new List<string>();
                for (string p = target; p.Length > 0; p = p.Contains('/') ? p.Substring(0, p.LastIndexOf('/')) : "") chain.Add(p);
                chain.Reverse();
                foreach (string p in chain)
                    if (!existing.ContainsKey(p) && seen.Add(p)) needed.Add(p);
            }
            plan.PathsToCreate = needed.OrderBy(p => p.Count(c => c == '/')).ThenBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
            return plan;
        }
    }

    /// <summary>
    /// Creates the missing folders of a list of target paths (parents first), create only: a folder is found by its name under its parent
    /// and never renamed, moved or deleted. Used by the import and by the replication of the lab.
    /// </summary>
    public sealed class FolderEnsurer
    {
        readonly MailboxApi _api;
        readonly string _mailbox;
        readonly Dictionary<string, FolderInfo> _folders = new Dictionary<string, FolderInfo>(StringComparer.OrdinalIgnoreCase);
        readonly HashSet<string> _failed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public int Created { get; private set; }
        public int Failed { get; private set; }
        public List<string> Problems { get; } = new List<string>();

        public FolderEnsurer(MailboxApi api, string mailboxId, IEnumerable<FolderInfo> existingTree)
        {
            _api = api; _mailbox = mailboxId;
            foreach (FolderInfo f in existingTree) if (!_folders.ContainsKey(f.Path)) _folders[f.Path] = f;
        }

        public FolderInfo Find(string path) { FolderInfo f; return _folders.TryGetValue(path, out f) ? f : null; }

        /// <param name="limiter">Optional: every creation takes a slot (the requests in flight stay under the limit).</param>
        public void EnsureAll(IEnumerable<string> pathsParentsFirst, Func<string, string> typeOf, Action<string> log, CancellationToken ct, AdaptiveLimiter limiter = null)
        {
            foreach (string path in pathsParentsFirst)
            {
                ct.ThrowIfCancellationRequested();
                if (_folders.ContainsKey(path)) continue;
                int slash = path.LastIndexOf('/');
                string parentPath = slash < 0 ? null : path.Substring(0, slash);
                if (parentPath != null && _failed.Contains(parentPath)) { _failed.Add(path); Failed++; continue; }
                FolderInfo parent = null;
                if (parentPath != null) _folders.TryGetValue(parentPath, out parent);
                string type = typeOf == null ? null : typeOf(path);
                try
                {
                    using (limiter == null ? null : limiter.Enter(ct)) _folders[path] = Create(path, parent, string.IsNullOrEmpty(type) ? null : type, log, ct);
                    Created++;
                    if (log != null) log("folder created: " + path);
                }
                catch (GraphException ex)
                {
                    if (ex.StatusCode == 401 || ex.StatusCode == 403) throw;
                    _failed.Add(path);
                    Failed++;
                    string why = "folder '" + path + "' not created: " + ex.Message;
                    Problems.Add(why);
                    if (log != null) log(why);
                }
            }
        }

        FolderInfo Create(string path, FolderInfo parent, string type, Action<string> log, CancellationToken ct)
        {
            string name = path.Substring(path.LastIndexOf('/') + 1);
            try
            {
                return _api.CreateFolder(_mailbox, name, parent == null ? null : parent.Id, type ?? "IPF.Note", parent == null ? null : parent.Path, ct);
            }
            catch (GraphException ex) when (ex.StatusCode == 400 && type != null && type != "IPF.Note")
            {
                if (log != null) log("folder '" + path + "': the service refused the type " + type + ", creating it as IPF.Note");
                return _api.CreateFolder(_mailbox, name, parent == null ? null : parent.Id, "IPF.Note", parent == null ? null : parent.Path, ct);
            }
        }
    }

    public sealed class ImportJob : TransferJob
    {
        readonly MailboxApi _api;
        readonly ExportStore _store;
        readonly ImportOptions _o;
        readonly ImportRunResult _result = new ImportRunResult();
        readonly ImportSessionPool _sessions;
        readonly ImportRouter _router;
        /// <summary>Unknown rows that this run produced and that are still Unknown (a row settled by the reconciliation leaves the count).</summary>
        long _unknownThisRun;
        /// <summary>Rows refused with nothing created (this run).</summary>
        long _failedThisRun;
        long _reconciledCreated, _reimported, _ambiguous, _exhausted;
        readonly ConcurrentDictionary<long, DateTime> _toReconcile = new ConcurrentDictionary<long, DateTime>();   // import key -> end of the request that left it Unknown (UTC)
        readonly ConcurrentDictionary<long, DateTime> _attemptStart = new ConcurrentDictionary<long, DateTime>();
        readonly ConcurrentDictionary<long, byte> _producedThisRun = new ConcurrentDictionary<long, byte>();
        readonly ConcurrentDictionary<long, int> _rounds = new ConcurrentDictionary<long, int>();
        readonly ImportGate _gate = new ImportGate();
        TimeSpan ReconcileDelay { get { return TimeSpan.FromSeconds(Math.Max(0, _o.ReconcileDelaySeconds)); } }

        /// <summary>The workers take it shared around an import; a reconciliation takes it alone, so that no item is being created while the folder is listed.</summary>
        sealed class ImportGate
        {
            readonly object _l = new object();
            int _active;
            bool _exclusive;
            public void Enter(CancellationToken ct)
            {
                lock (_l)
                {
                    while (_exclusive) { ct.ThrowIfCancellationRequested(); Monitor.Wait(_l, 250); }
                    _active++;
                }
            }
            public void Exit() { lock (_l) { _active--; Monitor.PulseAll(_l); } }
            public void BeginExclusive(CancellationToken ct)
            {
                lock (_l)
                {
                    while (_exclusive) { ct.ThrowIfCancellationRequested(); Monitor.Wait(_l, 250); }
                    _exclusive = true;
                    while (_active > 0)
                    {
                        if (ct.IsCancellationRequested) { _exclusive = false; Monitor.PulseAll(_l); ct.ThrowIfCancellationRequested(); }
                        Monitor.Wait(_l, 250);
                    }
                }
            }
            public void EndExclusive() { lock (_l) { _exclusive = false; Monitor.PulseAll(_l); } }
        }
        public ImportRunResult Result { get { return _result; } }
        public ImportSessionPool Sessions { get { return _sessions; } }

        ImportJob(MailboxApi api, ExportStore store, ImportOptions options)
        {
            _api = api; _store = store; _o = options;
            Graph = api.Graph;
            _sessions = new ImportSessionPool(api);
            _router = new ImportRouter(api, store);
            _router.Learner = new AuxLocationLearner(store, null, new[] { options.TargetMailboxId, options.SourceMailboxId }, Log);
            Graph.MailboxRedirected = _router.Learner.OnRedirect;   // a 308 answered to a folder request names an auxiliary archive too
            InitLimiter(options.Concurrency);
        }

        public static ImportJob Start(MailboxApi api, ExportStore store, ImportOptions options)
        {
            if (string.IsNullOrEmpty(options.SourceMailboxId)) throw new ArgumentException("SourceMailboxId is required.");
            if (string.IsNullOrEmpty(options.TargetMailboxId)) throw new ArgumentException("TargetMailboxId is required.");
            TransferLimits.CheckConcurrency(options.Concurrency);
            var job = new ImportJob(api, store, options);
            job.Begin(job.Run);
            return job;
        }

        void Problem(string text)
        {
            lock (_result.Problems) _result.Problems.Add(text);
            Log(text);
        }

        void Run()
        {
            try { RunCore(); }
            finally { FillJobTotals(); }
        }

        /// <summary>The state of the whole job in the store (all runs together), whatever way this run ended.</summary>
        void FillJobTotals()
        {
            try
            {
                if (_result.JobId == 0) return;
                long done = 0, failed = 0, unknown = 0, pending = 0;
                foreach (StoreStateStat st in _store.GetImportStats(_result.JobId))
                    switch (st.State)
                    {
                        case "Done": done = st.Count; break;
                        case "Failed": failed = st.Count; break;
                        case "Unknown": unknown = st.Count; break;
                        default: pending += st.Count; break;   // Pending and InFlight
                    }
                _result.JobDone = done; _result.JobFailed = failed; _result.JobUnknown = unknown; _result.JobPending = pending;
                _result.JobTotal = done + failed + unknown + pending;
                _result.JobUnknownEarlier = Math.Max(0, unknown - Interlocked.Read(ref _unknownThisRun));
                _result.JobKnown = true;
            }
            catch (Exception ex) { Log("the state of the job could not be read: " + ex.Message); }
        }

        void RunCore()
        {
            CancellationToken ct = Cts.Token;
            string target = _o.TargetMailboxId;

            Metrics.Phase = "folders";
            List<FolderInfo> tree = _api.GetFolderTree(target, ct);
            ImportPlan plan = ImportPlanner.Build(_store, _o, tree);
            var types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (ImportFolderPlan fp in plan.Folders) types[fp.TargetPath] = fp.SourceType;
            var folders = new FolderEnsurer(_api, target, tree);
            folders.EnsureAll(plan.PathsToCreate, p => { string t; return types.TryGetValue(p, out t) ? t : null; }, Log, ct, Limiter);
            _result.FoldersCreated = folders.Created;
            _result.FoldersFailed = folders.Failed;
            lock (_result.Problems) _result.Problems.AddRange(folders.Problems);

            _result.JobName = _o.JobName;
            StoreImportJob job = _store.FindImportJob(_o.JobName);
            bool resumed = job != null;
            if (job == null) job = _store.GetImportJob(_store.CreateImportJob(_o.JobName, target));
            else if (!string.Equals(job.TargetMailbox, target, StringComparison.Ordinal))
                throw new InvalidOperationException("The import job '" + _o.JobName + "' of the store targets another mailbox (" + job.TargetMailbox + "). Give another -JobName.");
            _result.JobId = job.JobId;
            _router.Load(job.JobId);
            foreach (ImportFolderPlan fp in plan.Folders)
            {
                FolderInfo known = folders.Find(fp.TargetPath);
                if (known != null) _router.SetPath(target, known.Id, fp.TargetPath);
            }
            if (_router.Count > 0) Log("routes: " + _router.Count + " folder(s) of this job are served by an auxiliary archive (learned by an earlier run)");
            int recovered;
            if (resumed && _store.LastRecovery != null && _store.LastRecovery.ImportInFlightToUnknownByJob.TryGetValue(job.JobId, out recovered) && recovered > 0)
            {
                _result.RecoveredToUnknown = recovered;
                Log("recovery: " + recovered + " rows of the job '" + _o.JobName + "' were InFlight when an earlier run stopped and are now Unknown (the items may exist)");
            }

            Metrics.Phase = "plan";
            long planned = 0;
            _result.NotQueued = plan.NotExported;
            if (plan.NotExported > 0) Problem(plan.NotExported + " item(s) of the store are not exported yet (Pending or Failed in the store): they are not part of this import. Finish the export, then run the import again.");
            foreach (ImportFolderPlan fp in plan.Folders)
            {
                if (fp.Items == 0) continue;
                FolderInfo tf = folders.Find(fp.TargetPath);
                if (tf == null) { _result.NotQueued += fp.Items; Problem("the " + fp.Items + " items of '" + fp.SourcePath + "' are not imported: the folder '" + fp.TargetPath + "' does not exist in the target"); continue; }
                const int pageSize = 1000;
                for (int page = 0; ; page++)
                {
                    StoreItemPage p = _store.GetItems(fp.SourceFolderKey, page, pageSize, StoreItemSort.ReceivedDate, false, StoreItemState.Exported);
                    planned += _store.AddImportItems(job.JobId, p.Items.Select(i => new StoreImportPlan { ItemKey = i.Key, TargetMailbox = target, TargetFolder = tf.Id }).ToList());
                    if (page + 1 >= p.PageCount) break;
                }
            }
            int requeued = _store.RequeueImportFailed(job.JobId, _o.MaxAttempts);
            StoreStateStat pending = _store.GetImportStats(job.JobId).First(s => s.State == "Pending");
            Log("plan: " + planned + " items added to the job '" + _o.JobName + "' (" + requeued + " failed ones tried again), " + pending.Count + " to import");
            if (resumed)
            {
                var before = _store.GetImportStats(job.JobId).ToDictionary(x => x.State, x => x.Count);
                Log("job '" + _o.JobName + "' resumed: " + before["Done"] + " done, " + before["Failed"] + " failed, " + before["Unknown"] + " unknown (" + (_o.AutoReconcile ? "looked up in the target folder by this run" : "never replayed by itself") + "), " + before["Pending"] + " pending");
            }

            _result.AutoReconcile = _o.AutoReconcile;
            if (_o.AutoReconcile)
            {
                List<StoreImportItem> left = _store.GetImportItems(job.JobId, StoreImportState.Unknown, 10000000);
                foreach (StoreImportItem u in left) _toReconcile[u.ImportKey] = u.UpdatedUtc ?? DateTime.UtcNow;
                if (left.Count > 0)
                    Log("reconcile: " + left.Count + " Unknown rows left by an earlier run will be looked up in their target folder " +
                        (ReconcileDelay.TotalSeconds > 0 ? "once " + Seconds(ReconcileDelay) + " have passed since their request ended" : "now") + " (-NoAutoReconcile leaves them Unknown)");
            }

            Metrics.Phase = "import";
            Metrics.BeginTransfer(pending.Count, pending.StoredBytes);
            _labels = new StoreFolderLabels(_store);
            foreach (StoreFolderCount c in _store.GetImportPendingByFolder(job.JobId)) Metrics.Folders.Expect(_labels.Label(c.FolderKey), c.Count);
            RunWorkers(_o.Concurrency, n => ImportLoop(job.JobId, ct));
            if (_o.AutoReconcile) ReconcileTail(job.JobId, ct);

            _result.Done = Metrics.Items;
            _result.Failed = Interlocked.Read(ref _failedThisRun);
            _result.Unknown = Interlocked.Read(ref _unknownThisRun);
            _result.Bytes = Metrics.Bytes;
            _result.ReconciledCreated = Interlocked.Read(ref _reconciledCreated);
            _result.ReconciledReimported = Interlocked.Read(ref _reimported);
            _result.ReconcileAmbiguous = Interlocked.Read(ref _ambiguous);
            _result.ReconcileExhausted = Interlocked.Read(ref _exhausted);
            _result.SessionsCreated = _sessions.Created;
            _result.FoldersInAux = _router.Count;
            _result.AuxArchives = _router.Mailboxes;
            foreach (TimeSpan t in _sessions.Lifetimes) _result.SessionLifetimeSeconds.Add(t.TotalSeconds);
            _result.Completed = !ct.IsCancellationRequested;
            Metrics.Phase = "done";
            _result.Final = Snapshot();
            string summary = _result.ReconcileSummary();
            if (summary != null) Log(summary);
            if (_unknownThisRun > 0)
                Log(_unknownThisRun + " items are Unknown (the service may have created them): " + (_o.AutoReconcile
                    ? "the lookup in the target folder could not settle them (several items could be them, or the run was stopped); list them with Get-MbtImportUnknown -Check"
                    : "list them with Get-MbtImportUnknown"));
        }

        static string Seconds(TimeSpan t) { return Math.Round(t.TotalSeconds, 1).ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + " s"; }

        /// <summary>One worker: takes the Pending rows one by one; when an Unknown row is due for its lookup, settles the Unknown rows first.</summary>
        void ImportLoop(long jobId, CancellationToken ct)
        {
            while (!ct.IsCancellationRequested)
            {
                if (ReconcileDue()) ReconcilePass(jobId, ct);
                _gate.Enter(ct);
                try
                {
                    using (Limiter.Enter(ct))
                    {
                        List<StoreImportItem> next = _store.NextImportItems(jobId, 1);
                        if (next.Count == 0) return;
                        ImportOne(next[0], ct);
                    }
                }
                finally { _gate.Exit(); }
            }
        }

        bool ReconcileDue()
        {
            if (_toReconcile.IsEmpty) return false;
            DateTime now = DateTime.UtcNow;
            TimeSpan delay = ReconcileDelay;
            foreach (KeyValuePair<long, DateTime> kv in _toReconcile) if (now - kv.Value >= delay) return true;
            return false;
        }

        /// <summary>The queue is empty: waits for the Unknown rows to settle, looks them up, imports again what was not created, until nothing is left to look up.</summary>
        void ReconcileTail(long jobId, CancellationToken ct)
        {
            bool announced = false;
            while (!ct.IsCancellationRequested && !_toReconcile.IsEmpty)
            {
                TimeSpan wait = _toReconcile.Values.Min() + ReconcileDelay - DateTime.UtcNow;
                if (wait > TimeSpan.Zero)
                {
                    Metrics.Phase = "reconcile";
                    if (!announced) { Log("reconcile: " + _toReconcile.Count + " Unknown row(s) to look up in the target after the settle delay, in " + Seconds(wait)); announced = true; }
                    ct.WaitHandle.WaitOne(wait < TimeSpan.FromMilliseconds(500) ? wait : TimeSpan.FromMilliseconds(500));
                    continue;
                }
                announced = false;
                try { ReconcilePass(jobId, ct); }
                catch (OperationCanceledException) { break; }   // stopped: the rows stay Unknown, the next run looks them up
                if (ct.IsCancellationRequested) break;
                if (_store.GetImportStats(jobId).First(x => x.State == "Pending").Count > 0)
                {
                    Metrics.Phase = "import";
                    RunWorkers(_o.Concurrency, n => ImportLoop(jobId, ct));
                }
            }
        }

        /// <summary>
        /// Looks up the Unknown rows whose delay is over, with nothing being imported (the gate): the folder is listed once, then each row is
        /// settled by what the listing shows (see the header). An item that could be the one of two rows is ambiguous for both.
        /// </summary>
        void ReconcilePass(long jobId, CancellationToken ct)
        {
            _gate.BeginExclusive(ct);
            try
            {
                DateTime now = DateTime.UtcNow;
                TimeSpan delay = ReconcileDelay;
                var due = new HashSet<long>(_toReconcile.Where(kv => now - kv.Value >= delay).Select(kv => kv.Key));
                if (due.Count == 0) return;
                List<StoreImportItem> unknown = _store.GetImportItems(jobId, StoreImportState.Unknown, 10000000);
                var present = new HashSet<long>(unknown.Select(u => u.ImportKey));
                foreach (long k in due.ToList())
                    if (!present.Contains(k)) { due.Remove(k); DateTime gone; _toReconcile.TryRemove(k, out gone); }   // settled meanwhile (by a person)
                if (due.Count == 0) return;
                HashSet<string> known = _store.GetDoneTargetItemIds(null);
                int created = 0, absent = 0, ambiguous = 0;
                int rounds = Math.Max(1, _o.ReconcileRounds);
                foreach (var group in unknown.GroupBy(u => u.TargetMailbox + "\n" + u.TargetFolder))
                {
                    List<StoreImportItem> rows = group.ToList();
                    var dueRows = new List<StoreImportItem>();
                    foreach (StoreImportItem u in rows.Where(r => due.Contains(r.ImportKey)))
                    {
                        int used;
                        _rounds.TryGetValue(u.ImportKey, out used);
                        if (used >= rounds)
                        {
                            DateTime gone; _toReconcile.TryRemove(u.ImportKey, out gone);
                            Interlocked.Increment(ref _exhausted);
                            Problem("import " + u.ImportKey + " left Unknown: it was looked up " + used + " time(s) (-ReconcileRounds " + rounds + ") without being settled and still ends Unknown");
                        }
                        else dueRows.Add(u);
                    }
                    if (dueRows.Count == 0) continue;
                    List<ItemInfo> listing;
                    try
                    {
                        using (Limiter.Enter(ct)) listing = _api.EnumerateItems(rows[0].TargetMailbox, rows[0].TargetFolder, 0, ct).ToList();
                    }
                    catch (GraphException ex)
                    {
                        if (ex.StatusCode == 401 || ex.StatusCode == 403) { Fatal(ex); return; }
                        foreach (StoreImportItem u in dueRows)
                        {
                            _rounds.AddOrUpdate(u.ImportKey, 1, (k, v) => v + 1);
                            _toReconcile[u.ImportKey] = DateTime.UtcNow;   // looked up again after another delay
                        }
                        Problem("reconcile: the target folder " + rows[0].TargetFolder + " could not be listed (" + Short(ex.Message) + "), " + dueRows.Count + " Unknown row(s) will be looked up again after the delay");
                        continue;
                    }
                    // the candidates of every Unknown row of the folder, due or not: an item wanted by two rows cannot be given to either
                    var candidates = new Dictionary<long, List<string>>();
                    var claims = new Dictionary<string, int>(StringComparer.Ordinal);
                    foreach (StoreImportItem u in rows)
                    {
                        DateTime start;
                        DateTime notBefore = ImportReconciler.NotBefore(_attemptStart.TryGetValue(u.ImportKey, out start) ? start : (DateTime?)null, u.UpdatedUtc);
                        List<string> c = ImportReconciler.Candidates(_store.GetItem(u.ItemKey).Size, listing, known, notBefore);
                        candidates[u.ImportKey] = c;
                        foreach (string id in c) { int n; claims.TryGetValue(id, out n); claims[id] = n + 1; }
                    }
                    foreach (StoreImportItem u in dueRows)
                    {
                        DateTime gone;
                        List<string> c = candidates[u.ImportKey];
                        _toReconcile.TryRemove(u.ImportKey, out gone);
                        if (c.Count == 1 && claims[c[0]] == 1)
                        {
                            _store.ResolveImportUnknown(u.ImportKey, true, c[0]);
                            Metrics.Folders.Settled(FolderLabel(u.ItemKey), true, u.Length ?? 0);
                            Metrics.ItemDone(u.Length ?? 0);
                            Interlocked.Increment(ref Metrics.TotalItems);
                            Interlocked.Increment(ref _reconciledCreated);
                            LeftUnknown(u.ImportKey);
                            created++;
                            Log("import " + u.ImportKey + " reconciled: the item was created, id " + c[0]);
                        }
                        else if (c.Count == 0)
                        {
                            _rounds.AddOrUpdate(u.ImportKey, 1, (k, v) => v + 1);
                            if (u.Attempts >= _o.MaxAttempts)
                            {
                                _store.MarkImportFailed(u.ImportKey, "no matching item in the target folder after the lookup: nothing was created, and the " + u.Attempts + " attempts allowed are used");
                                Metrics.Folders.SettledFailed(FolderLabel(u.ItemKey));
                                Interlocked.Increment(ref _failedThisRun);
                                Log("import " + u.ImportKey + " reconciled: no item was created, but its " + u.Attempts + " attempts are used (Failed, nothing created)");
                            }
                            else
                            {
                                _store.ResolveImportUnknown(u.ImportKey, false, null);
                                Metrics.Folders.Settled(FolderLabel(u.ItemKey), false);
                                Interlocked.Increment(ref Metrics.TotalItems);
                                Interlocked.Increment(ref _reimported);
                                Log("import " + u.ImportKey + " reconciled: no item was created in the target folder, it is imported again (attempt " + (u.Attempts + 1) + " of " + _o.MaxAttempts + ")");
                            }
                            LeftUnknown(u.ImportKey);
                            absent++;
                        }
                        else
                        {
                            Interlocked.Increment(ref _ambiguous);
                            ambiguous++;
                            Problem("import " + u.ImportKey + " left Unknown: " +
                                (c.Count == 1 ? "the one item that could be it (" + c[0] + ") could also be the item of another Unknown row" : c.Count + " items in the target folder could be it (" + string.Join(", ", c) + ")") +
                                "; settle it with Resolve-MbtImportUnknown");
                        }
                    }
                }
                Log("reconcile: " + (created + absent + ambiguous) + " row(s) looked up: " + created + " created, " + absent + " not created, " + ambiguous + " ambiguous");
            }
            finally { _gate.EndExclusive(); }
        }

        /// <summary>A row of this run is no longer Unknown (settled by the lookup).</summary>
        void LeftUnknown(long importKey)
        {
            byte b;
            if (_producedThisRun.TryRemove(importKey, out b)) Interlocked.Decrement(ref _unknownThisRun);
        }

        /// <summary>The request ended Unknown: the row is looked up after the settle delay (unless the run is being stopped: the next run does it).</summary>
        void UnknownOnce(long importKey, CancellationToken ct)
        {
            Interlocked.Increment(ref _unknownThisRun);
            _producedThisRun[importKey] = 1;
            if (_o.AutoReconcile && !ct.IsCancellationRequested) _toReconcile[importKey] = DateTime.UtcNow;
        }

        void ImportOne(StoreImportItem it, CancellationToken ct)
        {
            if (ct.IsCancellationRequested) { _store.MarkImportPending(it.ImportKey); return; }
            _attemptStart[it.ImportKey] = DateTime.UtcNow;
            try
            {
                ImportResult r = ImportHelper.ImportWithSession(_api, _sessions, _store, it.TargetMailbox, it.TargetFolder, it.ItemKey, it.Length ?? 0, ct, Log,
                    _router, (m, f) => _store.SetImportTarget(it.ImportKey, m, f));
                _store.MarkImportDone(it.ImportKey, r.ItemId);
                Metrics.ItemDone(r.Bytes);
                Metrics.Folders.Done(FolderLabel(it.ItemKey), r.Bytes);
                Succeeded();
            }
            catch (OperationCanceledException)
            {
                _store.MarkImportUnknown(it.ImportKey, "interrupted while the request was in flight: the item may exist");
                UnknownOnce(it.ImportKey, ct);
                Metrics.ItemFailed();
                Metrics.Folders.Unknown(FolderLabel(it.ItemKey));
            }
            catch (Exception ex)
            {
                Metrics.ItemFailed();
                if (ImportHelper.NothingCreated(ex))
                {
                    _store.MarkImportFailed(it.ImportKey, ex.Message);
                    Interlocked.Increment(ref _failedThisRun);
                    Metrics.Folders.Fail(FolderLabel(it.ItemKey));
                    Log("import " + it.ImportKey + " failed (nothing created): " + Short(ex.Message));
                    var ge = ex as GraphException;
                    if (ge != null && (ge.StatusCode == 401 || ge.StatusCode == 403)) Fatal(ex);
                }
                else
                {
                    _store.MarkImportUnknown(it.ImportKey, ex.Message);
                    UnknownOnce(it.ImportKey, ct);
                    Metrics.Folders.Unknown(FolderLabel(it.ItemKey));
                    Log("import " + it.ImportKey + " UNKNOWN (the item may have been created): " + Short(ex.Message));
                }
                FailedOnce(ex.Message);
            }
        }

        StoreFolderLabels _labels;
        string FolderLabel(long itemKey)
        {
            if (_labels == null) _labels = new StoreFolderLabels(_store);
            StoreItem si = _store.GetItem(itemKey);
            return _labels.Label(si == null ? null : si.FolderKey);
        }

        static string Short(string s) { return s != null && s.Length > 300 ? s.Substring(0, 300) + "..." : s; }
    }
}
