// =============================================================================
//  Mailbox Transfer - engine, part 7: export of a whole mailbox into the store
// -----------------------------------------------------------------------------
//  Phases of an ExportJob (all in the engine, on its own threads):
//    tree       folder tree of the mailbox -> store (mailbox, folders)
//    inventory  items of every folder, page after page -> store (Pending; a changed changeKey puts an Exported item back to
//               Pending: the store does it in UpsertItems). Skipped with SkipInventory (resume without listing again).
//    export     Pending items in batches of at most BatchItems items and BatchBytes (announced sizes) per exportItems
//               request, taken by Concurrency workers (4 at most: the service allows 4 concurrent requests per mailbox).
//               Each answer is streamed to a temporary file, copied into a pack of the store (SHA-256 compared), decoded for its
//               preview, then the temporary file is deleted. A request that fails is asked again item by item.
//  A new run does only what is left: Pending items, and Failed items that were tried fewer than MaxAttempts times.
//
//  Auto-expanding archive (Kind Archive): a folder of the main archive can live in an auxiliary archive. The folder is recorded ONCE, in the tree
//  of the main archive (it is the placeholder the main archive lists); its items are recorded with the mailbox that really holds them:
//    - the listing of the items of such a folder answers 308 to the auxiliary archive: the pages come from there, the mailbox id of the URL that
//      answered is the mailbox of the items (ItemInfo.MailboxId);
//    - an item asked from the main archive can also answer ErrorArchiveFolderMovedPermanently with the URL of the auxiliary archive: the item
//      is exported from there and its record moves to that mailbox.
//  Every auxiliary archive met is recorded as a location (kind AuxArchive, how and when) and as a mailbox of the store. In the main archive every
//  folder is listed, whatever its announced item count (an expanded folder may announce none). Kind AuxArchive exports one auxiliary archive by
//  itself (its own tree); a folder it holds that the main archive tree already records is not recorded a second time.
//
//  A folder being moved to an auxiliary archive (measured live 2026-10-31): the auxiliary archive lists the folder skeleton first and the items later; the
//  items of such a folder answer 500 "Cannot query rows in a table" there while the main archive still holds them. The listing is asked again a bounded number
//  of times (MailboxApi.ContentAttempts), then the folder is DEFERRED (ExportResult.FoldersDeferred: path, mailbox, the message of the service), never counted as
//  empty; the other folders go on. The same folder in the main archive and in an auxiliary archive is recognised by its folder entry key (the folder part of the
//  Exchange id, see ExchangeEntryId), not by the id string (the mailbox part differs): the auxiliary archive's items are recorded under the main archive's folder
//  row, and an item listed on both sides under two ids with the same message entry key is recorded once (ExportResult.ItemsDeduplicated). An export that
//  deferred a folder, failed an item or did not finish is "resumable": run the same command again (ExportResult.NeedsAnotherRun).
// =============================================================================
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer
{
    public sealed class ExportOptions
    {
        public string MailboxId;
        public string Kind = "Primary";
        public string Upn;
        public int Concurrency = TransferLimits.MaxConcurrency;
        public int BatchItems = MailboxApi.MaxExportBatch;
        /// <summary>Announced size of the items of one request: bounds the size of the JSON answer (a third larger than the items).</summary>
        public long BatchBytes = 200L * 1024 * 1024;
        /// <summary>A Failed item goes back to Pending at the start of a run while it was tried fewer times than this.</summary>
        public int MaxAttempts = 3;
        public bool SkipInventory;
        public int ItemPageSize;
        /// <summary>Only this folder and the folders below it (path as in the store), null for the whole mailbox.</summary>
        public string FolderPath;
        /// <summary>
        /// What a browse selection asked for, in the ids of the import-export API (MailboxBrowser.ResolvePlan): these folders (and the ones below them
        /// when the entry says so), or only these items of a folder. Exclusive with FolderPath. The tree of the whole mailbox is still recorded.
        /// </summary>
        public ExportPlan Plan;
        /// <summary>
        /// After this many folders in a row whose content is not available yet (a location being filled: every folder refuses), the next folders of the run are asked ONCE
        /// instead of ContentAttempts times, so that hundreds of folders are not 14 s each; the folders that refuse are deferred all the same. A folder that lists again
        /// puts the full asks back. 0: always the full asks.
        /// </summary>
        public int NotReadyTrip = 3;
        public string TempPath;
        /// <summary>
        /// Also export Recoverable Items (what was deleted from Deleted Items, purged, held, versions), read through the import-export API: they are recorded in the store under one
        /// top folder, "Recoverable Items/Deletions"... The folders are RecoverableItemsFolders, else Deletions, Purges, Versions and DiscoveryHolds (Calendar Logging and
        /// SubstrateHolds are Microsoft's own copies: only when named). Exclusive with Plan (the plan names its folders: select them in the browse view) and with FolderPath, except a
        /// FolderPath below "Recoverable Items", which asks for them by itself. Primary, Archive and AuxArchive.
        /// </summary>
        public bool IncludeRecoverableItems;
        /// <summary>Names of the folders of Recoverable Items to export (see RecoverableItems.KnownFolders); null or empty: the default set.</summary>
        public string[] RecoverableItemsFolders;
    }

    public sealed class ExportResult
    {
        public long Exported, Failed, Bytes, Inventoried, PendingAtStart, RequeuedFailed;
        public int Folders, FoldersNotListed, PreviewWarnings;
        public List<string> Problems = new List<string>();
        public TransferSnapshot Final;
        public bool Completed;
        /// <summary>Recoverable Items: the folders recorded (the root included) in this run.</summary>
        public int RecoverableFolders;
        /// <summary>Auto-expanding archive: folders of the main archive whose items live in an auxiliary archive, those mailboxes, and the items recorded in them (this run).</summary>
        public int FoldersMoved;
        public List<string> AuxMailboxes = new List<string>();
        public long ItemsMoved;
        /// <summary>Folders whose items could not be listed because the service cannot read their content yet (being moved to an auxiliary archive): not exported in this run, not empty.</summary>
        public List<DeferredFolder> FoldersDeferred = new List<DeferredFolder>();
        /// <summary>Auxiliary archive export: folders that are in the main archive store and in the auxiliary archive with their items on one side only or on both (a move in progress), one sentence each.</summary>
        public List<string> MovesInProgress = new List<string>();
        /// <summary>Items listed on the other side of a move under another id but the same message entry key: recorded once. Records of items not exported yet that followed the item to its new mailbox and id.</summary>
        public long ItemsDeduplicated, ItemsRehomed;
        /// <summary>Deferred folders that were asked once because the location kept refusing (ExportOptions.NotReadyTrip).</summary>
        public int FoldersDeferredFast;
        /// <summary>Something is left for another run of the same command: a deferred folder, a folder not listed, a failed item, or the run did not finish. Entry script exit code 2.</summary>
        public bool NeedsAnotherRun { get { return !Completed || Failed > 0 || FoldersNotListed > 0 || FoldersDeferred.Count > 0; } }
    }

    /// <summary>A folder the export could not list because its content is not available yet.</summary>
    public sealed class DeferredFolder
    {
        public string Path, MailboxId, FolderId, ErrorCode, Message;
        public int StatusCode, Attempts;
        public override string ToString() { return "'" + Path + "' in " + MailboxId + ": content not available yet (HTTP " + StatusCode + " " + ErrorCode + ": " + Message + "), asked " + Attempts + " time(s); not exported, run again later"; }
    }

    public sealed class ExportJob : TransferJob
    {
        readonly MailboxApi _api;
        readonly ExportStore _store;
        readonly ExportOptions _o;
        readonly ExportResult _result = new ExportResult();
        public ExportResult Result { get { return _result; } }

        readonly StoreMailboxKind _kind;
        readonly AuxLocationLearner _learner;
        // AuxArchive kind: folder id in this mailbox -> mailbox and folder id of the record the main archive tree already holds (same folder entry key)
        readonly Dictionary<string, Tuple<string, string>> _folderOwner = new Dictionary<string, Tuple<string, string>>(StringComparer.Ordinal);
        List<FolderInfo> _tree = new List<FolderInfo>();
        readonly HashSet<string> _movedFolders = new HashSet<string>(StringComparer.Ordinal);

        ExportJob(MailboxApi api, ExportStore store, ExportOptions options)
        {
            _api = api; _store = store; _o = options;
            Graph = api.Graph;
            _kind = KindOf(options.Kind);
            _learner = new AuxLocationLearner(store, options.Upn, new[] { options.MailboxId }, Log);
            Graph.MailboxRedirected = _learner.OnRedirect;
            InitLimiter(options.Concurrency);
        }

        /// <summary>Primary (default), Archive / MainArchive, AuxArchive.</summary>
        public static StoreMailboxKind KindOf(string kind)
        {
            if (string.Equals(kind, "Archive", StringComparison.OrdinalIgnoreCase) || string.Equals(kind, "MainArchive", StringComparison.OrdinalIgnoreCase)) return StoreMailboxKind.Archive;
            if (string.Equals(kind, "AuxArchive", StringComparison.OrdinalIgnoreCase) || string.Equals(kind, "Aux", StringComparison.OrdinalIgnoreCase)) return StoreMailboxKind.AuxArchive;
            return StoreMailboxKind.Primary;
        }

        public static ExportJob Start(MailboxApi api, ExportStore store, ExportOptions options)
        {
            if (string.IsNullOrEmpty(options.MailboxId)) throw new ArgumentException("MailboxId is required.");
            if (string.IsNullOrEmpty(options.TempPath)) throw new ArgumentException("TempPath is required.");
            TransferLimits.CheckConcurrency(options.Concurrency);
            if (options.Plan != null && !string.IsNullOrEmpty(options.FolderPath)) throw new ArgumentException("Plan and FolderPath are exclusive.");
            RecoverableItems.CheckNames(options.RecoverableItemsFolders);
            if (options.RecoverableItemsFolders != null && options.RecoverableItemsFolders.Any(x => !string.IsNullOrWhiteSpace(x)) && !options.IncludeRecoverableItems && options.Plan == null && !RecoverableItems.IsPath(options.FolderPath))
                throw new ArgumentException("RecoverableItemsFolders needs IncludeRecoverableItems.");
            if (options.IncludeRecoverableItems && options.Plan != null)
                throw new ArgumentException("IncludeRecoverableItems and Plan are exclusive: the plan names its folders (select the folders of Recoverable Items in the browse view).");
            if (options.IncludeRecoverableItems && !string.IsNullOrEmpty(options.FolderPath) && !RecoverableItems.IsPath(options.FolderPath))
                throw new ArgumentException("IncludeRecoverableItems adds Recoverable Items to the export of the whole mailbox; it cannot be combined with the folder '" + options.FolderPath + "'. Name the folder 'Recoverable Items/Deletions' instead.");
            if (options.Plan != null && !options.Plan.IsComplete)
                throw new ArgumentException("The export plan has entries that cannot be exported: " + string.Join(" ", options.Plan.Problems));
            if (options.Plan != null && options.Plan.Folders.Count == 0) throw new ArgumentException("The export plan is empty.");
            if (options.Plan != null && options.Plan.Folders.Any(f => f.MailboxId != null && !MailboxApi.SameMailbox(f.MailboxId, options.MailboxId)))
                throw new ArgumentException("The export plan has folders of other mailboxes (a folder that moved to an auxiliary archive): export each mailbox with plan.ForMailbox(mailboxId).");
            if (options.BatchItems < 1 || options.BatchItems > MailboxApi.MaxExportBatch)
                throw new ArgumentOutOfRangeException("options", "BatchItems must be between 1 and " + MailboxApi.MaxExportBatch + ".");
            var job = new ExportJob(api, store, options);
            job.Begin(job.Run);
            return job;
        }

        void Run()
        {
            CancellationToken ct = Cts.Token;
            string mailbox = _o.MailboxId;

            Metrics.Phase = "tree";
            List<FolderInfo> tree = _api.GetFolderTree(mailbox, ct);
            tree = AddRecoverableItems(tree, ct);
            _tree = tree;
            _store.UpsertMailbox(new StoreMailbox { Id = mailbox, Kind = _kind, Upn = _o.Upn, DisplayName = _kind + " mailbox" });
            _store.RecordLocation(new StoreMailboxLocation { MailboxId = mailbox, Kind = _kind, Upn = _o.Upn, Source = "ExportTarget", Detail = "mailbox exported", Rank = StoreMailboxLocation.RankGiven });
            if (_kind == StoreMailboxKind.AuxArchive) MapOwnedFolders(tree);
            PrepareScope(tree);
            List<FolderInfo> scope = tree.Where(InScope).ToList();
            // The store records the folders that are exported and the path to them, not the whole folder tree of the mailbox: a PST made from the store
            // (and the browse of the store) then holds the exported folders only, not hundreds of empty ones around them.
            HashSet<string> recorded = FoldersOnPath(tree, scope);
            _store.UpsertFolders(tree.Where(f => !_folderOwner.ContainsKey(f.Id) && recorded.Contains(f.Id)).Select(f => new StoreFolder
            {
                MailboxId = mailbox, FolderId = f.Id, ParentFolderId = f.ParentFolderId, DisplayName = f.DisplayName, FolderType = f.Type,
                SourceItemCount = f.TotalItemCount, SourceSize = f.TotalSize >= 0 ? (long?)f.TotalSize : null
            }).ToList());
            _result.Folders = tree.Count;
            Log("tree: " + tree.Count + " folders" + (_folderOwner.Count > 0 ? " (" + _folderOwner.Count + " of them already recorded in the main archive tree)" : ""));
            foreach (string move in _result.MovesInProgress) Log(move);

            if (!string.IsNullOrEmpty(_o.FolderPath) && scope.Count == 0) throw new InvalidOperationException("Folder '" + _o.FolderPath + "' not found in this mailbox.");
            if (_o.Plan != null && _planFolderIds.Any(id => !tree.Any(f => f.Id == id)))
                throw new InvalidOperationException("Folder(s) of the export plan not found in this mailbox: " + string.Join(", ", _planFolderIds.Where(id => !tree.Any(f => f.Id == id))));
            if (_o.Plan != null) Log("plan: " + _o.Plan.Folders.Count + " entr" + (_o.Plan.Folders.Count == 1 ? "y" : "ies") + ", " + scope.Count + " folder(s) in scope" + (_planItems.Count > 0 ? ", " + _planItems.Values.Sum(s => s.Count) + " item(s) named" : ""));

            if (!_o.SkipInventory)
            {
                Metrics.Phase = "inventory";
                Inventory(scope, ct);
                if (ct.IsCancellationRequested) return;
            }

            List<string> sources = ExportMailboxes();
            foreach (string m in sources) _result.RequeuedFailed += _store.RequeueFailed(m, _o.MaxAttempts);
            long pendingCount = 0, pendingBytes = 0;
            foreach (string m in sources)
            {
                StoreStateStat st = _store.GetItemStats(m).First(s => s.State == "Pending");
                pendingCount += st.Count; pendingBytes += st.AnnouncedBytes;
            }
            StoreStateStat pending = new StoreStateStat { State = "Pending", Count = pendingCount, AnnouncedBytes = pendingBytes };
            _result.PendingAtStart = pending.Count;
            Metrics.Phase = "export";
            Metrics.BeginTransfer(pending.Count, pending.AnnouncedBytes);
            Log("export: " + pending.Count + " items to export" + (_result.RequeuedFailed > 0 ? " (" + _result.RequeuedFailed + " failed ones tried again)" : ""));
            ExportPending(ct);

            _result.Exported = Metrics.Items;
            _result.Failed = Metrics.Failed;
            _result.Bytes = Metrics.Bytes;
            _result.Completed = !ct.IsCancellationRequested;
            Metrics.Phase = "done";
            _result.Final = Snapshot();
        }

        // Recoverable Items below the root that a plan did not ask for by name (they are not in the subtree of "Recoverable Items" unless the plan names them)
        readonly HashSet<string> _recoverableLeftOut = new HashSet<string>(StringComparer.Ordinal);

        /// <summary>
        /// Recoverable Items are not in the top-level listing of the import-export API: when the export asks for them (IncludeRecoverableItems, a FolderPath below
        /// "Recoverable Items", or a plan whose folders are not in the tree) they are read by their well-known name and added to the tree under the top folder
        /// "Recoverable Items". A plan names its folders: every folder of Recoverable Items is read, and the subtree of the root only takes the folders asked for (the default set,
        /// or RecoverableItemsFolders) unless the plan names a folder itself.
        /// </summary>
        List<FolderInfo> AddRecoverableItems(List<FolderInfo> tree, CancellationToken ct)
        {
            bool byPlan = _o.Plan != null && _o.Plan.Folders.Any(f => !string.IsNullOrEmpty(f.ExportFolderId) && (f.MailboxId == null || MailboxApi.SameMailbox(f.MailboxId, _o.MailboxId)) && !tree.Any(t => t.Id == f.ExportFolderId));
            bool byPath = RecoverableItems.IsPath(_o.FolderPath);
            if (!_o.IncludeRecoverableItems && !byPlan && !byPath) return tree;
            bool explicitList = _o.RecoverableItemsFolders != null && _o.RecoverableItemsFolders.Any(x => !string.IsNullOrWhiteSpace(x));
            string[] chosen = RecoverableItems.Chosen(RecoverableItems.Names(_o.RecoverableItemsFolders));
            // FolderPath "Recoverable Items" takes the folders asked for (the default set, or the explicit list); a path below it names its folder itself
            bool wholeRoot = byPath && string.Equals(_o.FolderPath.Trim('/'), RecoverableItems.TopName, StringComparison.OrdinalIgnoreCase);
            RecoverableTree rt;
            try { rt = _api.GetRecoverableItemsTree(_o.MailboxId, _o.IncludeRecoverableItems || wholeRoot ? chosen : null, ct); }
            catch (GraphException ex) when (ex.StatusCode == 404)
            {
                Log("Recoverable Items: this mailbox has none (HTTP 404 " + ex.ErrorCode + ")");
                return tree;
            }
            if (byPlan)
                foreach (FolderInfo f in rt.Folders)
                    if (!RecoverableItems.Wanted(f.Path, chosen)) _recoverableLeftOut.Add(f.Id);
            foreach (string n in rt.Missing)
            {
                if (explicitList || (byPath && !wholeRoot)) Problem("Recoverable Items has no folder '" + n + "' in this mailbox");
                else Log("Recoverable Items: no folder '" + n + "' in this mailbox (nothing to export from it)");
            }
            var all = new List<FolderInfo>(tree);
            foreach (FolderInfo f in rt.Folders) if (!all.Any(x => x.Id == f.Id)) all.Add(f);
            _result.RecoverableFolders = rt.Folders.Count;
            Log("Recoverable Items: " + rt.Folders.Count + " folder(s) read through the import-export API (" + string.Join(", ", rt.Folders.Skip(1).Select(f => f.Path.Substring(RecoverableItems.TopName.Length + 1) + " " + f.TotalItemCount)) + ")"
                + (rt.Skipped.Count > 0 && !byPlan && (!byPath || wholeRoot) ? "; not asked for: " + string.Join(", ", rt.Skipped) : ""));
            return all.OrderBy(f => f.Path, StringComparer.OrdinalIgnoreCase).ToList();
        }

        /// <summary>The mailbox exported and, for the main archive, the auxiliary archives of the store (their items are exported from there).</summary>
        List<string> ExportMailboxes()
        {
            var list = new List<string> { _o.MailboxId };
            if (_kind == StoreMailboxKind.Archive)
                foreach (StoreMailbox m in _store.GetMailboxes())
                    if (m.Kind == StoreMailboxKind.AuxArchive && !MailboxApi.SameMailbox(m.Id, _o.MailboxId)) list.Add(m.Id);
            _result.AuxMailboxes = list.Skip(1).ToList();
            return list;
        }

        /// <summary>
        /// Auxiliary archive: a folder of its tree that the main archive tree of the store already records is the same folder, recognised by the folder entry key of the
        /// id (the id strings differ: the mailbox part is not the same); the items are recorded under the main archive's folder row. Ids that are not in the Exchange
        /// format are compared as they are (the key of such an id is the id).
        /// </summary>
        void MapOwnedFolders(List<FolderInfo> tree)
        {
            var byKey = new Dictionary<string, Tuple<string, string, long?>>(StringComparer.Ordinal);
            foreach (StoreMailbox m in _store.GetMailboxes().Where(x => x.Kind == StoreMailboxKind.Archive))
                foreach (StoreFolder sf in _store.GetFolders(m.Id))
                    byKey[ExchangeEntryId.FolderKeyOf(sf.FolderId)] = Tuple.Create(m.Id, sf.FolderId, sf.SourceItemCount);
            foreach (FolderInfo f in tree)
            {
                Tuple<string, string, long?> owner;
                if (!byKey.TryGetValue(ExchangeEntryId.FolderKeyOf(f.Id), out owner)) continue;
                _folderOwner[f.Id] = Tuple.Create(owner.Item1, owner.Item2);
                long inArchive = owner.Item3 ?? 0;
                if (f.TotalItemCount == 0 && inArchive > 0) _result.MovesInProgress.Add("folder '" + f.Path + "' is being moved to this auxiliary archive: " + inArchive + " item(s) were in the main archive when it was exported, the folder is still a skeleton here");
                else if (f.TotalItemCount > 0 && inArchive > 0) _result.MovesInProgress.Add("folder '" + f.Path + "' has items on both sides of a move: " + inArchive + " in the main archive (when it was exported), " + f.TotalItemCount + " here; an item listed on both sides is recorded once");
            }
        }

        // Selection plan: the folders in scope (by id, with the subtrees asked), and the items named for a folder (id -> item ids).
        readonly HashSet<string> _planScope = new HashSet<string>(StringComparer.Ordinal);
        readonly List<string> _planFolderIds = new List<string>();
        readonly Dictionary<string, HashSet<string>> _planItems = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        readonly HashSet<string> _planWhole = new HashSet<string>(StringComparer.Ordinal);   // folders asked whole (all their items)

        void PrepareScope(List<FolderInfo> tree)
        {
            if (_o.Plan == null) return;
            var byId = tree.ToDictionary(f => f.Id, StringComparer.Ordinal);
            var subtrees = new HashSet<string>(StringComparer.Ordinal);
            foreach (ExportPlanFolder p in _o.Plan.Folders)
            {
                _planFolderIds.Add(p.ExportFolderId);
                _planScope.Add(p.ExportFolderId);
                if (p.IncludeSubfolders) subtrees.Add(p.ExportFolderId);
                if (p.ItemIds == null) _planWhole.Add(p.ExportFolderId);
                if (p.ItemIds != null)
                {
                    HashSet<string> set;
                    if (!_planItems.TryGetValue(p.ExportFolderId, out set)) _planItems[p.ExportFolderId] = set = new HashSet<string>(StringComparer.Ordinal);
                    foreach (string id in p.ItemIds) set.Add(id);
                }
            }
            // a folder below a folder whose subtree is asked
            foreach (FolderInfo f in tree)
            {
                var seen = new HashSet<string>(StringComparer.Ordinal);
                FolderInfo cursor = f;
                while (cursor != null && seen.Add(cursor.Id))
                {
                    if (subtrees.Contains(cursor.Id))
                    {
                        if (_recoverableLeftOut.Contains(f.Id) && !_planFolderIds.Contains(f.Id)) break;
                        _planScope.Add(f.Id); _planWhole.Add(f.Id); break;
                    }
                    FolderInfo parent = null;
                    if (cursor.ParentFolderId != null) byId.TryGetValue(cursor.ParentFolderId, out parent);
                    cursor = parent;
                }
            }
        }

        /// <summary>Is this item asked for? Every item of a folder asked whole (or below a folder whose subtree is asked), only the named ones otherwise.</summary>
        bool ItemWanted(string folderId, string itemId)
        {
            HashSet<string> named;
            return _o.Plan == null || _planWhole.Contains(folderId) || (_planItems.TryGetValue(folderId, out named) && named.Contains(itemId));
        }

        /// <summary>The ids of the folders of <paramref name="scope"/> and of the folders above them (the path to each one).</summary>
        static HashSet<string> FoldersOnPath(List<FolderInfo> tree, List<FolderInfo> scope)
        {
            var byId = new Dictionary<string, FolderInfo>(StringComparer.Ordinal);
            foreach (FolderInfo f in tree) byId[f.Id] = f;
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (FolderInfo f in scope)
            {
                FolderInfo cursor = f;
                while (cursor != null && ids.Add(cursor.Id))
                {
                    FolderInfo parent = null;
                    if (cursor.ParentFolderId != null) byId.TryGetValue(cursor.ParentFolderId, out parent);
                    cursor = parent;
                }
            }
            return ids;
        }

        bool InScope(FolderInfo f)
        {
            if (_o.Plan != null) return _planScope.Contains(f.Id);
            if (string.IsNullOrEmpty(_o.FolderPath)) return true;
            string p = _o.FolderPath.Trim('/');
            return string.Equals(f.Path, p, StringComparison.OrdinalIgnoreCase) || f.Path.StartsWith(p + "/", StringComparison.OrdinalIgnoreCase);
        }

        // ---------------------------------------------------------------------------------------------
        //  Inventory
        // ---------------------------------------------------------------------------------------------

        void Inventory(List<FolderInfo> folders, CancellationToken ct)
        {
            // Main archive: every folder is listed, an expanded folder (its items live in an auxiliary archive) may announce no item.
            bool listAll = _kind == StoreMailboxKind.Archive;
            var todo = new ConcurrentQueue<FolderInfo>(folders.Where(f => listAll || f.TotalItemCount > 0));
            long listed = 0;
            RunWorkers(_o.Concurrency, n =>
            {
                FolderInfo f;
                while (!ct.IsCancellationRequested && todo.TryDequeue(out f))
                {
                    try
                    {
                        using (Limiter.Enter(ct))
                        {
                            long count = 0;
                            Tuple<string, string> owner;
                            _folderOwner.TryGetValue(f.Id, out owner);
                            string recordedFolderId = owner == null ? f.Id : owner.Item2;
                            string recordedFolderMailbox = owner == null ? _o.MailboxId : owner.Item1;
                            var batch = new List<StoreItemEntry>();
                            Dictionary<string, List<StoreItem>> recorded = null;
                            bool inAux = false;
                            try
                            {
                                bool once = _o.NotReadyTrip > 0 && Volatile.Read(ref _notReadyStreak) >= _o.NotReadyTrip;
                                foreach (ItemInfo i in _api.EnumerateItems(_o.MailboxId, f.Id, _o.ItemPageSize, ct, once ? 1 : 0))
                                {
                                    if (!ItemWanted(f.Id, i.Id)) continue;
                                    string real = string.IsNullOrEmpty(i.MailboxId) ? _o.MailboxId : i.MailboxId;
                                    if (!MailboxApi.SameMailbox(real, _o.MailboxId))
                                    {
                                        inAux = true;
                                        _learner.Learn(real, "Redirect308", "the listing of the items of folder '" + f.Path + "' of " + _o.MailboxId + " was answered by this mailbox");
                                        lock (_result.AuxMailboxes) if (!_result.AuxMailboxes.Contains(real, StringComparer.OrdinalIgnoreCase)) _result.AuxMailboxes.Add(real);
                                    }
                                    // the folder is recorded once (where its tree is: the main archive), the item with the mailbox that holds it
                                    string folderMailbox = owner != null ? owner.Item1 : (MailboxApi.SameMailbox(real, _o.MailboxId) ? null : _o.MailboxId);
                                    batch.Add(new StoreItemEntry { MailboxId = real, ItemId = i.Id, FolderId = recordedFolderId, FolderMailboxId = folderMailbox, ChangeKey = i.ChangeKey, ItemType = i.Type, Size = i.Size });
                                    count++;
                                    if (batch.Count >= 2000) { _store.UpsertItems(OnePerMessage(batch, ref recorded, recordedFolderMailbox, recordedFolderId)); batch = new List<StoreItemEntry>(); }
                                }
                                Interlocked.Exchange(ref _notReadyStreak, 0);   // listed to the end (empty or not): the location answers again
                            }
                            catch (ContentNotAvailableException ex)
                            {
                                // not empty, not failed for good: the service cannot read the content yet (folder being moved). What was read is kept; the folder is reported.
                                var d = new DeferredFolder { Path = f.Path, MailboxId = _o.MailboxId, FolderId = f.Id, StatusCode = ex.StatusCode, ErrorCode = ex.ErrorCode, Message = ex.ServiceMessage, Attempts = ex.Attempts };
                                lock (_result.FoldersDeferred) _result.FoldersDeferred.Add(d);
                                if (ex.Attempts == 1 && _o.NotReadyTrip > 0 && _api.ContentAttempts > 1) { Interlocked.Increment(ref _fastDeferred); _result.FoldersDeferredFast = _fastDeferred; }
                                int streak = Interlocked.Increment(ref _notReadyStreak);
                                if (_o.NotReadyTrip > 0 && streak == _o.NotReadyTrip)
                                    Log("the last " + streak + " folders of this location were not readable yet: the next ones are asked once (not " + _api.ContentAttempts + " times) and deferred; a folder that lists puts the full asks back");
                                Problem("folder " + d);
                            }
                            if (batch.Count > 0) _store.UpsertItems(OnePerMessage(batch, ref recorded, recordedFolderMailbox, recordedFolderId));
                            if (inAux)
                            {
                                bool first;
                                lock (_movedFolders) first = _movedFolders.Add(f.Id);
                                if (first)
                                {
                                    Interlocked.Increment(ref _foldersMoved);
                                    Log("folder '" + f.Path + "' is an expanded folder: its items live in an auxiliary archive (recorded once in the main archive tree, items with their real mailbox)");
                                }
                            }
                            Interlocked.Add(ref listed, count);
                        }
                    }
                    catch (GraphException ex)
                    {
                        if (ex.StatusCode == 401 || ex.StatusCode == 403) throw;
                        Interlocked.Increment(ref _foldersNotListed);
                        string why = "folder '" + f.Path + "' not listed: " + ex.Message;
                        lock (_result.Problems) _result.Problems.Add(why);
                        Log(why);
                    }
                }
            });
            _result.Inventoried = listed;
            _result.FoldersMoved = _foldersMoved;
            _result.FoldersNotListed = _foldersNotListed;
            Log("inventory: " + listed + " items listed in " + folders.Count(f => f.TotalItemCount > 0) + " folders" +
                (_foldersNotListed > 0 ? ", " + _foldersNotListed + " folders could not be listed" : "") +
                (_result.FoldersDeferred.Count > 0 ? ", " + _result.FoldersDeferred.Count + " folder(s) deferred (content not available yet)" : ""));
        }

        /// <summary>
        /// The same message listed on the two sides of a move has two ids (the mailbox part differs) and one message entry key. The record that exists for it under
        /// another mailbox is kept when its bytes are exported (the new listing is dropped), otherwise it follows the message to the mailbox that lists it now.
        /// The same id in two mailboxes is not handled here: the store moves that record itself (UpsertItems).
        /// </summary>
        List<StoreItemEntry> OnePerMessage(List<StoreItemEntry> batch, ref Dictionary<string, List<StoreItem>> recorded, string folderMailbox, string folderId)
        {
            if (_kind == StoreMailboxKind.Primary) return OnePerMessageInMailbox(batch);
            if (recorded == null)
            {
                recorded = new Dictionary<string, List<StoreItem>>(StringComparer.Ordinal);
                // every folder row of the store that is this folder: the main archive's and an auxiliary archive's own (an auxiliary archive exported alone first records its folders itself)
                string key = ExchangeEntryId.FolderKeyOf(folderId);
                var rows = new List<long>();
                long? mine = _store.FindFolderKey(folderMailbox, folderId);
                if (mine.HasValue) rows.Add(mine.Value);
                lock (_rowsLock)
                {
                    if (_rowsByKey == null)
                    {
                        _rowsByKey = new Dictionary<string, List<long>>(StringComparer.Ordinal);
                        foreach (StoreMailbox m in _store.GetMailboxes().Where(x => x.Kind == StoreMailboxKind.Archive || x.Kind == StoreMailboxKind.AuxArchive))
                            foreach (StoreFolder sf in _store.GetFolders(m.Id))
                            {
                                List<long> l;
                                string k = ExchangeEntryId.FolderKeyOf(sf.FolderId);
                                if (!_rowsByKey.TryGetValue(k, out l)) _rowsByKey[k] = l = new List<long>();
                                l.Add(sf.Key);
                            }
                    }
                    List<long> same;
                    if (_rowsByKey.TryGetValue(key, out same)) foreach (long k2 in same) if (!rows.Contains(k2)) rows.Add(k2);
                }
                foreach (long row in rows)
                    foreach (StoreItem r in _store.GetFolderItemRefs(row))
                    {
                        List<StoreItem> l;
                        string k = ExchangeEntryId.MessageKeyOf(r.ItemId);
                        if (!recorded.TryGetValue(k, out l)) recorded[k] = l = new List<StoreItem>();
                        l.Add(r);
                    }
            }
            if (recorded.Count == 0) return batch;
            var keep = new List<StoreItemEntry>(batch.Count);
            foreach (StoreItemEntry e in batch)
            {
                List<StoreItem> same;
                StoreItem other = recorded.TryGetValue(ExchangeEntryId.MessageKeyOf(e.ItemId), out same)
                    ? same.FirstOrDefault(r => !MailboxApi.SameMailbox(r.MailboxId, e.MailboxId) && !string.Equals(r.ItemId, e.ItemId, StringComparison.Ordinal)) : null;
                if (other != null)
                {
                    if (other.State == StoreItemState.Exported) { Interlocked.Increment(ref _dedup); _result.ItemsDeduplicated = Interlocked.Read(ref _dedup); continue; }
                    if (_store.RehomeItem(other.Key, e.MailboxId, e.ItemId)) { Interlocked.Increment(ref _rehomed); _result.ItemsRehomed = Interlocked.Read(ref _rehomed); }
                }
                keep.Add(e);
            }
            return keep;
        }

        /// <summary>
        /// One record per message in the primary mailbox too: a message moved to another folder while the export runs has a new item id and the same message
        /// entry key (database guid and global counter), so the listing of its new folder would record it a second time. The record that already exists
        /// is kept when its bytes are exported (the new listing is dropped), otherwise it takes the new id and the new listing updates it.
        /// </summary>
        List<StoreItemEntry> OnePerMessageInMailbox(List<StoreItemEntry> batch)
        {
            lock (_mailboxKeysLock)
            {
                if (_mailboxKeys == null)
                {
                    _mailboxKeys = new Dictionary<string, KeyValuePair<string, string>>(StringComparer.Ordinal);
                    foreach (StoreFolder sf in _store.GetFolders(_o.MailboxId))
                        foreach (StoreItem r in _store.GetFolderItemRefs(sf.Key))
                        {
                            string rk = EntryKeyOf(r.ItemId);
                            if (rk != null) _mailboxKeys[rk] = new KeyValuePair<string, string>(r.MailboxId, r.ItemId);
                        }
                }
                var keep = new List<StoreItemEntry>(batch.Count);
                foreach (StoreItemEntry e in batch)
                {
                    string k = EntryKeyOf(e.ItemId);
                    KeyValuePair<string, string> known;
                    if (k != null && _mailboxKeys.TryGetValue(k, out known) && !string.Equals(known.Value, e.ItemId, StringComparison.Ordinal) && MailboxApi.SameMailbox(known.Key, e.MailboxId))
                    {
                        long? otherKey = _store.FindItemKey(known.Key, known.Value);
                        StoreItem other = otherKey.HasValue ? _store.GetItem(otherKey.Value) : null;
                        if (other != null)
                        {
                            if (other.State == StoreItemState.Exported) { Interlocked.Increment(ref _dedup); _result.ItemsDeduplicated = Interlocked.Read(ref _dedup); continue; }
                            if (_store.RehomeItem(other.Key, e.MailboxId, e.ItemId)) { Interlocked.Increment(ref _rehomed); _result.ItemsRehomed = Interlocked.Read(ref _rehomed); }
                        }
                    }
                    if (k != null) _mailboxKeys[k] = new KeyValuePair<string, string>(e.MailboxId, e.ItemId);
                    keep.Add(e);
                }
                return keep;
            }
        }

        /// <summary>The identity of the message alone (not of the folder it is in): null for an id that is not in the Exchange format.</summary>
        static string EntryKeyOf(string itemId)
        {
            ExchangeEntryId e = ExchangeEntryId.Parse(itemId);
            return e == null || e.MessageKey == null ? null : e.MessageKey;
        }

        readonly object _mailboxKeysLock = new object();
        Dictionary<string, KeyValuePair<string, string>> _mailboxKeys;

        long _dedup, _rehomed;
        readonly object _rowsLock = new object();
        Dictionary<string, List<long>> _rowsByKey;

        int _notReadyStreak, _fastDeferred;
        int _foldersNotListed;
        int _foldersMoved;

        // ---------------------------------------------------------------------------------------------
        //  Export
        // ---------------------------------------------------------------------------------------------

        void ExportPending(CancellationToken ct)
        {
            HashSet<long> allowed = null;
            Dictionary<long, string> folderIds = null;
            if (_o.Plan != null)
            {
                // the store may hold items of earlier runs outside the selection: only the folders in scope, and the named items of the others
                List<StoreFolder> storeFolders = _store.GetFolders(_o.MailboxId);
                folderIds = storeFolders.ToDictionary(f => f.Key, f => f.FolderId);
                allowed = new HashSet<long>(storeFolders.Where(f => _planScope.Contains(f.FolderId)).Select(f => f.Key));
                // folders of this auxiliary archive that the main archive tree records: their rows are the main archive's, the plan names the auxiliary ids
                foreach (var kv in _folderOwner)
                {
                    if (!_planScope.Contains(kv.Key)) continue;
                    long? k = _store.FindFolderKey(kv.Value.Item1, kv.Value.Item2);
                    if (k.HasValue) { allowed.Add(k.Value); folderIds[k.Value] = kv.Key; }
                }
            }
            else if (!string.IsNullOrEmpty(_o.FolderPath))
            {
                string p = _o.FolderPath.Trim('/');
                allowed = new HashSet<long>(_store.GetFolders(_o.MailboxId)
                    .Where(f => string.Equals(f.Path, p, StringComparison.OrdinalIgnoreCase) || (f.Path != null && f.Path.StartsWith(p + "/", StringComparison.OrdinalIgnoreCase)))
                    .Select(f => f.Key));
                foreach (var kv in _folderOwner)
                {
                    FolderInfo f = _tree.FirstOrDefault(t => t.Id == kv.Key);
                    if (f == null || !InScope(f)) continue;
                    long? k = _store.FindFolderKey(kv.Value.Item1, kv.Value.Item2);
                    if (k.HasValue) allowed.Add(k.Value);
                }
            }

            PlanFolders(allowed, folderIds, ct);
            var queue = new BlockingCollection<List<StoreItem>>(_o.Concurrency * 2);
            var producer = Task.Factory.StartNew(() =>
            {
                try
                {
                    // one request asks one mailbox: the main archive first, then each auxiliary archive that holds items of it
                    foreach (string source in ExportMailboxes())
                    {
                        long after = 0;
                        var batch = new List<StoreItem>();
                        long batchBytes = 0;
                        while (!ct.IsCancellationRequested)
                        {
                            List<StoreItem> page = _store.GetPendingItems(source, 500, after);
                            if (page.Count == 0) break;
                            after = page[page.Count - 1].Key;
                            foreach (StoreItem it in page)
                            {
                                if (allowed != null && !(it.FolderKey.HasValue && allowed.Contains(it.FolderKey.Value))) continue;
                                if (folderIds != null && !ItemWanted(folderIds[it.FolderKey.Value], it.ItemId)) continue;
                                if (batch.Count > 0 && (batch.Count >= _o.BatchItems || batchBytes + it.Size > _o.BatchBytes))
                                {
                                    queue.Add(batch, ct);
                                    batch = new List<StoreItem>(); batchBytes = 0;
                                }
                                batch.Add(it); batchBytes += it.Size;
                            }
                        }
                        if (batch.Count > 0 && !ct.IsCancellationRequested) queue.Add(batch, ct);
                    }
                }
                catch (OperationCanceledException) { }
                catch (Exception ex) { Fatal(ex); }
                finally { queue.CompleteAdding(); }
            }, TaskCreationOptions.LongRunning);

            RunWorkers(_o.Concurrency, n =>
            {
                string tmp = Path.Combine(_o.TempPath, "export-tmp", "w" + n);
                if (Directory.Exists(tmp)) Directory.Delete(tmp, true);
                Directory.CreateDirectory(tmp);
                try
                {
                    foreach (List<StoreItem> batch in queue.GetConsumingEnumerable(ct))
                    {
                        using (Limiter.Enter(ct)) ExportBatch(batch, tmp, ct);
                    }
                }
                finally { try { Directory.Delete(tmp, true); } catch (IOException) { } }
            });
            if (ct.IsCancellationRequested) { try { while (queue.TryTake(out _)) { } } catch (ObjectDisposedException) { } }
            producer.Wait();
        }

        StoreFolderLabels _labels;
        string FolderLabel(StoreItem it) { if (_labels == null) _labels = new StoreFolderLabels(_store); return _labels.Label(it.FolderKey); }

        /// <summary>Tells the progress per folder how many items each folder has to export in this run (the same items the producer will queue).</summary>
        void PlanFolders(HashSet<long> allowed, Dictionary<long, string> folderIds, CancellationToken ct)
        {
            if (_labels == null) _labels = new StoreFolderLabels(_store);
            foreach (string source in ExportMailboxes())
            {
                if (folderIds == null)
                {
                    foreach (StoreFolderCount c in _store.GetPendingByFolder(source))
                        if (allowed == null || allowed.Contains(c.FolderKey)) Metrics.Folders.Expect(_labels.Label(c.FolderKey), c.Count);
                    continue;
                }
                // a plan can name single items of a folder: count them one by one
                long after = 0;
                while (!ct.IsCancellationRequested)
                {
                    List<StoreItem> page = _store.GetPendingItems(source, 500, after);
                    if (page.Count == 0) break;
                    after = page[page.Count - 1].Key;
                    foreach (StoreItem it in page)
                    {
                        if (allowed != null && !(it.FolderKey.HasValue && allowed.Contains(it.FolderKey.Value))) continue;
                        if (!ItemWanted(folderIds[it.FolderKey.Value], it.ItemId)) continue;
                        Metrics.Folders.Expect(_labels.Label(it.FolderKey), 1);
                    }
                }
            }
        }

        void ExportBatch(List<StoreItem> items, string tmp, CancellationToken ct)
        {
            List<ExportEntry> entries;
            try
            {
                entries = _api.ExportItems(items[0].MailboxId, items.Select(i => i.ItemId).ToList(), tmp, 0, ct);
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                CleanDirectory(tmp);
                var ge = ex as GraphException;
                if (ge != null && (ge.StatusCode == 401 || ge.StatusCode == 403)) throw;
                if (ct.IsCancellationRequested) throw new OperationCanceledException();
                if (items.Count > 1)
                {
                    Log("request for " + items.Count + " items failed (" + Short(ex.Message) + "), asking again item by item");
                    foreach (StoreItem it in items) ExportBatch(new List<StoreItem> { it }, tmp, ct);
                    return;
                }
                Fail(items[0], ex.Message);
                return;
            }

            var byId = items.ToDictionary(i => i.ItemId, StringComparer.Ordinal);
            var handled = new HashSet<long>();
            try
            {
                foreach (ExportEntry e in entries)
                {
                    StoreItem it;
                    if (e.ItemId == null || !byId.TryGetValue(e.ItemId, out it)) { Log("the answer holds an item that was not asked for, ignored"); continue; }
                    handled.Add(it.Key);
                    if (e.HasError || e.FilePath == null) { Fail(it, ((e.ErrorCode ?? "") + " " + (e.ErrorMessage ?? "no data")).Trim()); continue; }
                    StoreWriteResult w;
                    using (var fs = new FileStream(e.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.SequentialScan))
                        w = _store.WriteItem(it.Key, fs, ct, true);
                    if (!string.Equals(w.Sha256, e.Sha256, StringComparison.OrdinalIgnoreCase))
                        Problem("item " + it.Key + ": the SHA-256 of the stored bytes differs from the one computed on arrival");
                    FxReadResult read = FxReader.ReadFile(e.FilePath);
                    if (read.Error != null || read.Truncated)
                    {
                        Interlocked.Increment(ref _previewWarnings);
                        _result.PreviewWarnings = _previewWarnings;
                        Problem("item " + it.Key + ": the FastTransfer reader reports " + (read.Error ?? "a truncated stream") + " (the stored bytes are the exported ones)");
                    }
                    _store.SetPreview(it.Key, StorePreviews.From(FxMessage.Summarize(read), it.Key));
                    _store.UpdateChangeKey(it.Key, e.ChangeKey);
                    File.Delete(e.FilePath);
                    // asked from the main archive, exported from an auxiliary one (ErrorArchiveFolderMovedPermanently): the item lives there
                    string real = MailboxApi.MailboxIdOfUrl(e.ExportUrl);
                    if (real != null && !MailboxApi.SameMailbox(real, it.MailboxId))
                    {
                        _learner.Learn(real, "ExportMoved", "item " + it.Key + " of " + it.MailboxId + " was exported from this mailbox");
                        if (_store.MoveItemToMailbox(it.Key, real)) Interlocked.Increment(ref _itemsMoved);
                        else Problem("item " + it.Key + " lives in " + real + " but that mailbox already holds a record of it: left as it is");
                        _result.ItemsMoved = Interlocked.Read(ref _itemsMoved);
                    }
                    Metrics.ItemDone(w.Length);
                    Metrics.Folders.Done(FolderLabel(it), w.Length);
                    Succeeded();
                }
            }
            finally { CleanDirectory(tmp); }
            foreach (StoreItem it in items) if (!handled.Contains(it.Key)) Fail(it, "not returned by exportItems");
        }

        int _previewWarnings;
        long _itemsMoved;

        void Fail(StoreItem it, string message)
        {
            _store.MarkItemFailed(it.Key, message);
            Metrics.ItemFailed();
            Metrics.Folders.Fail(FolderLabel(it));
            Log("item " + it.Key + " failed: " + Short(message));
            FailedOnce(message);
        }

        void Problem(string text)
        {
            lock (_result.Problems) _result.Problems.Add(text);
            Log(text);
        }

        static string Short(string s) { return s != null && s.Length > 300 ? s.Substring(0, 300) + "..." : s; }

        static void CleanDirectory(string dir)
        {
            foreach (string f in Directory.EnumerateFiles(dir)) { try { File.Delete(f); } catch (IOException) { } }
        }
    }
}
