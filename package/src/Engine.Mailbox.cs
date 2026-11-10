// =============================================================================
//  Mailbox Transfer - engine, part 3: mailbox import and export API
// -----------------------------------------------------------------------------
//  Microsoft Graph "mailbox import and export" (v1.0):
//    GET  /admin/exchange/mailboxes/{mailboxId}/folders
//    GET  /admin/exchange/mailboxes/{mailboxId}/folders/{folderId}/childFolders
//    GET  /admin/exchange/mailboxes/{mailboxId}/folders/{folderId}/items
//    POST /admin/exchange/mailboxes/{mailboxId}/exportItems          (20 item ids at most)
//    POST /admin/exchange/mailboxes/{mailboxId}/createImportSession  -> importUrl (pre-authenticated)
//    POST <importUrl>                                                 (no Authorization header)
//  A mailbox id is the physical id "MBX:<mailbox guid>@<tenant guid>" (primary or archive mailbox).
// =============================================================================
using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer
{
    public sealed class FolderInfo
    {
        public string Id;
        public string DisplayName;
        public string ParentFolderId;
        public string WellKnownName;
        public string Type;
        public int ChildFolderCount;
        public int TotalItemCount;
        /// <summary>Slash separated names from the top of the tree, filled by GetFolderTree.</summary>
        public string Path;
        public int Depth;
        /// <summary>Sum of the item sizes in bytes, -1 until measured.</summary>
        public long TotalSize = -1;
        /// <summary>The mailbox that answered, when it is not the one asked (the request was redirected, 308, to the auxiliary archive the folder moved to); null otherwise. Set by GetFolder.</summary>
        public string MailboxId;
    }

    public sealed class ItemInfo
    {
        public string Id;
        public string Type;
        public string ChangeKey;
        public long Size;
        public DateTimeOffset? Created;
        public DateTimeOffset? Modified;
        /// <summary>
        /// The mailbox that answered the listing this item comes from. It differs from the mailbox asked when the folder is an expanded
        /// folder of an auto-expanding archive (308 to the auxiliary archive): the item then lives in that auxiliary archive.
        /// </summary>
        public string MailboxId;
    }

    /// <summary>
    /// The listing of the items of a folder cannot be read now, and the service says it is a state, not a refusal: the folder is being moved to an
    /// auxiliary archive of an auto-expanding archive (the folder skeleton exists in the auxiliary archive before its content: 500 ErrorInternalServerError
    /// "Cannot query rows in a table", measured on the lab tenant 2026-10-31), or the mailbox is moving / its store is unavailable (EWS error codes
    /// ErrorMailboxMoveInProgress, ErrorMailboxStoreUnavailable, ErrorInternalServerTransientError). It is thrown after the bounded retries of
    /// <see cref="MailboxApi.EnumerateItems"/>; the folder is deferred, not empty, and not failed for good: ask again later.
    /// </summary>
    public sealed class ContentNotAvailableException : Exception
    {
        public string MailboxId { get; }
        public string FolderId { get; }
        public int StatusCode { get; }
        public string ErrorCode { get; }
        /// <summary>The message of the service, as it came (without the URL).</summary>
        public string ServiceMessage { get; }
        /// <summary>How many times the listing was asked (the first try included).</summary>
        public int Attempts { get; }

        public ContentNotAvailableException(string mailboxId, string folderId, GraphException cause, int attempts)
            : base("The content of folder " + folderId + " in " + mailboxId + " is not available yet (HTTP " + cause.StatusCode + " " + cause.ErrorCode + ": " + MailboxApi.ServiceMessageOf(cause) + "), asked " + attempts + " time(s).", cause)
        {
            MailboxId = mailboxId; FolderId = folderId; StatusCode = cause.StatusCode; ErrorCode = cause.ErrorCode;
            ServiceMessage = MailboxApi.ServiceMessageOf(cause); Attempts = attempts;
        }
    }

    /// <summary>
    /// The mailbox has no route in Exchange Online: a soft-deleted mailbox (the user was deleted less than 30 days ago) or an inactive mailbox (kept on hold
    /// after the user was deleted) cannot be reached by the mailbox import-export API. Measured live 2026-11-09: the folder listing of such an id answers a bare
    /// 500 (no code, no message), or 404 ErrorInvalidUser. The message says what to do, in the words of the administrator.
    /// </summary>
    public sealed class MailboxNotReachableException : Exception
    {
        public string MailboxId { get; }
        public int StatusCode { get; }
        public string ErrorCode { get; }

        public MailboxNotReachableException(string mailboxId, GraphException cause)
            : base(MessageFor(mailboxId, cause), cause)
        {
            MailboxId = mailboxId; StatusCode = cause.StatusCode; ErrorCode = cause.ErrorCode;
        }

        /// <summary>Nothing is exported from this mailbox as it is; the supported routes are in the text.</summary>
        public static string MessageFor(string mailboxId, GraphException cause)
        {
            return "The mailbox " + mailboxId + " cannot be reached by the mailbox import-export API (the service answered " + cause.StatusCode + (string.IsNullOrEmpty(cause.ErrorCode) ? " with no explanation" : " " + cause.ErrorCode) + "). "
                + "This is what happens to a soft-deleted mailbox (its user was deleted less than 30 days ago) and to an inactive mailbox (kept on hold after its user was deleted): they have no route in Exchange Online, "
                + "so the API cannot read them, and neither can this tool. Supported ways: for a soft-deleted mailbox, restore the user within 30 days (Microsoft 365 admin center, Deleted users, "
                + "or the Exchange Online cmdlet Undo-SoftDeletedMailbox) and export it once it is active again; for an inactive mailbox, recover it into an active mailbox "
                + "(restore its content into an active mailbox with New-MailboxRestoreRequest, or recover it as a new mailbox) and export that active mailbox. "
                + "If the mailbox is active, check the mailbox id (Get-Mailbox -Identity <user> | Format-List ExchangeGuid) and try again.";
        }
    }

    /// <summary>The folders of Recoverable Items read through the import-export API: the root and the chosen folders below it, with their paths ("Recoverable Items/Deletions").</summary>
    public sealed class RecoverableTree
    {
        public List<FolderInfo> Folders = new List<FolderInfo>();
        /// <summary>Names asked for that the mailbox does not have.</summary>
        public List<string> Missing = new List<string>();
        /// <summary>The folders of the mailbox below the root that were left out because they were not asked for.</summary>
        public List<string> Skipped = new List<string>();
    }

    /// <summary>
    /// Recoverable Items (what a user deleted from Deleted Items, kept for a while; what a hold keeps; the versions of changed items). Measured live 2026-11-09 on the
    /// import-export API: the top-level folder listing does NOT list Recoverable Items, but GET .../folders/{id}, .../items and .../childFolders work on its root and on
    /// the folders below it (well-known names recoverableitemsroot, recoverableitemsdeletions, ... work as folder ids); exportItems takes the ids of their items
    /// (full fidelity, attachments included; the time of the deletion is the last modification time of the item). The folders below the root are Calendar Logging,
    /// Deletions, DiscoveryHolds, Purges, SubstrateHolds and Versions.
    ///
    /// In the store and in a PST they live under one top folder, "Recoverable Items", never mixed with the normal folders. An import of such an export never goes
    /// back into Recoverable Items (nothing can be imported there): the folders below the top folder are created in a normal folder, "Restored deleted items".
    /// </summary>
    public static class RecoverableItems
    {
        public const string TopName = "Recoverable Items";
        public const string RestoredName = "Restored deleted items";
        public const string RootWellKnown = "recoverableitemsroot";
        /// <summary>What -IncludeRecoverableItems exports by default.</summary>
        public static readonly string[] DefaultFolders = { "Deletions", "Purges", "Versions", "DiscoveryHolds" };
        /// <summary>The folders of the root as the service lists them. Calendar Logging and SubstrateHolds are Microsoft's own copies: only on request.</summary>
        public static readonly string[] KnownFolders = { "Calendar Logging", "Deletions", "DiscoveryHolds", "Purges", "SubstrateHolds", "Versions" };

        /// <summary>Is this path the top folder or below it?</summary>
        public static bool IsPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            string p = path.Trim('/');
            return string.Equals(p, TopName, StringComparison.OrdinalIgnoreCase) || p.StartsWith(TopName + "/", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>The folders to take: the explicit list, else the default set.</summary>
        public static string[] Chosen(IEnumerable<string> explicitList)
        {
            string[] list = explicitList == null ? new string[0] : explicitList.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim().Trim('/')).ToArray();
            return list.Length == 0 ? DefaultFolders : list;
        }

        /// <summary>The explicit list is a list of names of the folders of the root; anything else is refused with the names that exist.</summary>
        public static void CheckNames(IEnumerable<string> explicitList)
        {
            if (explicitList == null) return;
            foreach (string n in explicitList)
            {
                if (string.IsNullOrWhiteSpace(n)) continue;
                string name = n.Trim().Trim('/');
                if (name.StartsWith(TopName + "/", StringComparison.OrdinalIgnoreCase)) name = name.Substring(TopName.Length + 1);
                if (!KnownFolders.Any(k => string.Equals(k, name, StringComparison.OrdinalIgnoreCase)))
                    throw new ArgumentException("'" + n + "' is not a folder of Recoverable Items. The folders are: " + string.Join(", ", KnownFolders) + ".");
            }
        }

        /// <summary>The explicit list with an optional "Recoverable Items/" in front of each name removed.</summary>
        public static string[] Names(IEnumerable<string> explicitList)
        {
            if (explicitList == null) return null;
            return explicitList.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x =>
            {
                string name = x.Trim().Trim('/');
                return name.StartsWith(TopName + "/", StringComparison.OrdinalIgnoreCase) ? name.Substring(TopName.Length + 1) : name;
            }).ToArray();
        }

        /// <summary>Is this folder (by its path) one the names ask for? The top folder is always wanted; a folder below a wanted one too; anything that is not Recoverable Items is not concerned (true).</summary>
        public static bool Wanted(string path, IEnumerable<string> names)
        {
            if (!IsPath(path)) return true;
            string p = path.Trim('/');
            if (string.Equals(p, TopName, StringComparison.OrdinalIgnoreCase)) return true;
            if (names == null) return true;
            string first = p.Substring(TopName.Length + 1);
            int slash = first.IndexOf('/');
            if (slash >= 0) first = first.Substring(0, slash);
            return names.Any(n => string.Equals(n, first, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// The target path of a source path of the import: a path in Recoverable Items becomes "Restored deleted items" (and the folders below it keep their names).
        /// Any other path is returned as it is.
        /// </summary>
        public static string Restore(string path)
        {
            if (!IsPath(path)) return path;
            string p = path.Trim('/');
            return p.Length == TopName.Length ? RestoredName : RestoredName + p.Substring(TopName.Length);
        }
    }

    /// <summary>
    /// The id of a folder or a message as Exchange Online gives it (REST id, base64url, "AAMk..."), taken apart. Verified on the lab tenant 2026-10-31
    /// (folders of the main archive and of its first auxiliary archive, 89 bytes each):
    ///   00 03 | length (UInt16 LE, 36) | the mailbox guid as text | 00 | length (UInt16 LE, 46 for a folder, 70 for a message) | EntryID
    ///   EntryID of a folder (MS-OXCDATA 2.2.4.1): flags (4, zero) | provider uid (16, differs per mailbox) | folder type (2, 1 = private) |
    ///       database guid (16) | global counter (6, the folder number) | pad (2)
    ///   EntryID of a message (MS-OXCDATA 2.2.4.2): flags (4) | provider uid (16) | message type (2, 7 = private) | folder database guid (16) |
    ///       folder global counter (6) | pad (2) | message database guid (16) | message global counter (6) | pad (2)
    /// A folder that exists in the main archive and, as a skeleton or moved, in an auxiliary archive has the same folder key in both: only the mailbox guid
    /// and the provider uid differ. Message keys follow the same logic (whether Exchange keeps the message part when it moves a message is not verified).
    /// </summary>
    public sealed class ExchangeEntryId
    {
        public string MailboxGuid;
        public bool IsFolder { get { return MessageKey == null; } }
        /// <summary>The provider uid of the entry (hex): per mailbox, not part of the identity of the folder.</summary>
        public string ProviderUid;
        public int EntryType;
        /// <summary>Identity of the folder, the same in every mailbox that holds it: "database guid-global counter" (hex). For a message id, the folder it is in.</summary>
        public string FolderKey;
        /// <summary>Identity of the message ("database guid-global counter", hex), null for a folder id.</summary>
        public string MessageKey;

        /// <summary>The parts of an id, or null when it is not in this format (a test id, an EWS id of another layout, an immutable id).</summary>
        public static ExchangeEntryId Parse(string id)
        {
            if (string.IsNullOrEmpty(id) || id.Length < 60) return null;
            byte[] b;
            try
            {
                string s = id.TrimEnd('=').Replace('-', '+').Replace('_', '/');
                s = s + new string('=', (4 - s.Length % 4) % 4);
                b = Convert.FromBase64String(s);
            }
            catch (FormatException) { return null; }
            if (b.Length < 4 + 36 + 3 || b[0] != 0 || b[1] != 3) return null;
            int guidLen = b[2] | (b[3] << 8);
            if (guidLen != 36 || b.Length < 4 + guidLen + 3) return null;
            Guid g;
            string guidText = Encoding.ASCII.GetString(b, 4, guidLen);
            if (!Guid.TryParse(guidText, out g)) return null;
            int p = 4 + guidLen;
            if (b[p] != 0) return null;
            int entryLen = b[p + 1] | (b[p + 2] << 8);
            p += 3;
            if (b.Length != p + entryLen) return null;
            if (entryLen != 46 && entryLen != 70) return null;
            if (b[p] != 0 || b[p + 1] != 0 || b[p + 2] != 0 || b[p + 3] != 0) return null;
            int type = b[p + 20] | (b[p + 21] << 8);
            if (entryLen == 46 && type != 1) return null;
            if (entryLen == 70 && type != 7) return null;
            var r = new ExchangeEntryId { MailboxGuid = guidText.ToLowerInvariant(), EntryType = type, ProviderUid = Hex(b, p + 4, 16) };
            r.FolderKey = Hex(b, p + 22, 16) + "-" + Hex(b, p + 38, 6);
            if (entryLen == 70) r.MessageKey = Hex(b, p + 46, 16) + "-" + Hex(b, p + 62, 6);
            return r;
        }

        static string Hex(byte[] b, int offset, int count)
        {
            var sb = new StringBuilder(count * 2);
            for (int i = 0; i < count; i++) sb.Append(b[offset + i].ToString("x2"));
            return sb.ToString();
        }

        /// <summary>The identity of the folder an id names (a folder id, or the folder part of a message id); the id itself when it is not in the Exchange format (test ids).</summary>
        public static string FolderKeyOf(string id)
        {
            ExchangeEntryId e = Parse(id);
            return e == null ? id : e.FolderKey;
        }

        /// <summary>The identity of a message id (folder key and message key); the id itself when it is not in the Exchange format.</summary>
        public static string MessageKeyOf(string id)
        {
            ExchangeEntryId e = Parse(id);
            return e == null || e.MessageKey == null ? id : e.FolderKey + "/" + e.MessageKey;
        }

        /// <summary>True when the id is in the Exchange format (the key is then a real entry key, not the id itself).</summary>
        public static bool IsDecodable(string id) { return Parse(id) != null; }
    }

    public sealed class ImportSession
    {
        public string MailboxId;
        public string ImportUrl;
        public DateTimeOffset? Expiration;
    }

    public sealed class ImportResult
    {
        public string ItemId;
        public string ChangeKey;
        public long Bytes;
        public double Seconds;
    }

    public sealed class MailboxApi
    {
        public const int MaxExportBatch = 20;
        static readonly Regex PhysicalId = new Regex(@"^MBX:[0-9A-Fa-f-]+@[0-9A-Fa-f-]+$", RegexOptions.Compiled);
        static readonly Regex MailboxInUrl = new Regex(@"[Mm]ailboxes(?:/|\('?)(?<id>MBX(?::|%3[Aa])[^/'?)]+)", RegexOptions.Compiled);

        public GraphClient Graph { get; }

        /// <summary>
        /// The physical mailbox id inside a Graph or Outlook URL ("/admin/exchange/mailboxes/MBX:x@y/..." or "Mailboxes('MBX:x@y')/importItem"),
        /// or null. Used to learn which mailbox a redirect or an ErrorArchiveFolderMovedPermanently message points to.
        /// </summary>
        public static string MailboxIdOfUrl(string url)
        {
            if (string.IsNullOrEmpty(url)) return null;
            Match m = MailboxInUrl.Match(url);
            if (!m.Success) return null;
            string id = Uri.UnescapeDataString(m.Groups["id"].Value);
            return PhysicalId.IsMatch(id) ? id : null;
        }

        /// <summary>The same id with the case of nothing to compare: physical ids are compared without case.</summary>
        public static bool SameMailbox(string a, string b) { return string.Equals(a, b, StringComparison.OrdinalIgnoreCase); }

        public MailboxApi(GraphClient graph) { Graph = graph; }

        /// <summary>"MBX:guid@guid" is written as is (as in the Microsoft documentation), any other id is escaped.</summary>
        public static string MailboxPath(string mailboxId)
        {
            string id = PhysicalId.IsMatch(mailboxId) ? mailboxId : Uri.EscapeDataString(mailboxId);
            return "/admin/exchange/mailboxes/" + id;
        }

        static string FolderPath(string mailboxId, string folderId)
        {
            return MailboxPath(mailboxId) + "/folders/" + Uri.EscapeDataString(folderId);
        }

        /// <summary>Splits ids in groups of 20, the maximum of one exportItems request.</summary>
        public static List<string[]> Chunk(IEnumerable<string> ids, int size = MaxExportBatch)
        {
            var groups = new List<string[]>();
            var current = new List<string>(size);
            foreach (string id in ids)
            {
                current.Add(id);
                if (current.Count == size) { groups.Add(current.ToArray()); current.Clear(); }
            }
            if (current.Count > 0) groups.Add(current.ToArray());
            return groups;
        }

        // ---------------------------------------------------------------------------------------------
        //  Folders
        // ---------------------------------------------------------------------------------------------

        static FolderInfo ParseFolder(JsonElement e)
        {
            return new FolderInfo
            {
                Id = Json.Str(e, "id"),
                DisplayName = Json.Str(e, "displayName"),
                ParentFolderId = Json.Str(e, "parentFolderId"),
                WellKnownName = Json.Str(e, "wellKnownName"),
                Type = Json.Str(e, "type"),
                ChildFolderCount = (int)Json.Long(e, "childFolderCount"),
                TotalItemCount = (int)Json.Long(e, "totalItemCount")
            };
        }

        /// <summary>
        /// Every folder of a mailbox with its path. The documentation of "List folders" is ambiguous (all folders, or only
        /// the first level), so the listing is completed: a folder whose children were not all returned is asked for its
        /// childFolders. Whatever the service does, the tree is complete.
        /// </summary>
        public List<FolderInfo> GetFolderTree(string mailboxId, CancellationToken ct = default(CancellationToken))
        {
            var byId = new Dictionary<string, FolderInfo>(StringComparer.Ordinal);
            var known = new Dictionary<string, int>(StringComparer.Ordinal);   // parent id -> children seen
            Action<JsonElement> add = e =>
            {
                FolderInfo f = ParseFolder(e);
                if (f.Id == null || byId.ContainsKey(f.Id)) return;
                byId[f.Id] = f;
                if (f.ParentFolderId != null)
                {
                    int n; known.TryGetValue(f.ParentFolderId, out n);
                    known[f.ParentFolderId] = n + 1;
                }
            };

            try { foreach (JsonElement e in Graph.EnumerateValues(Graph.Url(MailboxPath(mailboxId) + "/folders"), ct)) add(e); }
            catch (GraphException ex) when (IsMailboxNotRoutable(ex)) { throw new MailboxNotReachableException(mailboxId, ex); }

            var asked = new HashSet<string>(StringComparer.Ordinal);
            while (true)
            {
                var incomplete = byId.Values.Where(f =>
                {
                    int n; known.TryGetValue(f.Id, out n);
                    return f.ChildFolderCount > n && !asked.Contains(f.Id);
                }).ToList();
                if (incomplete.Count == 0) break;
                foreach (FolderInfo f in incomplete)
                {
                    asked.Add(f.Id);
                    foreach (JsonElement e in Graph.EnumerateValues(Graph.Url(FolderPath(mailboxId, f.Id) + "/childFolders"), ct)) add(e);
                }
            }

            foreach (FolderInfo f in byId.Values) SetPath(f, byId);
            return byId.Values
                .OrderBy(f => f.Path, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        /// <summary>The folders of the first level only (one listing: the service answers the top of the tree, the lazy browse view asks the children as the user opens a folder).</summary>
        public List<FolderInfo> GetRootFolders(string mailboxId, CancellationToken ct = default(CancellationToken))
        {
            var list = new List<FolderInfo>();
            try { foreach (JsonElement e in Graph.EnumerateValues(Graph.Url(MailboxPath(mailboxId) + "/folders"), ct)) list.Add(ParseFolder(e)); }
            catch (GraphException ex) when (IsMailboxNotRoutable(ex)) { throw new MailboxNotReachableException(mailboxId, ex); }
            return list;
        }

        /// <summary>
        /// The service has no route to this mailbox (a soft-deleted or inactive mailbox): a bare 500 (no code, no message) or ErrorInvalidUser, measured live 2026-11-09 on the
        /// folder listing of a physical mailbox id. A 500 that says something (a folder being moved: "Cannot query rows in a table") is not this.
        /// </summary>
        public static bool IsMailboxNotRoutable(GraphException ex)
        {
            if (ex == null) return false;
            if (string.Equals(ex.ErrorCode, "ErrorInvalidUser", StringComparison.OrdinalIgnoreCase)) return true;
            // GraphClient writes "GET url -> 500 <code>: <message>": a bare answer ends right after the status
            return ex.StatusCode == 500 && string.IsNullOrEmpty(ex.ErrorCode) && (ex.Message ?? "").TrimEnd().EndsWith("-> 500", StringComparison.Ordinal);
        }

        /// <summary>
        /// Recoverable Items: the root and the folders below it that <paramref name="names"/> asks for (null: all of them), as folders of a tree whose top is "Recoverable Items"
        /// (paths "Recoverable Items/Deletions"; the root has no parent). The top-level listing of the service does not have them, so they are asked by their well-known name:
        /// GET .../folders/recoverableitemsroot, then its childFolders, then the childFolders of a folder that announces some. A GraphException (404 when the mailbox has no
        /// Recoverable Items) is the caller's to handle.
        /// </summary>
        public RecoverableTree GetRecoverableItemsTree(string mailboxId, IEnumerable<string> names, CancellationToken ct = default(CancellationToken))
        {
            var result = new RecoverableTree();
            var wanted = names == null ? null : names.ToList();
            FolderInfo root = GetFolder(mailboxId, RecoverableItems.RootWellKnown, ct);
            if (string.IsNullOrEmpty(root.Id)) throw new InvalidOperationException("The service did not give the id of the Recoverable Items folder.");
            root.DisplayName = RecoverableItems.TopName; root.ParentFolderId = null; root.Path = RecoverableItems.TopName; root.Depth = 0;
            result.Folders.Add(root);
            var found = new List<string>();
            foreach (FolderInfo c in GetChildFolders(mailboxId, root.Id, ct).OrderBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase))
            {
                ct.ThrowIfCancellationRequested();
                if (wanted != null && !wanted.Any(n => string.Equals(n, c.DisplayName, StringComparison.OrdinalIgnoreCase))) { result.Skipped.Add(c.DisplayName); continue; }
                found.Add(c.DisplayName);
                c.ParentFolderId = root.Id;
                AddRecoverableFolder(mailboxId, c, RecoverableItems.TopName, 1, result.Folders, ct);
            }
            if (wanted != null) foreach (string n in wanted) if (!found.Any(f => string.Equals(f, n, StringComparison.OrdinalIgnoreCase))) result.Missing.Add(n);
            return result;
        }

        void AddRecoverableFolder(string mailboxId, FolderInfo f, string parentPath, int depth, List<FolderInfo> into, CancellationToken ct)
        {
            f.Path = parentPath + "/" + (f.DisplayName ?? "(no name)"); f.Depth = depth;
            into.Add(f);
            if (f.ChildFolderCount > 0 && depth < 8)
                foreach (FolderInfo c in GetChildFolders(mailboxId, f.Id, ct).OrderBy(x => x.DisplayName, StringComparer.OrdinalIgnoreCase))
                {
                    c.ParentFolderId = f.Id;
                    AddRecoverableFolder(mailboxId, c, f.Path, depth + 1, into, ct);
                }
        }

        /// <summary>The children of one folder (GET .../folders/{id}/childFolders).</summary>
        public List<FolderInfo> GetChildFolders(string mailboxId, string folderId, CancellationToken ct = default(CancellationToken))
        {
            var list = new List<FolderInfo>();
            foreach (JsonElement e in Graph.EnumerateValues(Graph.Url(FolderPath(mailboxId, folderId) + "/childFolders"), ct)) list.Add(ParseFolder(e));
            return list;
        }

        /// <summary>One folder by id (GET .../folders/{id}); GraphException 404 when this mailbox has no such folder. Path and Depth are not filled.</summary>
        public FolderInfo GetFolder(string mailboxId, string folderId, CancellationToken ct = default(CancellationToken))
        {
            string answered = null;
            using (JsonDocument doc = Graph.GetJson(Graph.Url(FolderPath(mailboxId, folderId)), ct, u => answered = u))
            {
                FolderInfo f = ParseFolder(doc.RootElement);
                string holder = MailboxIdOfUrl(answered);
                if (holder != null && !SameMailbox(holder, mailboxId)) f.MailboxId = holder;
                return f;
            }
        }

        static void SetPath(FolderInfo f, Dictionary<string, FolderInfo> byId)
        {
            var names = new List<string>();
            var seen = new HashSet<string>();
            FolderInfo cursor = f;
            while (cursor != null && seen.Add(cursor.Id))
            {
                names.Add(cursor.DisplayName ?? "(no name)");
                FolderInfo parent = null;
                if (cursor.ParentFolderId != null) byId.TryGetValue(cursor.ParentFolderId, out parent);
                cursor = parent;
            }
            names.Reverse();
            f.Depth = names.Count - 1;
            f.Path = string.Join("/", names);
        }

        /// <summary>
        /// Creates one folder (create only, nothing is changed or deleted): POST .../folders for a top-level folder,
        /// POST .../folders/{parentFolderId}/childFolders below another one. displayName and type are the two properties
        /// the API takes (MailboxFolder.ReadWrite). Returns the folder with its path.
        /// </summary>
        public FolderInfo CreateFolder(string mailboxId, string displayName, string parentFolderId = null, string type = "IPF.Note",
            string parentPath = null, CancellationToken ct = default(CancellationToken))
        {
            string path = parentFolderId == null ? MailboxPath(mailboxId) + "/folders" : FolderPath(mailboxId, parentFolderId) + "/childFolders";
            string body = JsonSerializer.Serialize(new { displayName = displayName, type = type });
            using (JsonDocument doc = Graph.PostJson(Graph.Url(path), body, ct))
            {
                FolderInfo f = ParseFolder(doc.RootElement);
                f.Path = string.IsNullOrEmpty(parentPath) ? f.DisplayName : parentPath + "/" + f.DisplayName;
                f.Depth = f.Path.Split('/').Length - 1;
                return f;
            }
        }

        // ---------------------------------------------------------------------------------------------
        //  Items
        // ---------------------------------------------------------------------------------------------

        /// <summary>
        /// A listing that failed in a way that means "not now": the folder content is not there yet (being moved to an auxiliary archive), or the mailbox
        /// is moving. Observed: 500 ErrorInternalServerError "Cannot query rows in a table" on the folder of an auxiliary archive whose content was not
        /// moved yet. Documented (EWS): ErrorMailboxMoveInProgress, ErrorMailboxStoreUnavailable, ErrorInternalServerTransientError.
        /// </summary>
        public static bool IsContentNotReady(GraphException ex)
        {
            if (ex == null) return false;
            if (ex.ErrorCode != null && (ex.ErrorCode.Equals("ErrorMailboxMoveInProgress", StringComparison.OrdinalIgnoreCase)
                || ex.ErrorCode.Equals("ErrorMailboxStoreUnavailable", StringComparison.OrdinalIgnoreCase)
                || ex.ErrorCode.Equals("ErrorInternalServerTransientError", StringComparison.OrdinalIgnoreCase))) return true;
            return ex.StatusCode == 500 && ex.Message != null && ex.Message.IndexOf("Cannot query rows in a table", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        /// <summary>The message of the service in a GraphException ("GET url -> 500 Code: message" without the request).</summary>
        public static string ServiceMessageOf(GraphException ex)
        {
            string m = ex.Message ?? "";
            int arrow = m.IndexOf(" -> ", StringComparison.Ordinal);
            if (arrow >= 0)
            {
                int colon = m.IndexOf(": ", arrow, StringComparison.Ordinal);
                if (colon >= 0) return m.Substring(colon + 2);
            }
            return m;
        }

        /// <summary>Times a listing is asked when the content is not available yet (the first try included), then <see cref="ContentNotAvailableException"/>.</summary>
        public int ContentAttempts = 3;
        /// <summary>Waits between those tries, in seconds (multiplied by GraphOptions.DelayScale: 0 in the tests); the last one is reused when the list is shorter.</summary>
        public double[] ContentBackoffSeconds = { 3, 10 };

        bool RetryContent(GraphException ex, int attempt, string mailboxId, string folderId, CancellationToken ct, int attempts = 0)
        {
            if (!IsContentNotReady(ex)) return false;
            if (attempts <= 0) attempts = ContentAttempts;
            if (attempt >= attempts) throw new ContentNotAvailableException(mailboxId, folderId, ex, attempt);
            double seconds = ContentBackoffSeconds == null || ContentBackoffSeconds.Length == 0 ? 0 : ContentBackoffSeconds[Math.Min(attempt - 1, ContentBackoffSeconds.Length - 1)];
            seconds *= Graph.Options.DelayScale;
            var notice = Graph.Notice;
            if (notice != null) { try { notice("content of folder ..." + (folderId.Length > 14 ? folderId.Substring(folderId.Length - 14) : folderId) + " not available (" + ServiceMessageOf(ex) + "), asking again in " + seconds.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) + " s (attempt " + attempt + " of " + attempts + ")"); } catch (Exception) { } }
            if (seconds > 0) ct.WaitHandle.WaitOne(TimeSpan.FromSeconds(seconds));
            ct.ThrowIfCancellationRequested();
            return true;
        }

        /// <summary>One page of an items listing (a URL of .../folders/{id}/items or a nextLink), asked again like <see cref="EnumerateItems"/> when the content is not available yet.</summary>
        public JsonDocument GetItemsPage(string url, string mailboxId, string folderId, CancellationToken ct = default(CancellationToken), Action<string> finalUrl = null)
        {
            for (int attempt = 1; ; attempt++)
            {
                try { return Graph.GetJson(url, ct, finalUrl); }
                catch (GraphException ex) { if (!RetryContent(ex, attempt, mailboxId, folderId, ct)) throw; }
            }
        }

        /// <summary>
        /// Items of a folder, page after page. pageSize 0 lets the service choose. A page the service cannot read yet (folder content being moved to an
        /// auxiliary archive, see <see cref="IsContentNotReady"/>) is asked again <see cref="ContentAttempts"/> times with a back-off, then
        /// <see cref="ContentNotAvailableException"/>; any other error is a GraphException as before.
        /// </summary>
        /// <param name="contentAttempts">How many times a page is asked when the content is not available yet; 0: <see cref="ContentAttempts"/>, 1: once (the export does that for a location that keeps refusing).</param>
        public IEnumerable<ItemInfo> EnumerateItems(string mailboxId, string folderId, int pageSize = 0, CancellationToken ct = default(CancellationToken), int contentAttempts = 0)
        {
            string url = Graph.Url(FolderPath(mailboxId, folderId) + "/items") + (pageSize > 0 ? "?$top=" + pageSize : "");
            string answering = mailboxId;
            foreach (JsonElement e in Graph.EnumerateValues(url, ct, u => { string m = MailboxIdOfUrl(u); if (m != null) answering = m; }, null,
                (ex, attempt) => RetryContent(ex, attempt, mailboxId, folderId, ct, contentAttempts)))
                yield return new ItemInfo
                {
                    MailboxId = answering,
                    Id = Json.Str(e, "id"),
                    Type = Json.Str(e, "type"),
                    ChangeKey = Json.Str(e, "changeKey"),
                    Size = Json.Long(e, "size"),
                    Created = Json.Date(e, "createdDateTime"),
                    Modified = Json.Date(e, "lastModifiedDateTime")
                };
        }

        // ---------------------------------------------------------------------------------------------
        //  Export
        // ---------------------------------------------------------------------------------------------

        /// <summary>
        /// Exports up to 20 items; every item with data becomes a file in <paramref name="directory"/> (streaming, see
        /// ExportResponseReader). Items of an auto-expanded archive folder answer ErrorArchiveFolderMovedPermanently with the
        /// URL to use: they are asked again there, as documented.
        /// </summary>
        public List<ExportEntry> ExportItems(string mailboxId, IList<string> itemIds, string directory, int indexBase = 0,
            CancellationToken ct = default(CancellationToken))
        {
            if (itemIds.Count == 0 || itemIds.Count > MaxExportBatch)
                throw new ArgumentException("exportItems takes 1 to " + MaxExportBatch + " item ids, " + itemIds.Count + " given.");
            var result = new List<ExportEntry>();
            ExportTo(Graph.Url(MailboxPath(mailboxId) + "/exportItems"), itemIds, directory, indexBase, result, 0, ct);
            return result;
        }

        void ExportTo(string url, IList<string> ids, string directory, int indexBase, List<ExportEntry> result, int depth, CancellationToken ct)
        {
            string body = JsonSerializer.Serialize(new { itemIds = ids });
            List<ExportEntry> entries;
            using (HttpResponseMessage resp = Graph.Send("POST", url,
                () => new StringContent(body, Encoding.UTF8, "application/json"),
                completion: HttpCompletionOption.ResponseHeadersRead, ct: ct))
            using (Stream s = Graph.ReadBody(resp, ct))
                entries = ExportResponseReader.Read(s, directory, indexBase + result.Count);

            var moved = new List<ExportEntry>();
            foreach (ExportEntry e in entries)
            {
                e.ExportUrl = url;
                if (depth < 2 && e.ErrorCode == "ErrorArchiveFolderMovedPermanently" && GraphClient.IsGraphUrl(e.ErrorMessage)) moved.Add(e);
                else result.Add(e);
            }
            foreach (var group in moved.GroupBy(e => e.ErrorMessage))
            {
                var observer = Graph.MailboxRedirected;
                string to = MailboxIdOfUrl(group.Key);
                if (observer != null && to != null && !SameMailbox(to, MailboxIdOfUrl(url)))
                {
                    try { observer(MailboxIdOfUrl(url), to, "ErrorArchiveFolderMovedPermanently " + GraphClient.Redact(group.Key)); } catch (Exception) { }
                }
                ExportTo(group.Key, group.Select(e => e.ItemId).ToList(), directory, indexBase, result, depth + 1, ct);
            }
        }

        // ---------------------------------------------------------------------------------------------
        //  Import
        // ---------------------------------------------------------------------------------------------

        public ImportSession CreateImportSession(string mailboxId, CancellationToken ct = default(CancellationToken))
        {
            using (JsonDocument doc = Graph.PostJson(Graph.Url(MailboxPath(mailboxId) + "/createImportSession"), null, ct))
            {
                string importUrl = Json.Str(doc.RootElement, "importUrl");
                if (string.IsNullOrEmpty(importUrl))
                    throw new GraphException(0, "NoImportUrl", "createImportSession answered without importUrl.", null);
                return new ImportSession { MailboxId = mailboxId, ImportUrl = importUrl, Expiration = Json.Date(doc.RootElement, "expirationDateTime") };
            }
        }

        /// <summary>
        /// Creates one item from an exported file ("Mode":"create"). The import URL is pre-authenticated: no Authorization
        /// header is sent. The body is streamed from the file. Not retried after a 504 or a network error: the item may exist.
        /// </summary>
        public ImportResult ImportItem(ImportSession session, string folderId, string filePath, CancellationToken ct = default(CancellationToken))
        {
            return ImportItem(session, folderId, () => new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true),
                new FileInfo(filePath).Length, ct);
        }

        /// <summary>
        /// Imports a FastTransfer stream that is not a file: any IFxStreamSource (an item of the store, FxMessageSource built from a
        /// PST message, ...). The body is written from the stream in pieces, so the memory used does not depend on the size of the item.
        /// </summary>
        public ImportResult ImportItem(ImportSession session, string folderId, IFxStreamSource source, CancellationToken ct = default(CancellationToken))
        {
            return ImportItem(session, folderId, () => source.OpenRead(), source.Length, ct);
        }

        /// <summary>Imports an item of the local store (a C# delegate opens its range of the pack: it runs on HttpClient threads).</summary>
        public ImportResult ImportStoreItem(ImportSession session, string folderId, ExportStore store, long itemKey, CancellationToken ct = default(CancellationToken))
        {
            StoreItem item = store.GetItem(itemKey);
            if (item == null || item.State != StoreItemState.Exported || !item.Length.HasValue)
                throw new InvalidOperationException("Store item " + itemKey + " is not exported.");
            return ImportItem(session, folderId, () => store.OpenItem(itemKey), item.Length.Value, ct);
        }

        /// <summary>
        /// Same, from any readable stream of known length (an item of the local store is a range of a pack file): the stream is
        /// opened by <paramref name="open"/> each time the body is written (and disposed after).
        /// </summary>
        public ImportResult ImportItem(ImportSession session, string folderId, Func<Stream> open, long size, CancellationToken ct = default(CancellationToken))
        {
            var watch = Stopwatch.StartNew();
            using (HttpResponseMessage resp = Graph.Send("POST", session.ImportUrl,
                () => new ImportContent(folderId, open, size), authorize: false, idempotent: false, ct: ct))
            {
                // The service answered 2xx: the item exists. An answer that cannot be read (empty, not JSON) is not a failure and must never make the caller send the item again:
                // an empty answer is an item created without a known id; an unreadable one is reported as Unknown (status 0, not "nothing created"), to be looked up in the folder.
                string text;
                try { text = resp.Content.ReadAsStringAsync().GetAwaiter().GetResult(); }
                catch (Exception ex) when (!(ex is OperationCanceledException))
                { throw new GraphException(0, "ImportAnswerUnreadable", "importItem answered " + (int)resp.StatusCode + " but its answer could not be read (" + ex.Message + "): the item may have been created.", session.ImportUrl); }
                watch.Stop();
                string itemId = null, changeKey = null;
                if (!string.IsNullOrWhiteSpace(text))
                {
                    try
                    {
                        using (JsonDocument doc = JsonDocument.Parse(text))
                        {
                            itemId = Json.Str(doc.RootElement, "itemId");
                            changeKey = Json.Str(doc.RootElement, "changeKey");
                        }
                    }
                    catch (JsonException ex)
                    { throw new GraphException(0, "ImportAnswerUnreadable", "importItem answered " + (int)resp.StatusCode + " but its answer is not readable (" + ex.Message + "): the item may have been created.", session.ImportUrl); }
                }
                return new ImportResult { ItemId = itemId, ChangeKey = changeKey, Bytes = size, Seconds = watch.Elapsed.TotalSeconds };
            }
        }
    }

    /// <summary>
    /// Body of importItem: {"FolderId":"...","Mode":"create","Data":"&lt;base64 of the file&gt;"} written from the file in pieces,
    /// with an exact Content-Length (no chunked encoding).
    /// </summary>
    public sealed class ImportContent : HttpContent, IUploadProgress
    {
        readonly Func<Stream> _open;
        readonly long _size;
        readonly byte[] _prefix;
        static readonly byte[] Suffix = Encoding.ASCII.GetBytes("\"}");
        long _written;
        volatile bool _writeFailed;
        volatile bool _anyWritten;

        /// <summary>Bytes of the body handed to the transport so far (this attempt).</summary>
        public long BytesWritten { get { return Interlocked.Read(ref _written); } }
        public long TotalBytes { get { long n; TryComputeLength(out n); return n; } }
        /// <summary>True when a write to the transport threw: the body did not reach its end (a failed read of the source does not count).</summary>
        public bool WriteFailed { get { return _writeFailed; } }
        /// <summary>True once a piece of the body was handed to the transport (sticky: if the transport asks for the body again, the first attempt counts).</summary>
        public bool AnyByteWritten { get { return _anyWritten; } }

        public ImportContent(string folderId, string path)
            : this(folderId, () => new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, true), new FileInfo(path).Length)
        {
        }

        public ImportContent(string folderId, Func<Stream> open, long size)
        {
            _open = open;
            _size = size;
            _prefix = Encoding.UTF8.GetBytes("{\"FolderId\":" + JsonSerializer.Serialize(folderId) + ",\"Mode\":\"create\",\"Data\":\"");
            Headers.ContentType = new MediaTypeHeaderValue("application/json");
        }

        protected override bool TryComputeLength(out long length)
        {
            length = _prefix.Length + (_size + 2) / 3 * 4 + Suffix.Length;
            return true;
        }

        protected override async Task SerializeToStreamAsync(Stream stream, TransportContext context)
        {
            Interlocked.Exchange(ref _written, 0);
            _writeFailed = false;
            await Write(stream, _prefix, _prefix.Length).ConfigureAwait(false);
            var raw = new byte[3 * 16384];
            var encoded = new byte[4 * 16384];
            using (Stream file = _open())
            {
                int n = await ReadFull(file, raw).ConfigureAwait(false);
                while (n > 0)
                {
                    bool last = n < raw.Length;   // a short block can only be the end of the file
                    int consumed, written;
                    if (Base64.EncodeToUtf8(new ReadOnlySpan<byte>(raw, 0, n), encoded, out consumed, out written, last) != System.Buffers.OperationStatus.Done)
                        throw new InvalidDataException("base64 encoding failed.");
                    await Write(stream, encoded, written).ConfigureAwait(false);
                    if (last) break;
                    n = await ReadFull(file, raw).ConfigureAwait(false);
                }
            }
            await Write(stream, Suffix, Suffix.Length).ConfigureAwait(false);
        }

        /// <summary>Hands a piece of the body to the transport; a failure here is recorded (the service did not get the whole body).</summary>
        async Task Write(Stream stream, byte[] buffer, int count)
        {
            try { await stream.WriteAsync(buffer, 0, count).ConfigureAwait(false); }
            catch (Exception) { _writeFailed = true; throw; }
            Interlocked.Add(ref _written, count);
            _anyWritten = true;
        }

        static async Task<int> ReadFull(Stream s, byte[] buffer)
        {
            int total = 0;
            while (total < buffer.Length)
            {
                int n = await s.ReadAsync(buffer, total, buffer.Length - total).ConfigureAwait(false);
                if (n == 0) break;
                total += n;
            }
            return total;
        }
    }
}
