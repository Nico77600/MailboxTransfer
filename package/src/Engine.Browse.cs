// =============================================================================
//  Mailbox Transfer - engine, part 8: live browse of a mailbox before the export
// -----------------------------------------------------------------------------
//  The user looks at a mailbox like a mailbox (folders with counts and sizes, messages with subject, sender, dates, an
//  excerpt of the body, the attachments), ticks what to export, and the selection becomes an export plan. Nothing here
//  writes to a console: every call takes a CancellationToken and an optional progress callback, and returns objects (the
//  WPF console will drive them).
//
//  Two APIs, one view:
//    Microsoft Graph Mail API (Mail.Read, delegated)   Primary and main Archive. The archive is addressed as a mailbox of its own,
//        /users/MBX:<mailbox guid>@<tenant id>/mailFolders... and /messages (how MailboxMessageReport reads it). Full headers.
//    Mailbox import and export API (what the export uses)   Auxiliary archives (the Mail API does not reach them), the folders the
//        Mail API does not list (Calendar, Contacts, Tasks, Notes, ...), and the fallback when the Mail API refuses a mailbox.
//        It gives per-item size and dates only; the subject and the rest come after the export, from the FastTransfer stream
//        (the previews of the store).
//
//  Correlation of the two APIs (measured on the lab tenant, 2026-10-30):
//    folders  the Mail API folder id and the import-export folder id are the SAME string (REST id) in the primary mailbox and in the
//             main archive: 41 of 41 folders of the Mail API tree of the primary mailbox are in the import-export tree with the same
//             id and the same item count. Recoverable Items (and its sub-folders) are not in the top-level listing of the import-export
//             API, but they are reachable there by their well-known name (recoverableitemsroot) and by id (measured 2026-11-09): the browse view reads their
//             tree and counts from the import-export API (MailboxApi.GetRecoverableItemsTree), the messages from the Mail API, and they are exportable.
//    items    the Mail API message id and the import-export item id are the SAME string too; PR_MESSAGE_SIZE (0x0E08, read with the
//             message) equals the size the import-export listing announces, byte for byte (3 of 3 items of the Inbox).
//    Both facts are checked again when a selection becomes an export plan (ResolvePlan): a folder is looked up in the import-export
//    API by its id (one GET); if the id is unknown there, the folder is found by its path in the import-export tree; a selected item
//    the import-export listing of the folder does not know by id is matched by size and dates (created or last modified), or
//    reported. Nothing is assumed silently: the plan says how each entry was matched (SameId, ByPath, Unresolved).
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;

namespace MailboxTransfer
{
    /// <summary>One mailbox of the user: Primary, Archive (main archive) or AuxArchive.</summary>
    public sealed class BrowseLocation
    {
        public string Kind;
        public string MailboxId;
        /// <summary>How the location was learned (Graph, OperatorFile, Parameter, a redirect...), informational.</summary>
        public string Source;
        /// <summary>The Mail API reaches Primary and the main archive, not the auxiliary archives.</summary>
        public bool HasMailApi { get { return !string.Equals(Kind, "AuxArchive", StringComparison.OrdinalIgnoreCase); } }
        public override string ToString() { return Kind + " " + MailboxId; }
    }

    public sealed class BrowseProgress
    {
        public string Stage;
        public int Done;
        /// <summary>-1 when not known.</summary>
        public int Total = -1;
        public string Message;
        public override string ToString() { return Stage + " " + Done + (Total >= 0 ? "/" + Total : "") + (Message == null ? "" : " " + Message); }
    }

    public sealed class BrowseFolder
    {
        /// <summary>Id in the API that listed it. For the Mail API folders of the primary mailbox and of the main archive it is also the import-export id.</summary>
        public string Id;
        public string ParentId;
        public string DisplayName;
        public string WellKnownName;
        /// <summary>Slash separated names from the top of the tree.</summary>
        public string Path;
        public int Depth;
        public int ChildFolderCount;
        public int TotalItemCount;
        /// <summary>-1 when the API does not tell.</summary>
        public int UnreadItemCount = -1;
        /// <summary>Size of the folder in bytes (Mail API sizeInBytes), -1 when the API does not tell (import-export API).</summary>
        public long SizeBytes = -1;
        public bool IsHidden;
        /// <summary>Folder class (IPF.Note, IPF.Appointment...) as the import-export API gives it; null for a Mail API folder.</summary>
        public string FolderClass;
        /// <summary>MailApi or ImportExportApi: where the folder and its items are read from.</summary>
        public string Source;
        /// <summary>A folder of Recoverable Items (the root "Recoverable Items" or a folder below it): read by the import-export API, exportable like any other.</summary>
        public bool Recoverable;
        /// <summary>False for a folder the import-export API cannot export (none today: Recoverable Items are exportable); ResolvePlan reports such an entry.</summary>
        public bool InExportTree = true;
        /// <summary>Id of the folder in the import-export API (the id the export takes); null when InExportTree is false. Verified by ResolvePlan.</summary>
        public string ExportFolderId;
        /// <summary>The mailbox the folder was listed in, and its kind (Primary, Archive, AuxArchive); set by the tree and by ListFolders.</summary>
        public string MailboxId;
        public string LocationKind;
        /// <summary>Available, or ContentNotAvailable once a listing of its items found the service unable to read them yet (set by ListItems).</summary>
        public BrowseFolderState State = BrowseFolderState.Available;
        /// <summary>The message of the service when State is ContentNotAvailable.</summary>
        public string StateMessage;
        /// <summary>Set by <see cref="FolderCorrelation.Annotate"/>: this folder exists in the main archive and in an auxiliary archive.</summary>
        public FolderMoveState MoveState = FolderMoveState.None;
        /// <summary>Readable form of MoveState ("being moved to aux archive 1"), null when the folder is not in a move.</summary>
        public string MoveText;
        /// <summary>The identity of the folder across mailboxes, taken from its id (<see cref="ExchangeEntryId"/>); null when the id is not in the Exchange format.</summary>
        public string EntryKey { get { ExchangeEntryId e = ExchangeEntryId.Parse(Id); return e == null ? null : e.FolderKey; } }
        /// <summary>True when the service gave the size of the folder (Mail API); the import-export API does not (SizeBytes -1).</summary>
        public bool SizeKnown { get { return SizeBytes >= 0; } }
        public override string ToString() { return Path + " (" + TotalItemCount + ")"; }
    }

    /// <summary>Whether the content of a folder can be read now. ContentNotAvailable: the service answered "not now" (the folder is being moved to an auxiliary archive).</summary>
    public enum BrowseFolderState { Available, ContentNotAvailable }

    /// <summary>Where a folder stands when it exists in the main archive and in an auxiliary archive (see <see cref="FolderCorrelation"/>).</summary>
    public enum FolderMoveState
    {
        None,
        /// <summary>The items are still in the main archive; the auxiliary archive holds the folder skeleton.</summary>
        BeingMoved,
        /// <summary>The items are in the auxiliary archive; the main archive holds an empty placeholder.</summary>
        Moved,
        /// <summary>Items are visible on both sides (a move in progress).</summary>
        BothSides
    }

    /// <summary>
    /// Totals over folders (or the entries of a selection) that never add up an unknown size: the import-export API does not give the size of a folder
    /// (-1), and -1 summed is not a size. FoldersWithUnknownSize says how many folders are left out of KnownSizeBytes.
    /// </summary>
    public sealed class BrowseTotals
    {
        public int Folders;
        public long Items;
        /// <summary>The sum of the sizes that are known.</summary>
        public long KnownSizeBytes;
        public int FoldersWithUnknownSize;
        /// <summary>No folder has an unknown size.</summary>
        public bool SizeComplete { get { return FoldersWithUnknownSize == 0; } }
        /// <summary>No folder has a known size (so there is no total at all).</summary>
        public bool SizeUnknown { get { return Folders > 0 && FoldersWithUnknownSize == Folders; } }

        public void Add(long items, long sizeBytes)
        {
            Folders++;
            Items += items;
            if (sizeBytes >= 0) KnownSizeBytes += sizeBytes; else FoldersWithUnknownSize++;
        }

        public static BrowseTotals Of(IEnumerable<BrowseFolder> folders)
        {
            var t = new BrowseTotals();
            foreach (BrowseFolder f in folders) t.Add(f.TotalItemCount, f.SizeBytes);
            return t;
        }

        public static BrowseTotals Of(BrowseSelection selection)
        {
            var t = new BrowseTotals();
            foreach (SelectionEntry e in selection.Entries) t.Add(e.Items != null && e.Items.Count > 0 ? e.Items.Count : e.ItemCount, e.Items != null && e.Items.Count > 0 ? e.Items.Where(i => i.Size < 0).Any() ? -1 : e.Items.Sum(i => i.Size) : e.SizeBytes);
            return t;
        }

        public static string FormatBytes(long bytes)
        {
            string[] units = { "B", "KiB", "MiB", "GiB", "TiB" };
            double v = bytes;
            int u = 0;
            while (v >= 1024 && u < units.Length - 1) { v /= 1024; u++; }
            return (u == 0 ? v.ToString("0") : v.ToString("0.0", System.Globalization.CultureInfo.InvariantCulture)) + " " + units[u];
        }

        /// <summary>"40 folder(s), 1234 item(s), 12.3 GiB" and, when some sizes are not known, " (n folders with unknown size, not counted)"; "size unknown" when none is.</summary>
        public override string ToString()
        {
            string size = SizeUnknown ? "size unknown" : FormatBytes(KnownSizeBytes);
            string unknown = FoldersWithUnknownSize > 0 && !SizeUnknown ? " (" + FoldersWithUnknownSize + " folder" + (FoldersWithUnknownSize == 1 ? "" : "s") + " with unknown size, not counted)" : "";
            return Folders + " folder(s), " + Items + " item(s), " + size + unknown;
        }
    }

    public sealed class BrowseTree
    {
        public BrowseLocation Location;
        public List<BrowseFolder> Folders = new List<BrowseFolder>();
        /// <summary>What the view could not give (Mail API refused, import-export folders not merged...), for the user.</summary>
        public List<string> Notes = new List<string>();
        /// <summary>The Mail API gave the tree; false: the import-export API (auxiliary archive, or the Mail API refused).</summary>
        public bool FromMailApi;
        public string FallbackReason;

        /// <summary>Items and size over the whole tree; a folder whose size the service does not give is counted apart (see <see cref="BrowseTotals"/>).</summary>
        public BrowseTotals Totals() { return BrowseTotals.Of(Folders); }

        public BrowseFolder Find(string idOrPath)
        {
            if (string.IsNullOrEmpty(idOrPath)) return null;
            BrowseFolder byId = Folders.FirstOrDefault(f => string.Equals(f.Id, idOrPath, StringComparison.Ordinal));
            if (byId != null) return byId;
            string p = idOrPath.Trim('/');
            var exact = Folders.Where(f => string.Equals(f.Path, p, StringComparison.OrdinalIgnoreCase)).ToList();
            if (exact.Count == 1) return exact[0];
            if (exact.Count > 1) throw new InvalidOperationException("Folder '" + idOrPath + "' is ambiguous, " + exact.Count + " folders have this path.");
            var tail = Folders.Where(f => f.Path.EndsWith("/" + p, StringComparison.OrdinalIgnoreCase)).ToList();
            if (tail.Count == 1) return tail[0];
            if (tail.Count > 1) throw new InvalidOperationException("Folder '" + idOrPath + "' is ambiguous, " + tail.Count + " folders end with it:" + string.Concat(tail.Select(t => Environment.NewLine + "  " + t.Path)));
            return null;
        }
    }

    public enum BrowseSort { ReceivedDescending, ReceivedAscending, SentDescending, SentAscending, SubjectAscending, SubjectDescending }

    public sealed class BrowseItemsQuery
    {
        public int PageSize = 50;
        public BrowseSort Sort = BrowseSort.ReceivedDescending;
        /// <summary>BrowsePage.NextToken of the previous page; null for the first page.</summary>
        public string Token;
        /// <summary>Read To and Cc (their counts). Exchange reads the recipients of each message for it: slower on folders of meeting messages.</summary>
        public bool Recipients = true;
    }

    public sealed class BrowseRecipient
    {
        public string Name;
        public string Address;
        public override string ToString() { return string.IsNullOrEmpty(Name) ? Address : string.IsNullOrEmpty(Address) ? Name : Name + " <" + Address + ">"; }
    }

    /// <summary>A row of the list of a folder. Source MailApi: every field; ImportExportApi: Id, ItemClass, Size, Created, Modified and ChangeKey only.</summary>
    public sealed class BrowseItem
    {
        public string Id;
        public string FolderId;
        public string Source;
        public string Subject;
        public string FromName, FromAddress;
        /// <summary>-1 when not read.</summary>
        public int ToCount = -1, CcCount = -1;
        public DateTimeOffset? Received, Sent, Created, Modified;
        /// <summary>Size in bytes (PR_MESSAGE_SIZE for the Mail API, the announced size for the import-export API), -1 when unknown.</summary>
        public long Size = -1;
        public bool? HasAttachments, IsRead, IsDraft;
        public string Importance, Sensitivity, ItemClass, FlagStatus, InternetMessageId, ConversationId, ChangeKey;
        public string[] Categories = new string[0];
        /// <summary>False for a row of the import-export API: the details come after the export, from the FastTransfer stream.</summary>
        public bool HasDetails { get { return Source == MailboxBrowser.MailApi; } }
    }

    public sealed class BrowsePage
    {
        public List<BrowseItem> Items = new List<BrowseItem>();
        /// <summary>Pass it as BrowseItemsQuery.Token to read the next page; null after the last one.</summary>
        public string NextToken;
        public bool HasMore { get { return NextToken != null; } }
        public string Source;
        public List<string> Notes = new List<string>();
        /// <summary>ContentNotAvailable: the service cannot read the items of this folder yet (after bounded retries); Items is empty and that is NOT the content of the folder.</summary>
        public BrowseFolderState State = BrowseFolderState.Available;
        /// <summary>The message of the service when State is ContentNotAvailable.</summary>
        public string StateMessage;
    }

    public sealed class BrowseAttachment
    {
        public string Id, Name, ContentType;
        public long Size;
        public bool IsInline;
        /// <summary>file, item (an attached message, event or contact) or reference.</summary>
        public string Kind = "file";
        /// <summary>The attached item, when Kind is item (and it could be read).</summary>
        public BrowsePreview Item;
    }

    public sealed class BrowsePreviewOptions
    {
        /// <summary>Characters of the text body kept (0: bodyPreview only, no body request).</summary>
        public int BodyChars = 2000;
        public bool Attachments = true;
        /// <summary>How deep the attached items are read (0: names only).</summary>
        public int NestedDepth = 1;
    }

    public sealed class BrowsePreview
    {
        public BrowseItem Header;
        public List<BrowseRecipient> To = new List<BrowseRecipient>(), Cc = new List<BrowseRecipient>(), Bcc = new List<BrowseRecipient>(), ReplyTo = new List<BrowseRecipient>();
        public string BodyPreview;
        /// <summary>The body as text, cut at BodyChars.</summary>
        public string BodyText;
        public bool BodyTrimmed;
        public List<BrowseAttachment> Attachments = new List<BrowseAttachment>();
        public string Source;
        /// <summary>What is missing and why (auxiliary archive: the details come after the export).</summary>
        public List<string> Notes = new List<string>();
    }

    // ---------------------------------------------------------------------------------------------
    //  Selection (what the user ticked) and export plan (the same, in the ids of the import-export API)
    // ---------------------------------------------------------------------------------------------

    public sealed class SelectionItem
    {
        public string Id;
        public long Size = -1;
        public DateTimeOffset? Created, Modified;
    }

    public sealed class SelectionEntry
    {
        /// <summary>Id of the folder in the browse view.</summary>
        public string FolderId;
        public string FolderPath;
        /// <summary>The id the import-export API is known to use for this folder when it is not FolderId (matched by path while the tree was read); null: FolderId.</summary>
        public string ExportFolderId;
        /// <summary>The folder and every folder below it (ignored when Items is set).</summary>
        public bool IncludeSubfolders;
        /// <summary>Only these items of the folder; null or empty: every item of the folder.</summary>
        public List<SelectionItem> Items;
        /// <summary>False for a folder the import-export tree does not hold (Recoverable Items): ResolvePlan reports it.</summary>
        public bool Exportable = true;
        /// <summary>Items and bytes the browse view announced (informational).</summary>
        public int ItemCount;
        public long SizeBytes = -1;
    }

    /// <summary>What the user ticked in the browse view of one mailbox. Serialised as JSON (a selection file).</summary>
    public sealed class BrowseSelection
    {
        public int Version = 1;
        public string Kind;
        public string MailboxId;
        public List<SelectionEntry> Entries = new List<SelectionEntry>();

        static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions
        {
            IncludeFields = true, WriteIndented = true,
            DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
        };

        public static BrowseSelection For(BrowseLocation location)
        {
            return new BrowseSelection { Kind = location.Kind, MailboxId = location.MailboxId };
        }

        /// <summary>Ticks a folder (replaces an earlier entry of the same folder).</summary>
        public SelectionEntry AddFolder(BrowseFolder folder, bool includeSubfolders = false)
        {
            Remove(folder.Id);
            var e = new SelectionEntry
            {
                FolderId = folder.Id, FolderPath = folder.Path, ExportFolderId = folder.ExportFolderId == folder.Id ? null : folder.ExportFolderId, IncludeSubfolders = includeSubfolders, Exportable = folder.InExportTree,
                ItemCount = folder.TotalItemCount, SizeBytes = folder.SizeBytes
            };
            Entries.Add(e);
            return e;
        }

        /// <summary>Ticks some items of a folder (added to the items already ticked there; the whole folder entry, if any, is replaced).</summary>
        public SelectionEntry AddItems(BrowseFolder folder, IEnumerable<BrowseItem> items)
        {
            SelectionEntry e = Entries.FirstOrDefault(x => string.Equals(x.FolderId, folder.Id, StringComparison.Ordinal));
            if (e == null || e.Items == null || e.Items.Count == 0)
            {
                Remove(folder.Id);
                e = new SelectionEntry { FolderId = folder.Id, FolderPath = folder.Path, ExportFolderId = folder.ExportFolderId == folder.Id ? null : folder.ExportFolderId, Exportable = folder.InExportTree, Items = new List<SelectionItem>(), ItemCount = folder.TotalItemCount, SizeBytes = folder.SizeBytes };
                Entries.Add(e);
            }
            foreach (BrowseItem i in items)
                if (!e.Items.Any(x => x.Id == i.Id)) e.Items.Add(new SelectionItem { Id = i.Id, Size = i.Size, Created = i.Created, Modified = i.Modified });
            return e;
        }

        public bool Remove(string folderId)
        {
            return Entries.RemoveAll(x => string.Equals(x.FolderId, folderId, StringComparison.Ordinal)) > 0;
        }

        public void Clear() { Entries.Clear(); }

        public string ToJson() { return JsonSerializer.Serialize(this, JsonOptions); }

        public static BrowseSelection FromJson(string json)
        {
            BrowseSelection s;
            try { s = JsonSerializer.Deserialize<BrowseSelection>(json, JsonOptions); }
            catch (JsonException ex) { throw new InvalidDataException("The selection is not valid JSON: " + ex.Message, ex); }
            if (s == null) throw new InvalidDataException("The selection is empty.");
            if (s.Version != 1) throw new InvalidDataException("Selection version " + s.Version + " is not supported (this build reads version 1).");
            if (string.IsNullOrEmpty(s.MailboxId) || string.IsNullOrEmpty(s.Kind)) throw new InvalidDataException("The selection names no mailbox (Kind and MailboxId are needed).");
            if (s.Entries == null) s.Entries = new List<SelectionEntry>();
            foreach (SelectionEntry e in s.Entries)
                if (string.IsNullOrEmpty(e.FolderId)) throw new InvalidDataException("A selection entry has no FolderId.");
            return s;
        }

        public static BrowseSelection Load(string path) { return FromJson(File.ReadAllText(path, Encoding.UTF8)); }
        public void Save(string path) { FileNaming.Require(path, true, "Selection file"); File.WriteAllText(path, ToJson() + Environment.NewLine, new UTF8Encoding(false)); }
    }

    /// <summary>How an entry of the selection was found in the import-export API.</summary>
    public static class PlanMatch
    {
        public const string SameId = "SameId";
        public const string ByPath = "ByPath";
        public const string Unresolved = "Unresolved";
        /// <summary>The folder (or its items) is in another mailbox than the selection's: it moved to an auxiliary archive since the selection was made.</summary>
        public const string Moved = "Moved";
    }

    public sealed class ExportPlanFolder
    {
        /// <summary>Folder id of the browse view.</summary>
        public string SourceFolderId;
        /// <summary>Folder id of the import-export API (what the export takes); null when Match is Unresolved.</summary>
        public string ExportFolderId;
        /// <summary>The mailbox ExportFolderId is a folder of; null: the mailbox of the plan. Set (Match Moved) when the folder, or its items, moved to another archive.</summary>
        public string MailboxId;
        public string Path;
        public bool IncludeSubfolders;
        /// <summary>Import-export item ids to export; null: every item of the folder.</summary>
        public List<string> ItemIds;
        /// <summary>SameId, ByPath or Unresolved (see PlanMatch).</summary>
        public string Match;
        /// <summary>How the items were matched: ById, BySizeAndDate (a count in the note), or null for a whole folder.</summary>
        public string Note;
        public bool Resolved { get { return Match != PlanMatch.Unresolved && ExportFolderId != null; } }
        /// <summary>Items this entry exports: the ticked items, or the items of the folder (and of the folders below it with IncludeSubfolders), as the service announces them now; -1 when not counted.</summary>
        public long Items = -1;
        /// <summary>Folders the entry covers (the folder, and the folders below it with IncludeSubfolders).</summary>
        public int FolderCount = 1;
        /// <summary>The size of the folders that give one (Mail API: sizeInBytes; the import-export API gives none), -1 when no folder does. Never a sum with unknown sizes in it: see FoldersUnknownSize.</summary>
        public long Bytes = -1;
        /// <summary>Folders (or ticked items) whose size is not known and is left out of Bytes.</summary>
        public int FoldersUnknownSize = 1;
        /// <summary>No size is missing from Bytes.</summary>
        public bool SizeComplete { get { return Bytes >= 0 && FoldersUnknownSize == 0; } }
        internal SelectionEntry Entry;
        internal FolderInfo Info;
    }

    public sealed class ExportPlan
    {
        public string Kind;
        public string MailboxId;
        public List<ExportPlanFolder> Folders = new List<ExportPlanFolder>();
        /// <summary>One sentence per entry or item that cannot be exported, and why.</summary>
        public List<string> Problems = new List<string>();
        /// <summary>What the resolution noticed that is not a problem (a mailbox it could not read while looking for a moved folder).</summary>
        public List<string> Notes = new List<string>();
        public bool IsComplete { get { return Problems.Count == 0 && Folders.All(f => f.Resolved); } }
        public IEnumerable<ExportPlanFolder> Resolved { get { return Folders.Where(f => f.Resolved); } }

        /// <summary>Items and size of what the plan exports (the entries that resolved), from the counts the service gives now; an unknown size is counted apart (see <see cref="BrowseTotals"/>), never summed.</summary>
        public BrowseTotals Totals()
        {
            var t = new BrowseTotals();
            foreach (ExportPlanFolder f in Resolved)
            {
                t.Folders += f.FolderCount;
                t.Items += Math.Max(0, f.Items);
                if (f.Bytes >= 0) t.KnownSizeBytes += f.Bytes;
                t.FoldersWithUnknownSize += f.FoldersUnknownSize;
            }
            return t;
        }
        /// <summary>The mailboxes the plan exports from: its own and the ones a moved folder or item sends it to. An export job takes one mailbox: see <see cref="ForMailbox"/>.</summary>
        public List<string> Mailboxes
        {
            get
            {
                var l = new List<string>();
                foreach (ExportPlanFolder f in Folders)
                {
                    string m = f.MailboxId ?? MailboxId;
                    if (m != null && !l.Contains(m, StringComparer.OrdinalIgnoreCase)) l.Add(m);
                }
                return l;
            }
        }
        /// <summary>The part of the plan that is exported from one mailbox (the export job of that mailbox takes it as its Plan).</summary>
        public ExportPlan ForMailbox(string mailboxId)
        {
            var p = new ExportPlan { Kind = null, MailboxId = mailboxId, Problems = Problems, Notes = Notes };
            p.Folders.AddRange(Folders.Where(f => MailboxApi.SameMailbox(f.MailboxId ?? MailboxId, mailboxId)));
            return p;
        }
    }

    // ---------------------------------------------------------------------------------------------
    //  One folder in the main archive and in an auxiliary archive (auto-expanding archive, a move in progress)
    // ---------------------------------------------------------------------------------------------

    /// <summary>A folder as one mailbox lists it.</summary>
    public sealed class FolderCopy
    {
        public BrowseLocation Location;
        public BrowseFolder Folder;
        public override string ToString() { return Location.Kind + " " + Folder.Path + " (" + Folder.TotalItemCount + ")"; }
    }

    /// <summary>The same folder in several mailboxes, with where its items are.</summary>
    public sealed class FolderMoveGroup
    {
        /// <summary>The folder entry key (database guid and global counter of the id) or "path:" + the path when an id could not be decoded.</summary>
        public string Key;
        /// <summary>EntryId (the folder part of the Exchange id is the same in every mailbox) or Path (fallback: an id that is not in the Exchange format).</summary>
        public string MatchedBy;
        public string Path;
        public List<FolderCopy> Copies = new List<FolderCopy>();
        public FolderMoveState State;
        public string Text;
        /// <summary>Mailbox id to "aux archive 1 (1a2b3c06)" ..., in the order the auxiliary archives were given.</summary>
        public Dictionary<string, string> AuxLabels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public override string ToString() { return Path + ": " + State + (Text == null ? "" : " - " + Text); }
    }

    /// <summary>
    /// Recognises the same folder in the main archive and in the auxiliary archives. The folder part of an Exchange id is the same in every mailbox that holds
    /// the folder (see <see cref="ExchangeEntryId"/>, verified live 2026-10-31); only when an id is not in that format (test ids, other layouts) the path decides.
    /// Two folders of the same path whose ids both decode to different keys are NOT the same folder (the auxiliary archive has its own "Archive" folder).
    /// Primary mailboxes are left out: nothing moves between the primary mailbox and the archive.
    /// </summary>
    public static class FolderCorrelation
    {
        public static string KeyOf(BrowseFolder f)
        {
            string k = f.EntryKey;
            return k != null ? k : "path:" + (f.Path ?? f.DisplayName ?? "").ToLowerInvariant();
        }

        public static string AuxLabel(int index, string mailboxId)
        {
            string g = mailboxId ?? "";
            int at = g.IndexOf(':'); if (at >= 0) g = g.Substring(at + 1);
            int at2 = g.IndexOf('@'); if (at2 >= 0) g = g.Substring(0, at2);
            return "aux archive " + (index + 1) + (g.Length >= 8 ? " (" + g.Substring(0, 8) + ")" : "");
        }

        /// <summary>The folders that exist in at least two of the mailboxes (Archive, AuxArchive), with their state.</summary>
        public static List<FolderMoveGroup> Correlate(IEnumerable<BrowseTree> trees)
        {
            var list = trees.Where(t => t.Location.Kind == "Archive" || t.Location.Kind == "AuxArchive").ToList();
            var labels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (BrowseTree t in list.Where(t => t.Location.Kind == "AuxArchive"))
                if (!labels.ContainsKey(t.Location.MailboxId)) labels[t.Location.MailboxId] = AuxLabel(labels.Count, t.Location.MailboxId);

            var byKey = new Dictionary<string, FolderMoveGroup>(StringComparer.Ordinal);
            var order = new List<FolderMoveGroup>();
            foreach (BrowseTree t in list)
                foreach (BrowseFolder f in t.Folders)
                {
                    string key = KeyOf(f);
                    FolderMoveGroup g;
                    if (!byKey.TryGetValue(key, out g))
                    {
                        g = new FolderMoveGroup { Key = key, MatchedBy = f.EntryKey != null ? "EntryId" : "Path", Path = f.Path, AuxLabels = labels };
                        byKey[key] = g; order.Add(g);
                    }
                    g.Copies.Add(new FolderCopy { Location = t.Location, Folder = f });
                }

            var result = new List<FolderMoveGroup>();
            foreach (FolderMoveGroup g in order)
            {
                // the path fallback only joins folders of different mailboxes; two folders of one mailbox with one path are not one folder
                if (g.Copies.Select(c => c.Location.MailboxId.ToLowerInvariant()).Distinct().Count() < 2) continue;
                if (g.Copies.GroupBy(c => c.Location.MailboxId.ToLowerInvariant()).Any(x => x.Count() > 1)) continue;
                Classify(g);
                result.Add(g);
            }
            return result;
        }

        static void Classify(FolderMoveGroup g)
        {
            var withItems = g.Copies.Where(c => c.Folder.TotalItemCount > 0).ToList();
            bool archiveHas = withItems.Any(c => c.Location.Kind == "Archive");
            bool auxHas = withItems.Any(c => c.Location.Kind == "AuxArchive");
            FolderCopy archive = g.Copies.FirstOrDefault(c => c.Location.Kind == "Archive");
            var auxes = g.Copies.Where(c => c.Location.Kind == "AuxArchive").ToList();
            string where = string.Join(", ", (auxHas ? withItems.Where(c => c.Location.Kind == "AuxArchive").ToList() : auxes).Select(c => g.AuxLabels[c.Location.MailboxId]).Distinct());
            if (archive == null || (!archiveHas && !auxHas)) { g.State = FolderMoveState.None; return; }
            if (archiveHas && !auxHas)
            {
                g.State = FolderMoveState.BeingMoved;
                g.Text = "being moved to " + where + " (the items are still in the main archive)";
                foreach (FolderCopy c in auxes) { c.Folder.MoveState = g.State; c.Folder.MoveText = "being moved here from the main archive (skeleton: the items are still there)"; }
                archive.Folder.MoveState = g.State; archive.Folder.MoveText = "being moved to " + where;
            }
            else if (!archiveHas)
            {
                g.State = FolderMoveState.Moved;
                g.Text = "moved to " + where + " (the main archive keeps an empty placeholder)";
                foreach (FolderCopy c in auxes) { c.Folder.MoveState = g.State; c.Folder.MoveText = "moved here from the main archive"; }
                archive.Folder.MoveState = g.State; archive.Folder.MoveText = "moved to " + where;
            }
            else
            {
                g.State = FolderMoveState.BothSides;
                g.Text = "items on both sides: main archive " + archive.Folder.TotalItemCount + ", " + string.Join(", ", auxes.Select(c => g.AuxLabels[c.Location.MailboxId] + " " + c.Folder.TotalItemCount)) + " (move in progress)";
                foreach (FolderCopy c in g.Copies) { c.Folder.MoveState = g.State; c.Folder.MoveText = "move in progress, items on both sides"; }
            }
        }

        /// <summary>Correlates the trees and writes the result on their folders (BrowseFolder.MoveState / MoveText). Returns the groups.</summary>
        public static List<FolderMoveGroup> Annotate(IEnumerable<BrowseTree> trees)
        {
            var all = trees.ToList();
            foreach (BrowseTree t in all) foreach (BrowseFolder f in t.Folders) { f.MoveState = FolderMoveState.None; f.MoveText = null; }
            return Correlate(all);
        }
    }

    // ---------------------------------------------------------------------------------------------
    //  The browser
    // ---------------------------------------------------------------------------------------------

    public sealed class MailboxBrowser
    {
        public const string MailApi = "MailApi";
        public const string ImportExportApi = "ImportExportApi";

        const string MessageSelect = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,createdDateTime,lastModifiedDateTime,changeKey,hasAttachments,importance,categories,flag,isRead,isDraft,internetMessageId,conversationId";
        // PR_MESSAGE_SIZE, PR_MESSAGE_CLASS, PR_SENSITIVITY: the Mail API has no property for them
        const string MessageExpand = "singleValueExtendedProperties($filter=id%20eq%20'Integer%200x0E08'%20or%20id%20eq%20'String%200x001A'%20or%20id%20eq%20'Integer%200x0036')";
        static readonly Regex ExtendedId = new Regex(@"^(?<t>\w+)\s+0x(?<n>[0-9A-Fa-f]+)$", RegexOptions.Compiled);

        public GraphClient Graph { get; }
        public MailboxApi Api { get; }
        /// <summary>Merge into the tree the folders the Mail API does not list (Calendar, Contacts, Tasks, Notes, Journal...), read from the import-export API.</summary>
        public bool IncludeNonMailFolders = true;
        /// <summary>Add Recoverable Items and its folders (tree and counts from the import-export API, messages from the Mail API; exportable).</summary>
        public bool IncludeRecoverableItems;
        /// <summary>Which folders of Recoverable Items the tree shows (names, see RecoverableItems.KnownFolders); null: all of them.</summary>
        public string[] RecoverableItemsFolders;

        public MailboxBrowser(MailboxApi api) { Api = api; Graph = api.Graph; }

        /// <summary>The locations of a user, as the browse view takes them: from rows of Get-MbtMailboxLocation (Kind, MailboxId, Source).</summary>
        public static BrowseLocation Location(string kind, string mailboxId, string source = null)
        {
            string k = ExportJob.KindOf(kind).ToString();
            return new BrowseLocation { Kind = k, MailboxId = mailboxId, Source = source };
        }

        static string User(BrowseLocation l) { return "/users/" + l.MailboxId; }
        static string Esc(string id) { return Uri.EscapeDataString(id); }

        static void Report(Action<BrowseProgress> progress, string stage, int done, int total, string message = null)
        {
            if (progress != null) progress(new BrowseProgress { Stage = stage, Done = done, Total = total, Message = message });
        }

        static bool IsRefusal(GraphException ex) { return ex.StatusCode == 403 || ex.StatusCode == 404 || ex.StatusCode == 400; }

        // ---------------------------------------------------------------------------------------------
        //  Folders
        // ---------------------------------------------------------------------------------------------

        static BrowseFolder FromMail(JsonElement e, bool recoverable)
        {
            JsonElement v;
            return new BrowseFolder
            {
                Id = Json.Str(e, "id"), ParentId = Json.Str(e, "parentFolderId"), DisplayName = Json.Str(e, "displayName"), WellKnownName = Json.Str(e, "wellKnownName"),
                ChildFolderCount = (int)Json.Long(e, "childFolderCount"), TotalItemCount = (int)Json.Long(e, "totalItemCount"),
                UnreadItemCount = (int)Json.Long(e, "unreadItemCount", -1), SizeBytes = Json.Long(e, "sizeInBytes", -1),
                IsHidden = Json.TryGet(e, "isHidden", out v) && v.ValueKind == JsonValueKind.True,
                Source = MailApi, Recoverable = recoverable, InExportTree = true
            };
        }

        /// <summary>A folder of Recoverable Items from the import-export API; mailApi: its messages are read through the Mail API (the folder ids are the same in both).</summary>
        static BrowseFolder FromRecoverable(FolderInfo f, bool mailApi)
        {
            return new BrowseFolder
            {
                Id = f.Id, ParentId = f.ParentFolderId, DisplayName = f.DisplayName, WellKnownName = f.WellKnownName, FolderClass = f.Type,
                ChildFolderCount = f.ChildFolderCount, TotalItemCount = f.TotalItemCount, Source = mailApi ? MailApi : ImportExportApi, ExportFolderId = f.Id,
                Recoverable = true, InExportTree = true
            };
        }

        static BrowseFolder FromExport(FolderInfo f)
        {
            return new BrowseFolder
            {
                Id = f.Id, ParentId = f.ParentFolderId, DisplayName = f.DisplayName, WellKnownName = f.WellKnownName, FolderClass = f.Type,
                ChildFolderCount = f.ChildFolderCount, TotalItemCount = f.TotalItemCount, Source = ImportExportApi, ExportFolderId = f.Id
            };
        }

        static void AssignPaths(IList<BrowseFolder> folders)
        {
            var byId = new Dictionary<string, BrowseFolder>(StringComparer.Ordinal);
            foreach (BrowseFolder f in folders) if (f.Id != null) byId[f.Id] = f;
            foreach (BrowseFolder f in folders)
            {
                var names = new List<string>();
                var seen = new HashSet<string>(StringComparer.Ordinal);
                BrowseFolder cursor = f;
                while (cursor != null && seen.Add(cursor.Id))
                {
                    names.Add(cursor.DisplayName ?? "(no name)");
                    BrowseFolder parent = null;
                    if (cursor.ParentId != null) byId.TryGetValue(cursor.ParentId, out parent);
                    cursor = parent;
                }
                names.Reverse();
                f.Depth = names.Count - 1;
                f.Path = string.Join("/", names);
            }
        }

        /// <summary>
        /// Every folder of the mailbox with counts and sizes. Primary and Archive: the Mail API (mailFolders/delta: the whole tree flat, 200 folders a
        /// page, with totalItemCount and sizeInBytes), plus the folders only the import-export API lists (IncludeNonMailFolders) and, if asked,
        /// Recoverable Items. Auxiliary archive, or a mailbox the Mail API refuses (no Mail.Read, 403/404): the import-export tree, with the item counts
        /// and without folder sizes; the reason is in Notes.
        /// </summary>
        public BrowseTree GetFolderTree(BrowseLocation location, CancellationToken ct = default(CancellationToken), Action<BrowseProgress> progress = null)
        {
            return GetFolderTree(location, IncludeRecoverableItems, ct, progress);
        }

        /// <summary>The same, with Recoverable Items added (the folders of RecoverableItemsFolders) or not, whatever IncludeRecoverableItems says.</summary>
        BrowseTree GetFolderTree(BrowseLocation location, bool includeRecoverable, CancellationToken ct, Action<BrowseProgress> progress)
        {
            var tree = new BrowseTree { Location = location };
            if (location.HasMailApi)
            {
                try
                {
                    ReadMailTree(location, tree, ct, progress);
                    tree.FromMailApi = true;
                }
                catch (GraphException ex) when (IsRefusal(ex))
                {
                    tree.Folders.Clear();
                    tree.FallbackReason = "the Mail API refused the mailbox (HTTP " + ex.StatusCode + " " + ex.ErrorCode + "): " + ex.Message;
                    tree.Notes.Add(tree.FallbackReason + ". Showing the import-export listing (item counts, no folder sizes, no message details).");
                }
            }
            else
            {
                tree.FallbackReason = "auxiliary archive: the Mail API does not reach it";
                tree.Notes.Add("Auxiliary archive: folders and items come from the import-export API (counts, sizes and dates of the items). The subject, sender and body are known after the export, from the FastTransfer stream (the previews of the store).");
            }

            if (!tree.FromMailApi)
            {
                Report(progress, "folders", 0, -1, "import-export tree");
                foreach (FolderInfo f in Api.GetFolderTree(location.MailboxId, ct)) tree.Folders.Add(FromExport(f));
            }
            else if (IncludeNonMailFolders)
            {
                try { MergeNonMailFolders(location, tree, ct, progress); }
                catch (GraphException ex) when (ex.StatusCode != 401)
                {
                    tree.Notes.Add("The folders only the import-export API lists (Calendar, Contacts, Tasks...) were not added: " + ex.Message);
                }
            }
            if (includeRecoverable)
            {
                try { ReadRecoverable(location, tree, ct, progress); }
                catch (GraphException ex) when (ex.StatusCode != 401)
                {
                    tree.Notes.Add("Recoverable Items were not read: " + ex.Message);
                }
            }
            AssignPaths(tree.Folders);
            foreach (BrowseFolder f in tree.Folders) { f.MailboxId = location.MailboxId; f.LocationKind = location.Kind; }
            tree.Folders.Sort((a, b) => StringComparer.OrdinalIgnoreCase.Compare(a.Path, b.Path));
            Report(progress, "folders", tree.Folders.Count, tree.Folders.Count, "done");
            return tree;
        }

        void ReadMailTree(BrowseLocation location, BrowseTree tree, CancellationToken ct, Action<BrowseProgress> progress)
        {
            string url = Graph.Url(User(location) + "/mailFolders/delta");
            var headers = new Dictionary<string, string> { { "Prefer", "odata.maxpagesize=200" } };
            int pages = 0;
            foreach (JsonElement e in Graph.EnumerateValues(url, ct, u => { pages++; Report(progress, "folders", tree.Folders.Count, -1, "page " + pages); }, headers))
            {
                JsonElement removed;
                if (Json.TryGet(e, "@removed", out removed)) continue;
                BrowseFolder f = FromMail(e, false);
                if (f.Id != null) { f.ExportFolderId = f.Id; tree.Folders.Add(f); }
            }
        }

        void MergeNonMailFolders(BrowseLocation location, BrowseTree tree, CancellationToken ct, Action<BrowseProgress> progress)
        {
            var known = new HashSet<string>(tree.Folders.Select(f => f.Id), StringComparer.Ordinal);
            Report(progress, "folders", tree.Folders.Count, -1, "folders only the import-export API lists");
            foreach (FolderInfo f in Api.GetRootFolders(location.MailboxId, ct))
                MergeExportFolder(location.MailboxId, f, null, tree, known, ct);
            // The Mail API announces every child in childFolderCount (Calendar, Contacts, Tasks... included) but lists the mail folders only: a folder
            // with fewer children listed than announced gets its children from the import-export API (one request for each such folder).
            var seen = new HashSet<string>(StringComparer.Ordinal);
            while (true)
            {
                var children = new Dictionary<string, int>(StringComparer.Ordinal);
                foreach (BrowseFolder f in tree.Folders)
                    if (f.ParentId != null) { int n; children.TryGetValue(f.ParentId, out n); children[f.ParentId] = n + 1; }
                var incomplete = tree.Folders.Where(f =>
                {
                    int n; children.TryGetValue(f.Id, out n);
                    return f.Source == MailApi && !f.Recoverable && f.ChildFolderCount > n && !seen.Contains(f.Id);
                }).ToList();
                if (incomplete.Count == 0) break;
                int done = 0;
                foreach (BrowseFolder parent in incomplete)
                {
                    ct.ThrowIfCancellationRequested();
                    seen.Add(parent.Id);
                    Report(progress, "folders", done++, incomplete.Count, "non-mail children of " + parent.DisplayName);
                    foreach (FolderInfo c in Api.GetChildFolders(location.MailboxId, parent.ExportFolderId ?? parent.Id, ct))
                        MergeExportFolder(location.MailboxId, c, parent, tree, known, ct);
                }
            }
        }

        /// <summary>
        /// A folder of the import-export API under a Mail API folder (null: at the top). Known by id: nothing to do. Not known by id but a Mail API
        /// folder of the same name is at the same place: it is that folder under another id (the lab tenant never does this, the ids are equal), its
        /// ExportFolderId says which and a note tells. Otherwise it is a folder the Mail API does not list (with its own children).
        /// </summary>
        void MergeExportFolder(string mailboxId, FolderInfo f, BrowseFolder mailParent, BrowseTree tree, HashSet<string> known, CancellationToken ct)
        {
            if (known.Contains(f.Id)) return;
            var byId = new HashSet<string>(tree.Folders.Select(x => x.Id), StringComparer.Ordinal);
            BrowseFolder twin = tree.Folders.FirstOrDefault(x => x.Source == MailApi && !x.Recoverable
                && (mailParent == null ? (x.ParentId == null || !byId.Contains(x.ParentId)) : x.ParentId == mailParent.Id)
                && string.Equals(x.DisplayName, f.DisplayName, StringComparison.OrdinalIgnoreCase));
            if (twin != null)
            {
                twin.ExportFolderId = f.Id;
                known.Add(f.Id);
                string note = "The import-export API uses other folder ids than the Mail API (folder '" + twin.DisplayName + "'): a selection is matched by path when it becomes an export plan.";
                if (!tree.Notes.Contains(note)) tree.Notes.Add(note);
                return;
            }
            AddExportSubtree(mailboxId, f, mailParent, tree.Folders, known, ct);
        }

        void AddExportSubtree(string mailboxId, FolderInfo f, BrowseFolder mailParent, List<BrowseFolder> into, HashSet<string> known, CancellationToken ct)
        {
            if (!known.Add(f.Id)) return;
            BrowseFolder b = FromExport(f);
            if (mailParent != null) b.ParentId = mailParent.Id;
            into.Add(b);
            if (f.ChildFolderCount > 0)
                foreach (FolderInfo c in Api.GetChildFolders(mailboxId, f.Id, ct))
                    AddExportSubtree(mailboxId, c, null, into, known, ct);
        }
        /// <summary>
        /// Recoverable Items: the tree and the item counts come from the import-export API (the Mail API counts messages only, and Recoverable Items hold appointments and others);
        /// when the tree is the Mail API's, each folder is also asked of the Mail API for its size and unread count, and its messages are listed through it (the folder ids are the same
        /// in both APIs, measured 2026-11-09). A folder the Mail API refuses is listed through the import-export API.
        /// </summary>
        void ReadRecoverable(BrowseLocation location, BrowseTree tree, CancellationToken ct, Action<BrowseProgress> progress)
        {
            Report(progress, "folders", tree.Folders.Count, -1, "Recoverable Items");
            RecoverableTree rt = Api.GetRecoverableItemsTree(location.MailboxId, RecoverableItemsFolders, ct);
            bool mailOk = tree.FromMailApi;
            foreach (FolderInfo f in rt.Folders)
            {
                BrowseFolder b = FromRecoverable(f, mailOk);
                if (mailOk)
                {
                    try
                    {
                        // no $select: sizeInBytes is returned by default but is not a property the service accepts in $select (400, measured 2026-11-09)
                        using (JsonDocument doc = Graph.GetJson(Graph.Url(User(location) + "/mailFolders/" + Esc(f.Id)), ct))
                        {
                            BrowseFolder m = FromMail(doc.RootElement, true);
                            b.SizeBytes = m.SizeBytes; b.UnreadItemCount = m.UnreadItemCount;
                        }
                    }
                    catch (GraphException ex) when (IsRefusal(ex)) { b.Source = ImportExportApi; }
                }
                tree.Folders.Add(b);
            }
            foreach (string n in rt.Missing) tree.Notes.Add("Recoverable Items has no folder '" + n + "' in this mailbox.");
        }

        List<BrowseFolder> ListMailChildren(BrowseLocation location, string parentId, bool recoverable, CancellationToken ct)
        {
            var list = new List<BrowseFolder>();
            string url = Graph.Url(User(location) + "/mailFolders/" + Esc(parentId) + "/childFolders?$top=100");
            foreach (JsonElement e in Graph.EnumerateValues(url, ct))
            {
                BrowseFolder f = FromMail(e, recoverable);
                if (f.Id != null) { f.ExportFolderId = f.Id; list.Add(f); }
            }
            return list;
        }

        /// <summary>
        /// Lazy view: the children of one folder (parent null: the first level). Opening a folder costs one request; the whole tree
        /// (GetFolderTree) costs one request per 200 folders. Paths follow the parent's.
        /// </summary>
        public List<BrowseFolder> ListFolders(BrowseLocation location, BrowseFolder parent = null, CancellationToken ct = default(CancellationToken))
        {
            var result = new List<BrowseFolder>();
            if (parent == null)
            {
                bool viaMail = false;
                if (location.HasMailApi)
                {
                    try { result.AddRange(ListMailChildren(location, "msgfolderroot", false, ct)); viaMail = true; }
                    catch (GraphException ex) when (IsRefusal(ex)) { viaMail = false; }
                }
                if (!viaMail || IncludeNonMailFolders)
                {
                    var known = new HashSet<string>(result.Select(f => f.Id), StringComparer.Ordinal);
                    try
                    {
                        foreach (FolderInfo f in Api.GetRootFolders(location.MailboxId, ct))
                        {
                            if (!known.Add(f.Id)) continue;
                            BrowseFolder twin = result.FirstOrDefault(x => x.Source == MailApi && string.Equals(x.DisplayName, f.DisplayName, StringComparison.OrdinalIgnoreCase));
                            if (twin != null) twin.ExportFolderId = f.Id;
                            else result.Add(FromExport(f));
                        }
                    }
                    catch (GraphException) when (viaMail) { }
                }
                if (IncludeRecoverableItems)
                {
                    try
                    {
                        FolderInfo r = Api.GetFolder(location.MailboxId, RecoverableItems.RootWellKnown, ct);
                        r.DisplayName = RecoverableItems.TopName; r.ParentFolderId = null;
                        result.Add(FromRecoverable(r, viaMail));
                    }
                    catch (GraphException ex) when (ex.StatusCode != 401) { }
                }
            }
            else if (parent.Recoverable)
            {
                // the folders of Recoverable Items: from the import-export API; below the root only the folders asked for
                bool isRoot = string.Equals(parent.Path, RecoverableItems.TopName, StringComparison.OrdinalIgnoreCase);
                foreach (FolderInfo c in Api.GetChildFolders(location.MailboxId, parent.ExportFolderId ?? parent.Id, ct))
                {
                    if (isRoot && RecoverableItemsFolders != null && !RecoverableItemsFolders.Any(n => string.Equals(n, c.DisplayName, StringComparison.OrdinalIgnoreCase))) continue;
                    BrowseFolder b = FromRecoverable(c, parent.Source == MailApi);
                    b.ParentId = parent.Id;
                    result.Add(b);
                }
            }
            else if (parent.Source == MailApi)
            {
                result.AddRange(ListMailChildren(location, parent.Id, parent.Recoverable, ct));
                // announced children the Mail API does not list (Calendar, Contacts...): from the import-export API
                if (IncludeNonMailFolders && !parent.Recoverable && parent.ChildFolderCount > result.Count)
                {
                    var known = new HashSet<string>(result.Select(f => f.Id), StringComparer.Ordinal);
                    try
                    {
                        foreach (FolderInfo c in Api.GetChildFolders(location.MailboxId, parent.ExportFolderId ?? parent.Id, ct))
                        {
                            if (!known.Add(c.Id)) continue;
                            BrowseFolder twin = result.FirstOrDefault(x => x.Source == MailApi && string.Equals(x.DisplayName, c.DisplayName, StringComparison.OrdinalIgnoreCase));
                            if (twin != null) twin.ExportFolderId = c.Id;
                            else { BrowseFolder added = FromExport(c); added.ParentId = parent.Id; result.Add(added); }
                        }
                    }
                    catch (GraphException ex) when (ex.StatusCode != 401) { }
                }
            }
            else
            {
                foreach (FolderInfo c in Api.GetChildFolders(location.MailboxId, parent.Id, ct)) result.Add(FromExport(c));
            }
            foreach (BrowseFolder f in result)
            {
                f.Depth = parent == null ? 0 : parent.Depth + 1;
                f.Path = parent == null ? f.DisplayName : parent.Path + "/" + f.DisplayName;
                f.MailboxId = location.MailboxId; f.LocationKind = location.Kind;
                if (f.ParentId == null && parent != null) f.ParentId = parent.Id;
            }
            result.Sort((a, b) => StringComparer.OrdinalIgnoreCase.Compare(a.Path, b.Path));
            return result;
        }

        // ---------------------------------------------------------------------------------------------
        //  Items
        // ---------------------------------------------------------------------------------------------

        static string OrderBy(BrowseSort sort)
        {
            switch (sort)
            {
                case BrowseSort.ReceivedAscending: return "receivedDateTime%20asc";
                case BrowseSort.SentDescending: return "sentDateTime%20desc";
                case BrowseSort.SentAscending: return "sentDateTime%20asc";
                case BrowseSort.SubjectAscending: return "subject%20asc";
                case BrowseSort.SubjectDescending: return "subject%20desc";
                default: return "receivedDateTime%20desc";
            }
        }

        static string SensitivityName(long v)
        {
            switch (v) { case 0: return "normal"; case 1: return "personal"; case 2: return "private"; case 3: return "confidential"; default: return v.ToString(); }
        }

        static BrowseRecipient ReadRecipient(JsonElement r)
        {
            JsonElement ea;
            if (!Json.TryGet(r, "emailAddress", out ea)) return null;
            return new BrowseRecipient { Name = Json.Str(ea, "name"), Address = Json.Str(ea, "address") };
        }

        static List<BrowseRecipient> ReadRecipients(JsonElement e, string name)
        {
            var list = new List<BrowseRecipient>();
            JsonElement arr;
            if (Json.TryGet(e, name, out arr) && arr.ValueKind == JsonValueKind.Array)
                foreach (JsonElement r in arr.EnumerateArray()) { BrowseRecipient x = ReadRecipient(r); if (x != null) list.Add(x); }
            return list;
        }

        static bool? Bool(JsonElement e, string name)
        {
            JsonElement v;
            if (Json.TryGet(e, name, out v)) { if (v.ValueKind == JsonValueKind.True) return true; if (v.ValueKind == JsonValueKind.False) return false; }
            return null;
        }

        static string[] Strings(JsonElement e, string name)
        {
            JsonElement arr;
            if (Json.TryGet(e, name, out arr) && arr.ValueKind == JsonValueKind.Array)
                return arr.EnumerateArray().Where(x => x.ValueKind == JsonValueKind.String).Select(x => x.GetString()).ToArray();
            return new string[0];
        }

        /// <summary>A message of the Mail API (list row or message) as a row; the extended properties give the size, the class and the sensitivity.</summary>
        static BrowseItem ReadMessage(JsonElement e, string folderId)
        {
            var item = new BrowseItem
            {
                Id = Json.Str(e, "id"), FolderId = folderId, Source = MailApi, Subject = Json.Str(e, "subject"),
                Received = Json.Date(e, "receivedDateTime"), Sent = Json.Date(e, "sentDateTime"), Created = Json.Date(e, "createdDateTime"), Modified = Json.Date(e, "lastModifiedDateTime"),
                ChangeKey = Json.Str(e, "changeKey"), HasAttachments = Bool(e, "hasAttachments"), IsRead = Bool(e, "isRead"), IsDraft = Bool(e, "isDraft"),
                Importance = Json.Str(e, "importance"), InternetMessageId = Json.Str(e, "internetMessageId"), ConversationId = Json.Str(e, "conversationId"),
                Categories = Strings(e, "categories")
            };
            JsonElement v;
            if (Json.TryGet(e, "from", out v)) { BrowseRecipient f = ReadRecipient(v); if (f != null) { item.FromName = f.Name; item.FromAddress = f.Address; } }
            if (Json.TryGet(e, "toRecipients", out v) && v.ValueKind == JsonValueKind.Array) item.ToCount = v.GetArrayLength();
            if (Json.TryGet(e, "ccRecipients", out v) && v.ValueKind == JsonValueKind.Array) item.CcCount = v.GetArrayLength();
            if (Json.TryGet(e, "flag", out v)) item.FlagStatus = Json.Str(v, "flagStatus");
            if (Json.TryGet(e, "singleValueExtendedProperties", out v) && v.ValueKind == JsonValueKind.Array)
                foreach (JsonElement p in v.EnumerateArray())
                {
                    Match m = ExtendedId.Match(Json.Str(p, "id") ?? "");
                    if (!m.Success) continue;
                    int tag = Convert.ToInt32(m.Groups["n"].Value, 16);
                    string value = Json.Str(p, "value");
                    long n;
                    if (tag == 0x0E08 && long.TryParse(value, out n)) item.Size = n;
                    else if (tag == 0x001A) item.ItemClass = value;
                    else if (tag == 0x0036 && long.TryParse(value, out n)) item.Sensitivity = SensitivityName(n);
                }
            return item;
        }

        static BrowseItem ReadExportItem(JsonElement e, string folderId)
        {
            return new BrowseItem
            {
                Id = Json.Str(e, "id"), FolderId = folderId, Source = ImportExportApi, ItemClass = Json.Str(e, "type"), ChangeKey = Json.Str(e, "changeKey"),
                Size = Json.Long(e, "size", -1), Created = Json.Date(e, "createdDateTime"), Modified = Json.Date(e, "lastModifiedDateTime"),
                Categories = Strings(e, "categories")
            };
        }

        /// <summary>
        /// One page of the rows of a folder, in the order asked (Mail API). The first page is built from the folder; the next ones from
        /// <see cref="BrowsePage.NextToken"/> (the nextLink of the service, followed as returned when it is https on graph.microsoft.com).
        /// A folder of the import-export API, or one the Mail API refuses (403/404/400), is listed by the import-export API: size and dates only,
        /// in the order of the service (it has no sort), with a note.
        /// </summary>
        public BrowsePage ListItems(BrowseLocation location, BrowseFolder folder, BrowseItemsQuery query = null, CancellationToken ct = default(CancellationToken), Action<BrowseProgress> progress = null)
        {
            if (query == null) query = new BrowseItemsQuery();
            if (query.PageSize < 1 || query.PageSize > 1000) throw new ArgumentOutOfRangeException("query", "PageSize must be between 1 and 1000.");
            BrowsePage result = ListItemsCore(location, folder, query, ct, progress);
            folder.State = result.State;
            folder.StateMessage = result.StateMessage;
            return result;
        }

        BrowsePage ListItemsCore(BrowseLocation location, BrowseFolder folder, BrowseItemsQuery query, CancellationToken ct, Action<BrowseProgress> progress)
        {
            bool tokenFromExport = query.Token != null && query.Token.IndexOf("/admin/exchange/", StringComparison.OrdinalIgnoreCase) >= 0;
            bool useMail = location.HasMailApi && folder.Source == MailApi && !tokenFromExport;
            if (useMail)
            {
                try
                {
                    BrowsePage mail = ListMailItems(location, folder, query, ct, progress);
                    // The Mail API lists the messages of a folder, not its calendar items, contacts or tasks (Deleted Items often holds some): when the
                    // whole folder came in one page and holds fewer messages than the folder counts items, say so; when it holds none, show the items
                    // the import-export API lists.
                    if (query.Token == null && !mail.HasMore && mail.Items.Count < folder.TotalItemCount)
                    {
                        string others = (folder.TotalItemCount - mail.Items.Count) + " of the " + folder.TotalItemCount + " item(s) of this folder are not messages (calendar items, contacts, tasks...)";
                        if (mail.Items.Count == 0)
                        {
                            BrowsePage export = ListExportItems(location, folder, query, ct, progress);
                            export.Notes.Insert(0, others + ": listed by the import-export API (size and dates).");
                            return export;
                        }
                        mail.Notes.Add(others + ": the import-export API lists them (size and dates only).");
                    }
                    return mail;
                }
                catch (GraphException ex) when (MailboxApi.IsContentNotReady(ex) && query.Token == null)
                {
                    // Measured live 2026-10-31: the Mail API answers 500 "Cannot query rows in a table" for a folder of the main archive whose content moved to an
                    // auxiliary archive; the import-export API follows the redirect to the auxiliary archive and lists it (size and dates only).
                    BrowsePage page = ListExportItems(location, folder, query, ct, progress);
                    page.Notes.Insert(0, "The Mail API cannot read this folder now (" + MailboxApi.ServiceMessageOf(ex) + "): its content is being moved to an auxiliary archive. Showing the import-export listing (size and dates only).");
                    return page;
                }
                catch (GraphException ex) when (IsRefusal(ex) && query.Token == null)
                {
                    BrowsePage page = ListExportItems(location, folder, query, ct, progress);
                    page.Notes.Insert(0, "The Mail API refused the folder (HTTP " + ex.StatusCode + " " + ex.ErrorCode + "): " + ex.Message + ". Showing the import-export listing.");
                    return page;
                }
            }
            return ListExportItems(location, folder, query, ct, progress);
        }

        /// <summary>Every row of a folder, page after page (lazy: stop whenever you want).</summary>
        public IEnumerable<BrowseItem> EnumerateItems(BrowseLocation location, BrowseFolder folder, BrowseItemsQuery query = null, CancellationToken ct = default(CancellationToken), Action<BrowseProgress> progress = null)
        {
            var q = new BrowseItemsQuery();
            if (query != null) { q.PageSize = query.PageSize; q.Sort = query.Sort; q.Recipients = query.Recipients; q.Token = query.Token; }
            while (true)
            {
                BrowsePage page = ListItems(location, folder, q, ct, progress);
                foreach (BrowseItem i in page.Items) yield return i;
                if (!page.HasMore) yield break;
                q.Token = page.NextToken;
            }
        }

        BrowsePage ListMailItems(BrowseLocation location, BrowseFolder folder, BrowseItemsQuery q, CancellationToken ct, Action<BrowseProgress> progress)
        {
            string url = q.Token;
            if (url == null)
            {
                string select = q.Recipients ? MessageSelect : MessageSelect.Replace("toRecipients,", "").Replace("ccRecipients,", "");
                url = Graph.Url(User(location) + "/mailFolders/" + Esc(folder.Id) + "/messages?$select=" + select + "&$top=" + q.PageSize + "&$orderby=" + OrderBy(q.Sort) + "&$expand=" + MessageExpand);
            }
            else if (!GraphClient.IsGraphUrl(q.Token)) throw new GraphException(0, "UnexpectedNextLink", "The page token is not a https URL on graph.microsoft.com.", GraphClient.Redact(q.Token));
            Report(progress, "items", 0, folder.TotalItemCount, folder.Path);
            var page = new BrowsePage { Source = MailApi };
            using (JsonDocument doc = Graph.GetJson(url, ct))
            {
                JsonElement value, link;
                if (Json.TryGet(doc.RootElement, "value", out value) && value.ValueKind == JsonValueKind.Array)
                    foreach (JsonElement e in value.EnumerateArray()) page.Items.Add(ReadMessage(e, folder.Id));
                if (Json.TryGet(doc.RootElement, "@odata.nextLink", out link) && link.ValueKind == JsonValueKind.String) page.NextToken = link.GetString();
            }
            Report(progress, "items", page.Items.Count, folder.TotalItemCount, folder.Path);
            return page;
        }

        BrowsePage ListExportItems(BrowseLocation location, BrowseFolder folder, BrowseItemsQuery q, CancellationToken ct, Action<BrowseProgress> progress)
        {
            string url = q.Token;
            string mailbox = location.MailboxId;
            if (url == null)
                url = Graph.Url(MailboxApi.MailboxPath(mailbox) + "/folders/" + Esc(folder.ExportFolderId ?? folder.Id) + "/items?$top=" + q.PageSize);
            else if (!GraphClient.IsGraphUrl(q.Token)) throw new GraphException(0, "UnexpectedNextLink", "The page token is not a https URL on graph.microsoft.com.", GraphClient.Redact(q.Token));
            Report(progress, "items", 0, folder.TotalItemCount, folder.Path);
            var page = new BrowsePage { Source = ImportExportApi };
            string answered = url;
            try
            {
                using (JsonDocument doc = Api.GetItemsPage(url, mailbox, folder.ExportFolderId ?? folder.Id, ct, u => answered = u))
                {
                    JsonElement value, link;
                    if (Json.TryGet(doc.RootElement, "value", out value) && value.ValueKind == JsonValueKind.Array)
                        foreach (JsonElement e in value.EnumerateArray()) page.Items.Add(ReadExportItem(e, folder.Id));
                    if (Json.TryGet(doc.RootElement, "@odata.nextLink", out link) && link.ValueKind == JsonValueKind.String) page.NextToken = link.GetString();
                }
            }
            catch (ContentNotAvailableException ex)
            {
                // Not an exception for the view: the folder is there, its content is not readable yet (it is being moved to an auxiliary archive).
                // The token is kept so that the same page can be asked again later.
                page.State = BrowseFolderState.ContentNotAvailable;
                page.StateMessage = ex.ServiceMessage;
                page.NextToken = q.Token;
                page.Notes.Add("The content of this folder is not available yet: the service answered \"" + ex.ServiceMessage + "\" (HTTP " + ex.StatusCode + " " + ex.ErrorCode + ") " + ex.Attempts + " time(s). " +
                    (location.Kind == "AuxArchive" ? "An auxiliary archive gets the folder skeleton first and the items after: the items are probably still in the main archive. Ask again later." : "Ask again later."));
                Report(progress, "items", 0, folder.TotalItemCount, folder.Path + " (content not available)");
                return page;
            }
            if (q.Token == null)
                page.Notes.Add(location.HasMailApi && folder.Source == MailApi
                    ? "Size and dates only, in the order of the service. The subject and the sender come after the export."
                    : "Size and dates only, in the order of the service (the import-export API has no sort). The subject, the sender and the body come after the export, from the FastTransfer stream (the previews of the store).");
            string holder = MailboxApi.MailboxIdOfUrl(answered);
            if (holder != null && !MailboxApi.SameMailbox(holder, mailbox) && q.Token == null)
                page.Notes.Add("This folder is an expanded folder: its items live in the auxiliary archive " + holder + ".");
            Report(progress, "items", page.Items.Count, folder.TotalItemCount, folder.Path);
            return page;
        }

        // ---------------------------------------------------------------------------------------------
        //  Preview
        // ---------------------------------------------------------------------------------------------

        /// <summary>
        /// Header fields, a text excerpt of the body and the attachments of one item. A row of the import-export API (auxiliary archive, non-mail
        /// folder) has no more to show than the row itself: the preview says so.
        /// </summary>
        public BrowsePreview GetItemPreview(BrowseLocation location, BrowseItem row, BrowsePreviewOptions options = null, CancellationToken ct = default(CancellationToken))
        {
            if (options == null) options = new BrowsePreviewOptions();
            var preview = new BrowsePreview { Header = row, Source = row.Source };
            if (!row.HasDetails || !location.HasMailApi)
            {
                preview.Source = ImportExportApi;
                preview.Notes.Add("Auxiliary archive or non-mail folder: only the size and the dates are known before the export. The subject, the sender, the body and the attachments are read from the FastTransfer stream after the export (the previews of the store).");
                return preview;
            }
            string user = User(location);
            string id = Esc(row.Id);
            string select = "id,subject,from,sender,toRecipients,ccRecipients,bccRecipients,replyTo,receivedDateTime,sentDateTime,createdDateTime,lastModifiedDateTime,changeKey,hasAttachments,importance,categories,flag,isRead,isDraft,internetMessageId,conversationId,bodyPreview" + (options.BodyChars > 0 ? ",body" : "");
            var headers = new Dictionary<string, string>();
            if (options.BodyChars > 0) headers["Prefer"] = "outlook.body-content-type=\"text\"";
            using (JsonDocument doc = Graph.GetJson(Graph.Url(user + "/messages/" + id + "?$select=" + select + "&$expand=" + MessageExpand), ct, null, headers))
            {
                JsonElement root = doc.RootElement, body;
                BrowseItem full = ReadMessage(root, row.FolderId);
                if (full.Size < 0) full.Size = row.Size;
                preview.Header = full;
                preview.To = ReadRecipients(root, "toRecipients");
                preview.Cc = ReadRecipients(root, "ccRecipients");
                preview.Bcc = ReadRecipients(root, "bccRecipients");
                preview.ReplyTo = ReadRecipients(root, "replyTo");
                preview.BodyPreview = Json.Str(root, "bodyPreview");
                if (options.BodyChars > 0 && Json.TryGet(root, "body", out body))
                {
                    string text = (Json.Str(body, "content") ?? "").Replace("\r\n", "\n").Trim();
                    if (text.Length > options.BodyChars) { text = text.Substring(0, options.BodyChars); preview.BodyTrimmed = true; }
                    preview.BodyText = text;
                }
            }
            if (options.Attachments && (preview.Header.HasAttachments ?? true))
                preview.Attachments = ReadAttachments(user, preview.Header.Id, options.NestedDepth, options, ct, preview.Notes);
            return preview;
        }

        List<BrowseAttachment> ReadAttachments(string user, string messageId, int depth, BrowsePreviewOptions options, CancellationToken ct, List<string> notes)
        {
            var list = new List<BrowseAttachment>();
            string url = Graph.Url(user + "/messages/" + Esc(messageId) + "/attachments?$select=id,name,contentType,size,isInline");
            foreach (JsonElement e in Graph.EnumerateValues(url, ct))
            {
                string type = Json.Str(e, "@odata.type") ?? "";
                var a = new BrowseAttachment
                {
                    Id = Json.Str(e, "id"), Name = Json.Str(e, "name"), ContentType = Json.Str(e, "contentType"), Size = Json.Long(e, "size"),
                    IsInline = Bool(e, "isInline") ?? false,
                    Kind = type.EndsWith("itemAttachment", StringComparison.OrdinalIgnoreCase) ? "item" : type.EndsWith("referenceAttachment", StringComparison.OrdinalIgnoreCase) ? "reference" : "file"
                };
                if (a.Kind == "item" && depth > 0)
                {
                    try { a.Item = ReadAttachedItem(user, messageId, a.Id, depth - 1, options, ct, notes); }
                    catch (GraphException ex) when (ex.StatusCode != 401) { notes.Add("Attached item '" + a.Name + "' not read: " + ex.Message); }
                }
                list.Add(a);
            }
            return list;
        }

        BrowsePreview ReadAttachedItem(string user, string messageId, string attachmentId, int depth, BrowsePreviewOptions options, CancellationToken ct, List<string> notes)
        {
            string url = Graph.Url(user + "/messages/" + Esc(messageId) + "/attachments/" + Esc(attachmentId) + "?$expand=microsoft.graph.itemAttachment/item");
            using (JsonDocument doc = Graph.GetJson(url, ct))
            {
                JsonElement item;
                if (!Json.TryGet(doc.RootElement, "item", out item) || item.ValueKind != JsonValueKind.Object) return null;
                BrowseItem header = ReadMessage(item, null);
                header.Id = null;
                string type = Json.Str(item, "@odata.type") ?? "";
                header.ItemClass = type.StartsWith("#microsoft.graph.", StringComparison.Ordinal) ? "microsoft.graph." + type.Substring("#microsoft.graph.".Length) : type;
                var nested = new BrowsePreview { Header = header, Source = MailApi };
                nested.To = ReadRecipients(item, "toRecipients");
                nested.Cc = ReadRecipients(item, "ccRecipients");
                nested.BodyPreview = Json.Str(item, "bodyPreview");
                JsonElement atts;
                if (Json.TryGet(item, "attachments", out atts) && atts.ValueKind == JsonValueKind.Array)
                    foreach (JsonElement a in atts.EnumerateArray())
                        nested.Attachments.Add(new BrowseAttachment { Id = Json.Str(a, "id"), Name = Json.Str(a, "name"), ContentType = Json.Str(a, "contentType"), Size = Json.Long(a, "size"), IsInline = Bool(a, "isInline") ?? false });
                return nested;
            }
        }

        // ---------------------------------------------------------------------------------------------
        //  Selection -> export plan
        // ---------------------------------------------------------------------------------------------

        /// <summary>
        /// Turns what the user ticked into folder ids and item ids of the import-export API (what ExportOptions.Plan takes). Read only. For each
        /// folder: GET it in the import-export API by its id (SameId); if the service does not know the id, find it by its path in the
        /// import-export tree (ByPath); else Unresolved (Recoverable Items, a folder that does not exist). For an entry with items (verify):
        /// the folder is listed in the import-export API; an item whose id is there is taken by id; one that is not is matched by its message entry key
        /// (the same message in another mailbox), then by size and dates (created or last modified within 2 s) when exactly one listed item fits,
        /// otherwise reported in Problems.
        ///
        /// Folders that moved: with <paramref name="alsoSearch"/> (the other Archive / AuxArchive locations of the user), a folder the selection mailbox no longer
        /// knows, or knows as an empty placeholder, is looked for in those mailboxes by its folder entry key (path as the fallback for an id that is not in
        /// the Exchange format). Found with items there, the plan gets an entry of its own for that mailbox (Match Moved, ExportPlanFolder.MailboxId): a selection
        /// made on the main archive keeps resolving after the folder went to an auxiliary archive. Items ticked by id are looked for on both sides.
        /// A folder whose content the service cannot read yet is reported in Problems (not exportable now), never resolved as empty.
        /// </summary>
        public ExportPlan ResolvePlan(BrowseSelection selection, bool verify = true, CancellationToken ct = default(CancellationToken), Action<BrowseProgress> progress = null,
            IList<BrowseLocation> alsoSearch = null)
        {
            var plan = new ExportPlan { Kind = selection.Kind, MailboxId = selection.MailboxId };
            List<FolderInfo> exportTree = null;
            var others = (alsoSearch ?? new List<BrowseLocation>()).Where(l => (l.Kind == "Archive" || l.Kind == "AuxArchive") && !MailboxApi.SameMailbox(l.MailboxId, selection.MailboxId)).ToList();
            var otherTrees = new Dictionary<string, List<FolderInfo>>(StringComparer.OrdinalIgnoreCase);
            Func<BrowseLocation, List<FolderInfo>> treeOf = l =>
            {
                List<FolderInfo> t;
                if (!otherTrees.TryGetValue(l.MailboxId, out t))
                {
                    try { t = Api.GetFolderTree(l.MailboxId, ct); }
                    catch (GraphException ex) when (ex.StatusCode != 401) { t = new List<FolderInfo>(); plan.Notes.Add("The folders of " + l.Kind + " " + l.MailboxId + " could not be read, a moved folder cannot be looked for there: " + ex.Message); }
                    otherTrees[l.MailboxId] = t;
                }
                return t;
            };
            int n = 0;
            foreach (SelectionEntry entry in selection.Entries)
            {
                ct.ThrowIfCancellationRequested();
                Report(progress, "plan", n++, selection.Entries.Count, entry.FolderPath);
                var pf = new ExportPlanFolder { SourceFolderId = entry.FolderId, Path = entry.FolderPath, IncludeSubfolders = entry.IncludeSubfolders && (entry.Items == null || entry.Items.Count == 0), Match = PlanMatch.Unresolved, Entry = entry };
                plan.Folders.Add(pf);
                if (!entry.Exportable)
                {
                    pf.Note = "the folder cannot be exported through the import-export API";
                    plan.Problems.Add("'" + entry.FolderPath + "': " + pf.Note + ".");
                    continue;
                }
                string tryId = entry.ExportFolderId ?? entry.FolderId;
                bool sameId = string.Equals(tryId, entry.FolderId, StringComparison.Ordinal);
                FolderInfo own = null;
                if (!verify)
                {
                    pf.ExportFolderId = tryId;
                    pf.Match = sameId ? PlanMatch.SameId : PlanMatch.ByPath;
                    pf.Note = "not verified";
                }
                else
                {
                    try
                    {
                        FolderInfo f = Api.GetFolder(selection.MailboxId, tryId, ct);
                        own = f;
                        pf.ExportFolderId = f.Id ?? tryId;
                        pf.Match = sameId ? PlanMatch.SameId : PlanMatch.ByPath;
                        if (!sameId) pf.Note = "the import-export API uses another id than the browse view; matched by path when the tree was read";
                        if (f.MailboxId != null)
                        {
                            // measured live 2026-10-31: the request for a folder that moved is redirected (308) to the auxiliary archive, which answers with its own id for the folder
                            pf.MailboxId = f.MailboxId; pf.Match = PlanMatch.Moved;
                            pf.Note = "the service redirected the folder to " + f.MailboxId + " (it moved to an auxiliary archive); exported from there";
                        }
                    }
                    catch (GraphException ex) when (ex.StatusCode == 404 || ex.StatusCode == 400)
                    {
                        if (exportTree == null) exportTree = Api.GetFolderTree(selection.MailboxId, ct);
                        var byPath = exportTree.Where(t => string.Equals(t.Path, entry.FolderPath, StringComparison.OrdinalIgnoreCase)).ToList();
                        if (byPath.Count == 1) { pf.ExportFolderId = byPath[0].Id; pf.Match = PlanMatch.ByPath; own = byPath[0]; pf.Note = "the import-export API does not know the id of the browse view; found by its path"; }
                        else if (byPath.Count == 0 && others.Count > 0)
                        {
                            // gone from this mailbox: moved to another one since the selection was made?
                            var moved = FindMoved(entry, others, treeOf, withItemsOnly: false);
                            if (moved.Count == 0)
                            {
                                pf.Note = "no such folder in the import-export tree, nor in the other archives";
                                plan.Problems.Add("'" + entry.FolderPath + "': " + pf.Note + ".");
                                continue;
                            }
                            // the first copy takes the place of the entry, the others (with items) follow as entries of their own (an entry of ticked items is resolved once, see ResolveItems)
                            bool first = true;
                            foreach (var m in entry.Items != null && entry.Items.Count > 0 ? moved.Take(1).ToList() : moved)
                            {
                                ExportPlanFolder target = first ? pf : new ExportPlanFolder { SourceFolderId = entry.FolderId, Path = entry.FolderPath, IncludeSubfolders = pf.IncludeSubfolders };
                                if (!first) plan.Folders.Add(target);
                                target.MailboxId = m.Item1.MailboxId; target.ExportFolderId = m.Item2.Id; target.Match = PlanMatch.Moved; target.Info = m.Item2; target.Entry = entry;
                                target.Note = "the folder is no longer in this mailbox: it moved to " + m.Item1.Kind + " " + m.Item1.MailboxId + " since the selection was made";
                                first = false;
                            }
                            if (entry.Items != null && entry.Items.Count > 0)
                            {
                                if (verify) ResolveItems(pf.MailboxId, entry, pf, plan, ct, others, treeOf);
                                else { pf.ItemIds = entry.Items.Select(i => i.Id).ToList(); pf.Note += "; items by id, not verified"; }
                            }
                            continue;
                        }
                        else
                        {
                            pf.Note = byPath.Count == 0 ? "no such folder in the import-export tree" : byPath.Count + " folders of the import-export tree have this path";
                            plan.Problems.Add("'" + entry.FolderPath + "': " + pf.Note + ".");
                            continue;
                        }
                    }
                }

                pf.Info = own;
                // the folder is here: an empty placeholder whose items went to another mailbox gets those mailboxes as entries of their own
                if (own != null && own.TotalItemCount == 0 && others.Count > 0 && (entry.Items == null || entry.Items.Count == 0))
                {
                    foreach (var m in FindMoved(entry, others, treeOf, withItemsOnly: true))
                    {
                        plan.Folders.Add(new ExportPlanFolder
                        {
                            SourceFolderId = entry.FolderId, Path = entry.FolderPath, IncludeSubfolders = pf.IncludeSubfolders, MailboxId = m.Item1.MailboxId, ExportFolderId = m.Item2.Id,
                            Match = PlanMatch.Moved, Note = "the items of this folder are in " + m.Item1.Kind + " " + m.Item1.MailboxId + " (the folder is an empty placeholder here)", Info = m.Item2, Entry = entry
                        });
                    }
                }

                if (entry.Items != null && entry.Items.Count > 0)
                {
                    if (!verify)
                    {
                        pf.ItemIds = entry.Items.Select(i => i.Id).ToList();
                        pf.Note = "items by id, not verified";
                    }
                    else ResolveItems(pf.MailboxId ?? selection.MailboxId, entry, pf, plan, ct, others, treeOf);
                }
            }
            Report(progress, "plan", selection.Entries.Count, selection.Entries.Count, "done");
            if (verify) CountPlan(plan, selection, ct);
            return plan;
        }

        /// <summary>
        /// Items and size of each entry of the plan (ExportPlanFolder.Items / Bytes): ticked items are counted; a whole folder is the count the service gives now (import-export API
        /// listing: auxiliary archives and moved folders), with the size the Mail API announced when the selection was made; a folder with its subfolders is added up over the tree of
        /// the mailbox (the Mail API for Primary and Archive: counts and sizes; the import-export API for the others: counts, no size). A tree that cannot be read leaves the counts of the folder itself.
        /// </summary>
        void CountPlan(ExportPlan plan, BrowseSelection selection, CancellationToken ct)
        {
            BrowseTree own = null;
            bool ownTried = false;
            var infoTrees = new Dictionary<string, List<FolderInfo>>(StringComparer.OrdinalIgnoreCase);
            foreach (ExportPlanFolder pf in plan.Folders.Where(f => f.Resolved))
            {
                ct.ThrowIfCancellationRequested();
                SelectionEntry e = pf.Entry;
                if (pf.ItemIds != null)
                {
                    pf.Items = pf.ItemIds.Count; pf.FolderCount = 1;
                    bool known = e != null && e.Items != null && e.Items.Count == pf.ItemIds.Count && e.Items.All(i => i.Size >= 0);
                    pf.Bytes = known ? e.Items.Sum(i => i.Size) : -1; pf.FoldersUnknownSize = known ? 0 : 1;
                    continue;
                }
                bool inOwn = pf.MailboxId == null || MailboxApi.SameMailbox(pf.MailboxId, selection.MailboxId);
                pf.Items = pf.Info != null ? pf.Info.TotalItemCount : e != null ? e.ItemCount : -1;
                pf.Bytes = inOwn && e != null && e.SizeBytes >= 0 ? e.SizeBytes : -1;
                pf.FoldersUnknownSize = pf.Bytes >= 0 ? 0 : 1;
                if (!pf.IncludeSubfolders) continue;
                try
                {
                    if (inOwn)
                    {
                        if (!ownTried)
                        {
                            ownTried = true;
                            // Recoverable Items are in the tree when the plan names one (they are not in the top-level listing of the service)
                            own = GetFolderTree(new BrowseLocation { Kind = selection.Kind, MailboxId = selection.MailboxId }, plan.Folders.Any(x => RecoverableItems.IsPath(x.Path)), ct, null);
                        }
                        BrowseFolder root = own == null ? null : own.Folders.FirstOrDefault(f => f.Id == pf.SourceFolderId || (pf.ExportFolderId != null && f.Id == pf.ExportFolderId))
                            ?? own.Folders.FirstOrDefault(f => string.Equals(f.Path, pf.Path, StringComparison.OrdinalIgnoreCase));
                        if (root == null) { plan.Notes.Add("'" + pf.Path + "': the folders below it could not be counted (the folder is not in the tree)."); continue; }
                        var covered = own.Folders.Where(f => (f.Path == root.Path || f.Path.StartsWith(root.Path + "/", StringComparison.OrdinalIgnoreCase))).ToList();
                        BrowseTotals t = BrowseTotals.Of(covered);
                        pf.Items = t.Items; pf.FolderCount = t.Folders; pf.Bytes = t.FoldersWithUnknownSize == t.Folders ? -1 : t.KnownSizeBytes; pf.FoldersUnknownSize = t.FoldersWithUnknownSize;
                    }
                    else
                    {
                        List<FolderInfo> tree;
                        if (!infoTrees.TryGetValue(pf.MailboxId, out tree)) { tree = Api.GetFolderTree(pf.MailboxId, ct); infoTrees[pf.MailboxId] = tree; }
                        FolderInfo root = tree.FirstOrDefault(f => f.Id == pf.ExportFolderId);
                        if (root == null) { plan.Notes.Add("'" + pf.Path + "': the folders below it could not be counted (the folder is not in the tree of " + pf.MailboxId + ")."); continue; }
                        var covered = tree.Where(f => f.Path == root.Path || (f.Path != null && f.Path.StartsWith(root.Path + "/", StringComparison.OrdinalIgnoreCase))).ToList();
                        pf.Items = covered.Sum(f => (long)f.TotalItemCount); pf.FolderCount = covered.Count; pf.Bytes = -1; pf.FoldersUnknownSize = covered.Count;
                    }
                }
                catch (GraphException ex) when (ex.StatusCode != 401)
                {
                    plan.Notes.Add("'" + pf.Path + "': the folders below it could not be counted (" + ex.Message + "); the count is the folder itself.");
                }
            }
        }

        /// <summary>The same folder (folder entry key, else path) in the other mailboxes: the mailbox and its folder. withItemsOnly keeps the copies that announce items; if none does and not withItemsOnly, the first copy.</summary>
        List<Tuple<BrowseLocation, FolderInfo>> FindMoved(SelectionEntry entry, List<BrowseLocation> others, Func<BrowseLocation, List<FolderInfo>> treeOf, bool withItemsOnly)
        {
            string key = ExchangeEntryId.Parse(entry.FolderId) == null ? null : ExchangeEntryId.FolderKeyOf(entry.FolderId);
            var all = new List<Tuple<BrowseLocation, FolderInfo>>();
            foreach (BrowseLocation l in others)
                foreach (FolderInfo t in treeOf(l))
                {
                    ExchangeEntryId e = ExchangeEntryId.Parse(t.Id);
                    bool same = key != null && e != null ? e.FolderKey == key : key == null && e == null && string.Equals(t.Path, entry.FolderPath, StringComparison.OrdinalIgnoreCase);
                    if (same) all.Add(Tuple.Create(l, t));
                }
            var withItems = all.Where(x => x.Item2.TotalItemCount > 0).ToList();
            if (withItems.Count > 0) return withItems;
            return withItemsOnly || all.Count == 0 ? new List<Tuple<BrowseLocation, FolderInfo>>() : new List<Tuple<BrowseLocation, FolderInfo>> { all[0] };
        }

        void ResolveItems(string mailboxId, SelectionEntry entry, ExportPlanFolder pf, ExportPlan plan, CancellationToken ct, List<BrowseLocation> others, Func<BrowseLocation, List<FolderInfo>> treeOf)
        {
            // every place the items of this folder may be: this mailbox's folder, and the same folder in the other mailboxes
            var places = new List<Tuple<string, string, List<ItemInfo>>>();   // mailbox, folder id, listing
            var unavailable = new List<string>();
            Action<string, string> list = (mbx, folderId) =>
            {
                try { places.Add(Tuple.Create(mbx, folderId, Api.EnumerateItems(mbx, folderId, 0, ct).ToList())); }
                catch (ContentNotAvailableException ex) { unavailable.Add(mbx + ": " + ex.ServiceMessage); }
            };
            list(mailboxId, pf.ExportFolderId);
            // the other mailboxes are looked at only when an item ticked is not in this folder's listing (the common case costs nothing more)
            var here = new HashSet<string>(places.SelectMany(p => p.Item3.Select(i => i.Id)), StringComparer.Ordinal);
            if (entry.Items.Any(i => !here.Contains(i.Id)))
            {
                string key = ExchangeEntryId.Parse(entry.FolderId) == null ? null : ExchangeEntryId.FolderKeyOf(entry.FolderId);
                foreach (BrowseLocation l in others)
                {
                    if (MailboxApi.SameMailbox(l.MailboxId, mailboxId)) continue;
                    foreach (FolderInfo t in treeOf(l))
                    {
                        ExchangeEntryId e = ExchangeEntryId.Parse(t.Id);
                        bool same = key != null && e != null ? e.FolderKey == key : key == null && e == null && string.Equals(t.Path, entry.FolderPath, StringComparison.OrdinalIgnoreCase);
                        if (same && t.TotalItemCount > 0) list(l.MailboxId, t.Id);
                    }
                }
            }
            if (places.Count == 0)
            {
                plan.Problems.Add("'" + entry.FolderPath + "': the items cannot be checked, the content of the folder is not available yet (" + string.Join("; ", unavailable) + ").");
                pf.ItemIds = new List<string>();
                pf.Note = "content not available";
                return;
            }
            var all = places.SelectMany(p => p.Item3.Select(i => new { Mailbox = p.Item1, Folder = p.Item2, Item = i })).ToList();
            var byId = new HashSet<string>(all.Select(x => x.Item.Id), StringComparer.Ordinal);
            var claimed = new HashSet<string>(entry.Items.Where(i => byId.Contains(i.Id)).Select(i => i.Id), StringComparer.Ordinal);
            var chosen = new List<Tuple<string, string, string>>();   // mailbox, folder id, item id
            int byId2 = 0, byKey = 0, bySize = 0, missing = 0;
            foreach (SelectionItem it in entry.Items)
            {
                var hit = all.FirstOrDefault(x => x.Item.Id == it.Id && MailboxApi.SameMailbox(x.Mailbox, mailboxId)) ?? all.FirstOrDefault(x => x.Item.Id == it.Id);
                if (hit != null) { chosen.Add(Tuple.Create(hit.Mailbox, hit.Folder, hit.Item.Id)); byId2++; continue; }
                if (ExchangeEntryId.IsDecodable(it.Id))
                {
                    string k = ExchangeEntryId.MessageKeyOf(it.Id);
                    var sameKey = all.Where(x => !claimed.Contains(x.Item.Id) && ExchangeEntryId.MessageKeyOf(x.Item.Id) == k).ToList();
                    if (sameKey.Count == 1) { chosen.Add(Tuple.Create(sameKey[0].Mailbox, sameKey[0].Folder, sameKey[0].Item.Id)); claimed.Add(sameKey[0].Item.Id); byKey++; continue; }
                }
                var fits = all.Where(l => !claimed.Contains(l.Item.Id) && it.Size >= 0 && l.Item.Size == it.Size && (Near(l.Item.Created, it.Created) || Near(l.Item.Modified, it.Modified))).ToList();
                if (fits.Count == 1) { chosen.Add(Tuple.Create(fits[0].Mailbox, fits[0].Folder, fits[0].Item.Id)); claimed.Add(fits[0].Item.Id); bySize++; }
                else
                {
                    missing++;
                    plan.Problems.Add("'" + entry.FolderPath + "': item " + it.Id + " " + (fits.Count == 0 ? "is not in the import-export listing of the folder (by id, nor by size and date)" : "matches " + fits.Count + " items of the import-export listing by size and date: not guessed") + ".");
                }
            }
            // an item seen on both sides of a move is taken once
            var seen = new HashSet<string>(StringComparer.Ordinal);
            var ids = new List<string>();
            foreach (var group in chosen.GroupBy(c => c.Item1, StringComparer.OrdinalIgnoreCase))
            {
                var ofMailbox = group.Where(c => seen.Add(ExchangeEntryId.MessageKeyOf(c.Item3))).ToList();
                if (MailboxApi.SameMailbox(group.Key, mailboxId)) ids.AddRange(ofMailbox.Select(c => c.Item3));
                else if (ofMailbox.Count > 0)
                {
                    foreach (var folderGroup in ofMailbox.GroupBy(c => c.Item2))
                        plan.Folders.Add(new ExportPlanFolder
                        {
                            SourceFolderId = entry.FolderId, Path = entry.FolderPath, MailboxId = group.Key, ExportFolderId = folderGroup.Key, Match = PlanMatch.Moved,
                            ItemIds = folderGroup.Select(c => c.Item3).ToList(), Note = folderGroup.Count() + " item(s) found in " + group.Key + ", where the folder moved"
                        });
                }
            }
            pf.ItemIds = ids;
            pf.Note = (pf.Note != null ? pf.Note + "; " : "") + byId2 + " item(s) by id" + (byKey > 0 ? ", " + byKey + " by entry key" : "") + (bySize > 0 ? ", " + bySize + " by size and date" : "") + (missing > 0 ? ", " + missing + " not found" : "") +
                (unavailable.Count > 0 ? "; content not available in " + string.Join("; ", unavailable) : "");
        }

        static bool Near(DateTimeOffset? a, DateTimeOffset? b)
        {
            return a.HasValue && b.HasValue && Math.Abs((a.Value - b.Value).TotalSeconds) <= 2;
        }
    }
}
