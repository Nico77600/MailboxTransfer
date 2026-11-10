// =============================================================================
//  Mailbox Transfer - engine, part 5h: PST repair (writes a NEW file; the source is never modified)
// -----------------------------------------------------------------------------
//  PstRepair.Repair(source, target, options, progress, ct) -> PstRepairReport
//  PstRepair.Plan(source, options, ...)                    -> the same report with the actions that would be taken (nothing written)
//
//  Architecture: repair = tolerant read of the source, then everything is written again with our own PstWriter, which produces a clean
//  file (counts, sizes, header counters, allocation maps, B-trees, tables, name-to-id map are all computed from what is written).
//  The levels only differ in how the folders and messages of the source are discovered:
//    Recompute  folders from the hierarchy tables, messages from the contents tables (the structure is trusted), a reference to a
//               node that does not exist is dropped and reported, a damaged part of a message is dropped and reported
//    Rebuild    folders and messages from the nodes of the node B-tree (the parent of each node), tables are not trusted: a
//               message not listed is kept in its folder, an unreachable folder and the messages whose folder is gone are put under
//               "Recovered items" (with the folder they came from when it is known)
//    Salvage    the file is scanned for valid blocks and pages (signature + CRC) and the B-trees are rebuilt from the scan
//               (PstScan); then as Rebuild
//  A level that cannot even start (header, B-tree, root folder unreadable) goes on to the next one (PstRepairOptions.Escalate).
//
//  Memory: a message is read, written and released; what stays is the folder tree, 8 bytes per message of the source for the
//  level Rebuild / Salvage (NID and parent) and the 250 bytes per message the writer keeps for the contents tables. The source is
//  read through the tolerant reader, a big value is verified by reading it through (block by block) before the message is written
//  so that a damaged value never leaves a half-written message in the target.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace MailboxTransfer
{
    public enum PstRepairLevel { Recompute = 1, Rebuild = 2, Salvage = 3 }

    public enum PstItemStatus { Recovered, Partial, Lost }

    public enum PstActionStatus { Planned, Applied, NotApplied, Event }

    public sealed class PstRepairOptions
    {
        public PstRepairLevel Level = PstRepairLevel.Recompute;
        /// <summary>When the source cannot be walked at this level (damaged header, B-tree or root folder), go on with the next level.</summary>
        public bool Escalate = true;
        /// <summary>Replace the target file when it exists.</summary>
        public bool Overwrite = false;
        /// <summary>Run the checker on the source (report.Before) / on the result (report.After).</summary>
        public bool CheckBefore = true, CheckAfter = true;
        /// <summary>The check of the source was already done (the plan of the automatic level): it is used as report.Before instead of reading the file again.</summary>
        public PstCheckReport BeforeReport;
        /// <summary>Options of both checks (default: all verifications).</summary>
        public PstCheckOptions CheckOptions;
        /// <summary>Items kept in report.Items: all of them up to this number, then only the ones not fully recovered (up to this number too). ItemSink sees every item.</summary>
        public int MaxItemRecords = 100000;
        /// <summary>Actions kept in report.Actions (more are only counted).</summary>
        public int MaxActions = 2000;
        /// <summary>Called for every item (recovered, partial or lost) as soon as it is processed; keeps nothing, for very large files.</summary>
        public Action<PstRepairItem> ItemSink;
        /// <summary>Read through every big value before writing the message (never a half-written message); off: faster, a damaged value then fails the message while it is written.</summary>
        public bool VerifyValues = true;
        public string RecoveredFolderName = "Recovered items";
        /// <summary>
        /// Salvage: also look for message property contexts that no node entry refers to (old versions of deleted or modified messages in
        /// free space look the same, so this is off by default); done anyway when the scan finds no message node at all (node B-tree
        /// destroyed or cut off, a truncated file). Such a message has lost its folder, recipients and attachments.
        /// </summary>
        public bool RecoverUnlinked = false;
        /// <summary>
        /// A file attachment whose data tree has unreadable blocks is kept with the blocks that read and zeros in place of the others,
        /// its names suffixed ("name.ext" -> "name (damaged).ext") and the byte ranges that are zeros listed in the item. Off (default):
        /// the data of such an attachment is dropped (the attachment stays with its other properties and 0 bytes) and reported.
        /// </summary>
        public bool KeepDamagedAttachments = false;
    }

    public sealed class PstRepairAction
    {
        public PstActionStatus Status;
        public string Code, Message, Where;
        public PstRepairLevel Level;
        public int Count = 1;
        public uint Nid;
        /// <summary>Codes of the checker findings this action deals with.</summary>
        public readonly List<string> Fixes = new List<string>();

        public override string ToString()
        {
            var sb = new StringBuilder();
            sb.Append(Status.ToString().ToUpperInvariant().PadRight(10)).Append(' ').Append(Code).Append(": ").Append(Message);
            if (Count > 1) sb.Append(" (").Append(Count).Append(')');
            if (Nid != 0) sb.Append(" [NID 0x").Append(Nid.ToString("X")).Append(']');
            if (!string.IsNullOrEmpty(Where)) sb.Append(" [").Append(Where).Append(']');
            if (Fixes.Count > 0) sb.Append(" <- ").Append(string.Join(", ", Fixes));
            return sb.ToString();
        }
    }

    public sealed class PstRepairItem
    {
        public PstItemStatus Status;
        public uint SourceNid, TargetNid;
        public string SourceFolder, TargetFolder, Subject;
        /// <summary>The parts that could not be recovered (Partial) or the reason (Lost).</summary>
        public readonly List<string> Lost = new List<string>();
        public bool Orphan;
        /// <summary>KeepDamagedAttachments: the attachments kept with zero-filled ranges (also listed in Lost).</summary>
        public readonly List<PstDamagedAttachment> Damaged = new List<PstDamagedAttachment>();
        public override string ToString()
        {
            return Status + " 0x" + SourceNid.ToString("X") + " [" + SourceFolder + "] " + Subject + (Lost.Count > 0 ? " - " + string.Join("; ", Lost) : "");
        }
    }

    public sealed class PstRepairReport
    {
        public string SourcePath, TargetPath;
        public PstRepairLevel Requested, Used;
        public bool Escalated, PlanOnly, Cancelled, Success;
        public string Error;
        public PstCheckReport Before, After;
        public readonly List<PstRepairAction> Actions = new List<PstRepairAction>();
        public int ActionsOmitted;
        public readonly List<PstRepairItem> Items = new List<PstRepairItem>();
        public long ItemsRecovered, ItemsPartial, ItemsLost, OrphanItems, DanglingDropped, FoldersWritten, FoldersRecovered, PropertiesLost, StructuralLosses;
        /// <summary>Where the name-to-id map came from (Original, RecoveredByScan, RebuiltFromUsage) and why; its number of entries and of candidates the scan found.</summary>
        public PstNameMapSource NameMapSource;
        public string NameMapDetail;
        public int NameMapEntries, NameMapCandidates;
        /// <summary>Named-property ids (0x8000 and up) used by the messages read, the ones the map did not resolve, and the properties dropped for that.</summary>
        public int NamedIdsUsed, NamedIdsUnresolved;
        public long NamedPropsDropped;
        /// <summary>KeepDamagedAttachments: attachments kept with zero-filled ranges and the total of those bytes.</summary>
        public long AttachmentsKeptDamaged, DamagedBytes;
        public long ItemsTotal { get { return ItemsRecovered + ItemsPartial + ItemsLost; } }
        /// <summary>Plan: the lowest level that deals with every error and warning of the check.</summary>
        public PstRepairLevel Recommended = PstRepairLevel.Recompute;
        public TimeSpan Elapsed;
        public long SourceLength, TargetLength;

        public double RecoveryRate { get { long t = ItemsTotal; return t == 0 ? 1.0 : (ItemsRecovered + ItemsPartial) / (double)t; } }

        public string Summary()
        {
            var sb = new StringBuilder();
            sb.AppendLine((PlanOnly ? "Plan for " : "Repair of ") + SourcePath + (PlanOnly ? "" : " -> " + TargetPath));
            sb.Append("  level ").Append(Requested);
            if (!PlanOnly) sb.Append(Used != Requested ? " (done at " + Used + ")" : "");
            sb.Append(PlanOnly ? ", recommended " + Recommended : "").AppendLine();
            if (Before != null) sb.AppendLine("  before: " + Before.Errors + " error(s), " + Before.Warnings + " warning(s), " + Before.Infos + " information(s)");
            if (!PlanOnly)
            {
                sb.AppendLine("  items: " + ItemsRecovered.ToString("N0") + " recovered, " + ItemsPartial.ToString("N0") + " partial, " + ItemsLost.ToString("N0") + " lost" +
                    (OrphanItems > 0 ? ", " + OrphanItems.ToString("N0") + " orphan(s) put in the recovered folder" : "") + (DanglingDropped > 0 ? ", " + DanglingDropped.ToString("N0") + " dangling reference(s) dropped" : ""));
                if (NameMapSource != PstNameMapSource.None) sb.AppendLine("  name map: " + NameMapSource + (NameMapSource == PstNameMapSource.Original ? "" : " - " + NameMapDetail) + (NamedIdsUnresolved > 0 ? "; " + NamedIdsUnresolved + " named-property id(s) without entry, " + NamedPropsDropped.ToString("N0") + " propert" + (NamedPropsDropped == 1 ? "y" : "ies") + " dropped" : ""));
                if (AttachmentsKeptDamaged > 0) sb.AppendLine("  damaged attachments kept: " + AttachmentsKeptDamaged.ToString("N0") + " (" + DamagedBytes.ToString("N0") + " bytes zero-filled)");
                if (After != null) sb.AppendLine("  after: " + After.Errors + " error(s), " + After.Warnings + " warning(s), " + After.Infos + " information(s)");
                if (Error != null) sb.AppendLine("  error: " + Error);
                if (Cancelled) sb.AppendLine("  cancelled");
            }
            sb.Append("  " + Actions.Count + " action(s) in " + Elapsed.TotalSeconds.ToString("N1") + " s");
            return sb.ToString();
        }
    }

    /// <summary>A progress sink that calls its handler on the thread doing the work (System.Progress posts to the thread pool; a PowerShell script block needs the runspace thread).</summary>
    public sealed class PstProgressSink : IProgress<PstProgress>
    {
        readonly Action<PstProgress> _handler;
        public PstProgressSink(Action<PstProgress> handler) { _handler = handler; }
        public void Report(PstProgress value) { if (_handler != null) _handler(value); }
    }

    /// <summary>The source cannot be walked at the requested level: nothing was written that matters, the next level may do better.</summary>
    public sealed class PstRepairDiscoveryException : Exception
    {
        public PstRepairDiscoveryException(string m, Exception inner) : base(m, inner) { }
    }

    public static class PstRepair
    {
        public static PstRepairReport Repair(string sourcePath, string targetPath, PstRepairOptions options)
        {
            return Repair(sourcePath, targetPath, options, null, CancellationToken.None);
        }

        public static PstRepairReport Repair(string sourcePath, string targetPath, PstRepairOptions options, IProgress<PstProgress> progress, CancellationToken ct)
        {
            options = options ?? new PstRepairOptions();
            var rep = new PstRepairReport { SourcePath = sourcePath, TargetPath = targetPath, Requested = options.Level, Used = options.Level };
            var sw = Stopwatch.StartNew();
            string tmp = null;
            try
            {
                if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath)) throw new FileNotFoundException("The source file does not exist.", sourcePath);
                if (string.IsNullOrEmpty(targetPath)) throw new ArgumentException("No target file.");
                FileNaming.Require(targetPath, true, "Repair target");
                string fs = Path.GetFullPath(sourcePath), ft = Path.GetFullPath(targetPath);
                if (string.Equals(fs, ft, StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("The target is the source: a repair writes a new file and never modifies the source.");
                if (File.Exists(ft) && !options.Overwrite) throw new IOException("The target exists (use Overwrite).");
                rep.SourceLength = new FileInfo(fs).Length;
                string dir = Path.GetDirectoryName(ft);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir)) throw new DirectoryNotFoundException("The folder of the target does not exist: " + dir);
                tmp = ft + ".repair-tmp";
                if (File.Exists(tmp)) File.Delete(tmp);

                if (options.CheckBefore)
                {
                    rep.Before = options.BeforeReport ?? RunCheck(fs, options, progress, ct);
                    if (rep.Before.Cancelled) { rep.Cancelled = true; return rep; }
                }
                PstRepairRun run = null;
                for (int lv = (int)options.Level; lv <= (int)PstRepairLevel.Salvage; lv++)
                {
                    ct.ThrowIfCancellationRequested();
                    run = new PstRepairRun(fs, tmp, options, rep, progress, ct);
                    try
                    {
                        run.Execute((PstRepairLevel)lv);
                        rep.Used = (PstRepairLevel)lv;
                        break;
                    }
                    catch (PstRepairDiscoveryException x)
                    {
                        run.Abandon();
                        if (File.Exists(tmp)) File.Delete(tmp);
                        if (!options.Escalate || lv == (int)PstRepairLevel.Salvage) throw new InvalidOperationException("The file cannot be repaired at level " + (PstRepairLevel)lv + ": " + x.Message, x);
                        rep.Escalated = true;
                        rep.AddAction(new PstRepairAction { Status = PstActionStatus.Event, Code = "ESCALATE", Level = (PstRepairLevel)lv, Message = "level " + (PstRepairLevel)lv + " cannot walk the file (" + x.Message + "); trying " + (PstRepairLevel)(lv + 1) }, options.MaxActions);
                        rep.ResetCounters();
                    }
                }
                if (File.Exists(ft)) File.Delete(ft);
                if (options.Escalate && rep.Used < PstRepairLevel.Salvage && rep.StructuralLosses > 0)
                {
                    // items lost to errors of the B-trees: the scan of the file may find their blocks. Keep the better of the two results.
                    var rep2 = new PstRepairReport { SourcePath = sourcePath, TargetPath = targetPath, Requested = options.Level, Used = PstRepairLevel.Salvage, Before = rep.Before };
                    string tmp2 = ft + ".repair-tmp2";
                    if (File.Exists(tmp2)) File.Delete(tmp2);
                    try
                    {
                        var run2 = new PstRepairRun(fs, tmp2, options, rep2, progress, ct);
                        try { run2.Execute(PstRepairLevel.Salvage); }
                        finally { run2.Abandon(); }
                        long first = rep.ItemsRecovered + rep.ItemsPartial, second = rep2.ItemsRecovered + rep2.ItemsPartial;
                        rep.AddAction(new PstRepairAction { Status = PstActionStatus.Event, Code = "ESCALATE", Level = rep.Used, Message = rep.StructuralLosses + " item(s) lost to errors of the B-trees at level " + rep.Used + "; level Salvage keeps " + second + " item(s) against " + first + (second >= first ? ": its result is used" : ": the first result is kept") }, options.MaxActions);
                        if (second >= first)
                        {
                            File.Delete(tmp); File.Move(tmp2, tmp);
                            var keep = rep.Actions.ToList(); int omitted = rep.ActionsOmitted;
                            rep.TakeFrom(rep2);
                            rep.Actions.InsertRange(0, keep); rep.ActionsOmitted += omitted;
                            rep.Used = PstRepairLevel.Salvage; rep.Escalated = true;
                        }
                        else if (File.Exists(tmp2)) File.Delete(tmp2);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception) { if (File.Exists(tmp2)) File.Delete(tmp2); }
                }
                File.Move(tmp, ft);
                tmp = null;
                rep.TargetLength = new FileInfo(ft).Length;
                PstRepairPlan.AddMapped(rep, rep.Before, rep.Used, false, options.MaxActions);
                if (options.CheckAfter)
                {
                    rep.After = RunCheck(ft, options, progress, ct);
                    rep.Cancelled = rep.After.Cancelled;
                }
                rep.Success = rep.After == null ? true : rep.After.Errors == 0;
            }
            catch (OperationCanceledException) { rep.Cancelled = true; }
            catch (Exception ex) { rep.Error = ex.Message; rep.Success = false; }
            finally
            {
                if (tmp != null) { try { if (File.Exists(tmp)) File.Delete(tmp); } catch (Exception) { } }
                rep.Elapsed = sw.Elapsed;
            }
            return rep;
        }

        /// <summary>What a repair would do: the check of the source and the actions it implies, with the lowest level that deals with all of it. Nothing is written.</summary>
        public static PstRepairReport Plan(string sourcePath, PstRepairOptions options, IProgress<PstProgress> progress, CancellationToken ct)
        {
            options = options ?? new PstRepairOptions();
            var rep = new PstRepairReport { SourcePath = sourcePath, Requested = options.Level, Used = options.Level, PlanOnly = true };
            var sw = Stopwatch.StartNew();
            try
            {
                if (string.IsNullOrEmpty(sourcePath) || !File.Exists(sourcePath)) throw new FileNotFoundException("The source file does not exist.", sourcePath);
                rep.SourceLength = new FileInfo(sourcePath).Length;
                rep.Before = RunCheck(Path.GetFullPath(sourcePath), options, progress, ct);
                rep.Cancelled = rep.Before.Cancelled;
                PstRepairPlan.AddMapped(rep, rep.Before, options.Level, true, options.MaxActions);
                rep.Success = true;
            }
            catch (OperationCanceledException) { rep.Cancelled = true; }
            catch (Exception ex) { rep.Error = ex.Message; }
            rep.Elapsed = sw.Elapsed;
            return rep;
        }

        static PstCheckReport RunCheck(string path, PstRepairOptions o, IProgress<PstProgress> progress, CancellationToken ct)
        {
            var co = o.CheckOptions ?? new PstCheckOptions();
            co.ProgressReport = progress;
            co.Cancellation = ct;
            return PstChecker.Check(path, co);
        }
    }

    /// <summary>The findings of the checker and what deals with them: the lowest level, the action.</summary>
    static class PstRepairPlan
    {
        sealed class Group
        {
            public string Action, Message;
            public PstRepairLevel Level;
            public string[] Codes;
        }

        static Group G(string action, PstRepairLevel level, string message, params string[] codes) { return new Group { Action = action, Level = level, Message = message, Codes = codes }; }

        static readonly Group[] Groups =
        {
            G("REBUILD-HEADER", PstRepairLevel.Recompute, "header counters recomputed (dwUnique, bidNextP / bidNextB, NID high-water marks, free counts, end of file)",
                "HDR-UNIQUE", "HDR-BIDNEXT", "HDR-BIDNEXTP", "HDR-NID-HIGHWATER", "HDR-AMAPLAST", "HDR-AMAPFREE", "HDR-PMAPFREE", "HDR-EOF", "HDR-EOF-DIFF", "HDR-AMAP-INVALID", "HDR-AMAP-VALUE", "HDR-CLIENTVER", "HDR-PLATFORM", "BBT-BID-NEXT", "BBT-BID0"),
            G("REBUILD-MAPS", PstRepairLevel.Recompute, "allocation maps (AMap, PMap, FMap, FPMap), DList and B-trees written again with a fresh layout; blocks no node uses are not copied",
                "AMAP-BAD", "AMAP-BIT-ORDER", "AMAP-MISSING", "AMAP-NONE", "AMAP-UNMARKED", "AMAP-UNUSED", "PMAP-NOT-FULL", "DLIST-BAD", "DLIST-COUNT", "DLIST-ENTRY", "DLIST-FREE", "ALLOC-CROSS", "ALLOC-OVERLAP", "ALLOC-RANGE",
                "BLOCK-ORPHAN", "BLOCK-CREF", "BLOCK-ALIGN", "BID-RESERVED-BIT", "DATATREE-NOT-FULL", "NODE-STRAY", "TC-ROWVER", "TC-SAMPLED", "TREE-SKIP"),
            G("RECOMPUTE-COUNTS", PstRepairLevel.Recompute, "folder counts (content, unread, subfolders, hidden) and the table rows that repeat them recomputed",
                "FOLDER-COUNT", "FOLDER-UNREAD", "FOLDER-SUBFOLDERS", "FOLDER-HIDDEN-COUNT", "HIERARCHY-ROW-VALUE", "CONTENTS-CELL", "FOLDER-PROP"),
            G("RECOMPUTE-SIZES", PstRepairLevel.Recompute, "message and attachment sizes recomputed",
                "MSG-SIZE", "MSG-SIZE-COMPUTED", "ATTACH-SIZE", "PC-SIZE", "STORE-PROP-SIZE", "MSG-HASATTACH"),
            G("DROP-DANGLING", PstRepairLevel.Recompute, "references to nodes that do not exist dropped",
                "CONTENTS-ROW", "HIERARCHY-ROW", "HIERARCHY-STALE"),
            G("REBUILD-TABLES", PstRepairLevel.Rebuild, "hierarchy and contents tables rebuilt from the folder and message nodes (messages not listed kept, folders not listed or with a wrong parent re-attached, folder loops cut)",
                "MSG-NOT-LISTED", "FOLDER-NOT-LISTED", "CONTENTS-PARENT", "CONTENTS-DUP", "HIERARCHY-DUP", "HIERARCHY-PARENT", "FOLDER-PARENT", "FOLDER-PARENT-MISMATCH", "FOLDER-LOOP", "FOLDER-TABLE", "ROOT-PARENT", "TC-ROWID-MISMATCH", "TC-ROWID", "TC-ROWID-DUP"),
            G("REBUILD-NAMEMAP", PstRepairLevel.Rebuild, "named-property map rebuilt from the properties the messages use; properties whose id has no entry dropped",
                "NAMEMAP-BUCKET-MISSING", "NAMEMAP-BUCKETS", "NAMEMAP-ENTRIES", "NAMEMAP-GUIDREF", "NAMEMAP-GUIDS", "NAMEMAP-HASH", "NAMEMAP-INDEX", "NAMEMAP-STRING", "NAMEDPROP-UNMAPPED", "MAP-BAD"),
            G("DROP-DAMAGED", PstRepairLevel.Recompute, "blocks that fail their checks (CRC, signature, trailer) or are missing cannot be recovered: the values, attachments and messages that use them are dropped (a message is partial or lost) and listed per item",
                "BLOCK-CRC", "BLOCK-SIG", "BLOCK-BID", "BLOCK-CB", "BLOCK-READ", "BLOCK-RANGE", "BLOCK-SIZE", "BID-DANGLING", "XBLOCK-CHILD", "XBLOCK-EMPTY", "XBLOCK-READ", "XBLOCK-SHAPE", "XBLOCK-TOTAL", "XBLOCK-TYPE",
                "SUBNODE-CYCLE", "SUBNODE-EMPTY", "SUBNODE-INTERNAL", "SUBNODE-LEVEL", "SUBNODE-ORDER", "SUBNODE-SHAPE", "SUBNODE-TYPE"),
            G("SALVAGE-INDEX", PstRepairLevel.Salvage, "block and node B-trees rebuilt from the pages found by scanning the file; header taken from what is valid",
                "HDR-MAGIC", "HDR-MAGIC-CLIENT", "HDR-CRC-FULL", "HDR-CRC-PARTIAL", "HDR-VERSION", "HDR-SHORT", "HDR-SENTINEL", "HDR-CRYPT", "HDR-ANSI", "FILE-OPEN", "CHECK-ABORT", "NBT-NID", "BBT-ORDER"),
            G("REWRITE-STORE-OBJECTS", PstRepairLevel.Recompute, "store objects (templates, receive-folder table, outgoing queue, message map, search objects) written again as the writer makes them",
                "TEMPLATE-COLUMN", "TEMPLATE-ROWS", "MSG-OUTGOING-QUEUE", "MSG-TEMPLATE-MISSING", "MSG-NODE-EC1", "HMP-MISSING", "HMP-ENTRY", "RECEIVE-TABLE-ROW", "MSG-RECEIVE-TABLE"),
            G("RECOVER-PARTIAL", PstRepairLevel.Recompute, "damaged messages, attachments and tables: the parts that decode are kept, the others are dropped and listed per item",
                "LTP-BROKEN", "TC-BROKEN", "TC-CBDATA", "TC-COLUMN-ORDER", "TC-IBIT", "MSG-BROKEN", "MSG-OPEN", "MSG-SKIP", "MSG-DEPTH", "ATTACH-BROKEN", "ATTACH-DATA", "ATTACH-EMBEDDED", "ATTACH-EMBEDDED-NID", "ATTACH-MISSING", "ATTACH-PROP",
                "MSG-ATTACH-MISMATCH", "MSG-ATTACH-TABLE", "MSG-RECIPIENTS", "MSG-MANDATORY", "MSG-PROP", "PC-OBJECT", "PC-STRING", "PC-TYPE", "PC-VALUE", "STORE-PROP", "STORE-ENTRYID", "STORE-ENTRYID-UID", "STORE-PASSWORD", "CONTENTS-DUP"),
        };

        static readonly Dictionary<string, Group> ByCode = BuildIndex();
        static Group IndexGroup { get { return Groups.First(x => x.Action == "SALVAGE-INDEX"); } }

        static Dictionary<string, Group> BuildIndex()
        {
            var d = new Dictionary<string, Group>(StringComparer.Ordinal);
            foreach (Group g in Groups) foreach (string c in g.Codes) if (!d.ContainsKey(c)) d[c] = g;
            return d;
        }

        /// <summary>Lowest level that deals with a finding (Salvage for an unknown error, Recompute for an unknown information or warning).</summary>
        public static PstRepairLevel LevelOf(PstFinding f)
        {
            Group g;
            if (ByCode.TryGetValue(f.Code, out g)) return g.Level;
            if (f.Code.StartsWith("NBT-") || f.Code.StartsWith("BBT-")) return PstRepairLevel.Salvage;
            return f.Severity == PstSeverity.Error ? PstRepairLevel.Salvage : PstRepairLevel.Recompute;
        }

        public static void AddMapped(PstRepairReport rep, PstCheckReport before, PstRepairLevel level, bool plan, int maxActions)
        {
            if (before == null) return;
            var counts = new Dictionary<Group, int>();
            var first = new Dictionary<Group, PstFinding>();
            var codes = new Dictionary<Group, SortedSet<string>>();
            var unknown = new SortedSet<string>(StringComparer.Ordinal);
            foreach (KeyValuePair<string, int> kv in before.CountByCode)
            {
                Group g;
                if (!ByCode.TryGetValue(kv.Key, out g) && !(kv.Key.StartsWith("NBT-") || kv.Key.StartsWith("BBT-")) ) { unknown.Add(kv.Key); continue; }
                if (g == null) g = IndexGroup;
                int n; counts.TryGetValue(g, out n); counts[g] = n + kv.Value;
                SortedSet<string> s; if (!codes.TryGetValue(g, out s)) codes[g] = s = new SortedSet<string>(StringComparer.Ordinal);
                s.Add(kv.Key);
            }
            foreach (PstFinding f in before.Findings)
            {
                Group g;
                if (!ByCode.TryGetValue(f.Code, out g) && (f.Code.StartsWith("NBT-") || f.Code.StartsWith("BBT-"))) g = IndexGroup;
                if (g != null && !first.ContainsKey(g)) first[g] = f;
                if (f.Severity != PstSeverity.Info)
                {
                    PstRepairLevel l = LevelOf(f);
                    if (l > rep.Recommended) rep.Recommended = l;
                }
            }
            foreach (Group g in Groups)
            {
                int n;
                if (!counts.TryGetValue(g, out n)) continue;
                var a = new PstRepairAction { Code = g.Action, Message = g.Message, Level = g.Level, Count = n };
                a.Status = plan ? (g.Level <= level ? PstActionStatus.Planned : PstActionStatus.NotApplied) : (g.Level <= level ? PstActionStatus.Applied : PstActionStatus.NotApplied);
                if (a.Status == PstActionStatus.NotApplied) a.Message += " - needs level " + g.Level;
                a.Fixes.AddRange(codes[g]);
                PstFinding ff;
                if (first.TryGetValue(g, out ff)) { a.Where = ff.Where; a.Nid = ff.Nid; }
                rep.AddAction(a, maxActions);
            }
            if (unknown.Count > 0)
            {
                var a = new PstRepairAction { Code = "REVIEW", Level = PstRepairLevel.Salvage, Message = "findings with no specific repair (the file is rewritten from what decodes)", Status = plan ? PstActionStatus.Planned : PstActionStatus.Applied };
                a.Fixes.AddRange(unknown);
                a.Count = unknown.Sum(c => before.CountByCode[c]);
                rep.AddAction(a, maxActions);
            }
        }
    }

    static class PstRepairReportExt
    {
        public static void AddAction(this PstRepairReport r, PstRepairAction a, int max)
        {
            if (r.Actions.Count < max) r.Actions.Add(a); else r.ActionsOmitted++;
        }

        public static void TakeFrom(this PstRepairReport r, PstRepairReport o)
        {
            r.Actions.Clear(); r.Actions.AddRange(o.Actions); r.ActionsOmitted = o.ActionsOmitted;
            r.Items.Clear(); r.Items.AddRange(o.Items);
            r.ItemsRecovered = o.ItemsRecovered; r.ItemsPartial = o.ItemsPartial; r.ItemsLost = o.ItemsLost; r.OrphanItems = o.OrphanItems; r.DanglingDropped = o.DanglingDropped;
            r.FoldersWritten = o.FoldersWritten; r.FoldersRecovered = o.FoldersRecovered; r.PropertiesLost = o.PropertiesLost; r.StructuralLosses = o.StructuralLosses;
            r.NameMapSource = o.NameMapSource; r.NameMapDetail = o.NameMapDetail; r.NameMapEntries = o.NameMapEntries; r.NameMapCandidates = o.NameMapCandidates;
            r.NamedIdsUsed = o.NamedIdsUsed; r.NamedIdsUnresolved = o.NamedIdsUnresolved; r.NamedPropsDropped = o.NamedPropsDropped;
            r.AttachmentsKeptDamaged = o.AttachmentsKeptDamaged; r.DamagedBytes = o.DamagedBytes;
        }

        public static void ResetCounters(this PstRepairReport r)
        {
            r.ItemsRecovered = r.ItemsPartial = r.ItemsLost = r.OrphanItems = r.DanglingDropped = r.FoldersWritten = r.FoldersRecovered = r.PropertiesLost = r.StructuralLosses = 0;
            r.NameMapSource = PstNameMapSource.None; r.NameMapDetail = null; r.NameMapEntries = r.NameMapCandidates = r.NamedIdsUsed = r.NamedIdsUnresolved = 0;
            r.NamedPropsDropped = r.AttachmentsKeptDamaged = r.DamagedBytes = 0;
            r.Items.Clear();
        }
    }

    /// <summary>A folder of the source as the repair sees it.</summary>
    sealed class SrcFolder
    {
        public uint Nid, Parent;
        public string Name = "", Class;
        public int StoredContent = -1, StoredUnread = -1, StoredAssoc = -1, StoredSub = -1;
        public bool Synthetic, InfoLost, Recovered;
        public readonly List<SrcFolder> Children = new List<SrcFolder>();
        public string Path = "";
    }

    sealed partial class PstRepairRun
    {
        readonly string _src, _tmp;
        readonly PstRepairOptions _o;
        readonly PstRepairReport _rep;
        readonly IProgress<PstProgress> _progress;
        readonly CancellationToken _ct;
        PstFile _file;
        PstReader _reader;
        PstWriter _writer;
        PstRepairLevel _level;
        PstFolder _recovered;
        readonly Dictionary<uint, PstFolder> _lostFolders = new Dictionary<uint, PstFolder>();
        uint _ipmNid, _deletedNid, _searchRootNid;
        string _storeName, _ipmName;
        long _itemsDone, _itemsTotal, _problemsKept;
        readonly Stopwatch _clock = Stopwatch.StartNew();
        long _lastReport;

        public PstRepairRun(string src, string tmp, PstRepairOptions o, PstRepairReport rep, IProgress<PstProgress> progress, CancellationToken ct)
        {
            _src = src; _tmp = tmp; _o = o; _rep = rep; _progress = progress; _ct = ct;
            if (rep.Before != null && rep.Before.Stats.ContainsKey("messages")) _itemsTotal = rep.Before.Stats["messages"];
        }

        public void Abandon()
        {
            if (_writer != null) { try { _writer.Dispose(); } catch (Exception) { } _writer = null; }
            if (_reader != null) { try { _reader.Dispose(); } catch (Exception) { } _reader = null; }
            else if (_file != null) { try { _file.Dispose(); } catch (Exception) { } }
            _file = null;
        }

        void Event(string code, string message, string where, uint nid)
        {
            _rep.AddAction(new PstRepairAction { Status = PstActionStatus.Event, Code = code, Level = _level, Message = message, Where = where, Nid = nid }, _o.MaxActions);
        }

        // ------------------------------------------------------------------ entry point
        public void Execute(PstRepairLevel level)
        {
            _level = level;
            try
            {
                Discover();
                Write();
            }
            finally { Abandon(); }
        }

        PstProp Try(PstPcReader pc, ushort id)
        {
            try { return pc.Get(id); } catch (Exception) { return null; }
        }

        static uint EntryIdNid(PstProp p) { return p == null || p.Data == null || p.Data.Length < 24 ? 0 : PstFormat.G32(p.Data, 20); }

        // ------------------------------------------------------------------ opening the source
        void OpenSource()
        {
            if (_level == PstRepairLevel.Salvage)
            {
                _file = PstFile.OpenSalvage(_src);
                var scan = PstScan.Scan(_file, _ct, _progress);
                _file.Index = scan;
                _file.CryptMethod = scan.DetectCrypt(_file, _file.CryptMethod);
                Event("SALVAGE-SCAN", scan.Describe(), null, 0);
                RecoverUnlinkedMessages(scan);
                _rep.StructuralLosses = 0;
            }
            else
            {
                try { _file = PstFile.Open(_src); }
                catch (Exception x) when (!(x is OperationCanceledException)) { throw new PstRepairDiscoveryException("header: " + x.Message, x); }
            }
            _reader = PstReader.Wrap(_file);
            _reader.NameMapErrors = new List<string>();
            ResolveNameMap(_level != PstRepairLevel.Recompute);
        }

        HashSet<uint> _unlinked;

        void RecoverUnlinkedMessages(PstScan scan)
        {
            bool anyMessage = scan.Nodes().Any(n => PstFormat.NidType(n.Nid) == PstFormat.NtNormalMessage);
            if (!_o.RecoverUnlinked && anyMessage) return;
            var found = new List<ulong>();
            foreach (ulong root in scan.UnreferencedRoots())
            {
                _ct.ThrowIfCancellationRequested();
                try
                {
                    var pc = new PstPcReader(new PstNodeView(_file, 0, root, 0));
                    PstProp cls = pc.Get(0x001A);
                    string c = PstConvert.StringOf(cls);
                    if (!string.IsNullOrEmpty(c) && pc.Find(0x0E07) != null) found.Add(root);
                }
                catch (Exception x) when (!(x is OperationCanceledException)) { }
            }
            if (found.Count == 0) return;
            _unlinked = new HashSet<uint>(scan.AddUnlinked(found));
            Event("UNLINKED-MESSAGES", found.Count + " message property context(s) that no node entry refers to " + (anyMessage ? "(requested)" : "(no message node found in the file)") + ": recovered without folder, recipients and attachments", null, 0);
        }

        void ReadStore()
        {
            try
            {
                var pc = new PstPcReader(_file.OpenNode(PstFormat.NidMessageStore));
                _storeName = PstConvert.StringOf(Try(pc, 0x3001));
                _ipmNid = EntryIdNid(Try(pc, 0x35E0));
                _deletedNid = EntryIdNid(Try(pc, 0x35E3));
                _searchRootNid = EntryIdNid(Try(pc, 0x35E7));
            }
            catch (Exception x) when (!(x is OperationCanceledException))
            {
                Event("STORE-UNREADABLE", "message store unreadable (" + x.Message + "): default display name and special folders", null, PstFormat.NidMessageStore);
            }
            if (string.IsNullOrEmpty(_storeName)) _storeName = Path.GetFileNameWithoutExtension(_src);
        }

        // ------------------------------------------------------------------ folders of the source
        SrcFolder _root;
        readonly Dictionary<uint, SrcFolder> _folders = new Dictionary<uint, SrcFolder>();
        ulong[] _msgs;                      // (parent << 32) | nid, sorted (levels Rebuild / Salvage)

        SrcFolder ReadFolderInfo(uint nid, uint parent)
        {
            var f = new SrcFolder { Nid = nid, Parent = parent };
            try
            {
                PstNbtEntry e;
                if (!_file.TryFindNode(nid, out e)) { f.Synthetic = true; f.InfoLost = true; return f; }
                var pc = new PstPcReader(new PstNodeView(_file, e.Nid, e.BidData, e.BidSub));
                PstProp n = Try(pc, 0x3001);
                if (n != null) f.Name = PstConvert.StringOf(n) ?? ""; else f.InfoLost = true;
                f.Class = PstConvert.StringOf(Try(pc, 0x3613));
                PstProp p = Try(pc, 0x3602); if (p != null) f.StoredContent = PstConvert.Int32Of(p, 0);
                p = Try(pc, 0x3603); if (p != null) f.StoredUnread = PstConvert.Int32Of(p, 0);
                p = Try(pc, 0x3617); f.StoredAssoc = p != null ? PstConvert.Int32Of(p, 0) : 0;
                p = Try(pc, 0x360A); if (p != null && p.Data != null) f.StoredSub = p.Data.Length > 0 && p.Data[0] != 0 ? 1 : 0;
            }
            catch (Exception x) when (!(x is OperationCanceledException))
            {
                f.InfoLost = true;
                Event("FOLDER-INFO-LOST", "folder properties unreadable (" + x.Message + ")", null, nid);
            }
            if (f.InfoLost && f.Name.Length == 0 && nid != PstFormat.NidRootFolder) f.Name = "Recovered folder 0x" + nid.ToString("X");
            return f;
        }

        void Discover()
        {
            OpenSource();
            ReadStore();
            if (_level == PstRepairLevel.Recompute) DiscoverTables(); else DiscoverNodes();
            // the writer's own names for the two containers, taken from the source
            if (_folders.ContainsKey(_ipmNid)) _ipmName = _folders[_ipmNid].Name;
        }

        // ---- level Recompute: the hierarchy tables
        void DiscoverTables()
        {
            var visited = new HashSet<uint>();
            try
            {
                _root = LoadByTables(PstFormat.NidRootFolder, PstFormat.NidRootFolder, 0, visited, true);
            }
            catch (Exception x) when (!(x is OperationCanceledException) && !(x is PstRepairDiscoveryException)) { throw new PstRepairDiscoveryException("root folder: " + x.Message, x); }
            if (_root == null) throw new PstRepairDiscoveryException("root folder: missing", null);
        }

        SrcFolder LoadByTables(uint nid, uint parent, int depth, HashSet<uint> visited, bool strict)
        {
            _ct.ThrowIfCancellationRequested();
            if (!visited.Add(nid)) { Event("FOLDER-LOOP", "folder listed twice (loop): dropped from its second parent", null, nid); return null; }
            if (depth > 64) { Event("FOLDER-DEPTH", "folder hierarchy deeper than 64 levels: subtree dropped (use level Rebuild)", null, nid); return null; }
            PstNbtEntry e;
            bool exists = _file.TryFindNode(nid, out e);
            if (!exists)
            {
                if (strict) throw new PstRepairDiscoveryException("folder node 0x" + nid.ToString("X") + " does not exist", null);
                _rep.DanglingDropped++;
                Event("DANGLING-FOLDER", "hierarchy row of a folder that does not exist dropped", null, nid);
                return null;
            }
            SrcFolder f = ReadFolderInfo(nid, parent);
            _folders[nid] = f;
            uint hn = PstFormat.Nid(PstFormat.NtHierarchyTable, PstFormat.NidIndex(nid));
            try
            {
                if (_file.HasNode(hn))
                {
                    var tc = new PstTcReader(_file.OpenNode(hn));
                    var kids = new List<uint>();
                    for (int i = 0; i < tc.RowCount; i++) kids.Add(tc.RowId(i));
                    foreach (uint k in kids)
                    {
                        SrcFolder c = LoadByTables(k, nid, depth + 1, visited, false);
                        if (c != null) f.Children.Add(c);
                    }
                }
            }
            catch (Exception x) when (!(x is OperationCanceledException) && !(x is PstRepairDiscoveryException))
            {
                if (strict) throw new PstRepairDiscoveryException("hierarchy table of the root folder: " + x.Message, x);
                Event("HIERARCHY-UNREADABLE", "hierarchy table unreadable (" + x.Message + "): subfolders not found at this level (use Rebuild)", null, nid);
            }
            return f;
        }

        // ---- levels Rebuild / Salvage: the nodes of the NBT and their parents
        void DiscoverNodes()
        {
            var list = new List<ulong>();
            int searchFolders = 0;
            try
            {
                foreach (PstNbtEntry n in _file.EnumerateNodes())
                {
                    if ((list.Count & 0xFFF) == 0) _ct.ThrowIfCancellationRequested();
                    uint t = PstFormat.NidType(n.Nid);
                    if (t == PstFormat.NtNormalFolder) _folders[n.Nid] = new SrcFolder { Nid = n.Nid, Parent = n.Parent };
                    else if (t == PstFormat.NtNormalMessage || t == PstFormat.NtAssocMessage) list.Add(((ulong)n.Parent << 32) | n.Nid);
                    else if (t == PstFormat.NtSearchFolder) searchFolders++;
                }
            }
            catch (Exception x) when (!(x is OperationCanceledException)) { throw new PstRepairDiscoveryException("node B-tree: " + x.Message, x); }
            list.Sort();
            _msgs = list.ToArray();
            _itemsTotal = _msgs.Length;
            if (searchFolders > 0) Event("SEARCH-FOLDERS-SKIPPED", searchFolders + " search folder(s) are not rewritten (the writer has no search folders; their criteria are lost, the messages they show are in their own folders)", null, 0);
            if (_folders.Count == 0 && _msgs.Length == 0) throw new PstRepairDiscoveryException("the file has no folder and no message node", null);

            foreach (SrcFolder f in _folders.Values.ToList())
            {
                SrcFolder i2 = ReadFolderInfo(f.Nid, f.Parent);
                _folders[f.Nid] = i2;
            }
            if (!_folders.TryGetValue(PstFormat.NidRootFolder, out _root))
            {
                _root = new SrcFolder { Nid = PstFormat.NidRootFolder, Parent = PstFormat.NidRootFolder, Synthetic = true };
                _folders[_root.Nid] = _root;
                Event("ROOT-REBUILT", "the root folder does not exist: a new one is used", null, PstFormat.NidRootFolder);
            }
            LinkFolders();
        }

        readonly List<SrcFolder> _recoveredTops = new List<SrcFolder>();

        void LinkFolders()
        {
            foreach (SrcFolder f in _folders.Values) f.Children.Clear();
            foreach (SrcFolder f in _folders.Values.OrderBy(x => x.Nid))
            {
                SrcFolder p;
                if (f == _root || f.Parent == f.Nid || !_folders.TryGetValue(f.Parent, out p)) continue;
                p.Children.Add(f);
            }
            var reached = Reach();
            if (reached.Count < _folders.Count)
            {
                // folders that cannot be reached (parent node gone, loop): the hierarchy tables of the others may still say where they were
                var hint = HierarchyHints();
                bool changed = false;
                foreach (SrcFolder f in _folders.Values.OrderBy(x => x.Nid))
                {
                    uint hp;
                    if (reached.Contains(f) || f == _root || !hint.TryGetValue(f.Nid, out hp) || hp == f.Parent || hp == f.Nid) continue;
                    SrcFolder p;
                    if (!_folders.TryGetValue(hp, out p)) continue;
                    SrcFolder old;
                    if (_folders.TryGetValue(f.Parent, out old)) old.Children.Remove(f);
                    f.Parent = hp; p.Children.Add(f); changed = true;
                    Event("FOLDER-REATTACHED", "folder '" + f.Name + "' cannot be reached from its parent node: re-attached to the folder whose hierarchy table lists it", null, f.Nid);
                }
                if (changed) reached = Reach();
            }
            if (reached.Count < _folders.Count)
            {
                foreach (SrcFolder f in _folders.Values.OrderBy(x => x.Nid))
                {
                    if (reached.Contains(f)) continue;
                    // walk up to the top of this component (a missing parent, a cycle, or a folder already reached)
                    SrcFolder top = f;
                    var seen = new HashSet<uint> { f.Nid };
                    SrcFolder p;
                    while (top.Parent != top.Nid && _folders.TryGetValue(top.Parent, out p) && !reached.Contains(p) && seen.Add(p.Nid)) top = p;
                    bool cycle = top.Parent != top.Nid && _folders.ContainsKey(top.Parent) && !reached.Contains(_folders[top.Parent]);
                    if (cycle) { SrcFolder q = _folders[top.Parent]; q.Children.Remove(top); }
                    else if (_folders.TryGetValue(top.Parent, out p) && reached.Contains(p)) { /* too deep: cut */ p.Children.Remove(top); }
                    top.Recovered = true;
                    _recoveredTops.Add(top);
                    string why = cycle ? "loop in the folder hierarchy" : _folders.ContainsKey(top.Parent) ? "deeper than 64 levels" : "parent folder 0x" + top.Parent.ToString("X") + " does not exist";
                    Event("FOLDER-ORPHAN", "folder '" + top.Name + "' moved under the recovered folder (" + why + ")", null, top.Nid);
                    _rep.FoldersRecovered++;
                    foreach (SrcFolder r in Subtree(top, reached)) reached.Add(r);
                }
            }
        }

        HashSet<SrcFolder> Reach()
        {
            var reached = new HashSet<SrcFolder>();
            foreach (SrcFolder r in Subtree(_root, reached, 64)) reached.Add(r);
            return reached;
        }

        static IEnumerable<SrcFolder> Subtree(SrcFolder top, HashSet<SrcFolder> already) { return Subtree(top, already, int.MaxValue); }

        static IEnumerable<SrcFolder> Subtree(SrcFolder top, HashSet<SrcFolder> already, int maxDepth)
        {
            var seen = new HashSet<SrcFolder>();
            var st = new Stack<KeyValuePair<SrcFolder, int>>();
            st.Push(new KeyValuePair<SrcFolder, int>(top, 0));
            while (st.Count > 0)
            {
                var kv = st.Pop();
                if (already.Contains(kv.Key) || !seen.Add(kv.Key)) continue;
                yield return kv.Key;
                if (kv.Value >= maxDepth) continue;
                foreach (SrcFolder c in kv.Key.Children) st.Push(new KeyValuePair<SrcFolder, int>(c, kv.Value + 1));
            }
        }

        Dictionary<uint, uint> HierarchyHints()
        {
            var hint = new Dictionary<uint, uint>();
            foreach (SrcFolder f in _folders.Values)
            {
                uint hn = PstFormat.Nid(PstFormat.NtHierarchyTable, PstFormat.NidIndex(f.Nid));
                try
                {
                    if (!_file.HasNode(hn)) continue;
                    var tc = new PstTcReader(_file.OpenNode(hn));
                    for (int i = 0; i < tc.RowCount; i++) { uint c = tc.RowId(i); if (!hint.ContainsKey(c)) hint[c] = f.Nid; }
                }
                catch (Exception x) when (!(x is OperationCanceledException)) { }
            }
            return hint;
        }

        // ------------------------------------------------------------------ writing
        void Write()
        {
            var wo = new PstWriterOptions { IpmSubtreeName = _ipmName };
            _writer = PstWriter.Create(_tmp, _storeName, wo);
            Report("repair", null, null);
            WriteTree(_root, null, 0, false);
            if (_level != PstRepairLevel.Recompute)
            {
                foreach (SrcFolder top in _recoveredTops) { top.Path = "(detached)/" + top.Name; WriteTree(top, RecoveredRoot(), 1, true); }
                WriteOrphanMessages();
            }
            _rep.NamedIdsUsed = _reader.NamedIdsUsed.Count; _rep.NamedIdsUnresolved = _reader.NamedIdsUnresolved.Count; _rep.NamedPropsDropped = _reader.NamedPropsDropped;
            if (_reader.NamedIdsUnresolved.Count > 0)
                Event("NAMEMAP-IDS-LOST", _reader.NamedIdsUnresolved.Count + " of " + _reader.NamedIdsUsed.Count + " named-property id(s) used by the messages have no entry in the name map: " + _reader.NamedPropsDropped.ToString("N0") + " propert" + (_reader.NamedPropsDropped == 1 ? "y" : "ies") + " dropped", null, PstFormat.NidNameToIdMap);
            _writer.Close();
            _writer = null;
            Report("repair", null, null, true);
        }

        PstFolder RecoveredRoot()
        {
            if (_recovered == null) _recovered = _writer.CreateFolder(null, _o.RecoveredFolderName, "IPF.Note");
            return _recovered;
        }

        string PathOf(SrcFolder f, SrcFolder parent)
        {
            if (parent == null || f == _root) return "";
            return parent.Path.Length == 0 ? "/" + f.Name : parent.Path + "/" + f.Name;
        }

        void WriteTree(SrcFolder f, PstFolder destParent, int depth, bool underRecovered)
        {
            _ct.ThrowIfCancellationRequested();
            PstFolder dest;
            bool special = true;
            if (!underRecovered && f.Nid == PstFormat.NidRootFolder) dest = _writer.Root;
            else if (!underRecovered && f.Nid == _ipmNid && _ipmNid != 0) dest = _writer.IpmSubtree;
            else if (!underRecovered && f.Nid == _deletedNid && _deletedNid != 0) dest = _writer.DeletedItems;
            else if (!underRecovered && f.Nid == _searchRootNid && _searchRootNid != 0) dest = _writer.SearchRoot;
            else
            {
                special = false;
                PstFolder parent = destParent;
                dest = _writer.CreateFolder(parent, f.Name, string.IsNullOrEmpty(f.Class) ? "IPF.Note" : f.Class);
            }
            _rep.FoldersWritten++;
            WriteFolderContents(f, dest);
            bool ipmUnknown = !_folders.ContainsKey(_ipmNid) || _ipmNid == 0;
            foreach (SrcFolder c in f.Children.ToList())
            {
                PstFolder cp = dest;
                if (!underRecovered && f.Nid == PstFormat.NidRootFolder && ipmUnknown && c.Nid != _searchRootNid) cp = _writer.IpmSubtree;
                c.Path = PathOf(c, f);
                WriteTree(c, cp, depth + 1, underRecovered);
            }
            if (!special && f.StoredSub >= 0 && (f.StoredSub == 1) != (f.Children.Count > 0))
                Event("RECOUNT-FOLDER", "subfolders flag " + (f.StoredSub == 1) + " -> " + (f.Children.Count > 0), f.Path, f.Nid);
        }

        void WriteFolderContents(SrcFolder f, PstFolder dest)
        {
            int written = 0, unread = 0, fai = 0;
            if (_level == PstRepairLevel.Recompute)
            {
                foreach (bool assoc in new[] { false, true })
                {
                    List<uint> nids = TableNids(f, assoc);
                    foreach (uint nid in nids)
                    {
                        int st = WriteMessage(nid, f, dest, f.Path, false, assoc);
                        if (st >= 0) { if (assoc) fai++; else { written++; if (st == 1) unread++; } }
                    }
                }
            }
            else
            {
                foreach (uint nid in MessagesOf(f.Nid))
                {
                    bool assoc = PstFormat.NidType(nid) == PstFormat.NtAssocMessage;
                    int st = WriteMessage(nid, f, dest, f.Path, false, assoc);
                    if (st >= 0) { if (assoc) fai++; else { written++; if (st == 1) unread++; } }
                }
            }
            if (f.StoredContent >= 0 && f.StoredContent != written) Event("RECOUNT-FOLDER", "content count " + f.StoredContent + " -> " + written, f.Path, f.Nid);
            if (f.StoredUnread >= 0 && f.StoredUnread != unread) Event("RECOUNT-FOLDER", "unread count " + f.StoredUnread + " -> " + unread, f.Path, f.Nid);
            if (f.StoredAssoc >= 0 && f.StoredAssoc != fai) Event("RECOUNT-FOLDER", "hidden (associated) count " + f.StoredAssoc + " -> " + fai, f.Path, f.Nid);
        }

        List<uint> TableNids(SrcFolder f, bool assoc)
        {
            var res = new List<uint>();
            if (f.Synthetic) return res;
            uint tn = PstFormat.Nid(assoc ? PstFormat.NtAssocContentsTable : PstFormat.NtContentsTable, PstFormat.NidIndex(f.Nid));
            try
            {
                if (!_file.HasNode(tn)) return res;
                var tc = new PstTcReader(_file.OpenNode(tn));
                var seen = new HashSet<uint>();
                for (int i = 0; i < tc.RowCount; i++) { uint id = tc.RowId(i); if (seen.Add(id)) res.Add(id); }
            }
            catch (Exception x) when (!(x is OperationCanceledException))
            {
                Event("CONTENTS-UNREADABLE", (assoc ? "hidden-items" : "contents") + " table unreadable (" + x.Message + "): the messages of this folder are not found at this level (use Rebuild)", f.Path, f.Nid);
            }
            return res;
        }

        IEnumerable<uint> MessagesOf(uint folderNid)
        {
            ulong lo = (ulong)folderNid << 32;
            int a = 0, b = _msgs.Length;
            while (a < b) { int m = (a + b) >> 1; if (_msgs[m] < lo) a = m + 1; else b = m; }
            for (int i = a; i < _msgs.Length && (_msgs[i] >> 32) == folderNid; i++) yield return (uint)_msgs[i];
        }

        void WriteOrphanMessages()
        {
            long i = 0;
            while (i < _msgs.Length)
            {
                uint parent = (uint)(_msgs[i] >> 32);
                long j = i;
                while (j < _msgs.Length && (uint)(_msgs[j] >> 32) == parent) j++;
                if (!_folders.ContainsKey(parent))
                {
                    PstFolder dest;
                    if (!_lostFolders.TryGetValue(parent, out dest))
                    {
                        dest = _writer.CreateFolder(RecoveredRoot(), parent == 0 ? "Unassigned" : "Lost folder 0x" + parent.ToString("X"), "IPF.Note");
                        _lostFolders[parent] = dest;
                    }
                    string src = parent == 0 ? "(no folder)" : "(folder 0x" + parent.ToString("X") + " does not exist)";
                    for (long k = i; k < j; k++)
                    {
                        uint nid = (uint)_msgs[k];
                        WriteMessage(nid, null, dest, src, true, PstFormat.NidType(nid) == PstFormat.NtAssocMessage);
                    }
                }
                i = j;
            }
        }

        // ------------------------------------------------------------------ one message
        /// <summary>A loss caused by an error of the B-trees (a salvage can find the blocks), not by a damaged block.</summary>
        static bool Structural(string m)
        {
            return m != null && (m.Contains("is not in the block B-tree") || m.Contains("Page CRC") || m.Contains("Page signature") || m.Contains("Page BID") || m.Contains("Unexpected page type") || m.Contains("B-tree deeper") || m.Contains("B-tree page"));
        }

        /// <summary>Returns -1 when nothing was written, else 1 when the written message is unread, 0 otherwise.</summary>
        int WriteMessage(uint nid, SrcFolder folder, PstFolder dest, string srcPath, bool orphan, bool assoc)
        {
            _ct.ThrowIfCancellationRequested();
            var item = new PstRepairItem { SourceNid = nid, SourceFolder = srcPath, TargetFolder = dest.Name, Orphan = orphan };
            int result = -1;
            try
            {
                if (_level == PstRepairLevel.Recompute && !_file.HasNode(nid))
                {
                    _rep.DanglingDropped++;
                    Event("DANGLING-ROW", "contents row of a message that does not exist dropped", srcPath, nid);
                    return -1;
                }
                var lost = new List<string>();
                FxMessage m = _reader.ReadTolerant(nid, lost, _o.VerifyValues, _ct, _o.KeepDamagedAttachments ? item.Damaged : null);
                if (_unlinked != null && _unlinked.Contains(nid)) lost.Insert(0, "node entry lost: folder, recipients and attachments");
                if (assoc) { m.Properties.RemoveAll(p => !p.IsNamed && p.Id == 0x67AA); m.Properties.Add(FxProperty.FromBool(0x67AA, true)); }
                FxProperty sj = m.Find(0x0037);
                if (sj != null && sj.Raw != null) item.Subject = FxValue.DecodeString(sj.Raw, sj.TypeInfo.BaseType, 1252).TrimEnd('\0');
                FxProperty fl = m.Find(0x0E07);
                bool unread = fl != null && fl.Raw != null && fl.Raw.Length >= 4 && (BitConverter.ToInt32(fl.Raw, 0) & 1) == 0;
                item.TargetNid = _writer.AddMessage(dest, m, _ct, null);
                item.Lost.AddRange(lost);
                foreach (PstDamagedAttachment d in item.Damaged) { _rep.AttachmentsKeptDamaged++; _rep.DamagedBytes += d.BadBytes; }
                if (_level != PstRepairLevel.Salvage && lost.Any(Structural)) _rep.StructuralLosses++;
                item.Status = lost.Count == 0 ? PstItemStatus.Recovered : PstItemStatus.Partial;
                _rep.PropertiesLost += lost.Count;
                result = unread ? 1 : 0;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception x)
            {
                item.Status = PstItemStatus.Lost;
                item.Lost.Add(x.Message);
                if (_level != PstRepairLevel.Salvage && Structural(x.Message)) _rep.StructuralLosses++;
            }
            if (orphan)
            {
                _rep.OrphanItems++;
                if (item.Status != PstItemStatus.Lost) Event("ORPHAN-RECOVERED", "message whose folder does not exist kept under '" + _o.RecoveredFolderName + "/" + dest.Name + "'", srcPath, nid);
            }
            switch (item.Status)
            {
                case PstItemStatus.Recovered: _rep.ItemsRecovered++; break;
                case PstItemStatus.Partial: _rep.ItemsPartial++; break;
                default: _rep.ItemsLost++; break;
            }
            _itemsDone++;
            if (_o.ItemSink != null) _o.ItemSink(item);
            if (_rep.Items.Count < _o.MaxItemRecords) _rep.Items.Add(item);
            else if (item.Status != PstItemStatus.Recovered && _problemsKept < _o.MaxItemRecords) { _rep.Items.Add(item); _problemsKept++; }
            if (_clock.ElapsedMilliseconds - _lastReport > 200) Report("repair", srcPath, item.Subject);
            return result;
        }

        void Report(string phase, string folder, string item) { Report(phase, folder, item, false); }

        void Report(string phase, string folder, string item, bool force)
        {
            _lastReport = _clock.ElapsedMilliseconds;
            if (_progress != null) _progress.Report(new PstProgress { Phase = phase, Items = _itemsDone, ItemsTotal = _itemsTotal, CurrentFolder = folder, CurrentItem = item });
        }
    }

    public sealed partial class PstReader
    {
        /// <summary>A reader over an already opened file (a salvage index, a damaged store): nothing is read until it is used.</summary>
        public static PstReader Wrap(PstFile f)
        {
            return new PstReader(f);
        }

        /// <summary>
        /// Repair: reads a message as far as it decodes. The parts that cannot be read (a property, a recipient, an attachment, an
        /// embedded message, a big value whose data tree is damaged) are left out and described in lost; a message whose property
        /// context cannot be read at all raises.
        /// </summary>
        internal FxMessage ReadTolerant(uint nid, List<string> lost, bool verify, System.Threading.CancellationToken ct)
        {
            return ReadTolerant(nid, lost, verify, ct, null);
        }

        /// <summary>A list in damaged: the attachments whose data tree has unreadable blocks are kept with zeros there (PstRepairOptions.KeepDamagedAttachments) and described in it.</summary>
        internal FxMessage ReadTolerant(uint nid, List<string> lost, bool verify, System.Threading.CancellationToken ct, List<PstDamagedAttachment> damaged)
        {
            _damagedSink = damaged;
            try { return ReadTol(File.OpenNode(nid), 0, lost, "", verify, ct); }
            finally { _damagedSink = null; }
        }

        List<PstDamagedAttachment> _damagedSink;

        static string Tag(ushort id, ushort type) { return "0x" + id.ToString("X4") + type.ToString("X4"); }

        static void VerifyBlob(FxBlob b, System.Threading.CancellationToken ct)
        {
            byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(64 * 1024);
            try
            {
                using (System.IO.Stream s = b.OpenRead())
                {
                    long left = b.Length;
                    while (left > 0)
                    {
                        ct.ThrowIfCancellationRequested();
                        int n = s.Read(buf, 0, (int)Math.Min(buf.Length, left));
                        if (n <= 0) throw new PstFormatException("The value ends before its length");
                        left -= n;
                    }
                }
            }
            finally { System.Buffers.ArrayPool<byte>.Shared.Return(buf); }
        }

        /// <summary>Converts a decoded property; null (and a note in lost) when it cannot be kept.</summary>
        FxProperty KeepProp(PstProp p, List<string> lost, string where, bool verify, System.Threading.CancellationToken ct)
        {
            if (p.ObjectView != null) return null;
            if (p.Id >= 0x8000)
            {
                PstNameMap.Entry en;
                NamedIdsUsed.Add(p.Id);
                if (!Names.TryGetValue(p.Id, out en)) { NamedIdsUnresolved.Add(p.Id); NamedPropsDropped++; lost.Add(where + "named property 0x" + p.Id.ToString("X4") + " (no entry in the name map)"); return null; }
            }
            if (verify && p.Source != null)
            {
                try { VerifyBlob(p.Source, ct); }
                catch (OperationCanceledException) { throw; }
                catch (Exception x) { lost.Add(where + "property " + Tag(p.Id, p.Type) + " (value damaged: " + x.Message + ")"); return null; }
            }
            return ToFxNamed(p);
        }

        /// <summary>
        /// The data of a file attachment (KeepDamagedAttachments): the value when it reads, else the blocks that read with zeros in place of
        /// the others (damaged describes it), else null (nothing readable, or a tree that cannot be walked: dropped as without the option).
        /// </summary>
        FxProperty KeepDamagedData(PstProp p, string where, int index, List<string> lost, System.Threading.CancellationToken ct, out PstDamagedAttachment damaged)
        {
            damaged = null;
            try { VerifyBlob(p.Source, ct); return ToFxNamed(p); }
            catch (OperationCanceledException) { throw; }
            catch (Exception x)
            {
                var db = (PstDataBlob)p.Source;
                PstSalvagedBlob sb = PstSalvagedBlob.TryCreate(db.File, db.Bid, ct);
                if (sb == null) { lost.Add(where + "property " + Tag(p.Id, p.Type) + " (value damaged: " + x.Message + ")"); return null; }
                damaged = new PstDamagedAttachment { Where = where.TrimEnd(), Length = sb.Length, BadBytes = sb.BadBytes };
                damaged.BadRanges.AddRange(sb.BadRanges);
                lost.Add(where + "data damaged, kept with zeros in place of the unreadable blocks: " + (sb.Length - sb.BadBytes).ToString("N0") + " of " + sb.Length.ToString("N0") + " bytes read; zero-filled " + string.Join(", ", sb.BadRanges.Take(8)) + (sb.BadRanges.Count > 8 ? ", ... (" + (sb.BadRanges.Count - 8) + " more)" : ""));
                return FxProperty.CreateLazy(0x3701, 0x0102, sb);
            }
        }

        /// <summary>"name.ext" becomes "name (damaged).ext" in the long file name and the display name (added when missing) and in a short file name equal to them; a nameless attachment gets "attachment (damaged)".</summary>
        static void MarkDamagedNames(FxAttachment a, PstDamagedAttachment d)
        {
            Func<ushort, FxProperty> find = id => a.Properties.FirstOrDefault(x => !x.IsNamed && x.Id == id && x.Raw != null && (x.TypeInfo.BaseType == 0x001F || x.TypeInfo.BaseType == 0x001E));
            Func<FxProperty, string> text = pr => FxValue.DecodeString(pr.Raw, pr.TypeInfo.BaseType, 1252).TrimEnd('\0');
            FxProperty lng = find(0x3707), disp = find(0x3001), shrt = find(0x3704);
            string original = lng != null ? text(lng) : disp != null ? text(disp) : shrt != null ? text(shrt) : null;
            string marked = PstDamagedName.Mark(original);
            d.OriginalName = original;
            d.NewName = marked;
            Action<ushort, FxProperty> put = (id, old) =>
            {
                int at = old == null ? -1 : a.Properties.IndexOf(old);
                if (at >= 0) a.Properties[at] = FxProperty.FromString(id, marked); else a.Properties.Add(FxProperty.FromString(id, marked));
            };
            put(0x3707, lng);
            put(0x3001, disp);
            if (shrt != null && (original == null || text(shrt) == original)) put(0x3704, shrt);
        }

        FxMessage ReadTol(PstNodeView node, int depth, List<string> lost, string where, bool verify, System.Threading.CancellationToken ct)
        {
            if (depth > 32) throw new PstFormatException("Embedded messages nested deeper than 32 levels");
            var m = new FxMessage();
            var pc = new PstPcReader(node) { LazyThreshold = LazyThreshold };
            foreach (PstPcReader.Entry e in pc.Entries)
            {
                ct.ThrowIfCancellationRequested();
                try
                {
                    FxProperty f = KeepProp(pc.Decode(e), lost, where, verify, ct);
                    if (f != null) m.Properties.Add(f);
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception x) { lost.Add(where + "property " + Tag(e.Id, e.Type) + " (" + x.Message + ")"); }
            }
            bool hasRec, hasAtt;
            try { hasRec = node.HasSubnode(PstFormat.NidRecipientTable); hasAtt = node.HasSubnode(PstFormat.NidAttachmentTable); }
            catch (Exception x) when (!(x is OperationCanceledException))
            {
                lost.Add(where + "recipients and attachments (subnode tree: " + x.Message + ")");
                return m;
            }
            if (hasRec)
            {
                try
                {
                    var tc = new PstTcReader(node.OpenSubnode(PstFormat.NidRecipientTable));
                    for (int i = 0; i < tc.RowCount; i++)
                    {
                        try
                        {
                            var r = new FxRecipient();
                            bool partial = false;
                            foreach (PstTcReader.Col c in tc.Columns)
                            {
                                if (c.Id == 0x67F2 || c.Id == 0x67F3) continue;
                                try
                                {
                                    PstProp p = tc.GetCell(i, c);
                                    if (p == null) continue;
                                    FxProperty f = KeepProp(p, lost, where + "recipient " + i + " ", verify, ct);
                                    if (f != null) r.Properties.Add(f);
                                }
                                catch (OperationCanceledException) { throw; }
                                catch (Exception) { partial = true; }
                            }
                            if (partial) lost.Add(where + "recipient " + i + " (some properties unreadable)");
                            m.Recipients.Add(r);
                        }
                        catch (OperationCanceledException) { throw; }
                        catch (Exception x) { lost.Add(where + "recipient " + i + " (" + x.Message + ")"); }
                    }
                }
                catch (OperationCanceledException) { throw; }
                catch (Exception x) { lost.Add(where + "recipient table (" + x.Message + ")"); }
            }
            if (hasAtt)
            {
                PstTcReader tc = null;
                try { tc = new PstTcReader(node.OpenSubnode(PstFormat.NidAttachmentTable)); }
                catch (OperationCanceledException) { throw; }
                catch (Exception x) { lost.Add(where + "attachment table (" + x.Message + ")"); }
                for (int i = 0; tc != null && i < tc.RowCount; i++)
                {
                    string aw = where + "attachment " + i + " ";
                    try
                    {
                        uint anid = tc.RowId(i);
                        PstNodeView an = node.OpenSubnode(anid);
                        // KeepDamagedAttachments: every long value of the attachment is a data tree read block by block, so that a bad block shows up as such
                        var apc = new PstPcReader(an) { LazyThreshold = _damagedSink != null ? 0 : LazyThreshold };
                        var a = new FxAttachment();
                        a.Properties.Add(FxProperty.FromInt32(0x0E21, i));
                        bool embeddedLost = false;
                        PstDamagedAttachment damagedHere = null;
                        foreach (PstPcReader.Entry e in apc.Entries)
                        {
                            ct.ThrowIfCancellationRequested();
                            try
                            {
                                PstProp p = apc.Decode(e);
                                if (p.ObjectView != null)
                                {
                                    try
                                    {
                                        a.Embedded = ReadTol(p.ObjectView, depth + 1, lost, aw + "embedded message: ", verify, ct);
                                        a.Properties.Add(FxProperty.Create(0x3701, 0x000D, new byte[0]));
                                    }
                                    catch (OperationCanceledException) { throw; }
                                    catch (Exception x) { embeddedLost = true; lost.Add(aw + "embedded message (" + x.Message + ")"); }
                                }
                                else
                                {
                                    FxProperty f = null;
                                    if (_damagedSink != null && p.Id == 0x3701 && p.Type == 0x0102 && p.Source is PstDataBlob) f = KeepDamagedData(p, aw, i, lost, ct, out damagedHere);
                                    else f = KeepProp(p, lost, aw, verify, ct);
                                    if (f != null) a.Properties.Add(f);
                                }
                            }
                            catch (OperationCanceledException) { throw; }
                            catch (Exception x) { lost.Add(aw + "property " + Tag(e.Id, e.Type) + " (" + x.Message + ")"); }
                        }
                        if (damagedHere != null) { MarkDamagedNames(a, damagedHere); _damagedSink.Add(damagedHere); }
                        if (a.Embedded == null)
                        {
                            FxProperty meth = a.Properties.FirstOrDefault(x => !x.IsNamed && x.Id == 0x3705);
                            bool method5 = meth != null && meth.Raw != null && meth.Raw.Length >= 4 && BitConverter.ToInt32(meth.Raw, 0) == 5;
                            if (embeddedLost || method5) a.Properties.RemoveAll(x => !x.IsNamed && (x.Id == 0x3705 || x.Id == 0x3701));
                            // the embedded message or the data is lost: an attachment of 0 bytes, the other properties (name, ...) are kept
                            if (!a.Properties.Any(x => !x.IsNamed && x.Id == 0x3701) && (embeddedLost || method5 || apc.Find(0x3701) != null))
                                a.Properties.Add(FxProperty.FromBinary(0x3701, new byte[0]));
                        }                        m.Attachments.Add(a);
                    }
                    catch (OperationCanceledException) { throw; }
                    catch (Exception x) { lost.Add(aw + "(" + x.Message + ")"); }
                }
            }
            return m;
        }
    }
}
