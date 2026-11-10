// =============================================================================
//  Mailbox Transfer - engine, part 5j: import from PST - journal and job
// -----------------------------------------------------------------------------
//  Creation only: nothing is ever replaced, moved or deleted in the target. The items come straight from the PST (no copy in a store):
//  each one is read with PstReader, made ready for importItem (FxImportStream) and streamed as base64 into the request (bounded memory).
//
//  Journal (SQLite, one file, tables in PstImportJournal): one row per message, keyed by (job, PST file, node id, fingerprint). The fingerprint
//  is computed from the node's data blocks, the identifying properties and the size, so a message that changed in the file is a new item and
//  an unchanged one is never added twice. States: Pending -> InFlight -> Done / Failed / Unknown / Refused:
//    Done     created (target item id kept, SHA-256 of the stream that was sent);
//    Failed   the service refused the request with nothing created (4xx other than a content refusal, 429, 503, upload interrupted): tried again by the
//             next run while attempts remain;
//    Refused  this item cannot be imported and trying again will not help (Failed-final): it cannot be read from the PST, cannot be expressed as a
//             FastTransfer stream, the service refuses its content (HTTP 400, 413, 415, 422), or it failed with nothing created on every attempt allowed
//             (-MaxAttempts). Reported with the reason; the run goes on. A Refused row is final: the job is complete when every row is Done or Refused
//             (exit code 1 when some are Refused, see PstImportRunResult.ExitCode);
//    Unknown  timeout, network error after the whole body was sent, 5xx, interruption: the item may exist. A crash turns the InFlight rows into Unknown.
//  Reconciliation (on by default) of an Unknown row: after the settle delay the target folder is listed; the items created after the attempt began that no
//  Done row accounts for are candidates; the service gives an item another size than the stream that was sent, so each candidate is exported and compared
//  with the message (class, subject, dates, message id, display name). One match = Done; none = imported again (attempts allowed); several = left Unknown.
//  A run that finds its journal again resumes: Done rows are never sent again, so a PST (or a part of a split PST) can be imported again without duplicates.
// =============================================================================
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using Microsoft.Data.Sqlite;

namespace MailboxTransfer
{
    public sealed class PstJournalItem
    {
        public long ItemKey, JobId;
        public string PstId, Fingerprint, PstPath, TargetPath, TargetMailbox, TargetFolder, Subject, MessageClass, State, Error, TargetItem, Sha256;
        public uint Nid;
        public long Size;
        public int Attempts;
        public DateTime? AttemptStartUtc, UpdatedUtc;
    }

    public sealed class PstJournalStat { public string State; public long Count, Bytes; }

    /// <summary>Rows of a state in one folder of the target (the progress per folder of the job).</summary>
    public sealed class PstJournalFolderCount { public string TargetPath; public long Count; }

    /// <summary>The journal of an import from PST (a SQLite file of its own, next to the store).</summary>
    public sealed class PstImportJournal : IDisposable
    {
        readonly SqliteConnection _c;
        readonly FileStream _lockFile;
        readonly object _l = new object();
        public int RecoveredToUnknown { get; private set; }

        public PstImportJournal(string path)
        {
            string dir = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!string.IsNullOrEmpty(dir)) Directory.CreateDirectory(dir);
            // one process at a time (like the export store): two runs of the same journal would send the same items twice
            try { _lockFile = new FileStream(Path.GetFullPath(path) + ".lock", FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.DeleteOnClose); }
            catch (IOException ex) { throw new IOException("The import journal is already open by another process: " + path, ex); }
            try
            {
            _c = new SqliteConnection(new SqliteConnectionStringBuilder { DataSource = path, Mode = SqliteOpenMode.ReadWriteCreate, Pooling = false }.ToString());
            _c.Open();
            // FULL: every state of a row (InFlight, Done...) is on disk when the call returns; a lost one could send an item twice
            Exec("PRAGMA journal_mode=WAL;"); Exec("PRAGMA synchronous=FULL;"); Exec("PRAGMA busy_timeout=10000;");
            Exec(@"CREATE TABLE IF NOT EXISTS job(jobId INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, targetMailbox TEXT NOT NULL, targetRoot TEXT, createdUtc TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS pst(jobId INTEGER NOT NULL, pstId TEXT NOT NULL, sourcePath TEXT, workPath TEXT, length INTEGER, lastWriteTicks INTEGER, checkedUtc TEXT, errors INTEGER, warnings INTEGER, repaired INTEGER, note TEXT, PRIMARY KEY(jobId, pstId));
CREATE TABLE IF NOT EXISTS item(itemKey INTEGER PRIMARY KEY AUTOINCREMENT, jobId INTEGER NOT NULL, pstId TEXT NOT NULL, nid INTEGER NOT NULL, fingerprint TEXT NOT NULL,
  pstPath TEXT, targetPath TEXT, targetMailbox TEXT, targetFolder TEXT, subject TEXT, class TEXT, size INTEGER, state TEXT NOT NULL, attempts INTEGER NOT NULL DEFAULT 0,
  error TEXT, targetItem TEXT, sha256 TEXT, attemptStartUtc TEXT, updatedUtc TEXT,
  UNIQUE(jobId, pstId, nid, fingerprint));
CREATE INDEX IF NOT EXISTS item_state ON item(jobId, state);");
            // a journal written before the options of a job were recorded
            using (SqliteCommand cmd = _c.CreateCommand())
            {
                cmd.CommandText = "SELECT COUNT(*) FROM pragma_table_info('job') WHERE name='options';";
                if (Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) == 0) Exec("ALTER TABLE job ADD COLUMN options TEXT;");
            }
            lock (_l) RecoveredToUnknown = Exec("UPDATE item SET state='Unknown', error='interrupted: the run stopped while the request was in flight, the item may exist', updatedUtc=@t WHERE state='InFlight';", "@t", Now());
            }
            catch
            {
                if (_c != null) _c.Dispose();
                _lockFile.Dispose();
                throw;
            }
        }

        public void Dispose() { lock (_l) _c.Dispose(); _lockFile.Dispose(); }

        static string Now() { return DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture); }
        static DateTime? Parse(object v) { return v == null || v is DBNull ? (DateTime?)null : DateTime.Parse((string)v, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind).ToUniversalTime(); }

        int Exec(string sql, params object[] nv)
        {
            using (SqliteCommand cmd = _c.CreateCommand())
            {
                cmd.CommandText = sql;
                for (int i = 0; i + 1 < nv.Length; i += 2) cmd.Parameters.AddWithValue((string)nv[i], nv[i + 1] ?? DBNull.Value);
                return cmd.ExecuteNonQuery();
            }
        }

        T Scalar<T>(string sql, params object[] nv)
        {
            using (SqliteCommand cmd = _c.CreateCommand())
            {
                cmd.CommandText = sql;
                for (int i = 0; i + 1 < nv.Length; i += 2) cmd.Parameters.AddWithValue((string)nv[i], nv[i + 1] ?? DBNull.Value);
                object o = cmd.ExecuteScalar();
                return o == null || o is DBNull ? default(T) : (T)Convert.ChangeType(o, typeof(T), CultureInfo.InvariantCulture);
            }
        }

        public long FindJob(string name)
        {
            lock (_l) return Scalar<long>("SELECT jobId FROM job WHERE name=@n;", "@n", name);
        }

        public string JobTarget(long jobId) { lock (_l) return Scalar<string>("SELECT targetMailbox FROM job WHERE jobId=@j;", "@j", jobId); }

        /// <summary>The target options the job was started with (see PstImportJob.TargetOptions), null for a job of an older journal.</summary>
        public string JobOptions(long jobId) { lock (_l) return Scalar<string>("SELECT options FROM job WHERE jobId=@j;", "@j", jobId); }
        public void SetJobOptions(long jobId, string options) { lock (_l) Exec("UPDATE job SET options=@o WHERE jobId=@j;", "@o", options, "@j", jobId); }

        /// <summary>
        /// The target folder of a path exists now: the rows of that folder that were left Failed because it could not be created (no target folder recorded) get it and go back to Pending.
        /// </summary>
        public int AttachFolder(long jobId, string pstId, string targetPath, string targetMailbox, string folderId)
        {
            lock (_l) return Exec("UPDATE item SET targetFolder=@f, targetMailbox=@m, state='Pending', error=NULL, attempts=0, updatedUtc=@t WHERE jobId=@j AND pstId=@p AND targetPath=@tp AND targetFolder IS NULL AND state='Failed';",
                "@f", folderId, "@m", targetMailbox, "@t", Now(), "@j", jobId, "@p", pstId, "@tp", targetPath);
        }

        public long CreateJob(string name, string targetMailbox, string targetRoot)
        {
            lock (_l)
            {
                Exec("INSERT INTO job(name, targetMailbox, targetRoot, createdUtc) VALUES(@n, @m, @r, @t);", "@n", name, "@m", targetMailbox, "@r", targetRoot, "@t", Now());
                return Scalar<long>("SELECT jobId FROM job WHERE name=@n;", "@n", name);
            }
        }

        public PstImportSource GetPst(long jobId, string pstId)
        {
            lock (_l)
                using (SqliteCommand cmd = _c.CreateCommand())
                {
                    cmd.CommandText = "SELECT sourcePath, workPath, length, lastWriteTicks, errors, warnings, repaired, note, checkedUtc FROM pst WHERE jobId=@j AND pstId=@p;";
                    cmd.Parameters.AddWithValue("@j", jobId); cmd.Parameters.AddWithValue("@p", pstId);
                    using (SqliteDataReader r = cmd.ExecuteReader())
                    {
                        if (!r.Read()) return null;
                        return new PstImportSource
                        {
                            PstId = pstId, SourcePath = r.IsDBNull(0) ? null : r.GetString(0), WorkPath = r.IsDBNull(1) ? null : r.GetString(1), Length = r.IsDBNull(2) ? 0 : r.GetInt64(2),
                            LastWriteTicks = r.IsDBNull(3) ? 0 : r.GetInt64(3), CheckErrors = r.IsDBNull(4) ? 0 : r.GetInt32(4), CheckWarnings = r.IsDBNull(5) ? 0 : r.GetInt32(5),
                            Repaired = !r.IsDBNull(6) && r.GetInt32(6) != 0, Note = r.IsDBNull(7) ? null : r.GetString(7), Checked = !r.IsDBNull(8)
                        };
                    }
                }
        }

        /// <summary>The identity under which an earlier run of the job recorded this path when it is not <paramref name="pstId"/> (the file changed), else null.</summary>
        public string OtherPstIdForPath(long jobId, string sourcePath, string pstId)
        {
            lock (_l) return Scalar<string>("SELECT pstId FROM pst WHERE jobId=@j AND sourcePath=@s AND pstId<>@p LIMIT 1;", "@j", jobId, "@s", sourcePath, "@p", pstId);
        }

        public void SavePst(long jobId, PstImportSource s)
        {
            lock (_l)
                Exec(@"INSERT OR REPLACE INTO pst(jobId, pstId, sourcePath, workPath, length, lastWriteTicks, checkedUtc, errors, warnings, repaired, note) VALUES(@j,@p,@s,@w,@l,@t,@c,@e,@wa,@r,@n);",
                    "@j", jobId, "@p", s.PstId, "@s", s.SourcePath, "@w", s.WorkPath, "@l", s.Length, "@t", s.LastWriteTicks, "@c", s.Checked ? Now() : null,
                    "@e", s.CheckErrors, "@wa", s.CheckWarnings, "@r", s.Repaired ? 1 : 0, "@n", s.Note);
        }

        /// <summary>Adds the rows that are not there yet (same job, file, node and fingerprint); returns how many were new.</summary>
        public int AddItems(long jobId, IEnumerable<PstJournalItem> items) { List<PstJournalItem> refused; return AddItems(jobId, items, out refused); }

        /// <summary>Same; <paramref name="addedRefused"/> gets the new rows that were added already Refused (unreadable in the PST).</summary>
        public int AddItems(long jobId, IEnumerable<PstJournalItem> items, out List<PstJournalItem> addedRefused)
        {
            int added = 0;
            addedRefused = new List<PstJournalItem>();
            lock (_l)
            {
                using (SqliteTransaction tx = _c.BeginTransaction())
                {
                    foreach (PstJournalItem it in items)
                    {
                        using (SqliteCommand cmd = _c.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            cmd.CommandText = @"INSERT OR IGNORE INTO item(jobId, pstId, nid, fingerprint, pstPath, targetPath, targetMailbox, targetFolder, subject, class, size, state, error, updatedUtc)
VALUES(@j,@p,@n,@f,@pp,@tp,@tm,@tf,@s,@c,@z,@st,@e,@u);";
                            cmd.Parameters.AddWithValue("@j", jobId); cmd.Parameters.AddWithValue("@p", it.PstId); cmd.Parameters.AddWithValue("@n", (long)it.Nid);
                            cmd.Parameters.AddWithValue("@f", it.Fingerprint); cmd.Parameters.AddWithValue("@pp", (object)it.PstPath ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@tp", (object)it.TargetPath ?? DBNull.Value); cmd.Parameters.AddWithValue("@tm", (object)it.TargetMailbox ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@tf", (object)it.TargetFolder ?? DBNull.Value); cmd.Parameters.AddWithValue("@s", (object)it.Subject ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@c", (object)it.MessageClass ?? DBNull.Value); cmd.Parameters.AddWithValue("@z", it.Size);
                            cmd.Parameters.AddWithValue("@st", it.State ?? "Pending"); cmd.Parameters.AddWithValue("@e", (object)it.Error ?? DBNull.Value);
                            cmd.Parameters.AddWithValue("@u", Now());
                            int n = cmd.ExecuteNonQuery();
                            added += n;
                            if (n > 0 && it.State == "Refused") addedRefused.Add(it);
                        }
                    }
                    tx.Commit();
                }
            }
            return added;
        }

        static PstJournalItem Read(SqliteDataReader r)
        {
            Func<int, string> s = i => r.IsDBNull(i) ? null : r.GetString(i);
            return new PstJournalItem
            {
                ItemKey = r.GetInt64(0), JobId = r.GetInt64(1), PstId = s(2), Nid = (uint)r.GetInt64(3), Fingerprint = s(4), PstPath = s(5), TargetPath = s(6), TargetMailbox = s(7), TargetFolder = s(8),
                Subject = s(9), MessageClass = s(10), Size = r.IsDBNull(11) ? 0 : r.GetInt64(11), State = s(12), Attempts = r.GetInt32(13), Error = s(14), TargetItem = s(15), Sha256 = s(16),
                AttemptStartUtc = Parse(r.IsDBNull(17) ? null : r.GetString(17)), UpdatedUtc = Parse(r.IsDBNull(18) ? null : r.GetString(18))
            };
        }

        const string Cols = "itemKey, jobId, pstId, nid, fingerprint, pstPath, targetPath, targetMailbox, targetFolder, subject, class, size, state, attempts, error, targetItem, sha256, attemptStartUtc, updatedUtc";

        public PstJournalItem GetItem(long itemKey)
        {
            lock (_l)
                using (SqliteCommand cmd = _c.CreateCommand())
                {
                    cmd.CommandText = "SELECT " + Cols + " FROM item WHERE itemKey=@k;"; cmd.Parameters.AddWithValue("@k", itemKey);
                    using (SqliteDataReader r = cmd.ExecuteReader()) return r.Read() ? Read(r) : null;
                }
        }

        /// <summary>The next Pending row, now InFlight (one more attempt, the attempt starts now), or null.</summary>
        public PstJournalItem NextPending(long jobId)
        {
            lock (_l)
            {
                long key = Scalar<long>("SELECT itemKey FROM item WHERE jobId=@j AND state='Pending' ORDER BY itemKey LIMIT 1;", "@j", jobId);
                if (key == 0) return null;
                string t = Now();
                Exec("UPDATE item SET state='InFlight', attempts=attempts+1, attemptStartUtc=@t, updatedUtc=@t WHERE itemKey=@k;", "@t", t, "@k", key);
                return GetItem(key);
            }
        }

        public List<PstJournalItem> GetItems(long jobId, string state, int limit = int.MaxValue)
        {
            var list = new List<PstJournalItem>();
            lock (_l)
                using (SqliteCommand cmd = _c.CreateCommand())
                {
                    cmd.CommandText = "SELECT " + Cols + " FROM item WHERE jobId=@j" + (state == null ? "" : " AND state=@s") + " ORDER BY itemKey LIMIT @n;";
                    cmd.Parameters.AddWithValue("@j", jobId); cmd.Parameters.AddWithValue("@n", limit);
                    if (state != null) cmd.Parameters.AddWithValue("@s", state);
                    using (SqliteDataReader r = cmd.ExecuteReader()) while (r.Read()) list.Add(Read(r));
                }
            return list;
        }

        public HashSet<string> DoneTargetItems(long jobId)
        {
            var set = new HashSet<string>(StringComparer.Ordinal);
            lock (_l)
                using (SqliteCommand cmd = _c.CreateCommand())
                {
                    cmd.CommandText = "SELECT targetItem FROM item WHERE jobId=@j AND state='Done' AND targetItem IS NOT NULL;"; cmd.Parameters.AddWithValue("@j", jobId);
                    using (SqliteDataReader r = cmd.ExecuteReader()) while (r.Read()) set.Add(r.GetString(0));
                }
            return set;
        }

        public List<PstJournalStat> Stats(long jobId)
        {
            var list = new List<PstJournalStat>();
            lock (_l)
                using (SqliteCommand cmd = _c.CreateCommand())
                {
                    cmd.CommandText = "SELECT state, COUNT(*), COALESCE(SUM(size),0) FROM item WHERE jobId=@j GROUP BY state;"; cmd.Parameters.AddWithValue("@j", jobId);
                    using (SqliteDataReader r = cmd.ExecuteReader()) while (r.Read()) list.Add(new PstJournalStat { State = r.GetString(0), Count = r.GetInt64(1), Bytes = r.GetInt64(2) });
                }
            return list;
        }

        public long Count(long jobId, string state) { var s = Stats(jobId).FirstOrDefault(x => x.State == state); return s == null ? 0 : s.Count; }

        /// <summary>The rows of a state, per folder of the target (path below the root of the mailbox).</summary>
        public List<PstJournalFolderCount> CountByTarget(long jobId, string state)
        {
            var list = new List<PstJournalFolderCount>();
            lock (_l)
                using (SqliteCommand cmd = _c.CreateCommand())
                {
                    cmd.CommandText = "SELECT COALESCE(targetPath,''), COUNT(*) FROM item WHERE jobId=@j AND state=@s GROUP BY targetPath;";
                    cmd.Parameters.AddWithValue("@j", jobId); cmd.Parameters.AddWithValue("@s", state);
                    using (SqliteDataReader r = cmd.ExecuteReader()) while (r.Read()) list.Add(new PstJournalFolderCount { TargetPath = r.GetString(0), Count = r.GetInt64(1) });
                }
            return list;
        }

        void Set(long key, string state, string error, string targetItem, string sha)
        {
            lock (_l)
                Exec("UPDATE item SET state=@s, error=@e, targetItem=COALESCE(@i, targetItem), sha256=COALESCE(@h, sha256), updatedUtc=@t WHERE itemKey=@k;",
                    "@s", state, "@e", error, "@i", targetItem, "@h", sha, "@t", Now(), "@k", key);
        }

        public void MarkDone(long key, string targetItem, string sha) { Set(key, "Done", null, targetItem, sha); }
        public void MarkFailed(long key, string error) { Set(key, "Failed", error, null, null); }
        public void MarkRefused(long key, string error) { Set(key, "Refused", error, null, null); }
        public void MarkUnknown(long key, string error) { Set(key, "Unknown", error, null, null); }
        public void MarkPending(long key)
        {
            lock (_l) Exec("UPDATE item SET state='Pending', attempts=MAX(attempts-1,0), updatedUtc=@t WHERE itemKey=@k;", "@t", Now(), "@k", key);
        }
        /// <summary>An Unknown row settled by a lookup (or by a person): Done with the item found, or back to Pending to be sent again.</summary>
        public void ResolveUnknown(long key, bool imported, string targetItem)
        {
            if (imported) Set(key, "Done", null, targetItem, null);
            else lock (_l) Exec("UPDATE item SET state='Pending', error=NULL, updatedUtc=@t WHERE itemKey=@k;", "@t", Now(), "@k", key);
        }

        public int RequeueFailed(long jobId, int maxAttempts)
        {
            lock (_l) return Exec("UPDATE item SET state='Pending', updatedUtc=@t WHERE jobId=@j AND state='Failed' AND attempts<@m AND targetFolder IS NOT NULL;", "@t", Now(), "@j", jobId, "@m", maxAttempts);
        }
    }

    /// <summary>One line per item that was not imported, for the report.</summary>
    public sealed class PstImportItemReport
    {
        public string Pst, Folder, Subject, MessageClass, State, Reason;
        public uint Nid;
        public long Size;
        public int Attempts;
    }

    public sealed class PstImportRunResult
    {
        public long JobId;
        public string JobName;
        public long Done, Failed, Unknown, Refused, Bytes;
        public int FoldersCreated, FoldersFailed, SessionsCreated;
        public bool JobKnown;
        public long JobTotal, JobDone, JobFailed, JobUnknown, JobPending, JobRefused, JobUnknownEarlier;
        public int RecoveredToUnknown;
        public bool AutoReconcile;
        public long ReconciledCreated, ReconciledReimported, ReconcileAmbiguous, ReconcileExhausted;
        public List<double> SessionLifetimeSeconds = new List<double>();
        public List<string> Problems = new List<string>();
        /// <summary>The items of the job that are not Done (Failed, Refused, Unknown), with the reason; at most MaxReports.</summary>
        public List<PstImportItemReport> NotImported = new List<PstImportItemReport>();
        public List<PstImportSource> Sources = new List<PstImportSource>();
        public long AssociatedSkipped;
        public TransferSnapshot Final;
        public bool Completed;
        public bool Planned;
        /// <summary>Nothing is left to do: every row is Done or Refused (refused for good, see the header of this file). A job with Refused rows is complete, but not clean.</summary>
        public bool JobComplete { get { return JobKnown && Planned && JobDone + JobRefused == JobTotal; } }
        /// <summary>The job is complete and every row is Done.</summary>
        public bool JobClean { get { return JobComplete && JobRefused == 0; } }

        /// <summary>
        /// 0 complete, every row Done; 3 Unknown rows (settle them); 2 not complete (Pending or Failed rows: run again); 1 complete, but some items were refused for good
        /// (nothing to run again, the report lists them). 1 is also the code of an error that stopped the run (decided by the caller).
        /// </summary>
        public int ExitCode { get { return JobUnknown > 0 ? 3 : !JobComplete ? 2 : JobRefused > 0 ? 1 : 0; } }

        public string ReconcileSummary()
        {
            if (!AutoReconcile || ReconciledCreated + ReconciledReimported + ReconcileAmbiguous + ReconcileExhausted == 0) return null;
            return "reconciliation: " + ReconciledCreated + " Unknown row(s) found created in the target (marked Done), " + ReconciledReimported + " found absent and imported again, " +
                ReconcileAmbiguous + " left Unknown because several items could be them" + (ReconcileExhausted > 0 ? ", " + ReconcileExhausted + " left Unknown after the allowed rounds" : "");
        }

        public string JobSummary()
        {
            string name = "job '" + JobName + "'";
            if (!JobKnown) return name + " state not read: the run stopped before the job was opened";
            if (JobClean) return name + " complete: " + JobDone + " of " + JobTotal + " done";
            if (JobComplete) return name + " complete, " + JobRefused + " item(s) refused for good: " + JobDone + " of " + JobTotal + " done (the report lists them and why; running again does not change it)";
            var parts = new List<string>();
            if (JobUnknown > 0) parts.Add(JobUnknown + " unknown (the item may exist, never replayed by itself)");
            if (JobFailed > 0) parts.Add(JobFailed + " failed (nothing created, the next run tries again while attempts remain)");
            if (JobRefused > 0) parts.Add(JobRefused + " refused for good (see the report)");
            if (JobPending > 0) parts.Add(JobPending + " pending");
            return name + " incomplete: " + JobDone + " of " + JobTotal + " done, " + string.Join(", ", parts);
        }
    }

    /// <summary>A stream source that hashes what is read (the last complete read wins).</summary>
    sealed class HashingSource : IFxStreamSource
    {
        readonly IFxStreamSource _inner;
        public string Sha256;
        public HashingSource(IFxStreamSource inner) { _inner = inner; }
        public long Length { get { return _inner.Length; } }
        public Stream OpenRead() { return new HashStream(_inner.OpenRead(), this); }

        sealed class HashStream : Stream
        {
            readonly Stream _s; readonly HashingSource _owner; readonly IncrementalHash _h = IncrementalHash.CreateHash(HashAlgorithmName.SHA256); long _n;
            public HashStream(Stream s, HashingSource owner) { _s = s; _owner = owner; }
            public override bool CanRead { get { return true; } }
            public override bool CanSeek { get { return false; } }
            public override bool CanWrite { get { return false; } }
            public override long Length { get { return _s.Length; } }
            public override long Position { get { return _n; } set { throw new NotSupportedException(); } }
            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }
            public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
            public override int Read(byte[] buffer, int offset, int count)
            {
                int n = _s.Read(buffer, offset, count);
                if (n > 0) { _h.AppendData(buffer, offset, n); _n += n; }
                else if (_n == _owner.Length) _owner.Sha256 = BitConverter.ToString(_h.GetHashAndReset()).Replace("-", "").ToLowerInvariant();
                return n;
            }
            protected override void Dispose(bool disposing) { if (disposing) { _s.Dispose(); _h.Dispose(); } base.Dispose(disposing); }
        }
    }

    /// <summary>The identifying properties of a message, to find it again among the items of a folder.</summary>
    public static class PstImportSignature
    {
        static readonly ushort[] Ids = { 0x001A, 0x0037, 0x0039, 0x0E06, 0x1035, 0x3001 };

        /// <summary>One line per identifying property the message has (text of a string, hex of the rest).</summary>
        public static Dictionary<ushort, string> Of(FxMessage m)
        {
            var d = new Dictionary<ushort, string>();
            foreach (ushort id in Ids)
            {
                FxProperty p = m.Find(id);
                if (p == null || p.IsLazy || p.IsMultiValued || p.RawLength > 4096) continue;
                d[id] = p.TypeInfo.IsString ? p.AsString().TrimEnd('\0') : Convert.ToBase64String(p.Raw);
            }
            return d;
        }

        /// <summary>The candidate has every identifying property of the message with the same value, and the message has at least one.</summary>
        public static bool Matches(Dictionary<ushort, string> ours, Dictionary<ushort, string> theirs)
        {
            if (ours.Count == 0) return false;
            foreach (var kv in ours)
            {
                string v;
                if (!theirs.TryGetValue(kv.Key, out v) || v != kv.Value) return false;
            }
            return true;
        }
    }

    public sealed class PstImportJob : TransferJob
    {
        readonly MailboxApi _api;
        readonly PstImportJournal _journal;
        readonly PstImportOptions _o;
        readonly PstImportRunResult _result = new PstImportRunResult();
        readonly ImportSessionPool _sessions;
        readonly Dictionary<string, PstImportSource> _sources = new Dictionary<string, PstImportSource>();
        readonly Dictionary<string, PstReader> _readers = new Dictionary<string, PstReader>();
        readonly object _readerLock = new object();
        readonly string _workDir;
        readonly IProgress<PstProgress> _progress;
        long _unknownThisRun, _failedThisRun, _refusedThisRun;
        long _reconciledCreated, _reimported, _ambiguous, _exhausted;
        readonly ConcurrentDictionary<long, DateTime> _toReconcile = new ConcurrentDictionary<long, DateTime>();
        readonly ConcurrentDictionary<long, byte> _producedThisRun = new ConcurrentDictionary<long, byte>();
        readonly ConcurrentDictionary<long, int> _rounds = new ConcurrentDictionary<long, int>();
        readonly Gate _gate = new Gate();
        /// <summary>How many not-imported rows the result lists (the report of the refused items is written from them).</summary>
        public int MaxReports = 100000;

        public PstImportRunResult Result { get { return _result; } }
        TimeSpan ReconcileDelay { get { return TimeSpan.FromSeconds(Math.Max(0, _o.ReconcileDelaySeconds)); } }

        sealed class Gate
        {
            readonly object _l = new object();
            int _active; bool _exclusive;
            public void Enter(CancellationToken ct) { lock (_l) { while (_exclusive) { ct.ThrowIfCancellationRequested(); Monitor.Wait(_l, 250); } _active++; } }
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

        PstImportJob(MailboxApi api, PstImportJournal journal, PstImportOptions options, string workDir, IProgress<PstProgress> progress)
        {
            _api = api; _journal = journal; _o = options; _workDir = workDir; _progress = progress;
            Graph = api.Graph;
            _sessions = new ImportSessionPool(api);
            InitLimiter(options.Concurrency);
        }

        public static PstImportJob Start(MailboxApi api, PstImportJournal journal, PstImportOptions options, string workDir = null, IProgress<PstProgress> progress = null)
        {
            if (string.IsNullOrEmpty(options.TargetMailboxId)) throw new ArgumentException("TargetMailboxId is required.");
            if (options.PstPaths == null || options.PstPaths.Count == 0) throw new ArgumentException("PstPaths is required.");
            if (string.IsNullOrWhiteSpace(options.TargetFolder) && !options.MergeWellKnown) throw new ArgumentException("TargetFolder is required (the folder below which the folders of the PST are created).");
            TransferLimits.CheckConcurrency(options.Concurrency);
            var job = new PstImportJob(api, journal, options, workDir, progress);
            job.Begin(job.Run);
            return job;
        }

        void Problem(string text) { lock (_result.Problems) _result.Problems.Add(text); Log(text); }

        /// <summary>What decides which folder of the target an item goes to, written the same way every time: the target folder, the well-known folders, the file subfolders, the map.</summary>
        public static string TargetOptions(PstImportOptions o)
        {
            string map = o.FolderMap == null ? "" : string.Join("|", o.FolderMap.OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase).Select(kv => kv.Key.Trim('/').ToLowerInvariant() + ">" + (kv.Value ?? "").Trim('/')));
            return "targetFolder=" + (o.TargetFolder ?? "").Trim('/') + "; mergeWellKnown=" + o.MergeWellKnown + "; perFileSubfolder=" + o.PerFileSubfolder + "; folderMap=" + map;
        }
        static string Short(string s) { return s != null && s.Length > 300 ? s.Substring(0, 300) + "..." : s; }
        static string Seconds(TimeSpan t) { return Math.Round(t.TotalSeconds, 1).ToString("0.#", CultureInfo.InvariantCulture) + " s"; }

        void Run()
        {
            try { RunCore(); }
            finally
            {
                FillJobTotals();
                lock (_readerLock) { foreach (PstReader r in _readers.Values) r.Dispose(); _readers.Clear(); }
            }
        }

        void FillJobTotals()
        {
            try
            {
                if (_result.JobId == 0) return;
                long done = 0, failed = 0, unknown = 0, pending = 0, refused = 0;
                foreach (PstJournalStat st in _journal.Stats(_result.JobId))
                    switch (st.State)
                    {
                        case "Done": done = st.Count; break;
                        case "Failed": failed = st.Count; break;
                        case "Unknown": unknown = st.Count; break;
                        case "Refused": refused = st.Count; break;
                        default: pending += st.Count; break;
                    }
                _result.JobDone = done; _result.JobFailed = failed; _result.JobUnknown = unknown; _result.JobPending = pending; _result.JobRefused = refused;
                _result.JobTotal = done + failed + unknown + pending + refused;
                _result.JobUnknownEarlier = Math.Max(0, unknown - Interlocked.Read(ref _unknownThisRun));
                _result.JobKnown = true;
                foreach (string state in new[] { "Refused", "Failed", "Unknown" })
                    foreach (PstJournalItem it in _journal.GetItems(_result.JobId, state, MaxReports))
                    {
                        if (_result.NotImported.Count >= MaxReports) break;
                        PstImportSource s;
                        _result.NotImported.Add(new PstImportItemReport
                        {
                            Pst = _sources.TryGetValue(it.PstId, out s) ? s.SourcePath : it.PstId, Folder = it.PstPath, Subject = it.Subject, MessageClass = it.MessageClass, State = it.State,
                            Reason = it.Error, Nid = it.Nid, Size = it.Size, Attempts = it.Attempts
                        });
                    }
            }
            catch (Exception ex) { Log("the state of the job could not be read: " + ex.Message); }
        }

        void RunCore()
        {
            CancellationToken ct = Cts.Token;
            string target = _o.TargetMailboxId;
            _result.JobName = _o.JobName;

            // the job of the journal (a name that exists is resumed)
            long jobId = _journal.FindJob(_o.JobName);
            bool resumed = jobId != 0;
            if (!resumed) jobId = _journal.CreateJob(_o.JobName, target, _o.TargetFolder);
            else if (!string.Equals(_journal.JobTarget(jobId), target, StringComparison.Ordinal))
                throw new InvalidOperationException("The import job '" + _o.JobName + "' of the journal targets another mailbox (" + _journal.JobTarget(jobId) + "). Give another -JobName.");
            // the same job name with other target options would send the rest of the items to other folders than the first ones
            string options = TargetOptions(_o);
            string was = _journal.JobOptions(jobId);
            if (was == null) _journal.SetJobOptions(jobId, options);
            else if (!string.Equals(was, options, StringComparison.Ordinal))
                throw new InvalidOperationException("The import job '" + _o.JobName + "' of the journal was started with other target options (" + was + "), these are " + options +
                    ". Resuming it would send the rest of the items to other folders than the first ones. Give the same options, or another -JobName.");
            _result.JobId = jobId;
            _result.RecoveredToUnknown = _journal.RecoveredToUnknown;
            if (resumed && _journal.RecoveredToUnknown > 0)
                Log("recovery: " + _journal.RecoveredToUnknown + " row(s) were InFlight when an earlier run stopped and are now Unknown (the items may exist)");

            Metrics.Phase = "check";
            long jid = jobId;
            List<PstImportSource> sources = PstImportPrepare.Run(_o, id => _journal.GetPst(jid, id), _workDir, false, _progress, ct);
            _result.Sources = sources;
            foreach (PstImportSource s in sources)
            {
                string earlier = resumed ? _journal.OtherPstIdForPath(jobId, s.SourcePath, s.PstId) : null;
                if (earlier != null)
                    Log("WARNING: " + s.SourcePath + " is not the file this job imported before (" + earlier + ", now " + s.PstId + "): it was modified or replaced, so its items are treated as new and may be imported a second time. " +
                        "Keep the PST files unchanged (do not open them in Outlook) during a job, or use another -JobName.");
                _journal.SavePst(jobId, s);
                _sources[s.PstId] = s;
                Log("file " + s.SourcePath + ": " + (s.Checked ? s.Note : "not checked"));
            }
            PstImportSource damaged = sources.FirstOrDefault(s => (s.Damaged && !s.Repaired) || (_o.Repair && s.Damaged && s.WorkPath == null));
            if (damaged != null)
                throw new InvalidOperationException("The PST file " + damaged.SourcePath + " is damaged (" + damaged.Note + "). Nothing was imported. " +
                    (_o.Repair ? "The repair did not produce a usable copy." : "Run again with -Repair to import from a repaired copy (the file itself is never modified)."));

            Metrics.Phase = "folders";
            List<FolderInfo> tree = _api.GetFolderTree(target, ct);
            var plans = new List<PstFolderPlan>();
            FolderFilter filter = FolderFilter.Create(_o.IncludeFolders, _o.ExcludeFolders);
            if (filter != null) Log("folders chosen: " + filter);
            foreach (PstImportSource s in sources)
            {
                string note, store;
                using (PstReader r = PstReader.Open(s.OpenPath)) store = r.DisplayName;
                plans.AddRange(PstImportPlanner.MapFolders(_o, s, tree, store, filter, out note));
                if (note != null) Log(note);
            }
            if (filter != null)
            {
                foreach (string u in filter.Unmatched()) Problem(u + " matches no folder of the PST file(s)");
                if (plans.Count == 0) throw new InvalidOperationException("No folder of the PST file(s) is chosen (" + filter + "): nothing was imported. Check the folder paths (they are relative to the top folder of the PST, for instance Inbox/Projects).");
            }
            var needed = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var existing = new HashSet<string>(tree.Select(f => f.Path), StringComparer.OrdinalIgnoreCase);
            var types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (PstFolderPlan fp in plans)
            {
                if (!types.ContainsKey(fp.TargetPath) && !string.IsNullOrEmpty(fp.ContainerClass)) types[fp.TargetPath] = fp.ContainerClass;
                var chain = new List<string>();
                for (string p = fp.TargetPath; p.Length > 0; p = p.Contains('/') ? p.Substring(0, p.LastIndexOf('/')) : "") chain.Add(p);
                chain.Reverse();
                foreach (string p in chain) if (!existing.Contains(p) && seen.Add(p)) needed.Add(p);
            }
            needed = needed.OrderBy(p => p.Count(c => c == '/')).ThenBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
            var folders = new FolderEnsurer(_api, target, tree);
            folders.EnsureAll(needed, p => { string t; return types.TryGetValue(p, out t) ? t : null; }, Log, ct, Limiter);
            _result.FoldersCreated = folders.Created; _result.FoldersFailed = folders.Failed;
            lock (_result.Problems) _result.Problems.AddRange(folders.Problems);

            Metrics.Phase = "plan";
            long planned = 0;
            foreach (PstImportSource s in sources)
            {
                using (PstReader r = PstReader.Open(s.OpenPath))
                {
                    var byNid = r.AllFolders().ToDictionary(f => f.Nid);
                    foreach (PstFolderPlan fp in plans.Where(p => p.PstId == s.PstId))
                    {
                        ct.ThrowIfCancellationRequested();
                        FolderInfo tf = folders.Find(fp.TargetPath);
                        PstReaderFolder rf = byNid[fp.Nid];
                        _result.AssociatedSkipped += fp.Associated;
                        // a folder that could not be created: its items are rows of the job, Failed (nothing was sent), so that the job is not complete and the next run, which tries to create
                        // the folder again, imports them
                        if (tf == null)
                        {
                            if (rf.ContentCount > 0) Problem("the " + rf.ContentCount + " items of '" + fp.PstPath + "' (" + s.SourcePath + ") are not imported: the folder '" + fp.TargetPath + "' does not exist in the target and could not be created");
                        }
                        else _journal.AttachFolder(jobId, s.PstId, fp.TargetPath, target, tf.Id);
                        var batch = new List<PstJournalItem>();
                        foreach (PstScanItem it in PstImportScanner.Scan(r, s, rf, ct))
                        {
                            batch.Add(new PstJournalItem
                            {
                                PstId = s.PstId, Nid = it.Nid, Fingerprint = it.Fingerprint, PstPath = fp.PstPath, TargetPath = fp.TargetPath, TargetMailbox = target, TargetFolder = tf == null ? null : tf.Id,
                                Subject = it.Subject, MessageClass = it.Class, Size = it.Size,
                                State = it.Error != null ? "Refused" : tf == null ? "Failed" : "Pending",
                                Error = it.Error ?? (tf == null ? "the folder '" + fp.TargetPath + "' could not be created in the target: nothing was sent, the next run tries again" : null)
                            });
                            if (batch.Count >= 500) { planned += AddBatch(jobId, batch); batch.Clear(); }
                        }
                        planned += AddBatch(jobId, batch);
                        if (_progress != null) _progress.Report(new PstProgress { Phase = "plan", Items = planned, CurrentFolder = fp.PstPath });
                    }
                }
            }
            _result.Planned = true;
            int requeued = _journal.RequeueFailed(jobId, _o.MaxAttempts);
            long pendingCount = _journal.Count(jobId, "Pending");
            PstJournalStat pendingStat = _journal.Stats(jobId).FirstOrDefault(x => x.State == "Pending");
            Log("plan: " + planned + " new item(s) added to the job '" + _o.JobName + "' (" + requeued + " failed ones tried again), " + pendingCount + " to import");
            if (_result.AssociatedSkipped > 0) Log(_result.AssociatedSkipped + " folder-associated item(s) (rules, views, settings) are not imported");
            if (resumed)
                Log("job '" + _o.JobName + "' resumed: " + _journal.Count(jobId, "Done") + " done, " + _journal.Count(jobId, "Failed") + " failed, " + _journal.Count(jobId, "Unknown") + " unknown, " + _journal.Count(jobId, "Refused") + " refused");

            _result.AutoReconcile = _o.AutoReconcile;
            if (_o.AutoReconcile)
            {
                List<PstJournalItem> left = _journal.GetItems(jobId, "Unknown");
                foreach (PstJournalItem u in left) _toReconcile[u.ItemKey] = u.UpdatedUtc ?? DateTime.UtcNow;
                if (left.Count > 0) Log("reconcile: " + left.Count + " Unknown row(s) left by an earlier run will be looked up in their target folder once " + Seconds(ReconcileDelay) + " have passed since their request ended");
            }

            Metrics.Phase = "import";
            Metrics.BeginTransfer(pendingCount, pendingStat == null ? 0 : pendingStat.Bytes);
            foreach (PstJournalFolderCount c in _journal.CountByTarget(jobId, "Pending")) Metrics.Folders.Expect(c.TargetPath, c.Count);
            RunWorkers(_o.Concurrency, n => ImportLoop(jobId, ct));
            if (_o.AutoReconcile) ReconcileTail(jobId, ct);

            _result.Done = Metrics.Items;
            _result.Failed = Interlocked.Read(ref _failedThisRun);
            _result.Unknown = Interlocked.Read(ref _unknownThisRun);
            _result.Refused = Interlocked.Read(ref _refusedThisRun);
            _result.Bytes = Metrics.Bytes;
            _result.ReconciledCreated = Interlocked.Read(ref _reconciledCreated);
            _result.ReconciledReimported = Interlocked.Read(ref _reimported);
            _result.ReconcileAmbiguous = Interlocked.Read(ref _ambiguous);
            _result.ReconcileExhausted = Interlocked.Read(ref _exhausted);
            _result.SessionsCreated = _sessions.Created;
            foreach (TimeSpan t in _sessions.Lifetimes) _result.SessionLifetimeSeconds.Add(t.TotalSeconds);
            _result.Completed = !ct.IsCancellationRequested;
            Metrics.Phase = "done";
            _result.Final = Snapshot();
            string summary = _result.ReconcileSummary();
            if (summary != null) Log(summary);
        }

        long AddBatch(long jobId, List<PstJournalItem> batch)
        {
            List<PstJournalItem> refused;
            int added = _journal.AddItems(jobId, batch, out refused);
            foreach (PstJournalItem it in refused)
            {
                Interlocked.Increment(ref _refusedThisRun);
                Log("item " + (it.PstPath ?? "") + ": " + (it.Subject ?? "(no subject)") + " (node " + it.Nid + ") refused: " + Short(it.Error));
            }
            return added;
        }

        // ------------------------------------------------------------------------------------------ reading the PST
        PstReader ReaderFor(string pstId)
        {
            lock (_readerLock)
            {
                PstReader r;
                if (!_readers.TryGetValue(pstId, out r)) { r = PstReader.Open(_sources[pstId].OpenPath); _readers[pstId] = r; }
                return r;
            }
        }

        /// <summary>The message as it is sent (made ready for importItem). The reader is shared: the call is serialized, the big values stay lazy.</summary>
        FxMessage ReadForImport(PstJournalItem it, bool forExchange)
        {
            PstReader r = ReaderFor(it.PstId);
            lock (r)
            {
                FxMessage m = r.ReadMessage(it.Nid);
                return forExchange ? FxImportStream.Prepare(m) : m;
            }
        }

        // ------------------------------------------------------------------------------------------ import
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
                        PstJournalItem next = _journal.NextPending(jobId);
                        if (next == null) return;
                        ImportOne(next, ct);
                    }
                }
                finally { _gate.Exit(); }
            }
        }

        static bool IsContentRefusal(GraphException ge)
        {
            return ge != null && (ge.StatusCode == 400 || ge.StatusCode == 413 || ge.StatusCode == 415 || ge.StatusCode == 422);
        }

        void Refuse(PstJournalItem it, string reason)
        {
            _journal.MarkRefused(it.ItemKey, reason);
            Interlocked.Increment(ref _refusedThisRun);
            Metrics.ItemFailed();
            Metrics.Folders.Fail(it.TargetPath);
            Log("item " + it.ItemKey + " (" + (it.PstPath ?? "") + ": " + (it.Subject ?? "(no subject)") + ") refused: " + Short(reason));
        }

        /// <summary>The attempts allowed are used and nothing was ever created: the item is refused for good (the reason keeps the last error).</summary>
        void GiveUp(PstJournalItem it, string lastError)
        {
            _journal.MarkRefused(it.ItemKey, "gave up after " + it.Attempts + " attempt(s), nothing was created. Last error: " + lastError);
            Interlocked.Increment(ref _refusedThisRun);
            Metrics.Folders.Fail(it.TargetPath);
            Log("item " + it.ItemKey + " (" + (it.PstPath ?? "") + ": " + (it.Subject ?? "(no subject)") + ") refused for good: gave up after " + it.Attempts + " attempt(s). Last error: " + Short(lastError));
        }

        void ImportOne(PstJournalItem it, CancellationToken ct)
        {
            if (ct.IsCancellationRequested) { _journal.MarkPending(it.ItemKey); return; }
            HashingSource src;
            try
            {
                FxMessage m = ReadForImport(it, true);
                src = new HashingSource(FxImportStream.Build(m));   // validates every property and checks the stream before anything is sent
            }
            catch (OperationCanceledException) { _journal.MarkPending(it.ItemKey); return; }
            catch (FxImportCheckException ex)
            {
                Refuse(it, "not sent, the stream built does not match the message: " + ex.Message);
                return;
            }
            catch (Exception ex)
            {
                Refuse(it, (ex is ArgumentException ? "cannot be expressed as a FastTransfer stream: " : "cannot be read from the PST: ") + ex.Message);
                return;
            }
            try
            {
                ImportResult r = ImportHelper.ImportWithSession(_api, _sessions, src.OpenRead, it.TargetMailbox, it.TargetFolder, src.Length, ct, Log, null, null);
                _journal.MarkDone(it.ItemKey, r.ItemId, src.Sha256);
                Metrics.ItemDone(r.Bytes);
                Metrics.Folders.Done(it.TargetPath, r.Bytes);
                Succeeded();
            }
            catch (OperationCanceledException)
            {
                _journal.MarkUnknown(it.ItemKey, "interrupted while the request was in flight: the item may exist");
                UnknownOnce(it.ItemKey, ct);
                Metrics.ItemFailed();
                Metrics.Folders.Unknown(it.TargetPath);
            }
            catch (Exception ex)
            {
                var ge = ex as GraphException;
                if (ge != null && (ge.StatusCode == 401 || ge.StatusCode == 403)) { _journal.MarkFailed(it.ItemKey, ex.Message); Metrics.ItemFailed(); Metrics.Folders.Fail(it.TargetPath); Fatal(ex); return; }
                if (IsContentRefusal(ge)) { Refuse(it, "refused by the service: " + (ge.ErrorCode ?? ("HTTP " + ge.StatusCode)) + " " + ge.Message); return; }
                Metrics.ItemFailed();
                if (ImportHelper.NothingCreated(ex))
                {
                    if (it.Attempts >= _o.MaxAttempts) { GiveUp(it, ex.Message); FailedOnce(ex.Message); return; }
                    _journal.MarkFailed(it.ItemKey, ex.Message);
                    Interlocked.Increment(ref _failedThisRun);
                    Metrics.Folders.Fail(it.TargetPath);
                    Log("item " + it.ItemKey + " failed (nothing created): " + Short(ex.Message));
                }
                else
                {
                    _journal.MarkUnknown(it.ItemKey, ex.Message);
                    UnknownOnce(it.ItemKey, ct);
                    Metrics.Folders.Unknown(it.TargetPath);
                    Log("item " + it.ItemKey + " UNKNOWN (the item may have been created): " + Short(ex.Message));
                }
                FailedOnce(ex.Message);
            }
        }

        void UnknownOnce(long key, CancellationToken ct)
        {
            Interlocked.Increment(ref _unknownThisRun);
            _producedThisRun[key] = 1;
            if (_o.AutoReconcile && !ct.IsCancellationRequested) _toReconcile[key] = DateTime.UtcNow;
        }

        void LeftUnknown(long key)
        {
            byte b;
            if (_producedThisRun.TryRemove(key, out b)) Interlocked.Decrement(ref _unknownThisRun);
        }

        // ------------------------------------------------------------------------------------------ reconciliation
        bool ReconcileDue()
        {
            if (_toReconcile.IsEmpty) return false;
            DateTime now = DateTime.UtcNow;
            foreach (KeyValuePair<long, DateTime> kv in _toReconcile) if (now - kv.Value >= ReconcileDelay) return true;
            return false;
        }

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
                catch (OperationCanceledException) { break; }
                if (ct.IsCancellationRequested) break;
                if (_journal.Count(jobId, "Pending") > 0)
                {
                    Metrics.Phase = "import";
                    RunWorkers(_o.Concurrency, n => ImportLoop(jobId, ct));
                }
            }
        }

        void ReconcilePass(long jobId, CancellationToken ct)
        {
            _gate.BeginExclusive(ct);
            try
            {
                DateTime now = DateTime.UtcNow;
                var due = new HashSet<long>(_toReconcile.Where(kv => now - kv.Value >= ReconcileDelay).Select(kv => kv.Key));
                if (due.Count == 0) return;
                List<PstJournalItem> unknown = _journal.GetItems(jobId, "Unknown");
                var present = new HashSet<long>(unknown.Select(u => u.ItemKey));
                foreach (long k in due.ToList()) if (!present.Contains(k)) { due.Remove(k); DateTime gone; _toReconcile.TryRemove(k, out gone); }
                if (due.Count == 0) return;
                HashSet<string> known = _journal.DoneTargetItems(jobId);
                int created = 0, absent = 0, ambiguous = 0, rounds = Math.Max(1, _o.ReconcileRounds);
                foreach (var group in unknown.GroupBy(u => u.TargetMailbox + "\n" + u.TargetFolder))
                {
                    List<PstJournalItem> rows = group.ToList();
                    var dueRows = new List<PstJournalItem>();
                    foreach (PstJournalItem u in rows.Where(r => due.Contains(r.ItemKey)))
                    {
                        int used; _rounds.TryGetValue(u.ItemKey, out used);
                        if (used >= rounds)
                        {
                            DateTime gone; _toReconcile.TryRemove(u.ItemKey, out gone);
                            Interlocked.Increment(ref _exhausted);
                            Problem("item " + u.ItemKey + " left Unknown: it was looked up " + used + " time(s) (-ReconcileRounds " + rounds + ") without being settled");
                        }
                        else dueRows.Add(u);
                    }
                    if (dueRows.Count == 0) continue;
                    List<ItemInfo> listing;
                    try { using (Limiter.Enter(ct)) listing = _api.EnumerateItems(rows[0].TargetMailbox, rows[0].TargetFolder, 0, ct).ToList(); }
                    catch (GraphException ex)
                    {
                        if (ex.StatusCode == 401 || ex.StatusCode == 403) { Fatal(ex); return; }
                        foreach (PstJournalItem u in dueRows) { _rounds.AddOrUpdate(u.ItemKey, 1, (k, v) => v + 1); _toReconcile[u.ItemKey] = DateTime.UtcNow; }
                        Problem("reconcile: the target folder " + rows[0].TargetFolder + " could not be listed (" + Short(ex.Message) + "), " + dueRows.Count + " Unknown row(s) will be looked up again after the delay");
                        continue;
                    }
                    // candidates of every Unknown row of the folder: the items created after the attempt began that nothing accounts for
                    var candidateIds = new Dictionary<long, List<string>>();
                    var all = new HashSet<string>(StringComparer.Ordinal);
                    foreach (PstJournalItem u in rows)
                    {
                        DateTime notBefore = ImportReconciler.NotBefore(u.AttemptStartUtc, u.UpdatedUtc);
                        var c = new List<string>();
                        foreach (ItemInfo i in listing)
                        {
                            if (known.Contains(i.Id)) continue;
                            DateTimeOffset? latest = i.Modified.HasValue && (!i.Created.HasValue || i.Modified.Value > i.Created.Value) ? i.Modified : i.Created;
                            if (latest.HasValue && latest.Value.UtcDateTime < notBefore) continue;
                            c.Add(i.Id); all.Add(i.Id);
                        }
                        candidateIds[u.ItemKey] = c;
                    }
                    // the service gives an item another size than the stream that was sent: each candidate is exported and compared with the message
                    Dictionary<string, Dictionary<ushort, string>> exported = ExportCandidates(rows[0].TargetMailbox, all.ToList(), ct);
                    var candidates = new Dictionary<long, List<string>>();
                    var claims = new Dictionary<string, int>(StringComparer.Ordinal);
                    foreach (PstJournalItem u in rows)
                    {
                        var mine = new List<string>();
                        if (candidateIds[u.ItemKey].Count > 0)
                        {
                            Dictionary<ushort, string> sig = null;
                            try { sig = PstImportSignature.Of(ReadForImport(u, false)); } catch (OperationCanceledException) { throw; } catch (Exception) { }
                            foreach (string id in candidateIds[u.ItemKey])
                            {
                                Dictionary<ushort, string> got;
                                if (sig != null && exported.TryGetValue(id, out got) && PstImportSignature.Matches(sig, got)) mine.Add(id);
                            }
                        }
                        candidates[u.ItemKey] = mine;
                        foreach (string id in mine) { int n; claims.TryGetValue(id, out n); claims[id] = n + 1; }
                    }
                    foreach (PstJournalItem u in dueRows)
                    {
                        DateTime gone; _toReconcile.TryRemove(u.ItemKey, out gone);
                        List<string> c = candidates[u.ItemKey];
                        if (c.Count == 1 && claims[c[0]] == 1)
                        {
                            _journal.ResolveUnknown(u.ItemKey, true, c[0]);
                            Metrics.Folders.Settled(u.TargetPath, true, u.Size);
                            Metrics.ItemDone(u.Size);
                            Interlocked.Increment(ref Metrics.TotalItems);
                            Interlocked.Increment(ref _reconciledCreated);
                            LeftUnknown(u.ItemKey);
                            created++;
                            Log("item " + u.ItemKey + " reconciled: the item was created, id " + c[0]);
                        }
                        else if (c.Count == 0)
                        {
                            _rounds.AddOrUpdate(u.ItemKey, 1, (k, v) => v + 1);
                            if (u.Attempts >= _o.MaxAttempts)
                            {
                                _journal.MarkRefused(u.ItemKey, "no matching item in the target folder after the lookup: nothing was created, and the " + u.Attempts + " attempts allowed are used. Last error: " + u.Error);
                                Metrics.Folders.SettledFailed(u.TargetPath);
                                Interlocked.Increment(ref _refusedThisRun);
                                Log("item " + u.ItemKey + " reconciled: no item was created, but its " + u.Attempts + " attempts are used (refused for good, nothing created)");
                            }
                            else
                            {
                                _journal.ResolveUnknown(u.ItemKey, false, null);
                                Metrics.Folders.Settled(u.TargetPath, false);
                                Interlocked.Increment(ref Metrics.TotalItems);
                                Interlocked.Increment(ref _reimported);
                                Log("item " + u.ItemKey + " reconciled: no item was created in the target folder, it is imported again (attempt " + (u.Attempts + 1) + " of " + _o.MaxAttempts + ")");
                            }
                            LeftUnknown(u.ItemKey);
                            absent++;
                        }
                        else
                        {
                            Interlocked.Increment(ref _ambiguous);
                            ambiguous++;
                            Problem("item " + u.ItemKey + " left Unknown: " + (c.Count == 1 ? "the one item that could be it (" + c[0] + ") could also be the item of another Unknown row" : c.Count + " items in the target folder could be it") + "; check the target folder");
                        }
                    }
                }
                Log("reconcile: " + (created + absent + ambiguous) + " row(s) looked up: " + created + " created, " + absent + " not created, " + ambiguous + " ambiguous");
            }
            finally { _gate.EndExclusive(); }
        }

        /// <summary>The identifying properties of the candidates (exported again, at most 20 per request). An item that cannot be exported is simply not a match.</summary>
        Dictionary<string, Dictionary<ushort, string>> ExportCandidates(string mailbox, List<string> ids, CancellationToken ct)
        {
            var result = new Dictionary<string, Dictionary<ushort, string>>(StringComparer.Ordinal);
            if (ids.Count == 0) return result;
            string dir = Path.Combine(Path.GetTempPath(), "mbt-pstimport-reconcile-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(dir);
            try
            {
                for (int i = 0; i < ids.Count; i += MailboxApi.MaxExportBatch)
                {
                    List<string> chunk = ids.Skip(i).Take(MailboxApi.MaxExportBatch).ToList();
                    try
                    {
                        List<ExportEntry> entries;
                        using (Limiter.Enter(ct)) entries = _api.ExportItems(mailbox, chunk, dir, i, ct);
                        foreach (ExportEntry e in entries)
                        {
                            if (e.HasError || e.FilePath == null) continue;
                            FxDocument d = FxCodec.Decode(FxStreamSource.FromFile(e.FilePath));
                            result[e.ItemId] = PstImportSignature.Of(d.Message);
                        }
                    }
                    catch (GraphException ex)
                    {
                        if (ex.StatusCode == 401 || ex.StatusCode == 403) throw;
                        Log("reconcile: " + chunk.Count + " candidate(s) could not be exported (" + Short(ex.Message) + ")");
                    }
                }
            }
            finally { try { Directory.Delete(dir, true); } catch (Exception) { } }
            return result;
        }
    }
}
