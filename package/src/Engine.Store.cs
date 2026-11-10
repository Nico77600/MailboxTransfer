// Engine.Store.cs - local storage of a mailbox export (catalog + packs).
//
// One export = one folder:
//   <export>\catalog.db            SQLite (WAL, synchronous=NORMAL): metadata, index, previews, import jobs.
//   <export>\packs\pack-NNNNNN.bin append-only archives (target ~1 GiB, then rotation). Each item is one contiguous
//                                  block (packId, offset, length, sha256 in the catalog). NEVER one file per item.
//   <export>\export.lock           exclusive lock held while the export is open for writing (one process only).
//
// Concurrency model
//   * Writes of item streams: each concurrent WriteItem takes an exclusive PACK WRITER from a pool (stack of idle writers).
//     When the pool is empty a pack is created (or an unsealed pack left by a previous session is reused), so there is
//     at most one open pack per concurrent writer (= per export worker): no lock is held while bytes are copied,
//     the workers never contend on a file. A writer goes back to the pool after each item, or the pack is sealed once it
//     reaches PackTargetSize (an item is never split, a pack can exceed the target by one item).
//   * Catalog: ONE write connection, serialised by a lock (statements are short; bulk upserts are chunked in transactions
//     of UpsertBatchSize rows so that item commits of the workers are never starved). Reads use a small pool of
//     read-only (query_only) connections: WAL lets them run in parallel with the writer.
//   * The commits that move an import row (InFlight, Done, ...) run with synchronous=FULL (StoreOptions.ImportSyncFull): a lost one could send an item twice.
//   * Durability order of an item: bytes written -> FileStream.Flush(true) (FlushFileBuffers) -> catalog row committed.
//     The item is "Exported" only after the last step. In WAL/synchronous=NORMAL a commit is not fsync'ed, but the data
//     is already on disk, and a lost commit only turns the item back into "to do".
//   * Recovery at Open: every pack is truncated to the end of its last committed item, items whose bytes are missing
//     are put back to Pending, and importItem rows InFlight become Unknown (never replayed automatically).
//   * No stream is ever fully loaded in memory: copy loop with a per-writer buffer (CopyBufferSize), incremental SHA-256,
//     reads through a bounded read-only sub-stream over the pack file.
//
// All public type names are prefixed Store* (ExportStore for the main class) to stay clear of the other Engine.*.cs files.

using System;
using System.Buffers;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Data.Sqlite;

namespace MailboxTransfer
{
    #region Enums, options and records

    /// <summary>Primary mailbox, main (In-Place) archive, or an auxiliary archive of an auto-expanding archive.</summary>
    public enum StoreMailboxKind { Primary = 0, Archive = 1, AuxArchive = 2 }

    public enum StoreItemState { Pending = 0, Exported = 1, Failed = 2, Skipped = 3 }

    public enum StoreImportState { Pending = 0, InFlight = 1, Done = 2, Failed = 3, Unknown = 4 }

    public enum StoreItemSort { ReceivedDate = 0, Size = 1, Subject = 2 }

    public sealed class StoreOptions
    {
        /// <summary>Target size of a pack before rotation (default 1 GiB). An item is never split.</summary>
        public long PackTargetSize = 1L << 30;
        /// <summary>Size of the copy buffer of each pack writer.</summary>
        public int CopyBufferSize = 256 * 1024;
        /// <summary>Flush(true) (FlushFileBuffers) after every item. Keep true except for benchmarks.</summary>
        public bool DurableFlush = true;
        /// <summary>Read connections kept open for reuse.</summary>
        public int MaxIdleReaders = 8;
        public int BusyTimeoutMs = 30000;
        /// <summary>Rows per transaction in UpsertFolders / UpsertItems / SetPreviews / AddImportItems.</summary>
        public int UpsertBatchSize = 2000;
        /// <summary>
        /// synchronous=FULL for the commits that move an import row (InFlight, Done, Failed, Unknown, Pending): a lost commit after a power cut could
        /// send an item twice. The export side stays on NORMAL (its bytes are flushed before the row, a lost commit only turns the item back into "to do").
        /// </summary>
        public bool ImportSyncFull = true;
    }

    public sealed class StoreMailbox
    {
        public string Id { get; set; }
        public StoreMailboxKind Kind { get; set; }
        public string Upn { get; set; }
        public string DisplayName { get; set; }
        public int Key { get; set; }
    }

    /// <summary>
    /// A mailbox location of a user (Primary, MainArchive = Archive, AuxArchive) with how and when it was learned: from Graph, from an
    /// operator file (Exchange Online Get-MailboxLocation), or from the redirects of an auto-expanding archive (308, ErrorArchiveFolderMovedPermanently, 409).
    /// </summary>
    public sealed class StoreMailboxLocation
    {
        /// <summary>Given by the operator or by Graph: wins over a location guessed from a redirect.</summary>
        public const int RankGiven = 2;
        /// <summary>Learned from a redirect of the service.</summary>
        public const int RankLearned = 1;

        public string MailboxId { get; set; }
        public StoreMailboxKind Kind { get; set; }
        public string Upn { get; set; }
        /// <summary>How it was discovered: "Graph", "OperatorFile", "Parameter", "Redirect308", "ExportMoved", "Import409", "ExportTarget"...</summary>
        public string Source { get; set; }
        public string Detail { get; set; }
        public int Rank { get; set; } = RankLearned;
        public DateTime DiscoveredUtc { get; set; } = DateTime.UtcNow;
    }

    public sealed class StoreFolder
    {
        public string MailboxId { get; set; }
        public string FolderId { get; set; }
        public string ParentFolderId { get; set; }
        public string DisplayName { get; set; }
        /// <summary>Output only: full path computed by the store ("Inbox/Sub").</summary>
        public string Path { get; set; }
        /// <summary>Folder type / container class (IPF.Note, ...).</summary>
        public string FolderType { get; set; }
        public long? SourceItemCount { get; set; }
        public long? SourceSize { get; set; }
        /// <summary>Output only.</summary>
        public long Key { get; set; }
    }

    /// <summary>Item description at enumeration time (UpsertItems).</summary>
    public sealed class StoreItemEntry
    {
        public string MailboxId { get; set; }
        public string ItemId { get; set; }
        public string FolderId { get; set; }
        /// <summary>
        /// Mailbox whose folder record <see cref="FolderId"/> is. Null: <see cref="MailboxId"/>. Differs for the items of an expanded folder of an
        /// auto-expanding archive: the folder is recorded once, in the main archive tree, while the items live (MailboxId) in an auxiliary archive.
        /// </summary>
        public string FolderMailboxId { get; set; }
        public string ChangeKey { get; set; }
        public string ItemType { get; set; }
        public long Size { get; set; }
    }

    /// <summary>Items of a state in one folder (the progress per folder of a job).</summary>
    public sealed class StoreFolderCount
    {
        public long FolderKey, Count, Bytes;
    }

    public sealed class StoreItem
    {
        public long Key { get; set; }
        public string MailboxId { get; set; }
        public string ItemId { get; set; }
        public long? FolderKey { get; set; }
        public string FolderId { get; set; }
        public string ChangeKey { get; set; }
        public string ItemType { get; set; }
        public long Size { get; set; }
        public StoreItemState State { get; set; }
        public int? PackId { get; set; }
        public long? Offset { get; set; }
        public long? Length { get; set; }
        public string Sha256 { get; set; }
        public DateTime? ExportedUtc { get; set; }
        public int Attempts { get; set; }
        public string LastError { get; set; }
        public string Subject { get; set; }
        public DateTime? ReceivedUtc { get; set; }
    }

    public sealed class StoreItemPage
    {
        public List<StoreItem> Items { get; set; }
        public long Total { get; set; }
        public int Page { get; set; }
        public int PageSize { get; set; }
        public int PageCount { get { return PageSize <= 0 ? 0 : (int)((Total + PageSize - 1) / PageSize); } }
    }

    public sealed class StorePreview
    {
        public long ItemKey { get; set; }
        public string Subject { get; set; }
        public string From { get; set; }
        public string To { get; set; }
        public DateTime? ReceivedUtc { get; set; }
        public DateTime? SentUtc { get; set; }
        public string MessageClass { get; set; }
        public bool HasAttachments { get; set; }
        public int AttachmentCount { get; set; }
        public string AttachmentNames { get; set; }
        /// <summary>Truncated to 1 KiB (UTF-8) when stored.</summary>
        public string BodyPreview { get; set; }
    }

    public sealed class StoreWriteResult
    {
        public long ItemKey { get; set; }
        public long Length { get; set; }
        public string Sha256 { get; set; }
        public int PackId { get; set; }
        public long Offset { get; set; }
    }

    public sealed class StoreVerifyResult
    {
        public long ItemKey { get; set; }
        public bool Ok { get; set; }
        public string Reason { get; set; }
        public long Length { get; set; }
        public string ExpectedSha256 { get; set; }
        public string ActualSha256 { get; set; }
    }

    public sealed class StoreAggregate
    {
        public long Count { get; set; }
        /// <summary>Sum of the sizes announced by the source.</summary>
        public long Size { get; set; }
        public long Pending { get; set; }
        public long Exported { get; set; }
        public long Failed { get; set; }
        public long Skipped { get; set; }
        /// <summary>Bytes actually stored in the packs.</summary>
        public long StoredBytes { get; set; }

        internal void Add(StoreAggregate o)
        {
            Count += o.Count; Size += o.Size; Pending += o.Pending; Exported += o.Exported;
            Failed += o.Failed; Skipped += o.Skipped; StoredBytes += o.StoredBytes;
        }
    }

    public sealed class StoreFolderNode
    {
        public long FolderKey { get; set; }
        public string FolderId { get; set; }
        public string ParentFolderId { get; set; }
        public string DisplayName { get; set; }
        public string Path { get; set; }
        public string FolderType { get; set; }
        public long? SourceItemCount { get; set; }
        public long? SourceSize { get; set; }
        public int Depth { get; set; }
        /// <summary>Items directly in this folder.</summary>
        public StoreAggregate Own { get; set; }
        /// <summary>This folder and all its sub-folders.</summary>
        public StoreAggregate Total { get; set; }
        public List<StoreFolderNode> Children { get; set; }
    }

    public sealed class StoreStateStat
    {
        public string State { get; set; }
        public long Count { get; set; }
        public long AnnouncedBytes { get; set; }
        public long StoredBytes { get; set; }
    }

    public sealed class StoreStats
    {
        public long Mailboxes { get; set; }
        public long Folders { get; set; }
        public long Items { get; set; }
        public long Previews { get; set; }
        public long Packs { get; set; }
        public long PackBytesOnDisk { get; set; }
        public long AnnouncedBytes { get; set; }
        public long StoredBytes { get; set; }
        public StoreStateStat[] ItemsByState { get; set; }
        public long ImportJobs { get; set; }
        public StoreStateStat[] ImportByState { get; set; }
    }

    public sealed class StoreRecoveryReport
    {
        public int PacksChecked { get; set; }
        public int TruncatedPacks { get; set; }
        public long TruncatedBytes { get; set; }
        /// <summary>Packs shorter than what the catalog says (external damage): the missing items go back to Pending.</summary>
        public int DamagedPacks { get; set; }
        public long LostItems { get; set; }
        public int ImportInFlightToUnknown { get; set; }
        /// <summary>The same count by import job (job id -> rows that were InFlight and became Unknown).</summary>
        public Dictionary<long, int> ImportInFlightToUnknownByJob { get; } = new Dictionary<long, int>();
        public int UnknownPackFiles { get; set; }
    }

    public sealed class StoreImportPlan
    {
        public long ItemKey { get; set; }
        /// <summary>Null = the target mailbox of the job.</summary>
        public string TargetMailbox { get; set; }
        public string TargetFolder { get; set; }
    }

    public sealed class StoreImportItem
    {
        public long ImportKey { get; set; }
        public long JobId { get; set; }
        public long ItemKey { get; set; }
        public string ItemId { get; set; }
        public string SourceMailboxId { get; set; }
        public string TargetMailbox { get; set; }
        public string TargetFolder { get; set; }
        public StoreImportState State { get; set; }
        public string TargetItemId { get; set; }
        public int Attempts { get; set; }
        public string LastError { get; set; }
        public long? Length { get; set; }
        /// <summary>Last change of the row: when it became InFlight, then Done / Failed / Unknown (so, for an Unknown row, when the request ended).</summary>
        public DateTime? UpdatedUtc { get; set; }
    }

    public sealed class StoreImportJob
    {
        public long JobId { get; set; }
        public string Name { get; set; }
        public string TargetMailbox { get; set; }
        public DateTime CreatedUtc { get; set; }
    }

    #endregion

    #region Internal streams and pack writer

    /// <summary>Read-only, seekable sub-stream [offset, offset+length) of a pack file. Owns the underlying FileStream.</summary>
    internal sealed class StoreBoundedStream : Stream
    {
        private readonly FileStream _fs;
        private readonly long _offset;
        private readonly long _length;
        private long _pos;

        public StoreBoundedStream(FileStream fs, long offset, long length)
        {
            _fs = fs; _offset = offset; _length = length; _pos = 0;
            _fs.Position = offset;
        }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return true; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { return _length; } }
        public override long Position
        {
            get { return _pos; }
            set
            {
                if (value < 0) throw new ArgumentOutOfRangeException("value");
                _pos = value;
                _fs.Position = _offset + Math.Min(value, _length);
            }
        }

        public override int Read(byte[] buffer, int offset, int count)
        {
            return Read(new Span<byte>(buffer, offset, count));
        }

        public override int Read(Span<byte> buffer)
        {
            long remaining = _length - _pos;
            if (remaining <= 0 || buffer.Length == 0) return 0;
            if (buffer.Length > remaining) buffer = buffer.Slice(0, (int)remaining);
            int n = _fs.Read(buffer);
            _pos += n;
            return n;
        }

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default(CancellationToken))
        {
            long remaining = _length - _pos;
            if (remaining <= 0 || buffer.Length == 0) return 0;
            if (buffer.Length > remaining) buffer = buffer.Slice(0, (int)remaining);
            int n = await _fs.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            _pos += n;
            return n;
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            return ReadAsync(new Memory<byte>(buffer, offset, count), cancellationToken).AsTask();
        }

        public override long Seek(long offset, SeekOrigin origin)
        {
            long target;
            switch (origin)
            {
                case SeekOrigin.Begin: target = offset; break;
                case SeekOrigin.Current: target = _pos + offset; break;
                default: target = _length + offset; break;
            }
            if (target < 0) throw new IOException("Seek before the beginning of the item.");
            Position = target;
            return _pos;
        }

        public override void Flush() { }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _fs.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>Exclusive writer on one pack file. Owned by one WriteItem at a time (or idle in the pool).</summary>
    internal sealed class StorePackWriter
    {
        public int PackId;
        public FileStream Stream;
        /// <summary>End of the last committed item = where the next item starts.</summary>
        public long End;
        public byte[] Buffer;
        public bool Dirty;
        public bool Closed;
    }

    #endregion

    /// <summary>
    /// Local store of an export. Thread-safe: any number of threads may call WriteItem/WriteItemAsync and the read APIs.
    /// </summary>
    public sealed class ExportStore : IDisposable
    {
        public const int SchemaVersion = 1;
        public const int MaxPreviewBodyBytes = 1024;

        private readonly string _root;
        private readonly string _packsDir;
        private readonly string _dbPath;
        private readonly StoreOptions _opt;
        private readonly bool _readOnly;

        private SqliteConnection _w;                       // the single write connection
        private readonly object _dbLock = new object();    // serialises everything that uses _w
        private FileStream _lockFile;

        private readonly object _poolLock = new object();
        private readonly Stack<StorePackWriter> _idle = new Stack<StorePackWriter>();
        private readonly Queue<int> _reusable = new Queue<int>();       // unsealed packs of a previous session
        private readonly List<StorePackWriter> _allWriters = new List<StorePackWriter>();

        private readonly object _readerLock = new object();
        private readonly Stack<SqliteConnection> _readers = new Stack<SqliteConnection>();

        private readonly object _stateLock = new object();
        private int _active;
        private bool _closing;
        private bool _closed;

        private readonly ConcurrentDictionary<string, int> _mailboxKeys = new ConcurrentDictionary<string, int>(StringComparer.Ordinal);
        private readonly ConcurrentDictionary<string, long> _folderKeys = new ConcurrentDictionary<string, long>(StringComparer.Ordinal);

        public string Root { get { return _root; } }
        public string PacksDirectory { get { return _packsDir; } }
        public string DatabasePath { get { return _dbPath; } }
        public bool IsReadOnly { get { return _readOnly; } }
        public StoreOptions Options { get { return _opt; } }
        public string Name { get; private set; }
        public string Source { get; private set; }
        public DateTime CreatedUtc { get; private set; }
        public StoreRecoveryReport LastRecovery { get; private set; }

        #region Create / Open / Dispose

        public static ExportStore Create(string path, string exportName, string source)
        {
            return Create(path, exportName, source, null);
        }

        /// <summary>Creates a new export in the folder <paramref name="path"/> (must not already contain a catalog).</summary>
        public static ExportStore Create(string path, string exportName, string source, StoreOptions options)
        {
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("path");
            string full = System.IO.Path.GetFullPath(path);
            FileNaming.Require(path, false, "Store folder");
            if (File.Exists(System.IO.Path.Combine(full, "catalog.db"))) throw new IOException("An export already exists in " + full);
            Directory.CreateDirectory(full);
            Directory.CreateDirectory(System.IO.Path.Combine(full, "packs"));
            return new ExportStore(full, options ?? new StoreOptions(), false, true, exportName, source);
        }

        public static ExportStore Open(string path) { return Open(path, null); }

        /// <summary>Opens an existing export for reading and writing (runs the crash recovery).</summary>
        public static ExportStore Open(string path, StoreOptions options)
        {
            string full = System.IO.Path.GetFullPath(path);
            if (!File.Exists(System.IO.Path.Combine(full, "catalog.db"))) throw new FileNotFoundException("No export in " + full);
            Directory.CreateDirectory(System.IO.Path.Combine(full, "packs"));
            return new ExportStore(full, options ?? new StoreOptions(), false, false, null, null);
        }

        /// <summary>Opens an export read-only (no lock, no recovery, no writes): browsing while another process exports.</summary>
        public static ExportStore OpenReadOnly(string path)
        {
            string full = System.IO.Path.GetFullPath(path);
            if (!File.Exists(System.IO.Path.Combine(full, "catalog.db"))) throw new FileNotFoundException("No export in " + full);
            return new ExportStore(full, new StoreOptions(), true, false, null, null);
        }

        private ExportStore(string root, StoreOptions opt, bool readOnly, bool create, string name, string source)
        {
            _root = root;
            _packsDir = System.IO.Path.Combine(root, "packs");
            _dbPath = System.IO.Path.Combine(root, "catalog.db");
            _opt = opt;
            _readOnly = readOnly;
            if (opt.PackTargetSize < 1) throw new ArgumentException("PackTargetSize");
            if (opt.CopyBufferSize < 4096) opt.CopyBufferSize = 4096;
            try
            {
                if (!readOnly)
                {
                    try
                    {
                        _lockFile = new FileStream(System.IO.Path.Combine(root, "export.lock"), FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1);
                    }
                    catch (IOException ex)
                    {
                        throw new IOException("The export is already open by another process: " + root, ex);
                    }
                    _w = OpenConnection(false);
                    Exec(_w, "PRAGMA journal_mode=WAL;");
                    Exec(_w, "PRAGMA synchronous=NORMAL;");
                    Exec(_w, "PRAGMA wal_autocheckpoint=2000;");
                    Exec(_w, "PRAGMA journal_size_limit=67108864;");
                    Exec(_w, "PRAGMA cache_size=-65536;");
                    if (create) InitSchema(name, source);
                }
                LoadMeta();
                if (!readOnly && !create) EnsureAdditiveSchema();
                if (!readOnly) Recover();
            }
            catch
            {
                CloseHandles(true);
                throw;
            }
        }

        private SqliteConnection OpenConnection(bool reader)
        {
            var csb = new SqliteConnectionStringBuilder
            {
                DataSource = _dbPath,
                Mode = (reader && _readOnly) ? SqliteOpenMode.ReadOnly : SqliteOpenMode.ReadWriteCreate,
                Pooling = false,
                Cache = SqliteCacheMode.Private
            };
            var c = new SqliteConnection(csb.ToString());
            c.Open();
            Exec(c, "PRAGMA busy_timeout=" + _opt.BusyTimeoutMs.ToString(CultureInfo.InvariantCulture) + ";");
            Exec(c, "PRAGMA foreign_keys=ON;");
            Exec(c, "PRAGMA temp_store=MEMORY;");
            if (reader)
            {
                Exec(c, "PRAGMA cache_size=-8192;");
                Exec(c, "PRAGMA query_only=ON;");
            }
            return c;
        }

        public void Dispose()
        {
            lock (_stateLock)
            {
                if (_closed) return;
                _closing = true;
                // Let the running operations finish (a stalled stream would block here: use Abandon() to cut it).
                DateTime limit = DateTime.UtcNow.AddSeconds(60);
                while (_active > 0 && DateTime.UtcNow < limit) Monitor.Wait(_stateLock, 200);
                _closed = true;
            }
            CloseHandles(false);
        }

        /// <summary>
        /// Crash simulation (tests): closes every handle without waiting for the running writes, without sealing, truncating
        /// or checkpointing anything. What is on disk is what a killed process would leave.
        /// </summary>
        public void Abandon()
        {
            lock (_stateLock) { _closing = true; _closed = true; }
            CloseHandles(true);
        }

        private void CloseHandles(bool abrupt)
        {
            lock (_poolLock)
            {
                foreach (var w in _allWriters)
                {
                    w.Closed = true;
                    try { if (w.Stream != null) w.Stream.Dispose(); } catch { }
                }
                _allWriters.Clear();
                _idle.Clear();
            }
            lock (_readerLock)
            {
                while (_readers.Count > 0) { try { _readers.Pop().Dispose(); } catch { } }
            }
            lock (_dbLock)
            {
                if (_w != null)
                {
                    try { if (!abrupt) Exec(_w, "PRAGMA wal_checkpoint(TRUNCATE);"); } catch { }
                    try { _w.Dispose(); } catch { }
                    _w = null;
                }
            }
            if (_lockFile != null) { try { _lockFile.Dispose(); } catch { } _lockFile = null; }
        }

        private void EnterOp(bool write)
        {
            lock (_stateLock)
            {
                if (_closing || _closed) throw new ObjectDisposedException("ExportStore");
                if (write && _readOnly) throw new InvalidOperationException("The export is opened read-only.");
                _active++;
            }
        }

        private void ExitOp()
        {
            lock (_stateLock) { _active--; Monitor.PulseAll(_stateLock); }
        }

        #endregion

        #region SQL helpers

        private static int Exec(SqliteConnection c, string sql)
        {
            using (var cmd = c.CreateCommand()) { cmd.CommandText = sql; return cmd.ExecuteNonQuery(); }
        }

        /// <summary>Raises the write connection to synchronous=FULL for one import commit (call under _dbLock, outside a transaction); <see cref="SyncBack"/> restores NORMAL.</summary>
        private bool SyncFull(SqliteConnection c)
        {
            if (!_opt.ImportSyncFull) return false;
            Exec(c, "PRAGMA synchronous=FULL;");
            return true;
        }

        private static void SyncBack(SqliteConnection c, bool full)
        {
            if (full) Exec(c, "PRAGMA synchronous=NORMAL;");
        }

        private static SqliteCommand Cmd(SqliteConnection c, string sql, SqliteTransaction tx, params object[] nameValue)
        {
            var cmd = c.CreateCommand();
            cmd.CommandText = sql;
            cmd.Transaction = tx;
            for (int i = 0; i + 1 < nameValue.Length; i += 2)
            {
                var p = cmd.CreateParameter();
                p.ParameterName = (string)nameValue[i];
                p.Value = nameValue[i + 1] ?? DBNull.Value;
                cmd.Parameters.Add(p);
            }
            return cmd;
        }

        private static SqliteParameter Param(SqliteCommand cmd, string name)
        {
            var p = cmd.CreateParameter();
            p.ParameterName = name;
            cmd.Parameters.Add(p);
            return p;
        }

        private static object Db(object v) { return v ?? DBNull.Value; }

        private static string Str(SqliteDataReader r, int i) { return r.IsDBNull(i) ? null : r.GetString(i); }
        private static long? Lng(SqliteDataReader r, int i) { return r.IsDBNull(i) ? (long?)null : r.GetInt64(i); }
        private static int? Int(SqliteDataReader r, int i) { return r.IsDBNull(i) ? (int?)null : r.GetInt32(i); }

        private static long NowMs() { return DateTimeOffset.UtcNow.ToUnixTimeMilliseconds(); }
        private static object ToMs(DateTime? d)
        {
            if (!d.HasValue) return DBNull.Value;
            DateTime u = d.Value.Kind == DateTimeKind.Unspecified ? DateTime.SpecifyKind(d.Value, DateTimeKind.Utc) : d.Value.ToUniversalTime();
            return new DateTimeOffset(u).ToUnixTimeMilliseconds();
        }
        private static DateTime? FromMs(long? ms)
        {
            return ms.HasValue ? DateTimeOffset.FromUnixTimeMilliseconds(ms.Value).UtcDateTime : (DateTime?)null;
        }
        private static DateTime? FromMs(SqliteDataReader r, int i) { return FromMs(Lng(r, i)); }

        private static string Hex(byte[] b) { return Convert.ToHexString(b).ToLowerInvariant(); }

        private static string TruncateUtf8(string s, int maxBytes)
        {
            if (string.IsNullOrEmpty(s)) return s;
            if (Encoding.UTF8.GetByteCount(s) <= maxBytes) return s;
            var sb = new StringBuilder();
            int bytes = 0;
            foreach (Rune r in s.EnumerateRunes())
            {
                int l = r.Utf8SequenceLength;
                if (bytes + l > maxBytes) break;
                bytes += l;
                sb.Append(r.ToString());
            }
            return sb.ToString();
        }

        private static string Clip(string s, int maxChars)
        {
            return (s != null && s.Length > maxChars) ? s.Substring(0, maxChars) : s;
        }

        private SqliteConnection W()
        {
            if (_w == null) throw new ObjectDisposedException("ExportStore");
            return _w;
        }

        private sealed class ReadLease : IDisposable
        {
            private readonly ExportStore _s;
            public readonly SqliteConnection Conn;
            public ReadLease(ExportStore s, SqliteConnection c) { _s = s; Conn = c; }
            public void Dispose() { _s.ReturnReader(Conn); }
        }

        private ReadLease Reader()
        {
            SqliteConnection c = null;
            lock (_readerLock) { if (_readers.Count > 0) c = _readers.Pop(); }
            if (c == null) c = OpenConnection(true);
            return new ReadLease(this, c);
        }

        private void ReturnReader(SqliteConnection c)
        {
            bool keep;
            lock (_readerLock)
            {
                keep = !_closed && _readers.Count < _opt.MaxIdleReaders;
                if (keep) _readers.Push(c);
            }
            if (!keep) { try { c.Dispose(); } catch { } }
        }

        #endregion

        #region Schema

        private const string SchemaSql = @"
CREATE TABLE export(
  id             INTEGER PRIMARY KEY CHECK (id = 1),
  name           TEXT NOT NULL,
  source         TEXT,
  created_utc    INTEGER NOT NULL,
  schema_version INTEGER NOT NULL
);
CREATE TABLE mailbox(
  mailbox_key  INTEGER PRIMARY KEY,
  mailbox_id   TEXT NOT NULL UNIQUE,          -- physical id of the mailbox (GUID)
  kind         INTEGER NOT NULL,              -- 0 Primary, 1 Archive
  upn          TEXT,
  display_name TEXT
);
CREATE TABLE folder(
  folder_key        INTEGER PRIMARY KEY,
  mailbox_key       INTEGER NOT NULL REFERENCES mailbox(mailbox_key),
  folder_id         TEXT NOT NULL,
  parent_folder_id  TEXT,
  display_name      TEXT,
  path              TEXT,                      -- computed by the store from the parent chain
  folder_type       TEXT,
  source_item_count INTEGER,                   -- as announced by the source
  source_size       INTEGER,
  UNIQUE (mailbox_key, folder_id)
);
CREATE INDEX ix_folder_parent ON folder(mailbox_key, parent_folder_id);
CREATE TABLE pack(
  pack_id     INTEGER PRIMARY KEY AUTOINCREMENT,   -- file: packs\pack-<id %06d>.bin
  created_utc INTEGER NOT NULL,
  sealed      INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE item(
  item_key     INTEGER PRIMARY KEY,
  mailbox_key  INTEGER NOT NULL REFERENCES mailbox(mailbox_key),
  item_id      TEXT NOT NULL,
  folder_key   INTEGER REFERENCES folder(folder_key),
  change_key   TEXT,
  item_type    TEXT,
  size         INTEGER NOT NULL DEFAULT 0,          -- as announced by the source
  state        INTEGER NOT NULL DEFAULT 0,          -- 0 Pending, 1 Exported, 2 Failed, 3 Skipped
  pack_id      INTEGER REFERENCES pack(pack_id),
  pack_offset  INTEGER,
  pack_length  INTEGER,
  sha256       BLOB,
  exported_utc INTEGER,
  attempts     INTEGER NOT NULL DEFAULT 0,
  last_error   TEXT,
  received_utc INTEGER,                              -- denormalised from preview (sorting)
  subject      TEXT,                                 -- denormalised from preview (sorting)
  UNIQUE (mailbox_key, item_id)
);
CREATE INDEX ix_item_pending ON item(mailbox_key, item_key) WHERE state = 0;
CREATE INDEX ix_item_agg     ON item(mailbox_key, folder_key, state, size, pack_length);
CREATE INDEX ix_item_size    ON item(folder_key, size);
CREATE INDEX ix_item_date    ON item(folder_key, received_utc);
CREATE INDEX ix_item_subject ON item(folder_key, subject COLLATE NOCASE);
CREATE INDEX ix_item_pack    ON item(pack_id) WHERE pack_id IS NOT NULL;
CREATE TABLE preview(
  item_key         INTEGER PRIMARY KEY REFERENCES item(item_key),
  subject          TEXT,
  from_text        TEXT,
  to_text          TEXT,
  received_utc     INTEGER,
  sent_utc         INTEGER,
  message_class    TEXT,
  has_attachments  INTEGER NOT NULL DEFAULT 0,
  attachment_count INTEGER NOT NULL DEFAULT 0,
  attachment_names TEXT,
  body_preview     TEXT                              -- <= 1 KiB UTF-8
);
CREATE TABLE importJob(
  job_id         INTEGER PRIMARY KEY,
  name           TEXT,
  target_mailbox TEXT NOT NULL,
  created_utc    INTEGER NOT NULL
);
CREATE TABLE importItem(
  import_key     INTEGER PRIMARY KEY,
  job_id         INTEGER NOT NULL REFERENCES importJob(job_id),
  item_key       INTEGER NOT NULL REFERENCES item(item_key),
  target_mailbox TEXT,                               -- null = importJob.target_mailbox
  target_folder  TEXT,
  state          INTEGER NOT NULL DEFAULT 0,         -- 0 Pending, 1 InFlight, 2 Done, 3 Failed, 4 Unknown
  target_item_id TEXT,
  attempts       INTEGER NOT NULL DEFAULT 0,
  last_error     TEXT,
  updated_utc    INTEGER,
  UNIQUE (job_id, item_key)
);
CREATE INDEX ix_importitem_state ON importItem(job_id, state, import_key);
" + AdditiveSchemaSql;

        // Tables added after schema 1 without raising the version (an older build ignores them): created when a store is created or opened for writing.
        private const string AdditiveSchemaSql = @"
CREATE TABLE IF NOT EXISTS mailboxLocation(
  mailbox_id     TEXT PRIMARY KEY,                   -- MBX:<guid>@<tenant>
  kind           INTEGER NOT NULL,                   -- 0 Primary, 1 Archive (MainArchive), 2 AuxArchive
  upn            TEXT,
  source         TEXT NOT NULL,                      -- how it was discovered
  detail         TEXT,
  rank           INTEGER NOT NULL DEFAULT 1,         -- 2 given (operator, Graph), 1 learned from a redirect
  discovered_utc INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS importRoute(
  job_id       INTEGER NOT NULL,
  from_mailbox TEXT NOT NULL,                        -- planned target mailbox
  from_folder  TEXT NOT NULL,                        -- planned target folder
  to_mailbox   TEXT NOT NULL,                        -- auxiliary archive that serves the folder
  to_folder    TEXT NOT NULL,                        -- folder id as that mailbox lists it
  learned_utc  INTEGER NOT NULL,
  PRIMARY KEY (job_id, from_mailbox, from_folder)
);
";

        private void InitSchema(string name, string source)
        {
            lock (_dbLock)
            {
                var c = W();
                using (var tx = c.BeginTransaction())
                {
                    foreach (var stmt in SplitStatements(SchemaSql))
                        using (var cmd = Cmd(c, stmt, tx)) cmd.ExecuteNonQuery();
                    using (var cmd = Cmd(c, "INSERT INTO export(id,name,source,created_utc,schema_version) VALUES(1,@n,@s,@t,@v)", tx,
                        "@n", name ?? System.IO.Path.GetFileName(_root), "@s", source, "@t", NowMs(), "@v", SchemaVersion))
                        cmd.ExecuteNonQuery();
                    using (var cmd = Cmd(c, "PRAGMA user_version=" + SchemaVersion.ToString(CultureInfo.InvariantCulture), tx)) cmd.ExecuteNonQuery();
                    tx.Commit();
                }
            }
        }

        private void EnsureAdditiveSchema()
        {
            lock (_dbLock)
            {
                var c = W();
                foreach (var stmt in SplitStatements(AdditiveSchemaSql))
                    using (var cmd = Cmd(c, stmt, null)) cmd.ExecuteNonQuery();
            }
        }

        private static IEnumerable<string> SplitStatements(string sql)
        {
            // The schema has no ';' inside literals/comments: a plain split is enough (comments after -- are kept in the statement).
            foreach (var part in sql.Split(';'))
            {
                string s = part.Trim();
                if (s.Length > 0) yield return s;
            }
        }

        private void LoadMeta()
        {
            using (var l = Reader())
            {
                using (var cmd = Cmd(l.Conn, "PRAGMA user_version", null))
                {
                    long v = Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                    if (v > SchemaVersion) throw new InvalidDataException("The export uses schema version " + v + ", this build supports up to " + SchemaVersion + ".");
                    if (v < 1) throw new InvalidDataException("Not a MailboxTransfer export catalog (user_version " + v + ").");
                }
                using (var cmd = Cmd(l.Conn, "SELECT name, source, created_utc FROM export WHERE id=1", null))
                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) throw new InvalidDataException("Export metadata missing.");
                    Name = r.GetString(0); Source = Str(r, 1); CreatedUtc = FromMs(r, 2).Value;
                }
            }
        }

        #endregion

        #region Recovery

        public string GetPackPath(int packId)
        {
            return System.IO.Path.Combine(_packsDir, "pack-" + packId.ToString("D6", CultureInfo.InvariantCulture) + ".bin");
        }

        private void Recover()
        {
            var rep = new StoreRecoveryReport();
            lock (_dbLock)
            {
                var c = W();
                var packs = new List<KeyValuePair<int, bool>>();
                using (var cmd = Cmd(c, "SELECT pack_id, sealed FROM pack ORDER BY pack_id", null))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) packs.Add(new KeyValuePair<int, bool>(r.GetInt32(0), r.GetInt32(1) != 0));
                var known = new HashSet<string>(packs.Select(p => System.IO.Path.GetFileName(GetPackPath(p.Key))), StringComparer.OrdinalIgnoreCase);

                foreach (var p in packs)
                {
                    rep.PacksChecked++;
                    string path = GetPackPath(p.Key);
                    long fileLen = File.Exists(path) ? new FileInfo(path).Length : 0;
                    long expected = CommittedEnd(c, p.Key);
                    if (fileLen < expected)
                    {
                        // Damage the process cannot cause (the bytes are flushed before the commit): treat the items as to do.
                        rep.DamagedPacks++;
                        using (var tx = c.BeginTransaction())
                        {
                            using (var cmd = Cmd(c, "UPDATE item SET state=0, pack_id=NULL, pack_offset=NULL, pack_length=NULL, sha256=NULL, exported_utc=NULL " +
                                "WHERE pack_id=@p AND state=1 AND pack_offset + pack_length > @len", tx, "@p", p.Key, "@len", fileLen))
                                rep.LostItems += cmd.ExecuteNonQuery();
                            tx.Commit();
                        }
                        expected = CommittedEnd(c, p.Key);
                    }
                    if (fileLen > expected)
                    {
                        using (var fs = new FileStream(path, FileMode.Open, FileAccess.Write, FileShare.None))
                        {
                            fs.SetLength(expected);
                            fs.Flush(true);
                        }
                        rep.TruncatedPacks++;
                        rep.TruncatedBytes += fileLen - expected;
                    }
                    if (!p.Value)
                    {
                        if (expected >= _opt.PackTargetSize)
                            using (var cmd = Cmd(c, "UPDATE pack SET sealed=1 WHERE pack_id=@p", null, "@p", p.Key)) cmd.ExecuteNonQuery();
                        else
                            lock (_poolLock) _reusable.Enqueue(p.Key);
                    }
                }
                if (Directory.Exists(_packsDir))
                    foreach (var f in Directory.EnumerateFiles(_packsDir))
                        if (!known.Contains(System.IO.Path.GetFileName(f))) rep.UnknownPackFiles++;

                using (var cmd = Cmd(c, "SELECT job_id, COUNT(*) FROM importItem WHERE state=1 GROUP BY job_id", null))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) rep.ImportInFlightToUnknownByJob[r.GetInt64(0)] = r.GetInt32(1);
                using (var cmd = Cmd(c, "UPDATE importItem SET state=4, updated_utc=@t WHERE state=1", null, "@t", NowMs()))
                    rep.ImportInFlightToUnknown = cmd.ExecuteNonQuery();
            }
            LastRecovery = rep;
        }

        private static long CommittedEnd(SqliteConnection c, int packId)
        {
            using (var cmd = Cmd(c, "SELECT MAX(pack_offset + pack_length) FROM item WHERE pack_id=@p AND state=1", null, "@p", packId))
            {
                object o = cmd.ExecuteScalar();
                return (o == null || o == DBNull.Value) ? 0 : Convert.ToInt64(o, CultureInfo.InvariantCulture);
            }
        }

        #endregion

        #region Keys

        private int MailboxKey(string mailboxId)
        {
            int k;
            if (_mailboxKeys.TryGetValue(mailboxId, out k)) return k;
            using (var l = Reader())
            using (var cmd = Cmd(l.Conn, "SELECT mailbox_key FROM mailbox WHERE mailbox_id=@m", null, "@m", mailboxId))
            {
                object o = cmd.ExecuteScalar();
                if (o == null) throw new KeyNotFoundException("Unknown mailbox: " + mailboxId);
                k = Convert.ToInt32(o, CultureInfo.InvariantCulture);
            }
            _mailboxKeys[mailboxId] = k;
            return k;
        }

        private static string FolderCacheKey(int mk, string folderId) { return mk.ToString(CultureInfo.InvariantCulture) + "|" + folderId; }

        /// <summary>Folder key of (mailbox, folderId), or null if unknown.</summary>
        public long? FindFolderKey(string mailboxId, string folderId)
        {
            EnterOp(false);
            try
            {
                int mk = MailboxKey(mailboxId);
                return FolderKeyR(mk, folderId);
            }
            finally { ExitOp(); }
        }

        private long? FolderKeyR(int mk, string folderId)
        {
            long k;
            string ck = FolderCacheKey(mk, folderId);
            if (_folderKeys.TryGetValue(ck, out k)) return k;
            using (var l = Reader())
            using (var cmd = Cmd(l.Conn, "SELECT folder_key FROM folder WHERE mailbox_key=@m AND folder_id=@f", null, "@m", mk, "@f", folderId))
            {
                object o = cmd.ExecuteScalar();
                if (o == null) return null;
                k = Convert.ToInt64(o, CultureInfo.InvariantCulture);
            }
            _folderKeys[ck] = k;
            return k;
        }

        /// <summary>Key of the item (mailbox, itemId), or null if unknown.</summary>
        public long? FindItemKey(string mailboxId, string itemId)
        {
            EnterOp(false);
            try
            {
                int mk = MailboxKey(mailboxId);
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT item_key FROM item WHERE mailbox_key=@m AND item_id=@i", null, "@m", mk, "@i", itemId))
                {
                    object o = cmd.ExecuteScalar();
                    return o == null ? (long?)null : Convert.ToInt64(o, CultureInfo.InvariantCulture);
                }
            }
            finally { ExitOp(); }
        }

        #endregion

        #region Mailboxes, folders, items (catalog)

        public int UpsertMailbox(StoreMailbox mailbox)
        {
            if (mailbox == null || string.IsNullOrEmpty(mailbox.Id)) throw new ArgumentException("mailbox.Id");
            EnterOp(true);
            try
            {
                lock (_dbLock)
                {
                    var c = W();
                    using (var cmd = Cmd(c, "INSERT INTO mailbox(mailbox_id,kind,upn,display_name) VALUES(@i,@k,@u,@d) " +
                        "ON CONFLICT(mailbox_id) DO UPDATE SET kind=excluded.kind, upn=excluded.upn, display_name=excluded.display_name", null,
                        "@i", mailbox.Id, "@k", (int)mailbox.Kind, "@u", mailbox.Upn, "@d", mailbox.DisplayName))
                        cmd.ExecuteNonQuery();
                    using (var cmd = Cmd(c, "SELECT mailbox_key FROM mailbox WHERE mailbox_id=@i", null, "@i", mailbox.Id))
                        mailbox.Key = Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                    _mailboxKeys[mailbox.Id] = mailbox.Key;
                    return mailbox.Key;
                }
            }
            finally { ExitOp(); }
        }

        /// <summary>Like <see cref="UpsertMailbox"/> but an existing mailbox is left as it is (its kind is not overwritten). Returns the key.</summary>
        public int EnsureMailbox(StoreMailbox mailbox)
        {
            if (mailbox == null || string.IsNullOrEmpty(mailbox.Id)) throw new ArgumentException("mailbox.Id");
            EnterOp(true);
            try
            {
                lock (_dbLock)
                {
                    var c = W();
                    using (var cmd = Cmd(c, "INSERT INTO mailbox(mailbox_id,kind,upn,display_name) VALUES(@i,@k,@u,@d) ON CONFLICT(mailbox_id) DO NOTHING", null,
                        "@i", mailbox.Id, "@k", (int)mailbox.Kind, "@u", mailbox.Upn, "@d", mailbox.DisplayName))
                        cmd.ExecuteNonQuery();
                    using (var cmd = Cmd(c, "SELECT mailbox_key FROM mailbox WHERE mailbox_id=@i", null, "@i", mailbox.Id))
                        mailbox.Key = Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                    _mailboxKeys[mailbox.Id] = mailbox.Key;
                    return mailbox.Key;
                }
            }
            finally { ExitOp(); }
        }

        public List<StoreMailbox> GetMailboxes()
        {
            EnterOp(false);
            try
            {
                var list = new List<StoreMailbox>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT mailbox_key, mailbox_id, kind, upn, display_name FROM mailbox ORDER BY mailbox_key", null))
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        list.Add(new StoreMailbox { Key = r.GetInt32(0), Id = r.GetString(1), Kind = (StoreMailboxKind)r.GetInt32(2), Upn = Str(r, 3), DisplayName = Str(r, 4) });
                return list;
            }
            finally { ExitOp(); }
        }

        /// <summary>
        /// Records a mailbox location (see <see cref="StoreMailboxLocation"/>). A location already known keeps its first discovery, unless the
        /// new one has a higher rank (given by the operator or Graph beats learned from a redirect). Returns true when the mailbox was not known.
        /// </summary>
        public bool RecordLocation(StoreMailboxLocation loc)
        {
            if (loc == null || string.IsNullOrEmpty(loc.MailboxId)) throw new ArgumentException("loc.MailboxId");
            if (string.IsNullOrEmpty(loc.Source)) throw new ArgumentException("loc.Source");
            EnterOp(true);
            try
            {
                lock (_dbLock)
                {
                    var c = W();
                    bool isNew;
                    using (var cmd = Cmd(c, "SELECT COUNT(*) FROM mailboxLocation WHERE mailbox_id=@i", null, "@i", loc.MailboxId))
                        isNew = Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) == 0;
                    using (var cmd = Cmd(c, "INSERT INTO mailboxLocation(mailbox_id,kind,upn,source,detail,rank,discovered_utc) VALUES(@i,@k,@u,@s,@d,@r,@t) " +
                        "ON CONFLICT(mailbox_id) DO UPDATE SET kind=excluded.kind, upn=COALESCE(excluded.upn, mailboxLocation.upn), source=excluded.source, " +
                        "detail=excluded.detail, rank=excluded.rank, discovered_utc=excluded.discovered_utc WHERE excluded.rank > mailboxLocation.rank", null,
                        "@i", loc.MailboxId, "@k", (int)loc.Kind, "@u", Db(loc.Upn), "@s", loc.Source, "@d", Db(loc.Detail == null ? null : Clip(loc.Detail, 500)),
                        "@r", loc.Rank, "@t", ToMs(loc.DiscoveredUtc)))
                        cmd.ExecuteNonQuery();
                    if (!string.IsNullOrEmpty(loc.Upn))
                        using (var cmd = Cmd(c, "UPDATE mailboxLocation SET upn=@u WHERE mailbox_id=@i AND upn IS NULL", null, "@u", loc.Upn, "@i", loc.MailboxId))
                            cmd.ExecuteNonQuery();
                    return isNew;
                }
            }
            finally { ExitOp(); }
        }

        /// <summary>Every known location of the store, Primary first, then the archive, then the auxiliary archives in the order they were discovered. Empty for a store written before the table existed and opened read-only.</summary>
        public List<StoreMailboxLocation> GetLocations()
        {
            EnterOp(false);
            try
            {
                var list = new List<StoreMailboxLocation>();
                try
                {
                    using (var l = Reader())
                    using (var cmd = Cmd(l.Conn, "SELECT mailbox_id, kind, upn, source, detail, rank, discovered_utc FROM mailboxLocation ORDER BY kind, discovered_utc, mailbox_id", null))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                            list.Add(new StoreMailboxLocation
                            {
                                MailboxId = r.GetString(0), Kind = (StoreMailboxKind)r.GetInt32(1), Upn = Str(r, 2), Source = r.GetString(3), Detail = Str(r, 4),
                                Rank = r.GetInt32(5), DiscoveredUtc = FromMs(r, 6) ?? DateTime.UtcNow
                            });
                }
                catch (SqliteException ex) when (ex.Message.IndexOf("no such table", StringComparison.OrdinalIgnoreCase) >= 0) { }
                return list;
            }
            finally { ExitOp(); }
        }

        /// <summary>Mailbox id of a folder record with this folder id in a mailbox of this kind (the main archive holds the record of an expanded folder), or null.</summary>
        public string FindFolderMailbox(string folderId, StoreMailboxKind kind)
        {
            EnterOp(false);
            try
            {
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT m.mailbox_id FROM folder f JOIN mailbox m ON m.mailbox_key=f.mailbox_key WHERE f.folder_id=@f AND m.kind=@k ORDER BY m.mailbox_key LIMIT 1", null,
                    "@f", folderId, "@k", (int)kind))
                {
                    object o = cmd.ExecuteScalar();
                    return o == null ? null : Convert.ToString(o, CultureInfo.InvariantCulture);
                }
            }
            finally { ExitOp(); }
        }

        /// <summary>
        /// An item turned out to live in another mailbox than the one it was listed in (an auxiliary archive answered ErrorArchiveFolderMovedPermanently):
        /// its record is moved there. False when that mailbox already holds a record of the same item id (nothing is changed then).
        /// </summary>
        public bool MoveItemToMailbox(long itemKey, string mailboxId)
        {
            EnterOp(true);
            try
            {
                lock (_dbLock)
                {
                    var c = W();
                    int mk = MailboxKeyW(c, null, mailboxId);
                    try
                    {
                        using (var cmd = Cmd(c, "UPDATE item SET mailbox_key=@m WHERE item_key=@k", null, "@m", mk, "@k", itemKey))
                            return cmd.ExecuteNonQuery() == 1;
                    }
                    catch (SqliteException ex) when (ex.SqliteErrorCode == 19) { return false; }   // UNIQUE (mailbox_key, item_id)
                }
            }
            finally { ExitOp(); }
        }

        /// <summary>The items recorded in one folder (whatever mailbox they are recorded under): key, mailbox, item id, size and state only.</summary>
        public List<StoreItem> GetFolderItemRefs(long folderKey)
        {
            EnterOp(false);
            try
            {
                var list = new List<StoreItem>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT i.item_key, m.mailbox_id, i.item_id, i.size, i.state FROM item i JOIN mailbox m ON m.mailbox_key=i.mailbox_key WHERE i.folder_key=@f", null, "@f", folderKey))
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        list.Add(new StoreItem { Key = r.GetInt64(0), MailboxId = r.GetString(1), ItemId = r.GetString(2), Size = r.GetInt64(3), State = (StoreItemState)r.GetInt32(4), FolderKey = folderKey });
                return list;
            }
            finally { ExitOp(); }
        }

        /// <summary>
        /// The record of an item that is not exported yet is now known by another mailbox and another item id (the same message listed by the auxiliary
        /// archive it moved to): the record follows, and a Failed one is Pending again with no attempt counted. False when that mailbox already holds the new id
        /// (nothing is changed) or the item is Exported (an exported record is never touched: the bytes are in the store).
        /// </summary>
        public bool RehomeItem(long itemKey, string mailboxId, string newItemId)
        {
            EnterOp(true);
            try
            {
                lock (_dbLock)
                {
                    var c = W();
                    int mk = MailboxKeyW(c, null, mailboxId);
                    try
                    {
                        using (var cmd = Cmd(c, "UPDATE item SET mailbox_key=@m, item_id=@i, state=CASE WHEN state=2 THEN 0 ELSE state END, attempts=CASE WHEN state=2 THEN 0 ELSE attempts END, last_error=CASE WHEN state=2 THEN NULL ELSE last_error END " +
                            "WHERE item_key=@k AND state<>1", null, "@m", mk, "@i", newItemId, "@k", itemKey))
                            return cmd.ExecuteNonQuery() == 1;
                    }
                    catch (SqliteException ex) when (ex.SqliteErrorCode == 19) { return false; }   // UNIQUE (mailbox_key, item_id)
                }
            }
            finally { ExitOp(); }
        }

        private static IEnumerable<List<T>> Chunks<T>(IEnumerable<T> source, int size)
        {
            var buf = new List<T>(size);
            foreach (var x in source)
            {
                buf.Add(x);
                if (buf.Count >= size) { yield return buf; buf = new List<T>(size); }
            }
            if (buf.Count > 0) yield return buf;
        }

        // Array overloads: PowerShell cannot convert object[] to IEnumerable<T> of a type compiled with Add-Type.
        public int UpsertFolders(StoreFolder[] folders) { return UpsertFolders((IEnumerable<StoreFolder>)folders); }
        public int UpsertItems(StoreItemEntry[] items) { return UpsertItems((IEnumerable<StoreItemEntry>)items); }
        public int SetPreviews(StorePreview[] previews) { return SetPreviews((IEnumerable<StorePreview>)previews); }
        public int AddImportItems(long jobId, StoreImportPlan[] plan) { return AddImportItems(jobId, (IEnumerable<StoreImportPlan>)plan); }

        /// <summary>Inserts or updates folders (transactions of UpsertBatchSize rows) and recomputes the full paths.</summary>
        public int UpsertFolders(IEnumerable<StoreFolder> folders)
        {
            if (folders == null) throw new ArgumentNullException("folders");
            EnterOp(true);
            try
            {
                int n = 0;
                var touched = new HashSet<int>();
                foreach (var chunk in Chunks(folders, _opt.UpsertBatchSize))
                {
                    lock (_dbLock)
                    {
                        var c = W();
                        using (var tx = c.BeginTransaction())
                        {
                            using (var up = Cmd(c, "INSERT INTO folder(mailbox_key,folder_id,parent_folder_id,display_name,folder_type,source_item_count,source_size) " +
                                "VALUES(@m,@f,@p,@d,@t,@c,@s) ON CONFLICT(mailbox_key,folder_id) DO UPDATE SET parent_folder_id=excluded.parent_folder_id, " +
                                "display_name=excluded.display_name, folder_type=excluded.folder_type, source_item_count=excluded.source_item_count, source_size=excluded.source_size", tx))
                            using (var sel = Cmd(c, "SELECT folder_key FROM folder WHERE mailbox_key=@m AND folder_id=@f", tx))
                            {
                                var pm = Param(up, "@m"); var pf = Param(up, "@f"); var pp = Param(up, "@p"); var pd = Param(up, "@d");
                                var pt = Param(up, "@t"); var pc = Param(up, "@c"); var ps = Param(up, "@s");
                                var sm = Param(sel, "@m"); var sf = Param(sel, "@f");
                                foreach (var f in chunk)
                                {
                                    if (f == null || string.IsNullOrEmpty(f.FolderId)) throw new ArgumentException("FolderId is required.");
                                    int mk = MailboxKeyW(c, tx, f.MailboxId);
                                    pm.Value = mk; pf.Value = f.FolderId; pp.Value = Db(f.ParentFolderId); pd.Value = Db(f.DisplayName);
                                    pt.Value = Db(f.FolderType); pc.Value = Db(f.SourceItemCount); ps.Value = Db(f.SourceSize);
                                    up.ExecuteNonQuery();
                                    sm.Value = mk; sf.Value = f.FolderId;
                                    long key = Convert.ToInt64(sel.ExecuteScalar(), CultureInfo.InvariantCulture);
                                    f.Key = key;
                                    _folderKeys[FolderCacheKey(mk, f.FolderId)] = key;
                                    touched.Add(mk);
                                    n++;
                                }
                            }
                            tx.Commit();
                        }
                    }
                }
                foreach (int mk in touched) RecomputePaths(mk);
                return n;
            }
            finally { ExitOp(); }
        }

        private int MailboxKeyW(SqliteConnection c, SqliteTransaction tx, string mailboxId)
        {
            int k;
            if (mailboxId != null && _mailboxKeys.TryGetValue(mailboxId, out k)) return k;
            using (var cmd = Cmd(c, "SELECT mailbox_key FROM mailbox WHERE mailbox_id=@m", tx, "@m", mailboxId))
            {
                object o = cmd.ExecuteScalar();
                if (o == null) throw new KeyNotFoundException("Unknown mailbox: " + mailboxId + " (call UpsertMailbox first).");
                k = Convert.ToInt32(o, CultureInfo.InvariantCulture);
            }
            _mailboxKeys[mailboxId] = k;
            return k;
        }

        private void RecomputePaths(int mk)
        {
            lock (_dbLock)
            {
                var c = W();
                var rows = new Dictionary<string, string[]>(StringComparer.Ordinal);   // folderId -> [key, parent, name, path]
                using (var cmd = Cmd(c, "SELECT folder_key, folder_id, parent_folder_id, display_name, path FROM folder WHERE mailbox_key=@m", null, "@m", mk))
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        rows[r.GetString(1)] = new[] { r.GetInt64(0).ToString(CultureInfo.InvariantCulture), Str(r, 2), Str(r, 3) ?? "", Str(r, 4) };
                var changes = new List<KeyValuePair<long, string>>();
                foreach (var kv in rows)
                {
                    var parts = new List<string>();
                    var seen = new HashSet<string>(StringComparer.Ordinal);
                    string cur = kv.Key;
                    while (cur != null && rows.ContainsKey(cur) && seen.Add(cur))
                    {
                        parts.Add(rows[cur][2]);
                        cur = rows[cur][1];
                    }
                    parts.Reverse();
                    string path = string.Join("/", parts);
                    if (!string.Equals(path, kv.Value[3], StringComparison.Ordinal))
                        changes.Add(new KeyValuePair<long, string>(long.Parse(kv.Value[0], CultureInfo.InvariantCulture), path));
                }
                if (changes.Count == 0) return;
                using (var tx = c.BeginTransaction())
                {
                    using (var up = Cmd(c, "UPDATE folder SET path=@p WHERE folder_key=@k", tx))
                    {
                        var pp = Param(up, "@p"); var pk = Param(up, "@k");
                        foreach (var ch in changes) { pp.Value = ch.Value; pk.Value = ch.Key; up.ExecuteNonQuery(); }
                    }
                    tx.Commit();
                }
            }
        }

        public List<StoreFolder> GetFolders(string mailboxId)
        {
            EnterOp(false);
            try
            {
                int mk = MailboxKey(mailboxId);
                var list = new List<StoreFolder>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT folder_key, folder_id, parent_folder_id, display_name, path, folder_type, source_item_count, source_size " +
                    "FROM folder WHERE mailbox_key=@m ORDER BY path", null, "@m", mk))
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        list.Add(new StoreFolder
                        {
                            Key = r.GetInt64(0), MailboxId = mailboxId, FolderId = r.GetString(1), ParentFolderId = Str(r, 2), DisplayName = Str(r, 3),
                            Path = Str(r, 4), FolderType = Str(r, 5), SourceItemCount = Lng(r, 6), SourceSize = Lng(r, 7)
                        });
                return list;
            }
            finally { ExitOp(); }
        }

        /// <summary>
        /// Registers the items found at enumeration time. An existing item keeps its state, except when its changeKey changed
        /// and it was Exported: it goes back to Pending (the old bytes stay in the pack as dead space).
        /// </summary>
        public int UpsertItems(IEnumerable<StoreItemEntry> items)
        {
            if (items == null) throw new ArgumentNullException("items");
            EnterOp(true);
            try
            {
                int n = 0;
                foreach (var chunk in Chunks(items, _opt.UpsertBatchSize))
                {
                    lock (_dbLock)
                    {
                        var c = W();
                        using (var tx = c.BeginTransaction())
                        {
                            using (var up = Cmd(c,
                                "INSERT INTO item(mailbox_key,item_id,folder_key,change_key,item_type,size) VALUES(@m,@i,@f,@c,@t,@s) " +
                                "ON CONFLICT(mailbox_key,item_id) DO UPDATE SET folder_key=excluded.folder_key, item_type=excluded.item_type, size=excluded.size, " +
                                "state = CASE WHEN item.state=1 AND excluded.change_key IS NOT NULL AND item.change_key IS NOT excluded.change_key THEN 0 ELSE item.state END, " +
                                "pack_id = CASE WHEN item.state=1 AND excluded.change_key IS NOT NULL AND item.change_key IS NOT excluded.change_key THEN NULL ELSE item.pack_id END, " +
                                "pack_offset = CASE WHEN item.state=1 AND excluded.change_key IS NOT NULL AND item.change_key IS NOT excluded.change_key THEN NULL ELSE item.pack_offset END, " +
                                "pack_length = CASE WHEN item.state=1 AND excluded.change_key IS NOT NULL AND item.change_key IS NOT excluded.change_key THEN NULL ELSE item.pack_length END, " +
                                "sha256 = CASE WHEN item.state=1 AND excluded.change_key IS NOT NULL AND item.change_key IS NOT excluded.change_key THEN NULL ELSE item.sha256 END, " +
                                "change_key = COALESCE(excluded.change_key, item.change_key)", tx))
                            {
                                var pm = Param(up, "@m"); var pi = Param(up, "@i"); var pf = Param(up, "@f");
                                var pc = Param(up, "@c"); var pt = Param(up, "@t"); var ps = Param(up, "@s");
                                foreach (var e in chunk)
                                {
                                    if (e == null || string.IsNullOrEmpty(e.ItemId)) throw new ArgumentException("ItemId is required.");
                                    int mk = MailboxKeyW(c, tx, e.MailboxId);
                                    int fmk = string.IsNullOrEmpty(e.FolderMailboxId) ? mk : MailboxKeyW(c, tx, e.FolderMailboxId);
                                    object fk = DBNull.Value;
                                    if (!string.IsNullOrEmpty(e.FolderId))
                                    {
                                        long k;
                                        string ck = FolderCacheKey(fmk, e.FolderId);
                                        if (!_folderKeys.TryGetValue(ck, out k))
                                        {
                                            using (var sel = Cmd(c, "SELECT folder_key FROM folder WHERE mailbox_key=@m AND folder_id=@f", tx, "@m", fmk, "@f", e.FolderId))
                                            {
                                                object o = sel.ExecuteScalar();
                                                if (o == null) throw new KeyNotFoundException("Unknown folder: " + e.FolderId + " (call UpsertFolders first).");
                                                k = Convert.ToInt64(o, CultureInfo.InvariantCulture);
                                            }
                                            _folderKeys[ck] = k;
                                        }
                                        fk = k;
                                    }
                                    if (fmk != mk)
                                    {
                                        // Listed through the main archive before the item was known to live in an auxiliary archive: the same item is not recorded twice.
                                        using (var mv = Cmd(c, "UPDATE item SET mailbox_key=@m WHERE mailbox_key=@o AND item_id=@i AND folder_key=@f " +
                                            "AND NOT EXISTS (SELECT 1 FROM item x WHERE x.mailbox_key=@m AND x.item_id=@i)", tx, "@m", mk, "@o", fmk, "@i", e.ItemId, "@f", fk))
                                            mv.ExecuteNonQuery();
                                    }
                                    pm.Value = mk; pi.Value = e.ItemId; pf.Value = fk; pc.Value = Db(e.ChangeKey); pt.Value = Db(e.ItemType); ps.Value = e.Size;
                                    up.ExecuteNonQuery();
                                    n++;
                                }
                            }
                            tx.Commit();
                        }
                    }
                }
                return n;
            }
            finally { ExitOp(); }
        }

        private const string ItemSelect =
            "SELECT i.item_key, m.mailbox_id, i.item_id, i.folder_key, f.folder_id, i.change_key, i.item_type, i.size, i.state, i.pack_id, i.pack_offset, " +
            "i.pack_length, i.sha256, i.exported_utc, i.attempts, i.last_error, i.subject, i.received_utc " +
            "FROM item i JOIN mailbox m ON m.mailbox_key = i.mailbox_key LEFT JOIN folder f ON f.folder_key = i.folder_key ";

        private static StoreItem ReadItem(SqliteDataReader r)
        {
            return new StoreItem
            {
                Key = r.GetInt64(0),
                MailboxId = r.GetString(1),
                ItemId = r.GetString(2),
                FolderKey = Lng(r, 3),
                FolderId = Str(r, 4),
                ChangeKey = Str(r, 5),
                ItemType = Str(r, 6),
                Size = r.GetInt64(7),
                State = (StoreItemState)r.GetInt32(8),
                PackId = Int(r, 9),
                Offset = Lng(r, 10),
                Length = Lng(r, 11),
                Sha256 = r.IsDBNull(12) ? null : Hex(r.GetFieldValue<byte[]>(12)),
                ExportedUtc = FromMs(r, 13),
                Attempts = r.GetInt32(14),
                LastError = Str(r, 15),
                Subject = Str(r, 16),
                ReceivedUtc = FromMs(r, 17)
            };
        }

        public StoreItem GetItem(long itemKey)
        {
            EnterOp(false);
            try
            {
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, ItemSelect + "WHERE i.item_key=@k", null, "@k", itemKey))
                using (var r = cmd.ExecuteReader())
                    return r.Read() ? ReadItem(r) : null;
            }
            finally { ExitOp(); }
        }

        public StoreItem GetItem(string mailboxId, string itemId)
        {
            long? k = FindItemKey(mailboxId, itemId);
            return k.HasValue ? GetItem(k.Value) : null;
        }

        public List<StoreItem> GetPendingItems(string mailboxId, int limit)
        {
            return GetPendingItems(mailboxId, limit, 0);
        }

        /// <summary>Pending items of a mailbox in key order, after <paramref name="afterKey"/> (cursor for a single producer feeding the workers).</summary>
        public List<StoreItem> GetPendingItems(string mailboxId, int limit, long afterKey)
        {
            EnterOp(false);
            try
            {
                int mk = MailboxKey(mailboxId);
                var list = new List<StoreItem>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, ItemSelect + "WHERE i.mailbox_key=@m AND i.state=0 AND i.item_key>@a ORDER BY i.item_key LIMIT @n", null,
                    "@m", mk, "@a", afterKey, "@n", limit))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(ReadItem(r));
                return list;
            }
            finally { ExitOp(); }
        }

        public StoreItemPage GetItems(long folderKey, int page, int pageSize, StoreItemSort sort, bool descending)
        {
            return GetItems(folderKey, page, pageSize, sort, descending, null);
        }

        /// <summary>One page (0-based) of the items of a folder, sorted by date / size / subject (ties by key: stable pages).</summary>
        public StoreItemPage GetItems(long folderKey, int page, int pageSize, StoreItemSort sort, bool descending, StoreItemState? stateFilter)
        {
            if (page < 0) throw new ArgumentOutOfRangeException("page");
            if (pageSize < 1) throw new ArgumentOutOfRangeException("pageSize");
            EnterOp(false);
            try
            {
                string col;
                switch (sort)
                {
                    case StoreItemSort.Size: col = "i.size"; break;
                    case StoreItemSort.Subject: col = "i.subject COLLATE NOCASE"; break;
                    default: col = "i.received_utc"; break;
                }
                string dir = descending ? " DESC" : "";
                string where = "WHERE i.folder_key=@f" + (stateFilter.HasValue ? " AND i.state=@s" : "");
                var res = new StoreItemPage { Page = page, PageSize = pageSize, Items = new List<StoreItem>() };
                using (var l = Reader())
                {
                    using (var cmd = Cmd(l.Conn, "SELECT COUNT(*) FROM item i " + where, null, "@f", folderKey, "@s", stateFilter.HasValue ? (object)(int)stateFilter.Value : null))
                        res.Total = Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                    using (var cmd = Cmd(l.Conn, ItemSelect + where + " ORDER BY " + col + dir + ", i.item_key" + dir + " LIMIT @n OFFSET @o", null,
                        "@f", folderKey, "@s", stateFilter.HasValue ? (object)(int)stateFilter.Value : null, "@n", pageSize, "@o", (long)page * pageSize))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read()) res.Items.Add(ReadItem(r));
                }
                return res;
            }
            finally { ExitOp(); }
        }

        /// <summary>Folder tree of a mailbox with the aggregates (own and recursive) of counts and sizes.</summary>
        public List<StoreFolderNode> GetFolderTree(string mailboxId)
        {
            EnterOp(false);
            try
            {
                int mk = MailboxKey(mailboxId);
                var byKey = new Dictionary<long, StoreFolderNode>();
                var byId = new Dictionary<string, StoreFolderNode>(StringComparer.Ordinal);
                using (var l = Reader())
                {
                    using (var cmd = Cmd(l.Conn, "SELECT folder_key, folder_id, parent_folder_id, display_name, path, folder_type, source_item_count, source_size " +
                        "FROM folder WHERE mailbox_key=@m", null, "@m", mk))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                        {
                            var n = new StoreFolderNode
                            {
                                FolderKey = r.GetInt64(0), FolderId = r.GetString(1), ParentFolderId = Str(r, 2), DisplayName = Str(r, 3), Path = Str(r, 4),
                                FolderType = Str(r, 5), SourceItemCount = Lng(r, 6), SourceSize = Lng(r, 7),
                                Own = new StoreAggregate(), Total = new StoreAggregate(), Children = new List<StoreFolderNode>()
                            };
                            byKey[n.FolderKey] = n; byId[n.FolderId] = n;
                        }
                    // The items of the folders of this mailbox: the ones recorded in it, and (auto-expanding archive) the ones that live in an
                    // auxiliary archive while their folder is recorded here once, in the main archive tree. Only a store that has an auxiliary
                    // archive pays for the second part.
                    bool hasAux;
                    using (var cmd = Cmd(l.Conn, "SELECT COUNT(*) FROM mailbox WHERE kind=2", null))
                        hasAux = Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) > 0;
                    string aggregateSql = "SELECT folder_key, state, COUNT(*), COALESCE(SUM(size),0), COALESCE(SUM(pack_length),0) FROM item " +
                        "WHERE mailbox_key=@m AND folder_key IS NOT NULL GROUP BY folder_key, state";
                    if (hasAux)
                        aggregateSql = "SELECT folder_key, state, SUM(c), SUM(s), SUM(p) FROM (" +
                            "SELECT folder_key, state, COUNT(*) c, COALESCE(SUM(size),0) s, COALESCE(SUM(pack_length),0) p FROM item WHERE mailbox_key=@m AND folder_key IS NOT NULL GROUP BY folder_key, state " +
                            "UNION ALL " +
                            "SELECT folder_key, state, COUNT(*), COALESCE(SUM(size),0), COALESCE(SUM(pack_length),0) FROM item WHERE mailbox_key<>@m AND folder_key IN (SELECT folder_key FROM folder WHERE mailbox_key=@m) GROUP BY folder_key, state" +
                            ") GROUP BY folder_key, state";
                    using (var cmd = Cmd(l.Conn, aggregateSql, null, "@m", mk))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                        {
                            StoreFolderNode n;
                            if (!byKey.TryGetValue(r.GetInt64(0), out n)) continue;
                            long cnt = r.GetInt64(2);
                            n.Own.Count += cnt; n.Own.Size += r.GetInt64(3);
                            switch ((StoreItemState)r.GetInt32(1))
                            {
                                case StoreItemState.Pending: n.Own.Pending += cnt; break;
                                case StoreItemState.Exported: n.Own.Exported += cnt; n.Own.StoredBytes += r.GetInt64(4); break;
                                case StoreItemState.Failed: n.Own.Failed += cnt; break;
                                case StoreItemState.Skipped: n.Own.Skipped += cnt; break;
                            }
                        }
                }
                var roots = new List<StoreFolderNode>();
                foreach (var n in byKey.Values)
                {
                    StoreFolderNode p;
                    if (n.ParentFolderId != null && n.ParentFolderId != n.FolderId && byId.TryGetValue(n.ParentFolderId, out p)) p.Children.Add(n);
                    else roots.Add(n);
                }
                // Nodes caught in a parent cycle are unreachable from the roots: promote them to roots.
                var reach = new HashSet<long>();
                var stack = new Stack<StoreFolderNode>(roots);
                while (stack.Count > 0) { var n = stack.Pop(); if (reach.Add(n.FolderKey)) foreach (var ch in n.Children) stack.Push(ch); }
                foreach (var n in byKey.Values.OrderBy(x => x.FolderKey))
                    if (!reach.Contains(n.FolderKey))
                    {
                        StoreFolderNode p;
                        if (byId.TryGetValue(n.ParentFolderId ?? "", out p)) p.Children.Remove(n);
                        roots.Add(n);
                        var st = new Stack<StoreFolderNode>(); st.Push(n);
                        while (st.Count > 0) { var x = st.Pop(); if (reach.Add(x.FolderKey)) foreach (var ch in x.Children) st.Push(ch); }
                    }
                Comparison<StoreFolderNode> byName = (a, b) => string.Compare(a.DisplayName, b.DisplayName, StringComparison.OrdinalIgnoreCase);
                roots.Sort(byName);
                foreach (var r in roots) Finish(r, 0, byName, new HashSet<long>());
                return roots;
            }
            finally { ExitOp(); }
        }

        private static void Finish(StoreFolderNode n, int depth, Comparison<StoreFolderNode> byName, HashSet<long> path)
        {
            n.Depth = depth;
            n.Total = new StoreAggregate();
            n.Total.Add(n.Own);
            n.Children.Sort(byName);
            path.Add(n.FolderKey);
            foreach (var ch in n.Children)
                if (!path.Contains(ch.FolderKey)) { Finish(ch, depth + 1, byName, path); n.Total.Add(ch.Total); }
            path.Remove(n.FolderKey);
        }

        #endregion

        #region Item state (failed / skipped / requeue)

        public void MarkItemFailed(long itemKey, string error)
        {
            SetItemState(itemKey, StoreItemState.Failed, error, true);
        }

        public void MarkItemSkipped(long itemKey, string reason)
        {
            SetItemState(itemKey, StoreItemState.Skipped, reason, false);
        }

        private void SetItemState(long itemKey, StoreItemState state, string message, bool countAttempt)
        {
            EnterOp(true);
            try
            {
                lock (_dbLock)
                using (var cmd = Cmd(W(), "UPDATE item SET state=@s, last_error=@e, attempts=attempts+@a WHERE item_key=@k AND state<>1", null,
                    "@s", (int)state, "@e", Clip(message, 500), "@a", countAttempt ? 1 : 0, "@k", itemKey))
                    if (cmd.ExecuteNonQuery() == 0) throw new InvalidOperationException("Item " + itemKey + " not found or already exported.");
            }
            finally { ExitOp(); }
        }

        /// <summary>Records the change key the service returned with the exported bytes (an item modified after its listing).</summary>
        public void UpdateChangeKey(long itemKey, string changeKey)
        {
            if (string.IsNullOrEmpty(changeKey)) return;
            EnterOp(true);
            try
            {
                lock (_dbLock)
                using (var cmd = Cmd(W(), "UPDATE item SET change_key=@c WHERE item_key=@k", null, "@c", changeKey, "@k", itemKey))
                    cmd.ExecuteNonQuery();
            }
            finally { ExitOp(); }
        }

        /// <summary>Items of one mailbox by state: count, announced size, size stored in the packs (every state is returned).</summary>
        public StoreStateStat[] GetItemStats(string mailboxId)
        {
            EnterOp(false);
            try
            {
                int mk = MailboxKey(mailboxId);
                var counts = new Dictionary<int, StoreStateStat>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT state, COUNT(*), COALESCE(SUM(size),0), COALESCE(SUM(pack_length),0) FROM item WHERE mailbox_key=@m GROUP BY state", null, "@m", mk))
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        counts[r.GetInt32(0)] = new StoreStateStat { State = ((StoreItemState)r.GetInt32(0)).ToString(), Count = r.GetInt64(1), AnnouncedBytes = r.GetInt64(2), StoredBytes = r.GetInt64(3) };
                return Enum.GetValues(typeof(StoreItemState)).Cast<StoreItemState>()
                    .Select(s => counts.ContainsKey((int)s) ? counts[(int)s] : new StoreStateStat { State = s.ToString() }).ToArray();
            }
            finally { ExitOp(); }
        }

        /// <summary>Pending items of one mailbox per folder (key of the folder row, -1 for an item without a folder): count and announced bytes.</summary>
        public List<StoreFolderCount> GetPendingByFolder(string mailboxId)
        {
            EnterOp(false);
            try
            {
                int mk = MailboxKey(mailboxId);
                var list = new List<StoreFolderCount>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT COALESCE(folder_key,-1), COUNT(*), COALESCE(SUM(size),0) FROM item WHERE mailbox_key=@m AND state=0 GROUP BY folder_key", null, "@m", mk))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(new StoreFolderCount { FolderKey = r.GetInt64(0), Count = r.GetInt64(1), Bytes = r.GetInt64(2) });
                return list;
            }
            finally { ExitOp(); }
        }

        /// <summary>Failed (attempts &lt; maxAttempts) items of the mailbox go back to Pending. Returns their number.</summary>
        public int RequeueFailed(string mailboxId, int maxAttempts)
        {
            EnterOp(true);
            try
            {
                int mk = MailboxKey(mailboxId);
                lock (_dbLock)
                using (var cmd = Cmd(W(), "UPDATE item SET state=0 WHERE mailbox_key=@m AND state=2 AND attempts<@a", null, "@m", mk, "@a", maxAttempts))
                    return cmd.ExecuteNonQuery();
            }
            finally { ExitOp(); }
        }

        #endregion

        #region Packs: writing items

        private StorePackWriter AcquirePack()
        {
            StorePackWriter w = null;
            int reuse = 0;
            lock (_poolLock)
            {
                if (_idle.Count > 0) return _idle.Pop();
                if (_reusable.Count > 0) reuse = _reusable.Dequeue();
            }
            int packId = reuse;
            if (packId == 0)
            {
                lock (_dbLock)
                {
                    var c = W();
                    using (var cmd = Cmd(c, "INSERT INTO pack(created_utc, sealed) VALUES(@t, 0)", null, "@t", NowMs())) cmd.ExecuteNonQuery();
                    using (var cmd = Cmd(c, "SELECT last_insert_rowid()", null)) packId = Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                }
            }
            try
            {
                var fs = new FileStream(GetPackPath(packId), FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read, 64 * 1024, FileOptions.None);
                w = new StorePackWriter { PackId = packId, Stream = fs, Buffer = new byte[_opt.CopyBufferSize] };
                w.End = fs.Length;
                fs.Seek(w.End, SeekOrigin.Begin);
            }
            catch
            {
                if (w != null && w.Stream != null) w.Stream.Dispose();
                if (reuse != 0) lock (_poolLock) _reusable.Enqueue(reuse);
                throw;
            }
            lock (_poolLock)
            {
                if (_closed) { w.Stream.Dispose(); throw new ObjectDisposedException("ExportStore"); }
                _allWriters.Add(w);
            }
            return w;
        }

        private void ReleasePack(StorePackWriter w)
        {
            if (w.Closed) return;
            if (w.Dirty)
            {
                // The file may hold bytes after the committed end: do not reuse it in this session (Open truncates it).
                CloseWriter(w);
                return;
            }
            if (w.End >= _opt.PackTargetSize)
            {
                try
                {
                    lock (_dbLock)
                        if (_w != null)
                            using (var cmd = Cmd(_w, "UPDATE pack SET sealed=1 WHERE pack_id=@p", null, "@p", w.PackId)) cmd.ExecuteNonQuery();
                }
                catch { }
                CloseWriter(w);
                return;
            }
            lock (_poolLock)
            {
                if (_closed || w.Closed) return;
                _idle.Push(w);
            }
        }

        private void CloseWriter(StorePackWriter w)
        {
            lock (_poolLock)
            {
                w.Closed = true;
                _allWriters.Remove(w);
            }
            try { w.Stream.Dispose(); } catch { }
        }

        private static void RollbackPack(StorePackWriter w, long start)
        {
            try
            {
                if (w.Closed) { w.Dirty = true; return; }
                w.Stream.SetLength(start);
                w.Stream.Seek(start, SeekOrigin.Begin);
                w.End = start;
            }
            catch { w.Dirty = true; }
        }

        private void CheckWritable(long itemKey, bool overwrite)
        {
            lock (_dbLock)
            using (var cmd = Cmd(W(), "SELECT state FROM item WHERE item_key=@k", null, "@k", itemKey))
            {
                object o = cmd.ExecuteScalar();
                if (o == null) throw new KeyNotFoundException("Unknown item key " + itemKey);
                if (Convert.ToInt32(o, CultureInfo.InvariantCulture) == (int)StoreItemState.Exported && !overwrite)
                    throw new InvalidOperationException("Item " + itemKey + " is already exported (use the overwrite overload to replace it).");
            }
        }

        private StoreWriteResult Commit(long itemKey, StorePackWriter w, long start, long total, byte[] sha)
        {
            // 1. bytes on disk
            w.Stream.Flush(_opt.DurableFlush);
            // 2. index row (the item is "Exported" only from here)
            lock (_dbLock)
            using (var cmd = Cmd(W(), "UPDATE item SET state=1, pack_id=@p, pack_offset=@o, pack_length=@l, sha256=@h, exported_utc=@t, attempts=attempts+1, last_error=NULL WHERE item_key=@k", null,
                "@p", w.PackId, "@o", start, "@l", total, "@h", sha, "@t", NowMs(), "@k", itemKey))
            {
                if (cmd.ExecuteNonQuery() != 1) throw new KeyNotFoundException("Unknown item key " + itemKey);
            }
            w.End = start + total;
            return new StoreWriteResult { ItemKey = itemKey, Length = total, Sha256 = Hex(sha), PackId = w.PackId, Offset = start };
        }

        public StoreWriteResult WriteItem(long itemKey, Stream source, CancellationToken ct)
        {
            return WriteItem(itemKey, source, ct, false);
        }

        public StoreWriteResult WriteItem(string mailboxId, string itemId, Stream source, CancellationToken ct)
        {
            long? k = FindItemKey(mailboxId, itemId);
            if (!k.HasValue) throw new KeyNotFoundException("Unknown item " + itemId);
            return WriteItem(k.Value, source, ct, false);
        }

        /// <summary>
        /// Streams <paramref name="source"/> into a pack (incremental SHA-256, constant memory), flushes it and commits the index row.
        /// On any error (source, cancellation, disk) the pack is truncated back and the item stays as it was.
        /// </summary>
        public StoreWriteResult WriteItem(long itemKey, Stream source, CancellationToken ct, bool overwrite)
        {
            if (source == null) throw new ArgumentNullException("source");
            EnterOp(true);
            StorePackWriter w = null;
            long start = 0;
            bool committed = false;
            try
            {
                CheckWritable(itemKey, overwrite);
                w = AcquirePack();
                start = w.End;
                using (var h = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
                {
                    long total = 0;
                    int n;
                    byte[] buf = w.Buffer;
                    while ((n = source.Read(buf, 0, buf.Length)) > 0)
                    {
                        ct.ThrowIfCancellationRequested();
                        h.AppendData(buf, 0, n);
                        w.Stream.Write(buf, 0, n);
                        total += n;
                    }
                    var res = Commit(itemKey, w, start, total, h.GetHashAndReset());
                    committed = true;
                    return res;
                }
            }
            catch
            {
                if (w != null && !committed) RollbackPack(w, start);
                throw;
            }
            finally
            {
                if (w != null) ReleasePack(w);
                ExitOp();
            }
        }

        public Task<StoreWriteResult> WriteItemAsync(long itemKey, Stream source, CancellationToken ct)
        {
            return WriteItemAsync(itemKey, source, ct, false);
        }

        /// <summary>Asynchronous variant of WriteItem (network sources).</summary>
        public async Task<StoreWriteResult> WriteItemAsync(long itemKey, Stream source, CancellationToken ct, bool overwrite)
        {
            if (source == null) throw new ArgumentNullException("source");
            EnterOp(true);
            StorePackWriter w = null;
            long start = 0;
            bool committed = false;
            try
            {
                CheckWritable(itemKey, overwrite);
                w = AcquirePack();
                start = w.End;
                using (var h = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
                {
                    long total = 0;
                    int n;
                    byte[] buf = w.Buffer;
                    while ((n = await source.ReadAsync(buf, 0, buf.Length, ct).ConfigureAwait(false)) > 0)
                    {
                        h.AppendData(buf, 0, n);
                        await w.Stream.WriteAsync(buf, 0, n, ct).ConfigureAwait(false);
                        total += n;
                    }
                    var res = Commit(itemKey, w, start, total, h.GetHashAndReset());
                    committed = true;
                    return res;
                }
            }
            catch
            {
                if (w != null && !committed) RollbackPack(w, start);
                throw;
            }
            finally
            {
                if (w != null) ReleasePack(w);
                ExitOp();
            }
        }

        #endregion

        #region Reading items

        private void Locate(long itemKey, out int packId, out long offset, out long length, out byte[] sha)
        {
            using (var l = Reader())
            using (var cmd = Cmd(l.Conn, "SELECT state, pack_id, pack_offset, pack_length, sha256 FROM item WHERE item_key=@k", null, "@k", itemKey))
            using (var r = cmd.ExecuteReader())
            {
                if (!r.Read()) throw new KeyNotFoundException("Unknown item key " + itemKey);
                if (r.GetInt32(0) != (int)StoreItemState.Exported || r.IsDBNull(1))
                    throw new InvalidOperationException("Item " + itemKey + " is not exported.");
                packId = r.GetInt32(1); offset = r.GetInt64(2); length = r.GetInt64(3);
                sha = r.IsDBNull(4) ? null : r.GetFieldValue<byte[]>(4);
            }
        }

        /// <summary>Read-only seekable stream limited to the bytes of the item. Dispose it to release the file.</summary>
        public Stream OpenItem(long itemKey)
        {
            EnterOp(false);
            try
            {
                int packId; long offset, length; byte[] sha;
                Locate(itemKey, out packId, out offset, out length, out sha);
                return OpenRange(packId, offset, length);
            }
            finally { ExitOp(); }
        }

        public Stream OpenItem(string mailboxId, string itemId)
        {
            long? k = FindItemKey(mailboxId, itemId);
            if (!k.HasValue) throw new KeyNotFoundException("Unknown item " + itemId);
            return OpenItem(k.Value);
        }

        private Stream OpenRange(int packId, long offset, long length)
        {
            int bufSize = (int)Math.Min(64 * 1024, Math.Max(4096, length));
            var fs = new FileStream(GetPackPath(packId), FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, bufSize, FileOptions.SequentialScan);
            try
            {
                if (offset + length > fs.Length) throw new InvalidDataException("Pack " + packId + " is shorter than the catalog says.");
                return new StoreBoundedStream(fs, offset, length);
            }
            catch { fs.Dispose(); throw; }
        }

        /// <summary>Re-reads the item in streaming and compares the length and the SHA-256 with the catalog.</summary>
        public StoreVerifyResult VerifyItem(long itemKey)
        {
            var res = new StoreVerifyResult { ItemKey = itemKey };
            EnterOp(false);
            try
            {
                int packId; long offset, length; byte[] sha;
                try { Locate(itemKey, out packId, out offset, out length, out sha); }
                catch (InvalidOperationException ex) { res.Reason = ex.Message; return res; }
                res.Length = length;
                res.ExpectedSha256 = sha == null ? null : Hex(sha);
                if (!File.Exists(GetPackPath(packId))) { res.Reason = "Pack file missing."; return res; }
                byte[] buf = ArrayPool<byte>.Shared.Rent(256 * 1024);
                try
                {
                using (var s = OpenRange(packId, offset, length))
                using (var h = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
                {
                    long total = 0; int n;
                    while ((n = s.Read(buf, 0, buf.Length)) > 0) { h.AppendData(buf, 0, n); total += n; }
                    res.ActualSha256 = Hex(h.GetHashAndReset());
                    if (total != length) res.Reason = "Short read (" + total + " of " + length + " bytes).";
                    else if (sha == null || !string.Equals(res.ActualSha256, res.ExpectedSha256, StringComparison.Ordinal)) res.Reason = "SHA-256 mismatch.";
                    else res.Ok = true;
                }
                }
                finally { ArrayPool<byte>.Shared.Return(buf); }
            }
            catch (InvalidDataException ex) { res.Reason = ex.Message; }
            catch (IOException ex) { res.Reason = ex.Message; }
            finally { ExitOp(); }
            return res;
        }

        /// <summary>Verifies all Exported items (streaming); returns the failures only.</summary>
        public List<StoreVerifyResult> VerifyAll(CancellationToken ct)
        {
            var bad = new List<StoreVerifyResult>();
            long after = 0;
            while (true)
            {
                var keys = new List<long>();
                EnterOp(false);
                try
                {
                    using (var l = Reader())
                    using (var cmd = Cmd(l.Conn, "SELECT item_key FROM item WHERE state=1 AND item_key>@a ORDER BY item_key LIMIT 5000", null, "@a", after))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read()) keys.Add(r.GetInt64(0));
                }
                finally { ExitOp(); }
                if (keys.Count == 0) break;
                foreach (long k in keys)
                {
                    ct.ThrowIfCancellationRequested();
                    var v = VerifyItem(k);
                    if (!v.Ok) bad.Add(v);
                }
                after = keys[keys.Count - 1];
            }
            return bad;
        }

        #endregion

        #region Previews

        public void SetPreview(long itemKey, StorePreview p)
        {
            if (p == null) throw new ArgumentNullException("p");
            p.ItemKey = itemKey;
            SetPreviews(new[] { p });
        }

        /// <summary>Writes previews (replaces existing ones) and the sort columns (subject, date) of the items. Uses p.ItemKey.</summary>
        public int SetPreviews(IEnumerable<StorePreview> previews)
        {
            if (previews == null) throw new ArgumentNullException("previews");
            EnterOp(true);
            try
            {
                int n = 0;
                foreach (var chunk in Chunks(previews, _opt.UpsertBatchSize))
                {
                    lock (_dbLock)
                    {
                        var c = W();
                        using (var tx = c.BeginTransaction())
                        {
                            using (var ins = Cmd(c, "INSERT OR REPLACE INTO preview(item_key,subject,from_text,to_text,received_utc,sent_utc,message_class,has_attachments,attachment_count,attachment_names,body_preview) " +
                                "VALUES(@k,@s,@f,@t,@r,@e,@m,@h,@c,@a,@b)", tx))
                            using (var upd = Cmd(c, "UPDATE item SET subject=@s, received_utc=@r WHERE item_key=@k", tx))
                            {
                                var ik = Param(ins, "@k"); var is_ = Param(ins, "@s"); var ifr = Param(ins, "@f"); var it = Param(ins, "@t");
                                var ir = Param(ins, "@r"); var ie = Param(ins, "@e"); var im = Param(ins, "@m"); var ih = Param(ins, "@h");
                                var ic = Param(ins, "@c"); var ia = Param(ins, "@a"); var ib = Param(ins, "@b");
                                var uk = Param(upd, "@k"); var us = Param(upd, "@s"); var ur = Param(upd, "@r");
                                foreach (var p in chunk)
                                {
                                    string subject = Clip(p.Subject, 512);
                                    object received = ToMs(p.ReceivedUtc);
                                    ik.Value = p.ItemKey; is_.Value = Db(subject); ifr.Value = Db(Clip(p.From, 512)); it.Value = Db(Clip(p.To, 2000));
                                    ir.Value = received; ie.Value = ToMs(p.SentUtc); im.Value = Db(Clip(p.MessageClass, 128));
                                    ih.Value = p.HasAttachments ? 1 : 0; ic.Value = p.AttachmentCount; ia.Value = Db(Clip(p.AttachmentNames, 2000));
                                    ib.Value = Db(TruncateUtf8(p.BodyPreview, MaxPreviewBodyBytes));
                                    ins.ExecuteNonQuery();
                                    uk.Value = p.ItemKey; us.Value = Db(subject); ur.Value = received;
                                    upd.ExecuteNonQuery();
                                    n++;
                                }
                            }
                            tx.Commit();
                        }
                    }
                }
                return n;
            }
            finally { ExitOp(); }
        }

        public StorePreview GetPreview(long itemKey)
        {
            EnterOp(false);
            try
            {
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT subject,from_text,to_text,received_utc,sent_utc,message_class,has_attachments,attachment_count,attachment_names,body_preview " +
                    "FROM preview WHERE item_key=@k", null, "@k", itemKey))
                using (var r = cmd.ExecuteReader())
                {
                    if (!r.Read()) return null;
                    return new StorePreview
                    {
                        ItemKey = itemKey, Subject = Str(r, 0), From = Str(r, 1), To = Str(r, 2), ReceivedUtc = FromMs(r, 3), SentUtc = FromMs(r, 4),
                        MessageClass = Str(r, 5), HasAttachments = r.GetInt32(6) != 0, AttachmentCount = r.GetInt32(7), AttachmentNames = Str(r, 8), BodyPreview = Str(r, 9)
                    };
                }
            }
            finally { ExitOp(); }
        }

        #endregion

        #region Import jobs

        public long CreateImportJob(string name, string targetMailbox)
        {
            if (string.IsNullOrEmpty(targetMailbox)) throw new ArgumentException("targetMailbox");
            EnterOp(true);
            try
            {
                lock (_dbLock)
                {
                    var c = W();
                    using (var cmd = Cmd(c, "INSERT INTO importJob(name,target_mailbox,created_utc) VALUES(@n,@t,@c)", null, "@n", name, "@t", targetMailbox, "@c", NowMs()))
                        cmd.ExecuteNonQuery();
                    using (var cmd = Cmd(c, "SELECT last_insert_rowid()", null)) return Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                }
            }
            finally { ExitOp(); }
        }

        public StoreImportJob GetImportJob(long jobId)
        {
            EnterOp(false);
            try
            {
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT job_id,name,target_mailbox,created_utc FROM importJob WHERE job_id=@j", null, "@j", jobId))
                using (var r = cmd.ExecuteReader())
                    return r.Read() ? new StoreImportJob { JobId = r.GetInt64(0), Name = Str(r, 1), TargetMailbox = r.GetString(2), CreatedUtc = FromMs(r, 3).Value } : null;
            }
            finally { ExitOp(); }
        }

        /// <summary>The import jobs of the store, oldest first.</summary>
        public List<StoreImportJob> GetImportJobs()
        {
            EnterOp(false);
            try
            {
                var list = new List<StoreImportJob>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT job_id,name,target_mailbox,created_utc FROM importJob ORDER BY job_id", null))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(new StoreImportJob { JobId = r.GetInt64(0), Name = Str(r, 1), TargetMailbox = r.GetString(2), CreatedUtc = FromMs(r, 3).Value });
                return list;
            }
            finally { ExitOp(); }
        }

        /// <summary>The most recent import job of that name, or null: a second run with the same name resumes it.</summary>
        public StoreImportJob FindImportJob(string name)
        {
            StoreImportJob found = null;
            foreach (var j in GetImportJobs()) if (string.Equals(j.Name, name, StringComparison.Ordinal)) found = j;
            return found;
        }

        /// <summary>Adds items to a job (only Exported items are accepted; duplicates are ignored). Returns the number added.</summary>
        public int AddImportItems(long jobId, IEnumerable<StoreImportPlan> plan)
        {
            if (plan == null) throw new ArgumentNullException("plan");
            EnterOp(true);
            try
            {
                int n = 0;
                foreach (var chunk in Chunks(plan, _opt.UpsertBatchSize))
                {
                    lock (_dbLock)
                    {
                        var c = W();
                        using (var tx = c.BeginTransaction())
                        {
                            using (var ins = Cmd(c, "INSERT OR IGNORE INTO importItem(job_id,item_key,target_mailbox,target_folder,state,updated_utc) " +
                                "SELECT @j, item_key, @m, @f, 0, @t FROM item WHERE item_key=@k AND state=1", tx))
                            {
                                var pj = Param(ins, "@j"); var pm = Param(ins, "@m"); var pf = Param(ins, "@f"); var pt = Param(ins, "@t"); var pk = Param(ins, "@k");
                                long now = NowMs();
                                foreach (var e in chunk)
                                {
                                    pj.Value = jobId; pm.Value = Db(e.TargetMailbox); pf.Value = Db(e.TargetFolder); pt.Value = now; pk.Value = e.ItemKey;
                                    n += ins.ExecuteNonQuery();
                                }
                            }
                            tx.Commit();
                        }
                    }
                }
                return n;
            }
            finally { ExitOp(); }
        }

        private const string ImportSelect =
            "SELECT ii.import_key, ii.job_id, ii.item_key, i.item_id, m.mailbox_id, COALESCE(ii.target_mailbox, j.target_mailbox), ii.target_folder, ii.state, " +
            "ii.target_item_id, ii.attempts, ii.last_error, i.pack_length, ii.updated_utc " +
            "FROM importItem ii JOIN importJob j ON j.job_id=ii.job_id JOIN item i ON i.item_key=ii.item_key JOIN mailbox m ON m.mailbox_key=i.mailbox_key ";

        private static StoreImportItem ReadImport(SqliteDataReader r)
        {
            return new StoreImportItem
            {
                ImportKey = r.GetInt64(0), JobId = r.GetInt64(1), ItemKey = r.GetInt64(2), ItemId = r.GetString(3), SourceMailboxId = r.GetString(4),
                TargetMailbox = Str(r, 5), TargetFolder = Str(r, 6), State = (StoreImportState)r.GetInt32(7), TargetItemId = Str(r, 8),
                Attempts = r.GetInt32(9), LastError = Str(r, 10), Length = Lng(r, 11), UpdatedUtc = FromMs(r, 12)
            };
        }

        /// <summary>
        /// Atomically takes up to <paramref name="limit"/> Pending items of the job and marks them InFlight (attempts+1).
        /// A crash while InFlight makes them Unknown at the next Open: they are never handed out again automatically.
        /// </summary>
        public List<StoreImportItem> NextImportItems(long jobId, int limit)
        {
            EnterOp(true);
            try
            {
                var list = new List<StoreImportItem>();
                lock (_dbLock)
                {
                    var c = W();
                    bool full = SyncFull(c);
                    try
                    {
                    using (var tx = c.BeginTransaction())
                    {
                        using (var cmd = Cmd(c, ImportSelect + "WHERE ii.job_id=@j AND ii.state=0 ORDER BY ii.import_key LIMIT @n", tx, "@j", jobId, "@n", limit))
                        using (var r = cmd.ExecuteReader())
                            while (r.Read()) list.Add(ReadImport(r));
                        using (var up = Cmd(c, "UPDATE importItem SET state=1, attempts=attempts+1, updated_utc=@t WHERE import_key=@k", tx))
                        {
                            var pt = Param(up, "@t"); var pk = Param(up, "@k");
                            long now = NowMs();
                            foreach (var it in list)
                            {
                                pt.Value = now; pk.Value = it.ImportKey; up.ExecuteNonQuery();
                                it.State = StoreImportState.InFlight; it.Attempts++;
                            }
                        }
                        tx.Commit();
                    }
                    }
                    finally { SyncBack(c, full); }
                }
                return list;
            }
            finally { ExitOp(); }
        }

        private void SetImport(long importKey, StoreImportState state, string targetItemId, string error, StoreImportState? onlyFrom)
        {
            EnterOp(true);
            try
            {
                lock (_dbLock)
                {
                    var c = W();
                    bool full = SyncFull(c);
                    try
                    {
                        using (var cmd = Cmd(c, "UPDATE importItem SET state=@s, target_item_id=COALESCE(@t, target_item_id), last_error=@e, updated_utc=@u WHERE import_key=@k" +
                            (onlyFrom.HasValue ? " AND state=@f" : ""), null,
                            "@s", (int)state, "@t", string.IsNullOrEmpty(targetItemId) ? null : targetItemId, "@e", string.IsNullOrEmpty(error) ? null : Clip(error, 500), "@u", NowMs(), "@k", importKey, "@f", onlyFrom.HasValue ? (object)(int)onlyFrom.Value : null))
                            if (cmd.ExecuteNonQuery() == 0) throw new InvalidOperationException("Import item " + importKey + " not found" + (onlyFrom.HasValue ? " or not " + onlyFrom.Value : "") + ".");
                    }
                    finally { SyncBack(c, full); }
                }
            }
            finally { ExitOp(); }
        }

        public void MarkImportDone(long importKey, string targetItemId) { SetImport(importKey, StoreImportState.Done, targetItemId, null, null); }
        public void MarkImportFailed(long importKey, string error) { SetImport(importKey, StoreImportState.Failed, null, error, null); }

        /// <summary>The service may or may not have created the item (timeout, network): never replayed by itself, resolved by a person.</summary>
        public void MarkImportUnknown(long importKey, string error) { SetImport(importKey, StoreImportState.Unknown, null, error, null); }

        /// <summary>
        /// The folder of an import row turned out to live in another mailbox (auxiliary archive of an auto-expanding archive: the service
        /// answered 409 "the target session is expected in mailbox X"): the row now targets that mailbox (and the folder id it lists), so that a retry,
        /// a resumed run and the reconciliation of an Unknown row all look there.
        /// </summary>
        public void SetImportTarget(long importKey, string targetMailbox, string targetFolder)
        {
            EnterOp(true);
            try
            {
                lock (_dbLock)
                using (var cmd = Cmd(W(), "UPDATE importItem SET target_mailbox=@m, target_folder=@f WHERE import_key=@k", null, "@m", targetMailbox, "@f", targetFolder, "@k", importKey))
                    if (cmd.ExecuteNonQuery() == 0) throw new InvalidOperationException("Import item " + importKey + " not found.");
            }
            finally { ExitOp(); }
        }

        /// <summary>Remembers, for a job, that the folder (fromMailbox, fromFolder) is served by (toMailbox, toFolder): a resumed run goes straight there.</summary>
        public void SaveImportRoute(long jobId, string fromMailbox, string fromFolder, string toMailbox, string toFolder)
        {
            EnterOp(true);
            try
            {
                lock (_dbLock)
                using (var cmd = Cmd(W(), "INSERT INTO importRoute(job_id,from_mailbox,from_folder,to_mailbox,to_folder,learned_utc) VALUES(@j,@a,@b,@c,@d,@t) " +
                    "ON CONFLICT(job_id,from_mailbox,from_folder) DO UPDATE SET to_mailbox=excluded.to_mailbox, to_folder=excluded.to_folder, learned_utc=excluded.learned_utc", null,
                    "@j", jobId, "@a", fromMailbox, "@b", fromFolder, "@c", toMailbox, "@d", toFolder, "@t", NowMs()))
                    cmd.ExecuteNonQuery();
            }
            finally { ExitOp(); }
        }

        /// <summary>The redirections learned by the runs of a job: key "fromMailbox|fromFolder" -> (toMailbox, toFolder).</summary>
        public Dictionary<string, KeyValuePair<string, string>> GetImportRoutes(long jobId)
        {
            EnterOp(false);
            try
            {
                var map = new Dictionary<string, KeyValuePair<string, string>>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    using (var l = Reader())
                    using (var cmd = Cmd(l.Conn, "SELECT from_mailbox, from_folder, to_mailbox, to_folder FROM importRoute WHERE job_id=@j", null, "@j", jobId))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read()) map[r.GetString(0) + "|" + r.GetString(1)] = new KeyValuePair<string, string>(r.GetString(2), r.GetString(3));
                }
                catch (SqliteException ex) when (ex.Message.IndexOf("no such table", StringComparison.OrdinalIgnoreCase) >= 0) { }
                return map;
            }
            finally { ExitOp(); }
        }

        /// <summary>Puts an item back to Pending (failed or unknown, once the caller knows it is safe to replay).</summary>
        public void MarkImportPending(long importKey) { SetImport(importKey, StoreImportState.Pending, null, null, null); }

        /// <summary>Resolves an Unknown item after a check on the target: imported (Done) or not (Pending).</summary>
        public void ResolveImportUnknown(long importKey, bool imported, string targetItemId)
        {
            SetImport(importKey, imported ? StoreImportState.Done : StoreImportState.Pending, imported ? targetItemId : null, null, StoreImportState.Unknown);
        }

        /// <summary>Failed items (attempts &lt; maxAttempts) go back to Pending. Unknown items are NOT touched.</summary>
        public int RequeueImportFailed(long jobId, int maxAttempts)
        {
            EnterOp(true);
            try
            {
                lock (_dbLock)
                using (var cmd = Cmd(W(), "UPDATE importItem SET state=0, updated_utc=@u WHERE job_id=@j AND state=3 AND attempts<@a", null, "@u", NowMs(), "@j", jobId, "@a", maxAttempts))
                    return cmd.ExecuteNonQuery();
            }
            finally { ExitOp(); }
        }

        public List<StoreImportItem> GetImportItems(long jobId, StoreImportState state, int limit)
        {
            EnterOp(false);
            try
            {
                var list = new List<StoreImportItem>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, ImportSelect + "WHERE ii.job_id=@j AND ii.state=@s ORDER BY ii.import_key LIMIT @n", null, "@j", jobId, "@s", (int)state, "@n", limit))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(ReadImport(r));
                return list;
            }
            finally { ExitOp(); }
        }

        /// <summary>The ids of the target items that import rows are Done with (one job, or every job of the store when jobId is null): items the store already accounts for.</summary>
        public HashSet<string> GetDoneTargetItemIds(long? jobId)
        {
            EnterOp(false);
            try
            {
                var ids = new HashSet<string>(StringComparer.Ordinal);
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT target_item_id FROM importItem WHERE state=2 AND target_item_id IS NOT NULL" + (jobId.HasValue ? " AND job_id=@j" : ""), null,
                    jobId.HasValue ? new object[] { "@j", jobId.Value } : new object[0]))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) ids.Add(r.GetString(0));
                return ids;
            }
            finally { ExitOp(); }
        }

        /// <summary>The rows of an import job still Pending, per folder of the source (key of the folder row): count and announced bytes.</summary>
        public List<StoreFolderCount> GetImportPendingByFolder(long jobId)
        {
            EnterOp(false);
            try
            {
                var list = new List<StoreFolderCount>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT COALESCE(i.folder_key,-1), COUNT(*), COALESCE(SUM(i.size),0) FROM importItem ii JOIN item i ON i.item_key=ii.item_key " +
                    "WHERE ii.job_id=@j AND ii.state=0 GROUP BY i.folder_key", null, "@j", jobId))
                using (var r = cmd.ExecuteReader())
                    while (r.Read()) list.Add(new StoreFolderCount { FolderKey = r.GetInt64(0), Count = r.GetInt64(1), Bytes = r.GetInt64(2) });
                return list;
            }
            finally { ExitOp(); }
        }

        public StoreStateStat[] GetImportStats(long jobId)
        {
            EnterOp(false);
            try
            {
                var counts = new Dictionary<int, StoreStateStat>();
                using (var l = Reader())
                using (var cmd = Cmd(l.Conn, "SELECT ii.state, COUNT(*), COALESCE(SUM(i.size),0), COALESCE(SUM(i.pack_length),0) FROM importItem ii JOIN item i ON i.item_key=ii.item_key " +
                    "WHERE ii.job_id=@j GROUP BY ii.state", null, "@j", jobId))
                using (var r = cmd.ExecuteReader())
                    while (r.Read())
                        counts[r.GetInt32(0)] = new StoreStateStat { State = ((StoreImportState)r.GetInt32(0)).ToString(), Count = r.GetInt64(1), AnnouncedBytes = r.GetInt64(2), StoredBytes = r.GetInt64(3) };
                return Enum.GetValues(typeof(StoreImportState)).Cast<StoreImportState>()
                    .Select(s => counts.ContainsKey((int)s) ? counts[(int)s] : new StoreStateStat { State = s.ToString() }).ToArray();
            }
            finally { ExitOp(); }
        }

        #endregion

        #region Stats

        public StoreStats GetStats()
        {
            EnterOp(false);
            try
            {
                var st = new StoreStats();
                var byState = new Dictionary<int, StoreStateStat>();
                using (var l = Reader())
                {
                    st.Mailboxes = Scalar(l.Conn, "SELECT COUNT(*) FROM mailbox");
                    st.Folders = Scalar(l.Conn, "SELECT COUNT(*) FROM folder");
                    st.Previews = Scalar(l.Conn, "SELECT COUNT(*) FROM preview");
                    st.Packs = Scalar(l.Conn, "SELECT COUNT(*) FROM pack");
                    st.ImportJobs = Scalar(l.Conn, "SELECT COUNT(*) FROM importJob");
                    using (var cmd = Cmd(l.Conn, "SELECT state, COUNT(*), COALESCE(SUM(size),0), COALESCE(SUM(pack_length),0) FROM item GROUP BY state", null))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                            byState[r.GetInt32(0)] = new StoreStateStat { State = ((StoreItemState)r.GetInt32(0)).ToString(), Count = r.GetInt64(1), AnnouncedBytes = r.GetInt64(2), StoredBytes = r.GetInt64(3) };
                    var imp = new Dictionary<int, StoreStateStat>();
                    using (var cmd = Cmd(l.Conn, "SELECT state, COUNT(*) FROM importItem GROUP BY state", null))
                    using (var r = cmd.ExecuteReader())
                        while (r.Read())
                            imp[r.GetInt32(0)] = new StoreStateStat { State = ((StoreImportState)r.GetInt32(0)).ToString(), Count = r.GetInt64(1) };
                    st.ImportByState = Enum.GetValues(typeof(StoreImportState)).Cast<StoreImportState>()
                        .Select(s => imp.ContainsKey((int)s) ? imp[(int)s] : new StoreStateStat { State = s.ToString() }).ToArray();
                }
                st.ItemsByState = Enum.GetValues(typeof(StoreItemState)).Cast<StoreItemState>()
                    .Select(s => byState.ContainsKey((int)s) ? byState[(int)s] : new StoreStateStat { State = s.ToString() }).ToArray();
                foreach (var s in st.ItemsByState) { st.Items += s.Count; st.AnnouncedBytes += s.AnnouncedBytes; st.StoredBytes += s.StoredBytes; }
                if (Directory.Exists(_packsDir))
                    foreach (var f in new DirectoryInfo(_packsDir).EnumerateFiles("pack-*.bin")) st.PackBytesOnDisk += f.Length;
                return st;
            }
            finally { ExitOp(); }
        }

        private static long Scalar(SqliteConnection c, string sql)
        {
            using (var cmd = Cmd(c, sql, null)) return Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
        }

        #endregion
    }
}
