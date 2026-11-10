// =============================================================================
//  Mailbox Transfer - engine, part 16: EML import (own MIME conversion, or Exchange's conversion) through the import-export API
// -----------------------------------------------------------------------------
//  EmlImporter.Run   a directory tree of .eml files (or one file) -> the same folder tree below a root folder of a mailbox, create only.
//
//  Conversion (EmlConversion)
//    Own        .eml -> MimeParser -> MimeToMapi (FxMessage) -> FastTransfer stream with the header record of the envelope -> importItem.
//               No extra permission, works for every location (auxiliary archives included).
//    Exchange   Exchange does the MIME -> MAPI conversion (the workaround published by Glen Scales): the .eml is posted as a DRAFT to the Mail
//               API (POST /users/{mailbox}/messages, text/plain, base64), exportItems gives Exchange's own FastTransfer stream, this engine
//               patches PR_MESSAGE_FLAGS (0x0E07: draft bit cleared, read bit as wanted) and PR_MESSAGE_DELIVERY_TIME (0x0E06), the draft is
//               deleted, the stream goes to importItem. Needs Mail.ReadWrite on the mailbox the drafts are created in; an error names it.
//  The target is never changed: folders are found by name or created, items are created, nothing is replaced or deleted (the draft of the
//  Exchange conversion is the one thing deleted: the one this run created).
//  Resume: import-state.jsonl (one line per file: Done, Failed, Unknown); a file Done is not imported again; a file Unknown (the request may
//  have created the item: timeout, 5xx) is not imported again either, it is reported.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace MailboxTransfer
{
    public enum EmlConversion { Own, Exchange }

    public sealed class EmlImportProgress
    {
        public string Stage, Folder, File;
        public long Done, Resumed, Failed, Unknown, Total;
        public override string ToString() { return Stage + " " + File + " done " + Done + " resumed " + Resumed + " failed " + Failed + " unknown " + Unknown; }
    }

    public sealed class EmlImportEntry
    {
        public string State { get; set; }       // Done, Failed, Unknown
        public string File { get; set; }        // relative to the import path, '/'
        public string Folder { get; set; }      // path below the target root
        public string TargetRoot { get; set; }
        public string ItemId { get; set; }
        public string Subject { get; set; }
        public long Bytes { get; set; }
        public double Seconds { get; set; }
        public string Error { get; set; }
        public string[] Notes { get; set; }
        public string Utc { get; set; }
    }

    public sealed class EmlImportRun
    {
        /// <summary>A directory (searched for *.eml recursively) or one .eml file.</summary>
        public string Path;
        public MailboxApi Api;
        /// <summary>Import-export mailbox id (MBX:guid@tenant) the items are created in.</summary>
        public string MailboxId;
        /// <summary>Top-level folder of the mailbox below which everything is created; created when missing.</summary>
        public string TargetRoot;
        public bool WhatIf;
        public bool Resume = true;
        /// <summary>Messages without a read state in the manifest are imported unread.</summary>
        public bool Unread;
        public EmlConversion Conversion = EmlConversion.Own;
        public EmlExchangeConverter Exchange;
        public string StatePath;
        public CancellationToken Cancellation;
        public IProgress<EmlImportProgress> Progress;
        /// <summary>Progress per folder for the console (the folder below the target root: files to import, imported, kept from an earlier run, failed, unknown); null: not followed.</summary>
        public FolderTracker Folders;
        /// <summary>Only the folders (paths below Path, see <see cref="FolderFilter"/>) these patterns choose; empty: every folder.</summary>
        public List<string> IncludeFolders = new List<string>();
        /// <summary>The folders these patterns choose, and what is below them, are left out; an exclusion wins.</summary>
        public List<string> ExcludeFolders = new List<string>();
    }

    public sealed class EmlImportResult
    {
        public long Done, Resumed, Failed, Unknown;
        /// <summary>Files left out because their folder is not chosen (IncludeFolders / ExcludeFolders).</summary>
        public long Filtered;
        /// <summary>The patterns that matched no folder (a typo).</summary>
        public List<string> UnmatchedPatterns = new List<string>();
        public int FoldersCreated;
        public bool Cancelled, WhatIf;
        public TimeSpan Elapsed;
        public string StatePath, TargetRootId;
        public List<EmlImportEntry> Entries = new List<EmlImportEntry>();
        public bool Complete { get { return Failed == 0 && Unknown == 0 && !Cancelled; } }
    }

    /// <summary>Exchange's own MIME -> MAPI conversion, through a draft (see the header of this file).</summary>
    public sealed class EmlExchangeConverter
    {
        readonly MailboxApi _api;
        readonly string _mailApiUser;
        readonly string _exportMailboxId;
        public string TempDirectory = Path.GetTempPath();

        /// <param name="mailApiUser">The user the drafts are created for in the Mail API: a UPN, or MBX:guid@tenant.</param>
        /// <param name="exportMailboxId">The import-export mailbox id (MBX:...) the draft is exported from.</param>
        public EmlExchangeConverter(MailboxApi api, string mailApiUser, string exportMailboxId)
        {
            _api = api; _mailApiUser = mailApiUser; _exportMailboxId = exportMailboxId;
        }

        public sealed class Converted : IDisposable
        {
            public IFxStreamSource Source;
            public string TempFile;
            /// <summary>The folder made for this conversion (it holds only TempFile): removed with it.</summary>
            public string TempDirectory;
            public void Dispose()
            {
                if (TempFile != null) { try { File.Delete(TempFile); } catch (IOException) { } }
                if (TempDirectory != null) { try { Directory.Delete(TempDirectory, true); } catch (IOException) { } catch (UnauthorizedAccessException) { } }
            }
        }

        /// <summary>The MIME of the message as base64 text, written in pieces (the request body of POST /messages with Content-Type text/plain): no copy of the message as a string or as a second byte array.</summary>
        sealed class Base64Content : HttpContent
        {
            readonly byte[] _data;
            public Base64Content(byte[] data) { _data = data; Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/plain"); }
            protected override bool TryComputeLength(out long length) { length = ((long)_data.Length + 2) / 3 * 4; return true; }
            protected override async System.Threading.Tasks.Task SerializeToStreamAsync(Stream stream, System.Net.TransportContext context)
            {
                const int Raw = 3 * 16384;
                var encoded = new byte[4 * 16384];
                for (int offset = 0; offset < _data.Length; offset += Raw)
                {
                    int n = Math.Min(Raw, _data.Length - offset);
                    int consumed, written;
                    if (System.Buffers.Text.Base64.EncodeToUtf8(new ReadOnlySpan<byte>(_data, offset, n), encoded, out consumed, out written, offset + n >= _data.Length) != System.Buffers.OperationStatus.Done)
                        throw new InvalidDataException("base64 encoding failed.");
                    await stream.WriteAsync(encoded, 0, written).ConfigureAwait(false);
                }
            }
        }

        public Converted Convert(byte[] eml, DateTimeOffset received, DateTimeOffset? submit, bool read, CancellationToken ct, List<string> notes)
        {
            GraphClient g = _api.Graph;
            string draftId = null;
            bool kept = false;
            string tmp = Path.Combine(TempDirectory, "mbt-eml-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(tmp);
            try
            {
                try
                {
                    using (HttpResponseMessage resp = g.Send("POST", g.Url("/users/" + _mailApiUser + "/messages"),
                        () => new Base64Content(eml),
                        idempotent: false, ct: ct))
                    using (Stream s = resp.Content.ReadAsStreamAsync().GetAwaiter().GetResult())
                    using (JsonDocument doc = JsonDocument.Parse(s))
                        draftId = Json.Str(doc.RootElement, "id");
                }
                catch (GraphException ex) when (ex.StatusCode == 403 || ex.StatusCode == 401)
                {
                    throw new InvalidOperationException("The Exchange conversion creates a draft with the Mail API and needs the delegated permission Mail.ReadWrite on the mailbox (" + ex.Message + "). The application does not have it: use the own conversion (-MimeConversion Own).", ex);
                }
                if (string.IsNullOrEmpty(draftId)) throw new InvalidOperationException("The Mail API created no draft (no id in the answer).");

                List<ExportEntry> exported = _api.ExportItems(_exportMailboxId, new[] { draftId }, tmp, 0, ct);
                ExportEntry e = exported.FirstOrDefault();
                if (e == null || e.HasError || e.FilePath == null)
                    throw new InvalidOperationException("exportItems of the draft failed: " + (e == null ? "no entry" : e.ErrorCode + " " + e.ErrorMessage));
                FxDocument d = FxCodec.DecodeFile(e.FilePath);
                if (d.Error != null) throw new InvalidOperationException("The exported draft cannot be decoded: " + d.Error);
                Patch(d.Message, received, submit, read);
                string output = Path.Combine(tmp, "patched.fxs");
                var envelope = d.Envelope;
                using (var fs = new FileStream(output, FileMode.Create, FileAccess.Write))
                    FxStreamSource.CopyTo(new FxMessageSource(d.Message, envelope), fs, ct, null);
                // lazy values of d.Message point to the exported file: the output is complete, keep only it
                var result = new Converted { TempFile = output, TempDirectory = tmp, Source = FxStreamSource.FromFile(output) };
                notes.Add("converted by Exchange (draft + exportItems)");
                kept = true;
                return result;
            }
            finally
            {
                if (draftId != null)
                {
                    try { using (g.Send("DELETE", g.Url("/users/" + _mailApiUser + "/messages/" + Uri.EscapeDataString(draftId)), ct: ct)) { } }
                    catch (Exception ex) { notes.Add("the staging draft " + draftId + " could not be deleted: " + ex.Message); }
                }
                try
                {
                    // success: only the patched file stays (the Converted removes it and the folder); failure: nothing stays
                    if (kept) { foreach (string f in Directory.GetFiles(tmp)) if (!f.EndsWith("patched.fxs")) File.Delete(f); }
                    else Directory.Delete(tmp, true);
                }
                catch (IOException) { }
                catch (UnauthorizedAccessException) { }
            }
        }

        /// <summary>PR_MESSAGE_FLAGS: draft (UNSENT 0x8) cleared, read bit as asked; PR_MESSAGE_DELIVERY_TIME; PR_CLIENT_SUBMIT_TIME when the draft has none.</summary>
        public static void Patch(FxMessage m, DateTimeOffset received, DateTimeOffset? submit, bool read)
        {
            int flags = MapiToMime.Int(m.Properties, Pid.MessageFlags) ?? 0;
            flags = (flags & ~0x8) & ~0x1;
            if (read) flags |= 0x1;
            Replace(m, FxProperty.FromInt32(Pid.MessageFlags, flags));
            Replace(m, FxProperty.FromTime(Pid.MessageDeliveryTime, received.UtcDateTime));
            if (submit.HasValue) Replace(m, FxProperty.FromTime(Pid.ClientSubmitTime, submit.Value.UtcDateTime));
        }

        static void Replace(FxMessage m, FxProperty p)
        {
            int i = m.Properties.FindIndex(x => x.Id == p.Id && !x.IsNamed);
            if (i >= 0) m.Properties[i] = p; else m.Properties.Add(p);
        }
    }

    public static class EmlImporter
    {
        public const string StateName = "import-state.jsonl";
        static readonly JsonSerializerOptions JsonOpts = new JsonSerializerOptions { DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull };

        /// <summary>The FastTransfer source of an .eml with the own conversion (public for the tests and the comparison tools).</summary>
        public static IFxStreamSource ConvertOwn(byte[] eml, EmlImportOptions options, List<string> notes, out FxMessage message)
        {
            MimeEntity root = MimeParser.Parse(eml);
            message = MimeToMapi.Build(root, options, notes);
            return FxImportStream.Build(message);
        }

        static string Rel(string root, string file)
        {
            string r = Path.GetFullPath(file).Substring(Path.GetFullPath(root).TrimEnd('\\', '/').Length).TrimStart('\\', '/');
            return r.Replace((char)92, (char)47);
        }

        public static List<EmlImportEntry> ReadState(string path)
        {
            var l = new List<EmlImportEntry>();
            if (!File.Exists(path)) return l;
            foreach (string line in File.ReadLines(path, Encoding.UTF8))
            {
                if (string.IsNullOrWhiteSpace(line)) continue;
                try { var e = JsonSerializer.Deserialize<EmlImportEntry>(line); if (e != null) l.Add(e); } catch (JsonException) { }
            }
            return l;
        }

        public static EmlImportResult Run(EmlImportRun r)
        {
            var started = System.Diagnostics.Stopwatch.StartNew();
            bool isFile = File.Exists(r.Path);
            if (!isFile && !Directory.Exists(r.Path)) throw new FileNotFoundException("Path not found: " + r.Path);
            string root = isFile ? Path.GetDirectoryName(Path.GetFullPath(r.Path)) : Path.GetFullPath(r.Path);
            if (string.IsNullOrWhiteSpace(r.TargetRoot)) throw new ArgumentException("TargetRoot is required: the folder everything is created below.");
            if (r.Conversion == EmlConversion.Exchange && r.Exchange == null) throw new ArgumentException("The Exchange conversion needs an EmlExchangeConverter.");
            string statePath = r.StatePath ?? Path.Combine(root, StateName);
            var result = new EmlImportResult { WhatIf = r.WhatIf, StatePath = statePath };

            var files = (isFile ? new[] { Path.GetFullPath(r.Path) } : Directory.EnumerateFiles(root, "*.eml", SearchOption.AllDirectories).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray());
            // sidecar manifest of an export: read state and received time of the original
            var manifest = new Dictionary<string, EmlManifestEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (EmlManifestEntry m in EmlExporter.ReadManifest(root)) if (m.State == "Exported" && m.File != null) manifest[m.File.Replace((char)92, (char)47)] = m;

            // the last line of a file in the state decides: Done is not imported again; Unknown and InFlight (the process stopped inside the request) may exist in the mailbox:
            // they are not imported again either, they are reported (Unknown) until a person settles them (Resolve); Failed is tried again
            var done = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var carried = new Dictionary<string, EmlImportEntry>(StringComparer.OrdinalIgnoreCase);
            if (r.Resume)
            {
                var last = new Dictionary<string, EmlImportEntry>(StringComparer.OrdinalIgnoreCase);
                foreach (EmlImportEntry e in ReadState(statePath))
                    if (string.Equals(e.TargetRoot, r.MailboxId + "|" + r.TargetRoot, StringComparison.OrdinalIgnoreCase)) last[e.File] = e;
                foreach (var kv in last)
                {
                    if (kv.Value.State == "Done") done.Add(kv.Key);
                    else if (kv.Value.State == "Unknown" || kv.Value.State == "InFlight") { done.Add(kv.Key); carried[kv.Key] = kv.Value; }
                }
            }

            var folderIds = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string rootId = null;
            ImportSession session = null;
            StreamWriter state = r.WhatIf ? null : new StreamWriter(new FileStream(statePath, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false));
            // the folder name of the source (manifest) wins over the name made safe for the file system
            Func<string, string> folderOf = rel2 =>
            {
                EmlManifestEntry m2;
                manifest.TryGetValue(rel2, out m2);
                return m2 != null && m2.Folder != null ? m2.Folder : rel2.Contains("/") ? rel2.Substring(0, rel2.LastIndexOf('/')) : "";
            };
            Func<string, string> label = f => r.TargetRoot + (f.Length > 0 ? "/" + f : "");
            // the folders chosen: a file is taken when its folder (the name of the source, else the directory) is chosen
            FolderFilter filter = FolderFilter.Create(r.IncludeFolders, r.ExcludeFolders);
            if (filter != null)
            {
                int before = files.Length;
                files = files.Where(f => filter.AllowsAny(folderOf(Rel(root, f)), Rel(root, Path.GetDirectoryName(f)))).ToArray();
                result.Filtered = before - files.Length;
                result.UnmatchedPatterns = filter.Unmatched();
            }
            if (r.Folders != null)
                foreach (string file in files)
                {
                    string rel0 = Rel(root, file);
                    if (!done.Contains(rel0)) r.Folders.Expect(label(folderOf(rel0)), 1);
                }
            try
            {
                foreach (string file in files)
                {
                    if (r.Cancellation.IsCancellationRequested) { result.Cancelled = true; break; }
                    string rel = Rel(root, file);
                    if (done.Contains(rel))
                    {
                        EmlImportEntry old;
                        if (carried.TryGetValue(rel, out old))
                        {
                            result.Unknown++;
                            result.Entries.Add(new EmlImportEntry { State = "Unknown", File = rel, Folder = old.Folder, TargetRoot = old.TargetRoot, Subject = old.Subject, Bytes = old.Bytes, Error = (old.State == "InFlight" ? "the run stopped while the request was in flight" : old.Error) + " (earlier run: not imported again; check the mailbox, then settle it with Resolve-MbtEmlUnknown)", Utc = old.Utc });
                            if (r.Folders != null) r.Folders.Unknown(label(old.Folder ?? ""));
                        }
                        else result.Resumed++;
                        continue;
                    }
                    EmlManifestEntry me;
                    manifest.TryGetValue(rel, out me);
                    string folderRel = folderOf(rel);
                    var entry = new EmlImportEntry { File = rel, Folder = folderRel, TargetRoot = r.MailboxId + "|" + r.TargetRoot, Bytes = new FileInfo(file).Length };
                    var notes = new List<string>();
                    var watch = System.Diagnostics.Stopwatch.StartNew();
                    bool sent = false;        // the request that creates the item is about to leave (or has left): from there on a failure may have created it
                    bool stop = false;
                    try
                    {
                        if (r.WhatIf) { result.Done++; entry.State = "Done"; entry.Notes = new[] { "WhatIf" }; result.Entries.Add(entry); if (r.Folders != null) r.Folders.Done(label(folderRel), entry.Bytes); continue; }
                        if (rootId == null) { rootId = EnsureRoot(r, result); result.TargetRootId = rootId; }
                        string folderId = EnsureFolder(r, rootId, folderRel, folderIds, result);

                        byte[] eml = File.ReadAllBytes(file);
                        bool read = me != null && me.IsRead.HasValue ? me.IsRead.Value : !r.Unread;
                        DateTimeOffset? received = null;
                        DateTimeOffset parsed;
                        if (me != null && me.Received != null && DateTimeOffset.TryParse(me.Received, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out parsed)) received = parsed;

                        MimeEntity entity = MimeParser.Parse(eml);
                        entry.Subject = entity.HeaderDecoded("Subject");
                        DateTimeOffset when = received ?? MimeToMapi.ReceivedFromHeaders(entity) ?? MimeText.ParseDate(entity.Header("Date")) ?? new DateTimeOffset(File.GetLastWriteTimeUtc(file), TimeSpan.Zero);

                        if (session == null || (session.Expiration.HasValue && session.Expiration.Value < DateTimeOffset.UtcNow.AddMinutes(2)))
                            session = r.Api.CreateImportSession(r.MailboxId, r.Cancellation);

                        ImportResult ir;
                        // a line "InFlight" before the request: a process that dies inside it leaves the file Unknown, never silently absent
                        Action inFlight = () =>
                        {
                            var line = new EmlImportEntry { State = "InFlight", File = rel, Folder = folderRel, TargetRoot = entry.TargetRoot, Subject = entry.Subject, Bytes = entry.Bytes, Utc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture) };
                            state.WriteLine(JsonSerializer.Serialize(line, JsonOpts)); state.Flush();
                            sent = true;
                        };
                        if (r.Conversion == EmlConversion.Exchange)
                        {
                            using (EmlExchangeConverter.Converted c = r.Exchange.Convert(eml, when, MimeText.ParseDate(entity.Header("Date")), read, r.Cancellation, notes))
                            {
                                inFlight();
                                ir = r.Api.ImportItem(session, folderId, c.Source, r.Cancellation);
                            }
                        }
                        else
                        {
                            FxMessage msg = MimeToMapi.Build(entity, new EmlImportOptions { Read = read, Received = received }, notes);
                            IFxStreamSource stream = FxImportStream.Build(msg);
                            inFlight();
                            ir = r.Api.ImportItem(session, folderId, stream, r.Cancellation);
                        }
                        entry.State = "Done"; entry.ItemId = ir.ItemId; entry.Seconds = watch.Elapsed.TotalSeconds;
                        result.Done++;
                        if (r.Folders != null) r.Folders.Done(label(folderRel), entry.Bytes);
                    }
                    catch (OperationCanceledException)
                    {
                        result.Cancelled = true; stop = true;
                        if (sent)
                        {
                            // stopped inside the request: the item may exist, so the file is Unknown, not forgotten (and not imported again by the next run)
                            entry.State = "Unknown"; entry.Error = "stopped while the request was in flight: the item may exist"; result.Unknown++;
                            if (r.Folders != null) r.Folders.Unknown(label(folderRel));
                        }
                        else break;
                    }
                    catch (Exception ex)
                    {
                        // before the request left (folder, session, conversion, draft) nothing can have been created: Failed. After it: the rules of the other imports (4xx, 429, 503 and an upload cut
                        // short: nothing created; anything else, a timeout, a 5xx, an answer that cannot be read: it may exist)
                        if (sent && !ImportHelper.NothingCreated(ex))
                        {
                            entry.State = "Unknown"; entry.Error = ex.Message; result.Unknown++;
                            if (r.Folders != null) r.Folders.Unknown(label(folderRel));
                        }
                        else
                        {
                            entry.State = "Failed"; entry.Error = ex.Message; result.Failed++;
                            if (r.Folders != null) r.Folders.Fail(label(folderRel));
                        }
                    }
                    entry.Notes = entry.Notes ?? (notes.Count > 0 ? notes.ToArray() : null);
                    entry.Utc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
                    result.Entries.Add(entry);
                    if (state != null) { state.WriteLine(JsonSerializer.Serialize(entry, JsonOpts)); state.Flush(); }
                    if (r.Progress != null) r.Progress.Report(new EmlImportProgress { Stage = "item", File = rel, Folder = folderRel, Done = result.Done, Resumed = result.Resumed, Failed = result.Failed, Unknown = result.Unknown, Total = files.Length });
                    if (stop) break;
                }
            }
            finally { if (state != null) state.Dispose(); }
            started.Stop();
            result.Elapsed = started.Elapsed;
            return result;
        }

        /// <summary>
        /// Settles the files left Unknown (or InFlight) by an earlier run: a person looked in the mailbox. imported: the item is there, the file becomes Done; else it is not there and
        /// the file becomes Failed, so that the next run imports it. files: relative paths as the state lists them; null: every Unknown file. Returns the files settled.
        /// </summary>
        public static List<string> Resolve(string statePath, bool imported, IEnumerable<string> files)
        {
            var last = new Dictionary<string, EmlImportEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (EmlImportEntry e in ReadState(statePath)) last[e.File + "\n" + e.TargetRoot] = e;
            HashSet<string> want = files == null ? null : new HashSet<string>(files.Select(f => f.Replace((char)92, (char)47)), StringComparer.OrdinalIgnoreCase);
            var settled = new List<string>();
            using (var w = new StreamWriter(new FileStream(statePath, FileMode.Append, FileAccess.Write, FileShare.Read), new UTF8Encoding(false)))
                foreach (EmlImportEntry e in last.Values.Where(x => (x.State == "Unknown" || x.State == "InFlight") && (want == null || want.Contains(x.File))).ToList())
                {
                    var line = new EmlImportEntry
                    {
                        State = imported ? "Done" : "Failed", File = e.File, Folder = e.Folder, TargetRoot = e.TargetRoot, Subject = e.Subject, Bytes = e.Bytes,
                        Error = imported ? null : "settled by hand: the item is not in the mailbox", Notes = new[] { imported ? "settled by hand: the item is in the mailbox" : "settled by hand: not in the mailbox, imported again by the next run" },
                        Utc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture)
                    };
                    w.WriteLine(JsonSerializer.Serialize(line, JsonOpts)); w.Flush();
                    settled.Add(e.File);
                }
            return settled;
        }

        static string EnsureRoot(EmlImportRun r, EmlImportResult result)
        {
            FolderInfo f = r.Api.GetRootFolders(r.MailboxId, r.Cancellation).FirstOrDefault(x => string.Equals(x.DisplayName, r.TargetRoot, StringComparison.OrdinalIgnoreCase));
            if (f != null) return f.Id;
            result.FoldersCreated++;
            return r.Api.CreateFolder(r.MailboxId, r.TargetRoot, null, "IPF.Note", null, r.Cancellation).Id;
        }

        static string EnsureFolder(EmlImportRun r, string rootId, string rel, Dictionary<string, string> cache, EmlImportResult result)
        {
            if (rel.Length == 0) return rootId;
            string cur = "";
            string parentId = rootId;
            foreach (string name in rel.Split('/'))
            {
                cur = cur.Length == 0 ? name : cur + "/" + name;
                string id;
                if (!cache.TryGetValue(cur, out id))
                {
                    FolderInfo f = r.Api.GetChildFolders(r.MailboxId, parentId, r.Cancellation).FirstOrDefault(x => string.Equals(x.DisplayName, name, StringComparison.OrdinalIgnoreCase));
                    if (f == null) { f = r.Api.CreateFolder(r.MailboxId, name, parentId, "IPF.Note", cur, r.Cancellation); result.FoldersCreated++; }
                    id = f.Id;
                    cache[cur] = id;
                }
                parentId = id;
            }
            return parentId;
        }
    }
}
