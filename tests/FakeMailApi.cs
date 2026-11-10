// =============================================================================
//  Mailbox Transfer - fake Microsoft Graph Mail API (tests only, no network)
// -----------------------------------------------------------------------------
//  An HttpMessageHandler that answers /users/MBX:<mailbox>@<tenant>/mailFolders... and /messages... for a few in-memory mailboxes (the
//  routes and the shapes measured on the lab tenant): mailFolders/delta in pages (Prefer odata.maxpagesize), childFolders, the messages of a
//  folder with $top, $orderby, $expand of the extended properties (PR_MESSAGE_SIZE 0x0E08, PR_MESSAGE_CLASS 0x001A, PR_SENSITIVITY 0x0036)
//  and nextLink, one message with its text body, attachments and attached items. Everything under /admin/exchange/ (the import-export API)
//  is handed to an inner FakeMailboxService, so one client talks to both APIs as it does against Graph. It can throttle (429 with
//  Retry-After), refuse a mailbox (403), and logs each request with its Prefer header.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer.Tests
{
    public sealed class FakeMailFolder
    {
        public string Id, Name, ParentId, WellKnown;
        public bool Recoverable;
        /// <summary>Children announced in childFolderCount that the Mail API does not list (Calendar, Contacts...: not mail folders).</summary>
        public int NonMailChildren;
        /// <summary>Items of the folder that are not messages (appointments, contacts...): counted in totalItemCount, never returned by /messages.</summary>
        public int NonMessageItems;
    }

    public sealed class FakeMailAttachment
    {
        public string Id, Name, ContentType = "application/octet-stream";
        public long Size;
        public bool IsInline;
        public FakeMailMessage Item;
    }

    public sealed class FakeMailMessage
    {
        public string Id, FolderId, Subject, FromName, FromAddress, MessageClass = "IPM.Note", Importance = "normal", Flag = "notFlagged", Body = "";
        public string[] To = new string[0], Cc = new string[0], Categories = new string[0];
        public DateTime ReceivedUtc = new DateTime(2026, 10, 1, 8, 0, 0, DateTimeKind.Utc);
        public DateTime SentUtc, CreatedUtc, ModifiedUtc;
        public long Size = 1000;
        public int Sensitivity;
        public bool IsRead = true, IsDraft;
        public List<FakeMailAttachment> Attachments = new List<FakeMailAttachment>();
    }

    public sealed class FakeMailApi : HttpMessageHandler
    {
        readonly object _lock = new object();
        readonly HttpMessageInvoker _inner;
        readonly Dictionary<string, List<FakeMailFolder>> _folders = new Dictionary<string, List<FakeMailFolder>>();
        readonly Dictionary<string, List<FakeMailMessage>> _messages = new Dictionary<string, List<FakeMailMessage>>();
        int _throttleLeft;
        double _throttleRetryAfter;
        string _throttleContains;

        /// <summary>Folders of a mailbox per page of mailFolders/delta at most (the service may return fewer than the Prefer odata.maxpagesize asked).</summary>
        public int DeltaPageSize = 200;
        public readonly List<string> Log = new List<string>();
        /// <summary>Mailboxes the Mail API refuses (403 ErrorAccessDenied: no Mail.Read, or a mailbox it does not reach).</summary>
        public readonly HashSet<string> Forbidden = new HashSet<string>();
        /// <summary>"mailbox|folder id": the Mail API cannot list the messages of that folder (live 2026-10-31: a folder of the main archive whose content moved to an auxiliary archive answers 500 "Cannot query rows in a table").</summary>
        public readonly HashSet<string> MessagesNotReady = new HashSet<string>();
        public int Throttled;

        public FakeMailApi(HttpMessageHandler importExport) { _inner = new HttpMessageInvoker(importExport); }

        public void AddMailbox(string id)
        {
            lock (_lock) { _folders[id] = new List<FakeMailFolder>(); _messages[id] = new List<FakeMailMessage>(); }
        }

        public FakeMailFolder AddFolder(string mailbox, string id, string name, string parentId = null, string wellKnown = null, bool recoverable = false, int nonMailChildren = 0)
        {
            var f = new FakeMailFolder { Id = id, Name = name, ParentId = string.IsNullOrEmpty(parentId) ? "ROOT" : parentId, WellKnown = string.IsNullOrEmpty(wellKnown) ? null : wellKnown, Recoverable = recoverable, NonMailChildren = nonMailChildren };
            lock (_lock) _folders[mailbox].Add(f);
            return f;
        }

        public FakeMailMessage AddMessage(string mailbox, string folderId, string id, string subject, DateTime receivedUtc, long size = 1000, string from = "Alice <alice@contoso.test>", string[] to = null, string[] cc = null)
        {
            var m = new FakeMailMessage { Id = id, FolderId = folderId, Subject = subject, ReceivedUtc = receivedUtc, SentUtc = receivedUtc.AddMinutes(-1), CreatedUtc = receivedUtc, ModifiedUtc = receivedUtc, Size = size };
            int lt = from.IndexOf('<');
            m.FromName = lt > 0 ? from.Substring(0, lt).Trim() : "";
            m.FromAddress = lt > 0 ? from.Substring(lt + 1).TrimEnd('>') : from;
            m.To = to ?? new[] { "Bob <bob@contoso.test>" };
            m.Cc = cc ?? new string[0];
            lock (_lock) _messages[mailbox].Add(m);
            return m;
        }

        public void Throttle(int count, double retryAfterSeconds, string urlContains = null)
        {
            lock (_lock) { _throttleLeft = count; _throttleRetryAfter = retryAfterSeconds; _throttleContains = urlContains; }
        }

        public int RequestsOf(string contains) { lock (_lock) return Log.Count(l => l.Contains(contains)); }
        public string[] Requests() { lock (_lock) return Log.ToArray(); }

        static HttpResponseMessage Json(HttpStatusCode code, string json)
        {
            return new HttpResponseMessage(code) { Content = new StringContent(json, Encoding.UTF8, "application/json") };
        }

        static string Q(string s) { return JsonSerializer.Serialize(s); }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            string url = request.RequestUri.OriginalString;
            if (url.Contains("/admin/exchange/") || url.Contains("outlook.office365.com") || url.Contains("/settings/exchange"))
                return await _inner.SendAsync(request, ct).ConfigureAwait(false);
            string prefer = request.Headers.TryGetValues("Prefer", out IEnumerable<string> values) ? string.Join(",", values) : null;
            lock (_lock)
            {
                Log.Add(request.Method.Method + " " + url + (prefer == null ? "" : " [Prefer: " + prefer + "]"));
                if (_throttleLeft > 0 && (_throttleContains == null || url.Contains(_throttleContains)))
                {
                    _throttleLeft--; Throttled++;
                    var r = Json((HttpStatusCode)429, "{\"error\":{\"code\":\"TooManyRequests\",\"message\":\"slow down\"}}");
                    r.Headers.TryAddWithoutValidation("Retry-After", _throttleRetryAfter.ToString(CultureInfo.InvariantCulture));
                    return r;
                }
                return Route(url, prefer);
            }
        }

        // ---------------------------------------------------------------------------------------------

        HttpResponseMessage Route(string url, string prefer)
        {
            string decoded = Uri.UnescapeDataString(url);
            int u = decoded.IndexOf("/users/", StringComparison.Ordinal);
            if (u < 0) return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"NoRoute\",\"message\":\"" + decoded.Replace("\"", "'") + "\"}}");
            string rest = decoded.Substring(u + 7);
            int slash = rest.IndexOf('/');
            string mailbox = slash < 0 ? rest : rest.Substring(0, slash);
            string tail = slash < 0 ? "" : rest.Substring(slash);
            string path = tail.Split('?')[0];
            string query = tail.Contains("?") ? tail.Substring(tail.IndexOf('?') + 1) : "";
            if (!_folders.ContainsKey(mailbox))
                return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorInvalidUser\",\"message\":\"The requested user '" + mailbox + "' is invalid.\"}}");
            if (Forbidden.Contains(mailbox))
                return Json(HttpStatusCode.Forbidden, "{\"error\":{\"code\":\"ErrorAccessDenied\",\"message\":\"Access is denied. Check credentials and try again.\"}}");

            if (path == "/mailFolders/delta") return Delta(mailbox, url, query, prefer);
            if (path == "/mailFolders/recoverableitemsroot") return FolderJsonResponse(mailbox, _folders[mailbox].FirstOrDefault(f => f.Recoverable && f.ParentId == "ROOT"));
            if (path.StartsWith("/mailFolders/") && path.EndsWith("/childFolders")) return ChildFolders(mailbox, path);
            if (path.StartsWith("/mailFolders/") && path.IndexOf('/', "/mailFolders/".Length) < 0 && path != "/mailFolders/delta")
            {
                string one = path.Substring("/mailFolders/".Length);
                return FolderJsonResponse(mailbox, _folders[mailbox].FirstOrDefault(f => f.Id == one));
            }
            if (path.StartsWith("/mailFolders/") && path.EndsWith("/messages"))
            {
                string fid = path.Substring("/mailFolders/".Length); fid = fid.Substring(0, fid.Length - "/messages".Length);
                if (MessagesNotReady.Contains(mailbox + "|" + fid))
                    return Json(HttpStatusCode.InternalServerError, "{\"error\":{\"code\":\"ErrorInternalServerError\",\"message\":\"An internal server error occurred. The operation failed., Cannot query rows in a table.\"}}");
                return Messages(mailbox, url, path, query);
            }
            if (path.StartsWith("/messages/")) return MessageRoute(mailbox, path, query, prefer);
            return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"NoRoute\",\"message\":\"" + path + "\"}}");
        }

        string FolderJson(string mailbox, FakeMailFolder f)
        {
            int kids = _folders[mailbox].Count(x => x.ParentId == f.Id) + f.NonMailChildren;
            var msgs = _messages[mailbox].Where(m => m.FolderId == f.Id).ToList();
            return "{\"id\":" + Q(f.Id) + ",\"displayName\":" + Q(f.Name) + ",\"parentFolderId\":" + Q(f.ParentId) + ",\"childFolderCount\":" + kids +
                ",\"unreadItemCount\":" + msgs.Count(m => !m.IsRead) + ",\"totalItemCount\":" + (msgs.Count + f.NonMessageItems) + ",\"sizeInBytes\":" + msgs.Sum(m => m.Size) +
                ",\"wellKnownName\":" + (f.WellKnown == null ? "null" : Q(f.WellKnown)) + ",\"isHidden\":false}";
        }

        HttpResponseMessage FolderJsonResponse(string mailbox, FakeMailFolder f)
        {
            if (f == null) return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"folder\"}}");
            return Json(HttpStatusCode.OK, FolderJson(mailbox, f));
        }

        HttpResponseMessage Delta(string mailbox, string url, string query, string prefer)
        {
            if (query.Contains("$top=")) return Json(HttpStatusCode.BadRequest, "{\"error\":{\"code\":\"ErrorInvalidUrlQuery\",\"message\":\"The following parameters are not supported with change tracking over the 'Folders' resource: '$orderby, $filter, $search, $top'.\"}}");
            int size = DeltaPageSize;
            if (prefer != null && prefer.Contains("odata.maxpagesize=")) size = Math.Min(DeltaPageSize, int.Parse(prefer.Substring(prefer.IndexOf("odata.maxpagesize=") + 18).Split(',', ';')[0].Trim()));
            int skip = 0;
            int at = query.IndexOf("$skiptoken=", StringComparison.Ordinal);
            if (at >= 0) skip = int.Parse(query.Substring(at + 11).Split('&')[0]);
            var all = _folders[mailbox].Where(f => !f.Recoverable).ToList();
            var page = all.Skip(skip).Take(size).ToList();
            var sb = new StringBuilder("{\"value\":[" + string.Join(",", page.Select(f => FolderJson(mailbox, f))) + "]");
            string baseUrl = "https://graph.microsoft.com/v1.0/users/" + mailbox + "/mailFolders/delta";
            if (skip + size < all.Count) sb.Append(",\"@odata.nextLink\":\"" + baseUrl + "?%24skiptoken=" + (skip + size) + "\"");
            else sb.Append(",\"@odata.deltaLink\":\"" + baseUrl + "?%24deltatoken=T\"");
            sb.Append("}");
            return Json(HttpStatusCode.OK, sb.ToString());
        }

        HttpResponseMessage ChildFolders(string mailbox, string path)
        {
            string id = path.Substring("/mailFolders/".Length);
            id = id.Substring(0, id.Length - "/childFolders".Length);
            string parent = id == "msgfolderroot" ? "ROOT" : id == "recoverableitemsroot" ? _folders[mailbox].First(f => f.Recoverable && f.ParentId == "ROOT").Id : id;
            var kids = _folders[mailbox].Where(f => f.ParentId == parent && (id == "msgfolderroot" ? !f.Recoverable : true)).ToList();
            return Json(HttpStatusCode.OK, "{\"value\":[" + string.Join(",", kids.Select(f => FolderJson(mailbox, f))) + "]}");
        }

        static string Param(string query, string name)
        {
            foreach (string part in query.Split('&'))
            {
                int eq = part.IndexOf('=');
                if (eq > 0 && part.Substring(0, eq).Replace("%24", "$") == name) return part.Substring(eq + 1);
            }
            return null;
        }

        HttpResponseMessage Messages(string mailbox, string url, string path, string query)
        {
            string folderId = path.Substring("/mailFolders/".Length);
            folderId = folderId.Substring(0, folderId.Length - "/messages".Length);
            if (!_folders[mailbox].Any(f => f.Id == folderId)) return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"folder\"}}");
            int top = int.Parse(Param(query, "$top") ?? "10");
            int skip = int.Parse(Param(query, "$skip") ?? "0");
            string order = (Param(query, "$orderby") ?? "receivedDateTime desc").Replace("%20", " ");
            if (query.Contains("sensitivity,") || query.Contains(",sensitivity")) return Json(HttpStatusCode.BadRequest, "{\"error\":{\"code\":\"BadRequest\",\"message\":\"Could not find a property named 'sensitivity' on type 'Microsoft.OutlookServices.Message'.\"}}");
            IEnumerable<FakeMailMessage> all = _messages[mailbox].Where(m => m.FolderId == folderId && m.MessageClass.StartsWith("IPM.Note"));
            switch (order)
            {
                case "receivedDateTime asc": all = all.OrderBy(m => m.ReceivedUtc); break;
                case "sentDateTime desc": all = all.OrderByDescending(m => m.SentUtc); break;
                case "sentDateTime asc": all = all.OrderBy(m => m.SentUtc); break;
                case "subject asc": all = all.OrderBy(m => m.Subject, StringComparer.OrdinalIgnoreCase); break;
                case "subject desc": all = all.OrderByDescending(m => m.Subject, StringComparer.OrdinalIgnoreCase); break;
                default: all = all.OrderByDescending(m => m.ReceivedUtc); break;
            }
            var list = all.ToList();
            var page = list.Skip(skip).Take(top).ToList();
            bool recipients = query.Contains("toRecipients");
            var sb = new StringBuilder("{\"value\":[" + string.Join(",", page.Select(m => MessageJson(m, recipients, false, false))) + "]");
            if (skip + top < list.Count)
            {
                string next = url.Contains("skip=") ? System.Text.RegularExpressions.Regex.Replace(url, @"%24skip=\d+|\$skip=\d+", "%24skip=" + (skip + top)) : url + "&%24skip=" + (skip + top);
                sb.Append(",\"@odata.nextLink\":" + Q(next));
            }
            sb.Append("}");
            return Json(HttpStatusCode.OK, sb.ToString());
        }

        static string Recipients(string[] list)
        {
            return "[" + string.Join(",", list.Select(r =>
            {
                int lt = r.IndexOf('<');
                string name = lt > 0 ? r.Substring(0, lt).Trim() : "", address = lt > 0 ? r.Substring(lt + 1).TrimEnd('>') : r;
                return "{\"emailAddress\":{\"name\":" + Q(name) + ",\"address\":" + Q(address) + "}}";
            })) + "]";
        }

        static string MessageJson(FakeMailMessage m, bool recipients, bool full, bool textBody)
        {
            var sb = new StringBuilder("{\"id\":" + Q(m.Id) + ",\"subject\":" + Q(m.Subject) +
                ",\"receivedDateTime\":" + Q(m.ReceivedUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")) + ",\"sentDateTime\":" + Q(m.SentUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")) +
                ",\"createdDateTime\":" + Q(m.CreatedUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")) + ",\"lastModifiedDateTime\":" + Q(m.ModifiedUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")) +
                ",\"changeKey\":\"ck-" + m.Id + "\",\"hasAttachments\":" + (m.Attachments.Count > 0 ? "true" : "false") + ",\"importance\":" + Q(m.Importance) +
                ",\"categories\":[" + string.Join(",", m.Categories.Select(Q)) + "],\"flag\":{\"flagStatus\":" + Q(m.Flag) + "},\"isRead\":" + (m.IsRead ? "true" : "false") +
                ",\"isDraft\":" + (m.IsDraft ? "true" : "false") + ",\"internetMessageId\":" + Q("<" + m.Id + "@contoso.test>") + ",\"conversationId\":" + Q("conv-" + m.Id) +
                ",\"from\":{\"emailAddress\":{\"name\":" + Q(m.FromName) + ",\"address\":" + Q(m.FromAddress) + "}}");
            if (recipients || full) sb.Append(",\"toRecipients\":" + Recipients(m.To) + ",\"ccRecipients\":" + Recipients(m.Cc));
            if (full)
            {
                sb.Append(",\"bccRecipients\":[],\"replyTo\":[],\"bodyPreview\":" + Q(m.Body.Length > 255 ? m.Body.Substring(0, 255) : m.Body));
                sb.Append(",\"body\":{\"contentType\":" + Q(textBody ? "text" : "html") + ",\"content\":" + Q(textBody ? m.Body : "<html><body>" + m.Body + "</body></html>") + "}");
            }
            sb.Append(",\"singleValueExtendedProperties\":[{\"id\":\"Integer 0xe08\",\"value\":\"" + m.Size + "\"},{\"id\":\"String 0x1a\",\"value\":" + Q(m.MessageClass) + "},{\"id\":\"Integer 0x36\",\"value\":\"" + m.Sensitivity + "\"}]}");
            return sb.ToString();
        }

        HttpResponseMessage MessageRoute(string mailbox, string path, string query, string prefer)
        {
            string rest = path.Substring("/messages/".Length);
            string[] parts = rest.Split('/');
            FakeMailMessage m = _messages[mailbox].FirstOrDefault(x => x.Id == parts[0]);
            if (m == null) return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"The specified object was not found in the store.\"}}");
            if (parts.Length == 1) return Json(HttpStatusCode.OK, MessageJson(m, true, true, prefer != null && prefer.Contains("body-content-type=\"text\"")));
            if (parts.Length == 2 && parts[1] == "attachments")
                return Json(HttpStatusCode.OK, "{\"value\":[" + string.Join(",", m.Attachments.Select(a => AttachmentJson(a, false))) + "]}");
            if (parts.Length == 3 && parts[1] == "attachments")
            {
                FakeMailAttachment a = m.Attachments.FirstOrDefault(x => x.Id == parts[2]);
                if (a == null) return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"attachment\"}}");
                return Json(HttpStatusCode.OK, AttachmentJson(a, query.Contains("itemAttachment/item")));
            }
            return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"NoRoute\",\"message\":\"" + path + "\"}}");
        }

        static string AttachmentJson(FakeMailAttachment a, bool withItem)
        {
            string type = a.Item != null ? "#microsoft.graph.itemAttachment" : "#microsoft.graph.fileAttachment";
            string json = "{\"@odata.type\":" + Q(type) + ",\"id\":" + Q(a.Id) + ",\"name\":" + Q(a.Name) + ",\"contentType\":" + Q(a.ContentType) + ",\"size\":" + a.Size + ",\"isInline\":" + (a.IsInline ? "true" : "false");
            if (withItem && a.Item != null)
            {
                string item = MessageJson(a.Item, true, true, false);
                item = "{\"@odata.type\":\"#microsoft.graph.eventMessage\"," + item.Substring(1);
                if (a.Item.Attachments.Count > 0)
                    item = item.Substring(0, item.Length - 1) + ",\"attachments\":[" + string.Join(",", a.Item.Attachments.Select(x => AttachmentJson(x, false))) + "]}";
                json += ",\"item\":" + item;
            }
            return json + "}";
        }
    }
}
