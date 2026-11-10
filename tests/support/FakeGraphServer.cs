// A tiny in-memory stand-in for the parts of Microsoft Graph v1.0 that tools\New-LabFxCorpus.ps1 calls (folders, calendars, contact
// folders, messages / events / contacts, attachments, PATCH, recurring instances, photo). Listens on 127.0.0.1 only. Tests only.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.Json;
using System.Threading;

namespace MailboxTransfer.Tests
{
    public sealed class FakeGraphServer : IDisposable
    {
        readonly HttpListener _listener = new HttpListener();
        readonly Thread _thread;
        readonly object _lock = new object();
        int _seq;
        // collection path (lower case) -> items {id, title}
        readonly Dictionary<string, List<string[]>> _collections = new Dictionary<string, List<string[]>>();
        readonly Dictionary<string, string> _parents = new Dictionary<string, string>();   // folder id -> parent id
        readonly Dictionary<string, string> _names = new Dictionary<string, string>();     // folder id -> display name
        readonly HashSet<string> _exceptions = new HashSet<string>();
        readonly HashSet<string> _cancelled = new HashSet<string>();

        public string Url { get; private set; }
        /// <summary>"METHOD path" of every request, in order.</summary>
        public List<string> Log { get; private set; }
        /// <summary>A request whose body contains one of these texts is answered 400.</summary>
        public List<string> RefuseBodiesContaining { get; private set; }
        // Like the real Graph: a calendar is not a mail folder, GET me/mailFolders/{calendarId} is a 404 (so the path of the calendar cannot be resolved).
        public bool CalendarsAreNotMailFolders { get; set; }
        readonly HashSet<string> _calendarIds = new HashSet<string>();
        public int Requests { get { lock (_lock) return Log.Count; } }

        public FakeGraphServer()
        {
            Log = new List<string>();
            RefuseBodiesContaining = new List<string>();
            int port;
            using (var probe = new System.Net.Sockets.TcpListener(IPAddress.Loopback, 0)) { probe.Start(); port = ((IPEndPoint)probe.LocalEndpoint).Port; probe.Stop(); }
            Url = "http://127.0.0.1:" + port;
            _listener.Prefixes.Add(Url + "/");
            _listener.Start();
            _thread = new Thread(Loop) { IsBackground = true };
            _thread.Start();
        }

        public void Dispose() { try { _listener.Stop(); _listener.Close(); } catch (Exception) { } }

        public IEnumerable<string> Requested(string methodAndPrefix)
        {
            lock (_lock) return Log.Where(l => l.StartsWith(methodAndPrefix, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        void Loop()
        {
            while (_listener.IsListening)
            {
                HttpListenerContext c;
                try { c = _listener.GetContext(); } catch (Exception) { return; }
                try { Handle(c); } catch (Exception ex) { Reply(c, 500, "{\"error\":{\"code\":\"fake\",\"message\":\"" + ex.Message.Replace("\"", "'") + "\"}}"); }
            }
        }

        static void Reply(HttpListenerContext c, int status, string json)
        {
            c.Response.StatusCode = status;
            if (json != null)
            {
                byte[] b = Encoding.UTF8.GetBytes(json);
                c.Response.ContentType = "application/json";
                c.Response.OutputStream.Write(b, 0, b.Length);
            }
            c.Response.Close();
        }

        string NewId(string prefix) { return prefix + (++_seq); }

        static string Esc(string s) { return JsonEncodedText.Encode(s ?? "").ToString(); }

        void Handle(HttpListenerContext c)
        {
            string method = c.Request.HttpMethod;
            string path = Uri.UnescapeDataString(c.Request.Url.AbsolutePath);
            path = path.TrimStart((char)47); if (path.StartsWith("v1.0/")) path = path.Substring(5);
            string query = c.Request.Url.Query ?? "";
            byte[] body;
            using (var ms = new MemoryStream()) { c.Request.InputStream.CopyTo(ms); body = ms.ToArray(); }
            string text = Encoding.UTF8.GetString(body);
            string contentType = c.Request.ContentType ?? "";
            lock (_lock)
            {
                Log.Add(method + " " + path);
                if (c.Request.Headers["Authorization"] == null) { Reply(c, 401, "{\"error\":{\"code\":\"InvalidAuthenticationToken\",\"message\":\"no token\"}}"); return; }
                if (method != "GET" && RefuseBodiesContaining.Any(r => text.Contains(r)))
                { Reply(c, 400, "{\"error\":{\"code\":\"ErrorInvalidProperty\",\"message\":\"refused by the fake\"}}"); return; }
                string[] p = path.Split('/');   // me / mailFolders / {id} / ...

                // ---- folders
                if (path == "me/mailFolders/msgfolderroot") { Reply(c, 200, "{\"id\":\"root\",\"displayName\":\"Top\"}"); return; }
                if (p.Length == 3 && (p[1] == "mailFolders" || p[1] == "contactFolders") && method == "GET")
                {
                    string id = p[2];
                    if (id == "calroot") { Reply(c, 200, "{\"id\":\"calroot\",\"displayName\":\"Calendar\",\"parentFolderId\":\"root\"}"); return; }
                    if (id == "contactsroot") { Reply(c, 200, "{\"id\":\"contactsroot\",\"displayName\":\"Contacts\",\"parentFolderId\":\"root\"}"); return; }
                    if (CalendarsAreNotMailFolders && _calendarIds.Contains(id)) { Reply(c, 404, "{\"error\":{\"code\":\"ErrorInvalidIdMalformed\",\"message\":\"not a mail folder\"}}"); return; }
                    if (!_names.ContainsKey(id)) { Reply(c, 404, "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"no folder\"}}"); return; }
                    Reply(c, 200, "{\"id\":\"" + id + "\",\"displayName\":\"" + Esc(_names[id]) + "\",\"parentFolderId\":\"" + _parents[id] + "\"}"); return;
                }
                if (p.Length == 4 && p[1] == "mailFolders" && p[3] == "childFolders") { Folders(c, method, path, p[2], text, "displayName"); return; }
                if (path == "me/contactFolders") { Folders(c, method, path, "contactsroot", text, "displayName"); return; }
                if (path == "me/calendars") { Folders(c, method, path, "calroot", text, "name"); return; }

                // ---- items of a folder
                if (p.Length == 4 && ((p[1] == "mailFolders" && p[3] == "messages") || (p[1] == "calendars" && p[3] == "events") || (p[1] == "contactFolders" && p[3] == "contacts")))
                {
                    string key = path.ToLowerInvariant();
                    List<string[]> list;
                    if (!_collections.TryGetValue(key, out list)) { list = new List<string[]>(); _collections[key] = list; }
                    if (method == "GET")
                    {
                        Reply(c, 200, "{\"value\":[" + string.Join(",", list.Select(i => "{\"id\":\"" + i[0] + "\",\"subject\":\"" + Esc(i[1]) + "\",\"displayName\":\"" + Esc(i[1]) + "\"}")) + "]}");
                        return;
                    }
                    string title;
                    if (contentType.StartsWith("text/plain"))
                    {
                        string mime = Encoding.UTF8.GetString(Convert.FromBase64String(text));
                        var line = mime.Split('\n').FirstOrDefault(l => l.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase));
                        title = line == null ? "" : line.Substring(8).Trim();
                    }
                    else
                    {
                        using (var doc = JsonDocument.Parse(text))
                        {
                            JsonElement e;
                            // The two refusals seen on the real service (400 "UnableToDeserializePostBody" / "RequestBodyRead").
                            if (p[1] == "mailFolders" && doc.RootElement.TryGetProperty("sensitivity", out e))
                            { Reply(c, 400, "{\"error\":{\"code\":\"ErrorInvalidRequest\",\"message\":\"UnableToDeserializePostBody: were unable to deserialize (sensitivity is not a message property)\"}}"); return; }
                            JsonElement rec, rng, rtype;
                            if (p[1] == "calendars" && doc.RootElement.TryGetProperty("recurrence", out rec) && rec.TryGetProperty("range", out rng) && rng.TryGetProperty("type", out rtype)
                                && !new[] { "endDate", "noEnd", "numbered" }.Contains(rtype.GetString()))
                            { Reply(c, 400, "{\"error\":{\"code\":\"BadRequest\",\"message\":\"RequestBodyRead: Requested value '" + rtype.GetString() + "' was not found.\"}}"); return; }
                            title = doc.RootElement.TryGetProperty("subject", out e) ? e.GetString() : doc.RootElement.TryGetProperty("displayName", out e) ? e.GetString() : "";
                        }
                    }
                    string newId = NewId(p[1] == "mailFolders" ? "msg" : p[1] == "calendars" ? "ev" : "ct");
                    list.Add(new[] { newId, title });
                    Reply(c, 201, "{\"id\":\"" + newId + "\",\"subject\":\"" + Esc(title) + "\"}");
                    return;
                }

                // ---- children of an item
                if (p.Length == 4 && p[3] == "attachments" && method == "POST") { Reply(c, 201, "{\"id\":\"" + NewId("att") + "\"}"); return; }
                if (p.Length == 4 && p[1] == "events" && p[3] == "instances" && method == "GET")
                {
                    string ev = p[2];
                    var start = new DateTime(2026, 11, 9, 13, 0, 0, DateTimeKind.Utc);
                    var rows = new List<string>();
                    for (int k = 0; k < 6; k++)
                    {
                        string id = "inst-" + ev + "-" + k;
                        if (_cancelled.Contains(id)) continue;
                        DateTime s = start.AddDays(7 * k);
                        rows.Add("{\"id\":\"" + id + "\",\"type\":\"" + (_exceptions.Contains(id) ? "exception" : "occurrence") + "\",\"subject\":\"x\",\"start\":{\"dateTime\":\"" +
                            s.ToString("yyyy-MM-ddTHH:mm:ss.0000000") + "\",\"timeZone\":\"UTC\"},\"end\":{\"dateTime\":\"" + s.AddMinutes(30).ToString("yyyy-MM-ddTHH:mm:ss.0000000") + "\",\"timeZone\":\"UTC\"}}");
                    }
                    Reply(c, 200, "{\"value\":[" + string.Join(",", rows) + "]}");
                    return;
                }
                if (p.Length == 3 && p[1] == "events" && p[2].StartsWith("inst-"))
                {
                    if (method == "PATCH") { _exceptions.Add(p[2]); Reply(c, 200, "{}"); return; }
                    if (method == "DELETE") { _cancelled.Add(p[2]); Reply(c, 204, null); return; }
                }
                if (p.Length == 3 && (p[1] == "messages" || p[1] == "events" || p[1] == "contacts") && method == "PATCH") { Reply(c, 200, "{}"); return; }
                if (path.EndsWith("/photo/$value") && method == "PUT") { Reply(c, 200, null); return; }

                Reply(c, 404, "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"the fake does not know " + method + " " + Esc(path) + "\"}}");
            }
        }

        void Folders(HttpListenerContext c, string method, string path, string parent, string text, string nameProperty)
        {
            string key = path.ToLowerInvariant();
            List<string[]> list;
            if (!_collections.TryGetValue(key, out list)) { list = new List<string[]>(); _collections[key] = list; }
            if (method == "GET")
            {
                Reply(c, 200, "{\"value\":[" + string.Join(",", list.Select(i => "{\"id\":\"" + i[0] + "\",\"" + nameProperty + "\":\"" + Esc(i[1]) + "\"}")) + "]}");
                return;
            }
            string name;
            using (var doc = JsonDocument.Parse(text)) name = doc.RootElement.GetProperty(nameProperty).GetString();
            string id = NewId("fld");
            list.Add(new[] { id, name });
            _names[id] = name; _parents[id] = parent == "msgfolderroot" ? "root" : parent;
            if (nameProperty == "name") _calendarIds.Add(id);
            Reply(c, 201, "{\"id\":\"" + id + "\",\"" + nameProperty + "\":\"" + Esc(name) + "\"}");
        }
    }
}
