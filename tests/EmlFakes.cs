// =============================================================================
//  Mailbox Transfer - fake Graph pieces for the EML tests (tests only, no network)
// -----------------------------------------------------------------------------
//  EmlFakeMail wraps a FakeMailApi (which hands /admin/exchange to a FakeMailboxService) and adds what the EML code uses on top:
//    GET  /users/{mailbox}/messages/{id}/$value   the MIME of a message (Mime["mailbox|id"])
//    POST /users/{mailbox}/messages               text/plain base64 MIME -> a draft (the Exchange conversion): the draft is registered in the fake
//                                                 import-export service with the FastTransfer stream the test gives in NextDraftData (what
//                                                 "Exchange" would export for it); DenyDrafts answers 403 (no Mail.ReadWrite)
//    DELETE /users/{mailbox}/messages/{id}        deletes the draft
// =============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer.Tests
{
    public sealed class EmlFakeMail : HttpMessageHandler
    {
        readonly HttpMessageInvoker _inner;
        readonly FakeMailboxService _svc;
        readonly object _lock = new object();
        int _drafts;

        public readonly Dictionary<string, byte[]> Mime = new Dictionary<string, byte[]>();
        public readonly List<string> Log = new List<string>();
        public bool DenyDrafts;
        public byte[] NextDraftData;
        public string DraftFolder = "ROOT";
        public byte[] LastDraftBody;
        public string LastDraftContentType;
        public readonly List<string> DraftsCreated = new List<string>();
        public readonly List<string> DraftsDeleted = new List<string>();
        /// <summary>GET $value answers 404 for these ids.</summary>
        public readonly HashSet<string> MissingMime = new HashSet<string>();

        public EmlFakeMail(FakeMailApi mail, FakeMailboxService svc) { _inner = new HttpMessageInvoker(mail); _svc = svc; }

        static HttpResponseMessage Json(HttpStatusCode code, string json)
        {
            return new HttpResponseMessage(code) { Content = new StringContent(json, Encoding.UTF8, "application/json") };
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
        {
            string path = Uri.UnescapeDataString(request.RequestUri.AbsolutePath);
            lock (_lock) Log.Add(request.Method + " " + path);
            int at = path.IndexOf("/users/", StringComparison.Ordinal);
            if (at >= 0)
            {
                string rest = path.Substring(at + 7);
                int slash = rest.IndexOf('/');
                string mailbox = slash < 0 ? rest : rest.Substring(0, slash);
                string tail = slash < 0 ? "" : rest.Substring(slash);
                if (request.Method == HttpMethod.Get && tail.StartsWith("/messages/") && tail.EndsWith("/$value"))
                {
                    string id = tail.Substring("/messages/".Length, tail.Length - "/messages/".Length - "/$value".Length);
                    byte[] data;
                    if (MissingMime.Contains(id) || !Mime.TryGetValue(mailbox + "|" + id, out data))
                        return Json(HttpStatusCode.NotFound, "{\"error\":{\"code\":\"ErrorItemNotFound\",\"message\":\"The specified object was not found in the store.\"}}");
                    var ok = new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(data) };
                    ok.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("text/plain");
                    return ok;
                }
                if (request.Method == HttpMethod.Post && tail == "/messages")
                {
                    byte[] body = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    LastDraftContentType = request.Content.Headers.ContentType == null ? null : request.Content.Headers.ContentType.MediaType;
                    if (DenyDrafts) return Json(HttpStatusCode.Forbidden, "{\"error\":{\"code\":\"ErrorAccessDenied\",\"message\":\"Access is denied. Check credentials and try again.\"}}");
                    LastDraftBody = Convert.FromBase64String(Encoding.ASCII.GetString(body));
                    string id = "DRAFT" + (++_drafts);
                    lock (_lock) DraftsCreated.Add(id);
                    if (NextDraftData == null) return Json(HttpStatusCode.InternalServerError, "{\"error\":{\"code\":\"x\",\"message\":\"no draft data prepared\"}}");
                    _svc.AddItem(mailbox, DraftFolder, id, NextDraftData);
                    return Json(HttpStatusCode.Created, "{\"id\":\"" + id + "\"}");
                }
                if (request.Method == HttpMethod.Delete && tail.StartsWith("/messages/"))
                {
                    lock (_lock) DraftsDeleted.Add(tail.Substring("/messages/".Length));
                    return new HttpResponseMessage(HttpStatusCode.NoContent);
                }
            }
            return await _inner.SendAsync(request, ct).ConfigureAwait(false);
        }
    }
}
