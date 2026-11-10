// =============================================================================
//  Mailbox Transfer - engine, part 14: MAPI object model (FxMessage) <-> MIME (MimeEntity)
// -----------------------------------------------------------------------------
//  MapiToMime.Build   FxMessage (export store, PST, any FastTransfer stream) -> MimeEntity -> .eml
//  MimeToMapi.Build   MimeEntity (parsed .eml)                                -> FxMessage  -> FastTransfer stream -> importItem
//
//  Rules, MAPI -> MIME
//    - Headers: PR_TRANSPORT_MESSAGE_HEADERS (0x007D) when the message has it, kept line for line (Content-* and MIME-Version are rebuilt);
//      what is missing (From, To, Cc, Subject, Date, Message-ID...) is written from the properties. Without transport headers every
//      header is built from the properties (sender / sent representing, recipients by type, PR_CLIENT_SUBMIT_TIME, ...).
//    - A message with no Received header gets X-MailboxTransfer-Received (PR_MESSAGE_DELIVERY_TIME): the date the message reached
//      the mailbox is not in any other header and the import gives it back.
//    - Body: PR_BODY (text/plain), PR_HTML (text/html in the charset of PR_INTERNET_CPID) or the HTML / text of PR_RTF_COMPRESSED
//      when it is the only body. Both -> multipart/alternative. Attachments referenced by cid: in the HTML -> multipart/related.
//    - Attachments by value -> base64 parts (streamed from the source, never loaded whole); embedded messages -> message/rfc822.
//      Not exported (reported in the notes): OLE attachments, attachments by reference.
//  Rules, MIME -> MAPI
//    - A text part without file name and not attached is the body (first text/plain, first text/html); the other leaves are attachments
//      (CID -> inline), message/rfc822 -> embedded message. The HTML is stored in UTF-8 (PR_INTERNET_CPID 65001).
//    - Dates: PR_CLIENT_SUBMIT_TIME = Date, PR_MESSAGE_DELIVERY_TIME = X-MailboxTransfer-Received, else the first (topmost) Received, else Date.
//    - The whole header block is kept in PR_TRANSPORT_MESSAGE_HEADERS.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace MailboxTransfer
{
    /// <summary>Tags used here (id only; the type is the one found in the stream).</summary>
    public static class Pid
    {
        public const ushort MessageClass = 0x001A, Importance = 0x0017, Priority = 0x0026, Sensitivity = 0x0036, Subject = 0x0037, ClientSubmitTime = 0x0039,
            SentRepresentingSearchKey = 0x003B, SubjectPrefix = 0x003D, SentRepresentingEntryId = 0x0041, SentRepresentingName = 0x0042,
            SentRepresentingAddrType = 0x0064, SentRepresentingEmail = 0x0065, ConversationTopic = 0x0070, TransportHeaders = 0x007D,
            ReadReceiptRequested = 0x0029, SenderEntryId = 0x0C19, SenderName = 0x0C1A, SenderSearchKey = 0x0C1D, SenderAddrType = 0x0C1E, SenderEmail = 0x0C1F,
            RecipientType = 0x0C15, MessageDeliveryTime = 0x0E06, MessageFlags = 0x0E07, HasAttach = 0x0E1B, NormalizedSubject = 0x0E1D, AttachNumber = 0x0E21,
            Body = 0x1000, RtfCompressed = 0x1009, Html = 0x1013, InternetMessageId = 0x1035, References = 0x1039, InReplyTo = 0x1042,
            DisplayName = 0x3001, AddrType = 0x3002, EmailAddress = 0x3003, CreationTime = 0x3007, LastModificationTime = 0x3008, RowId = 0x3000,
            AttachDataBin = 0x3701, AttachExtension = 0x3703, AttachFilename = 0x3704, AttachMethod = 0x3705, AttachLongFilename = 0x3707, AttachMimeTag = 0x370E,
            AttachContentBase = 0x3711, AttachContentId = 0x3712, AttachContentLocation = 0x3713, AttachFlags = 0x3714, AttachContentDisposition = 0x3716,
            InternetCpid = 0x3FDE, SmtpAddress = 0x39FE, SenderSmtpAddress = 0x5D01, SentRepresentingSmtpAddress = 0x5D02, AttachmentHidden = 0x7FFE,
            ObjectType = 0x0FFE, DisplayType = 0x3900, Responsibility = 0x0E0F, RecipientDisplayName = 0x5FF6, AttachSize = 0x0E20;
    }

    public sealed class MimeBuildOptions
    {
        /// <summary>Use PR_TRANSPORT_MESSAGE_HEADERS for the header block (default). False: every header is built from the properties.</summary>
        public bool UseTransportHeaders = true;
        /// <summary>Add X-MailboxTransfer-Received when the message has no Received header.</summary>
        public bool AddReceivedTime = true;
    }

    public static class MapiToMime
    {
        static readonly string[] Rebuilt = { "content-type", "content-transfer-encoding", "content-disposition", "mime-version", "content-id", "content-description", "content-length", "content-md5" };

        // ---- property readers -------------------------------------------------------------------------------------------

        public static FxProperty Prop(List<FxProperty> props, ushort id)
        {
            foreach (FxProperty p in props) if (p.Id == id && !p.IsNamed) return p;
            return null;
        }

        public static string Str(List<FxProperty> props, ushort id)
        {
            FxProperty p = Prop(props, id);
            if (p == null) return null;
            try
            {
                FxTypeInfo t = p.TypeInfo;
                if (t.IsString && !t.MultiValued) return p.AsString();
                if (t.Valid && !t.MultiValued && p.Type == 0x0102 && p.RawLength < 1 << 20) return Encoding.UTF8.GetString(p.Raw).TrimEnd('\0');
            }
            catch (Exception) { }
            return null;
        }

        public static int? Int(List<FxProperty> props, ushort id)
        {
            FxProperty p = Prop(props, id);
            if (p == null || p.TypeInfo.MultiValued) return null;
            try
            {
                if (p.Type == 0x0003 || p.Type == 0x000A) return p.AsInt32();
                if (p.Type == 0x0002) return p.AsInt16();
                if (p.Type == 0x000B) return p.AsBool() ? 1 : 0;
            }
            catch (Exception) { }
            return null;
        }

        public static DateTimeOffset? Time(List<FxProperty> props, ushort id)
        {
            FxProperty p = Prop(props, id);
            if (p == null || p.Type != 0x0040) return null;
            try { return p.AsTime(); } catch (Exception) { return null; }
        }

        static bool Bool(List<FxProperty> props, ushort id)
        {
            FxProperty p = Prop(props, id);
            try { return p != null && p.Type == 0x000B && p.AsBool(); } catch (Exception) { return false; }
        }

        // ---- mapping ------------------------------------------------------------------------------------------------------

        public sealed class Address { public string Name, Email; public int Type; public bool Resolved; }

        static string SmtpOf(List<FxProperty> props, ushort smtpId, ushort typeId, ushort emailId)
        {
            string smtp = Str(props, smtpId);
            if (!string.IsNullOrEmpty(smtp) && smtp.Contains("@")) return smtp;
            string type = Str(props, typeId);
            string email = Str(props, emailId);
            if (!string.IsNullOrEmpty(email) && (string.Equals(type, "SMTP", StringComparison.OrdinalIgnoreCase) || (type == null && email.Contains("@")))) return email;
            return null;
        }

        static List<Address> Recipients(FxMessage m, List<string> notes)
        {
            var list = new List<Address>();
            foreach (FxRecipient r in m.Recipients)
            {
                var a = new Address { Type = Int(r.Properties, Pid.RecipientType) ?? 1 };
                a.Name = Str(r.Properties, Pid.DisplayName) ?? Str(r.Properties, Pid.RecipientDisplayName);
                a.Email = SmtpOf(r.Properties, Pid.SmtpAddress, Pid.AddrType, Pid.EmailAddress);
                if (a.Email == null)
                {
                    string dn = Str(r.Properties, Pid.EmailAddress);
                    if (!string.IsNullOrEmpty(dn))
                    {
                        a.Email = dn;   // an Exchange distinguished name (EX) the properties do not resolve: kept as it is, reported
                        notes.Add("recipient '" + a.Name + "' has no SMTP address (" + (Str(r.Properties, Pid.AddrType) ?? "?") + " address kept)");
                    }
                }
                if (a.Email == null && a.Name == null) continue;
                list.Add(a);
            }
            return list;
        }

        static string AddressList(IEnumerable<Address> list)
        {
            return string.Join(", ", list.Select(a => MimeText.FormatAddress(a.Name, a.Email)));
        }

        public static bool IsMailClass(string messageClass)
        {
            if (string.IsNullOrEmpty(messageClass)) return true;
            string c = messageClass.ToUpperInvariant();
            string[] nonMail = { "IPM.APPOINTMENT", "IPM.CONTACT", "IPM.DISTLIST", "IPM.TASK", "IPM.STICKYNOTE", "IPM.ACTIVITY", "IPM.CONFIGURATION", "IPM.MICROSOFT.", "IPM.OLE.CLASS.", "IPM.POST", "IPM.REMINDER" };
            foreach (string p in nonMail) if (c.StartsWith(p, StringComparison.Ordinal)) return false;
            return true;
        }

        /// <summary>The MIME tree of a message. notes receives what could not be carried over.</summary>
        public static MimeEntity Build(FxMessage m, MimeBuildOptions options = null, List<string> notes = null)
        {
            if (options == null) options = new MimeBuildOptions();
            if (notes == null) notes = new List<string>();
            List<FxProperty> p = m.Properties;
            var headers = new List<MimeHeader>();
            bool haveTransport = false;

            string transport = options.UseTransportHeaders ? Str(p, Pid.TransportHeaders) : null;
            if (!string.IsNullOrWhiteSpace(transport))
            {
                MimeEntity parsed = MimeParser.Parse(Encoding.UTF8.GetBytes(transport.TrimEnd('\0').Replace("\r\n", "\n").Replace("\n", "\r\n") + "\r\n\r\n"));
                foreach (MimeHeader h in parsed.Headers)
                    if (!Rebuilt.Contains(h.Name.ToLowerInvariant())) headers.Add(h);
                haveTransport = headers.Count > 0;
            }
            Func<string, bool> has = name => headers.Any(h => string.Equals(h.Name, name, StringComparison.OrdinalIgnoreCase));
            Action<string, string> add = (name, value) => { if (!string.IsNullOrEmpty(value)) headers.Add(new MimeHeader { Name = name, Value = value }); };

            List<Address> rcpt = Recipients(m, notes);
            string senderName = Str(p, Pid.SenderName), senderEmail = SmtpOf(p, Pid.SenderSmtpAddress, Pid.SenderAddrType, Pid.SenderEmail);
            string repName = Str(p, Pid.SentRepresentingName), repEmail = SmtpOf(p, Pid.SentRepresentingSmtpAddress, Pid.SentRepresentingAddrType, Pid.SentRepresentingEmail);
            string fromName = repEmail != null ? repName : senderName, fromEmail = repEmail ?? senderEmail;

            DateTimeOffset? sent = Time(p, Pid.ClientSubmitTime);
            DateTimeOffset? delivered = Time(p, Pid.MessageDeliveryTime);
            string subject = Str(p, Pid.Subject);
            string msgId = Str(p, Pid.InternetMessageId);

            if (!has("Received") && options.AddReceivedTime && delivered.HasValue)
                headers.Insert(0, new MimeHeader { Name = "X-MailboxTransfer-Received", Value = MimeText.FormatDate(delivered.Value.ToUniversalTime()) });
            if (!has("From") && (fromEmail != null || fromName != null)) add("From", MimeText.FormatAddress(fromName, fromEmail));
            if (!has("Sender") && senderEmail != null && repEmail != null && !string.Equals(senderEmail, repEmail, StringComparison.OrdinalIgnoreCase))
                add("Sender", MimeText.FormatAddress(senderName, senderEmail));
            if (!has("To")) add("To", AddressList(rcpt.Where(a => a.Type == 1)));
            if (!has("Cc")) add("Cc", AddressList(rcpt.Where(a => a.Type == 2)));
            if (!has("Bcc")) add("Bcc", AddressList(rcpt.Where(a => a.Type == 3)));
            if (!has("Subject") && subject != null) add("Subject", MimeText.EncodeText(subject));
            if (!has("Date")) { DateTimeOffset? d = sent ?? delivered; if (d.HasValue) add("Date", MimeText.FormatDate(d.Value.ToUniversalTime())); }
            if (!has("Message-ID") && !string.IsNullOrEmpty(msgId)) add("Message-ID", msgId.StartsWith("<") ? msgId : "<" + msgId + ">");
            if (!has("In-Reply-To")) { string v = Str(p, Pid.InReplyTo); if (!string.IsNullOrEmpty(v)) add("In-Reply-To", v); }
            if (!has("References")) { string v = Str(p, Pid.References); if (!string.IsNullOrEmpty(v)) add("References", v); }
            if (!has("Reply-To"))
            {
                FxProperty rp = Prop(p, 0x004F);
                if (rp != null && !rp.TypeInfo.MultiValued && rp.RawLength > 0 && rp.RawLength < 1 << 20)
                    add("Reply-To", string.Join(", ", FlatEntries(rp.Raw).Select(x => MimeText.FormatAddress(x.Key, x.Value))));
            }
            if (!has("Thread-Topic")) { string v = Str(p, Pid.ConversationTopic); if (!string.IsNullOrEmpty(v)) add("Thread-Topic", MimeText.EncodeText(v)); }
            if (!has("Thread-Index"))
            {
                FxProperty ci = Prop(p, 0x0071);
                if (ci != null && !ci.TypeInfo.MultiValued && ci.RawLength > 0 && ci.RawLength < 4096) add("Thread-Index", System.Convert.ToBase64String(ci.Raw));
            }
            if (!has("Importance") && !has("X-Priority"))
            {
                int? imp = Int(p, Pid.Importance);
                if (imp == 2) { add("Importance", "high"); add("X-Priority", "1"); }
                else if (imp == 0) { add("Importance", "low"); add("X-Priority", "5"); }
            }
            if (!has("Sensitivity"))
            {
                int? sens = Int(p, Pid.Sensitivity);
                if (sens == 1) add("Sensitivity", "personal"); else if (sens == 2) add("Sensitivity", "private"); else if (sens == 3) add("Sensitivity", "company-confidential");
            }
            if (!haveTransport && Bool(p, Pid.ReadReceiptRequested) && !has("Disposition-Notification-To") && fromEmail != null) add("Disposition-Notification-To", fromEmail);

            CheckSkeleton(m, notes);
            MimeEntity body = BuildBody(m, options, notes);
            var root = body;
            // message headers first, then the content headers the body already carries
            var content = root.Headers;
            root.Headers = new List<MimeHeader>(headers);
            root.Add("MIME-Version", "1.0");
            root.Headers.AddRange(content);
            return root;
        }

        /// <summary>The one-off entries (name, address) of a flat entry list such as PR_REPLY_RECIPIENT_ENTRIES (MS-OXCDATA 2.2.5.1 and 2.2.1.2).</summary>
        public static List<KeyValuePair<string, string>> FlatEntries(byte[] b)
        {
            var list = new List<KeyValuePair<string, string>>();
            try
            {
                int count = BitConverter.ToInt32(b, 0);
                int pos = 8;
                for (int i = 0; i < count && pos + 4 <= b.Length; i++)
                {
                    int size = BitConverter.ToInt32(b, pos); pos += 4;
                    int start = pos;
                    if (size >= 24 && start + size <= b.Length)
                    {
                        int flags = BitConverter.ToUInt16(b, start + 22);
                        bool unicode = (flags & 0x8000) != 0;
                        int p0 = start + 24;
                        var parts = new List<string>();
                        for (int k = 0; k < 3 && p0 < start + size; k++)
                        {
                            int e = p0;
                            if (unicode) { while (e + 1 < start + size && (b[e] != 0 || b[e + 1] != 0)) e += 2; parts.Add(Encoding.Unicode.GetString(b, p0, e - p0)); p0 = e + 2; }
                            else { while (e < start + size && b[e] != 0) e++; parts.Add(Encoding.GetEncoding(1252).GetString(b, p0, e - p0)); p0 = e + 1; }
                        }
                        if (parts.Count == 3 && parts[1].Equals("SMTP", StringComparison.OrdinalIgnoreCase)) list.Add(new KeyValuePair<string, string>(parts[0], parts[2]));
                    }
                    pos = start + ((size + 3) & ~3);
                }
            }
            catch (Exception) { }
            return list;
        }

        /// <summary>
        /// PR_MIME_SKELETON (0x64F0) is the MIME of a message created from MIME, with a placeholder (X-Exchange-Mime-Skeleton-Content-Id) for each part whose content
        /// lives in the message. When it lists more attachments than the message holds, the attachments were lost before this export (observed live: an import
        /// without the FXDelProp markers keeps the skeleton and drops the attachments, and the Mail API $value then shows empty attachment parts). Reported in the notes.
        /// </summary>
        static void CheckSkeleton(FxMessage m, List<string> notes)
        {
            FxProperty sk = Prop(m.Properties, 0x64F0);
            if (sk == null || sk.TypeInfo.MultiValued || sk.RawLength == 0 || sk.RawLength > 1 << 20) return;
            try
            {
                MimeEntity skel = MimeParser.Parse(sk.Raw);
                int listed = skel.Walk().Count(e => !e.IsMultipart && e.Header("X-Exchange-Mime-Skeleton-Content-Id") != null && (e.Disposition == "attachment" || e.Disposition == "inline" || e.Message != null || e.ContentType.MediaType == "message/rfc822"));
                if (listed > m.Attachments.Count)
                    notes.Add("the MIME skeleton lists " + listed + " attachment(s), the message holds " + m.Attachments.Count + ": the attachment content was lost before this export (the item in the mailbox has no data for them)");
            }
            catch (Exception) { }
        }

        // ---- body -----------------------------------------------------------------------------------------------------------

        sealed class AttachInfo
        {
            public FxAttachment Att;
            public string Name, ContentType, Cid, Location;
            public bool Inline;
            public bool Hidden;
            public int Method;
        }

        static MimeEntity BuildBody(FxMessage m, MimeBuildOptions options, List<string> notes)
        {
            List<FxProperty> p = m.Properties;
            string plain = Str(p, Pid.Body);
            byte[] htmlBytes = null; string htmlCharset = null;
            FxProperty hp = Prop(p, Pid.Html);
            int cpid = Int(p, Pid.InternetCpid) ?? 0;
            if (hp != null)
            {
                if (hp.TypeInfo.IsString && !hp.TypeInfo.MultiValued) { htmlBytes = new UTF8Encoding(false).GetBytes(hp.AsString()); htmlCharset = "utf-8"; }
                else if (!hp.TypeInfo.MultiValued && hp.RawLength > 0) { htmlBytes = hp.Raw; }
                if (htmlBytes != null && htmlBytes.Length > 0 && htmlCharset == null)
                {
                    Encoding enc = null;
                    if (cpid > 0 && cpid != 1200 && cpid != 1201 && cpid != 12000) { try { enc = Encoding.GetEncoding(cpid); } catch (Exception) { } }
                    if (enc != null) htmlCharset = MimeText.CharsetName(enc);
                    else htmlCharset = MimeText.IsValidUtf8(htmlBytes) ? "utf-8" : "windows-1252";
                    if (cpid == 1200 || cpid == 1201) { htmlBytes = new UTF8Encoding(false).GetBytes(Encoding.GetEncoding(cpid).GetString(htmlBytes)); htmlCharset = "utf-8"; }
                }
            }
            if (htmlBytes != null && htmlBytes.Length == 0) htmlBytes = null;
            if (htmlBytes == null && string.IsNullOrEmpty(plain))
            {
                FxProperty rp = Prop(p, Pid.RtfCompressed);
                if (rp != null && !rp.TypeInfo.MultiValued && rp.RawLength > 0 && rp.RawLength < 64 * 1024 * 1024)
                {
                    byte[] rtf = Rtf.Decompress(rp.Raw);
                    if (rtf != null)
                    {
                        byte[] h = Rtf.ExtractHtml(rtf);
                        if (h != null) { htmlBytes = h; htmlCharset = "utf-8"; notes.Add("body taken from the HTML encapsulated in the RTF"); }
                        else { plain = Rtf.ToText(rtf); notes.Add("body taken from the RTF (converted to text)"); }
                    }
                    else notes.Add("RTF body could not be decompressed");
                }
            }
            string html = htmlBytes == null ? null : MimeText.DecodeText(htmlBytes, htmlCharset);
            if (string.IsNullOrEmpty(plain) && html != null)
            {
                // Exchange keeps one body and builds the other one when MIME is asked for: the same is done here (text from the HTML)
                plain = HtmlText.FromHtml(html);
                notes.Add("text alternative built from the HTML body");
            }

            // attachments
            var inlines = new List<AttachInfo>();
            var files = new List<AttachInfo>();
            foreach (FxAttachment a in m.Attachments)
            {
                int method = Int(a.Properties, Pid.AttachMethod) ?? 1;
                var info = new AttachInfo { Att = a, Method = method };
                info.Name = Str(a.Properties, Pid.AttachLongFilename) ?? Str(a.Properties, Pid.AttachFilename) ?? Str(a.Properties, Pid.DisplayName);
                info.ContentType = Str(a.Properties, Pid.AttachMimeTag);
                info.Cid = Str(a.Properties, Pid.AttachContentId);
                info.Location = Str(a.Properties, Pid.AttachContentLocation);
                info.Hidden = Bool(a.Properties, Pid.AttachmentHidden);
                if (a.Embedded == null && method != 1)
                {
                    notes.Add("attachment '" + (info.Name ?? "?") + "' skipped (attach method " + method + (method == 6 ? ", OLE object" : method >= 2 && method <= 4 ? ", by reference" : "") + ")");
                    continue;
                }
                if (a.Embedded == null && FindData(a) == null) { notes.Add("attachment '" + (info.Name ?? "?") + "' has no data, skipped"); continue; }
                bool referenced = html != null && !string.IsNullOrEmpty(info.Cid) && html.IndexOf("cid:" + info.Cid, StringComparison.OrdinalIgnoreCase) >= 0;
                info.Inline = referenced && a.Embedded == null;
                (info.Inline ? inlines : files).Add(info);
            }

            MimeEntity bodyPart = null;
            MimeEntity plainPart = plain != null ? TextPart("plain", Encoding.UTF8.GetBytes(plain), "utf-8") : null;
            MimeEntity htmlPart = htmlBytes != null ? TextPart("html", htmlBytes, htmlCharset) : null;
            if (htmlPart != null && inlines.Count > 0)
            {
                var related = new MimeEntity { ContentType = MultiType("related") };
                related.ContentType.Parameters["type"] = "text/html";
                SetContentType(related, "multipart/related; type=\"text/html\"");
                related.Parts.Add(htmlPart);
                foreach (AttachInfo i in inlines) related.Parts.Add(FilePart(i, true, notes));
                htmlPart = related;
            }
            else if (inlines.Count > 0) files.InsertRange(0, inlines);
            if (plainPart != null && htmlPart != null)
            {
                bodyPart = new MimeEntity { ContentType = MultiType("alternative") };
                SetContentType(bodyPart, "multipart/alternative");
                bodyPart.Parts.Add(plainPart);
                bodyPart.Parts.Add(htmlPart);
            }
            else bodyPart = htmlPart ?? plainPart ?? TextPart("plain", new byte[0], "utf-8");

            if (files.Count == 0) return bodyPart;
            var mixed = new MimeEntity { ContentType = MultiType("mixed") };
            SetContentType(mixed, "multipart/mixed");
            mixed.Parts.Add(bodyPart);
            foreach (AttachInfo f in files) mixed.Parts.Add(FilePart(f, false, notes));
            return mixed;
        }

        static MimeContentType MultiType(string sub) { return new MimeContentType { Type = "multipart", SubType = sub }; }

        static void SetContentType(MimeEntity e, string header)
        {
            e.Headers.RemoveAll(h => string.Equals(h.Name, "Content-Type", StringComparison.OrdinalIgnoreCase));
            e.Headers.Insert(0, new MimeHeader { Name = "Content-Type", Value = header });
            e.ContentType = MimeContentType.Parse(header);
            e.HasContentType = true;
        }

        static MimeEntity TextPart(string sub, byte[] bytes, string charset)
        {
            var e = new MimeEntity();
            e.Content = new FxBytesBlob(bytes);
            bool ascii = bytes.All(b => b < 128 && b != 0);
            int high = bytes.Count(b => b >= 128 || b == 0);
            string cte = ascii && !LongLines(bytes) ? "7bit" : high * 3 > bytes.Length ? "base64" : "quoted-printable";
            e.TransferEncoding = cte;
            SetContentType(e, "text/" + sub + "; charset=\"" + (ascii ? "us-ascii" : charset) + "\"");
            e.Headers.Add(new MimeHeader { Name = "Content-Transfer-Encoding", Value = cte });
            return e;
        }

        static bool LongLines(byte[] b)
        {
            int run = 0;
            foreach (byte x in b) { if (x == '\n' || x == '\r') run = 0; else if (++run > 990) return true; }
            return false;
        }

        static FxProperty FindData(FxAttachment a)
        {
            return a.Properties.FirstOrDefault(x => x.Id == Pid.AttachDataBin && !x.IsNamed);
        }

        static readonly Dictionary<string, string> Types = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { ".pdf", "application/pdf" }, { ".png", "image/png" }, { ".jpg", "image/jpeg" }, { ".jpeg", "image/jpeg" }, { ".gif", "image/gif" },
            { ".bmp", "image/bmp" }, { ".tif", "image/tiff" }, { ".tiff", "image/tiff" }, { ".svg", "image/svg+xml" }, { ".txt", "text/plain" }, { ".htm", "text/html" },
            { ".html", "text/html" }, { ".csv", "text/csv" }, { ".xml", "application/xml" }, { ".json", "application/json" }, { ".ics", "text/calendar" },
            { ".vcf", "text/vcard" }, { ".zip", "application/zip" }, { ".doc", "application/msword" }, { ".xls", "application/vnd.ms-excel" },
            { ".ppt", "application/vnd.ms-powerpoint" },
            { ".docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
            { ".xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
            { ".pptx", "application/vnd.openxmlformats-officedocument.presentationml.presentation" },
            { ".msg", "application/vnd.ms-outlook" }, { ".eml", "message/rfc822" }, { ".mp3", "audio/mpeg" }, { ".mp4", "video/mp4" }, { ".rtf", "application/rtf" }
        };

        static string ContentTypeFor(AttachInfo i)
        {
            if (!string.IsNullOrEmpty(i.ContentType) && Regex.IsMatch(i.ContentType, @"^[A-Za-z0-9!#$&^_.+-]+/[A-Za-z0-9!#$&^_.+-]+$")) return i.ContentType.ToLowerInvariant();
            string ext = i.Name == null ? "" : Path.GetExtension(i.Name);
            string t;
            return !string.IsNullOrEmpty(ext) && Types.TryGetValue(ext, out t) ? t : "application/octet-stream";
        }

        static MimeEntity FilePart(AttachInfo i, bool inline, List<string> notes)
        {
            var e = new MimeEntity();
            string name = i.Name;
            if (i.Att.Embedded != null)
            {
                FxMessage em = i.Att.Embedded;
                e.Message = Build(em, new MimeBuildOptions { UseTransportHeaders = true, AddReceivedTime = false }, notes);
                SetContentType(e, "message/rfc822");
                if (name == null || !name.EndsWith(".eml", StringComparison.OrdinalIgnoreCase)) name = (Str(em.Properties, Pid.Subject) ?? name ?? "message") + ".eml";
                e.Headers.Add(new MimeHeader { Name = "Content-Disposition", Value = "attachment; " + MimeText.FormatParameter("filename", SafeName(name)) });
                return e;
            }
            string ct = ContentTypeFor(i);
            e.Content = FindData(i.Att).Blob ?? new FxBytesBlob(new byte[0]);
            e.TransferEncoding = "base64";
            SetContentType(e, ct + (name != null ? "; " + MimeText.FormatParameter("name", SafeName(name)) : ""));
            e.Headers.Add(new MimeHeader { Name = "Content-Transfer-Encoding", Value = "base64" });
            e.Headers.Add(new MimeHeader
            {
                Name = "Content-Disposition",
                Value = (inline ? "inline" : "attachment") + (name != null ? "; " + MimeText.FormatParameter("filename", SafeName(name)) : "")
            });
            if (inline && !string.IsNullOrEmpty(i.Cid)) e.Headers.Add(new MimeHeader { Name = "Content-ID", Value = "<" + i.Cid.Trim('<', '>') + ">" });
            if (!string.IsNullOrEmpty(i.Location)) e.Headers.Add(new MimeHeader { Name = "Content-Location", Value = i.Location });
            return e;
        }

        static string SafeName(string n) { return n.Replace("\r", " ").Replace("\n", " "); }

        public static byte[] ToBytes(FxMessage m, MimeBuildOptions options = null, List<string> notes = null)
        {
            return MimeWriter.ToBytes(Build(m, options, notes));
        }

        public static void Write(FxMessage m, Stream dest, MimeBuildOptions options = null, List<string> notes = null, System.Threading.CancellationToken ct = default(System.Threading.CancellationToken))
        {
            MimeWriter.Write(Build(m, options, notes), dest, new MimeWriteOptions { Cancellation = ct });
        }
    }

    // ---------------------------------------------------------------------------------------------------------------------
    //  MIME -> MAPI
    // ---------------------------------------------------------------------------------------------------------------------

    public sealed class EmlImportOptions
    {
        /// <summary>Message flag read (default true: an archive import is not new mail). The manifest or a per-file setting can override it.</summary>
        public bool Read = true;
        /// <summary>Received time to use whatever the headers say (null: from the headers).</summary>
        public DateTimeOffset? Received;
        public string MessageClass = "IPM.Note";
        /// <summary>Keep the header block in PR_TRANSPORT_MESSAGE_HEADERS.</summary>
        public bool KeepHeaders = true;
    }

    public static class MimeToMapi
    {
        static readonly Guid OneOffUid = new Guid("a41f2b81-a3be-1910-9d6e-00dd010f5402");
        static readonly Regex SubjectPrefix = new Regex(@"^(?<p>[A-Za-zÀ-ſ]{1,3}(?::|\s?:)\s)(?<rest>.*)$", RegexOptions.Compiled | RegexOptions.Singleline);

        /// <summary>FlatEntryList of MS-OXCDATA 2.2.1.2: count, size, then each entry as size + bytes padded to 4.</summary>
        public static byte[] FlatEntryList(List<byte[]> entries)
        {
            var body = new MemoryStream();
            var w = new BinaryWriter(body);
            foreach (byte[] e in entries) { w.Write(e.Length); w.Write(e); for (int i = e.Length; (i & 3) != 0; i++) w.Write((byte)0); }
            var ms = new MemoryStream();
            var o = new BinaryWriter(ms);
            o.Write(entries.Count); o.Write((int)body.Length); o.Write(body.ToArray());
            return ms.ToArray();
        }

        public static byte[] OneOffEntryId(string name, string addrType, string email)
        {
            var ms = new MemoryStream();
            var w = new BinaryWriter(ms);
            w.Write(0);                          // flags
            w.Write(OneOffUid.ToByteArray());    // provider UID of the one-off provider
            w.Write((ushort)0);                  // version
            w.Write((ushort)0x8000);             // flags: Unicode
            foreach (string s in new[] { name ?? "", addrType, email ?? "" }) w.Write(Encoding.Unicode.GetBytes(s + "\0"));
            return ms.ToArray();
        }

        /// <summary>The topmost Received header's date, if any (the first one in the file is the last hop, the arrival in the mailbox).</summary>
        public static DateTimeOffset? ReceivedFromHeaders(MimeEntity e)
        {
            string mbt = e.Header("X-MailboxTransfer-Received");
            DateTimeOffset? d = MimeText.ParseDate(mbt);
            if (d.HasValue) return d;
            foreach (string r in e.HeaderAll("Received"))
            {
                int semi = r.LastIndexOf(';');
                if (semi < 0) continue;
                d = MimeText.ParseDate(r.Substring(semi + 1));
                if (d.HasValue) return d;
            }
            return null;
        }

        sealed class Collected
        {
            public MimeEntity Plain, Html;
            public List<MimeEntity> Files = new List<MimeEntity>();
        }

        static bool IsAttachmentPart(MimeEntity e)
        {
            return e.Disposition == "attachment" || (e.FileName != null && e.ContentType.Type == "text" && e.Disposition != "inline") || e.FileName != null;
        }

        static void Collect(MimeEntity e, Collected c)
        {
            if (e.IsMultipart)
            {
                foreach (MimeEntity part in e.Parts) Collect(part, c);
                return;
            }
            if (e.Message != null) { c.Files.Add(e); return; }
            if (e.ContentType.Type == "text" && !IsAttachmentPart(e) && e.Content != null)
            {
                if (e.ContentType.SubType == "plain" && c.Plain == null) { c.Plain = e; return; }
                if (e.ContentType.SubType == "html" && c.Html == null) { c.Html = e; return; }
            }
            if (e.Content != null) c.Files.Add(e);
        }

        static FxProperty Str(ushort id, string v) { return FxProperty.FromString(id, v); }

        /// <summary>
        /// The MAPI message of a parsed .eml, ready for importItem: what the service needs in the stream (FXDelProp, PidTagRowid, PidTagMid,
        /// no empty PidTagAttachDataObject) is added by FxImportStream.Prepare, the one place that knows these rules.
        /// </summary>
        public static FxMessage Build(MimeEntity root, EmlImportOptions options = null, List<string> notes = null)
        {
            return FxImportStream.Prepare(BuildCore(root, options, notes));
        }

        static FxMessage BuildCore(MimeEntity root, EmlImportOptions options, List<string> notes)
        {
            if (options == null) options = new EmlImportOptions();
            if (notes == null) notes = new List<string>();
            var m = new FxMessage();

            List<FxProperty> p = m.Properties;
            var c = new Collected();
            Collect(root, c);
            if (!root.IsMultipart && root.Message == null && c.Plain == null && c.Html == null && root.Content != null && !c.Files.Contains(root)) c.Files.Add(root);

            string subject = root.HeaderDecoded("Subject") ?? "";
            string normalized = subject, prefix = "";
            Match sm = SubjectPrefix.Match(subject);
            if (sm.Success) { prefix = sm.Groups["p"].Value; normalized = sm.Groups["rest"].Value; }

            DateTimeOffset now = DateTimeOffset.UtcNow;
            DateTimeOffset? date = MimeText.ParseDate(root.Header("Date"));
            DateTimeOffset received = options.Received ?? ReceivedFromHeaders(root) ?? date ?? now;
            DateTimeOffset submit = date ?? received;

            p.Add(Str(Pid.MessageClass, options.MessageClass ?? "IPM.Note"));
            p.Add(Str(Pid.Subject, subject));
            p.Add(Str(Pid.NormalizedSubject, normalized));
            p.Add(Str(Pid.SubjectPrefix, prefix));
            p.Add(Str(Pid.ConversationTopic, normalized));

            // sender
            List<MimeAddress> from = MimeText.ParseAddressList(root.Header("From"));
            List<MimeAddress> sender = MimeText.ParseAddressList(root.Header("Sender"));
            MimeAddress fa = from.FirstOrDefault(), sa = sender.FirstOrDefault() ?? fa;
            if (sa != null) AddSender(p, sa, Pid.SenderName, Pid.SenderAddrType, Pid.SenderEmail, Pid.SenderSearchKey, Pid.SenderEntryId, Pid.SenderSmtpAddress);
            if (fa != null) AddSender(p, fa, Pid.SentRepresentingName, Pid.SentRepresentingAddrType, Pid.SentRepresentingEmail, Pid.SentRepresentingSearchKey, Pid.SentRepresentingEntryId, Pid.SentRepresentingSmtpAddress);

            // Reply-To: PR_REPLY_RECIPIENT_ENTRIES (flat list of one-off entries) and PR_REPLY_RECIPIENT_NAMES
            List<MimeAddress> reply = root.HeaderAll("Reply-To").SelectMany(MimeText.ParseAddressList).Where(x => !string.IsNullOrEmpty(x.Address)).ToList();
            if (reply.Count > 0)
            {
                p.Add(FxProperty.FromBinary(0x004F, FlatEntryList(reply.Select(x => OneOffEntryId(string.IsNullOrEmpty(x.Name) ? x.Address : x.Name, "SMTP", x.Address)).ToList())));
                p.Add(Str(0x0050, string.Join("; ", reply.Select(x => string.IsNullOrEmpty(x.Name) ? x.Address : x.Name))));
            }

            // recipients
            int row = 0;
            foreach (var group in new[] { new { H = "To", T = 1 }, new { H = "Cc", T = 2 }, new { H = "Bcc", T = 3 } })
            {
                foreach (string hv in root.HeaderAll(group.H))
                    foreach (MimeAddress a in MimeText.ParseAddressList(hv))
                    {
                        var r = new FxRecipient();
                        string name = string.IsNullOrEmpty(a.Name) ? a.Address : a.Name;
                        r.Properties.Add(FxProperty.FromInt32(Pid.RowId, row++));
                        r.Properties.Add(FxProperty.FromInt32(Pid.RecipientType, group.T));
                        r.Properties.Add(FxProperty.FromInt32(Pid.ObjectType, 6));
                        r.Properties.Add(FxProperty.FromInt32(Pid.DisplayType, 0));
                        r.Properties.Add(Str(Pid.DisplayName, name));
                        r.Properties.Add(Str(Pid.AddrType, "SMTP"));
                        r.Properties.Add(Str(Pid.EmailAddress, a.Address));
                        r.Properties.Add(Str(Pid.SmtpAddress, a.Address));
                        r.Properties.Add(Str(Pid.RecipientDisplayName, name));
                        r.Properties.Add(FxProperty.FromBinary(0x0FFF, OneOffEntryId(name, "SMTP", a.Address)));
                        r.Properties.Add(FxProperty.FromBool(Pid.Responsibility, true));
                        m.Recipients.Add(r);
                    }
            }

            // times and ids
            p.Add(FxProperty.FromTime(Pid.ClientSubmitTime, submit.UtcDateTime));
            p.Add(FxProperty.FromTime(Pid.MessageDeliveryTime, received.UtcDateTime));
            p.Add(FxProperty.FromTime(Pid.CreationTime, now.UtcDateTime));
            p.Add(FxProperty.FromTime(Pid.LastModificationTime, now.UtcDateTime));
            string mid = root.Header("Message-ID"); if (!string.IsNullOrWhiteSpace(mid)) p.Add(Str(Pid.InternetMessageId, mid.Trim()));
            string irt = root.Header("In-Reply-To"); if (!string.IsNullOrWhiteSpace(irt)) p.Add(Str(Pid.InReplyTo, irt.Trim()));
            string refs = root.Header("References"); if (!string.IsNullOrWhiteSpace(refs)) p.Add(Str(Pid.References, refs.Trim()));
            if (options.KeepHeaders && root.RawHeaderBlock != null) p.Add(Str(Pid.TransportHeaders, root.RawHeaderBlock.Replace("\r\n", "\n").Replace("\n", "\r\n") + "\r\n"));

            // importance, sensitivity, receipt
            int importance = 1;
            string imp = (root.Header("Importance") ?? "").Trim().ToLowerInvariant(), xp = (root.Header("X-Priority") ?? "").Trim();
            if (imp == "high" || xp.StartsWith("1") || xp.StartsWith("2")) importance = 2;
            else if (imp == "low" || xp.StartsWith("4") || xp.StartsWith("5")) importance = 0;
            p.Add(FxProperty.FromInt32(Pid.Importance, importance));
            p.Add(FxProperty.FromInt32(Pid.Priority, importance == 2 ? 1 : importance == 0 ? -1 : 0));
            string sens = (root.Header("Sensitivity") ?? "").Trim().ToLowerInvariant();
            p.Add(FxProperty.FromInt32(Pid.Sensitivity, sens == "personal" ? 1 : sens == "private" ? 2 : sens.StartsWith("company") ? 3 : 0));
            if (root.Header("Disposition-Notification-To") != null) p.Add(FxProperty.FromBool(Pid.ReadReceiptRequested, true));

            // bodies
            if (c.Plain != null) p.Add(Str(Pid.Body, c.Plain.Text()));
            if (c.Html != null)
            {
                string html = c.Html.Text();
                html = Regex.Replace(html, @"(<meta\b[^>]*?charset\s*=\s*[""']?)[A-Za-z0-9_\-:.]+", "${1}utf-8", RegexOptions.IgnoreCase);
                p.Add(FxProperty.FromBinary(Pid.Html, new UTF8Encoding(false).GetBytes(html)));
                p.Add(FxProperty.FromInt32(Pid.InternetCpid, 65001));
            }
            else p.Add(FxProperty.FromInt32(Pid.InternetCpid, 65001));

            // attachments
            int num = 0;
            foreach (MimeEntity f in c.Files)
            {
                FxAttachment a = MakeAttachment(f, num++, now, notes, options);
                m.Attachments.Add(a);
            }

            int flags = (options.Read ? 1 : 0) | (m.Attachments.Count > 0 ? 0x10 : 0);
            p.Add(FxProperty.FromInt32(Pid.MessageFlags, flags));
            p.Add(FxProperty.FromBool(Pid.HasAttach, m.Attachments.Count > 0));
            if (root.Warnings.Count > 0) foreach (string w in root.Warnings) notes.Add("MIME: " + w);
            return m;
        }

        static void AddSender(List<FxProperty> p, MimeAddress a, ushort name, ushort type, ushort email, ushort searchKey, ushort entryId, ushort smtp)
        {
            string display = string.IsNullOrEmpty(a.Name) ? a.Address : a.Name;
            p.Add(Str(name, display));
            p.Add(Str(type, "SMTP"));
            p.Add(Str(email, a.Address));
            p.Add(Str(smtp, a.Address));
            p.Add(FxProperty.FromBinary(searchKey, Encoding.ASCII.GetBytes("SMTP:" + (a.Address ?? "").ToUpperInvariant() + "\0")));
            p.Add(FxProperty.FromBinary(entryId, OneOffEntryId(display, "SMTP", a.Address)));
        }

        static FxAttachment MakeAttachment(MimeEntity f, int num, DateTimeOffset now, List<string> notes, EmlImportOptions options)
        {
            var a = new FxAttachment();
            List<FxProperty> p = a.Properties;
            string name = f.FileName;
            p.Add(FxProperty.FromInt32(Pid.AttachNumber, num));
            if (f.Message != null)
            {
                FxMessage inner = BuildCore(f.Message, new EmlImportOptions { Read = true, KeepHeaders = options.KeepHeaders }, notes);
                inner.Properties.RemoveAll(x => x.Id == Pid.MessageFlags);
                inner.Properties.Add(FxProperty.FromInt32(Pid.MessageFlags, 1));
                if (name == null) name = (f.Message.HeaderDecoded("Subject") ?? "message") + ".eml";
                a.Embedded = inner;
                p.Add(FxProperty.FromInt32(Pid.AttachMethod, 5));
                p.Add(Str(Pid.DisplayName, f.Message.HeaderDecoded("Subject") ?? name));
                p.Add(Str(Pid.AttachLongFilename, name));
                p.Add(Str(Pid.AttachFilename, ShortName(name)));
                p.Add(Str(Pid.AttachExtension, ".msg"));
                p.Add(FxProperty.FromTime(Pid.CreationTime, now.UtcDateTime));
                p.Add(FxProperty.FromTime(Pid.LastModificationTime, now.UtcDateTime));
                return a;
            }
            if (name == null)
            {
                string ext = f.ContentType.MediaType == "text/calendar" ? ".ics" : f.ContentType.MediaType == "text/html" ? ".html" : f.ContentType.MediaType == "text/plain" ? ".txt" : "";
                name = f.ContentType.MediaType == "application/pkcs7-signature" ? "smime.p7s" : "attachment" + (num + 1) + ext;
            }
            string cid = f.ContentId;
            bool inline = f.Disposition == "inline" && cid != null || (cid != null && f.Disposition == null);
            p.Add(FxProperty.FromInt32(Pid.AttachMethod, 1));
            p.Add(Str(Pid.DisplayName, name));
            p.Add(Str(Pid.AttachLongFilename, name));
            p.Add(Str(Pid.AttachFilename, ShortName(name)));
            p.Add(Str(Pid.AttachExtension, Path.GetExtension(name) ?? ""));
            p.Add(Str(Pid.AttachMimeTag, f.ContentType.MediaType));
            if (cid != null) p.Add(Str(Pid.AttachContentId, cid));
            string loc = f.ContentLocation; if (loc != null) p.Add(Str(Pid.AttachContentLocation, loc));
            p.Add(FxProperty.FromInt32(Pid.AttachFlags, inline ? 4 : 0));
            p.Add(FxProperty.FromBool(Pid.AttachmentHidden, inline));
            p.Add(FxProperty.FromInt32(Pid.AttachSize, (int)Math.Min(int.MaxValue, f.ContentLength)));
            p.Add(FxProperty.FromTime(Pid.CreationTime, now.UtcDateTime));
            p.Add(FxProperty.FromTime(Pid.LastModificationTime, now.UtcDateTime));
            p.Add(f.Content.Length <= int.MaxValue && f.Content.Length < 64 * 1024 ? FxProperty.FromBinary(Pid.AttachDataBin, f.Content.ToArray()) : FxProperty.CreateLazy(Pid.AttachDataBin, 0x0102, f.Content));
            return a;
        }

        /// <summary>8.3-like name for PR_ATTACH_FILENAME.</summary>
        static string ShortName(string n)
        {
            string ext = Path.GetExtension(n) ?? "";
            string stem = Path.GetFileNameWithoutExtension(n) ?? "";
            string clean = new string(stem.Where(ch => ch < 128 && (char.IsLetterOrDigit(ch) || ch == '_')).ToArray());
            if (clean.Length == stem.Length && clean.Length <= 8 && ext.Length <= 4) return n;
            if (clean.Length > 6) clean = clean.Substring(0, 6);
            if (clean.Length == 0) clean = "ATT";
            return clean + "~1" + (ext.Length > 4 ? ext.Substring(0, 4) : ext);
        }
    }
}
