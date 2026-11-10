// =============================================================================
//  Mailbox Transfer - engine, part 18: preview of a message of a PST (the reading pane of the console, the same model as the previews of a mailbox and of a store)
// -----------------------------------------------------------------------------
//  PstReader.GetPreview(nid)   header (subject, sender, To/Cc/Bcc, dates, class, size), the body as text (and as HTML when the message has one), the list of the
//                              attachments (name, type, size, embedded message). Read only: nothing is written, an attachment is never loaded (its size is the size of
//                              its data). The body is cut at BodyChars characters; a message with a body of 40 MiB shows its first characters.
//  The text body is, in this order: PR_BODY, the text of PR_HTML (tags removed), the text of the compressed RTF.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public sealed class PstPreviewAttachment
    {
        public string Name, MimeType, ContentId;
        /// <summary>Size of the data of the attachment in bytes (the size of the embedded message is not known without reading it: -1).</summary>
        public long Size = -1;
        public bool IsMessage;
        public string EmbeddedSubject;
        public bool Inline { get { return !string.IsNullOrEmpty(ContentId); } }
        public override string ToString() { return Name + (Size >= 0 ? " (" + BrowseTotals.FormatBytes(Size) + ")" : ""); }
    }

    public sealed class PstMessagePreview
    {
        public uint Nid;
        public string Subject, MessageClass, FromName, FromAddress;
        public List<string> To = new List<string>(), Cc = new List<string>(), Bcc = new List<string>();
        public DateTime? SentUtc, ReceivedUtc;
        public long SizeBytes = -1;
        public int Importance = 1;
        /// <summary>The body as text, at most BodyChars characters ("" when the message has no body).</summary>
        public string BodyText = "";
        /// <summary>The HTML body when the message has one (at most MaxHtmlChars characters), else null.</summary>
        public string BodyHtml;
        /// <summary>text, html, rtf or none: where BodyText comes from.</summary>
        public string BodySource = "none";
        /// <summary>The body is longer than what is returned.</summary>
        public bool BodyTruncated;
        public List<PstPreviewAttachment> Attachments = new List<PstPreviewAttachment>();
        /// <summary>What could not be read (a body in an unknown format, an unreadable attachment...).</summary>
        public List<string> Notes = new List<string>();
        public string FromDisplay { get { return string.IsNullOrEmpty(FromName) ? FromAddress : string.IsNullOrEmpty(FromAddress) || FromName == FromAddress ? FromName : FromName + " <" + FromAddress + ">"; } }
    }

    public sealed partial class PstReader
    {
        public const int MaxHtmlChars = 1000000;
        const long MaxRtfBytes = 8L << 20;

        /// <summary>The preview of the message of this node. bodyChars: the longest text body returned (default 20,000 characters).</summary>
        public PstMessagePreview GetPreview(uint nid, int bodyChars = 20000)
        {
            FxMessage m = ReadMessage(nid);
            var p = new PstMessagePreview { Nid = nid };
            p.Subject = Printable(m.GetString(0x0037));
            p.MessageClass = m.GetString(0x001A);
            string senderType = m.GetString(0x0C1E);
            p.FromName = m.GetString(0x0C1A);
            p.FromAddress = m.GetString(0x5D01);
            if (string.IsNullOrEmpty(p.FromAddress) && string.Equals(senderType, "SMTP", StringComparison.OrdinalIgnoreCase)) p.FromAddress = m.GetString(0x0C1F);
            if (string.IsNullOrEmpty(p.FromName) && string.IsNullOrEmpty(p.FromAddress)) { p.FromName = m.GetString(0x0042); p.FromAddress = m.GetString(0x5D02); }
            p.SentUtc = TimeOf(m, 0x0039);
            p.ReceivedUtc = TimeOf(m, 0x0E06) ?? p.SentUtc ?? TimeOf(m, 0x3007);
            FxProperty size = m.Find(0x0E08);
            if (size != null && !size.IsMultiValued && size.RawLength >= 4) p.SizeBytes = size.AsInt32();
            FxProperty imp = m.Find(0x0017);
            if (imp != null && !imp.IsMultiValued && imp.RawLength >= 4) p.Importance = imp.AsInt32();

            foreach (FxRecipient r in m.Recipients)
            {
                int type = 1;
                FxProperty t = r.Properties.FirstOrDefault(x => x.Id == 0x0C15 && !x.IsNamed);
                if (t != null && t.RawLength >= 4) type = t.AsInt32();
                string name = Str(r.Properties, 0x3001), smtp = Str(r.Properties, 0x39FE), addr = Str(r.Properties, 0x3003), addrType = Str(r.Properties, 0x3002);
                if (string.IsNullOrEmpty(smtp) && string.Equals(addrType, "SMTP", StringComparison.OrdinalIgnoreCase)) smtp = addr;
                string text = string.IsNullOrEmpty(name) ? smtp : string.IsNullOrEmpty(smtp) || name == smtp ? name : name + " <" + smtp + ">";
                if (string.IsNullOrEmpty(text)) continue;
                (type == 2 ? p.Cc : type == 3 ? p.Bcc : p.To).Add(text);
            }

            FillBody(m, p, bodyChars);

            foreach (FxAttachment a in m.Attachments)
            {
                var pa = new PstPreviewAttachment();
                pa.Name = Str(a.Properties, 0x3707) ?? Str(a.Properties, 0x3704) ?? Str(a.Properties, 0x3001);
                pa.MimeType = Str(a.Properties, 0x370E);
                pa.ContentId = Str(a.Properties, 0x3712);
                if (a.Embedded != null)
                {
                    pa.IsMessage = true;
                    pa.EmbeddedSubject = Printable(a.Embedded.GetString(0x0037));
                    if (string.IsNullOrEmpty(pa.Name)) pa.Name = string.IsNullOrEmpty(pa.EmbeddedSubject) ? "(attached message)" : pa.EmbeddedSubject;
                }
                else
                {
                    FxProperty data = a.Properties.FirstOrDefault(x => x.Id == 0x3701 && !x.IsNamed);
                    FxProperty asz = a.Properties.FirstOrDefault(x => x.Id == 0x0E20 && !x.IsNamed);
                    if (data != null && data.TypeInfo.BaseType != 0x000D) pa.Size = data.RawLength;
                    else if (asz != null && asz.RawLength >= 4) pa.Size = asz.AsInt32();
                }
                if (string.IsNullOrEmpty(pa.Name)) pa.Name = "(no name)";
                p.Attachments.Add(pa);
            }
            return p;
        }

        static string Printable(string s)
        {
            if (s == null) return null;
            var sb = new StringBuilder(s.Length);
            foreach (char c in s) if (c >= ' ') sb.Append(c);
            return sb.ToString();
        }

        static string Str(List<FxProperty> props, ushort id)
        {
            FxProperty p = props.FirstOrDefault(x => x.Id == id && !x.IsNamed);
            if (p == null || p.IsMultiValued || !p.TypeInfo.IsString) return null;
            string s = p.AsString();
            return string.IsNullOrEmpty(s) ? null : s;
        }

        static DateTime? TimeOf(FxMessage m, ushort id)
        {
            FxProperty p = m.Find(id);
            if (p == null || p.IsMultiValued || p.RawLength < 8) return null;
            DateTimeOffset? t = p.AsTime();
            return t.HasValue ? (DateTime?)t.Value.UtcDateTime : null;
        }

        /// <summary>The first <paramref name="maxBytes"/> bytes of a value (lazy values are read from the PST, not loaded whole), and whether there is more.</summary>
        static byte[] Head(FxProperty p, long maxBytes, out bool more)
        {
            more = p.RawLength > maxBytes;
            if (!p.IsLazy) { byte[] raw = p.Raw ?? new byte[0]; if (raw.Length <= maxBytes) return raw; var cut = new byte[maxBytes]; Array.Copy(raw, cut, (int)maxBytes); return cut; }
            using (Stream s = p.OpenValue())
            {
                var buf = new byte[Math.Min(maxBytes, p.RawLength)];
                int n = 0, k;
                while (n < buf.Length && (k = s.Read(buf, n, buf.Length - n)) > 0) n += k;
                if (n < buf.Length) Array.Resize(ref buf, n);
                return buf;
            }
        }

        static string Decode(FxProperty p, byte[] bytes, int internetCodePage)
        {
            FxTypeInfo t = p.TypeInfo;
            if (t.IsString) return FxValue.DecodeString(bytes, t.BaseType, t.CodePageString ? t.CodePage : 1252);
            Encoding enc;
            try { enc = internetCodePage > 0 ? Encoding.GetEncoding(internetCodePage) : new UTF8Encoding(false); } catch (Exception) { enc = new UTF8Encoding(false); }
            return enc.GetString(bytes).TrimEnd('\0');
        }

        static void FillBody(FxMessage m, PstMessagePreview p, int bodyChars)
        {
            int cp = 0;
            FxProperty cpid = m.Find(0x3FDE);
            if (cpid != null && !cpid.IsMultiValued && cpid.RawLength >= 4) cp = cpid.AsInt32();
            try
            {
                FxProperty html = m.Find(0x1013);
                if (html != null && !html.IsMultiValued && html.RawLength > 0)
                {
                    bool more;
                    byte[] hb = Head(html, MaxHtmlChars * 4L, out more);
                    string h = Decode(html, hb, cp);
                    if (h != null) { p.BodyHtml = h.Length > MaxHtmlChars ? h.Substring(0, MaxHtmlChars) : h; if (h.Length > MaxHtmlChars || more) p.Notes.Add("the HTML body is cut after " + MaxHtmlChars + " characters"); }
                }
                FxProperty text = m.Find(0x1000);
                if (text != null && !text.IsMultiValued && text.RawLength > 0)
                {
                    bool more;
                    string t = Decode(text, Head(text, bodyChars * 4L + 8, out more), cp);
                    if (!string.IsNullOrEmpty(t)) { SetText(p, t, "text", bodyChars, more); return; }
                }
                if (p.BodyHtml != null && p.BodyHtml.Length > 0)
                {
                    string t = HtmlText.FromHtml(p.BodyHtml);
                    if (!string.IsNullOrWhiteSpace(t)) { SetText(p, t, "html", bodyChars, false); return; }
                }
                FxProperty rtf = m.Find(0x1009);
                if (rtf != null && !rtf.IsMultiValued && rtf.RawLength > 0)
                {
                    if (rtf.RawLength > MaxRtfBytes) { p.Notes.Add("the RTF body is too large to show (" + BrowseTotals.FormatBytes(rtf.RawLength) + ")"); return; }
                    byte[] plain = Rtf.Decompress(rtf.Raw);
                    string t = plain == null ? null : Rtf.ToText(plain);
                    if (!string.IsNullOrWhiteSpace(t)) { SetText(p, t, "rtf", bodyChars, false); return; }
                    p.Notes.Add("the RTF body could not be turned into text");
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException)) { p.Notes.Add("the body could not be read: " + ex.Message); }
        }

        static void SetText(PstMessagePreview p, string text, string source, int bodyChars, bool moreInSource)
        {
            p.BodySource = source;
            if (bodyChars > 0 && text.Length > bodyChars) { p.BodyText = text.Substring(0, bodyChars); p.BodyTruncated = true; }
            else { p.BodyText = text; p.BodyTruncated = moreInSource; }
        }
    }
}
