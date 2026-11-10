// =============================================================================
//  Mailbox Transfer - engine, part 5: readable summary of an exported message
// -----------------------------------------------------------------------------
//  Built on the elements found by FxReader (Engine.FastTransfer.cs); the exact object model (FxMessage, FxCodec) is in
//  Engine.FxModel.cs. Properties are matched on their id only, whatever the type:
//  a subject can be PtypString (0x001F), PtypString8 (0x001E) or a code page type (0x84B0).
//  Nothing here loads an attachment: only the size of its data is kept.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public sealed class FxRecipientSummary
    {
        public int Type;
        public string TypeName;
        public string DisplayName;
        public string AddressType;
        public string Address;
        public string Smtp;
    }

    public sealed class FxAttachmentSummary
    {
        public string FileName;
        public string DisplayName;
        public string MimeTag;
        public string ContentId;
        public int? Method;
        /// <summary>PidTagAttachSize as stored by Exchange, -1 if absent.</summary>
        public long Size = -1;
        /// <summary>Size of PidTagAttachDataBinary in the stream, -1 if absent.</summary>
        public long DataBytes = -1;
        public bool DataTruncated;
        public bool EmbeddedMessage;
        public string EmbeddedSubject;
    }

    public sealed class FxMessageSummary
    {
        public string Envelope;
        public long StreamBytes;
        public int PropertyCount;
        public int MarkerCount;
        public bool Truncated;
        public string Error;

        public string Subject;
        public string MessageClass;
        public string SenderName, SenderAddressType, SenderAddress, SenderSmtp;
        public string SentRepresentingName, SentRepresentingAddress, SentRepresentingSmtp;
        public DateTimeOffset? ClientSubmitTime, DeliveryTime, CreationTime, LastModificationTime;
        public int? Importance, Sensitivity;
        public long? MessageSize;
        public bool? HasAttachments;
        public string InternetMessageId;

        public string BodyTextPreview;
        public long BodyTextBytes = -1;
        public string BodyHtmlPreview;
        public long BodyHtmlBytes = -1;
        public bool HasCompressedRtf;
        public long RtfBytes = -1;

        public List<FxRecipientSummary> Recipients = new List<FxRecipientSummary>();
        public List<FxAttachmentSummary> Attachments = new List<FxAttachmentSummary>();

        /// <summary>Folder or location properties found at message level (PidTagParentFolderId, PidTagFolderId, ...).</summary>
        public List<string> FolderInfoProperties = new List<string>();
        /// <summary>True when PidTagFolderId, PidTagParentFolderId or PidTagParentSourceKey is present: the folder is in the stream.</summary>
        public bool HasFolderInfo;
        /// <summary>Named properties found at message level (categories, ...), as text.</summary>
        public List<string> NamedProperties = new List<string>();
    }

    /// <summary>The object model is in Engine.FxModel.cs; this part only holds the summary of a flat read (Summarize).</summary>
    public sealed partial class FxMessage
    {
        const int PreviewChars = 400;

        static string Str(FxElement e) { return e.Value as string; }

        static string Excerpt(string s)
        {
            if (s == null) return null;
            s = s.Replace("\r", " ").Replace("\n", " ");
            return s.Length > PreviewChars ? s.Substring(0, PreviewChars) + "..." : s;
        }

        static DateTimeOffset? Time(FxElement e) { return e.Value is DateTimeOffset ? (DateTimeOffset?)e.Value : null; }

        static int? Int(FxElement e)
        {
            if (e.Value is int) return (int)e.Value;
            if (e.Value is short) return (short)e.Value;
            return null;
        }

        public static FxMessageSummary Summarize(FxReadResult r)
        {
            var m = new FxMessageSummary
            {
                Envelope = r.Envelope,
                StreamBytes = r.BytesRead,
                PropertyCount = r.PropertyCount,
                MarkerCount = r.MarkerCount,
                Truncated = r.Truncated,
                Error = r.Error
            };

            var recipients = new Dictionary<string, FxRecipientSummary>();
            var attachments = new Dictionary<string, FxAttachmentSummary>();

            foreach (FxElement e in r.Elements)
            {
                if (e.IsMarker)
                {
                    if (e.Name == "StartEmbed")
                    {
                        FxAttachmentSummary a = AttachmentOf(e.Path, attachments);
                        if (a != null) a.EmbeddedMessage = true;
                    }
                    continue;
                }

                if (e.Path == "message") MessageLevel(m, e);
                else if (e.Path.StartsWith("message/recipient[", StringComparison.Ordinal) && e.Path.IndexOf('/', 8) < 0)
                {
                    FxRecipientSummary rc;
                    if (!recipients.TryGetValue(e.Path, out rc))
                    {
                        rc = new FxRecipientSummary { TypeName = "?" };
                        recipients[e.Path] = rc;
                        m.Recipients.Add(rc);
                    }
                    RecipientLevel(rc, e);
                }
                else if (e.Path.StartsWith("message/attachment[", StringComparison.Ordinal))
                {
                    FxAttachmentSummary a = AttachmentOf(e.Path, attachments);
                    if (a != null && !m.Attachments.Contains(a)) m.Attachments.Add(a);
                    if (e.Path.IndexOf("/embedded", StringComparison.Ordinal) > 0)
                    {
                        if (e.Id == 0x0037 && a != null && a.EmbeddedSubject == null) a.EmbeddedSubject = Str(e);
                    }
                    else if (a != null) AttachmentLevel(a, e);
                }
            }

            // attachments of which only the marker was seen are still attachments
            foreach (FxAttachmentSummary a in attachments.Values) if (!m.Attachments.Contains(a)) m.Attachments.Add(a);
            return m;
        }

        /// <summary>The attachment of "message/attachment[n]" or of a path below it (embedded message).</summary>
        static FxAttachmentSummary AttachmentOf(string path, Dictionary<string, FxAttachmentSummary> map)
        {
            if (!path.StartsWith("message/attachment[", StringComparison.Ordinal)) return null;
            int slash = path.IndexOf('/', "message/".Length);
            string key = slash < 0 ? path : path.Substring(0, slash);
            FxAttachmentSummary a;
            if (!map.TryGetValue(key, out a)) { a = new FxAttachmentSummary(); map[key] = a; }
            return a;
        }

        static void MessageLevel(FxMessageSummary m, FxElement e)
        {
            switch (e.Id)
            {
                case 0x0037: m.Subject = Str(e); break;
                case 0x001A: m.MessageClass = Str(e); break;
                case 0x0C1A: m.SenderName = Str(e); break;
                case 0x0C1E: m.SenderAddressType = Str(e); break;
                case 0x0C1F: m.SenderAddress = Str(e); break;
                case 0x5D01: m.SenderSmtp = Str(e); break;
                case 0x0042: m.SentRepresentingName = Str(e); break;
                case 0x0065: m.SentRepresentingAddress = Str(e); break;
                case 0x5D02: m.SentRepresentingSmtp = Str(e); break;
                case 0x0039: m.ClientSubmitTime = Time(e); break;
                case 0x0E06: m.DeliveryTime = Time(e); break;
                case 0x3007: m.CreationTime = Time(e); break;
                case 0x3008: m.LastModificationTime = Time(e); break;
                case 0x0017: m.Importance = Int(e); break;
                case 0x0036: m.Sensitivity = Int(e); break;
                case 0x0E08: m.MessageSize = Int(e); break;
                case 0x0E1B: m.HasAttachments = e.Value is bool ? (bool?)e.Value : null; break;
                case 0x1035: m.InternetMessageId = Str(e); break;
                case 0x1000:
                    m.BodyTextBytes = e.ValueBytes;
                    m.BodyTextPreview = Excerpt(Str(e));
                    break;
                case 0x1013:   // binary (HTML in the message code page) or string
                    m.BodyHtmlBytes = e.ValueBytes;
                    m.BodyHtmlPreview = Excerpt(Str(e) ?? (e.Preview == null ? null : Encoding.UTF8.GetString(e.Preview)));
                    break;
                case 0x1009:
                    m.HasCompressedRtf = true;
                    m.RtfBytes = e.ValueBytes;
                    break;
            }

            if (Array.IndexOf(FxNames.FolderInfoIds, e.Id) >= 0 && !e.IsNamed)
            {
                m.FolderInfoProperties.Add((FxNames.Name(e.Id) ?? "id") + " " + e.TagText + " = " + e.Display);
                if (e.Id == 0x6748 || e.Id == 0x6749 || e.Id == 0x65E1) m.HasFolderInfo = true;
            }
            if (e.IsNamed)
                m.NamedProperties.Add(e.TagText + " " + e.NamedText + " (" + e.TypeName + ") = " + e.Display);
        }

        static void RecipientLevel(FxRecipientSummary rc, FxElement e)
        {
            switch (e.Id)
            {
                case 0x0C15:
                    int? t = Int(e);
                    rc.Type = t ?? 0;
                    rc.TypeName = t == 1 ? "To" : t == 2 ? "Cc" : t == 3 ? "Bcc" : "Type" + t;
                    break;
                case 0x3001: rc.DisplayName = Str(e); break;
                case 0x3002: rc.AddressType = Str(e); break;
                case 0x3003: rc.Address = Str(e); break;
                case 0x39FE: rc.Smtp = Str(e); break;
            }
        }

        static void AttachmentLevel(FxAttachmentSummary a, FxElement e)
        {
            switch (e.Id)
            {
                case 0x3707: a.FileName = Str(e); break;
                case 0x3704: if (a.FileName == null) a.FileName = Str(e); break;
                case 0x3001: a.DisplayName = Str(e); break;
                case 0x370E: a.MimeTag = Str(e); break;
                case 0x3712: a.ContentId = Str(e); break;
                case 0x3705: a.Method = Int(e); break;
                case 0x0E20:
                    int? size = Int(e);
                    if (size != null) a.Size = size.Value;
                    break;
                case 0x3701:
                    if ((e.Type & 0x0FFF) == 0x000D) a.EmbeddedMessage = true;
                    else { a.DataBytes = e.ValueBytes; a.DataTruncated = e.Incomplete; }
                    break;
            }
        }
    }
}
