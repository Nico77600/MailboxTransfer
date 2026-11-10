// =============================================================================
//  Mailbox Transfer - engine, part 5h: the FastTransfer stream that importItem takes (one construction for every import)
// -----------------------------------------------------------------------------
//  The import from PST and the import of .eml files send the same kind of stream, so one helper builds it and checks it. What the
//  service needs was measured live (on a test tenant, see MailboxTransfer-Guide.md, chapter 7.4):
//    - envelope: one type-1 header record (8 bytes, 00 00 00 00 01 00 00 00) then the data records (type 2); a bare stream is refused
//      (400 ErrorUploadItemsCorruptedData, "unexpected opcode"); the type-4 record an export has is not needed and is left out;
//    - FXDelProp(PidTagMessageRecipients) before the recipients and FXDelProp(PidTagMessageAttachments) before the attachments, on the
//      TOP-LEVEL message only (an Exchange export has none inside an embedded message); without them the service accepts the item
//      and silently drops the recipients and the attachments;
//    - PidTagRowid (0x3000) is the first property of every recipient, otherwise 500 ErrorInternalServerError;
//    - an embedded message starts with PidTagMid (0x674A, Int64, any value) and its attachment has no empty PidTagAttachDataObject
//      (0x3701) marker before StartEmbed, otherwise 500.
//  Prepare enforces these rules on a message model; Build prepares, encodes, and decodes the stream again with our own decoder to
//  assert that the recipients, attachments and embedded messages are all there and that the rules hold (so that a silent loss can
//  never be produced by us). The stream that Exchange itself made (the Exchange conversion of an .eml) is not rebuilt here.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Linq;

namespace MailboxTransfer
{
    /// <summary>Recipients, attachments and embedded messages of a message, embedded messages included.</summary>
    public sealed class FxImportCounts
    {
        public int Recipients, Attachments, Embedded;

        public static FxImportCounts Of(FxMessage m)
        {
            var c = new FxImportCounts();
            FxDiff.Totals(m, out c.Recipients, out c.Attachments, out c.Embedded);
            return c;
        }

        public bool SameAs(FxImportCounts o) { return o != null && Recipients == o.Recipients && Attachments == o.Attachments && Embedded == o.Embedded; }
        public override string ToString() { return Recipients + " recipients, " + Attachments + " attachments, " + Embedded + " embedded messages"; }
    }

    /// <summary>The stream we built does not say what the model says: nothing is sent.</summary>
    public sealed class FxImportCheckException : Exception
    {
        public FxImportCheckException(string message) : base(message) { }
    }

    public static class FxImportStream
    {
        public const ushort TagRowId = 0x3000, TagMid = 0x674A, TagAttachDataObject = 0x3701, TagFxDelProp = 0x4016;

        /// <summary>The envelope importItem needs in front of the stream: one header record, the data records follow.</summary>
        public static FxEnvelope NewEnvelope()
        {
            var env = new FxEnvelope();
            env.Entries.Add(new FxEnvelopeEntry { Type = 1, Length = 8, Payload = new byte[] { 0, 0, 0, 0, 1, 0, 0, 0 } });
            return env;
        }

        static FxProperty DelProp(uint tag) { return FxProperty.Create(TagFxDelProp, 0x0003, BitConverter.GetBytes(tag)); }

        /// <summary>Makes a message ready for importItem (changes it in place and returns it). Safe to call again.</summary>
        public static FxMessage Prepare(FxMessage m)
        {
            Prepare(m, false, new long[1]);
            return m;
        }

        static void Prepare(FxMessage m, bool embedded, long[] midCounter)
        {
            if (embedded)
            {
                m.DelPropRecipients = null;
                m.DelPropAttachments = null;
                int mid = m.Properties.FindIndex(p => p.Id == TagMid && !p.IsNamed);
                if (mid < 0) m.Properties.Insert(0, FxProperty.FromInt64(TagMid, ++midCounter[0]));
                else if (mid > 0) { FxProperty p = m.Properties[mid]; m.Properties.RemoveAt(mid); m.Properties.Insert(0, p); }
            }
            else
            {
                if (m.DelPropRecipients == null) m.DelPropRecipients = DelProp(FxMessage.TagRecipients);
                if (m.DelPropAttachments == null) m.DelPropAttachments = DelProp(FxMessage.TagAttachments);
            }
            for (int i = 0; i < m.Recipients.Count; i++)
            {
                List<FxProperty> ps = m.Recipients[i].Properties;
                int k = ps.FindIndex(p => p.Id == TagRowId && !p.IsNamed);
                if (k > 0) { FxProperty row = ps[k]; ps.RemoveAt(k); ps.Insert(0, row); }
                else if (k < 0) ps.Insert(0, FxProperty.FromInt32(TagRowId, i));
            }
            foreach (FxAttachment a in m.Attachments)
            {
                if (a.Embedded == null) continue;
                a.Properties.RemoveAll(p => p.Id == TagAttachDataObject && !p.IsNamed);
                Prepare(a.Embedded, true, midCounter);
            }
        }

        /// <summary>
        /// The source of the stream to send: the message is prepared, every property is validated (ArgumentException when it cannot be
        /// expressed), the stream is built with the envelope and checked (FxImportCheckException). Nothing has been sent when it throws.
        /// </summary>
        public static IFxStreamSource Build(FxMessage m)
        {
            Prepare(m);
            FxCodec.PayloadLength(m);
            var source = new FxMessageSource(m, NewEnvelope());
            Verify(source, m);
            return source;
        }

        /// <summary>
        /// Decodes the stream with our decoder and asserts the rules above and that its recipients, attachments and embedded messages equal
        /// those of the model.
        /// </summary>
        public static void Verify(IFxStreamSource source, FxMessage model)
        {
            FxDocument doc = FxCodec.Decode(source);
            if (doc.Error != null || doc.Truncated) throw new FxImportCheckException("the stream built does not decode: " + (doc.Error ?? "truncated") + " (offset " + doc.ErrorOffset + ")");
            FxEnvelope env = doc.Envelope;
            if (env == null || env.Entries.Count == 0 || env.Entries[0].Type != 1 || env.Entries[0].Length != 8) throw new FxImportCheckException("the stream built does not start with the type-1 header record");
            if (env.Entries.Any(e => e.Type == 4)) throw new FxImportCheckException("the stream built carries a type-4 record");
            FxMessage got = doc.Message;
            if (got.DelPropRecipients == null || got.DelPropAttachments == null) throw new FxImportCheckException("the stream built has no FXDelProp before the recipients or before the attachments: the service would drop them");
            FxImportCounts want = FxImportCounts.Of(model), have = FxImportCounts.Of(got);
            if (!want.SameAs(have)) throw new FxImportCheckException("the stream built holds (" + have + ") where the message has (" + want + ")");
            CheckRules(got, false);
        }

        static void CheckRules(FxMessage m, bool embedded)
        {
            if (embedded)
            {
                if (m.DelPropRecipients != null || m.DelPropAttachments != null) throw new FxImportCheckException("an embedded message carries FXDelProp");
                if (m.Properties.Count == 0 || m.Properties[0].Id != TagMid || m.Properties[0].IsNamed) throw new FxImportCheckException("an embedded message does not start with PidTagMid");
            }
            foreach (FxRecipient r in m.Recipients)
                if (r.Properties.Count == 0 || r.Properties[0].Id != TagRowId || r.Properties[0].IsNamed) throw new FxImportCheckException("a recipient does not start with PidTagRowid");
            foreach (FxAttachment a in m.Attachments)
            {
                if (a.Embedded == null) continue;
                if (a.Properties.Any(p => p.Id == TagAttachDataObject && !p.IsNamed)) throw new FxImportCheckException("an embedded message attachment carries a PidTagAttachDataObject marker");
                CheckRules(a.Embedded, true);
            }
        }
    }
}
