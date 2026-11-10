// =============================================================================
//  Mailbox Transfer - engine, part 4: FastTransfer stream reader (MS-OXCFXICS)
// -----------------------------------------------------------------------------
//  The data of exportItems is "opaque" for Microsoft (not meant to be parsed). This reader exists to LOOK at it
//  (what is inside, is there folder information, is there an envelope), never to rewrite it: the file that goes
//  to importItem is always the exported file, byte for byte.
//
//  Format, from MS-OXCFXICS 2.2.4 (and the Open Specifications blog "How to parse the FastTransfer Stream"):
//    stream  = element*
//    element = marker | propValue
//    marker  = 4 bytes (little endian) found in the table of 2.2.4.1.4 (StartMessage, NewAttach, ...)
//    propValue = tag (type: 2 bytes, id: 2 bytes, little endian)
//                [named property: GUID (16) + kind (1): 0 = LID (4 bytes), 1 = name (UTF-16LE, null terminated)]
//                value
//    value   = fixed size (int16 2, int32/float32/error 4, int64/double/currency/time 8, bool 2 (not 1!), guid 16)
//            | variable: 4 bytes length + bytes (string8, string UTF-16LE with null, binary, server id)
//            | multi-valued (type | 0x1000): 4 bytes count + values
//            | code page string (type >= 0x8000, e.g. 0x84B0 = code page 1200 = UTF-16LE): 4 bytes length + bytes
//            | object (0x000D): no bytes, a marker (StartEmbed) follows
//
//  The Graph export adds an envelope in front of the stream (observed in the Learn examples of importItem and on a real
//  export of 100 MiB): records "type (4 bytes) length (4 bytes) payload", the first one being 01 00 00 00 08 00 00 00
//  (type 1, 8 bytes), then type 4 (6 bytes) and sometimes type 3 (24 bytes, the GUID of IMessage). The data is in the
//  records of type 2: they are consecutive pieces of ONE FastTransfer stream (3311 pieces of 31,680 bytes, the last one
//  shorter, for the real export of a 100 MiB attachment), a piece may end in the middle of a property, so the parser reads
//  their concatenation. This is a deduction from samples, not documented: the reader reports the records it finds and
//  tolerates their absence.
//
//  Two things seen on real streams that MS-OXCFXICS does not describe, both handled:
//    - a tag whose type has the bit 0x4000 (PtypBinary 0x4102, PtypTime 0x4040, ...) is followed by two bytes (0x0200 so far)
//      before the value; the type without the bit is the real type;
//    - the properties are not sorted by tag: the body comes after the named properties, just before the recipients.
//  The markers (2.2.4.1.4 on Microsoft Learn) are the ones of the table below: all checked, none was wrong. FXDelProp is
//  followed by a tag (0x0E12000D before the recipients, 0x0E13000D before the attachments).
//
//  A code page type with the bit 0x1000 (0x94B0 = multi-valued string in code page 1200, seen on a real export) is a multi-valued
//  code page string: reading it as a single string made the stream drift (the count was taken for a length) up to a tag with
//  type 0x0000 (see FxTypes.Describe).
//
//  The lexical structure (FxLexer: tags, named properties, every type, multi-valued) is shared by two readers: FxReader below,
//  which looks at a stream (flat list of elements, previews, an attachment is read and discarded, only its first
//  bytes are kept) and never throws, and FxCodec (Engine.FxModel.cs), which keeps every byte to build the object model and encode it
//  again. Unknown types stop the walk (no way to find the next element) and are reported in Error; a truncated stream is reported
//  in Truncated.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public static class FxNames
    {
        /// <summary>
        /// Markers of 2.2.4.1.4: 4 bytes, never a value. MetaTagFXDelProp (0x40160003), MetaTagEcWarning (0x400F0003) and
        /// MetaTagNewFXFolder (0x40110102) are meta-PROPERTIES with a value (the tag deleted, the error code, the replica info):
        /// the lexer reads them like any property; FxReader shows FXDelProp as a marker (as it always did).
        /// </summary>
        public static readonly Dictionary<uint, string> Markers = new Dictionary<uint, string>
        {
            { 0x40000003, "NewAttach" }, { 0x40010003, "StartEmbed" }, { 0x40020003, "EndEmbed" },
            { 0x40030003, "StartRecip" }, { 0x40040003, "EndToRecip" }, { 0x40090003, "StartTopFld" },
            { 0x400A0003, "StartSubFld" }, { 0x400B0003, "EndFolder" }, { 0x400C0003, "StartMessage" },
            { 0x400D0003, "EndMessage" }, { 0x400E0003, "EndAttach" },
            { 0x40100003, "StartFAIMsg" }, { 0x40120003, "IncrSyncChg" },
            { 0x40130003, "IncrSyncDel" }, { 0x40140003, "IncrSyncEnd" }, { 0x40150003, "IncrSyncMessage" },
            { 0x402F0003, "IncrSyncRead" }, { 0x403A0003, "IncrSyncStateBegin" },
            { 0x403B0003, "IncrSyncStateEnd" }, { 0x407D0003, "IncrSyncChgPartial" }, { 0x40180003, "FXErrorInfo" },
            { 0x4074000B, "IncrSyncProgressMode" }, { 0x4075000B, "IncrSyncProgressPerMsg" }, { 0x407B0102, "IncrSyncGroupInfo" }
        };

        /// <summary>MetaTagFXDelProp: a property whose Int32 value is the tag to delete (before the recipients and the attachments).</summary>
        public const uint FXDelProp = 0x40160003;

        static readonly Dictionary<ushort, string> Ids = new Dictionary<ushort, string>
        {
            { 0x0017, "PidTagImportance" }, { 0x001A, "PidTagMessageClass" }, { 0x0036, "PidTagSensitivity" },
            { 0x0037, "PidTagSubject" }, { 0x0039, "PidTagClientSubmitTime" }, { 0x003D, "PidTagSubjectPrefix" },
            { 0x0042, "PidTagSentRepresentingName" }, { 0x0064, "PidTagSentRepresentingAddressType" },
            { 0x0065, "PidTagSentRepresentingEmailAddress" }, { 0x0070, "PidTagConversationTopic" },
            { 0x0E02, "PidTagDisplayBcc" }, { 0x0E03, "PidTagDisplayCc" }, { 0x0E04, "PidTagDisplayTo" },
            { 0x0E06, "PidTagMessageDeliveryTime" }, { 0x0E07, "PidTagMessageFlags" }, { 0x0E08, "PidTagMessageSize" },
            { 0x0E09, "PidTagParentEntryId" }, { 0x0E1B, "PidTagHasAttachments" }, { 0x0E1D, "PidTagNormalizedSubject" },
            { 0x0E1F, "PidTagRtfInSync" }, { 0x0E20, "PidTagAttachSize" }, { 0x0C15, "PidTagRecipientType" },
            { 0x0C1A, "PidTagSenderName" }, { 0x0C1E, "PidTagSenderAddressType" }, { 0x0C1F, "PidTagSenderEmailAddress" },
            { 0x0FF6, "PidTagInstanceKey" }, { 0x0FF9, "PidTagRecordKey" }, { 0x0FFB, "PidTagStoreEntryId" },
            { 0x0FFF, "PidTagEntryId" }, { 0x1000, "PidTagBody" }, { 0x1009, "PidTagRtfCompressed" },
            { 0x1013, "PidTagHtml" }, { 0x1035, "PidTagInternetMessageId" }, { 0x3001, "PidTagDisplayName" },
            { 0x3002, "PidTagAddressType" }, { 0x3003, "PidTagEmailAddress" }, { 0x3007, "PidTagCreationTime" },
            { 0x3008, "PidTagLastModificationTime" }, { 0x300B, "PidTagSearchKey" }, { 0x39FE, "PidTagSmtpAddress" },
            { 0x3701, "PidTagAttachDataBinary/Object" }, { 0x3704, "PidTagAttachFilename" },
            { 0x3705, "PidTagAttachMethod" }, { 0x3707, "PidTagAttachLongFilename" }, { 0x370B, "PidTagRenderingPosition" },
            { 0x370E, "PidTagAttachMimeTag" }, { 0x3712, "PidTagAttachContentId" }, { 0x3713, "PidTagAttachContentLocation" },
            { 0x3714, "PidTagAttachFlags" }, { 0x3FDE, "PidTagInternetCodepage" }, { 0x3FFD, "PidTagMessageCodepage" },
            { 0x5D01, "PidTagSenderSmtpAddress" }, { 0x5D02, "PidTagSentRepresentingSmtpAddress" },
            { 0x6748, "PidTagFolderId" }, { 0x6749, "PidTagParentFolderId" }, { 0x674A, "PidTagMid" },
            { 0x674D, "PidTagInstID" }, { 0x674E, "PidTagInstanceNum" }, { 0x67A4, "PidTagChangeNumber" },
            { 0x65E0, "PidTagSourceKey" }, { 0x65E1, "PidTagParentSourceKey" }, { 0x65E2, "PidTagChangeKey" },
            { 0x65E3, "PidTagPredecessorChangeList" }, { 0x6639, "PidTagRights" }, { 0x7FFE, "PidTagAttachmentHidden" }
        };

        /// <summary>Properties that say where an item lives (folder, source key, entry id): searched by the summary.</summary>
        public static readonly ushort[] FolderInfoIds =
            { 0x6748, 0x6749, 0x674A, 0x67A4, 0x674D, 0x674E, 0x65E0, 0x65E1, 0x65E2, 0x65E3, 0x0E09, 0x0FFF, 0x0FFB, 0x0FF9, 0x0FF6, 0x300B };

        public static string Name(ushort id) { string n; return Ids.TryGetValue(id, out n) ? n : null; }

        public static string TypeName(ushort type)
        {
            FxTypeInfo info = FxTypes.Describe(type);
            if (info.CodePageString) return "CodePage(" + info.CodePage + ")" + (info.MultiValued ? "Multiple" : "");
            string mv = info.MultiValued ? "Multiple" : "";
            switch (type & 0x0FFF)
            {
                case 0x0001: return "PtypNull";
                case 0x0002: return "PtypInteger16" + mv;
                case 0x0003: return "PtypInteger32" + mv;
                case 0x0004: return "PtypFloating32" + mv;
                case 0x0005: return "PtypFloating64" + mv;
                case 0x0006: return "PtypCurrency" + mv;
                case 0x0007: return "PtypFloatingTime" + mv;
                case 0x000A: return "PtypErrorCode" + mv;
                case 0x000B: return "PtypBoolean" + mv;
                case 0x000D: return "PtypObject";
                case 0x0014: return "PtypInteger64" + mv;
                case 0x001E: return "PtypString8" + mv;
                case 0x001F: return "PtypString" + mv;
                case 0x0040: return "PtypTime" + mv;
                case 0x0048: return "PtypGuid" + mv;
                case 0x00FB: return "PtypServerId" + mv;
                case 0x0102: return "PtypBinary" + mv;
            }
            return "Unknown(0x" + type.ToString("X4") + ")";
        }

        /// <summary>The code page of a code page string type (0x84B0 is 1200, 0xFDE9 is 65001), 0 for another type.</summary>
        public static int CodePageOf(ushort type) { return FxTypes.Describe(type).CodePage; }
    }

    public sealed class FxElement
    {
        public bool IsMarker;
        public long Offset;
        public uint Tag;
        public string Name;
        public string Path;
        public ushort Type;
        public ushort Id;
        public string TypeName;
        public Guid? NamedGuid;
        public long? NamedLid;
        public string NamedString;
        /// <summary>Size in bytes of the value as encoded (an attachment: its size), length prefix excluded.</summary>
        public long ValueBytes;
        /// <summary>Decoded value: number, bool, DateTimeOffset, Guid, string (preview), object[] for multi-valued; null for binary.</summary>
        public object Value;
        /// <summary>First bytes of a binary or string value.</summary>
        public byte[] Preview;
        /// <summary>The value is longer than what is kept (preview) or the stream ends inside it.</summary>
        public bool Truncated;
        /// <summary>The stream ends inside this value.</summary>
        public bool Incomplete;
        /// <summary>The tag carried the bit 0x4000 in its type and a 2-byte field before the value (ExtendedFlags): see ParseElements.</summary>
        public bool Extended;
        public ushort ExtendedFlags;
        public bool IsNamed { get { return Id >= 0x8000; } }

        public string Display
        {
            get
            {
                if (IsMarker) return "";
                object v = Value;
                if (v == null) return Preview == null ? "" : ValueBytes + " bytes: " + FxReader.Hex(Preview, 16);
                var list = v as object[];
                if (list != null) return "[" + string.Join("; ", list.Take(5).Select(Show)) + (list.Length > 5 ? "; ..." : "") + "] (" + list.Length + ")";
                return Show(v) + (Truncated ? " ..." : "");
            }
        }

        static string Show(object v)
        {
            var s = v as string;
            if (s != null) return "\"" + (s.Length > 80 ? s.Substring(0, 80) + "..." : s).Replace("\r", "\\r").Replace("\n", "\\n") + "\"";
            if (v is DateTimeOffset) return ((DateTimeOffset)v).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture);
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        public string TagText { get { return "0x" + Tag.ToString("X8"); } }

        public string NamedText
        {
            get
            {
                if (!IsNamed || NamedGuid == null) return null;
                return NamedGuid + (NamedLid != null ? " lid 0x" + NamedLid.Value.ToString("X") : NamedString != null ? " name " + NamedString : "");
            }
        }
    }

    public sealed class FxEnvelopeRecord
    {
        public uint Type;
        public long Length;
        public long Offset;
        public bool IsData;
        public string HeadHex;
    }

    public sealed class FxReadResult
    {
        /// <summary>None: the file starts with the FastTransfer stream. Tlv: records in front of it (see header of this file).</summary>
        public string Envelope = "None";
        public List<FxEnvelopeRecord> EnvelopeRecords = new List<FxEnvelopeRecord>();
        /// <summary>Offset in the file of the payload of the first data record.</summary>
        public long DataOffset;
        /// <summary>Total announced length of the data records read (-1 if none), and their number.</summary>
        public long DataLength = -1;
        public int DataRecords;
        /// <summary>Records seen in the envelope; EnvelopeRecords keeps only the first 64 of them.</summary>
        public int EnvelopeRecordCount;
        public long BytesRead;
        public string HeadHex;
        public List<FxElement> Elements = new List<FxElement>();
        public bool Truncated;
        public string Error;
        public long ErrorOffset;
        public int PropertyCount { get { return Elements.Count(e => !e.IsMarker); } }
        public int MarkerCount { get { return Elements.Count(e => e.IsMarker); } }
    }

    public sealed class FxReaderOptions
    {
        public int MaxPreviewBytes = 16384;
        public int MaxMultiValues = 64;
    }

    /// <summary>Forward-only byte source with a peek and an optional end (the end of the data record).</summary>
    sealed class FxSource
    {
        readonly Stream _s;
        byte[] _buf = new byte[64 * 1024];
        int _pos, _len;
        public long Position { get; private set; }
        public long Limit = long.MaxValue;

        public FxSource(Stream s) { _s = s; }

        public long Remaining { get { return Limit == long.MaxValue ? long.MaxValue : Limit - Position; } }

        bool Fill()
        {
            if (_pos < _len) return true;
            _pos = 0;
            _len = _s.Read(_buf, 0, _buf.Length);
            return _len > 0;
        }

        public byte[] Peek(int n)
        {
            if (_len - _pos < n)
            {
                Buffer.BlockCopy(_buf, _pos, _buf, 0, _len - _pos);
                _len -= _pos; _pos = 0;
                while (_len < n)
                {
                    int got = _s.Read(_buf, _len, _buf.Length - _len);
                    if (got <= 0) break;
                    _len += got;
                }
            }
            int count = Math.Min(n, _len - _pos);
            var r = new byte[count];
            Buffer.BlockCopy(_buf, _pos, r, 0, count);
            return r;
        }

        /// <summary>Reads up to count bytes, never past the limit. Returns the number read (less only at the end).</summary>
        public int Read(byte[] dst, int offset, int count)
        {
            if (count > Remaining) count = (int)Remaining;
            int total = 0;
            while (total < count && Fill())
            {
                int n = Math.Min(count - total, _len - _pos);
                Buffer.BlockCopy(_buf, _pos, dst, offset + total, n);
                _pos += n; total += n;
            }
            Position += total;
            return total;
        }

        public byte[] ReadBytes(int count)
        {
            var b = new byte[count];
            int n = Read(b, 0, count);
            if (n < count) Array.Resize(ref b, n);
            return b;
        }

        public long Skip(long count)
        {
            if (count > Remaining) count = Remaining;
            long total = 0;
            while (total < count && Fill())
            {
                int n = (int)Math.Min(count - total, _len - _pos);
                _pos += n; total += n;
            }
            Position += total;
            return total;
        }
    }

    /// <summary>Reads the envelope records and exposes the concatenated payloads of the records of type 2 as one stream.</summary>
    sealed class FxEnvelopeStream : Stream
    {
        const int MaxKeptRecords = 64;
        /// <summary>Largest payload kept for a record that is not data (they are 8, 24 and 6 bytes in the exports seen).</summary>
        const int MaxKeptPayload = 1 << 20;
        readonly FxSource _raw;
        readonly FxReadResult _result;
        long _left;
        public bool Truncated;
        /// <summary>Every record read so far (the model re-encodes the envelope from them).</summary>
        public readonly List<FxEnvelopeEntry> Entries = new List<FxEnvelopeEntry>();

        public FxEnvelopeStream(FxSource raw, FxReadResult result) { _raw = raw; _result = result; }

        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { throw new NotSupportedException(); } }
        public override long Position { get { throw new NotSupportedException(); } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }

        public override int Read(byte[] buffer, int offset, int count)
        {
            while (_left == 0)
                if (!NextRecord()) return 0;
            int n = _raw.Read(buffer, offset, (int)Math.Min(count, _left));
            _left -= n;
            if (n == 0) { Truncated = true; _left = 0; return 0; }
            return n;
        }

        /// <summary>Moves to the payload of the next record of type 2 (the others are noted and skipped). False at the end.</summary>
        bool NextRecord()
        {
            while (true)
            {
                long offset = _raw.Position;
                byte[] h = _raw.ReadBytes(8);
                if (h.Length == 0) return false;
                if (h.Length < 8) { Truncated = true; return false; }
                var rec = new FxEnvelopeRecord { Type = BitConverter.ToUInt32(h, 0), Length = BitConverter.ToUInt32(h, 4), Offset = offset };
                rec.IsData = rec.Type == 2;
                var entry = new FxEnvelopeEntry { Type = rec.Type, Length = rec.Length, RawOffset = _raw.Position };
                Entries.Add(entry);
                if (rec.IsData)
                {
                    if (_result.DataRecords == 0) _result.DataOffset = _raw.Position;
                    _result.DataRecords++;
                    _result.DataLength = Math.Max(0, _result.DataLength) + rec.Length;
                }
                _result.EnvelopeRecordCount++;
                if (_result.EnvelopeRecords.Count < MaxKeptRecords) _result.EnvelopeRecords.Add(rec);
                if (rec.IsData) { _left = rec.Length; if (_left > 0) return true; continue; }
                int keep = (int)Math.Min(rec.Length, MaxKeptPayload);
                byte[] payload = _raw.ReadBytes(keep);
                entry.Payload = payload;
                rec.HeadHex = FxReader.Hex(payload.Length > 64 ? payload.Take(64).ToArray() : payload, 64);
                long rest = rec.Length - payload.Length;
                if (payload.Length < keep || _raw.Skip(rest) < rest) { Truncated = true; return false; }
            }
        }
    }

    /// <summary>One element of the lexical structure: a marker (4 bytes) or a property with its value.</summary>
    sealed class FxToken
    {
        public bool IsMarker;
        public uint Tag;
        public long Offset;
        public FxProperty Property;
        /// <summary>The lexer failed in this token (stream cut or unknown type): the value is not usable.</summary>
        public bool Bad;
    }

    /// <summary>
    /// The lexical structure of MS-OXCFXICS 2.2.4.1 on a byte source: tag (type, id), the 2 bytes of an extended type, the
    /// identity of a named property, then the value (fixed size, length + bytes, count + values, object). Values longer than
    /// the limits are read and discarded, only their first bytes are kept (preview); no limit keeps everything (the model).
    /// Never throws: after a problem (Truncated, Error) Next returns null; the token whose value was cut is returned first.
    /// </summary>
    sealed class FxLexer
    {
        /// <summary>A multi-valued property with more values than this is taken for garbage.</summary>
        public const uint MaxMultiValues = 1000000;
        /// <summary>A value larger than this is read in pieces of this size.</summary>
        const long PieceSize = 1L << 20;
        const int MaxNamedChars = 16384;

        readonly FxSource _src;
        readonly long _keep, _keepItem;
        readonly FxPayload _payload;
        readonly long _lazyAbove = long.MaxValue;
        bool _stopped;
        public bool Truncated;
        public string Error;
        public long ErrorOffset;

        public FxLexer(FxSource src, long keepSingle, long keepPerMultiValue) { _src = src; _keep = keepSingle; _keepItem = keepPerMultiValue; }

        /// <summary>
        /// With a payload, a variable-size single value longer than <paramref name="lazyAbove"/> is read and forgotten: the property keeps
        /// an FxSegmentBlob (its place in the payload) instead of the bytes.
        /// </summary>
        public FxLexer(FxSource src, long keepSingle, long keepPerMultiValue, FxPayload payload, long lazyAbove)
            : this(src, keepSingle, keepPerMultiValue)
        {
            _payload = payload;
            if (payload != null) _lazyAbove = lazyAbove;
        }

        public bool Failed { get { return Truncated || Error != null; } }

        void Cut(FxProperty p)
        {
            if (p != null) { p.Truncated = true; p.Incomplete = true; }
            Truncated = true; _stopped = true;
        }

        void Fail(string text, long offset) { Error = text; ErrorOffset = offset; _stopped = true; }

        public FxToken Next()
        {
            if (_stopped) return null;
            long offset = _src.Position;
            byte[] tb = _src.ReadBytes(4);
            if (tb.Length == 0) { _stopped = true; return null; }
            if (tb.Length < 4) { Cut(null); return null; }
            uint tag = BitConverter.ToUInt32(tb, 0);
            if (FxNames.Markers.ContainsKey(tag)) return new FxToken { IsMarker = true, Tag = tag, Offset = offset };

            var p = new FxProperty { Offset = offset, Id = (ushort)(tag >> 16), Type = (ushort)tag };
            if (p.Type < 0x8000 && (p.Type & 0x4000) != 0)
            {
                // Seen on real streams, not in MS-OXCFXICS: the type carries the bit 0x4000 and two bytes (0x0200 so far)
                // sit between the tag and the value.
                byte[] flags = _src.ReadBytes(2);
                if (flags.Length < 2) { Cut(null); return null; }
                p.Extended = true;
                p.ExtendedFlags = BitConverter.ToUInt16(flags, 0);
                p.Type = (ushort)(p.Type & ~0x4000);
            }
            var token = new FxToken { Tag = p.Tag, Offset = offset, Property = p };
            if (p.IsNamed && !ReadNamedInfo(p, offset)) { token.Bad = true; return token; }
            ReadValue(p, offset);
            token.Bad = Failed;
            return token;
        }

        bool ReadNamedInfo(FxProperty p, long offset)
        {
            byte[] g = _src.ReadBytes(17);
            if (g.Length < 17) { Cut(p); return false; }
            p.NamedGuid = new Guid(g.Take(16).ToArray());
            if (g[16] == 0)
            {
                byte[] lid = _src.ReadBytes(4);
                if (lid.Length < 4) { Cut(p); return false; }
                p.NamedLid = BitConverter.ToUInt32(lid, 0);
                return true;
            }
            if (g[16] == 1)
            {
                var chars = new StringBuilder();
                while (chars.Length < MaxNamedChars)
                {
                    byte[] c = _src.ReadBytes(2);
                    if (c.Length < 2) { Cut(p); return false; }
                    if (c[0] == 0 && c[1] == 0) { p.NamedName = chars.ToString(); return true; }
                    chars.Append(BitConverter.ToChar(c, 0));
                }
                Fail("Name of a named property longer than " + MaxNamedChars + " characters (tag 0x" + p.Tag.ToString("X8") + ").", offset);
                return false;
            }
            Fail("Unknown kind " + g[16] + " of named property (tag 0x" + p.Tag.ToString("X8") + "): 0 (LID) or 1 (name) expected.", offset);
            return false;
        }

        void ReadValue(FxProperty p, long offset)
        {
            FxTypeInfo t = FxTypes.Describe(p.Type);
            bool maybeMulti = p.Type < 0x8000 && (p.Type & 0x1000) != 0;
            if (!t.Valid)
            {
                Fail((maybeMulti ? "Unknown multi-valued type 0x" : "Unknown property type 0x") + p.Type.ToString("X4") + " (tag 0x" + p.Tag.ToString("X8") + "): the stream cannot be followed after it.", offset);
                return;
            }

            if (!t.MultiValued)
            {
                if (t.FixedSize >= 0)
                {
                    p.ValueLength = t.FixedSize;
                    p.Raw = _src.ReadBytes(t.FixedSize);
                    if (p.Raw.Length < t.FixedSize) Cut(p);
                    return;
                }
                if (t.IsObject)
                {
                    // The grammar gives PtypObject a length and bytes; the first reader (and what a StartEmbed stream looks like to
                    // it) took it for a tag alone, so a StartEmbed marker right after the tag means "no value at all".
                    byte[] next = _src.Peek(4);
                    if (next.Length == 4 && BitConverter.ToUInt32(next, 0) == 0x40010003) { p.BareObject = true; p.Raw = new byte[0]; return; }
                }
                byte[] raw; long length;
                ReadVariable(p, _keep, true, out raw, out length);
                return;
            }

            byte[] cb = _src.ReadBytes(4);
            if (cb.Length < 4) { Cut(p); return; }
            uint count = BitConverter.ToUInt32(cb, 0);
            if (count > MaxMultiValues) { Fail("Multi-valued property with an absurd count (" + count + ") for tag 0x" + p.Tag.ToString("X8") + ".", offset); return; }
            p.Items = new List<byte[]>();
            for (uint i = 0; i < count; i++)
            {
                if (t.FixedSize >= 0)
                {
                    byte[] raw = _src.ReadBytes(t.FixedSize);
                    p.ValueLength += raw.Length;
                    if (raw.Length < t.FixedSize) { Cut(p); return; }
                    p.Items.Add(raw);
                }
                else
                {
                    byte[] raw; long length;
                    if (!ReadVariable(p, _keepItem, false, out raw, out length)) return;
                    p.Items.Add(raw);
                    if (length != raw.Length && p.ItemLengths == null)
                        p.ItemLengths = new List<long>(p.Items.Take(p.Items.Count - 1).Select(b => (long)b.Length));
                    if (p.ItemLengths != null) p.ItemLengths.Add(length);
                    p.ValueLength += length;
                }
            }
        }

        /// <summary>Length + bytes. Keeps the first <paramref name="keep"/> bytes, discards the rest. False if the stream ends inside the value.</summary>
        bool ReadVariable(FxProperty p, long keep, bool single, out byte[] kept, out long length)
        {
            kept = new byte[0];
            length = 0;
            byte[] lb = _src.ReadBytes(4);
            if (lb.Length < 4) { Cut(p); return false; }
            length = BitConverter.ToUInt32(lb, 0);
            if (single) p.ValueLength = length;
            if (single && _payload != null && length > _lazyAbove && !FxTypes.Describe(p.Type).IsObject)
            {
                long at = _src.Position;
                if (_src.Skip(length) < length) { p.Raw = kept; Cut(p); return false; }
                p.Lazy = new FxSegmentBlob(_payload, at, length);
                return true;
            }
            long want = Math.Min(length, keep);
            bool complete;
            kept = ReadKeep(want, out complete);
            if (single) p.Raw = kept;
            if (!complete) { Cut(p); return false; }
            long rest = length - want;
            if (rest > 0)
            {
                p.Truncated = true;     // longer than what is kept
                if (_src.Skip(rest) < rest) { Cut(p); return false; }
            }
            return true;
        }

        byte[] ReadKeep(long want, out bool complete)
        {
            if (want <= PieceSize)
            {
                byte[] b = _src.ReadBytes((int)want);
                complete = b.Length == want;
                return b;
            }
            // read in pieces: a length announced by garbage must not allocate gigabytes before the stream proves it has the bytes
            var pieces = new List<byte[]>();
            long got = 0;
            while (got < want)
            {
                byte[] piece = _src.ReadBytes((int)Math.Min(PieceSize, want - got));
                if (piece.Length > 0) pieces.Add(piece);
                got += piece.Length;
                if (piece.Length < PieceSize && got < want) break;
            }
            complete = got == want;
            var all = new byte[got];
            long at = 0;
            foreach (byte[] piece in pieces) { Buffer.BlockCopy(piece, 0, all, (int)at, piece.Length); at += piece.Length; }
            return all;
        }
    }

    public static class FxReader
    {
        static FxReader() { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); }

        public static string Hex(byte[] b, int max)
        {
            if (b == null) return "";
            int n = Math.Min(b.Length, max);
            var sb = new StringBuilder(n * 3);
            for (int i = 0; i < n; i++) { if (i > 0) sb.Append(' '); sb.Append(b[i].ToString("x2")); }
            if (b.Length > n) sb.Append(" ...");
            return sb.ToString();
        }

        static uint U32(byte[] b, int o) { return BitConverter.ToUInt32(b, o); }

        sealed class Frame
        {
            public string Name;
            public int Recipients, Attachments;
            public Frame(string name) { Name = name; }
        }

        /// <summary>Reads a whole file or stream. Never throws: problems are in the result (Error, Truncated).</summary>
        public static FxReadResult Read(Stream stream, FxReaderOptions options = null)
        {
            var o = options ?? new FxReaderOptions();
            var r = new FxReadResult();
            var src = new FxSource(stream);
            try
            {
                byte[] head = src.Peek(64);
                r.HeadHex = Hex(head, 64);
                if (head.Length >= 8 && U32(head, 0) == 1 && U32(head, 4) == 8) ReadEnvelope(src, r, o);
                else { r.DataOffset = 0; ParseElements(src, r, o); }
            }
            catch (Exception ex)
            {
                if (r.Error == null) { r.Error = ex.GetType().Name + ": " + ex.Message; r.ErrorOffset = src.Position; }
            }
            r.BytesRead = src.Position;
            return r;
        }

        public static FxReadResult ReadFile(string path, FxReaderOptions options = null)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 81920))
                return Read(fs, options);
        }

        /// <summary>
        /// The envelope is a sequence of records. The records of type 2 are consecutive pieces of ONE FastTransfer stream (a
        /// piece may end in the middle of a property: on a real export of 100 MiB, 3311 pieces of 31,680 bytes): the parser
        /// reads their concatenation, the other records are only noted.
        /// </summary>
        static void ReadEnvelope(FxSource src, FxReadResult r, FxReaderOptions o)
        {
            r.Envelope = "Tlv";
            var data = new FxEnvelopeStream(src, r);
            ParseElements(new FxSource(data), r, o);
            if (data.Truncated) r.Truncated = true;
            if (r.DataRecords == 0 && r.Error == null) r.Error = "Envelope without a record of type 2 (the FastTransfer data).";
        }

        static void ParseElements(FxSource src, FxReadResult r, FxReaderOptions o)
        {
            var frames = new Stack<Frame>();
            frames.Push(new Frame("message"));
            Func<string> path = () => string.Join("/", frames.Reverse().Select(f => f.Name));
            var lexer = new FxLexer(src, o.MaxPreviewBytes, Math.Min(o.MaxPreviewBytes, 1024));

            FxToken t;
            while ((t = lexer.Next()) != null)
            {
                if (t.IsMarker)
                {
                    string marker = FxNames.Markers[t.Tag];
                    Enter(marker, frames);
                    r.Elements.Add(new FxElement { IsMarker = true, Offset = t.Offset, Tag = t.Tag, Name = marker, Path = path(), TypeName = "Marker" });
                }
                else if (t.Tag == FxNames.FXDelProp)   // followed by the tag of the deleted property (0x0E12000D before the recipients, 0x0E13000D before the attachments)
                    r.Elements.Add(new FxElement { IsMarker = true, Offset = t.Offset, Tag = t.Tag, Name = "FXDelProp", Path = path(), TypeName = "Marker" });
                else
                    r.Elements.Add(ToElement(t.Property, path(), o));
            }
            if (lexer.Truncated) r.Truncated = true;
            if (lexer.Error != null) { r.Error = lexer.Error; r.ErrorOffset = lexer.ErrorOffset; }
        }

        static FxElement ToElement(FxProperty p, string path, FxReaderOptions o)
        {
            var el = new FxElement
            {
                Offset = p.Offset, Tag = p.Tag, Type = p.Type, Id = p.Id, Path = path,
                Extended = p.Extended, ExtendedFlags = p.ExtendedFlags,
                NamedGuid = p.NamedGuid, NamedLid = p.NamedLid, NamedString = p.NamedName,
                Truncated = p.Truncated, Incomplete = p.Incomplete, ValueBytes = p.ValueLength
            };
            el.TypeName = FxNames.TypeName(el.Type);
            el.Name = FxNames.Name(el.Id);

            FxTypeInfo t = p.TypeInfo;
            if (!t.Valid || t.IsObject) return el;   // an unknown type stops the lexer; PtypObject: a StartEmbed marker follows
            int cp = t.CodePageString ? t.CodePage : 1252;

            if (!t.MultiValued)
            {
                if (t.FixedSize >= 0)
                {
                    if (!p.Incomplete) el.Value = FxValue.DecodeFixed(p.Raw, t.BaseType);
                    return el;
                }
                el.Preview = p.Raw;
                if (!p.Incomplete) el.Value = t.BaseType == 0x0102 || t.BaseType == 0x00FB ? null : FxValue.DecodeString(p.Raw, t.BaseType, cp);
                return el;
            }

            int shown = p.Items == null ? 0 : Math.Min(p.Items.Count, o.MaxMultiValues);
            var values = new object[shown];
            for (int i = 0; i < shown; i++)
            {
                byte[] raw = p.Items[i];
                if (t.FixedSize >= 0) values[i] = FxValue.DecodeFixed(raw, t.BaseType);
                else if (t.BaseType == 0x0102 || t.BaseType == 0x00FB) values[i] = (p.ItemLengths != null ? p.ItemLengths[i] : raw.Length) + " bytes";
                else values[i] = FxValue.DecodeString(raw, t.BaseType, cp);
            }
            el.Value = values;
            return el;
        }

        static void Enter(string marker, Stack<Frame> frames)
        {
            Frame top = frames.Peek();
            switch (marker)
            {
                case "StartRecip": top.Recipients++; frames.Push(new Frame("recipient[" + (top.Recipients - 1) + "]")); break;
                case "NewAttach": top.Attachments++; frames.Push(new Frame("attachment[" + (top.Attachments - 1) + "]")); break;
                case "StartEmbed": frames.Push(new Frame("embedded")); break;
                case "EndToRecip": PopIf(frames, "recipient"); break;
                case "EndAttach": PopIf(frames, "attachment"); break;
                case "EndEmbed": PopIf(frames, "embedded"); break;
            }
        }

        static void PopIf(Stack<Frame> frames, string prefix)
        {
            if (frames.Count > 1 && frames.Peek().Name.StartsWith(prefix, StringComparison.Ordinal)) frames.Pop();
        }
    }
}
