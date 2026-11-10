// =============================================================================
//  Mailbox Transfer - engine, part 4b: FastTransfer object model, encoder and round-trip validator (MS-OXCFXICS)
// -----------------------------------------------------------------------------
//  FxReader (Engine.FastTransfer.cs) LOOKS at a stream (flat list of elements, previews, never keeps an attachment).
//  This file is the exact side, the foundation of a shared MAPI object model:
//
//    FxCodec.Decode      stream (with or without the Graph envelope)  ->  FxDocument { Envelope, Message, Warnings }
//    FxCodec.Encode      FxDocument / FxMessage                       ->  stream
//    FxCodec.RoundTrip   decode, encode again, compare with the original byte for byte (first differing offset)
//
//  Model (what MS-OXCFXICS 2.2.4.2 calls messageContent):
//    FxMessage    = properties, [FXDelProp for the recipients], recipients[], [FXDelProp for the attachments], attachments[]
//    FxRecipient  = StartRecip properties EndToRecip
//    FxAttachment = NewAttach properties [StartEmbed FxMessage EndEmbed] EndAttach
//    FxProperty   = tag (id + type), optional named identity (GUID + LID | name), raw value bytes + typed accessors
//  A value is kept as the bytes found in the stream (without the length prefix); the accessors decode them. That is what makes the
//  round trip exact: nothing is normalized (a string keeps its terminator, a Boolean its 2 bytes, an unused byte stays).
//  Decode(IFxStreamSource) does not keep a single variable-size value above 64 KiB: FxProperty.Lazy remembers where it is in the
//  source (Engine.FxStream.cs) and the encoder, the PST writer and the comparison read it again in pieces, so that the memory
//  does not depend on the size of an item (a value is less than 4 GiB: its length is 32 bits in the stream).
//
//  What is kept so that the bytes come back identical:
//    - the order of the properties (the body may follow the named properties, the stream is not sorted by tag);
//    - the bit 0x4000 of a type and the two bytes that follow the tag (Extended, ExtendedFlags): seen on real streams, not in the spec;
//    - a PtypObject written without a length right before StartEmbed (BareObject): the grammar says length + bytes, the first
//      reader and its tests assumed no bytes at all, so both are accepted and each one is written back as it was read;
//    - the envelope: the records in their order, their payloads, the length of each data record (the service cuts the stream every
//      31,680 bytes, a value may straddle two records);
//    - the wrapper StartMessage / StartFAIMsg ... EndMessage when the stream has one (the Graph export has none).
//  Where identity is impossible by design, the decoder says so in Warnings (a property after the recipients, a marker the grammar
//  does not allow there, bytes after the end of the message): such a stream is still decoded, but encoded in the canonical order
//  of the grammar, so the round trip reports a difference (and the warning says why).
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    /// <summary>What the 16 bits of a property type mean in a FastTransfer stream.</summary>
    public sealed class FxTypeInfo
    {
        /// <summary>The type, without the extended bit 0x4000.</summary>
        public ushort Type;
        public bool Valid;
        public bool MultiValued;
        /// <summary>A code page string type (0x8000 | code page, 0x9000 | code page for the multi-valued form).</summary>
        public bool CodePageString;
        /// <summary>Base type of one value: 0x0003 for 0x1003; 0x001E for every code page string (the code page is in CodePage).</summary>
        public int BaseType;
        public int CodePage;
        /// <summary>Bytes of one value of a fixed-size type (PtypBoolean is 2 in a FastTransfer stream), -1 for a variable-size one.</summary>
        public int FixedSize = -1;
        public bool IsObject;
        public bool IsString { get { return BaseType == 0x001E || BaseType == 0x001F; } }
    }

    public static class FxTypes
    {
        static readonly Dictionary<ushort, FxTypeInfo> Cache = new Dictionary<ushort, FxTypeInfo>();

        static FxTypes() { Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); }

        static bool IsCodePage(int cp)
        {
            if (cp <= 0) return false;
            try { Encoding.GetEncoding(cp); return true; } catch (Exception) { return false; }
        }

        /// <summary>The fixed size of a base type, -1 if it is not a fixed-size one.</summary>
        public static int FixedSize(int baseType)
        {
            switch (baseType)
            {
                case 0x0001: return 0;                      // PtypNull
                case 0x0002: return 2;                      // PtypInteger16
                case 0x0003: case 0x0004: case 0x000A: return 4;   // Integer32, Floating32, ErrorCode
                case 0x0005: case 0x0006: case 0x0007: case 0x0014: case 0x0040: return 8;   // Floating64, Currency, FloatingTime, Integer64, Time
                case 0x000B: return 2;                      // PtypBoolean: 2 bytes in a FastTransfer stream (1 in MS-OXCDATA)
                case 0x0048: return 16;                     // PtypGuid
            }
            return -1;
        }

        /// <summary>Length-prefixed base types: String8, String, Binary, ServerId and Object.</summary>
        static bool IsVariable(int baseType)
        {
            return baseType == 0x001E || baseType == 0x001F || baseType == 0x0102 || baseType == 0x00FB || baseType == 0x000D;
        }

        /// <summary>
        /// Classifies a property type (the bit 0x4000 of the extended form already removed). A type from 0x8000 is a code page string:
        /// 0x84B0 is code page 1200 (UTF-16LE), 0xFDE9 is 65001 (UTF-8, it does not fit in 15 bits); the bit 0x1000 on top of
        /// a code page makes it multi-valued (0x94B0, seen on a real stream: PtypMultipleString in code page 1200). A code page that
        /// itself has the bit 0x1000 (KOI8-R 20866 = 0x5182 gives 0xD182) stays single when the type is a code page as it stands.
        /// </summary>
        public static FxTypeInfo Describe(ushort type)
        {
            FxTypeInfo info;
            lock (Cache) { if (Cache.TryGetValue(type, out info)) return info; }
            info = Classify(type);
            lock (Cache) { Cache[type] = info; }
            return info;
        }

        static FxTypeInfo Classify(ushort type)
        {
            var i = new FxTypeInfo { Type = type };
            if (type >= 0x8000)
            {
                int cp = type == 0xFDE9 ? 65001 : type & 0x7FFF;
                bool multi = false;
                if ((type & 0x1000) != 0 && type != 0xFDE9 && !IsCodePage(cp) && IsCodePage(cp & ~0x1000)) { multi = true; cp &= ~0x1000; }
                i.Valid = true; i.CodePageString = true; i.CodePage = cp; i.BaseType = 0x001E; i.MultiValued = multi;
                return i;
            }
            i.MultiValued = (type & 0x1000) != 0;
            i.BaseType = type & 0x0FFF;
            if ((type & 0x6000) != 0) return i;    // 0x2000 (MVI) and 0x4000 are not part of a FastTransfer type
            i.FixedSize = FixedSize(i.BaseType);
            if (!i.MultiValued)
            {
                i.IsObject = i.BaseType == 0x000D;
                i.Valid = i.FixedSize >= 0 || IsVariable(i.BaseType);
            }
            else i.Valid = i.FixedSize > 0 || (i.FixedSize < 0 && IsVariable(i.BaseType) && i.BaseType != 0x000D);
            return i;
        }
    }

    /// <summary>Decoding of the values (shared by the model and by the reader).</summary>
    public static class FxValue
    {
        static readonly DateTime FileTimeEpoch = new DateTime(1601, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        /// <summary>Last FILETIME that DateTimeOffset can hold.</summary>
        const long MaxFileTime = 2650467743990000000L;

        public static object DecodeFixed(byte[] b, int baseType)
        {
            switch (baseType)
            {
                case 0x0001: return null;
                case 0x0002: return BitConverter.ToInt16(b, 0);
                case 0x0003: return BitConverter.ToInt32(b, 0);
                case 0x0004: return BitConverter.ToSingle(b, 0);
                case 0x0005: case 0x0007: return BitConverter.ToDouble(b, 0);
                case 0x0006: return BitConverter.ToInt64(b, 0) / 10000m;
                case 0x000A: return BitConverter.ToUInt32(b, 0);
                case 0x000B: return b[0] != 0 || b[1] != 0;
                case 0x0014: return BitConverter.ToInt64(b, 0);
                case 0x0040:
                    long ticks = BitConverter.ToInt64(b, 0);
                    if (ticks <= 0 || ticks > MaxFileTime) return ticks;   // not a date: keep the raw value
                    return new DateTimeOffset(FileTimeEpoch.AddTicks(ticks), TimeSpan.Zero);
                case 0x0048: return new Guid(b);
            }
            return null;
        }

        /// <summary>A string value: UTF-16LE for PtypString and code page 1200, else the code page (PtypString8: 1252). Trailing nulls removed.</summary>
        public static string DecodeString(byte[] bytes, int baseType, int codePage)
        {
            if (bytes == null) return null;
            try
            {
                if (baseType == 0x001F || codePage == 1200) return Encoding.Unicode.GetString(bytes, 0, bytes.Length & ~1).TrimEnd('\0');
                if (codePage == 1201) return Encoding.BigEndianUnicode.GetString(bytes, 0, bytes.Length & ~1).TrimEnd('\0');
                Encoding enc;
                try { enc = Encoding.GetEncoding(codePage); } catch (Exception) { enc = Encoding.Latin1; }
                return enc.GetString(bytes).TrimEnd('\0');
            }
            catch (Exception) { return null; }
        }

        public static DateTimeOffset? ToTime(byte[] b)
        {
            long ticks = BitConverter.ToInt64(b, 0);
            if (ticks <= 0 || ticks > MaxFileTime) return null;
            return new DateTimeOffset(FileTimeEpoch.AddTicks(ticks), TimeSpan.Zero);
        }
    }

    /// <summary>One property of the stream: tag, optional named identity and the value as it is encoded.</summary>
    public sealed class FxProperty
    {
        public ushort Id;
        /// <summary>The type without the extended bit 0x4000 (that bit is in Extended).</summary>
        public ushort Type;
        /// <summary>The type carried the bit 0x4000 and two bytes sat between the tag and the value (real streams, not in the spec).</summary>
        public bool Extended;
        public ushort ExtendedFlags;
        /// <summary>Named property (Id >= 0x8000): the property set, and a numeric id (LID) or a name.</summary>
        public Guid? NamedGuid;
        public uint? NamedLid;
        public string NamedName;
        byte[] _raw;
        FxBlob _lazy;
        /// <summary>
        /// Single-valued: the value without its length prefix (a fixed-size value: its bytes; a string: with its terminator). For a
        /// lazy property (see Lazy) it reads the whole value into a new array each time (not above 2 GiB): code that must stay
        /// bounded uses Blob / OpenValue.
        /// </summary>
        public byte[] Raw
        {
            get { return _raw ?? (_lazy != null ? _lazy.ToArray() : null); }
            set { _raw = value; if (value != null) _lazy = null; }
        }
        /// <summary>
        /// Single-valued, variable size: the value is not in memory but in its source (a segment of the stream that was decoded,
        /// a data tree of a PST, ...). Set by FxCodec.Decode(IFxStreamSource) for values above FxDecodeOptions.LazyAbove.
        /// </summary>
        public FxBlob Lazy
        {
            get { return _lazy; }
            set { _lazy = value; if (value != null) _raw = null; }
        }
        public bool IsLazy { get { return _raw == null && _lazy != null; } }
        /// <summary>Length of a single value, whether it is in memory or lazy.</summary>
        public long RawLength { get { return _raw != null ? _raw.Length : _lazy != null ? _lazy.Length : 0; } }
        /// <summary>The single value as a blob (lazy or in memory), null if there is none.</summary>
        public FxBlob Blob { get { return _lazy != null && _raw == null ? _lazy : _raw != null ? new FxBytesBlob(_raw) : null; } }
        /// <summary>A stream of the single value (its bytes as encoded, without length prefix).</summary>
        public Stream OpenValue() { FxBlob b = Blob; return b == null ? new MemoryStream(new byte[0], false) : b.OpenRead(); }
        /// <summary>Multi-valued: one entry per value, same convention as Raw.</summary>
        public List<byte[]> Items;
        /// <summary>A PtypObject property written without length, just before StartEmbed.</summary>
        public bool BareObject;

        // Set by the decoder, not needed to encode.
        public long Offset;
        /// <summary>Size of the value in the stream (single: the value; multi-valued: all values), whatever was kept.</summary>
        public long ValueLength;
        /// <summary>The value is longer than what was kept (preview mode) or the stream ends inside it.</summary>
        public bool Truncated;
        public bool Incomplete;
        /// <summary>Multi-valued, preview mode only: the true length of each value, set when some value was cut (Items keeps the first bytes).</summary>
        public List<long> ItemLengths;

        public uint Tag { get { return ((uint)Id << 16) | Type; } }
        public bool IsNamed { get { return Id >= 0x8000; } }
        public FxTypeInfo TypeInfo { get { return FxTypes.Describe(Type); } }
        public bool IsMultiValued { get { return TypeInfo.MultiValued; } }

        public string NamedText
        {
            get
            {
                if (!IsNamed || NamedGuid == null) return null;
                return NamedGuid + (NamedLid != null ? " lid 0x" + NamedLid.Value.ToString("X") : NamedName != null ? " name " + NamedName : "");
            }
        }

        public override string ToString()
        {
            return "0x" + Tag.ToString("X8") + (NamedText != null ? " " + NamedText : "") + " " + FxNames.TypeName(Type);
        }

        // ---- decoded value ------------------------------------------------------------------------------------------------

        /// <summary>
        /// The value as a .NET object: number, bool, DateTimeOffset (a long if the FILETIME is out of range), Guid, string, byte[] (binary,
        /// server id, object); object[] for a multi-valued property; null for PtypNull.
        /// </summary>
        public object Value
        {
            get
            {
                FxTypeInfo t = TypeInfo;
                if (!t.Valid) return Raw;
                if (!t.MultiValued) return DecodeOne(t, Raw);
                var list = new object[Items == null ? 0 : Items.Count];
                for (int i = 0; i < list.Length; i++) list[i] = DecodeOne(t, Items[i]);
                return list;
            }
        }

        static object DecodeOne(FxTypeInfo t, byte[] raw)
        {
            if (raw == null) return null;
            if (t.FixedSize >= 0) return raw.Length < t.FixedSize ? null : FxValue.DecodeFixed(raw, t.BaseType);
            if (t.IsString) return FxValue.DecodeString(raw, t.BaseType, t.CodePageString ? t.CodePage : 1252);
            return raw;
        }

        InvalidOperationException Wrong(string wanted) { return new InvalidOperationException("Property " + this + " is not " + wanted + "."); }

        void NeedSingleFixed(int size, string wanted)
        {
            if (TypeInfo.MultiValued || Raw == null || Raw.Length < size) throw Wrong(wanted);
        }

        public short AsInt16() { NeedSingleFixed(2, "an integer of 16 bits"); return BitConverter.ToInt16(Raw, 0); }
        public int AsInt32() { NeedSingleFixed(4, "an integer of 32 bits"); return BitConverter.ToInt32(Raw, 0); }
        public long AsInt64() { NeedSingleFixed(8, "an integer of 64 bits"); return BitConverter.ToInt64(Raw, 0); }
        public double AsDouble() { NeedSingleFixed(8, "a floating point number of 64 bits"); return BitConverter.ToDouble(Raw, 0); }
        public bool AsBool() { NeedSingleFixed(2, "a Boolean"); return Raw[0] != 0 || Raw[1] != 0; }
        public Guid AsGuid() { NeedSingleFixed(16, "a GUID"); return new Guid(Raw); }
        public DateTimeOffset? AsTime() { NeedSingleFixed(8, "a time"); return FxValue.ToTime(Raw); }
        public byte[] AsBytes() { if (TypeInfo.MultiValued) throw Wrong("single-valued"); return Raw; }

        public string AsString()
        {
            FxTypeInfo t = TypeInfo;
            if (t.MultiValued || !t.IsString) throw Wrong("a string");
            return FxValue.DecodeString(Raw, t.BaseType, t.CodePageString ? t.CodePage : 1252);
        }

        public string[] AsStrings()
        {
            FxTypeInfo t = TypeInfo;
            if (!t.MultiValued || !t.IsString) throw Wrong("a multi-valued string");
            return Items.Select(b => FxValue.DecodeString(b, t.BaseType, t.CodePageString ? t.CodePage : 1252)).ToArray();
        }

        // ---- creation (for the writers built on this model) -----------------------------------------------------------------

        /// <summary>A single-valued property from its encoded value (fixed-size: exactly the bytes of the type).</summary>
        public static FxProperty Create(ushort id, ushort type, byte[] raw) { return new FxProperty { Id = id, Type = type, Raw = raw, ValueLength = raw == null ? 0 : raw.Length }; }

        /// <summary>A single-valued variable-size property whose value stays in its source (written by copying from it).</summary>
        public static FxProperty CreateLazy(ushort id, ushort type, FxBlob value) { return new FxProperty { Id = id, Type = type, Lazy = value, ValueLength = value.Length }; }

        public static FxProperty CreateMulti(ushort id, ushort type, IEnumerable<byte[]> items)
        {
            var p = new FxProperty { Id = id, Type = type, Items = new List<byte[]>(items) };
            p.ValueLength = p.Items.Sum(b => (long)b.Length);
            return p;
        }

        public static FxProperty FromInt16(ushort id, short v) { return Create(id, 0x0002, BitConverter.GetBytes(v)); }
        public static FxProperty FromInt32(ushort id, int v) { return Create(id, 0x0003, BitConverter.GetBytes(v)); }
        public static FxProperty FromInt64(ushort id, long v) { return Create(id, 0x0014, BitConverter.GetBytes(v)); }
        public static FxProperty FromDouble(ushort id, double v) { return Create(id, 0x0005, BitConverter.GetBytes(v)); }
        public static FxProperty FromBool(ushort id, bool v) { return Create(id, 0x000B, new byte[] { (byte)(v ? 1 : 0), 0 }); }
        public static FxProperty FromGuid(ushort id, Guid v) { return Create(id, 0x0048, v.ToByteArray()); }
        public static FxProperty FromTime(ushort id, DateTime utc) { return Create(id, 0x0040, BitConverter.GetBytes(utc.ToFileTimeUtc())); }
        public static FxProperty FromBinary(ushort id, byte[] v) { return Create(id, 0x0102, v); }
        /// <summary>PtypString: UTF-16LE with its terminator.</summary>
        public static FxProperty FromString(ushort id, string v) { return Create(id, 0x001F, Encoding.Unicode.GetBytes(v + "\0")); }

        public static FxProperty FromCodePageString(ushort id, int codePage, string v)
        {
            ushort type = codePage == 65001 ? (ushort)0xFDE9 : (ushort)(0x8000 | codePage);
            return Create(id, type, Encoding.GetEncoding(codePage).GetBytes(v + "\0"));
        }

        public static FxProperty FromStrings(ushort id, IEnumerable<string> v) { return CreateMulti(id, 0x101F, v.Select(s => Encoding.Unicode.GetBytes(s + "\0"))); }

        public FxProperty Named(Guid set, uint lid) { NamedGuid = set; NamedLid = lid; NamedName = null; return this; }
        public FxProperty Named(Guid set, string name) { NamedGuid = set; NamedName = name; NamedLid = null; return this; }
    }

    public sealed class FxRecipient
    {
        public List<FxProperty> Properties = new List<FxProperty>();
    }

    public sealed class FxAttachment
    {
        public List<FxProperty> Properties = new List<FxProperty>();
        /// <summary>The embedded message (StartEmbed ... EndEmbed), null for a file attachment.</summary>
        public FxMessage Embedded;
    }

    /// <summary>messageContent of MS-OXCFXICS 2.2.4.2. Summarize (the readable summary of a flat read) is in Engine.Message.cs.</summary>
    public sealed partial class FxMessage
    {
        public const uint MarkerStartMessage = 0x400C0003, MarkerEndMessage = 0x400D0003, MarkerStartFAIMsg = 0x40100003;
        public const uint TagRecipients = 0x0E12000D, TagAttachments = 0x0E13000D;

        /// <summary>0, or StartMessage / StartFAIMsg when the stream wraps the content in StartMessage ... EndMessage. An embedded message never has one.</summary>
        public uint StartMarker;
        public List<FxProperty> Properties = new List<FxProperty>();
        /// <summary>FXDelProp followed by PidTagMessageRecipients, written before the recipients (null: absent).</summary>
        public FxProperty DelPropRecipients;
        public List<FxRecipient> Recipients = new List<FxRecipient>();
        /// <summary>FXDelProp followed by PidTagMessageAttachments, written before the attachments (null: absent).</summary>
        public FxProperty DelPropAttachments;
        public List<FxAttachment> Attachments = new List<FxAttachment>();

        public FxProperty Find(ushort id)
        {
            return Properties.FirstOrDefault(p => p.Id == id && !p.IsNamed);
        }

        public FxProperty Find(Guid set, uint lid)
        {
            return Properties.FirstOrDefault(p => p.IsNamed && p.NamedGuid == set && p.NamedLid == lid);
        }

        public FxProperty Find(Guid set, string name)
        {
            return Properties.FirstOrDefault(p => p.IsNamed && p.NamedGuid == set && string.Equals(p.NamedName, name, StringComparison.Ordinal));
        }

        /// <summary>The string value of a property of the message (any single-valued string type), null if absent.</summary>
        public string GetString(ushort id)
        {
            FxProperty p = Find(id);
            return p != null && p.TypeInfo.IsString && !p.IsMultiValued ? p.AsString() : null;
        }

        /// <summary>Properties of the message, its recipients, its attachments and, recursively, its embedded messages.</summary>
        public int CountProperties()
        {
            int n = Properties.Count + (DelPropRecipients != null ? 1 : 0) + (DelPropAttachments != null ? 1 : 0);
            foreach (FxRecipient r in Recipients) n += r.Properties.Count;
            foreach (FxAttachment a in Attachments) { n += a.Properties.Count; if (a.Embedded != null) n += a.Embedded.CountProperties(); }
            return n;
        }

        /// <summary>0 for a message without embedded message, 1 if an attachment holds one, 2 if that one holds one too, ...</summary>
        public int EmbeddedDepth()
        {
            int d = 0;
            foreach (FxAttachment a in Attachments) if (a.Embedded != null) d = Math.Max(d, 1 + a.Embedded.EmbeddedDepth());
            return d;
        }
    }

    /// <summary>One record of the Graph envelope (type, length, payload), in file order.</summary>
    public sealed class FxEnvelopeEntry
    {
        public uint Type;
        public long Length;
        /// <summary>Type 2: a piece of the FastTransfer stream. Its bytes are not kept here: the stream is encoded again and cut again.</summary>
        public bool IsData { get { return Type == 2; } }
        /// <summary>The bytes of a record that is not data (the first 1 MiB at most: a larger payload would make the round trip differ).</summary>
        public byte[] Payload;
        /// <summary>Offset of the payload of the record in the file that was decoded (set by the decoder).</summary>
        public long RawOffset;
    }

    public sealed class FxEnvelope
    {
        public List<FxEnvelopeEntry> Entries = new List<FxEnvelopeEntry>();
        public int DataRecords { get { return Entries.Count(e => e.IsData); } }
        public long DataLength { get { return Entries.Where(e => e.IsData).Sum(e => e.Length); } }
    }

    public sealed class FxDocument
    {
        /// <summary>The records in front of the stream (Graph export), null when the file is the bare FastTransfer stream.</summary>
        public FxEnvelope Envelope;
        public FxMessage Message = new FxMessage();
        /// <summary>Everything the decoder saw that the syntax of messageContent does not allow (see the header of this file).</summary>
        public List<string> Warnings = new List<string>();
        public bool Truncated;
        public string Error;
        public long ErrorOffset;
        /// <summary>Bytes of the FastTransfer stream itself (the envelope excluded).</summary>
        public long PayloadBytes;
        public bool Ok { get { return Error == null && !Truncated && Warnings.Count == 0; } }
    }

    public sealed class FxDecodeOptions
    {
        /// <summary>Largest value kept whole (default: no limit, the round trip needs every byte).</summary>
        public long MaxValueBytes = long.MaxValue;
        /// <summary>
        /// Decode(IFxStreamSource) only: a variable-size single value longer than this is not kept but read again from the source
        /// when it is used (FxProperty.Lazy). The decoders that take a byte[] or a Stream keep everything.
        /// </summary>
        public long LazyAbove = 64 * 1024;
        /// <summary>Cancels a long decode (a big value is read through to follow the stream): OperationCanceledException, the only exception Decode throws.</summary>
        public System.Threading.CancellationToken Cancellation;
    }

    public sealed class FxRoundTripResult
    {
        public FxDocument Document;
        public bool Identical;
        /// <summary>Offset of the first byte that differs (or the length of the shorter one if one is a prefix of the other); -1 if identical.</summary>
        public long FirstDifference = -1;
        public long OriginalBytes;
        public long EncodedBytes;
        /// <summary>Empty when identical; else what is known about the difference (the error, the first warning, ...).</summary>
        public string Note = "";
    }

    public static class FxCodec
    {
        const uint MarkerStartRecip = 0x40030003, MarkerEndToRecip = 0x40040003, MarkerNewAttach = 0x40000003, MarkerEndAttach = 0x400E0003;
        const uint MarkerStartEmbed = 0x40010003, MarkerEndEmbed = 0x40020003;
        public const int DefaultChunk = 31680;
        const int MaxDepth = 64;

        // ---- decode ------------------------------------------------------------------------------------------------------

        public static FxDocument Decode(byte[] bytes, FxDecodeOptions options = null)
        {
            using (var ms = new MemoryStream(bytes, false)) return Decode(ms, options);
        }

        public static FxDocument DecodeFile(string path, FxDecodeOptions options = null)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 1 << 20, FileOptions.SequentialScan))
                return Decode(fs, options);
        }

        /// <summary>
        /// Same, from a source that can be opened again: a variable-size single value longer than options.LazyAbove (64 KiB by default)
        /// is not kept, the property holds an FxBlob (FxProperty.Lazy) that reads it again from the source when it is used. The
        /// memory of the document does not depend on the size of the values, but the source must stay readable (and unchanged) as
        /// long as the document is used.
        /// </summary>
        public static FxDocument Decode(IFxStreamSource source, FxDecodeOptions options = null)
        {
            if (source == null) throw new ArgumentNullException("source");
            using (Stream s = source.OpenRead()) return DecodeCore(s, source, options);
        }

        /// <summary>Never throws: problems are in Error, Truncated and Warnings. Keeps every value in memory.</summary>
        public static FxDocument Decode(Stream stream, FxDecodeOptions options = null)
        {
            return DecodeCore(stream, null, options);
        }

        static FxDocument DecodeCore(Stream stream, IFxStreamSource lazySource, FxDecodeOptions options)
        {
            var o = options ?? new FxDecodeOptions();
            var doc = new FxDocument();
            var src = new FxSource(o.Cancellation.CanBeCanceled ? new FxCancellableStream(stream, o.Cancellation) : stream);
            FxLexer lexer = null;
            FxEnvelopeStream envelope = null;
            try
            {
                byte[] head = src.Peek(8);
                FxSource data = src;
                FxPayload payload = null;
                if (head.Length >= 8 && BitConverter.ToUInt32(head, 0) == 1 && BitConverter.ToUInt32(head, 4) == 8)
                {
                    envelope = new FxEnvelopeStream(src, new FxReadResult());
                    doc.Envelope = new FxEnvelope { Entries = envelope.Entries };
                    data = new FxSource(envelope);
                    if (lazySource != null) payload = new FxPayload(lazySource, envelope.Entries);
                }
                else if (lazySource != null) payload = new FxPayload(lazySource, null);
                lexer = new FxLexer(data, o.MaxValueBytes, o.MaxValueBytes, payload, o.LazyAbove);
                new ModelBuilder(lexer, doc).Build();
                doc.PayloadBytes = data.Position;
            }
            catch (OperationCanceledException) { throw; }
            catch (Exception ex)
            {
                if (doc.Error == null) doc.Error = ex.GetType().Name + ": " + ex.Message;
            }
            if (lexer != null)
            {
                if (lexer.Error != null && doc.Error == null) { doc.Error = lexer.Error; doc.ErrorOffset = lexer.ErrorOffset; }
                if (lexer.Truncated) doc.Truncated = true;
            }
            if (envelope != null && envelope.Truncated) doc.Truncated = true;
            if (envelope != null && !envelope.Entries.Any(e => e.IsData) && doc.Error == null) doc.Error = "Envelope without a record of type 2 (the FastTransfer data).";
            return doc;
        }

        /// <summary>Builds the message from the tokens of the lexer, following the syntax of 2.2.4.2 (see the header for what it tolerates).</summary>
        sealed class ModelBuilder
        {
            readonly FxLexer _lex;
            readonly FxDocument _doc;
            FxToken _peek;
            bool _hasPeek;

            public ModelBuilder(FxLexer lex, FxDocument doc) { _lex = lex; _doc = doc; }

            FxToken Peek() { if (!_hasPeek) { _peek = _lex.Next(); _hasPeek = true; } return _peek; }
            FxToken Next() { FxToken t = Peek(); _hasPeek = false; _peek = null; return t; }

            void Warn(FxToken t, string text) { _doc.Warnings.Add((t == null ? "at the end of the stream" : "at offset " + t.Offset) + ": " + text); }

            static string Name(FxToken t)
            {
                string n;
                return t.IsMarker && FxNames.Markers.TryGetValue(t.Tag, out n) ? n : "0x" + t.Tag.ToString("X8");
            }

            /// <summary>A property cut short by the end of the stream, or of an unknown type, is not added to the model (the document says why).</summary>
            static bool Whole(FxToken t) { return !t.Bad; }

            public void Build()
            {
                FxMessage m = _doc.Message;
                FxToken first = Peek();
                if (first != null && first.IsMarker && (first.Tag == FxMessage.MarkerStartMessage || first.Tag == FxMessage.MarkerStartFAIMsg))
                {
                    Next();
                    m.StartMarker = first.Tag;
                    Content(m, 0);
                    FxToken end = Peek();
                    if (end != null && end.IsMarker && end.Tag == FxMessage.MarkerEndMessage) Next();
                    else if (end != null || !_lex.Failed) Warn(end, "EndMessage expected");
                }
                else Content(m, 0);
                FxToken rest = Peek();
                if (rest != null)
                {
                    Warn(rest, "data after the end of the message: " + Name(rest));
                    while (Next() != null) { }
                }
            }

            static bool IsEnd(uint tag)
            {
                return tag == FxMessage.MarkerEndMessage || tag == MarkerEndEmbed || tag == MarkerEndAttach || tag == MarkerEndToRecip;
            }

            /// <summary>messageContent: propList messageChildren. Stops (without consuming) at an end marker or at the end of the stream.</summary>
            void Content(FxMessage m, int depth)
            {
                int phase = 0;   // 0 properties, 1 recipients, 2 attachments
                while (true)
                {
                    FxToken t = Peek();
                    if (t == null) return;
                    if (!Whole(t)) { Next(); return; }
                    if (t.IsMarker)
                    {
                        if (IsEnd(t.Tag)) return;
                        if (t.Tag == MarkerStartRecip)
                        {
                            Next();
                            if (phase > 1) Warn(t, "recipient after the attachments (moved before them when encoded)");
                            var r = new FxRecipient();
                            Props(r.Properties, MarkerEndToRecip, "recipient");
                            m.Recipients.Add(r);
                            if (phase < 1) phase = 1;
                        }
                        else if (t.Tag == MarkerNewAttach)
                        {
                            Next();
                            m.Attachments.Add(Attachment(depth));
                            phase = 2;
                        }
                        else { Warn(t, "marker " + Name(t) + " is not allowed in the content of a message (dropped)"); Next(); }
                        continue;
                    }

                    FxProperty p = t.Property;
                    if (p.Tag == FxNames.FXDelProp && p.Raw != null && p.Raw.Length == 4)
                    {
                        uint what = BitConverter.ToUInt32(p.Raw, 0);
                        if (what == FxMessage.TagRecipients && phase == 0 && m.DelPropRecipients == null) { Next(); m.DelPropRecipients = p; phase = 1; continue; }
                        if (what == FxMessage.TagAttachments && phase <= 1 && m.DelPropAttachments == null) { Next(); m.DelPropAttachments = p; phase = 2; continue; }
                    }
                    Next();
                    if (phase > 0) Warn(t, "property " + p + " after the recipients or the attachments (moved before them when encoded)");
                    m.Properties.Add(p);
                }
            }

            void Props(List<FxProperty> into, uint endMarker, string what)
            {
                while (true)
                {
                    FxToken t = Peek();
                    if (t == null) { if (!_lex.Failed) Warn(t, what + " not closed"); return; }
                    if (!Whole(t)) { Next(); return; }
                    if (t.IsMarker)
                    {
                        if (t.Tag == endMarker) { Next(); return; }
                        if (IsEnd(t.Tag) || t.Tag == MarkerStartRecip || t.Tag == MarkerNewAttach) { Warn(t, what + " not closed before " + Name(t)); return; }
                        Warn(t, "marker " + Name(t) + " is not allowed in a " + what + " (dropped)");
                        Next();
                        continue;
                    }
                    Next();
                    into.Add(t.Property);
                }
            }

            FxAttachment Attachment(int depth)
            {
                var a = new FxAttachment();
                while (true)
                {
                    FxToken t = Peek();
                    if (t == null) { if (!_lex.Failed) Warn(t, "attachment not closed"); return a; }
                    if (!Whole(t)) { Next(); return a; }
                    if (t.IsMarker)
                    {
                        if (t.Tag == MarkerEndAttach) { Next(); return a; }
                        if (t.Tag == MarkerStartEmbed)
                        {
                            Next();
                            if (a.Embedded != null) Warn(t, "second embedded message in an attachment");
                            if (depth + 1 > MaxDepth) { Warn(t, "embedded messages nested more than " + MaxDepth + " levels"); a.Embedded = new FxMessage(); }
                            else
                            {
                                var e = new FxMessage();
                                Content(e, depth + 1);
                                a.Embedded = e;
                            }
                            FxToken end = Peek();
                            if (end != null && end.IsMarker && end.Tag == MarkerEndEmbed) Next();
                            else if (end != null || !_lex.Failed) Warn(end, "EndEmbed expected");
                            continue;
                        }
                        if (IsEnd(t.Tag) || t.Tag == MarkerStartRecip || t.Tag == MarkerNewAttach) { Warn(t, "attachment not closed before " + Name(t)); return a; }
                        Warn(t, "marker " + Name(t) + " is not allowed in an attachment (dropped)");
                        Next();
                        continue;
                    }
                    Next();
                    if (a.Embedded != null) Warn(t, "property " + t.Property + " after the embedded message (moved before it when encoded)");
                    a.Properties.Add(t.Property);
                }
            }
        }

        // ---- encode ------------------------------------------------------------------------------------------------------

        /// <summary>The FastTransfer stream of a message as a sequence of parts: bytes in memory and the lazy values to copy from their source.</summary>
        internal static IEnumerable<FxPart> Parts(FxMessage message)
        {
            if (message == null) throw new ArgumentNullException("message");
            return new PartWriter().Parts(message);
        }

        /// <summary>Length of the FastTransfer stream of a message (the lazy values are not read).</summary>
        public static long PayloadLength(FxMessage message)
        {
            long n = 0;
            foreach (FxPart p in Parts(message)) n += p.Length;
            return n;
        }

        /// <summary>The FastTransfer stream of a message alone (no envelope).</summary>
        public static byte[] EncodePayload(FxMessage message)
        {
            using (var ms = new MemoryStream())
            {
                EncodePayload(message, ms);
                return ms.ToArray();
            }
        }

        /// <summary>Writes the stream; the lazy values are copied from their source in pieces (never held whole).</summary>
        public static void EncodePayload(FxMessage message, Stream dest)
        {
            EncodePayload(message, dest, System.Threading.CancellationToken.None, null);
        }

        /// <summary>Same, cancellable between two pieces of a big value; onBytes is told the size of each piece written.</summary>
        public static void EncodePayload(FxMessage message, Stream dest, System.Threading.CancellationToken ct, Action<long> onBytes)
        {
            foreach (FxPart part in Parts(message)) { ct.ThrowIfCancellationRequested(); part.WriteTo(dest, ct, onBytes); }
        }

        public static byte[] Encode(FxDocument doc)
        {
            using (var ms = new MemoryStream())
            {
                Encode(doc, ms);
                return ms.ToArray();
            }
        }

        /// <summary>
        /// The records of the envelope with the data records laid out for a stream of <paramref name="total"/> bytes: as they were read if
        /// they add up to that length, else one run of records of 31,680 bytes where the first data record was.
        /// </summary>
        internal static List<FxEnvelopeEntry> BuildLayout(FxEnvelope envelope, long total)
        {
            List<FxEnvelopeEntry> entries = envelope.Entries;
            var layout = new List<FxEnvelopeEntry>();
            if (entries.Where(e => e.IsData).Sum(e => e.Length) == total)
                layout.AddRange(entries);
            else
            {
                bool dataDone = false;
                foreach (FxEnvelopeEntry e in entries)
                {
                    if (!e.IsData) { layout.Add(e); continue; }
                    if (dataDone) continue;
                    dataDone = true;
                    for (long left = total; left > 0; left -= DefaultChunk) layout.Add(new FxEnvelopeEntry { Type = 2, Length = Math.Min(DefaultChunk, left) });
                }
                if (!dataDone) for (long left = total; left > 0; left -= DefaultChunk) layout.Add(new FxEnvelopeEntry { Type = 2, Length = Math.Min(DefaultChunk, left) });
            }
            return layout;
        }

        /// <summary>
        /// The file: the envelope (if the document has one) with the stream cut in the records it was read in, or in records of
        /// 31,680 bytes if the message changed size; else the bare stream.
        /// </summary>
        public static void Encode(FxDocument doc, Stream dest)
        {
            if (doc == null) throw new ArgumentNullException("doc");
            if (doc.Envelope == null) { EncodePayload(doc.Message, dest); return; }
            List<FxEnvelopeEntry> layout = BuildLayout(doc.Envelope, PayloadLength(doc.Message));
            var ew = new EnvelopeWriter(dest, layout);
            EncodePayload(doc.Message, ew);
            ew.Finish();
        }

        /// <summary>Receives the stream and writes it into the data records of the layout, between the other records.</summary>
        sealed class EnvelopeWriter : Stream
        {
            readonly Stream _dest;
            readonly List<FxEnvelopeEntry> _layout;
            int _next;
            long _left;

            public EnvelopeWriter(Stream dest, List<FxEnvelopeEntry> layout) { _dest = dest; _layout = layout; }

            public override bool CanRead { get { return false; } }
            public override bool CanSeek { get { return false; } }
            public override bool CanWrite { get { return true; } }
            public override long Length { get { throw new NotSupportedException(); } }
            public override long Position { get { throw new NotSupportedException(); } set { throw new NotSupportedException(); } }
            public override void Flush() { _dest.Flush(); }
            public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }

            /// <summary>Writes the records up to the next data record (header included, payload not) and returns true if there is one.</summary>
            bool OpenNextData()
            {
                while (_next < _layout.Count)
                {
                    FxEnvelopeEntry e = _layout[_next++];
                    _dest.Write(BitConverter.GetBytes(e.Type), 0, 4);
                    _dest.Write(BitConverter.GetBytes((uint)e.Length), 0, 4);
                    if (e.IsData) { _left = e.Length; if (_left > 0) return true; continue; }
                    if (e.Payload != null) _dest.Write(e.Payload, 0, e.Payload.Length);
                }
                return false;
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                while (count > 0)
                {
                    while (_left == 0) if (!OpenNextData()) throw new InvalidOperationException("The envelope has fewer data records than the stream needs.");
                    int n = (int)Math.Min(count, _left);
                    _dest.Write(buffer, offset, n);
                    offset += n; count -= n; _left -= n;
                }
            }

            /// <summary>The records after the last data record.</summary>
            public void Finish()
            {
                if (_left != 0) throw new InvalidOperationException("The stream is shorter than the data records of the envelope.");
                OpenNextData();
            }
        }

        /// <summary>
        /// Little-endian producer of the elements of a stream, as parts: the small pieces are gathered in buffers (a buffer, once handed
        /// out, is never written again), a value that is lazy (or big) is a part of its own, copied from its source by whoever reads.
        /// </summary>
        sealed class PartWriter
        {
            const int BufSize = 16 * 1024, Direct = 8 * 1024;
            byte[] _buf;
            int _n;
            readonly Queue<FxPart> _q = new Queue<FxPart>();

            void Flush()
            {
                if (_n > 0) { _q.Enqueue(new FxPart { Bytes = _buf, Count = _n }); _buf = null; _n = 0; }
            }

            void Put(byte[] b, int offset, int count)
            {
                if (count == 0) return;
                if (count >= Direct) { Flush(); _q.Enqueue(new FxPart { Bytes = b, Offset = offset, Count = count }); return; }
                if (_buf == null) _buf = new byte[BufSize];
                if (_n + count > _buf.Length) { Flush(); _buf = new byte[BufSize]; }
                Buffer.BlockCopy(b, offset, _buf, _n, count);
                _n += count;
            }

            void Put(byte[] b) { Put(b, 0, b.Length); }
            void U16(ushort v) { Put(BitConverter.GetBytes(v)); }
            void U32(uint v) { Put(BitConverter.GetBytes(v)); }

            void PutBlob(FxBlob blob)
            {
                Flush();
                if (blob.Length > 0) _q.Enqueue(new FxPart { Blob = blob });
            }

            IEnumerable<FxPart> Drain()
            {
                while (_q.Count > 0) yield return _q.Dequeue();
            }

            public IEnumerable<FxPart> Parts(FxMessage m)
            {
                if (m.StartMarker != 0) U32(m.StartMarker);
                foreach (FxPart x in Content(m)) yield return x;
                if (m.StartMarker != 0) U32(FxMessage.MarkerEndMessage);
                Flush();
                foreach (FxPart x in Drain()) yield return x;
            }

            IEnumerable<FxPart> Content(FxMessage m)
            {
                foreach (FxProperty p in m.Properties) { Property(p); foreach (FxPart x in Drain()) yield return x; }
                if (m.DelPropRecipients != null) Property(m.DelPropRecipients);
                foreach (FxRecipient r in m.Recipients)
                {
                    U32(MarkerStartRecip);
                    foreach (FxProperty p in r.Properties) { Property(p); foreach (FxPart x in Drain()) yield return x; }
                    U32(MarkerEndToRecip);
                }
                if (m.DelPropAttachments != null) Property(m.DelPropAttachments);
                foreach (FxAttachment a in m.Attachments)
                {
                    U32(MarkerNewAttach);
                    foreach (FxProperty p in a.Properties) { Property(p); foreach (FxPart x in Drain()) yield return x; }
                    if (a.Embedded != null)
                    {
                        U32(MarkerStartEmbed);
                        foreach (FxPart x in Content(a.Embedded)) yield return x;
                        U32(MarkerEndEmbed);
                    }
                    U32(MarkerEndAttach);
                }
            }

            static ArgumentException Bad(FxProperty p, string why) { return new ArgumentException("Property " + p + " cannot be encoded: " + why); }

            void Property(FxProperty p)
            {
                FxTypeInfo t = FxTypes.Describe(p.Type);
                if (!t.Valid) throw Bad(p, "unknown type");
                if (p.IsNamed != (p.NamedGuid != null)) throw Bad(p, p.IsNamed ? "named property without a property set" : "a property set on an id below 0x8000");
                U16((ushort)(p.Extended ? p.Type | 0x4000 : p.Type)); U16(p.Id);
                if (p.Extended) U16(p.ExtendedFlags);
                if (p.IsNamed)
                {
                    Put(p.NamedGuid.Value.ToByteArray());
                    if (p.NamedLid != null) { Put(new byte[] { 0 }); U32(p.NamedLid.Value); }
                    else if (p.NamedName != null)
                    {
                        Put(new byte[] { 1 });
                        Put(Encoding.Unicode.GetBytes(p.NamedName));
                        Put(new byte[] { 0, 0 });
                    }
                    else throw Bad(p, "named property with neither a LID nor a name");
                }

                if (!t.MultiValued)
                {
                    if (p.IsLazy)
                    {
                        if (t.FixedSize >= 0 || t.IsObject) throw Bad(p, "only a variable-size value can be lazy");
                        long length = p.Lazy.Length;
                        if (length > uint.MaxValue) throw Bad(p, "the value has " + length + " bytes, the length of a value is 32 bits in a FastTransfer stream (less than 4 GiB)");
                        U32((uint)length);
                        PutBlob(p.Lazy);
                        return;
                    }
                    byte[] raw = p.Raw ?? new byte[0];
                    if (t.FixedSize >= 0)
                    {
                        if (raw.Length != t.FixedSize) throw Bad(p, "the value has " + raw.Length + " bytes, the type needs " + t.FixedSize);
                        Put(raw);
                    }
                    else if (t.IsObject && p.BareObject) { if (raw.Length != 0) throw Bad(p, "a bare object has no bytes"); }
                    else { U32((uint)raw.Length); Put(raw); }
                    return;
                }

                List<byte[]> items = p.Items ?? new List<byte[]>();
                U32((uint)items.Count);
                foreach (byte[] raw in items)
                {
                    if (t.FixedSize >= 0)
                    {
                        if (raw.Length != t.FixedSize) throw Bad(p, "a value has " + raw.Length + " bytes, the type needs " + t.FixedSize);
                        Put(raw);
                    }
                    else { U32((uint)raw.Length); Put(raw); }
                }
            }
        }
        // ---- round trip --------------------------------------------------------------------------------------------------

        public static FxRoundTripResult RoundTrip(byte[] bytes)
        {
            return RoundTrip(() => new MemoryStream(bytes, false));
        }

        /// <summary>The same for a file, with the large values left in the file (memory does not depend on their size).</summary>
        public static FxRoundTripResult RoundTripFile(string path)
        {
            return RoundTrip(FxStreamSource.FromFile(path));
        }

        /// <summary>Decodes the source (large values stay lazy), encodes the document again and compares with the source while writing.</summary>
        public static FxRoundTripResult RoundTrip(IFxStreamSource source)
        {
            var result = new FxRoundTripResult();
            result.Document = Decode(source);
            FxDocument doc = result.Document;
            using (Stream b = source.OpenRead())
            {
                var cmp = new CompareStream(b);
                try { Encode(doc, cmp); }
                catch (Exception ex) { result.Note = "encoding failed: " + ex.Message; }
                cmp.Finish();
                result.EncodedBytes = cmp.Written;
                result.OriginalBytes = cmp.OriginalBytes;
                result.FirstDifference = cmp.FirstDifference;
            }
            result.Identical = result.FirstDifference < 0 && result.Note.Length == 0 && doc.Error == null && !doc.Truncated;
            if (!result.Identical && result.Note.Length == 0)
                result.Note = doc.Error ?? (doc.Truncated ? "the stream is truncated" : doc.Warnings.Count > 0 ? doc.Warnings[0] : "the encoded stream differs");
            return result;
        }

        /// <summary>
        /// Decodes the stream, encodes the document again and compares with the original while writing (nothing is kept in memory
        /// besides the document). <paramref name="open"/> is called twice and must give the same bytes.
        /// </summary>
        public static FxRoundTripResult RoundTrip(Func<Stream> open)
        {
            var result = new FxRoundTripResult();
            using (Stream a = open())
                result.Document = Decode(a);
            FxDocument doc = result.Document;
            using (Stream b = open())
            {
                var cmp = new CompareStream(b);
                try { Encode(doc, cmp); }
                catch (Exception ex) { result.Note = "encoding failed: " + ex.Message; }
                cmp.Finish();
                result.EncodedBytes = cmp.Written;
                result.OriginalBytes = cmp.OriginalBytes;
                result.FirstDifference = cmp.FirstDifference;
            }
            result.Identical = result.FirstDifference < 0 && result.Note.Length == 0 && doc.Error == null && !doc.Truncated;
            if (!result.Identical && result.Note.Length == 0)
                result.Note = doc.Error ?? (doc.Truncated ? "the stream is truncated" : doc.Warnings.Count > 0 ? doc.Warnings[0] : "the encoded stream differs");
            return result;
        }

        /// <summary>A write-only stream that compares what it receives with another stream.</summary>
        sealed class CompareStream : Stream
        {
            readonly Stream _original;
            readonly byte[] _buf = new byte[64 * 1024];
            long _consumed;
            public long Written;
            public long FirstDifference = -1;
            /// <summary>Length of the original (known once Finish has run).</summary>
            public long OriginalBytes;

            public CompareStream(Stream original) { _original = original; }

            public override bool CanRead { get { return false; } }
            public override bool CanSeek { get { return false; } }
            public override bool CanWrite { get { return true; } }
            public override long Length { get { return Written; } }
            public override long Position { get { return Written; } set { throw new NotSupportedException(); } }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
            public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
            public override void SetLength(long value) { throw new NotSupportedException(); }

            int ReadOriginal(int want)
            {
                int got = 0;
                while (got < want)
                {
                    int n = _original.Read(_buf, got, want - got);
                    if (n <= 0) break;
                    got += n;
                }
                _consumed += got;
                return got;
            }

            public override void Write(byte[] buffer, int offset, int count)
            {
                int done = 0;
                while (done < count && FirstDifference < 0)
                {
                    int want = Math.Min(_buf.Length, count - done);
                    int got = ReadOriginal(want);
                    for (int i = 0; i < got; i++)
                        if (_buf[i] != buffer[offset + done + i]) { FirstDifference = Written + done + i; break; }
                    if (FirstDifference < 0 && got < want) FirstDifference = Written + done + got;   // the original is shorter
                    done += want;
                }
                Written += count;
            }

            /// <summary>Called when everything was written: the original must be at its end too.</summary>
            public void Finish()
            {
                if (FirstDifference < 0 && ReadOriginal(1) > 0) FirstDifference = Written;   // the original is longer
                int n;
                while ((n = _original.Read(_buf, 0, _buf.Length)) > 0) _consumed += n;
                OriginalBytes = _consumed;
            }
        }
    }
}
