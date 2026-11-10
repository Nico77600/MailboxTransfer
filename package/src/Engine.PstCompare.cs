// =============================================================================
//  Mailbox Transfer - engine, part 5f: comparison of a message written to a PST and read back
// -----------------------------------------------------------------------------
//  What the PST does not keep (PstConvert.NotStored, PidTagMessageSize, PidTagAttachSize, PidTagAttachNumber) is ignored in the
//  original; what the writer adds when it is missing (message class, conversation topic, times, search key, status, display lines, attachment method
//  and position) is ignored in the copy. A string compares by text (the PST stores Unicode only: a string8 or a code page string
//  comes back as PtypString), a Boolean by value, everything else byte for byte.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MailboxTransfer
{
    public static class PstFxCompare
    {
        static readonly HashSet<ushort> MsgIgnoreOriginal = new HashSet<ushort>(PstConvert.NotStored.Concat(new ushort[] { 0x0E08 }));
        static readonly HashSet<ushort> MsgAdded = new HashSet<ushort> { 0x001A, 0x0070, 0x0E02, 0x0E03, 0x0E04, 0x0E07, 0x0E08, 0x0E17, 0x3007, 0x3008, 0x300B };
        static readonly HashSet<ushort> AttIgnoreOriginal = new HashSet<ushort>(PstConvert.NotStored.Concat(new ushort[] { 0x0E20, 0x0E21, 0x3701 }));
        static readonly HashSet<ushort> AttAdded = new HashSet<ushort> { 0x0E20, 0x0E21, 0x3705, 0x370B, 0x3701 };

        static string Key(FxProperty p)
        {
            if (p.IsNamed && p.NamedGuid != null)
                return p.NamedGuid.Value.ToString("N") + (p.NamedLid != null ? ":L" + p.NamedLid.Value : ":S" + p.NamedName);
            return p.Id.ToString("X4");
        }

        static string Describe(FxProperty p) { return p.IsNamed ? "named " + p.NamedText : "0x" + p.Id.ToString("X4"); }

        static bool SameBytes(byte[] a, byte[] b)
        {
            if (a == null || b == null) return a == b;
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
            return true;
        }

        /// <summary>Compares two values read in pieces (nothing is held whole).</summary>
        static bool BlobsEqual(FxBlob a, FxBlob b)
        {
            if (a.Length != b.Length) return false;
            byte[] x = System.Buffers.ArrayPool<byte>.Shared.Rent(64 * 1024), y = System.Buffers.ArrayPool<byte>.Shared.Rent(64 * 1024);
            try
            {
                using (Stream sa = a.OpenRead())
                using (Stream sb = b.OpenRead())
                {
                    long left = a.Length;
                    while (left > 0)
                    {
                        int n = (int)Math.Min(x.Length, left);
                        if (FxIo.ReadFull(sa, x, 0, n) < n || FxIo.ReadFull(sb, y, 0, n) < n) return false;
                        for (int i = 0; i < n; i++) if (x[i] != y[i]) return false;
                        left -= n;
                    }
                }
                return true;
            }
            finally { System.Buffers.ArrayPool<byte>.Shared.Return(x); System.Buffers.ArrayPool<byte>.Shared.Return(y); }
        }

        static bool Utf16(FxTypeInfo t) { return t.CodePageString ? t.CodePage == 1200 : t.BaseType == 0x001F; }

        static FxBlob WithoutTerminator(FxBlob b)
        {
            long n = b.Length;
            if (n >= 2) { byte[] tail = b.ReadRange(n - 2, 2); if (tail[0] == 0 && tail[1] == 0) return b.Slice(0, n - 2); }
            return b;
        }

        static string Text(FxTypeInfo t, byte[] raw)
        {
            return FxValue.DecodeString(raw, t.BaseType, t.CodePageString ? t.CodePage : 1252).TrimEnd('\0');
        }

        static bool SameValue(FxProperty a, FxProperty b, out string why)
        {
            why = null;
            FxTypeInfo ta = a.TypeInfo, tb = b.TypeInfo;
            if (ta.MultiValued != tb.MultiValued) { why = "single / multi-valued differs"; return false; }
            bool sa = ta.IsString, sb = tb.IsString;
            if (sa != sb) { why = "type 0x" + a.Type.ToString("X4") + " became 0x" + b.Type.ToString("X4"); return false; }
            if (sa)
            {
                if (!ta.MultiValued && (a.IsLazy || b.IsLazy) && Utf16(ta) && Utf16(tb))
                {
                    if (!BlobsEqual(WithoutTerminator(a.Blob), WithoutTerminator(b.Blob))) { why = "text differs"; return false; }
                    return true;
                }
                if (!ta.MultiValued) { if (Text(ta, a.Raw) != Text(tb, b.Raw)) { why = "text differs"; return false; } return true; }
                if (a.Items.Count != b.Items.Count) { why = "number of values " + a.Items.Count + " became " + b.Items.Count; return false; }
                for (int i = 0; i < a.Items.Count; i++) if (Text(ta, a.Items[i]) != Text(tb, b.Items[i])) { why = "text of value " + i + " differs"; return false; }
                return true;
            }
            if (a.Type != b.Type && !(ta.BaseType == 0x00FB && tb.BaseType == 0x0102)) { why = "type 0x" + a.Type.ToString("X4") + " became 0x" + b.Type.ToString("X4"); return false; }
            if (ta.BaseType == 0x000B && !ta.MultiValued)
            {
                if (a.AsBool() != b.AsBool()) { why = "Boolean differs"; return false; }
                return true;
            }
            if (!ta.MultiValued)
            {
                if (a.IsLazy || b.IsLazy)
                {
                    FxBlob x = a.Blob, y = b.Blob;
                    if (x == null || y == null ? x != y : !BlobsEqual(x, y)) { why = "value differs (" + a.RawLength + " / " + b.RawLength + " bytes)"; return false; }
                    return true;
                }
                if (!SameBytes(a.Raw, b.Raw)) { why = "value differs (" + (a.Raw == null ? 0 : a.Raw.Length) + " / " + (b.Raw == null ? 0 : b.Raw.Length) + " bytes)"; return false; }
                return true;
            }
            if (a.Items.Count != b.Items.Count) { why = "number of values " + a.Items.Count + " became " + b.Items.Count; return false; }
            for (int i = 0; i < a.Items.Count; i++) if (!SameBytes(a.Items[i], b.Items[i])) { why = "value " + i + " differs"; return false; }
            return true;
        }

        static void CompareProps(string where, IEnumerable<FxProperty> orig, IEnumerable<FxProperty> copy, HashSet<ushort> ignoreOrig, HashSet<ushort> added, List<string> diffs, bool message)
        {
            var c = new Dictionary<string, FxProperty>();
            foreach (FxProperty p in copy) c[Key(p)] = p;
            var used = new HashSet<string>();
            foreach (FxProperty p in orig)
            {
                if (!p.IsNamed && ignoreOrig.Contains(p.Id)) continue;
                FxTypeInfo t = p.TypeInfo;
                if (!t.Valid || t.BaseType == 0x0001 || t.IsObject) continue;
                string k = Key(p);
                FxProperty q;
                if (!c.TryGetValue(k, out q)) { diffs.Add(where + ": " + Describe(p) + " (type 0x" + p.Type.ToString("X4") + ") is missing in the copy"); continue; }
                used.Add(k);
                string why;
                if (message && !p.IsNamed && p.Id == 0x0E07)
                {
                    if ((p.AsInt32() & ~0x10) != (q.AsInt32() & ~0x10)) diffs.Add(where + ": PidTagMessageFlags differs (" + p.AsInt32().ToString("X") + " / " + q.AsInt32().ToString("X") + ")");
                    continue;
                }
                if (!SameValue(p, q, out why)) diffs.Add(where + ": " + Describe(p) + ": " + why);
            }
            foreach (KeyValuePair<string, FxProperty> kv in c)
            {
                if (used.Contains(kv.Key)) continue;
                FxProperty q = kv.Value;
                if (!q.IsNamed && added.Contains(q.Id)) continue;
                diffs.Add(where + ": " + Describe(q) + " appears in the copy only");
            }
        }

        /// <summary>Appends one line per difference; returns true if there is none.</summary>
        public static bool Compare(FxMessage original, FxMessage copy, List<string> diffs)
        {
            int before = diffs.Count;
            Compare("message", original, copy, diffs);
            return diffs.Count == before;
        }

        static void Compare(string where, FxMessage a, FxMessage b, List<string> diffs)
        {
            CompareProps(where, a.Properties, b.Properties, MsgIgnoreOriginal, MsgAdded, diffs, true);
            if (a.Recipients.Count != b.Recipients.Count) diffs.Add(where + ": " + a.Recipients.Count + " recipients became " + b.Recipients.Count);
            else
                for (int i = 0; i < a.Recipients.Count; i++)
                    CompareProps(where + " recipient " + i, a.Recipients[i].Properties, b.Recipients[i].Properties, PstConvert.NotStored, new HashSet<ushort>(), diffs, false);
            if (a.Attachments.Count != b.Attachments.Count) { diffs.Add(where + ": " + a.Attachments.Count + " attachments became " + b.Attachments.Count); return; }
            for (int i = 0; i < a.Attachments.Count; i++)
            {
                FxAttachment x = a.Attachments[i], y = b.Attachments[i];
                string w = where + " attachment " + i;
                CompareProps(w, x.Properties, y.Properties, AttIgnoreOriginal, AttAdded, diffs, false);
                FxProperty xd = x.Properties.FirstOrDefault(p => p.Id == 0x3701 && !p.IsNamed), yd = y.Properties.FirstOrDefault(p => p.Id == 0x3701 && !p.IsNamed);
                if (x.Embedded != null)
                {
                    if (y.Embedded == null) diffs.Add(w + ": the embedded message is missing in the copy");
                    else Compare(w + " / embedded", x.Embedded, y.Embedded, diffs);
                }
                else if (xd != null)
                {
                    if (yd == null) diffs.Add(w + ": PidTagAttachDataBinary is missing in the copy");
                    else { string why; if (!SameValue(xd, yd, out why)) diffs.Add(w + ": PidTagAttachDataBinary: " + why); }
                }
            }
        }
    }
}
