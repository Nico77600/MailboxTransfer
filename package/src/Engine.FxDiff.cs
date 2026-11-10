// =============================================================================
//  Mailbox Transfer - engine, part 5g: difference between a message sent to Exchange and the same message exported again
// -----------------------------------------------------------------------------
//  FxDiff.Compare(sent, got) lists, property by property, what Exchange kept (nothing listed), changed, dropped or added. Used by the
//  import-from-PST verification of the lab tools (tools\Test-PstImportLive.ps1) and by the tests; the result is a plain list, so a UI
//  can show it. Recipients and attachments are matched by position, embedded messages follow their attachment.
//  A string compares by text (the terminator and the string8 / Unicode form do not count), everything else byte for byte; a value that
//  is lazy is compared in pieces (never held whole).
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace MailboxTransfer
{
    public enum FxDiffKind { Changed, Dropped, Added }

    public sealed class FxDiffEntry
    {
        public FxDiffKind Kind;
        /// <summary>"message", "recipient 2", "attachment 1", "attachment 1 / embedded message", ...</summary>
        public string Where;
        public string PropertyKey;
        public string PropertyName;
        /// <summary>The value sent and the value found (short forms: text, number or the first bytes in hex), null when there is none.</summary>
        public string Sent, Got;
        public string Detail;
        public override string ToString()
        {
            return Kind + " " + Where + " " + PropertyName + (Detail != null ? " (" + Detail + ")" : "") +
                (Sent != null || Got != null ? ": " + (Sent ?? "-") + " -> " + (Got ?? "-") : "");
        }
    }

    public static class FxDiff
    {
        public static List<FxDiffEntry> Compare(FxMessage sent, FxMessage got)
        {
            var list = new List<FxDiffEntry>();
            CompareMessage("message", sent, got, list);
            return list;
        }

        /// <summary>
        /// Recipients, attachments and embedded messages of a message, embedded messages included (a recipient or an attachment the service silently
        /// dropped must never pass a verification: compare these totals of the message sent and of the message exported again).
        /// </summary>
        public static void Totals(FxMessage m, out int recipients, out int attachments, out int embedded)
        {
            recipients = m.Recipients.Count; attachments = m.Attachments.Count; embedded = 0;
            foreach (FxAttachment a in m.Attachments)
            {
                if (a.Embedded == null) continue;
                int r, at, e;
                Totals(a.Embedded, out r, out at, out e);
                recipients += r; attachments += at; embedded += 1 + e;
            }
        }

        /// <summary>"3 recipients, 2 attachments, 1 embedded message" (the same text for the same totals).</summary>
        public static string TotalsText(FxMessage m)
        {
            int r, a, e;
            Totals(m, out r, out a, out e);
            return r + " recipients, " + a + " attachments, " + e + " embedded messages";
        }
        static string KeyOf(FxProperty p)
        {
            if (p.IsNamed && p.NamedGuid != null) return p.NamedGuid.Value.ToString("N") + (p.NamedLid != null ? ":L" + p.NamedLid.Value : ":S" + p.NamedName);
            return p.Id.ToString("X4");
        }

        static string NameOf(FxProperty p)
        {
            if (p.IsNamed) return "named " + p.NamedText + " " + FxNames.TypeName(p.Type);
            string n = FxNames.Name(p.Id);
            return "0x" + p.Id.ToString("X4") + (n != null ? " " + n : "") + " " + FxNames.TypeName(p.Type);
        }

        static void CompareMessage(string where, FxMessage a, FxMessage b, List<FxDiffEntry> list)
        {
            CompareProps(where, a.Properties, b.Properties, list);
            if (a.Recipients.Count != b.Recipients.Count)
                list.Add(new FxDiffEntry { Kind = FxDiffKind.Changed, Where = where, PropertyKey = "recipients", PropertyName = "recipients", Sent = a.Recipients.Count.ToString(), Got = b.Recipients.Count.ToString(), Detail = "number of recipients" });
            for (int i = 0; i < Math.Min(a.Recipients.Count, b.Recipients.Count); i++)
                CompareProps(where + " / recipient " + (i + 1), a.Recipients[i].Properties, b.Recipients[i].Properties, list);
            if (a.Attachments.Count != b.Attachments.Count)
                list.Add(new FxDiffEntry { Kind = FxDiffKind.Changed, Where = where, PropertyKey = "attachments", PropertyName = "attachments", Sent = a.Attachments.Count.ToString(), Got = b.Attachments.Count.ToString(), Detail = "number of attachments" });
            for (int i = 0; i < Math.Min(a.Attachments.Count, b.Attachments.Count); i++)
            {
                string w = where + " / attachment " + (i + 1);
                FxAttachment x = a.Attachments[i], y = b.Attachments[i];
                CompareProps(w, x.Properties, y.Properties, list);
                if ((x.Embedded == null) != (y.Embedded == null))
                    list.Add(new FxDiffEntry { Kind = x.Embedded == null ? FxDiffKind.Added : FxDiffKind.Dropped, Where = w, PropertyKey = "embedded", PropertyName = "embedded message" });
                else if (x.Embedded != null) CompareMessage(w + " / embedded message", x.Embedded, y.Embedded, list);
            }
        }

        static void CompareProps(string where, List<FxProperty> a, List<FxProperty> b, List<FxDiffEntry> list)
        {
            // Same identity (id or named identity) with a different type is a change of type, not a drop and an addition.
            var bm = new Dictionary<string, FxProperty>();
            foreach (FxProperty p in b) bm[KeyOf(p)] = p;
            var seen = new HashSet<string>();
            foreach (FxProperty p in a)
            {
                string k = KeyOf(p);
                if (!seen.Add(k)) continue;
                FxProperty q;
                if (!bm.TryGetValue(k, out q)) { list.Add(new FxDiffEntry { Kind = FxDiffKind.Dropped, Where = where, PropertyKey = k, PropertyName = NameOf(p), Sent = Short(p) }); continue; }
                string why;
                if (!SameValue(p, q, out why))
                    list.Add(new FxDiffEntry { Kind = FxDiffKind.Changed, Where = where, PropertyKey = k, PropertyName = NameOf(p), Sent = Short(p), Got = Short(q), Detail = why });
            }
            foreach (FxProperty q in b)
            {
                string k = KeyOf(q);
                if (seen.Add(k)) list.Add(new FxDiffEntry { Kind = FxDiffKind.Added, Where = where, PropertyKey = k, PropertyName = NameOf(q), Got = Short(q) });
            }
        }

        static string Short(FxProperty p)
        {
            try
            {
                FxTypeInfo t = p.TypeInfo;
                if (t.MultiValued)
                {
                    if (p.Items == null) return "(no values)";
                    return p.Items.Count + " values: " + string.Join(" | ", p.Items.Take(4).Select(v => ShortValue(t, v))) + (p.Items.Count > 4 ? " ..." : "");
                }
                if (p.RawLength > 256 && !t.IsString) return p.RawLength + " bytes";
                if (p.RawLength > 2048) return p.RawLength + " bytes of text";
                return ShortValue(t, p.Raw);
            }
            catch (Exception ex) { return "(" + ex.Message + ")"; }
        }

        static string ShortValue(FxTypeInfo t, byte[] raw)
        {
            if (raw == null) return "";
            if (t.IsString)
            {
                string s = FxValue.DecodeString(raw, t.BaseType, t.CodePageString ? t.CodePage : 1252).TrimEnd('\0');
                return s.Length > 90 ? s.Substring(0, 90) + "..." : s;
            }
            switch (t.BaseType)
            {
                case 0x0002: return BitConverter.ToInt16(raw, 0).ToString();
                case 0x0003: return BitConverter.ToInt32(raw, 0).ToString();
                case 0x0014: return BitConverter.ToInt64(raw, 0).ToString();
                case 0x000B: return (raw.Length > 0 && raw[0] != 0) ? "true" : "false";
                case 0x0040: { DateTimeOffset? d = FxValue.ToTime(raw); return d == null ? "(time)" : d.Value.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss'Z'"); }
                case 0x0048: return raw.Length == 16 ? new Guid(raw).ToString() : "(guid)";
            }
            return raw.Length + " bytes " + BitConverter.ToString(raw, 0, Math.Min(raw.Length, 12));
        }

        static bool SameValue(FxProperty a, FxProperty b, out string why)
        {
            why = null;
            FxTypeInfo ta = a.TypeInfo, tb = b.TypeInfo;
            if (ta.MultiValued != tb.MultiValued) { why = "single / multi-valued differs"; return false; }
            if (ta.IsString != tb.IsString) { why = "type " + FxNames.TypeName(a.Type) + " became " + FxNames.TypeName(b.Type); return false; }
            if (ta.IsString)
            {
                if (!ta.MultiValued)
                {
                    if (a.IsLazy || b.IsLazy)
                    {
                        // A big text: compared as bytes after the terminator is left out, only when both are UTF-16 (else decoded in memory).
                        if (ta.BaseType == 0x001F && tb.BaseType == 0x001F)
                        {
                            if (!BlobsEqual(WithoutTerminator(a.Blob), WithoutTerminator(b.Blob))) { why = "text differs"; return false; }
                            return true;
                        }
                    }
                    if (Text(ta, a.Raw) != Text(tb, b.Raw)) { why = "text differs"; return false; }
                    return true;
                }
                if (a.Items.Count != b.Items.Count) { why = "number of values " + a.Items.Count + " became " + b.Items.Count; return false; }
                for (int i = 0; i < a.Items.Count; i++) if (Text(ta, a.Items[i]) != Text(tb, b.Items[i])) { why = "text of value " + i + " differs"; return false; }
                return true;
            }
            if (a.Type != b.Type) { why = "type " + FxNames.TypeName(a.Type) + " became " + FxNames.TypeName(b.Type); return false; }
            if (ta.MultiValued)
            {
                if (a.Items.Count != b.Items.Count) { why = "number of values " + a.Items.Count + " became " + b.Items.Count; return false; }
                for (int i = 0; i < a.Items.Count; i++) if (!SameBytes(a.Items[i], b.Items[i])) { why = "value " + i + " differs"; return false; }
                return true;
            }
            if (ta.BaseType == 0x000B) { if (a.AsBool() != b.AsBool()) { why = "Boolean differs"; return false; } return true; }
            FxBlob x = a.Blob, y = b.Blob;
            if (x == null || y == null) { if (x != y) { why = "value differs"; return false; } return true; }
            if (!BlobsEqual(x, y)) { why = "value differs (" + x.Length + " / " + y.Length + " bytes)"; return false; }
            return true;
        }

        static string Text(FxTypeInfo t, byte[] raw) { return FxValue.DecodeString(raw ?? new byte[0], t.BaseType, t.CodePageString ? t.CodePage : 1252).TrimEnd('\0'); }

        static bool SameBytes(byte[] a, byte[] b)
        {
            if (a == null || b == null) return a == b;
            if (a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
            return true;
        }

        static FxBlob WithoutTerminator(FxBlob b)
        {
            long n = b.Length;
            if (n >= 2) { byte[] tail = b.ReadRange(n - 2, 2); if (tail[0] == 0 && tail[1] == 0) return b.Slice(0, n - 2); }
            return b;
        }

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
    }
}
