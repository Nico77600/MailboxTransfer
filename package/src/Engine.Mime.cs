// =============================================================================
//  Mailbox Transfer - engine, part 12: MIME parser and writer (RFC 5322 / 2045-2049 / 2047 / 2231), no third-party code
// -----------------------------------------------------------------------------
//  MimeParser.Parse   bytes of an .eml  ->  MimeEntity tree (tolerant: it never throws on broken MIME, it notes what it repaired)
//  MimeWriter.Write   MimeEntity tree   ->  bytes of an .eml (CRLF, folded headers, RFC 2047 / 2231 for non-ASCII, base64 / quoted-printable)
//  MimeText           header helpers: encoded words, parameters, address lists, dates, quoted-printable and base64
//
//  The model is the MIME tree only. The mapping to and from the MAPI object model (FxMessage) is in Engine.EmlMapi.cs.
//  What is tolerated: bare LF, missing blank line, mbox "From " line, folded headers with junk, unterminated multipart, a boundary that
//  never shows up, 8-bit bytes in headers (UTF-8 else Windows-1252), unknown charsets and transfer encodings, damaged base64 and
//  quoted-printable, message/rfc822 with a transfer encoding, nesting up to 48 levels (deeper content is kept as an opaque part).
//  A parsed entity keeps its decoded content in memory (FxBytesBlob): the parser is meant for messages, not for gigabyte files.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;

namespace MailboxTransfer
{
    public sealed class MimeHeader
    {
        public string Name;
        /// <summary>The value unfolded (folding white space replaced by one space), not decoded.</summary>
        public string Value;
        /// <summary>The value as it was in the file, with its line breaks (used to rebuild PR_TRANSPORT_MESSAGE_HEADERS).</summary>
        public string RawLine;
        public string Decoded { get { return MimeText.DecodeWords(Value); } }
        public override string ToString() { return Name + ": " + Value; }
    }

    public sealed class MimeAddress
    {
        public string Name;
        public string Address;
        public override string ToString() { return string.IsNullOrEmpty(Name) ? Address : Name + " <" + Address + ">"; }
    }

    /// <summary>A structured header value: type/subtype and its parameters (RFC 2231 continuations and encodings are already resolved).</summary>
    public sealed class MimeContentType
    {
        public string Type = "text";
        public string SubType = "plain";
        public readonly Dictionary<string, string> Parameters = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public string MediaType { get { return Type + "/" + SubType; } }
        public string this[string name] { get { string v; return Parameters.TryGetValue(name, out v) ? v : null; } }
        public bool Is(string type, string sub = null) { return Type == type && (sub == null || SubType == sub); }

        public static MimeContentType Parse(string value)
        {
            var ct = new MimeContentType();
            if (string.IsNullOrWhiteSpace(value)) return ct;
            string first;
            Dictionary<string, string> p = MimeText.ParseParameters(value, out first);
            string media = (first ?? "").Trim().ToLowerInvariant();
            int slash = media.IndexOf('/');
            if (slash > 0 && slash < media.Length - 1) { ct.Type = media.Substring(0, slash).Trim(); ct.SubType = media.Substring(slash + 1).Trim(); }
            else if (media.Length > 0 && media.All(char.IsLetter)) { ct.Type = media; ct.SubType = media == "text" ? "plain" : "x-unknown"; }
            foreach (var kv in p) ct.Parameters[kv.Key] = kv.Value;
            return ct;
        }
    }

    public sealed class MimeEntity
    {
        public List<MimeHeader> Headers = new List<MimeHeader>();
        public MimeContentType ContentType = new MimeContentType();
        public bool HasContentType;
        /// <summary>Content-Transfer-Encoding as declared, lower case (7bit, 8bit, binary, base64, quoted-printable); null when absent. Parser: what the file said. Writer: what to produce.</summary>
        public string TransferEncoding;
        /// <summary>attachment or inline, null when there is no Content-Disposition.</summary>
        public string Disposition;
        public MimeContentType DispositionParams = new MimeContentType();
        /// <summary>Decoded content of a leaf part (no multipart, no message/rfc822).</summary>
        public FxBlob Content;
        public List<MimeEntity> Parts = new List<MimeEntity>();
        /// <summary>The message of a message/rfc822 part.</summary>
        public MimeEntity Message;
        /// <summary>What the parser repaired or could not understand (empty for a clean message).</summary>
        public List<string> Warnings = new List<string>();
        /// <summary>The header block exactly as read (parser only), from the first to the last header line, line breaks kept.</summary>
        public string RawHeaderBlock;

        public bool IsMultipart { get { return ContentType.Type == "multipart"; } }
        public bool IsMessage { get { return Message != null; } }

        public string Header(string name)
        {
            foreach (MimeHeader h in Headers) if (string.Equals(h.Name, name, StringComparison.OrdinalIgnoreCase)) return h.Value;
            return null;
        }

        public IEnumerable<string> HeaderAll(string name)
        {
            foreach (MimeHeader h in Headers) if (string.Equals(h.Name, name, StringComparison.OrdinalIgnoreCase)) yield return h.Value;
        }

        public string HeaderDecoded(string name) { string v = Header(name); return v == null ? null : MimeText.DecodeWords(v); }
        public void Add(string name, string value) { Headers.Add(new MimeHeader { Name = name, Value = value }); }

        public string ContentId
        {
            get
            {
                string v = Header("Content-ID");
                if (v == null) return null;
                v = v.Trim();
                if (v.StartsWith("<") && v.EndsWith(">")) v = v.Substring(1, v.Length - 2);
                return v.Length == 0 ? null : v;
            }
        }

        public string ContentLocation { get { string v = Header("Content-Location"); return v == null ? null : v.Trim(); } }

        /// <summary>File name from Content-Disposition (filename), else Content-Type (name); RFC 2231 and RFC 2047 forms resolved.</summary>
        public string FileName
        {
            get
            {
                string n = DispositionParams["filename"] ?? ContentType["name"];
                return n == null ? null : MimeText.DecodeWords(n).Trim();
            }
        }

        public long ContentLength { get { return Content == null ? 0 : Content.Length; } }

        /// <summary>The content decoded as text with the charset of the part (UTF-8 when there is none and the bytes are valid UTF-8, else Windows-1252).</summary>
        public string Text()
        {
            if (Content == null) return "";
            return MimeText.DecodeText(Content.ToArray(), ContentType["charset"]);
        }

        public IEnumerable<MimeEntity> Walk()
        {
            yield return this;
            foreach (MimeEntity p in Parts) foreach (MimeEntity c in p.Walk()) yield return c;
        }
    }

    public static class MimeParser
    {
        public const int MaxDepth = 48;

        public static MimeEntity Parse(byte[] data)
        {
            if (data == null) throw new ArgumentNullException("data");
            return ParseEntity(data, 0, data.Length, 0, true);
        }

        public static MimeEntity ParseFile(string path) { return Parse(File.ReadAllBytes(path)); }

        public static MimeEntity Parse(Stream s)
        {
            using (var ms = new MemoryStream()) { s.CopyTo(ms); return Parse(ms.ToArray()); }
        }

        // ---- line helpers --------------------------------------------------------------------------------------------------

        /// <summary>Index of the first byte of the next line (after LF), or end when there is none; contentEnd is the end of the line without its break.</summary>
        static int LineEnd(byte[] d, int pos, int end, out int contentEnd)
        {
            int i = pos;
            while (i < end && d[i] != (byte)'\n') i++;
            if (i >= end) { contentEnd = end; if (contentEnd > pos && d[contentEnd - 1] == (byte)'\r') contentEnd--; return end; }
            contentEnd = i;
            if (contentEnd > pos && d[contentEnd - 1] == (byte)'\r') contentEnd--;
            return i + 1;
        }

        static string HeaderString(byte[] d, int pos, int len)
        {
            if (len <= 0) return "";
            try { return new UTF8Encoding(false, true).GetString(d, pos, len); }
            catch (Exception) { return MimeText.Latin1252.GetString(d, pos, len); }
        }

        static MimeEntity ParseEntity(byte[] d, int start, int end, int depth, bool topLevel)
        {
            var e = new MimeEntity();
            int pos = start;
            // mbox separator line
            if (topLevel && end - pos > 5 && d[pos] == 'F' && d[pos + 1] == 'r' && d[pos + 2] == 'o' && d[pos + 3] == 'm' && d[pos + 4] == ' ')
            {
                int ce; pos = LineEnd(d, pos, end, out ce);
                e.Warnings.Add("mbox 'From ' line skipped");
            }
            int headerStart = pos;
            int bodyStart = end;
            bool sawBlank = false;
            MimeHeader current = null;
            var raw = new StringBuilder();
            int firstHeaderPos = pos, lastHeaderEnd = pos;
            while (pos < end)
            {
                int ce;
                int next = LineEnd(d, pos, end, out ce);
                if (ce == pos) { bodyStart = next; sawBlank = true; break; }   // empty line: end of the header block
                if (d[pos] == (byte)' ' || d[pos] == (byte)'\t')
                {
                    if (current != null)
                    {
                        string cont = HeaderString(d, pos, ce - pos);
                        current.Value = current.Value + " " + cont.Trim();
                        current.RawLine += "\r\n" + cont;
                        lastHeaderEnd = next;
                        pos = next;
                        continue;
                    }
                }
                int colon = pos;
                while (colon < ce && d[colon] != (byte)':' && d[colon] > 32 && d[colon] < 127) colon++;
                if (colon < ce && d[colon] == (byte)':' && colon > pos)
                {
                    string name = Encoding.ASCII.GetString(d, pos, colon - pos);
                    string value = HeaderString(d, colon + 1, ce - colon - 1).Trim();
                    current = new MimeHeader { Name = name, Value = value, RawLine = HeaderString(d, pos, ce - pos) };
                    e.Headers.Add(current);
                    lastHeaderEnd = next;
                }
                else
                {
                    if (e.Headers.Count == 0)
                    {
                        // the first line is not a header: the whole thing is a body without headers
                        e.Warnings.Add("no header block, content taken as text");
                        bodyStart = headerStart;
                        pos = end;
                        break;
                    }
                    if (!sawBlank && e.Headers.Count > 0 && !LooksLikeHeaderLine(d, pos, ce))
                    {
                        // junk inside the header block: a body that starts without the blank line is the likely cause
                        e.Warnings.Add("line without colon inside the headers, taken as the start of the body");
                        bodyStart = pos;
                        break;
                    }
                }
                pos = next;
                if (pos >= end) bodyStart = end;
            }
            if (!sawBlank && pos >= end && e.Headers.Count > 0 && bodyStart == end) e.Warnings.Add("no blank line after the headers");
            if (e.Headers.Count > 0)
            {
                int rawEnd = lastHeaderEnd;
                e.RawHeaderBlock = HeaderString(d, firstHeaderPos, Math.Max(0, rawEnd - firstHeaderPos));
            }

            string cts = e.Header("Content-Type");
            if (cts != null) { e.ContentType = MimeContentType.Parse(cts); e.HasContentType = true; }
            string cte = e.Header("Content-Transfer-Encoding");
            if (cte != null) e.TransferEncoding = cte.Trim().ToLowerInvariant();
            string cd = e.Header("Content-Disposition");
            if (cd != null)
            {
                string first;
                Dictionary<string, string> p = MimeText.ParseParameters(cd, out first);
                e.Disposition = (first ?? "").Trim().ToLowerInvariant();
                if (e.Disposition.Length == 0) e.Disposition = null;
                foreach (var kv in p) e.DispositionParams.Parameters[kv.Key] = kv.Value;
            }

            int bodyLen = Math.Max(0, end - bodyStart);
            if (e.ContentType.Type == "multipart")
            {
                string boundary = e.ContentType["boundary"];
                if (string.IsNullOrEmpty(boundary) || depth >= MaxDepth)
                {
                    e.Warnings.Add(depth >= MaxDepth ? "nesting too deep, kept as an opaque part" : "multipart without boundary, kept as an opaque part");
                    e.ContentType = new MimeContentType { Type = "application", SubType = "octet-stream" };
                    e.Content = new FxBytesBlob(Decode(d, bodyStart, bodyLen, e.TransferEncoding, e.Warnings));
                    return e;
                }
                ParseMultipart(e, d, bodyStart, end, boundary, depth);
                return e;
            }
            byte[] content = Decode(d, bodyStart, bodyLen, e.TransferEncoding, e.Warnings);
            if (e.ContentType.Type == "message" && (e.ContentType.SubType == "rfc822" || e.ContentType.SubType == "global") && depth < MaxDepth)
            {
                e.Message = ParseEntity(content, 0, content.Length, depth + 1, false);
                return e;
            }
            e.Content = new FxBytesBlob(content);
            return e;
        }

        static bool LooksLikeHeaderLine(byte[] d, int pos, int ce)
        {
            // "Name:" with a plausible name was handled above; here anything else is junk
            return false;
        }

        static void ParseMultipart(MimeEntity e, byte[] d, int bodyStart, int end, string boundary, int depth)
        {
            byte[] delim = Encoding.ASCII.GetBytes("--" + boundary);
            var starts = new List<int>();   // start of the content of each part
            var ends = new List<int>();     // end of the content of each part (before the line break that precedes the next delimiter)
            int pos = bodyStart;
            int partStart = -1;
            bool closed = false;
            while (pos < end)
            {
                int ce;
                int next = LineEnd(d, pos, end, out ce);
                int len = ce - pos;
                if (len >= delim.Length && Matches(d, pos, delim))
                {
                    bool isClose = len >= delim.Length + 2 && d[pos + delim.Length] == '-' && d[pos + delim.Length + 1] == '-';
                    bool pad = true;
                    int after = pos + delim.Length + (isClose ? 2 : 0);
                    for (int i = after; i < ce; i++) if (d[i] != ' ' && d[i] != '\t') { pad = false; break; }
                    if (pad)
                    {
                        if (partStart >= 0)
                        {
                            int partEnd = pos;
                            // the line break before the delimiter belongs to the delimiter
                            if (partEnd > partStart && d[partEnd - 1] == '\n') partEnd--;
                            if (partEnd > partStart && d[partEnd - 1] == '\r') partEnd--;
                            starts.Add(partStart); ends.Add(partEnd);
                        }
                        if (isClose) { closed = true; partStart = -1; break; }
                        partStart = next;
                    }
                }
                pos = next;
            }
            if (partStart >= 0 && !closed)
            {
                e.Warnings.Add("multipart not closed");
                int partEnd = end;
                starts.Add(partStart); ends.Add(Math.Max(partStart, partEnd));
            }
            if (starts.Count == 0)
            {
                e.Warnings.Add("multipart boundary not found, kept as text");
                e.ContentType = new MimeContentType { Type = "text", SubType = "plain" };
                e.Content = new FxBytesBlob(Slice(d, bodyStart, end - bodyStart));
                return;
            }
            for (int i = 0; i < starts.Count; i++)
            {
                MimeEntity child = ParseEntity(d, starts[i], ends[i], depth + 1, false);
                // a part with no Content-Type inside multipart/digest is a message
                if (!child.HasContentType && e.ContentType.SubType == "digest" && child.Content != null && child.Headers.Count > 0)
                {
                    child.ContentType = new MimeContentType { Type = "message", SubType = "rfc822" };
                }
                e.Parts.Add(child);
            }
        }

        static bool Matches(byte[] d, int pos, byte[] what)
        {
            for (int i = 0; i < what.Length; i++) if (d[pos + i] != what[i]) return false;
            return true;
        }

        static byte[] Slice(byte[] d, int pos, int len)
        {
            if (len <= 0) return new byte[0];
            var r = new byte[len];
            Buffer.BlockCopy(d, pos, r, 0, len);
            return r;
        }

        static byte[] Decode(byte[] d, int pos, int len, string cte, List<string> warnings)
        {
            switch (cte)
            {
                case "base64":
                    return MimeText.FromBase64(d, pos, len, warnings);
                case "quoted-printable":
                    return MimeText.FromQuotedPrintable(d, pos, len);
                case "x-uuencode": case "uuencode":
                    return MimeText.FromUuencode(d, pos, len, warnings);
                case null: case "": case "7bit": case "8bit": case "binary":
                    return Slice(d, pos, len);
                default:
                    warnings.Add("unknown transfer encoding '" + cte + "', content kept as is");
                    return Slice(d, pos, len);
            }
        }
    }

    // ---------------------------------------------------------------------------------------------------------------------
    //  Text helpers
    // ---------------------------------------------------------------------------------------------------------------------

    public static class MimeText
    {
        public static readonly Encoding Latin1252;
        static MimeText()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            Latin1252 = Encoding.GetEncoding(1252);
        }

        // ---- charsets ------------------------------------------------------------------------------------------------------

        public static Encoding GetEncoding(string charset)
        {
            if (string.IsNullOrWhiteSpace(charset)) return null;
            string c = charset.Trim().Trim('"').ToLowerInvariant();
            int star = c.IndexOf('*');
            if (star > 0) c = c.Substring(0, star);
            switch (c)
            {
                case "utf8": case "utf-8": case "unicode-1-1-utf-8": return new UTF8Encoding(false);
                case "us-ascii": case "ascii": case "ansi_x3.4-1968": case "us": return Encoding.GetEncoding(1252);   // Windows-1252 reads ASCII and tolerates 8-bit bytes
                case "iso-8859-8-i": c = "iso-8859-8"; break;
                case "iso-8859-8-e": c = "iso-8859-8"; break;
                case "x-unknown": case "unknown-8bit": case "default": case "unknown": return null;
                case "cp1252": c = "windows-1252"; break;
                case "ks_c_5601-1987": c = "euc-kr"; break;
                case "gb2312": c = "gb2312"; break;
                case "utf7": c = "utf-7"; break;
            }
            try { return Encoding.GetEncoding(c); }
            catch (Exception) { return null; }
        }

        /// <summary>Name of a charset for a Content-Type header (never null).</summary>
        public static string CharsetName(Encoding e)
        {
            return e.CodePage == 65001 ? "utf-8" : e.WebName;
        }

        public static string DecodeText(byte[] bytes, string charset)
        {
            if (bytes == null || bytes.Length == 0) return "";
            Encoding enc = GetEncoding(charset);
            if (enc != null)
            {
                try
                {
                    // keep the declared charset, but fall back when the bytes are clearly not in it (declared us-ascii, 8-bit UTF-8 inside)
                    if (enc.CodePage == 1252 && (charset ?? "").ToLowerInvariant().Contains("ascii") && HasHighBytes(bytes) && IsValidUtf8(bytes))
                        return new UTF8Encoding(false).GetString(bytes);
                    return StripBom(enc.GetString(bytes));
                }
                catch (Exception) { }
            }
            if (IsValidUtf8(bytes)) return StripBom(new UTF8Encoding(false).GetString(bytes));
            return Latin1252.GetString(bytes);
        }

        static string StripBom(string s) { return s.Length > 0 && s[0] == '﻿' ? s.Substring(1) : s; }

        static bool HasHighBytes(byte[] b) { foreach (byte x in b) if (x >= 0x80) return true; return false; }

        public static bool IsValidUtf8(byte[] bytes)
        {
            try { new UTF8Encoding(false, true).GetString(bytes); return true; }
            catch (Exception) { return false; }
        }

        // ---- transfer encodings --------------------------------------------------------------------------------------------

        public static byte[] FromBase64(byte[] d, int pos, int len, List<string> warnings)
        {
            var outp = new MemoryStream(len * 3 / 4 + 4);
            uint acc = 0; int bits = 0; bool invalid = false; int pad = 0;
            for (int i = pos; i < pos + len; i++)
            {
                byte c = d[i];
                int v;
                if (c >= 'A' && c <= 'Z') v = c - 'A';
                else if (c >= 'a' && c <= 'z') v = c - 'a' + 26;
                else if (c >= '0' && c <= '9') v = c - '0' + 52;
                else if (c == '+' || c == '-') v = 62;     // '-' : url-safe variants
                else if (c == '/' || c == '_') v = 63;
                else if (c == '=') { pad++; continue; }
                else { if (c != '\r' && c != '\n' && c != ' ' && c != '\t') invalid = true; continue; }
                if (pad > 0) { invalid = true; pad = 0; }   // data after padding: keep decoding (concatenated blocks)
                acc = (acc << 6) | (uint)v; bits += 6;
                if (bits >= 8) { bits -= 8; outp.WriteByte((byte)((acc >> bits) & 0xFF)); }
            }
            if (invalid && warnings != null) warnings.Add("base64 content with invalid characters, skipped");
            return outp.ToArray();
        }

        public static byte[] FromQuotedPrintable(byte[] d, int pos, int len)
        {
            var outp = new MemoryStream(len);
            int end = pos + len;
            for (int i = pos; i < end; i++)
            {
                byte c = d[i];
                if (c != '=') { outp.WriteByte(c); continue; }
                if (i + 1 < end && (d[i + 1] == '\r' || d[i + 1] == '\n'))
                {
                    // soft line break
                    if (d[i + 1] == '\r' && i + 2 < end && d[i + 2] == '\n') i += 2; else i += 1;
                    continue;
                }
                if (i + 2 < end && Hex(d[i + 1]) >= 0 && Hex(d[i + 2]) >= 0)
                {
                    outp.WriteByte((byte)(Hex(d[i + 1]) * 16 + Hex(d[i + 2])));
                    i += 2;
                    continue;
                }
                // '=' at the very end of the data, or not followed by two hex digits: kept
                if (i + 1 >= end) continue;   // trailing soft break without line break
                outp.WriteByte(c);
            }
            return outp.ToArray();
        }

        static int Hex(byte c)
        {
            if (c >= '0' && c <= '9') return c - '0';
            if (c >= 'A' && c <= 'F') return c - 'A' + 10;
            if (c >= 'a' && c <= 'f') return c - 'a' + 10;
            return -1;
        }

        public static byte[] FromUuencode(byte[] d, int pos, int len, List<string> warnings)
        {
            var outp = new MemoryStream();
            string text = Encoding.ASCII.GetString(d, pos, len);
            bool started = false;
            foreach (string raw in text.Split('\n'))
            {
                string line = raw.TrimEnd('\r');
                if (!started) { if (line.StartsWith("begin ")) started = true; continue; }
                if (line == "end") break;
                if (line.Length == 0) continue;
                int n = (line[0] - 32) & 0x3F;
                if (n == 0) continue;
                var buf = new List<byte>();
                for (int i = 1; i + 3 < line.Length + 3 && buf.Count < n; i += 4)
                {
                    int a = Uu(line, i), b = Uu(line, i + 1), c = Uu(line, i + 2), dd = Uu(line, i + 3);
                    buf.Add((byte)((a << 2) | (b >> 4))); buf.Add((byte)(((b & 15) << 4) | (c >> 2))); buf.Add((byte)(((c & 3) << 6) | dd));
                }
                for (int i = 0; i < n && i < buf.Count; i++) outp.WriteByte(buf[i]);
            }
            if (!started && warnings != null) warnings.Add("uuencode content without 'begin' line");
            return outp.ToArray();
        }

        static int Uu(string s, int i) { return i < s.Length ? (s[i] - 32) & 0x3F : 0; }

        // ---- encoded words (RFC 2047) --------------------------------------------------------------------------------------

        static readonly Regex EncodedWord = new Regex(@"=\?(?<cs>[^?\s]+)\?(?<enc>[bBqQ])\?(?<txt>[^?\s]*)\?=", RegexOptions.Compiled);

        /// <summary>Decodes the encoded words of a header value; white space between two adjacent encoded words is dropped (RFC 2047 6.2).</summary>
        public static string DecodeWords(string value)
        {
            if (string.IsNullOrEmpty(value) || value.IndexOf("=?", StringComparison.Ordinal) < 0) return value;
            var sb = new StringBuilder();
            int pos = 0;
            MatchCollection ms = EncodedWord.Matches(value);
            int i = 0;
            while (i < ms.Count)
            {
                Match m = ms[i];
                sb.Append(value, pos, m.Index - pos);
                // a run of adjacent encoded words with the same charset and encoding is decoded as one (multi-byte characters split between words)
                var bytes = new MemoryStream();
                string cs = m.Groups["cs"].Value;
                int lastEnd = m.Index + m.Length;
                int j = i;
                while (true)
                {
                    Match w = ms[j];
                    byte[] part = DecodeWordBytes(w);
                    bytes.Write(part, 0, part.Length);
                    lastEnd = w.Index + w.Length;
                    if (j + 1 < ms.Count && string.Equals(ms[j + 1].Groups["cs"].Value, cs, StringComparison.OrdinalIgnoreCase)
                        && IsBlank(value, lastEnd, ms[j + 1].Index)) j++;
                    else break;
                }
                // same charset runs only: different charsets are decoded separately but still lose the white space between them
                Encoding enc = GetEncoding(StripLanguage(cs)) ?? Encoding.UTF8;
                sb.Append(enc.GetString(bytes.ToArray()));
                pos = lastEnd;
                i = j + 1;
                if (i < ms.Count && IsBlank(value, pos, ms[i].Index)) pos = ms[i].Index;
            }
            sb.Append(value, pos, value.Length - pos);
            return sb.ToString();
        }

        static string StripLanguage(string cs) { int s = cs.IndexOf('*'); return s > 0 ? cs.Substring(0, s) : cs; }

        static bool IsBlank(string s, int from, int to)
        {
            for (int i = from; i < to; i++) if (s[i] != ' ' && s[i] != '\t' && s[i] != '\r' && s[i] != '\n') return false;
            return true;
        }

        static byte[] DecodeWordBytes(Match w)
        {
            string txt = w.Groups["txt"].Value;
            if (char.ToUpperInvariant(w.Groups["enc"].Value[0]) == 'B')
            {
                byte[] b = Encoding.ASCII.GetBytes(txt);
                return FromBase64(b, 0, b.Length, null);
            }
            byte[] q = Encoding.ASCII.GetBytes(txt.Replace('_', ' '));
            return FromQuotedPrintable(q, 0, q.Length);
        }

        /// <summary>
        /// A text for a header value: as is when it is printable ASCII, else encoded words (UTF-8; Q when most of the text is ASCII, else B),
        /// each at most 75 characters. The words are separated by a space; the folding of the header puts them on separate lines.
        /// </summary>
        public static string EncodeText(string text)
        {
            if (text == null) return "";
            bool ascii = true;
            foreach (char c in text) if (c < 32 || c > 126) { ascii = false; break; }
            if (ascii) return text;
            text = text.Replace("\r\n", " ").Replace('\r', ' ').Replace('\n', ' ');
            int nonAscii = text.Count(c => c > 126 || c < 32);
            bool useQ = nonAscii * 3 < text.Length;
            var words = new List<string>();
            var cur = new StringBuilder();
            int budget = useQ ? 60 : 42;   // bytes of the payload per word before encoding (B: 3 bytes -> 4 chars, 75 - 12 for the delimiters)
            var enc = new UTF8Encoding(false);
            var chunk = new List<byte>();
            int idx = 0;
            while (idx < text.Length)
            {
                int cl = char.IsHighSurrogate(text[idx]) && idx + 1 < text.Length ? 2 : 1;
                byte[] cb = enc.GetBytes(text.Substring(idx, cl));
                int size = useQ ? QSize(cb) : cb.Length;
                int current = useQ ? chunk.Sum(b => QSize(new[] { b })) : chunk.Count;
                if (current + size > budget && chunk.Count > 0) { words.Add(Word(chunk.ToArray(), useQ)); chunk.Clear(); }
                chunk.AddRange(cb);
                idx += cl;
            }
            if (chunk.Count > 0) words.Add(Word(chunk.ToArray(), useQ));
            return string.Join(" ", words);
        }

        static int QSize(byte[] bytes)
        {
            int n = 0;
            foreach (byte b in bytes) n += IsQSafe(b) ? 1 : 3;
            return n;
        }

        static bool IsQSafe(byte b) { return (b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z') || (b >= '0' && b <= '9') || b == '!' || b == '*' || b == '+' || b == '-' || b == '/'; }

        static string Word(byte[] bytes, bool q)
        {
            if (!q) return "=?UTF-8?B?" + Convert.ToBase64String(bytes) + "?=";
            var sb = new StringBuilder("=?UTF-8?Q?");
            foreach (byte b in bytes)
            {
                if (IsQSafe(b)) sb.Append((char)b);
                else if (b == ' ') sb.Append('_');
                else sb.Append('=').Append(b.ToString("X2"));
            }
            return sb.Append("?=").ToString();
        }

        // ---- parameters (RFC 2045 / 2231) ----------------------------------------------------------------------------------

        /// <summary>
        /// Splits "first; name=value; name2="quoted"; name3*0*=utf-8''%C3%A9; name3*1=..." into the first token and the parameters
        /// (names lower-cased, continuations joined, percent-encoding decoded with the charset given).
        /// </summary>
        public static Dictionary<string, string> ParseParameters(string value, out string first)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var raw = new List<string[]>();   // name, value, quoted
            int i = 0, n = value.Length;
            first = null;
            int semi = IndexOfUnquoted(value, ';', 0);
            first = (semi < 0 ? value : value.Substring(0, semi)).Trim();
            i = semi < 0 ? n : semi + 1;
            while (i < n)
            {
                while (i < n && (value[i] == ' ' || value[i] == '\t' || value[i] == ';' || value[i] == '\r' || value[i] == '\n')) i++;
                if (i >= n) break;
                int eq = i;
                while (eq < n && value[eq] != '=' && value[eq] != ';') eq++;
                if (eq >= n || value[eq] == ';') { i = eq; continue; }   // parameter without value
                string name = value.Substring(i, eq - i).Trim().ToLowerInvariant();
                i = eq + 1;
                while (i < n && (value[i] == ' ' || value[i] == '\t')) i++;
                string v; bool quoted = false;
                if (i < n && value[i] == '"')
                {
                    quoted = true;
                    var sb = new StringBuilder();
                    i++;
                    while (i < n && value[i] != '"')
                    {
                        if (value[i] == '\\' && i + 1 < n) i++;
                        sb.Append(value[i]); i++;
                    }
                    i++;
                    v = sb.ToString();
                }
                else
                {
                    int s = i;
                    while (i < n && value[i] != ';') i++;
                    v = value.Substring(s, i - s).Trim();
                }
                if (name.Length > 0) raw.Add(new[] { name, v, quoted ? "q" : "" });
            }
            // RFC 2231
            var groups = new Dictionary<string, SortedDictionary<int, KeyValuePair<string, bool>>>(StringComparer.OrdinalIgnoreCase);
            foreach (string[] r in raw)
            {
                string name = r[0];
                int star = name.IndexOf('*');
                if (star < 0) { if (!result.ContainsKey(name)) result[name] = r[1]; continue; }
                string baseName = name.Substring(0, star);
                string rest = name.Substring(star + 1);
                bool encoded = name.EndsWith("*");
                rest = rest.TrimEnd((char)42);
                int idx = 0;
                if (rest.Length > 0 && !int.TryParse(rest, NumberStyles.None, CultureInfo.InvariantCulture, out idx)) { result[name] = r[1]; continue; }
                SortedDictionary<int, KeyValuePair<string, bool>> g;
                if (!groups.TryGetValue(baseName, out g)) groups[baseName] = g = new SortedDictionary<int, KeyValuePair<string, bool>>();
                g[idx] = new KeyValuePair<string, bool>(r[1], encoded);
            }
            foreach (var kv in groups)
            {
                string charset = null;
                var bytes = new MemoryStream();
                var plain = new StringBuilder();
                bool first2 = true;
                bool anyEncoded = false;
                foreach (var part in kv.Value)
                {
                    string text = part.Value.Key;
                    if (part.Value.Value)
                    {
                        anyEncoded = true;
                        if (first2)
                        {
                            int q1 = text.IndexOf('\'');
                            int q2 = q1 < 0 ? -1 : text.IndexOf('\'', q1 + 1);
                            if (q1 >= 0 && q2 >= 0) { charset = text.Substring(0, q1); text = text.Substring(q2 + 1); }
                        }
                        byte[] pb = PercentDecode(text);
                        bytes.Write(pb, 0, pb.Length);
                    }
                    else
                    {
                        byte[] pb = Encoding.UTF8.GetBytes(text);
                        bytes.Write(pb, 0, pb.Length);
                    }
                    first2 = false;
                }
                string decoded = anyEncoded ? DecodeText(bytes.ToArray(), charset) : Encoding.UTF8.GetString(bytes.ToArray());
                if (!result.ContainsKey(kv.Key)) result[kv.Key] = decoded;
                else if (anyEncoded) result[kv.Key] = decoded;   // the extended form wins over a plain fallback
            }
            return result;
        }

        static int IndexOfUnquoted(string s, char c, int from)
        {
            bool q = false;
            for (int i = from; i < s.Length; i++)
            {
                if (s[i] == '"') q = !q;
                else if (s[i] == '\\' && q) i++;
                else if (s[i] == c && !q) return i;
            }
            return -1;
        }

        static byte[] PercentDecode(string s)
        {
            var ms = new MemoryStream();
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] == '%' && i + 2 < s.Length && Hex((byte)s[i + 1]) >= 0 && Hex((byte)s[i + 2]) >= 0)
                { ms.WriteByte((byte)(Hex((byte)s[i + 1]) * 16 + Hex((byte)s[i + 2]))); i += 2; }
                else { byte[] b = Encoding.UTF8.GetBytes(s[i].ToString()); ms.Write(b, 0, b.Length); }
            }
            return ms.ToArray();
        }

        /// <summary>name="value" for printable ASCII without specials, name*=UTF-8''percent-encoded otherwise.</summary>
        public static string FormatParameter(string name, string value)
        {
            bool plain = true;
            foreach (char c in value) if (c < 32 || c > 126) { plain = false; break; }
            if (plain) return name + "=\"" + value.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
            var sb = new StringBuilder(name + "*=UTF-8''");
            foreach (byte b in Encoding.UTF8.GetBytes(value))
            {
                if ((b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z') || (b >= '0' && b <= '9') || b == '.' || b == '-' || b == '_' || b == '~') sb.Append((char)b);
                else sb.Append('%').Append(b.ToString("X2"));
            }
            return sb.ToString();
        }

        // ---- addresses -----------------------------------------------------------------------------------------------------

        /// <summary>Mailboxes of an address header (groups are flattened, comments and unusable entries ignored). Names are decoded.</summary>
        public static List<MimeAddress> ParseAddressList(string value)
        {
            var list = new List<MimeAddress>();
            if (string.IsNullOrWhiteSpace(value)) return list;
            foreach (string item in SplitAddresses(value))
            {
                string s = item.Trim();
                if (s.Length == 0) continue;
                MimeAddress a = ParseMailbox(s);
                if (a != null) list.Add(a);
            }
            return list;
        }

        static IEnumerable<string> SplitAddresses(string v)
        {
            var sb = new StringBuilder();
            bool q = false; int angle = 0, paren = 0;
            for (int i = 0; i < v.Length; i++)
            {
                char c = v[i];
                if (q) { sb.Append(c); if (c == '\\' && i + 1 < v.Length) sb.Append(v[++i]); else if (c == '"') q = false; continue; }
                if (paren > 0) { sb.Append(c); if (c == '(') paren++; else if (c == ')') paren--; continue; }
                switch (c)
                {
                    case '"': q = true; sb.Append(c); break;
                    case '(': paren++; sb.Append(c); break;
                    case '<': angle++; sb.Append(c); break;
                    case '>': if (angle > 0) angle--; sb.Append(c); break;
                    case ':': if (angle == 0) { sb.Clear(); } else sb.Append(c); break;   // group name ends here: "Undisclosed recipients:;"
                    case ';': if (angle == 0) { if (sb.ToString().Trim().Length > 0) yield return sb.ToString(); sb.Clear(); } else sb.Append(c); break;
                    case ',': if (angle == 0) { yield return sb.ToString(); sb.Clear(); } else sb.Append(c); break;
                    default: sb.Append(c); break;
                }
            }
            if (sb.ToString().Trim().Length > 0) yield return sb.ToString();
        }

        static MimeAddress ParseMailbox(string s)
        {
            string name = null, addr = null;
            int lt = IndexOfUnquoted(s, '<', 0);
            if (lt >= 0)
            {
                int gt = s.IndexOf('>', lt + 1);
                addr = (gt < 0 ? s.Substring(lt + 1) : s.Substring(lt + 1, gt - lt - 1)).Trim();
                name = s.Substring(0, lt).Trim();
                // route address <@a,@b:user@host>
                int colon = addr.LastIndexOf(':');
                if (addr.StartsWith("@") && colon > 0) addr = addr.Substring(colon + 1);
            }
            else
            {
                // bare address, maybe with a (comment) name
                string noComment = StripComments(s, out name);
                addr = noComment.Trim();
            }
            if (string.IsNullOrEmpty(addr) || addr.IndexOf('@') < 0) { if (string.IsNullOrEmpty(addr) && string.IsNullOrEmpty(name)) return null; }
            if (name != null)
            {
                string unq = UnquoteName(name);
                name = DecodeWords(unq).Trim();
                if (name.Length == 0) name = null;
            }
            addr = addr == null ? null : addr.Replace("\"", "").Trim();
            if (string.IsNullOrEmpty(addr)) return name == null ? null : new MimeAddress { Name = name, Address = "" };
            return new MimeAddress { Name = name, Address = addr };
        }

        static string StripComments(string s, out string comment)
        {
            var sb = new StringBuilder(); var cm = new StringBuilder();
            int depth = 0; bool q = false;
            foreach (char c in s)
            {
                if (q) { sb.Append(c); if (c == '"') q = false; continue; }
                if (c == '"' && depth == 0) { q = true; sb.Append(c); continue; }
                if (c == '(') { if (depth > 0) cm.Append(c); depth++; continue; }
                if (c == ')' && depth > 0) { depth--; if (depth > 0) cm.Append(c); continue; }
                if (depth > 0) cm.Append(c); else sb.Append(c);
            }
            comment = cm.Length > 0 ? cm.ToString() : null;
            return sb.ToString();
        }

        static string UnquoteName(string n)
        {
            n = n.Trim();
            if (n.Length >= 2 && n[0] == '"' && n[n.Length - 1] == '"')
            {
                var sb = new StringBuilder();
                for (int i = 1; i < n.Length - 1; i++) { if (n[i] == '\\' && i + 1 < n.Length - 1) i++; sb.Append(n[i]); }
                return sb.ToString();
            }
            if (n.Contains("\"")) return n.Replace("\"", "");
            return n;
        }

        /// <summary>Display-name &lt;address&gt; for a header: the name is quoted or encoded when it needs to be.</summary>
        public static string FormatAddress(string name, string address)
        {
            if (string.IsNullOrEmpty(name) || name == address) return address ?? "";
            bool ascii = name.All(c => c >= 32 && c <= 126);
            string n;
            if (!ascii) n = EncodeText(name);
            else if (Regex.IsMatch(name, @"^[A-Za-z0-9!#$%&'*+/=?^_`{|}~ -]+$") && name.Trim() == name) n = name;
            else n = "\"" + name.Replace("\\", "\\\\").Replace("\"", "\\\"") + "\"";
            if (string.IsNullOrEmpty(address)) return n;
            return n + " <" + address + ">";
        }

        // ---- dates (RFC 5322 3.3, with the obsolete forms) ------------------------------------------------------------------

        static readonly string[] Months = { "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec" };
        static readonly Regex DateRx = new Regex(
            @"(?:(?:mon|tue|wed|thu|fri|sat|sun)[a-z]*\s*,?\s*)?(?<d>\d{1,2})[\s-]+(?<mon>[A-Za-z]{3,9})[\s-]+(?<y>\d{2,4})\s+(?<h>\d{1,2}):(?<mi>\d{2})(?::(?<s>\d{2}))?(?:\s*(?<z>[+-]\d{4}|[A-Za-z]{1,5}))?",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public static DateTimeOffset? ParseDate(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            string v = value;
            Match m = DateRx.Match(v);
            if (m.Success)
            {
                int mon = Array.IndexOf(Months, m.Groups["mon"].Value.Substring(0, 3).ToLowerInvariant());
                int d, y, h, mi, s = 0;
                if (mon >= 0 && int.TryParse(m.Groups["d"].Value, out d) && int.TryParse(m.Groups["y"].Value, out y)
                    && int.TryParse(m.Groups["h"].Value, out h) && int.TryParse(m.Groups["mi"].Value, out mi))
                {
                    if (m.Groups["s"].Success) int.TryParse(m.Groups["s"].Value, out s);
                    if (m.Groups["y"].Value.Length == 2) y += y < 50 ? 2000 : 1900;
                    else if (m.Groups["y"].Value.Length == 3) y += 1900;
                    TimeSpan off = ParseZone(m.Groups["z"].Success ? m.Groups["z"].Value : "");
                    try { return new DateTimeOffset(new DateTime(y, mon + 1, d, h, mi, Math.Min(s, 59), DateTimeKind.Unspecified), off); }
                    catch (Exception) { }
                }
            }
            DateTimeOffset dto;
            string cleaned = Regex.Replace(v, @"\([^)]*\)", "").Trim();
            if (DateTimeOffset.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.AssumeUniversal, out dto)) return dto;
            return null;
        }

        static TimeSpan ParseZone(string z)
        {
            if (string.IsNullOrEmpty(z)) return TimeSpan.Zero;
            if ((z[0] == '+' || z[0] == '-') && z.Length == 5)
            {
                int hh = int.Parse(z.Substring(1, 2)), mm = int.Parse(z.Substring(3, 2));
                if (hh > 14) return TimeSpan.Zero;
                var ts = new TimeSpan(hh, mm, 0);
                return z[0] == '-' ? -ts : ts;
            }
            switch (z.ToUpperInvariant())
            {
                case "EST": return TimeSpan.FromHours(-5);
                case "EDT": return TimeSpan.FromHours(-4);
                case "CST": return TimeSpan.FromHours(-6);
                case "CDT": return TimeSpan.FromHours(-5);
                case "MST": return TimeSpan.FromHours(-7);
                case "MDT": return TimeSpan.FromHours(-6);
                case "PST": return TimeSpan.FromHours(-8);
                case "PDT": return TimeSpan.FromHours(-7);
                default: return TimeSpan.Zero;   // GMT, UT, Z, military letters (RFC 5322 4.3: treated as +0000), unknown
            }
        }

        /// <summary>"Tue, 15 Mar 2024 08:31:07 +0100" with the offset of the value.</summary>
        public static string FormatDate(DateTimeOffset d)
        {
            string[] days = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };
            string[] mons = { "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec" };
            TimeSpan o = d.Offset;
            string sign = o < TimeSpan.Zero ? "-" : "+";
            o = o.Duration();
            return string.Format(CultureInfo.InvariantCulture, "{0}, {1} {2} {3:0000} {4:00}:{5:00}:{6:00} {7}{8:00}{9:00}",
                days[(int)d.DayOfWeek], d.Day, mons[d.Month - 1], d.Year, d.Hour, d.Minute, d.Second, sign, o.Hours, o.Minutes);
        }

        // ---- header folding -------------------------------------------------------------------------------------------------

        /// <summary>"Name: value" folded at 78 columns on white space, with CRLF; no trailing CRLF.</summary>
        public static string FoldHeader(string name, string value)
        {
            string line = name + ": " + (value ?? "");
            if (line.Length <= 78) return line;
            var sb = new StringBuilder();
            int lineStart = 0;
            int lastSpace = -1;
            for (int i = 0; i < line.Length; i++)
            {
                if (line[i] == ' ' && i > name.Length + 1) lastSpace = i;
                if (i - lineStart >= 77 && lastSpace > lineStart)
                {
                    sb.Append(line, lineStart, lastSpace - lineStart).Append("\r\n");
                    lineStart = lastSpace;   // the space starts the continuation line
                    lastSpace = -1;
                }
            }
            sb.Append(line, lineStart, line.Length - lineStart);
            return sb.ToString();
        }

        // ---- encoders ------------------------------------------------------------------------------------------------------

        public static void WriteBase64(Stream source, Stream dest, CancellationToken ct)
        {
            var buf = new byte[57 * 64];
            var line = new byte[78];
            int n;
            while ((n = FxIo.ReadFull(source, buf, 0, buf.Length)) > 0)
            {
                ct.ThrowIfCancellationRequested();
                for (int off = 0; off < n; off += 57)
                {
                    int take = Math.Min(57, n - off);
                    string b64 = Convert.ToBase64String(buf, off, take);
                    for (int i = 0; i < b64.Length; i++) line[i] = (byte)b64[i];
                    line[b64.Length] = (byte)'\r'; line[b64.Length + 1] = (byte)'\n';
                    dest.Write(line, 0, b64.Length + 2);
                }
                if (n < buf.Length) break;
            }
        }

        /// <summary>Quoted-printable of text bytes: lines of at most 76 characters, line breaks (CRLF, LF or CR) written as CRLF.</summary>
        public static void WriteQuotedPrintable(byte[] text, Stream dest)
        {
            var outp = new StringBuilder();
            int i = 0;
            int lineLen = 0;
            Action<string> put = piece =>
            {
                if (lineLen + piece.Length > 75) { outp.Append("=\r\n"); lineLen = 0; }
                outp.Append(piece); lineLen += piece.Length;
            };
            while (i < text.Length)
            {
                byte b = text[i];
                if (b == '\r' || b == '\n')
                {
                    if (b == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    outp.Append("\r\n"); lineLen = 0; i++;
                    continue;
                }
                bool nextIsBreak = i + 1 >= text.Length || text[i + 1] == '\r' || text[i + 1] == '\n';
                if (b == ' ' || b == '\t')
                {
                    if (nextIsBreak) put("=" + b.ToString("X2")); else put(((char)b).ToString());
                }
                else if ((b >= 33 && b <= 126 && b != '=') ) put(((char)b).ToString());
                else put("=" + b.ToString("X2"));
                i++;
            }
            byte[] bytes = Encoding.ASCII.GetBytes(outp.ToString());
            dest.Write(bytes, 0, bytes.Length);
        }

        public static byte[] NormalizeToCrlf(byte[] text)
        {
            bool needs = false;
            for (int i = 0; i < text.Length; i++)
            {
                if (text[i] == '\n' && (i == 0 || text[i - 1] != '\r')) { needs = true; break; }
                if (text[i] == '\r' && (i + 1 >= text.Length || text[i + 1] != '\n')) { needs = true; break; }
            }
            if (!needs) return text;
            var ms = new MemoryStream(text.Length + 64);
            for (int i = 0; i < text.Length; i++)
            {
                byte c = text[i];
                if (c == '\r') { if (i + 1 < text.Length && text[i + 1] == '\n') i++; ms.WriteByte((byte)'\r'); ms.WriteByte((byte)'\n'); }
                else if (c == '\n') { ms.WriteByte((byte)'\r'); ms.WriteByte((byte)'\n'); }
                else ms.WriteByte(c);
            }
            return ms.ToArray();
        }
    }

    // ---------------------------------------------------------------------------------------------------------------------
    //  Writer
    // ---------------------------------------------------------------------------------------------------------------------

    public sealed class MimeWriteOptions
    {
        public CancellationToken Cancellation;
    }

    public static class MimeWriter
    {
        static readonly byte[] Crlf = { 13, 10 };

        /// <summary>
        /// Writes the entity as it is described: the headers in their order (already encoded), then the body. Multipart: a boundary is taken
        /// from Content-Type (parameter boundary) or generated and added; leaf: Content written with TransferEncoding (base64,
        /// quoted-printable, 7bit, 8bit, binary), text parts are normalised to CRLF; message/rfc822: the Message written inside.
        /// </summary>
        public static void Write(MimeEntity e, Stream dest, MimeWriteOptions options = null)
        {
            Write(e, dest, options ?? new MimeWriteOptions(), 0);
        }

        public static byte[] ToBytes(MimeEntity e)
        {
            using (var ms = new MemoryStream()) { Write(e, ms); return ms.ToArray(); }
        }

        static void Write(MimeEntity e, Stream dest, MimeWriteOptions o, int depth)
        {
            if (depth > 64) throw new InvalidDataException("MIME nesting deeper than 64 levels.");
            o.Cancellation.ThrowIfCancellationRequested();
            string boundary = null;
            if (e.IsMultipart)
            {
                boundary = e.ContentType["boundary"];
                if (boundary == null) boundary = "----=_MBT_" + Guid.NewGuid().ToString("N");
            }
            foreach (MimeHeader h in e.Headers)
            {
                if (h.RawLine != null && !(boundary != null && string.Equals(h.Name, "Content-Type", StringComparison.OrdinalIgnoreCase)))
                {
                    // a header taken from a file is written back as it was (folding included)
                    WriteAscii(dest, h.RawLine.Replace("\r\n", "\n").Replace("\n", "\r\n"));
                    dest.Write(Crlf, 0, 2);
                    continue;
                }
                string value = h.Value;
                if (boundary != null && string.Equals(h.Name, "Content-Type", StringComparison.OrdinalIgnoreCase) && e.ContentType["boundary"] == null)
                    value = value + "; boundary=\"" + boundary + "\"";
                WriteAscii(dest, MimeText.FoldHeader(h.Name, value));
                dest.Write(Crlf, 0, 2);
            }
            dest.Write(Crlf, 0, 2);

            if (e.IsMultipart)
            {
                foreach (MimeEntity part in e.Parts)
                {
                    WriteAscii(dest, "--" + boundary + "\r\n");
                    Write(part, dest, o, depth + 1);
                    dest.Write(Crlf, 0, 2);
                }
                WriteAscii(dest, "--" + boundary + "--\r\n");
                return;
            }
            if (e.Message != null) { Write(e.Message, dest, o, depth + 1); return; }
            if (e.Content == null || e.Content.Length == 0) return;
            string cte = e.TransferEncoding;
            bool isText = e.ContentType.Type == "text";
            switch (cte)
            {
                case "base64":
                    using (Stream s = e.Content.OpenRead()) MimeText.WriteBase64(s, dest, o.Cancellation);
                    break;
                case "quoted-printable":
                    MimeText.WriteQuotedPrintable(e.Content.ToArray(), dest);
                    dest.Write(Crlf, 0, 2);
                    break;
                default:
                    byte[] b = e.Content.ToArray();
                    if (isText) b = MimeText.NormalizeToCrlf(b);
                    dest.Write(b, 0, b.Length);
                    if (b.Length < 2 || b[b.Length - 2] != '\r' || b[b.Length - 1] != '\n') dest.Write(Crlf, 0, 2);
                    break;
            }
        }

        static void WriteAscii(Stream dest, string s)
        {
            byte[] b = Encoding.UTF8.GetBytes(s);
            dest.Write(b, 0, b.Length);
        }
    }
}
