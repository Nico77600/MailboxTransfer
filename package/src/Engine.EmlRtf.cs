// =============================================================================
//  Mailbox Transfer - engine, part 13: compressed RTF (PR_RTF_COMPRESSED, MS-OXRTFCP) and its two uses for a MIME body
// -----------------------------------------------------------------------------
//  Rtf.Decompress     LZFu / MELA container -> RTF bytes
//  Rtf.ExtractHtml    RTF that encapsulates HTML (\fromhtml1, MS-OXRTFEX) -> the HTML, null when the RTF does not carry any
//  Rtf.ToText         any RTF -> plain text (destinations skipped, \par, \tab, \'hh and \uN handled)
//  A message that only has an RTF body (Outlook "Rich Text" format) is turned into HTML or text with these when it is written as MIME.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MailboxTransfer
{
    public static class Rtf
    {
        const string Prebuf =
            "{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript \\fdecor MS Sans SerifSymbolArialTimes New RomanCourier{\\colortbl\\red0\\green0\\blue0\r\n\\par \\pard\\plain\\f0\\fs20\\b\\i\\u\\tx";

        /// <summary>The RTF of a PR_RTF_COMPRESSED value, null when the container is not valid.</summary>
        public static byte[] Decompress(byte[] src)
        {
            if (src == null || src.Length < 16) return null;
            uint compSize = BitConverter.ToUInt32(src, 0);
            uint rawSize = BitConverter.ToUInt32(src, 4);
            uint magic = BitConverter.ToUInt32(src, 8);
            if (rawSize > 512u * 1024 * 1024) return null;
            if (magic == 0x414C454D)   // MELA: not compressed
            {
                int n = (int)Math.Min(rawSize, (uint)(src.Length - 16));
                var r = new byte[n];
                Buffer.BlockCopy(src, 16, r, 0, n);
                return r;
            }
            if (magic != 0x75465A4C) return null;   // LZFu
            var dict = new byte[4096];
            byte[] pre = Encoding.ASCII.GetBytes(Prebuf);
            Buffer.BlockCopy(pre, 0, dict, 0, pre.Length);
            int wr = pre.Length;
            var outp = new MemoryStream((int)Math.Min(rawSize + 16, int.MaxValue));
            int pos = 16;
            int limit = (int)Math.Min((long)src.Length, (long)compSize + 4);
            while (pos < limit)
            {
                int ctl = src[pos++];
                for (int bit = 0; bit < 8 && pos < limit; bit++, ctl >>= 1)
                {
                    if ((ctl & 1) == 0)
                    {
                        byte b = src[pos++];
                        outp.WriteByte(b); dict[wr] = b; wr = (wr + 1) & 4095;
                    }
                    else
                    {
                        if (pos + 1 >= src.Length) return outp.ToArray();
                        int hi = src[pos++], lo = src[pos++];
                        int off = (hi << 4) | (lo >> 4);
                        int len = (lo & 15) + 2;
                        if (off == wr) return outp.ToArray();   // end marker
                        for (int k = 0; k < len; k++)
                        {
                            byte b = dict[(off + k) & 4095];
                            outp.WriteByte(b); dict[wr] = b; wr = (wr + 1) & 4095;
                        }
                    }
                }
            }
            return outp.ToArray();
        }

        static bool IsLetter(byte b) { return (b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z'); }

        sealed class State { public bool HtmlRtf; public bool Skip; public bool InTag; public int Uc = 1; }

        /// <summary>The HTML an RTF encapsulates (\fromhtml), null when there is none.</summary>
        public static byte[] ExtractHtml(byte[] rtf)
        {
            if (rtf == null || rtf.Length == 0) return null;
            string head = Encoding.ASCII.GetString(rtf, 0, Math.Min(rtf.Length, 2000));
            if (head.IndexOf("\\fromhtml", StringComparison.Ordinal) < 0) return null;
            int cp = AnsiCodePage(head);
            Encoding enc = Encoding.GetEncoding(cp);
            var o = new MemoryStream();
            var stack = new Stack<State>();
            var st = new State();
            int i = 0, n = rtf.Length;
            while (i < n)
            {
                byte c = rtf[i];
                if (c == '{') { stack.Push(st); st = new State { HtmlRtf = st.HtmlRtf, Skip = st.Skip, InTag = st.InTag, Uc = st.Uc }; i++; continue; }
                if (c == '}') { if (stack.Count > 0) st = stack.Pop(); i++; continue; }
                if (c == '\r' || c == '\n') { i++; continue; }
                if (c == '\\')
                {
                    i++;
                    if (i >= n) break;
                    byte d = rtf[i];
                    if (d == '\\' || d == '{' || d == '}') { if (!st.Skip && (!st.HtmlRtf || st.InTag)) o.WriteByte(d); i++; continue; }
                    if (d == '*') { i++; st.Skip = st.Skip || false; MarkDestination(rtf, ref i, ref st); continue; }
                    if (d == '\'')
                    {
                        if (i + 2 < n)
                        {
                            int v = Convert.ToInt32(Encoding.ASCII.GetString(rtf, i + 1, 2), 16);
                            i += 3;
                            if (!st.Skip && (!st.HtmlRtf || st.InTag)) { byte[] u = Encoding.UTF8.GetBytes(enc.GetString(new[] { (byte)v })); o.Write(u, 0, u.Length); }
                        }
                        else i = n;
                        continue;
                    }
                    if (IsLetter(d))
                    {
                        int s = i;
                        while (i < n && IsLetter(rtf[i])) i++;
                        string word = Encoding.ASCII.GetString(rtf, s, i - s);
                        int sign = 1, val = int.MinValue;
                        if (i < n && rtf[i] == '-') { sign = -1; i++; }
                        int ds = i;
                        while (i < n && rtf[i] >= '0' && rtf[i] <= '9') i++;
                        if (i > ds) val = sign * int.Parse(Encoding.ASCII.GetString(rtf, ds, i - ds));
                        if (i < n && rtf[i] == ' ') i++;
                        switch (word)
                        {
                            case "htmlrtf": st.HtmlRtf = val != 0; break;
                            case "htmltag": st.InTag = true; st.HtmlRtf = false; break;
                            case "mhtmltag": st.Skip = true; break;
                            case "par": if (!st.Skip && !st.HtmlRtf) o.WriteByte((byte)'\n'); break;
                            case "tab": if (!st.Skip && !st.HtmlRtf) o.WriteByte((byte)'\t'); break;
                            case "uc": if (val >= 0) st.Uc = val; break;
                            case "u":
                                if (val != int.MinValue)
                                {
                                    int cpnt = val < 0 ? val + 65536 : val;
                                    if (!st.Skip && (!st.HtmlRtf || st.InTag)) { byte[] u = Encoding.UTF8.GetBytes(((char)cpnt).ToString()); o.Write(u, 0, u.Length); }
                                    // skip the Uc fallback characters
                                    for (int k = 0; k < st.Uc && i < n; k++)
                                    {
                                        if (rtf[i] == '\\' && i + 3 < n && rtf[i + 1] == '\'') i += 4; else if (rtf[i] != '\\' && rtf[i] != '{' && rtf[i] != '}') i++;
                                    }
                                }
                                break;
                            case "fonttbl": case "colortbl": case "stylesheet": case "info": case "pict": st.Skip = true; break;
                        }
                        continue;
                    }
                    i++;   // \~ \- \_ and other symbols: ignored
                    continue;
                }
                if (!st.Skip && (!st.HtmlRtf || st.InTag)) o.WriteByte(c);
                i++;
            }
            return o.Length == 0 ? null : o.ToArray();
        }

        static void MarkDestination(byte[] rtf, ref int i, ref State st)
        {
            // after "\*": "\htmltagN" keeps its content (it is HTML), any other destination is skipped
            if (i < rtf.Length && rtf[i] == '\\')
            {
                int s = i + 1, e = s;
                while (e < rtf.Length && IsLetter(rtf[e])) e++;
                string word = Encoding.ASCII.GetString(rtf, s, e - s);
                if (word == "htmltag") { st.InTag = true; st.HtmlRtf = false; i = e; while (i < rtf.Length && rtf[i] >= '0' && rtf[i] <= '9') i++; if (i < rtf.Length && rtf[i] == ' ') i++; return; }
            }
            st.Skip = true;
        }

        static int AnsiCodePage(string head)
        {
            int k = head.IndexOf("\\ansicpg", StringComparison.Ordinal);
            if (k >= 0)
            {
                int s = k + 8, e = s;
                while (e < head.Length && char.IsDigit(head[e])) e++;
                int cp;
                if (e > s && int.TryParse(head.Substring(s, e - s), out cp))
                {
                    try { Encoding.GetEncoding(cp); return cp; } catch (Exception) { }
                }
            }
            return 1252;
        }

        /// <summary>Plain text of an RTF.</summary>
        public static string ToText(byte[] rtf)
        {
            if (rtf == null || rtf.Length == 0) return "";
            string head = Encoding.ASCII.GetString(rtf, 0, Math.Min(rtf.Length, 2000));
            Encoding enc = Encoding.GetEncoding(AnsiCodePage(head));
            var sb = new StringBuilder();
            var stack = new Stack<bool>();
            bool skip = false;
            int uc = 1;
            int i = 0, n = rtf.Length;
            while (i < n)
            {
                byte c = rtf[i];
                if (c == '{') { stack.Push(skip); i++; continue; }
                if (c == '}') { if (stack.Count > 0) skip = stack.Pop(); i++; continue; }
                if (c == '\r' || c == '\n') { i++; continue; }
                if (c == '\\')
                {
                    i++;
                    if (i >= n) break;
                    byte d = rtf[i];
                    if (d == '\\' || d == '{' || d == '}') { if (!skip) sb.Append((char)d); i++; continue; }
                    if (d == '*') { skip = true; i++; continue; }
                    if (d == '\'')
                    {
                        if (i + 2 < n) { int v = Convert.ToInt32(Encoding.ASCII.GetString(rtf, i + 1, 2), 16); i += 3; if (!skip) sb.Append(enc.GetString(new[] { (byte)v })); }
                        else i = n;
                        continue;
                    }
                    if (IsLetter(d))
                    {
                        int s = i;
                        while (i < n && IsLetter(rtf[i])) i++;
                        string word = Encoding.ASCII.GetString(rtf, s, i - s);
                        int sign = 1, val = int.MinValue;
                        if (i < n && rtf[i] == '-') { sign = -1; i++; }
                        int ds = i;
                        while (i < n && rtf[i] >= '0' && rtf[i] <= '9') i++;
                        if (i > ds) val = sign * int.Parse(Encoding.ASCII.GetString(rtf, ds, i - ds));
                        if (i < n && rtf[i] == ' ') i++;
                        switch (word)
                        {
                            case "par": case "line": if (!skip) sb.Append('\n'); break;
                            case "tab": if (!skip) sb.Append('\t'); break;
                            case "uc": if (val >= 0) uc = val; break;
                            case "u":
                                if (val != int.MinValue)
                                {
                                    if (!skip) sb.Append((char)(val < 0 ? val + 65536 : val));
                                    for (int k = 0; k < uc && i < n; k++)
                                    {
                                        if (rtf[i] == '\\' && i + 3 < n && rtf[i + 1] == '\'') i += 4; else if (rtf[i] != '\\' && rtf[i] != '{' && rtf[i] != '}') i++;
                                    }
                                }
                                break;
                            case "fonttbl": case "colortbl": case "stylesheet": case "info": case "pict": case "header": case "footer": skip = true; break;
                        }
                        continue;
                    }
                    i++;
                    continue;
                }
                if (!skip) sb.Append(enc.GetString(new[] { c }));
                i++;
            }
            return sb.ToString().Trim('\n', ' ');
        }
    }

}
