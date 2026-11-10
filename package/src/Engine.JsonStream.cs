// =============================================================================
//  Mailbox Transfer - engine, part 2: streaming of the exportItems response
// -----------------------------------------------------------------------------
//  The answer of exportItems is JSON: {"value":[{"itemId","changeKey","data":"<base64>"}, ...]}.
//  "data" is the whole item, so up to tens of megabytes per item, 20 items per request. It is never held
//  in memory: this small reader walks the bytes of the HTTP response, decodes the base64 text while it
//  arrives and writes it to a file. JSON escapes inside the text ("\/", "+") are handled.
//  The other values are small and read as ordinary strings.
// =============================================================================
using System;
using System.Buffers.Text;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace MailboxTransfer
{
    /// <summary>One element of the exportItems answer: a file on disk, or an error for this item.</summary>
    public sealed class ExportEntry
    {
        public int Index;
        public string ItemId;
        public string ChangeKey;
        public string FilePath;
        public string ExportUrl;
        public long Bytes;
        public string Sha256;
        public string ErrorCode;
        public string ErrorMessage;
        public bool HasError { get { return ErrorCode != null || ErrorMessage != null; } }
    }

    /// <summary>Decodes base64 given in pieces (any size) and writes the bytes to a stream.</summary>
    public sealed class Base64Writer
    {
        readonly Stream _dest;
        readonly IncrementalHash _hash;
        readonly byte[] _carry = new byte[4];
        int _carryLen;
        byte[] _out = new byte[48 * 1024];
        public long Bytes { get; private set; }

        public Base64Writer(Stream dest, IncrementalHash hash = null) { _dest = dest; _hash = hash; }

        public void Feed(ReadOnlySpan<byte> ascii)
        {
            if (_carryLen > 0)
            {
                while (_carryLen < 4 && !ascii.IsEmpty) { _carry[_carryLen++] = ascii[0]; ascii = ascii.Slice(1); }
                if (_carryLen < 4) return;
                Decode(new ReadOnlySpan<byte>(_carry, 0, 4));
                _carryLen = 0;
            }
            int bulk = (ascii.Length / 4) * 4;
            if (bulk > 0) Decode(ascii.Slice(0, bulk));
            for (int i = bulk; i < ascii.Length; i++) _carry[_carryLen++] = ascii[i];
        }

        void Decode(ReadOnlySpan<byte> quads)
        {
            int need = quads.Length / 4 * 3;
            if (_out.Length < need) _out = new byte[need];
            int consumed, written;
            if (Base64.DecodeFromUtf8(quads, _out, out consumed, out written, true) != System.Buffers.OperationStatus.Done)
                throw new InvalidDataException("Invalid base64 in the exportItems response.");
            _dest.Write(_out, 0, written);
            if (_hash != null) _hash.AppendData(_out, 0, written);
            Bytes += written;
        }

        public void Finish()
        {
            if (_carryLen != 0) throw new InvalidDataException("Truncated base64 in the exportItems response.");
        }
    }

    /// <summary>Forward-only JSON byte reader, just enough for the exportItems answer.</summary>
    sealed class JsonStreamReader
    {
        readonly Stream _s;
        readonly byte[] _buf = new byte[64 * 1024];
        int _pos, _len;

        public JsonStreamReader(Stream s) { _s = s; }

        bool Fill()
        {
            _pos = 0;
            _len = _s.Read(_buf, 0, _buf.Length);
            return _len > 0;
        }

        public int Peek()
        {
            if (_pos == _len && !Fill()) return -1;
            return _buf[_pos];
        }

        int Next()
        {
            int c = Peek();
            if (c >= 0) _pos++;
            return c;
        }

        byte NextRequired()
        {
            int c = Next();
            if (c < 0) throw new InvalidDataException("The exportItems response ends in the middle of a JSON value.");
            return (byte)c;
        }

        public void SkipWs()
        {
            int c;
            while ((c = Peek()) == ' ' || c == '\t' || c == '\r' || c == '\n') _pos++;
        }

        public void Expect(char ch)
        {
            SkipWs();
            byte c = NextRequired();
            if (c != ch) throw new InvalidDataException("JSON: '" + ch + "' expected, '" + (char)c + "' found.");
        }

        /// <summary>Escape sequence after the backslash, as UTF-8 bytes.</summary>
        void ReadEscape(List<byte> into)
        {
            byte e = NextRequired();
            switch ((char)e)
            {
                case '"': into.Add((byte)'"'); return;
                case '\\': into.Add((byte)'\\'); return;
                case '/': into.Add((byte)'/'); return;
                case 'b': into.Add(8); return;
                case 'f': into.Add(12); return;
                case 'n': into.Add(10); return;
                case 'r': into.Add(13); return;
                case 't': into.Add(9); return;
                case 'u':
                    int cp = ReadHex4();
                    if (cp >= 0xD800 && cp <= 0xDBFF && Peek() == '\\')
                    {
                        _pos++;
                        if (NextRequired() != 'u') throw new InvalidDataException("JSON: bad escape after a high surrogate.");
                        int low = ReadHex4();
                        cp = (low >= 0xDC00 && low <= 0xDFFF) ? 0x10000 + ((cp - 0xD800) << 10) + (low - 0xDC00) : 0xFFFD;
                    }
                    else if (cp >= 0xD800 && cp <= 0xDFFF) cp = 0xFFFD;
                    into.AddRange(Encoding.UTF8.GetBytes(char.ConvertFromUtf32(cp)));
                    return;
                default: throw new InvalidDataException("JSON: unknown escape \\" + (char)e);
            }
        }

        int ReadHex4()
        {
            int v = 0;
            for (int i = 0; i < 4; i++)
            {
                byte c = NextRequired();
                int d = c >= '0' && c <= '9' ? c - '0' : c >= 'a' && c <= 'f' ? c - 'a' + 10 : c >= 'A' && c <= 'F' ? c - 'A' + 10 : -1;
                if (d < 0) throw new InvalidDataException("JSON: bad \\u escape.");
                v = v * 16 + d;
            }
            return v;
        }

        public string ReadString()
        {
            Expect('"');
            var bytes = new List<byte>();
            while (true)
            {
                byte c = NextRequired();
                if (c == '"') break;
                if (c == '\\') ReadEscape(bytes); else bytes.Add(c);
            }
            return Encoding.UTF8.GetString(bytes.ToArray());
        }

        /// <summary>Reads a JSON string made of base64 text and hands it to the writer, piece by piece.</summary>
        public void ReadBase64(Base64Writer writer)
        {
            Expect('"');
            var esc = new List<byte>(4);
            while (true)
            {
                if (_pos == _len && !Fill()) throw new InvalidDataException("The exportItems response ends in the middle of the data.");
                int i = _pos;
                while (i < _len && _buf[i] != '"' && _buf[i] != '\\') i++;
                if (i > _pos) writer.Feed(new ReadOnlySpan<byte>(_buf, _pos, i - _pos));
                _pos = i;
                if (_pos == _len) continue;
                byte c = _buf[_pos++];
                if (c == '"') return;
                esc.Clear();
                ReadEscape(esc);
                if (esc.Count != 1 || esc[0] >= 0x80) throw new InvalidDataException("Invalid escape in a base64 string.");
                writer.Feed(esc.ToArray());
            }
        }

        public void SkipValue()
        {
            SkipWs();
            int c = Peek();
            if (c == '"') { SkipString(); return; }
            if (c == '{' || c == '[')
            {
                int depth = 0;
                do
                {
                    SkipWs();
                    c = Peek();
                    if (c < 0) throw new InvalidDataException("JSON ends unexpectedly.");
                    if (c == '"') { SkipString(); continue; }
                    _pos++;
                    if (c == '{' || c == '[') depth++;
                    else if (c == '}' || c == ']') depth--;
                } while (depth > 0);
                return;
            }
            while ((c = Peek()) >= 0 && c != ',' && c != '}' && c != ']' && c != ' ' && c != '\r' && c != '\n' && c != '\t') _pos++;
        }

        void SkipString()
        {
            Expect('"');
            while (true)
            {
                byte c = NextRequired();
                if (c == '"') return;
                if (c == '\\') NextRequired();
            }
        }
    }

    public static class ExportResponseReader
    {
        /// <summary>
        /// Reads an exportItems answer. Each item with data becomes a file "NNN-xxxx.fxs" in <paramref name="directory"/>
        /// (NNN = indexBase + position in the answer, xxxx = end of the item id). The decoded data is written while
        /// the response is read. On any error the unfinished file is removed and the exception is thrown;
        /// the files of the entries already complete stay.
        /// </summary>
        public static List<ExportEntry> Read(Stream body, string directory, int indexBase = 0)
        {
            Directory.CreateDirectory(directory);
            var entries = new List<ExportEntry>();
            var r = new JsonStreamReader(body);
            r.Expect('{');
            while (true)
            {
                r.SkipWs();
                if (r.Peek() == '}') break;
                string name = r.ReadString();
                r.Expect(':');
                if (string.Equals(name, "value", StringComparison.OrdinalIgnoreCase)) ReadEntries(r, directory, indexBase, entries);
                else r.SkipValue();
                r.SkipWs();
                if (r.Peek() == ',') r.Expect(',');
            }
            return entries;
        }

        static void ReadEntries(JsonStreamReader r, string directory, int indexBase, List<ExportEntry> entries)
        {
            r.Expect('[');
            while (true)
            {
                r.SkipWs();
                if (r.Peek() == ']') { r.Expect(']'); return; }
                entries.Add(ReadEntry(r, directory, indexBase + entries.Count));
                r.SkipWs();
                if (r.Peek() == ',') r.Expect(',');
            }
        }

        static ExportEntry ReadEntry(JsonStreamReader r, string directory, int index)
        {
            var entry = new ExportEntry { Index = index };
            string part = null;
            try
            {
                r.Expect('{');
                while (true)
                {
                    r.SkipWs();
                    if (r.Peek() == '}') { r.Expect('}'); break; }
                    string name = r.ReadString().ToLowerInvariant();
                    r.Expect(':');
                    r.SkipWs();
                    if (name == "itemid") entry.ItemId = r.ReadString();
                    else if (name == "changekey") entry.ChangeKey = r.ReadString();
                    else if (name == "data" && r.Peek() == '"')
                    {
                        part = Path.Combine(directory, "export-" + Guid.NewGuid().ToString("N") + ".part");
                        using (var file = new FileStream(part, FileMode.Create, FileAccess.Write, FileShare.None, 81920))
                        using (var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256))
                        {
                            var writer = new Base64Writer(file, hash);
                            r.ReadBase64(writer);
                            writer.Finish();
                            entry.Bytes = writer.Bytes;
                            entry.Sha256 = Convert.ToHexString(hash.GetHashAndReset()).ToLowerInvariant();
                        }
                    }
                    else if (name == "error" && r.Peek() == '{') ReadError(r, entry);
                    else r.SkipValue();
                    r.SkipWs();
                    if (r.Peek() == ',') r.Expect(',');
                }
                if (part != null)
                {
                    string final = Path.Combine(directory, SafeName(index, entry.ItemId));
                    File.Move(part, final, true);
                    entry.FilePath = final;
                    part = null;
                }
                return entry;
            }
            finally
            {
                if (part != null) try { File.Delete(part); } catch (IOException) { }
            }
        }

        static void ReadError(JsonStreamReader r, ExportEntry entry)
        {
            r.Expect('{');
            while (true)
            {
                r.SkipWs();
                if (r.Peek() == '}') { r.Expect('}'); break; }
                string name = r.ReadString().ToLowerInvariant();
                r.Expect(':');
                r.SkipWs();
                string value = r.Peek() == '"' ? r.ReadString() : null;
                if (value == null) r.SkipValue();
                if (name == "code") entry.ErrorCode = value ?? "Unknown";
                else if (name == "message") entry.ErrorMessage = value;
                r.SkipWs();
                if (r.Peek() == ',') r.Expect(',');
            }
            if (entry.ErrorCode == null) entry.ErrorCode = "Unknown";
        }

        /// <summary>"003-" followed by the last 24 letters, digits, '-' or '_' of the item id, then ".fxs".</summary>
        public static string SafeName(int index, string itemId)
        {
            var sb = new StringBuilder();
            foreach (char c in itemId ?? "")
                if ((c < 128 && char.IsLetterOrDigit(c)) || c == '-' || c == '_') sb.Append(c);
            string tail = sb.Length > 24 ? sb.ToString(sb.Length - 24, 24) : sb.ToString();
            return index.ToString("D3") + "-" + (tail.Length == 0 ? "item" : tail) + ".fxs";
        }
    }
}
