// =============================================================================
//  Mailbox Transfer - test fixtures (compiled by tests\MailboxTransfer.Tests.ps1)
// -----------------------------------------------------------------------------
//  FakeHandler : an HttpMessageHandler that answers from a queue and records what was sent, so the Graph client
//                is tested without any network.
//  FxBuilder   : writes FastTransfer streams from the layout of MS-OXCFXICS 2.2.4 (properties of every kind,
//                markers, named properties, multi-valued, code page strings), so the reader is tested on streams
//                built independently of the reader.
// =============================================================================
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace MailboxTransfer.Tests
{
    /// <summary>A token provider that works from any thread (a PowerShell script block does not).</summary>
    public static class FixedToken
    {
        public static readonly Func<string> Provider = () => "test-token";
    }

    public sealed class RecordedRequest
    {
        public string Method;
        public string Url;
        public string Body;
        public byte[] BodyBytes;
        public bool HasAuthorization;
        public long? ContentLength;
        public string ContentType;
    }

    /// <summary>A write-only stream that accepts a number of bytes and then fails like a connection closed by the server.</summary>
    public sealed class DroppingStream : Stream
    {
        readonly long _limit; long _accepted;
        public DroppingStream(long limit) { _limit = limit; }
        public override bool CanRead { get { return false; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return true; } }
        public override long Length { get { throw new NotSupportedException(); } }
        public override long Position { get { throw new NotSupportedException(); } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count)
        {
            if (_accepted + count > _limit) throw new IOException("Unable to write data to the transport connection: an existing connection was forcibly closed by the remote host.");
            _accepted += count;
        }
    }

    public sealed class FakeHandler : HttpMessageHandler
    {
        // DropAfter >= 0: the connection drops once that many bytes of the body were accepted (the upload is interrupted).
        // NotSent: the request fails before the body is asked for (connection refused, TLS, DNS): no byte reaches the transport.
        sealed class Step { public Func<HttpRequestMessage, HttpResponseMessage> Respond; public long DropAfter = -1; public string NotSent; }
        readonly Queue<Step> _queue = new Queue<Step>();
        public List<RecordedRequest> Requests = new List<RecordedRequest>();

        public void Enqueue(int status, string body = "{}", string retryAfter = null, string location = null)
        {
            _queue.Enqueue(new Step { Respond = r =>
            {
                var resp = new HttpResponseMessage((HttpStatusCode)status) { Content = new StringContent(body, Encoding.UTF8, "application/json") };
                if (retryAfter != null) resp.Headers.TryAddWithoutValidation("Retry-After", retryAfter);
                if (location != null) resp.Headers.Location = new Uri(location);
                return resp;
            } });
        }

        /// <summary>Fails after the whole body was read (a failure while waiting for the answer).</summary>
        public void EnqueueFailure(string message = "connection reset")
        {
            _queue.Enqueue(new Step { Respond = r => { throw new HttpRequestException(message); } });
        }

        /// <summary>Fails before any byte of the body is handed to the transport (connect, TLS or DNS failure): the body is never read.</summary>
        public void EnqueueNotSent(string message = "No connection could be made because the target machine actively refused it.")
        {
            _queue.Enqueue(new Step { NotSent = message });
        }

        /// <summary>
        /// The server drops the connection after it accepted <paramref name="afterBytes"/> bytes of the body: writing the body to the transport
        /// fails with an IOException, which HttpClient reports as HttpRequestException "Error while copying content to a stream."
        /// </summary>
        public void EnqueueInterruptedUpload(long afterBytes)
        {
            _queue.Enqueue(new Step { DropAfter = afterBytes });
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var rec = new RecordedRequest
            {
                Method = request.Method.Method,
                Url = request.RequestUri.OriginalString,
                HasAuthorization = request.Headers.Authorization != null
            };
            Step step = _queue.Count == 0 ? null : _queue.Peek();
            if (step != null && step.NotSent != null)
            {
                _queue.Dequeue();
                lock (Requests) Requests.Add(rec);
                throw new HttpRequestException(step.NotSent);
            }
            if (step != null && step.DropAfter >= 0 && request.Content != null)
            {
                _queue.Dequeue();
                lock (Requests) Requests.Add(rec);
                try { await request.Content.CopyToAsync(new DroppingStream(step.DropAfter)).ConfigureAwait(false); }
                catch (IOException ex) { throw new HttpRequestException("Error while copying content to a stream.", ex); }
                throw new InvalidOperationException("the body was smaller than the limit of the interrupted upload");
            }
            if (request.Content != null)
            {
                rec.BodyBytes = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                rec.Body = Encoding.UTF8.GetString(rec.BodyBytes);
                rec.ContentLength = request.Content.Headers.ContentLength;
                rec.ContentType = request.Content.Headers.ContentType == null ? null : request.Content.Headers.ContentType.MediaType;
            }
            lock (Requests) Requests.Add(rec);
            if (_queue.Count == 0)
                return new HttpResponseMessage(HttpStatusCode.InternalServerError) { Content = new StringContent("{\"error\":{\"code\":\"NothingQueued\",\"message\":\"no response queued\"}}") };
            return _queue.Dequeue().Respond(request);
        }
    }

    /// <summary>
    /// An HttpMessageHandler that answers by route (method + text contained in the URL, first route added first tested), for
    /// whole scripts that make many calls in an order that depends on the script. Several answers on one route are given in
    /// turn, the last one then repeats. No route: 404 with the request in the message, so a missing route is easy to see.
    /// </summary>
    public sealed class RouteHandler : HttpMessageHandler
    {
        sealed class Route { public string Method, Contains; public Queue<Tuple<int, string>> Answers = new Queue<Tuple<int, string>>(); }
        readonly List<Route> _routes = new List<Route>();
        public List<RecordedRequest> Requests = new List<RecordedRequest>();

        public void Add(string method, string urlContains, int status, string body)
        {
            Route r = _routes.FirstOrDefault(x => x.Method == method && x.Contains == urlContains);
            if (r == null) { r = new Route { Method = method, Contains = urlContains }; _routes.Add(r); }
            r.Answers.Enqueue(Tuple.Create(status, body));
        }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var rec = new RecordedRequest
            {
                Method = request.Method.Method,
                Url = request.RequestUri.OriginalString,
                HasAuthorization = request.Headers.Authorization != null
            };
            if (request.Content != null)
            {
                rec.BodyBytes = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                rec.Body = Encoding.UTF8.GetString(rec.BodyBytes);
                rec.ContentLength = request.Content.Headers.ContentLength;
                rec.ContentType = request.Content.Headers.ContentType == null ? null : request.Content.Headers.ContentType.MediaType;
            }
            lock (Requests) Requests.Add(rec);
            Tuple<int, string> answer = null;
            lock (_routes)
            {
                foreach (Route r in _routes)
                    if (r.Method == request.Method.Method && rec.Url.Contains(r.Contains))
                    {
                        answer = r.Answers.Count > 1 ? r.Answers.Dequeue() : r.Answers.Peek();
                        break;
                    }
            }
            if (answer == null)
                return new HttpResponseMessage(HttpStatusCode.NotFound) { Content = new StringContent("{\"error\":{\"code\":\"NoRoute\",\"message\":\"no route for " + request.Method.Method + " " + rec.Url.Replace("\"", "'") + "\"}}", Encoding.UTF8, "application/json") };
            return new HttpResponseMessage((HttpStatusCode)answer.Item1) { Content = new StringContent(answer.Item2, Encoding.UTF8, "application/json") };
        }
    }

    /// <summary>Source of a large attachment without a large array: bytes generated on the fly.</summary>
    public sealed class PatternStream : Stream
    {
        long _length, _position;
        public PatternStream(long length) { _length = length; }
        public override bool CanRead { get { return true; } }
        public override bool CanSeek { get { return false; } }
        public override bool CanWrite { get { return false; } }
        public override long Length { get { return _length; } }
        public override long Position { get { return _position; } set { throw new NotSupportedException(); } }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count)
        {
            int n = (int)Math.Min(count, _length - _position);
            for (int i = 0; i < n; i++) buffer[offset + i] = (byte)((_position + i) * 31);
            _position += n;
            return n;
        }
        public override long Seek(long offset, SeekOrigin origin) { throw new NotSupportedException(); }
        public override void SetLength(long value) { throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { throw new NotSupportedException(); }
    }

    /// <summary>Builds a FastTransfer stream element by element. Everything is little endian.</summary>
    public sealed class FxBuilder
    {
        readonly Stream _s;
        public FxBuilder(Stream s) { _s = s; }
        public FxBuilder() { _s = new MemoryStream(); }

        public byte[] ToArray() { return ((MemoryStream)_s).ToArray(); }

        void W(byte[] b) { _s.Write(b, 0, b.Length); }
        void U16(ushort v) { W(BitConverter.GetBytes(v)); }
        void U32(uint v) { W(BitConverter.GetBytes(v)); }
        void Tag(ushort type, ushort id) { U16(type); U16(id); }

        public FxBuilder Marker(uint marker) { U32(marker); return this; }
        public FxBuilder Int16(ushort id, short v) { Tag(0x0002, id); W(BitConverter.GetBytes(v)); return this; }
        public FxBuilder Int32(ushort id, int v) { Tag(0x0003, id); W(BitConverter.GetBytes(v)); return this; }
        public FxBuilder Int64(ushort id, long v) { Tag(0x0014, id); W(BitConverter.GetBytes(v)); return this; }
        public FxBuilder Double(ushort id, double v) { Tag(0x0005, id); W(BitConverter.GetBytes(v)); return this; }
        public FxBuilder Bool(ushort id, bool v) { Tag(0x000B, id); U16((ushort)(v ? 1 : 0)); return this; }
        public FxBuilder Time(ushort id, DateTime utc) { Tag(0x0040, id); W(BitConverter.GetBytes(utc.ToFileTimeUtc())); return this; }
        public FxBuilder GuidValue(ushort id, Guid g) { Tag(0x0048, id); W(g.ToByteArray()); return this; }

        /// <summary>PtypString: length in bytes including the terminator, then UTF-16LE.</summary>
        public FxBuilder Str(ushort id, string s)
        {
            Tag(0x001F, id);
            byte[] b = Encoding.Unicode.GetBytes(s + "\0");
            U32((uint)b.Length); W(b);
            return this;
        }

        public FxBuilder Str8(ushort id, string s)
        {
            Tag(0x001E, id);
            byte[] b = Encoding.Latin1.GetBytes(s + "\0");
            U32((uint)b.Length); W(b);
            return this;
        }

        /// <summary>A string tagged with a code page type, like 0x84B0 (1200, UTF-16LE) in the importItem sample of Microsoft Learn.</summary>
        public FxBuilder CodePageStr(ushort id, ushort type, string s, Encoding enc)
        {
            Tag(type, id);
            byte[] b = enc.GetBytes(s + "\0");
            U32((uint)b.Length); W(b);
            return this;
        }

        public FxBuilder Bin(ushort id, byte[] data)
        {
            Tag(0x0102, id); U32((uint)data.Length); W(data);
            return this;
        }

        /// <summary>A binary value of the given size generated on the fly (an attachment of tens of megabytes).</summary>
        public FxBuilder BinLarge(ushort id, long size)
        {
            Tag(0x0102, id); U32((uint)size);
            var chunk = new byte[65536];
            long done = 0;
            while (done < size)
            {
                int n = (int)Math.Min(chunk.Length, size - done);
                for (int i = 0; i < n; i++) chunk[i] = (byte)((done + i) * 31);
                _s.Write(chunk, 0, n);
                done += n;
            }
            return this;
        }

        public FxBuilder Object(ushort id) { Tag(0x000D, id); return this; }

        public FxBuilder MvInt32(ushort id, int[] values)
        {
            Tag(0x1003, id); U32((uint)values.Length);
            foreach (int v in values) W(BitConverter.GetBytes(v));
            return this;
        }

        public FxBuilder MvStr(ushort id, string[] values)
        {
            Tag(0x101F, id); U32((uint)values.Length);
            foreach (string s in values) { byte[] b = Encoding.Unicode.GetBytes(s + "\0"); U32((uint)b.Length); W(b); }
            return this;
        }

        /// <summary>Named property (id >= 0x8000) identified by a GUID and a numeric id, then a PtypInteger32 value.</summary>
        public FxBuilder NamedLidInt32(ushort id, Guid g, uint lid, int value)
        {
            Tag(0x0003, id); W(g.ToByteArray()); _s.WriteByte(0); U32(lid); W(BitConverter.GetBytes(value));
            return this;
        }

        /// <summary>Named property identified by a GUID and a name, then a PtypString value.</summary>
        public FxBuilder NamedNameStr(ushort id, Guid g, string name, string value)
        {
            Tag(0x001F, id); W(g.ToByteArray()); _s.WriteByte(1); W(Encoding.Unicode.GetBytes(name + "\0"));
            byte[] b = Encoding.Unicode.GetBytes(value + "\0");
            U32((uint)b.Length); W(b);
            return this;
        }

        // Generic forms, written from the grammar of 2.2.4.1 (propType propInfo value), for the tests of the exact decoder.

        /// <summary>The tag alone (type then id): the caller writes what follows with Raw.</summary>
        public FxBuilder TagOnly(ushort type, ushort id) { Tag(type, id); return this; }
        /// <summary>A fixed-size value: exactly the bytes given.</summary>
        public FxBuilder Fixed(ushort type, ushort id, byte[] value) { Tag(type, id); W(value); return this; }
        /// <summary>A variable-size value: length, then the bytes.</summary>
        public FxBuilder Variable(ushort type, ushort id, byte[] value) { Tag(type, id); U32((uint)value.Length); W(value); return this; }
        /// <summary>A multi-valued property: count, then the values (each one with its length when <paramref name="prefixed"/>).</summary>
        public FxBuilder Multi(ushort type, ushort id, byte[][] values, bool prefixed)
        {
            Tag(type, id); U32((uint)values.Length);
            foreach (byte[] v in values) { if (prefixed) U32((uint)v.Length); W(v); }
            return this;
        }
        /// <summary>The tag of a named property identified by a GUID and a numeric id; the value follows (Raw).</summary>
        public FxBuilder NamedLidTag(ushort type, ushort id, Guid g, uint lid) { Tag(type, id); W(g.ToByteArray()); _s.WriteByte(0); U32(lid); return this; }
        /// <summary>The tag of a named property identified by a GUID and a name; the value follows (Raw).</summary>
        public FxBuilder NamedNameTag(ushort type, ushort id, Guid g, string name) { Tag(type, id); W(g.ToByteArray()); _s.WriteByte(1); W(Encoding.Unicode.GetBytes(name + "\0")); return this; }
        /// <summary>The tag of an extended property (bit 0x4000 and two bytes); the value follows (Raw).</summary>
        public FxBuilder ExtTag(ushort type, ushort id, ushort flags) { Tag((ushort)(type | 0x4000), id); U16(flags); return this; }

        public FxBuilder Raw(byte[] b) { W(b); return this; }

        /// <summary>
        /// A property seen on real streams (not in MS-OXCFXICS): the type carries the bit 0x4000 and two bytes (0x0200) sit between
        /// the tag and the value.
        /// </summary>
        FxBuilder Ext(ushort type, ushort id) { Tag((ushort)(type | 0x4000), id); U16(0x0200); return this; }
        public FxBuilder ExtTime(ushort id, DateTime utc) { Ext(0x0040, id); W(BitConverter.GetBytes(utc.ToFileTimeUtc())); return this; }
        public FxBuilder ExtInt16(ushort id, short v) { Ext(0x0002, id); W(BitConverter.GetBytes(v)); return this; }
        public FxBuilder ExtBool(ushort id, bool v) { Ext(0x000B, id); U16((ushort)(v ? 1 : 0)); return this; }
        public FxBuilder ExtBin(ushort id, byte[] data) { Ext(0x0102, id); U32((uint)data.Length); W(data); return this; }

        /// <summary>FXDelProp marker followed by the tag of the deleted property (what a real stream does before the recipients and the attachments).</summary>
        public FxBuilder DelProp(uint tag) { U32(0x40160003); U32(tag); return this; }

        /// <summary>
        /// The envelope of the real exports: header records, then the data in records of type 2 of at most <paramref name="chunk"/>
        /// bytes (31,680 in a real export; a record may end in the middle of a property). chunk 0: one record, as in the Learn samples.
        /// </summary>
        public static void WriteEnvelope(Stream dest, Stream data, int chunk, bool guidRecord)
        {
            Action<uint, byte[]> rec = (t, p) =>
            {
                dest.Write(BitConverter.GetBytes(t), 0, 4);
                dest.Write(BitConverter.GetBytes((uint)p.Length), 0, 4);
                dest.Write(p, 0, p.Length);
            };
            rec(1, new byte[] { 0, 0, 0, 0, 1, 0, 0, 0 });
            if (guidRecord) rec(3, new byte[] { 1, 0, 0, 0, 7, 3, 2, 0, 0, 0, 0, 0, 0xC0, 0, 0, 0, 0, 0, 0, 0x46, 0, 0x24, 0, 0x80 });
            rec(4, new byte[] { 0x14, 0x0F, 0xD6, 0x91, 0x86, 0x17 });
            if (chunk <= 0)
            {
                var all = new MemoryStream();
                data.CopyTo(all);
                rec(2, all.ToArray());
                return;
            }
            var buf = new byte[chunk];
            int n;
            while ((n = ReadFull(data, buf)) > 0)
            {
                dest.Write(BitConverter.GetBytes(2u), 0, 4);
                dest.Write(BitConverter.GetBytes((uint)n), 0, 4);
                dest.Write(buf, 0, n);
            }
        }

        static int ReadFull(Stream s, byte[] buf)
        {
            int total = 0;
            while (total < buf.Length)
            {
                int n = s.Read(buf, total, buf.Length - total);
                if (n <= 0) break;
                total += n;
            }
            return total;
        }

        /// <summary>The envelope seen in front of the data in the Graph samples: records (type, length, payload), data record type 2.</summary>
        public static byte[] Envelope(byte[] data) { return Envelope(data, 0, true); }

        public static byte[] Envelope(byte[] data, int chunk, bool guidRecord)
        {
            var ms = new MemoryStream();
            WriteEnvelope(ms, new MemoryStream(data), chunk, guidRecord);
            return ms.ToArray();
        }
    }
}
