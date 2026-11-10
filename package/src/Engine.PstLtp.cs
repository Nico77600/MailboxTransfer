// =============================================================================
//  Mailbox Transfer - engine, part 5b: MS-PST, LTP layer writer (Heap-on-Node, BTree-on-Heap, Property Context, Table Context)
// -----------------------------------------------------------------------------
//  PstProp      a property in its PST form (strings in UTF-16 without terminator, Boolean on 1 byte, multi-valued as a list)
//  PstHeap      Heap-on-Node over several data blocks: HNHDR / HNPAGEHDR / HNBITMAPHDR, page map, fill levels, HIDs
//  PstBth       BTree-on-Heap (header, leaf and index levels, nodes of at most 3,580 bytes)
//  PstPc        Property Context: BTH of (property id, type, HNID); values inline, in the heap (<= 3,580 bytes) or in a subnode
//  PstTable     Table Context: TCINFO + column descriptors, Row Matrix (heap or subnode, whole rows in each 8,176-byte block),
//               Cell Existence Block, RowIndex BTH
//  The builders produce a PstNode (blocks + subnodes); PstNdbWriter writes it.
// =============================================================================
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MailboxTransfer
{
    public sealed class PstProp
    {
        public ushort Id;
        public ushort Type;
        /// <summary>Single value in its PST form (a string: UTF-16LE without terminator; a Boolean: one byte).</summary>
        public byte[] Data;
        /// <summary>
        /// A single value too big to keep in memory, in its PST form (same convention as Data): the writer copies it from its source
        /// into a data tree, the reader (PstPcReader.LazyThreshold) gives it as the data tree itself, read block by block.
        /// Data is null when Source is set.
        /// </summary>
        public FxBlob Source;
        /// <summary>Multi-valued property: one entry per value (same convention as Data).</summary>
        public List<byte[]> Items;
        /// <summary>PtypObject: the embedded message, written as a subnode.</summary>
        public PstNode Object;
        /// <summary>Reader side: the subnode of a PtypObject value.</summary>
        public PstNodeView ObjectView;
        public long ObjectSize;
        /// <summary>NID to give to the embedded message in the subnode tree (0: a new LTP NID).</summary>
        public uint ObjectNid;
        /// <summary>Store the value in a subnode even when it would fit in the heap (PidTagAttachDataBinary: Outlook measures the attachment size there).</summary>
        public bool ForceSubnode;

        public uint Tag { get { return ((uint)Id << 16) | Type; } }
        public bool IsMulti { get { return (Type & 0x1000) != 0; } }

        /// <summary>Bytes of a single value, in memory or in its source.</summary>
        public long DataLength { get { return Source != null ? Source.Length : Data == null ? 0 : Data.Length; } }

        /// <summary>The same property with its value in memory (a table cell is written at Close, when the source may be gone).</summary>
        public PstProp Materialized()
        {
            if (Source == null) return this;
            return new PstProp { Id = Id, Type = Type, Data = Source.ToArray() };
        }

        public static PstProp Make(ushort id, ushort type, byte[] data) { return new PstProp { Id = id, Type = type, Data = data }; }
        public static PstProp Int32(ushort id, int v) { return Make(id, 0x0003, BitConverter.GetBytes(v)); }
        public static PstProp Int64(ushort id, long v) { return Make(id, 0x0014, BitConverter.GetBytes(v)); }
        public static PstProp Bool(ushort id, bool v) { return Make(id, 0x000B, new byte[] { (byte)(v ? 1 : 0) }); }
        public static PstProp Time(ushort id, DateTime utc) { return Make(id, 0x0040, BitConverter.GetBytes(utc.ToFileTimeUtc())); }
        public static PstProp Str(ushort id, string s) { return Make(id, 0x001F, Encoding.Unicode.GetBytes(s ?? "")); }
        public static PstProp Bin(ushort id, byte[] b) { return Make(id, 0x0102, b); }
    }

    public static class PstTypes
    {
        /// <summary>Size of the value kept in the HNID of a PC record for the types stored inline (0 if not inline).</summary>
        public static int InlineSize(ushort type)
        {
            switch (type)
            {
                case 0x0002: return 2;
                case 0x0003: case 0x0004: case 0x000A: return 4;
                case 0x000B: return 1;
            }
            return 0;
        }

        /// <summary>Size of the cell of a column in the Row Matrix: 8, 4, 2 or 1 (4 for an HNID).</summary>
        public static int CellSize(ushort type)
        {
            switch (type)
            {
                case 0x0005: case 0x0006: case 0x0007: case 0x0014: case 0x0040: return 8;
                case 0x0002: return 2;
                case 0x000B: return 1;
            }
            return 4;
        }

        public static bool IsFixedInTc(ushort type)
        {
            switch (type)
            {
                case 0x0002: case 0x0003: case 0x0004: case 0x000A: case 0x000B:
                case 0x0005: case 0x0006: case 0x0007: case 0x0014: case 0x0040: return true;
            }
            return false;
        }

        /// <summary>Fixed-size single types (and their size in a PC / TC value).</summary>
        public static int FixedSize(ushort baseType)
        {
            switch (baseType)
            {
                case 0x0002: return 2;
                case 0x0003: case 0x0004: case 0x000A: return 4;
                case 0x000B: return 1;
                case 0x0005: case 0x0006: case 0x0007: case 0x0014: case 0x0040: return 8;
                case 0x0048: return 16;
            }
            return -1;
        }

        /// <summary>Bytes of a multi-valued property as stored in the heap or in a subnode (2.3.3.4).</summary>
        public static byte[] PackMulti(ushort type, List<byte[]> items)
        {
            int fixedSize = FixedSize((ushort)(type & 0x0FFF));
            if (items == null) items = new List<byte[]>();
            if (fixedSize > 0)
            {
                var r = new byte[fixedSize * items.Count];
                for (int i = 0; i < items.Count; i++) Buffer.BlockCopy(items[i], 0, r, i * fixedSize, fixedSize);
                return r;
            }
            int head = 4 + 4 * items.Count, total = head;
            foreach (byte[] it in items) total += it.Length;
            var b = new byte[total];
            PstFormat.P32(b, 0, (uint)items.Count);
            int pos = head;
            for (int i = 0; i < items.Count; i++)
            {
                PstFormat.P32(b, 4 + 4 * i, (uint)pos);
                Buffer.BlockCopy(items[i], 0, b, pos, items[i].Length);
                pos += items[i].Length;
            }
            return b;
        }
    }

    /// <summary>Heap-on-Node (2.3.1): allocations are never split across blocks; HID = block index (16 bits), index (11 bits), type 0.</summary>
    public sealed class PstHeap
    {
        sealed class Blk
        {
            public byte[] Buf = new byte[PstFormat.MaxBlockData];
            public int Hdr, Used;
            public readonly List<int> Offs = new List<int>();
        }

        readonly byte _clientSig;
        readonly List<Blk> _blocks = new List<Blk>();

        public PstHeap(byte clientSig) { _clientSig = clientSig; }

        static int HeaderSize(int block)
        {
            if (block == 0) return 12;
            if (block >= 8 && (block - 8) % 128 == 0) return 66;
            return 2;
        }

        public int BlockCount { get { return _blocks.Count; } }

        bool Fits(Blk b, int len)
        {
            int dataEnd = b.Hdr + b.Used + len;
            int map = (dataEnd + 1) & ~1;
            return map + 4 + 2 * (b.Offs.Count + 2) <= PstFormat.MaxBlockData && b.Offs.Count + 1 < 2047;
        }

        /// <summary>Reserves len bytes (1..3,580) and returns the HID; the bytes are zero until Set / Alloc fills them.</summary>
        public uint Reserve(int len)
        {
            if (len <= 0 || len > PstFormat.MaxHeapAlloc) throw new ArgumentOutOfRangeException("len", "A heap allocation is 1.." + PstFormat.MaxHeapAlloc + " bytes.");
            if (_blocks.Count == 0 || !Fits(_blocks[_blocks.Count - 1], len))
            {
                if (_blocks.Count >= 0xFFFF) throw new InvalidOperationException("Heap too large.");
                _blocks.Add(new Blk { Hdr = HeaderSize(_blocks.Count) });
                _blocks[_blocks.Count - 1].Used = 0;
            }
            int bi = _blocks.Count - 1;
            Blk b = _blocks[bi];
            b.Offs.Add(b.Hdr + b.Used);
            b.Used += len;
            return ((uint)bi << 16) | ((uint)b.Offs.Count << 5);
        }

        public uint Alloc(byte[] data, int off, int len)
        {
            uint hid = Reserve(len);
            Blk b = _blocks[(int)(hid >> 16)];
            Buffer.BlockCopy(data, off, b.Buf, b.Offs[b.Offs.Count - 1], len);
            return hid;
        }

        public uint Alloc(byte[] data) { return Alloc(data, 0, data.Length); }

        public void Set(uint hid, byte[] data)
        {
            Blk b = _blocks[(int)(hid >> 16)];
            int idx = (int)((hid >> 5) & 0x7FF) - 1;
            Buffer.BlockCopy(data, 0, b.Buf, b.Offs[idx], data.Length);
        }

        public List<byte[]> Finish(uint hidUserRoot)
        {
            if (_blocks.Count == 0) _blocks.Add(new Blk { Hdr = 12 });
            int n = _blocks.Count;
            var cbs = new int[n];
            var res = new List<byte[]>(n);
            for (int i = 0; i < n; i++)
            {
                Blk b = _blocks[i];
                int dataEnd = b.Hdr + b.Used;
                int ibHnpm = (dataEnd + 1) & ~1;
                int cb = ibHnpm + 4 + 2 * (b.Offs.Count + 1);
                cbs[i] = cb;
                // every block of a data tree but the last must be full (8,176 bytes): Outlook rejects a node otherwise
                var o = new byte[i < n - 1 ? PstFormat.MaxBlockData : cb];
                Buffer.BlockCopy(b.Buf, b.Hdr, o, b.Hdr, b.Used);
                PstFormat.P16(o, 0, ibHnpm);
                PstFormat.P16(o, ibHnpm, b.Offs.Count);
                PstFormat.P16(o, ibHnpm + 2, 0);
                for (int k = 0; k < b.Offs.Count; k++) PstFormat.P16(o, ibHnpm + 4 + 2 * k, b.Offs[k]);
                PstFormat.P16(o, ibHnpm + 4 + 2 * b.Offs.Count, dataEnd);
                res.Add(o);
            }
            // headers
            byte[] first = res[0];
            first[2] = 0xEC; first[3] = _clientSig;
            PstFormat.P32(first, 4, hidUserRoot);
            for (int i = 0; i < Math.Min(8, n); i++)
            {
                int lvl = PstFormat.FillLevel(PstFormat.MaxBlockData - cbs[i]);
                first[8 + i / 2] |= (byte)(lvl << (4 * (i % 2)));
            }
            for (int i = 8; i < n; i += 128)
            {
                if ((i - 8) % 128 != 0) continue;
                byte[] bm = res[i];
                for (int k = 0; k < 128 && i + k < n; k++)
                {
                    int lvl = PstFormat.FillLevel(PstFormat.MaxBlockData - cbs[i + k]);
                    bm[2 + k / 2] |= (byte)(lvl << (4 * (k % 2)));
                }
            }
            return res;
        }
    }

    /// <summary>BTree-on-Heap (2.3.2). Records are given sorted by key (unsigned, little endian).</summary>
    public static class PstBth
    {
        /// <summary>Allocates the BTHHEADER first (so that it is HID 0x20 in an empty heap), then the nodes. Returns the HID of the header.</summary>
        public static uint Build(PstHeap heap, int cbKey, int cbEnt, List<byte[]> records) { return Build(heap, cbKey, cbEnt, records, 0); }

        /// <summary>Same, with a header allocated by the caller before the values the records point to (hdr != 0).</summary>
        public static uint Build(PstHeap heap, int cbKey, int cbEnt, List<byte[]> records, uint hdr)
        {
            if (hdr == 0) hdr = heap.Reserve(8);
            int rec = cbKey + cbEnt;
            uint hidRoot = 0;
            int levels = 0;
            if (records.Count > 0)
            {
                int perLeaf = Math.Max(1, PstFormat.MaxHeapAlloc / rec);
                var keys = new List<byte[]>();
                var hids = new List<uint>();
                for (int i = 0; i < records.Count; i += perLeaf)
                {
                    int n = Math.Min(perLeaf, records.Count - i);
                    var node = new byte[n * rec];
                    for (int k = 0; k < n; k++) Buffer.BlockCopy(records[i + k], 0, node, k * rec, rec);
                    keys.Add(records[i]);
                    hids.Add(heap.Alloc(node));
                }
                int perIdx = Math.Max(2, PstFormat.MaxHeapAlloc / (cbKey + 4));
                while (hids.Count > 1)
                {
                    var nk = new List<byte[]>();
                    var nh = new List<uint>();
                    for (int i = 0; i < hids.Count; i += perIdx)
                    {
                        int n = Math.Min(perIdx, hids.Count - i);
                        var node = new byte[n * (cbKey + 4)];
                        for (int k = 0; k < n; k++)
                        {
                            Buffer.BlockCopy(keys[i + k], 0, node, k * (cbKey + 4), cbKey);
                            PstFormat.P32(node, k * (cbKey + 4) + cbKey, hids[i + k]);
                        }
                        nk.Add(keys[i]);
                        nh.Add(heap.Alloc(node));
                    }
                    keys = nk; hids = nh; levels++;
                }
                hidRoot = hids[0];
            }
            var h = new byte[8];
            h[0] = 0xB5; h[1] = (byte)cbKey; h[2] = (byte)cbEnt; h[3] = (byte)levels;
            PstFormat.P32(h, 4, hidRoot);
            heap.Set(hdr, h);
            return hdr;
        }
    }

    public static class PstLtp
    {
        /// <summary>Stores bytes in the heap (<= 3,580 bytes) or in a new subnode of the node; returns the HNID (0 for no data).</summary>
        public static uint PlaceBytes(PstHeap heap, PstNode node, byte[] data)
        {
            if (data == null || data.Length == 0) return 0;
            if (data.Length <= PstFormat.MaxHeapAlloc) return heap.Alloc(data, 0, data.Length);
            uint nid = node.NewLtpNid();
            node.Subnodes[nid] = new PstNode { Raw = data };
            return nid;
        }

        /// <summary>HNID of a variable-size or multi-valued or object property.</summary>
        public static uint PlaceValue(PstHeap heap, PstNode node, PstProp p)
        {
            if ((p.Type & 0x0FFF) == 0x000D)
            {
                uint nid = p.ObjectNid != 0 ? p.ObjectNid : node.NewLtpNid();
                node.Subnodes[nid] = p.Object;
                var b = new byte[8];
                PstFormat.P32(b, 0, nid); PstFormat.P32(b, 4, (uint)Math.Min(p.ObjectSize, (long)uint.MaxValue));
                return heap.Alloc(b);
            }
            if (p.Source != null && !p.IsMulti)
            {
                if (p.Source.Length == 0) return 0;
                uint snid = node.NewLtpNid();
                node.Subnodes[snid] = new PstNode { Source = p.Source };
                return snid;
            }
            if (p.ForceSubnode && !p.IsMulti && p.Data != null && p.Data.Length > 0)
            {
                uint nid = node.NewLtpNid();
                node.Subnodes[nid] = new PstNode { Raw = p.Data };
                return nid;
            }
            return PlaceBytes(heap, node, p.IsMulti ? PstTypes.PackMulti(p.Type, p.Items) : p.Data);
        }

        /// <summary>The value of a fixed-size property as it is stored in a cell (size bytes).</summary>
        public static byte[] FixedBytes(PstProp p, int size)
        {
            var b = new byte[size];
            if (p.Data != null) Buffer.BlockCopy(p.Data, 0, b, 0, Math.Min(size, p.Data.Length));
            return b;
        }
    }

    /// <summary>Property Context (2.3.3).</summary>
    public static class PstPc
    {
        public static PstNode Build(IEnumerable<PstProp> props)
        {
            var list = new List<PstProp>();
            var seen = new HashSet<ushort>();
            foreach (PstProp p in props.OrderBy(x => x.Id)) if (seen.Add(p.Id)) list.Add(p);
            var heap = new PstHeap(0xBC);
            var node = new PstNode();
            var recs = new List<byte[]>(list.Count);
            uint hdr = heap.Reserve(8);       // BTHHEADER first (HID 0x20 = hidUserRoot), then the values, then the BTH nodes
            var values = new uint[list.Count];
            for (int i = 0; i < list.Count; i++)
            {
                PstProp p = list[i];
                int inl = PstTypes.InlineSize(p.Type);
                if (inl > 0 && !p.IsMulti) continue;
                values[i] = PstLtp.PlaceValue(heap, node, p);
            }
            for (int i = 0; i < list.Count; i++)
            {
                PstProp p = list[i];
                var r = new byte[8];
                PstFormat.P16(r, 0, p.Id); PstFormat.P16(r, 2, p.Type);
                int inl = PstTypes.InlineSize(p.Type);
                if (inl > 0 && !p.IsMulti)
                {
                    if (p.Data != null) Buffer.BlockCopy(p.Data, 0, r, 4, Math.Min(inl, p.Data.Length));
                }
                else PstFormat.P32(r, 4, values[i]);
                recs.Add(r);
            }
            PstBth.Build(heap, 2, 6, recs, hdr);
            node.Blocks = heap.Finish(hdr);
            return node;
        }
    }

    /// <summary>Table Context (2.3.4). Columns are (property id, type); the row id and row version columns are added.</summary>
    public sealed class PstTable
    {
        public sealed class Row { public uint RowId, RowVer; public PstProp[] Cells; }

        public readonly List<uint> Columns;     // tags, sorted, unique, without PidTagLtpRowId / PidTagLtpRowVer
        readonly Dictionary<uint, int> _index = new Dictionary<uint, int>();
        readonly List<Row> _rows = new List<Row>();

        public PstTable(IEnumerable<uint> tags)
        {
            Columns = tags.Where(t => t != 0x67F20003 && t != 0x67F30003).Distinct().OrderBy(t => t).ToList();
            for (int i = 0; i < Columns.Count; i++) _index[Columns[i]] = i;
        }

        public int RowCount { get { return _rows.Count; } }

        /// <summary>Removes the last row (used only to build deliberately damaged files for the checker tests).</summary>
        public void RemoveLastRow() { if (_rows.Count > 0) _rows.RemoveAt(_rows.Count - 1); }
        public int ColumnIndex(uint tag) { int i; return _index.TryGetValue(tag, out i) ? i : -1; }

        /// <summary>Adds a row; cells are aligned with Columns (null: the column does not exist in this row).</summary>
        public void AddRow(uint rowId, PstProp[] cells) { AddRow(rowId, cells, 0); }

        public void AddRow(uint rowId, PstProp[] cells, uint rowVer)
        {
            if (cells.Length != Columns.Count) throw new ArgumentException("cells");
            _rows.Add(new Row { RowId = rowId, RowVer = rowVer, Cells = cells });
        }

        /// <summary>Row built from a property bag: only the columns of the table are taken, with a matching type.</summary>
        public PstProp[] Pick(IEnumerable<PstProp> props)
        {
            var cells = new PstProp[Columns.Count];
            foreach (PstProp p in props)
            {
                int i = ColumnIndex(p.Tag);
                if (i >= 0 && cells[i] == null) cells[i] = p.Materialized();
            }
            return cells;
        }

        public PstNode Build()
        {
            var cols = new List<uint>(Columns);
            cols.Add(0x67F20003); cols.Add(0x67F30003);
            cols.Sort();
            int cCols = cols.Count;
            // layout of a row
            var ibData = new Dictionary<uint, int>();
            var cbData = new Dictionary<uint, int>();
            var iBit = new Dictionary<uint, int>();
            ibData[0x67F20003] = 0; cbData[0x67F20003] = 4; iBit[0x67F20003] = 0;
            ibData[0x67F30003] = 4; cbData[0x67F30003] = 4; iBit[0x67F30003] = 1;
            int next = 2;
            foreach (uint t in Columns) iBit[t] = next++;
            int pos = 8, end4 = 8, end2 = 8, end1 = 8;
            foreach (int size in new[] { 8, 4, 2, 1 })
            {
                foreach (uint t in Columns)
                {
                    if (PstTypes.CellSize((ushort)(t & 0xFFFF)) != size) continue;
                    ibData[t] = pos; cbData[t] = size; pos += size;
                }
                if (size == 4) end4 = pos;
                if (size == 2) end2 = pos;
                if (size == 1) end1 = pos;
            }
            int cebSize = (cCols + 7) / 8;
            int rowSize = end1 + cebSize;

            var heap = new PstHeap(0x7C);
            var node = new PstNode();
            uint info = heap.Reserve(22 + 8 * cCols);          // HID 0x20

            // rows
            int nRows = _rows.Count;
            if ((long)nRows * rowSize > int.MaxValue) throw new NotSupportedException("Table too large.");
            var matrix = nRows == 0 ? null : new byte[nRows * rowSize];
            var rowIndex = new List<byte[]>(nRows);
            for (int r = 0; r < nRows; r++)
            {
                Row row = _rows[r];
                int rb = r * rowSize;
                PstFormat.P32(matrix, rb, row.RowId);
                PstFormat.P32(matrix, rb + 4, row.RowVer);
                for (int c = 0; c < Columns.Count; c++)
                {
                    PstProp p = row.Cells[c];
                    if (p == null) continue;
                    uint tag = Columns[c];
                    int ib = rb + ibData[tag];
                    int cb = cbData[tag];
                    ushort type = (ushort)(tag & 0xFFFF);
                    if (PstTypes.IsFixedInTc(type) && !p.IsMulti)
                    {
                        byte[] v = PstLtp.FixedBytes(p, cb);
                        Buffer.BlockCopy(v, 0, matrix, ib, cb);
                    }
                    else PstFormat.P32(matrix, ib, PstLtp.PlaceValue(heap, node, p));
                    int bit = iBit[tag];
                    matrix[rb + end1 + bit / 8] |= (byte)(0x80 >> (bit % 8));
                }
                int rbit0 = 0, rbit1 = 1;
                matrix[rb + end1 + rbit0 / 8] |= (byte)(0x80 >> (rbit0 % 8));
                matrix[rb + end1 + rbit1 / 8] |= (byte)(0x80 >> (rbit1 % 8));
                var ri = new byte[8];
                PstFormat.P32(ri, 0, row.RowId); PstFormat.P32(ri, 4, (uint)r);
                rowIndex.Add(ri);
            }
            rowIndex.Sort((a, b) => PstFormat.G32(a, 0).CompareTo(PstFormat.G32(b, 0)));
            for (int i = 1; i < rowIndex.Count; i++)
                if (PstFormat.G32(rowIndex[i], 0) == PstFormat.G32(rowIndex[i - 1], 0)) throw new InvalidOperationException("Duplicate row id 0x" + PstFormat.G32(rowIndex[i], 0).ToString("X"));
            uint hidIndex = PstBth.Build(heap, 4, 4, rowIndex);

            uint hnidRows = 0;
            if (nRows > 0)
            {
                if (matrix.Length <= PstFormat.MaxHeapAlloc) hnidRows = heap.Alloc(matrix);
                else
                {
                    if (rowSize > PstFormat.MaxBlockData) throw new NotSupportedException("A row of " + rowSize + " bytes does not fit in a block.");
                    int perBlock = PstFormat.MaxBlockData / rowSize;
                    var blocks = new List<byte[]>();
                    for (int r = 0; r < nRows; r += perBlock)
                    {
                        int n = Math.Min(perBlock, nRows - r);
                        bool last = r + n >= nRows;
                        var b = new byte[last ? n * rowSize : PstFormat.MaxBlockData];
                        Buffer.BlockCopy(matrix, r * rowSize, b, 0, n * rowSize);
                        blocks.Add(b);
                    }
                    uint nid = node.NewLtpNid();
                    node.Subnodes[nid] = new PstNode { Blocks = blocks };
                    hnidRows = nid;
                }
            }

            var ti = new byte[22 + 8 * cCols];
            ti[0] = 0x7C; ti[1] = (byte)cCols;
            PstFormat.P16(ti, 2, end4); PstFormat.P16(ti, 4, end2); PstFormat.P16(ti, 6, end1); PstFormat.P16(ti, 8, rowSize);
            PstFormat.P32(ti, 10, hidIndex);
            PstFormat.P32(ti, 14, hnidRows);
            PstFormat.P32(ti, 18, 0);
            for (int c = 0; c < cCols; c++)
            {
                uint t = cols[c];
                int o = 22 + 8 * c;
                PstFormat.P32(ti, o, t);
                PstFormat.P16(ti, o + 4, ibData[t]);
                ti[o + 6] = (byte)cbData[t];
                ti[o + 7] = (byte)iBit[t];
            }
            heap.Set(info, ti);
            node.Blocks = heap.Finish(info);
            return node;
        }
    }
}
