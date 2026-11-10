#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of our own MS-PST writer and reader (Engine.PstNdb.cs, Engine.PstLtp.cs, Engine.PstFile.cs,
    Engine.PstWriter.cs, Engine.PstReader.cs): NDB layer (header, CRC, allocation maps, B-trees, data trees, subnode trees),
    LTP layer (heaps, BTH, property contexts, table contexts), Messaging layer (store, folders, messages, named properties,
    recipients, attachments, embedded messages) and, when out\fxics-corpus2 exists (git-ignored), the real items.
    Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\Pst.Tests.ps1 -Output Detailed
#>

BeforeDiscovery {
    $dir = Join-Path (Split-Path $PSScriptRoot -Parent) 'out\fxics-corpus2'
    $script:HasCorpus2 = (Test-Path -LiteralPath $dir) -and @(Get-ChildItem -LiteralPath $dir -Filter '*.fxs' -File).Count -gt 0
}

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    $script:Corpus2 = Join-Path $script:RepoRoot 'out\fxics-corpus2'

    function New-Prop { param([ushort]$Id, [ushort]$Type, $Data) [MailboxTransfer.PstProp]::Make($Id, $Type, [byte[]]$Data) }

    function Find-PstNode { param($File, [uint32]$Nid) $e = [MailboxTransfer.PstNbtEntry]::new(); if ($File.TryFindNode($Nid, [ref]$e)) { $e } }

    function New-PstPath { param([string]$Name) Join-Path $TestDrive ($Name + '.pst') }

    function Get-PageLevel { param($File, $Root, [byte]$Type) $File.ReadPage([long]$Root.Ib, $Type, $Root.Bid)[491] }

    # A message with: strings in every encoding, every fixed type, multi-valued properties, named properties (LID and name),
    # recipients, a file attachment, an embedded message holding an attachment that holds an embedded message.
    function New-SyntheticMessage {
        param([string]$Subject = 'Synthetic', [int]$AttachBytes = 5000, [int]$BodyChars = 6000)
        $P = [MailboxTransfer.FxProperty]
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($P::FromString(0x001A, 'IPM.Note'))
        $m.Properties.Add($P::FromString(0x0037, $Subject))
        $m.Properties.Add($P::FromString(0x1000, ('body ' * ($BodyChars / 5))))
        $m.Properties.Add($P::FromCodePageString(0x0070, 1252, 'topic é'))
        $m.Properties.Add($P::FromInt32(0x0E07, 0x09))
        $m.Properties.Add($P::FromInt32(0x0017, 1))
        $m.Properties.Add($P::FromBool(0x0023, $true))
        $m.Properties.Add($P::FromInt64(0x6000, 1234567890123))
        $m.Properties.Add($P::FromDouble(0x6001, 3.5))
        $m.Properties.Add($P::FromTime(0x0039, [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)))
        $m.Properties.Add($P::FromGuid(0x6002, [guid]'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'))
        $m.Properties.Add($P::FromInt16(0x6003, -5))
        $m.Properties.Add($P::FromBinary(0x1013, [byte[]](1..200)))
        $m.Properties.Add($P::FromStrings(0x6004, [string[]]@('one', 'two', 'three é')))
        $l5 = [Collections.Generic.List[byte[]]]::new(); $l5.Add([BitConverter]::GetBytes(7)); $l5.Add([BitConverter]::GetBytes(-9))
        $m.Properties.Add($P::CreateMulti(0x6005, 0x1003, $l5))
        $l6 = [Collections.Generic.List[byte[]]]::new(); $l6.Add([byte[]](1, 2, 3)); $l6.Add([byte[]]@()); $l6.Add([byte[]](9))
        $m.Properties.Add($P::CreateMulti(0x6006, 0x1102, $l6))
        $m.Properties.Add($P::FromString(0x6007, ('x' * 5000)))
        $g = [guid]'11111111-2222-3333-4444-555555555555'
        $m.Properties.Add($P::FromString(0x8001, 'named by lid').Named($g, [uint32]0x8501))
        $m.Properties.Add($P::FromInt32(0x8002, 42).Named($g, 'ByName'))
        $m.Properties.Add($P::FromInt32(0x8003, 7).Named([guid]'00062008-0000-0000-C000-000000000046', [uint32]0x8205))
        foreach ($i in 1..3) {
            $r = [MailboxTransfer.FxRecipient]::new()
            $r.Properties.Add($P::FromInt32(0x0C15, $i))
            $r.Properties.Add($P::FromString(0x3001, "Recipient $i"))
            $r.Properties.Add($P::FromString(0x39FE, "r$i@example.com"))
            $r.Properties.Add($P::FromBinary(0x0FFF, [byte[]](1..30)))
            $m.Recipients.Add($r)
        }
        $a = [MailboxTransfer.FxAttachment]::new()
        $bytes = [byte[]]::new($AttachBytes); [Random]::new(3).NextBytes($bytes)
        $a.Properties.Add($P::FromInt32(0x0E21, 0))
        $a.Properties.Add($P::FromInt32(0x3705, 1))
        $a.Properties.Add($P::FromString(0x3707, 'file.bin'))
        $a.Properties.Add($P::FromBinary(0x3701, $bytes))
        $m.Attachments.Add($a)
        # embedded level 1 holding an attachment that holds an embedded level 2
        $e1 = [MailboxTransfer.FxMessage]::new()
        $e1.Properties.Add($P::FromString(0x0037, 'Level 1'))
        $e1.Properties.Add($P::FromString(0x001A, 'IPM.Note'))
        $r1 = [MailboxTransfer.FxRecipient]::new(); $r1.Properties.Add($P::FromString(0x3001, 'Carol')); $e1.Recipients.Add($r1)
        $e2 = [MailboxTransfer.FxMessage]::new()
        $e2.Properties.Add($P::FromString(0x0037, 'Level 2'))
        $e2.Properties.Add($P::FromBinary(0x1009, [byte[]](9, 8, 7)))
        $ia = [MailboxTransfer.FxAttachment]::new()
        $ia.Properties.Add($P::FromInt32(0x3705, 5)); $ia.Properties.Add($P::Create(0x3701, 0x000D, [byte[]]@()))
        $ia.Embedded = $e2
        $e1.Attachments.Add($ia)
        $oa = [MailboxTransfer.FxAttachment]::new()
        $oa.Properties.Add($P::FromInt32(0x3705, 5)); $oa.Properties.Add($P::Create(0x3701, 0x000D, [byte[]]@()))
        $oa.Embedded = $e1
        $m.Attachments.Add($oa)
        return $m
    }

    function Assert-SameMessage {
        param($Original, $Copy)
        $d = [Collections.Generic.List[string]]::new()
        [void][MailboxTransfer.PstFxCompare]::Compare($Original, $Copy, $d)
        $d | Should -BeNullOrEmpty
    }
}

Describe 'CRC, signature and header (MS-PST 5.3, 5.5, 2.2.2.6)' {
    It 'computes the CRCs of the sample header of the specification (dwCRCPartial 0x379AA90E, dwCRCFull 0x1FD283D6)' {
        $hex = '534d1700130001015c070000d07b990b0400000001000000540200000000000045000000000400000004000004040000004000000200010004040000000400000004000000800000000400000004000000040000000400000404000004040000040400000004000000040000000400000004000000040000000400000004000000040000000400000004000000040000000400000004000000040000000400000f04000000000000000000000000000000249f000000000000449b000000000040f212000000000000000000000000004b0200000000000000529000000000005302000000000000000a9000000000000200000000000000ffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff800100003414000000000000'
        $bytes = [Convert]::FromHexString($hex)      # bytes 8..523 of the header
        $bytes.Length | Should -Be 516
        [MailboxTransfer.PstFormat]::Crc($bytes, 0, 471) | Should -Be 0x379AA90E
        [MailboxTransfer.PstFormat]::Crc($bytes, 0, 516) | Should -Be 0x1FD283D6
    }
    It 'has no pre / post inversion: the CRC of nothing is the seed, and it chains' {
        $b = [byte[]](1..100)
        [MailboxTransfer.PstFormat]::Crc($b, 0, 0) | Should -Be 0
        $whole = [MailboxTransfer.PstFormat]::Crc($b, 0, 100)
        $part = [MailboxTransfer.PstFormat]::Crc($b, 0, 37)
        [MailboxTransfer.PstFormat]::Crc($b, 37, 63, $part) | Should -Be $whole
    }
    It 'computes the block signature (ib xor bid, high word xor low word)' {
        [MailboxTransfer.PstFormat]::Sig(0x5000, 0x44) | Should -Be ((0x5000 -bxor 0x44) -band 0xFFFF)
        [MailboxTransfer.PstFormat]::Sig(0x123450000, 0x4) | Should -Be ((((0x23450000 -bxor 4) -shr 16) -bxor (0x23450000 -bxor 4)) -band 0xFFFF)
    }
    It 'turns the permutation of the specification into its inverse and back (and cyclic is symmetric)' {
        $b = [byte[]](0..255)
        $c = $b.Clone()
        [MailboxTransfer.PstCrypt]::Decode(1, $c, 0, 256, 4, $true)
        $c[0] | Should -Be 65
        $c[1] | Should -Be 54
        [MailboxTransfer.PstCrypt]::Decode(1, $c, 0, 256, 4, $false)
        $c | Should -Be $b
        $d = $b.Clone()
        [MailboxTransfer.PstCrypt]::Decode(2, $d, 0, 256, 0x1234, $true)
        $d | Should -Not -Be $b
        [MailboxTransfer.PstCrypt]::Decode(2, $d, 0, 256, 0x1234, $false)
        $d | Should -Be $b
    }
    It 'writes a header the reader accepts, with the documented fields' {
        $path = New-PstPath 'hdr'
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $w.Close()
        $raw = [IO.File]::ReadAllBytes($path)
        [Text.Encoding]::ASCII.GetString($raw, 0, 4) | Should -Be '!BDN'
        [BitConverter]::ToUInt16($raw, 10) | Should -Be 23
        [BitConverter]::ToUInt16($raw, 12) | Should -Be 19
        $raw[512] | Should -Be 0x80
        $raw[513] | Should -Be 0
        $raw[248] | Should -Be 2
        ($raw.Length - 0x4400) % 253952 | Should -Be 0
        $p = [MailboxTransfer.PstFile]::Open($path)
        try { $p.HeaderCrcPartialOk | Should -BeTrue; $p.HeaderCrcFullOk | Should -BeTrue; $p.FileEof | Should -Be $raw.Length } finally { $p.Dispose() }
    }
}

Describe 'NDB layer' {
    It 'builds BBT and NBT of three levels over thousands of nodes, in order, with the right counts' {
        $path = New-PstPath 'btree'
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $n = 9000
        for ($i = 1; $i -le $n; $i++) {
            $node = [MailboxTransfer.PstNode]::new()
            $node.Raw = [byte[]]@(($i % 251), 1, 2)
            $w.WriteTopLevelNode([uint32]($i * 32 + 4), $node, 0)
        }
        $w.Close()
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            (Get-PageLevel $p $p.NbtRoot 0x81) | Should -BeGreaterOrEqual 2
            (Get-PageLevel $p $p.BbtRoot 0x80) | Should -BeGreaterOrEqual 2
            $all = @($p.EnumerateNodes())
            $all.Count | Should -Be $n
            $all[0].Nid | Should -Be 36
            for ($i = 1; $i -le $n; $i += 97) {
                $e = Find-PstNode $p ([uint32]($i * 32 + 4))
                $e | Should -Not -BeNullOrEmpty
                $p.ReadData($e.BidData)[0] | Should -Be ($i % 251)
            }
            (Find-PstNode $p 0x7FFFFFF4) | Should -BeNullOrEmpty
            # BIDs: data blocks increment by 4 and have bit 1 clear
            ($all[0].BidData -band 3) | Should -Be 0
            ($all[1].BidData - $all[0].BidData) | Should -Be 4
        } finally { $p.Dispose() }
    }
    It 'stores data trees of every shape: 1 byte, 8176, 8177, 3 blocks, one XBLOCK, an XXBLOCK, 70 MB' -ForEach @(
        @{ Size = 1 }, @{ Size = 8176 }, @{ Size = 8177 }, @{ Size = 24000 }, @{ Size = 8176 * 1021 }, @{ Size = 8176 * 1021 + 1 }, @{ Size = 12MB }, @{ Size = 70MB }
    ) {
        $path = New-PstPath "tree$Size"
        $data = [byte[]]::new($Size)
        [Random]::new($Size).NextBytes($data)
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $node = [MailboxTransfer.PstNode]::new(); $node.Raw = $data
        $w.WriteTopLevelNode(0x8004, $node, 0)
        $w.Close()
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            $e = Find-PstNode $p 0x8004
            $back = $p.ReadData($e.BidData)
            $back.Length | Should -Be $Size
            [Linq.Enumerable]::SequenceEqual($back, $data) | Should -BeTrue
            $blocks = $p.ReadDataTreeBlocks($e.BidData)
            $blocks.Count | Should -Be ([math]::Ceiling($Size / 8176))
            if ($Size -gt 8176) { ($e.BidData -band 2) | Should -Be 2 } else { ($e.BidData -band 2) | Should -Be 0 }
            if ($Size -gt 8176 * 1021) {
                # XXBLOCK: level 2
                $p.ReadBlock($e.BidData)[1] | Should -Be 2
            } elseif ($Size -gt 8176) { $p.ReadBlock($e.BidData)[1] | Should -Be 1 }
        } finally { $p.Dispose() }
    }
    It 'stores subnode trees: a leaf block, then an SIBLOCK over several SLBLOCKs' -ForEach @(@{ Count = 5 }, @{ Count = 340 }, @{ Count = 341 }, @{ Count = 900 }) {
        $path = New-PstPath "sub$Count"
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $node = [MailboxTransfer.PstNode]::new(); $node.Raw = [byte[]](1, 2, 3)
        for ($s = 1; $s -le $Count; $s++) {
            $sn = [MailboxTransfer.PstNode]::new(); $sn.Raw = [byte[]]::new($s % 50 + 1)
            if ($s % 100 -eq 0) { $deep = [MailboxTransfer.PstNode]::new(); $deep.Raw = [byte[]](7, 7); $sn.Subnodes[0x25] = $deep }
            $node.Subnodes[[uint32]($s * 32 + 5)] = $sn
        }
        $w.WriteTopLevelNode(0x8004, $node, 0)
        $w.Close()
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            $v = $p.OpenNode(0x8004)
            @($v.SubnodeNids).Count | Should -Be $Count
            $v.OpenSubnode([uint32](5 * 32 + 5)).ReadAll().Length | Should -Be (5 % 50 + 1)
            if ($Count -ge 100) { $v.OpenSubnode([uint32](100 * 32 + 5)).OpenSubnode(0x25).ReadAll().Length | Should -Be 2 }
            $p.ReadBlock($v.BidSub)[0] | Should -Be 2
            $p.ReadBlock($v.BidSub)[1] | Should -Be $(if ($Count -gt 340) { 1 } else { 0 })
        } finally { $p.Dispose() }
    }
    It 'keeps AMaps, DList and the maps of the density pages consistent over several allocation sections' {
        $path = New-PstPath 'maps'
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $node = [MailboxTransfer.PstNode]::new(); $node.Raw = [byte[]]::new(34MB)
        $w.WriteTopLevelNode(0x8004, $node, 0)
        $w.Close()
        $raw = [IO.File]::ReadAllBytes($path)
        $sections = ($raw.Length - 0x4400) / 253952
        $sections | Should -BeGreaterThan 128     # 34 MB over 248 KB sections: more than 128, so the first FMap page of the file exists
        $crc = [MailboxTransfer.PstFormat]
        for ($s = 0; $s -lt $sections; $s++) {
            $at = 0x4400 + $s * 253952
            $raw[$at + 496] | Should -Be 0x84
            $raw[$at + 497] | Should -Be 0x84
            [BitConverter]::ToUInt32($raw, $at + 500) | Should -Be ($crc::Crc($raw[$at..($at + 495)], 0, 496))
            [BitConverter]::ToUInt64($raw, $at + 504) | Should -Be $at
            $raw[$at] | Should -Be 0xFF         # the AMap maps itself
            if ($s % 8 -eq 0) { $raw[$at + 512 + 496] | Should -Be 0x83 }
        }
        $raw[0x4200 + 496] | Should -Be 0x86
        $raw[0x4600 + 496] | Should -Be 0x83                                   # the first PMap
        $fm = 0x4400 + 128 * 253952 + 1024                                      # the first FMap of the file: section 128, after its AMap and PMap
        $raw[$fm + 496] | Should -Be 0x82
        [BitConverter]::ToUInt32($raw, $fm + 500) | Should -Be ([MailboxTransfer.PstFormat]::Crc($raw[$fm..($fm + 495)], 0, 496))
        [BitConverter]::ToUInt64($raw, $fm + 504) | Should -Be $fm
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            # free space accounting: the sum of the free bits of the AMaps equals cbAMapFree
            $free = 0
            for ($s = 0; $s -lt $sections; $s++) { $at = 0x4400 + $s * 253952; foreach ($b in $raw[$at..($at + 495)]) { $free += 8 - [Numerics.BitOperations]::PopCount([uint32]$b) } }
            $p.AMapFree | Should -Be ($free * 64)
            $p.AMapLast | Should -Be (0x4400 + ($sections - 1) * 253952)
        } finally { $p.Dispose() }
    }
}

Describe 'LTP layer' {
    It 'lays out a heap over many blocks (bitmap headers at blocks 8 and 136), every allocation readable through its HID' {
        $path = New-PstPath 'heap'
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $props = [Collections.Generic.List[MailboxTransfer.PstProp]]::new()
        $rnd = [Random]::new(11)
        $expect = @{}
        for ($i = 0; $i -lt 450; $i++) {
            $b = [byte[]]::new(3000); $rnd.NextBytes($b)
            $props.Add((New-Prop (0x3000 + $i) 0x0102 $b)); $expect[0x3000 + $i] = $b
        }
        $node = [MailboxTransfer.PstPc]::Build($props)
        $node.Blocks.Count | Should -BeGreaterThan 140
        $w.WriteTopLevelNode(0x8022, $node, 0)
        $w.Close()
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            $v = $p.OpenNode(0x8022)
            $h = [MailboxTransfer.PstHeapReader]::new($v.Blocks)
            $h.ClientSig | Should -Be 0xBC
            $h.UserRoot | Should -Be 0x20
            $v.Blocks[0][2] | Should -Be 0xEC
            # block 8 and block 136 start with the abridged header + 64-byte fill map: the first allocation starts at 66
            [BitConverter]::ToUInt16($v.Blocks[8], 2 + 64 + 0) | Should -Not -Be 0   # data follows the map
            $pc = [MailboxTransfer.PstPcReader]::new($v)
            $pc.Entries.Count | Should -Be 450
            foreach ($i in 0, 1, 7, 8, 100, 300, 449) {
                [Linq.Enumerable]::SequenceEqual($pc.Get(0x3000 + $i).Data, [byte[]]$expect[0x3000 + $i]) | Should -BeTrue
            }
        } finally { $p.Dispose() }
    }
    It 'builds a multi-level BTH with thousands of records and finds every key' {
        $path = New-PstPath 'bth'
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $props = [Collections.Generic.List[MailboxTransfer.PstProp]]::new()
        for ($i = 0; $i -lt 6000; $i++) { $props.Add([MailboxTransfer.PstProp]::Int32([ushort](0x2000 + $i), $i * 3)) }
        $node = [MailboxTransfer.PstPc]::Build($props)
        $w.WriteTopLevelNode(0x8022, $node, 0)
        $w.Close()
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            $pc = [MailboxTransfer.PstPcReader]::new($p.OpenNode(0x8022))
            $pc.Entries.Count | Should -Be 6000
            $hdr = $pc.Ltp.Heap.Get($pc.Ltp.Heap.UserRoot)
            $hdr[3] | Should -BeGreaterThan 0       # bIdxLevels: at least one index level
            for ($i = 0; $i -lt 6000; $i += 61) { [BitConverter]::ToInt32($pc.Get(0x2000 + $i).Data, 0) | Should -Be ($i * 3) }
            $pc.Find(0x1FFF) | Should -BeNullOrEmpty
        } finally { $p.Dispose() }
    }
    It 'keeps every property type: inline, 8-byte, GUID, strings, large values in subnodes, multi-valued, empty values' {
        $path = New-PstPath 'pc'
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $big = [byte[]]::new(100000); [Random]::new(1).NextBytes($big)
        $props = @(
            [MailboxTransfer.PstProp]::Int32(0x0001, -7), (New-Prop 0x0002 0x0002 ([BitConverter]::GetBytes([int16]-3))),
            [MailboxTransfer.PstProp]::Bool(0x0003, $true), [MailboxTransfer.PstProp]::Int64(0x0004, [long]::MaxValue),
            [MailboxTransfer.PstProp]::Time(0x0005, [datetime]::UtcNow), (New-Prop 0x0006 0x0048 ([guid]::NewGuid().ToByteArray())),
            [MailboxTransfer.PstProp]::Str(0x0007, 'héllo wörld'), [MailboxTransfer.PstProp]::Str(0x0008, ''),
            [MailboxTransfer.PstProp]::Bin(0x0009, $big), [MailboxTransfer.PstProp]::Str(0x000A, ('long ' * 3000)),
            (New-Prop 0x000B 0x0005 ([BitConverter]::GetBytes([double]2.5)))
        )
        $mv = [MailboxTransfer.PstProp]::new(); $mv.Id = 0x000C; $mv.Type = 0x101F
        $mv.Items = [Collections.Generic.List[byte[]]]::new(); foreach ($s in 'a', 'bb', '') { $mv.Items.Add([Text.Encoding]::Unicode.GetBytes($s)) }
        $mv2 = [MailboxTransfer.PstProp]::new(); $mv2.Id = 0x000D; $mv2.Type = 0x1003
        $mv2.Items = [Collections.Generic.List[byte[]]]::new(); foreach ($n in 1, 2, 3) { $mv2.Items.Add([BitConverter]::GetBytes($n)) }
        $all = [Collections.Generic.List[MailboxTransfer.PstProp]]::new(); foreach ($x in @($props) + @($mv, $mv2)) { $all.Add($x) }
        $node = [MailboxTransfer.PstPc]::Build($all)
        $node.Subnodes.Count | Should -Be 2         # the 100,000 bytes and the 30,000 characters
        $w.WriteTopLevelNode(0x8022, $node, 0)
        $w.Close()
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            $pc = [MailboxTransfer.PstPcReader]::new($p.OpenNode(0x8022))
            [BitConverter]::ToInt32($pc.Get(1).Data, 0) | Should -Be -7
            [BitConverter]::ToInt16($pc.Get(2).Data, 0) | Should -Be -3
            $pc.Get(3).Data[0] | Should -Be 1
            [BitConverter]::ToInt64($pc.Get(4).Data, 0) | Should -Be ([long]::MaxValue)
            ([Guid]::new([byte[]]$pc.Get(6).Data)) | Should -BeOfType [guid]
            [Text.Encoding]::Unicode.GetString($pc.Get(7).Data) | Should -Be 'héllo wörld'
            $pc.Get(8).Data.Length | Should -Be 0
            [Linq.Enumerable]::SequenceEqual($pc.Get(9).Data, $big) | Should -BeTrue
            $pc.Get(10).Data.Length | Should -Be 30000
            [BitConverter]::ToDouble($pc.Get(11).Data, 0) | Should -Be 2.5
            $pc.Get(12).Items.Count | Should -Be 3
            [Text.Encoding]::Unicode.GetString($pc.Get(12).Items[1]) | Should -Be 'bb'
            $pc.Get(12).Items[2].Length | Should -Be 0
            ($pc.Get(13).Items | % { [BitConverter]::ToInt32($_, 0) }) | Should -Be @(1, 2, 3)
        } finally { $p.Dispose() }
    }
    It 'builds tables whose rows span many data blocks, with a Cell Existence Block and a sparse row index' {
        $path = New-PstPath 'tc'
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $cols = [uint32[]](0x0037001F, 0x00170003, 0x00390040, 0x0E080003, 0x30080040, 0x001A001F, 0x360A000B, 0x0002001F, 0x00030003, 0x10131102, 0x6000001F, 0x6001000B, 0x60020014, 0x60030002, 0x60040048)
        $tc = [MailboxTransfer.PstTable]::new($cols)
        $rows = 20000
        for ($r = 1; $r -le $rows; $r++) {
            $cells = [MailboxTransfer.PstProp[]]::new($tc.Columns.Count)
            $cells[$tc.ColumnIndex(0x0037001F)] = [MailboxTransfer.PstProp]::Str(0x37, "Subject $r")
            $cells[$tc.ColumnIndex(0x00170003)] = [MailboxTransfer.PstProp]::Int32(0x17, $r)
            $cells[$tc.ColumnIndex(0x60020014)] = [MailboxTransfer.PstProp]::Int64(0x6002, [long]$r * 1000000007)
            if ($r % 3 -eq 0) { $cells[$tc.ColumnIndex(0x360A000B)] = [MailboxTransfer.PstProp]::Bool(0x360A, ($r % 2 -eq 0)) }
            if ($r % 5 -eq 0) { $cells[$tc.ColumnIndex(0x6000001F)] = [MailboxTransfer.PstProp]::Str(0x6000, ('v' * 4000)) }   # in a subnode
            $tc.AddRow([uint32]($r * 8 + 2), $cells)
        }
        $node = $tc.Build()
        $node.Subnodes.Count | Should -BeGreaterThan 1       # the row matrix and the large strings
        $w.WriteTopLevelNode(0x60E, $node, 0)
        $w.Close()
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            $t = [MailboxTransfer.PstTcReader]::new($p.OpenNode(0x60E))
            $t.RowCount | Should -Be $rows
            # rows never straddle two blocks: floor(8176 / rowSize) per block, every block but the last has cb 8176
            $perBlock = [math]::Floor(8176 / $t.RowSize)
            $t.RowSize | Should -BeGreaterThan 40
            $t.Rgib[3] | Should -Be $t.RowSize
            ($t.Columns | ? { $_.Id -eq 0x67F2 }).IBit | Should -Be 0
            ($t.Columns | ? { $_.Id -eq 0x67F2 }).IbData | Should -Be 0
            ($t.Columns | ? { $_.Id -eq 0x67F3 }).IBit | Should -Be 1
            ($t.Columns | ? { $_.Id -eq 0x67F3 }).IbData | Should -Be 4
            $sorted = $t.Columns.Tag | Sort-Object
            $t.Columns.Tag | Should -Be $sorted
            foreach ($r in 0, 1, 2, 3, 4, 5, ($perBlock - 1), $perBlock, ($perBlock + 1), 9999, 19999) {
                $n = $r + 1
                $t.RowId($r) | Should -Be ($n * 8 + 2)
                [Text.Encoding]::Unicode.GetString($t.GetCell($r, 0x37).Data) | Should -Be "Subject $n"
                [BitConverter]::ToInt64($t.GetCell($r, 0x6002).Data, 0) | Should -Be ([long]$n * 1000000007)
                ($null -ne $t.GetCell($r, 0x360A)) | Should -Be ($n % 3 -eq 0)
                if ($n % 5 -eq 0) { $t.GetCell($r, 0x6000).Data.Length | Should -Be 8000 } else { $t.GetCell($r, 0x6000) | Should -BeNullOrEmpty }
            }
            $rowMatrixBlocks = $null
            foreach ($sn in $t.Ltp.Node.SubnodeNids) { $bl = $t.Ltp.Node.OpenSubnode([uint32]$sn).Blocks; if ($null -eq $rowMatrixBlocks -or $bl.Count -gt $rowMatrixBlocks.Count) { $rowMatrixBlocks = $bl } }
            $rowMatrixBlocks.Count | Should -Be ([math]::Ceiling($rows / $perBlock))
            $rowMatrixBlocks[0].Length | Should -Be 8176
        } finally { $p.Dispose() }
    }
    It 'builds an empty table and a table small enough to live in the heap' {
        $path = New-PstPath 'tcsmall'
        $w = [MailboxTransfer.PstNdbWriter]::new($path)
        $w.WriteTopLevelNode(0x60D, ([MailboxTransfer.PstTable]::new([uint32[]](0x3001001F, 0x36020003))).Build(), 0)
        $tc = [MailboxTransfer.PstTable]::new([uint32[]](0x3001001F, 0x36020003))
        $tc.AddRow(0x8022, @([MailboxTransfer.PstProp]::Str(0x3001, 'Only'), [MailboxTransfer.PstProp]::Int32(0x3602, 5)))
        $w.WriteTopLevelNode(0x60E, $tc.Build(), 0)
        $w.Close()
        $p = [MailboxTransfer.PstFile]::Open($path)
        try {
            $e = [MailboxTransfer.PstTcReader]::new($p.OpenNode(0x60D)); $e.RowCount | Should -Be 0
            $s = [MailboxTransfer.PstTcReader]::new($p.OpenNode(0x60E)); $s.RowCount | Should -Be 1
            [Text.Encoding]::Unicode.GetString($s.GetCell(0, 0x3001).Data) | Should -Be 'Only'
            $s.Ltp.Node.SubnodeNids | Should -BeNullOrEmpty       # row matrix in the heap
        } finally { $p.Dispose() }
    }
}

Describe 'Messaging layer' {
    It 'writes a store with the mandatory nodes and an empty hierarchy the reader and the properties of the store agree on' {
        $path = New-PstPath 'empty'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'My archive')
        $w.Close()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $r.DisplayName | Should -Be 'My archive'
            $r.RecordKey.Length | Should -Be 16
            foreach ($nid in 0x21, 0x61, 0x122, 0x12D, 0x12E, 0x12F, 0x1E1, 0x201, 0x60D, 0x60E, 0x60F, 0x610, 0x671, 0x692, 0x8022, 0x802D, 0x802E, 0x802F, 0x8042, 0x804D, 0x804E, 0x804F, 0x8062, 0x806D, 0x806E, 0x806F, 0x62B, 0x64C, 0x6B6, 0x6D7, 0x6F8, 0xC01, 0xEC1) {
                $r.File.HasNode([uint32]$nid) | Should -BeTrue -Because ('node 0x{0:X}' -f $nid)
            }
            $names = @($r.AllFolders() | % Name)
            $names | Should -Contain 'My archive'
            $names | Should -Contain 'Deleted Items'
            $names | Should -Contain 'Search Root'
            # entry ids of the store point to existing folders
            $sp = [MailboxTransfer.PstPcReader]::new($r.File.OpenNode(0x21))
            $ipm = $sp.Get(0x35E0).Data
            $ipm.Length | Should -Be 24
            [BitConverter]::ToUInt32($ipm, 20) | Should -Be 0x8022
            [BitConverter]::ToUInt32($sp.Get(0x35E3).Data, 20) | Should -Be 0x8062
            [BitConverter]::ToUInt32($sp.Get(0x35E7).Data, 20) | Should -Be 0x8042
            [Linq.Enumerable]::SequenceEqual([byte[]]($ipm[4..19]), [byte[]]$r.RecordKey) | Should -BeTrue
        } finally { $r.Dispose() }
    }
    It 'writes the optional spam search folder when asked' {
        $path = New-PstPath 'spam'
        $o = [MailboxTransfer.PstWriterOptions]::new(); $o.SpamSearchFolder = $true
        $w = [MailboxTransfer.PstWriter]::Create($path, 'S', $o); $w.Close()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try { $r.File.HasNode(0x2223) | Should -BeTrue; ($r.AllFolders() | % Name) | Should -Contain 'SPAM Search Folder 2' } finally { $r.Dispose() }
    }
    It 'round-trips a message with every property type, named properties, recipients, attachments and two levels of embedded messages' {
        $path = New-PstPath 'msg'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'T')
        $f = $w.CreateFolder($null, 'Folder', 'IPF.Note')
        $m = New-SyntheticMessage
        $nid = $w.AddMessage($f, $m)
        $w.Close()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $copy = $r.ReadMessage($nid)
            Assert-SameMessage $m $copy
            $copy.Attachments.Count | Should -Be 2
            $copy.Attachments[1].Embedded.Attachments[0].Embedded.GetString(0x37) | Should -Be 'Level 2'
            $copy.Find([guid]'11111111-2222-3333-4444-555555555555', 'ByName').AsInt32() | Should -Be 42
            $copy.Find([guid]'11111111-2222-3333-4444-555555555555', [uint32]0x8501).AsString() | Should -Be 'named by lid'
            $copy.Find([guid]'00062008-0000-0000-C000-000000000046', [uint32]0x8205).AsInt32() | Should -Be 7
            ($copy.Find(0x0E04)).AsString() | Should -Be 'Recipient 1'       # PidTagDisplayTo built from the To recipients
            ($copy.Find(0x0E07)).AsInt32() -band 0x10 | Should -Be 0x10       # mfHasAttach
            $copy.Find(0x0E08).AsInt32() | Should -BeGreaterThan 11000       # the size covers the body and the attachment
        } finally { $r.Dispose() }
    }
    It 'assigns one id per named property, shared by messages, in the name-to-id map with its hash buckets' {
        $path = New-PstPath 'names'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'T')
        $f = $w.CreateFolder($null, 'F', 'IPF.Note')
        $g = [guid]'11111111-2222-3333-4444-555555555555'
        $ids = @()
        foreach ($i in 1..40) {
            $m = [MailboxTransfer.FxMessage]::new()
            $m.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x37, "m$i"))
            $m.Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x8000 + ($i % 7), $i).Named($g, [uint32](0x9000 + ($i % 7))))
            $m.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x8010, "x$i").Named($g, "Name$($i % 5)"))
            [void]$w.AddMessage($f, $m)
        }
        $w.Close()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $pc = [MailboxTransfer.PstPcReader]::new($r.File.OpenNode(0x61))
            [BitConverter]::ToInt32($pc.Get(1).Data, 0) | Should -Be 251
            $entries = $pc.Get(3).Data
            ($entries.Length / 8) | Should -Be 12         # 7 numeric + 5 string names
            $pc.Get(2).Data.Length | Should -Be 16        # one GUID
            # every entry is in the bucket the formula gives
            $found = 0
            for ($b = 0; $b -lt 251; $b++) { $e = $pc.Get(0x1000 + $b); if ($e) { $found += $e.Data.Length / 8 } }
            $found | Should -Be 12
            for ($i = 0; $i -lt 12; $i++) {
                $dw0 = [BitConverter]::ToUInt32($entries, 8 * $i)
                $name = [BitConverter]::ToUInt16($entries, 8 * $i + 4)
                if (($name -band 1) -eq 1) { $str = $pc.Get(4).Data; $len = [BitConverter]::ToUInt32($str, $dw0); $dw0 = [MailboxTransfer.PstNameMap]::NameHash([Text.Encoding]::Unicode.GetString($str, $dw0 + 4, $len)) }
                $bucket = ($dw0 -bxor $name) % 251
                $bytes = $pc.Get(0x1000 + $bucket).Data
                $has = $false
                for ($k = 0; $k -lt $bytes.Length; $k += 8) { if ([BitConverter]::ToUInt32($bytes, $k) -eq $dw0 -and [BitConverter]::ToUInt16($bytes, $k + 6) -eq $i) { $has = $true } }
                $has | Should -BeTrue
            }
            foreach ($nid in $r.MessageNids(($r.AllFolders() | ? Name -EQ 'F'), $false)) {
                $c = $r.ReadMessage($nid)
                $idv = $c.Find($g, [uint32]0x9001)
                if ($idv) { $idv.Id | Should -BeGreaterOrEqual 0x8000 }
            }
        } finally { $r.Dispose() }
    }
    It 'keeps folder counts and rows of the contents tables, and puts associated messages in the FAI table' {
        $path = New-PstPath 'counts'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'T')
        $a = $w.CreateFolder($null, 'A', 'IPF.Note')
        $b = $w.CreateFolder($a, 'B', 'IPF.Appointment')
        foreach ($i in 1..30) {
            $m = [MailboxTransfer.FxMessage]::new()
            $m.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x37, "m$i"))
            $m.Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x0E07, $(if ($i % 3 -eq 0) { 1 } else { 0 })))
            [void]$w.AddMessage($b, $m)
        }
        $fai = [MailboxTransfer.FxMessage]::new(); $fai.StartMarker = [MailboxTransfer.FxMessage]::MarkerStartFAIMsg
        $fai.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x1A, 'IPM.Microsoft.FolderDesign.NamedView'))
        [void]$w.AddMessage($b, $fai)
        $w.MessagesWritten | Should -Be 31
        $w.Close()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $fa = $r.AllFolders() | ? Name -EQ 'A'; $fb = $r.AllFolders() | ? Name -EQ 'B'
            $fa.HasSubfolders | Should -BeTrue; $fb.HasSubfolders | Should -BeFalse
            $fb.Path | Should -Be 'T/A/B'
            $fb.ContentCount | Should -Be 30
            $fb.UnreadCount | Should -Be 20
            $fb.ContainerClass | Should -Be 'IPF.Appointment'
            $r.MessageNids($fb, $false).Count | Should -Be 30
            $r.MessageNids($fb, $true).Count | Should -Be 1
            ($r.MessageNids($fb, $true)[0] -band 0x1F) | Should -Be 8     # NID_TYPE_ASSOC_MESSAGE
            ($r.MessageNids($fb, $false)[0] -band 0x1F) | Should -Be 4
            # the hierarchy table of A carries the final counts of B
            $ht = [MailboxTransfer.PstTcReader]::new($r.File.OpenNode([MailboxTransfer.PstFormat]::Nid(13, [MailboxTransfer.PstFormat]::NidIndex($fa.Nid))))
            $ht.RowCount | Should -Be 1
            [BitConverter]::ToInt32($ht.GetCell(0, 0x3602).Data, 0) | Should -Be 30
        } finally { $r.Dispose() }
    }
    It 'round-trips a 12 MB attachment through the data tree of a subnode, and a message of 60 recipients' {
        $path = New-PstPath 'big'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'T')
        $f = $w.CreateFolder($null, 'F', 'IPF.Note')
        $m = New-SyntheticMessage -AttachBytes 12MB
        foreach ($i in 4..60) {
            $r = [MailboxTransfer.FxRecipient]::new()
            $r.Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x0C15, 2))
            $r.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x3001, "Recipient $i"))
            $m.Recipients.Add($r)
        }
        $nid = $w.AddMessage($f, $m)
        $w.Close()
        $rd = [MailboxTransfer.PstReader]::Open($path)
        try { Assert-SameMessage $m $rd.ReadMessage($nid) } finally { $rd.Dispose() }
    }
    It 'reports a projected size that is at least what the file ends up with, to roll to a new PST before a limit' {
        $path = New-PstPath 'size'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'T')
        $f = $w.CreateFolder($null, 'F', 'IPF.Note')
        $last = 0
        foreach ($i in 1..300) {
            $m = New-SyntheticMessage -Subject "s$i" -AttachBytes 20000
            [MailboxTransfer.PstWriter]::EstimateSize($m) | Should -BeGreaterThan 20000
            [void]$w.AddMessage($f, $m)
            $w.ProjectedSize | Should -BeGreaterOrEqual $last
            $last = $w.ProjectedSize
        }
        $projected = $w.ProjectedSize
        $w.Close()
        (Get-Item $path).Length | Should -BeLessOrEqual ($projected + 253952)
        (Get-Item $path).Length | Should -BeGreaterThan 6000000
    }
}

Describe 'split by size' {
    It 'rolls to a new PST before the limit, recreates the folders in each file and loses no message' {
        $path = New-PstPath 'split'
        $limit = 300KB
        $split = [MailboxTransfer.PstSplitWriter]::new($path, 'Split', $limit)
        $total = 60
        foreach ($i in 1..$total) {
            $m = New-SyntheticMessage -Subject "m$i" -AttachBytes 20000
            [void]$split.AddMessage($(if ($i % 2) { 'Mail/Odd' } else { 'Mail/Even' }), 'IPF.Note', $m)
        }
        $split.Close()
        $split.Files.Count | Should -BeGreaterThan 3
        $split.MessagesWritten | Should -Be $total
        (Split-Path $split.Files[0] -Leaf) | Should -Be 'split.pst'
        (Split-Path $split.Files[1] -Leaf) | Should -Be 'split-0002.pst'
        $count = 0
        foreach ($f in $split.Files) {
            (Get-Item $f).Length | Should -BeLessOrEqual ($limit + 3 * 253952)
            $r = [MailboxTransfer.PstReader]::Open($f)
            try { foreach ($fo in $r.AllFolders()) { $count += $r.MessageNids($fo, $false).Count } } finally { $r.Dispose() }
        }
        $count | Should -Be $total
    }
    It 'writes a message larger than the limit alone in its own file' {
        $path = New-PstPath 'split-big'
        $split = [MailboxTransfer.PstSplitWriter]::new($path, 'Split', 100KB)
        [void]$split.AddMessage('A', 'IPF.Note', (New-SyntheticMessage -Subject 'small'))
        [void]$split.AddMessage('A', 'IPF.Note', (New-SyntheticMessage -Subject 'huge' -AttachBytes 1MB))
        [void]$split.AddMessage('A', 'IPF.Note', (New-SyntheticMessage -Subject 'small 2'))
        $split.Close()
        $split.Files.Count | Should -Be 3
    }
}

Describe 'real items (out\fxics-corpus2, git-ignored)' -Skip:(-not $script:HasCorpus2) {
    It 'writes the 21+ varied items to a PST and reads them back identical, property by property' {
        $path = New-PstPath 'corpus2'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Corpus2')
        $f = $w.CreateFolder($null, 'Items', 'IPF.Note')
        $nids = [ordered]@{}
        foreach ($file in Get-ChildItem -LiteralPath $script:Corpus2 -Filter '*.fxs' | Sort-Object Name) {
            $doc = [MailboxTransfer.FxCodec]::DecodeFile($file.FullName)
            $doc.Error | Should -BeNullOrEmpty
            $nids[$file.Name] = $w.AddMessage($f, $doc.Message)
        }
        $w.Close()
        $nids.Count | Should -BeGreaterThan 20
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            foreach ($k in $nids.Keys) {
                $orig = [MailboxTransfer.FxCodec]::DecodeFile((Join-Path $script:Corpus2 $k)).Message
                $copy = $r.ReadMessage([uint32]$nids[$k])
                $d = [Collections.Generic.List[string]]::new()
                [void][MailboxTransfer.PstFxCompare]::Compare($orig, $copy, $d)
                $d | Should -BeNullOrEmpty -Because $k
            }
        } finally { $r.Dispose() }
    }
}
