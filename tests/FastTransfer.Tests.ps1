#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the exact FastTransfer decoder, encoder and round trip (Engine.FxModel.cs, MS-OXCFXICS 2.2.4).
    The streams are written by FxBuilder (tests\Fixtures.cs) from the grammar of the specification, independently of the codec
    that reads them. Nothing here reaches the network.

    The last test runs over every stream of out\fxics-corpus (git-ignored: exported streams, see tools\Test-FastTransferCorpus.ps1)
    and is skipped when that folder does not exist.

    Run:  Invoke-Pester -Path .\tests\FastTransfer.Tests.ps1 -Output Detailed
#>

BeforeDiscovery {
    $dir = Join-Path (Split-Path $PSScriptRoot -Parent) 'out\fxics-corpus'
    $script:HasCorpus = (Test-Path -LiteralPath $dir) -and @(Get-ChildItem -LiteralPath $dir -Filter '*.fxs' -File).Count -gt 0
}

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    [Text.Encoding]::RegisterProvider([Text.CodePagesEncodingProvider]::Instance)
    $script:Corpus = Join-Path $script:RepoRoot 'out\fxics-corpus'

    function New-Builder { [MailboxTransfer.Tests.FxBuilder]::new() }

    # An array of byte arrays that PowerShell does not flatten.
    function Get-Jagged {
        $r = [byte[][]]::new($args.Count)
        for ($i = 0; $i -lt $args.Count; $i++) { $r[$i] = [byte[]]$args[$i] }
        return , $r
    }

    function Get-Utf16([string]$Text) { [Text.Encoding]::Unicode.GetBytes($Text + "`0") }

    # Decodes, checks that the document is clean (no error, no warning) and that it comes back byte for byte.
    function Assert-RoundTrip([byte[]]$Bytes, [string]$Because = '') {
        $rt = [MailboxTransfer.FxCodec]::RoundTrip($Bytes)
        $rt.Document.Error | Should -BeNullOrEmpty -Because $Because
        $rt.Document.Truncated | Should -BeFalse -Because $Because
        @($rt.Document.Warnings) | Should -BeNullOrEmpty -Because $Because
        $rt.Identical | Should -BeTrue -Because "$Because $($rt.Note) (first difference at $($rt.FirstDifference))"
        return $rt.Document
    }

    # A message with every structure of the grammar: wrapper, FXDelProp, recipients, a file attachment, an embedded message holding
    # an attachment that holds an embedded message (2 levels).
    function New-NestedMessage {
        param([switch]$Wrapper, [int]$BigBytes = 0)
        $cp = [Text.Encoding]::Unicode
        $b = New-Builder
        if ($Wrapper) { $null = $b.Marker(0x400C0003) }
        $null = $b.CodePageStr(0x001A, 0x84B0, 'IPM.Note', $cp).CodePageStr(0x0037, 0x84B0, 'Top', $cp).Int32(0x0E07, 1)
        $null = $b.NamedNameStr(0x8001, [guid]'11111111-2222-3333-4444-555555555555', 'Custom', 'value')
        if ($BigBytes -gt 0) { $null = $b.BinLarge(0x1013, $BigBytes) }
        $null = $b.DelProp(0x0E12000D)
        $null = $b.Marker(0x40030003).Int32(0x0C15, 1).Str(0x3001, 'Alice').Marker(0x40040003)
        $null = $b.Marker(0x40030003).Int32(0x0C15, 2).Str(0x3001, 'Bob').Marker(0x40040003)
        $null = $b.DelProp(0x0E13000D)
        $null = $b.Marker(0x40000003).Int32(0x0E21, 0).Str(0x3707, 'file.txt').Bin(0x3701, [byte[]](1..200)).Marker(0x400E0003)
        # attachment 1: embedded message (level 1) holding a recipient and an attachment that holds an embedded message (level 2)
        $null = $b.Marker(0x40000003).Int32(0x0E21, 1).Int32(0x3705, 5).Object(0x3701).Marker(0x40010003)
        $null = $b.Str(0x0037, 'Level 1').Marker(0x40030003).Str(0x3001, 'Carol').Marker(0x40040003)
        $null = $b.Marker(0x40000003).Int32(0x0E21, 0).Int32(0x3705, 5).Object(0x3701).Marker(0x40010003)
        $null = $b.Str(0x0037, 'Level 2').Bin(0x1009, [byte[]](9, 8, 7)).Marker(0x40020003).Marker(0x400E0003)
        $null = $b.Marker(0x40020003).Marker(0x400E0003)
        if ($Wrapper) { $null = $b.Marker(0x400D0003) }
        return $b.ToArray()
    }
}

Describe 'the defect of the real item: a multi-valued code page string' {
    It 'follows the stream past 0x94B0 (PtypMultipleString in code page 1200): the count was taken for a length' {
        # as on the real item: named property by name, type 0x94B0, count 1, one value of 80 bytes, then more properties
        $g = [guid]'0b63e350-9ccc-11d0-bcdb-00805fccce04'
        $b = New-Builder
        $null = $b.Int32(0x0017, 1).NamedLidInt32(0x805A, $g, 0x1234, -1)
        $null = $b.NamedNameTag(0x94B0, 0x805B, $g, 'ErrorTags').Raw([BitConverter]::GetBytes([uint32]1))
        $text = Get-Utf16 'parsererror_ParsingError_301002_unknown'
        $null = $b.Raw([BitConverter]::GetBytes([uint32]$text.Length)).Raw($text)
        $null = $b.CodePageStr(0x0037, 0x84B0, 'After', [Text.Encoding]::Unicode)
        $data = $b.ToArray()
        $read = [MailboxTransfer.FxReader]::Read([IO.MemoryStream]::new($data))
        $read.Error | Should -BeNullOrEmpty
        $read.Truncated | Should -BeFalse
        @($read.Elements).Count | Should -Be 4
        $tags = $read.Elements[2]
        $tags.TypeName | Should -Be 'CodePage(1200)Multiple'
        $tags.NamedString | Should -Be 'ErrorTags'
        $tags.Value[0] | Should -Be 'parsererror_ParsingError_301002_unknown'
        $read.Elements[3].Value | Should -Be 'After'
        $doc = Assert-RoundTrip $data
        $doc.Message.Properties[2].AsStrings()[0] | Should -Be 'parsererror_ParsingError_301002_unknown'
        $doc.Message.Properties[2].IsMultiValued | Should -BeTrue
    }
    It 'still reports a tag of type 0x0000 as unknown (the stream cannot be followed after it)' {
        $b = New-Builder
        $null = $b.Int32(0x0017, 1).Raw([byte[]](0, 0, 0, 0x70, 1, 2, 3, 4))
        $read = [MailboxTransfer.FxReader]::Read([IO.MemoryStream]::new($b.ToArray()))
        $read.Error | Should -BeLike 'Unknown property type 0x0000 (tag 0x70000000)*'
        $doc = [MailboxTransfer.FxCodec]::Decode($b.ToArray())
        $doc.Error | Should -BeLike 'Unknown property type 0x0000*'
    }
}

Describe 'lexical structure: every type' {
    It 'decodes and re-encodes every fixed-size type' {
        $t = [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)
        $g = [guid]'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
        $b = New-Builder
        $null = $b.Fixed(0x0001, 0x7001, [byte[]]@())
        $null = $b.Fixed(0x0002, 0x7002, [BitConverter]::GetBytes([int16]-2))
        $null = $b.Fixed(0x0003, 0x7003, [BitConverter]::GetBytes([int]-70000))
        $null = $b.Fixed(0x0004, 0x7004, [BitConverter]::GetBytes([single]1.5))
        $null = $b.Fixed(0x0005, 0x7005, [BitConverter]::GetBytes([double]-2.25))
        $null = $b.Fixed(0x0006, 0x7006, [BitConverter]::GetBytes([long]12345678))
        $null = $b.Fixed(0x0007, 0x7007, [BitConverter]::GetBytes([double]45000.5))
        $null = $b.Fixed(0x000A, 0x7008, [BitConverter]::GetBytes([uint32]2147746055))
        $null = $b.Fixed(0x000B, 0x7009, [byte[]](1, 0))
        $null = $b.Fixed(0x000B, 0x700A, [byte[]](0, 0))
        $null = $b.Fixed(0x0014, 0x700B, [BitConverter]::GetBytes([long]-5000000000))
        $null = $b.Fixed(0x0040, 0x700C, [BitConverter]::GetBytes($t.ToFileTimeUtc()))
        $null = $b.Fixed(0x0048, 0x700D, $g.ToByteArray())
        $data = $b.ToArray()
        $doc = Assert-RoundTrip $data
        $p = $doc.Message.Properties
        $p.Count | Should -Be 13
        $p[0].Value | Should -BeNullOrEmpty
        $p[1].AsInt16() | Should -Be -2
        $p[2].AsInt32() | Should -Be -70000
        $p[3].Value | Should -Be 1.5
        $p[4].AsDouble() | Should -Be -2.25
        $p[5].Value | Should -Be 1234.5678
        $p[6].Value | Should -Be 45000.5
        $p[7].Value | Should -Be ([uint32]2147746055)
        $p[8].AsBool() | Should -BeTrue
        $p[9].AsBool() | Should -BeFalse
        $p[10].AsInt64() | Should -Be -5000000000
        $p[11].AsTime().UtcDateTime | Should -Be $t
        $p[12].AsGuid() | Should -Be $g
        $p[8].Raw.Length | Should -Be 2 -Because 'a PtypBoolean is 2 bytes in a FastTransfer stream'
        { $p[2].AsString() } | Should -Throw '*is not a string*'
        # the flat reader agrees
        $read = [MailboxTransfer.FxReader]::Read([IO.MemoryStream]::new($data))
        $read.Error | Should -BeNullOrEmpty
        @($read.Elements).Count | Should -Be 13
    }
    It 'decodes and re-encodes the variable-size types (String8, String, Binary, ServerId, Object with a length, empty values)' {
        $b = New-Builder
        $null = $b.Variable(0x001E, 0x7101, [byte[]](0x63, 0x61, 0x66, 0xE9, 0))
        $null = $b.Variable(0x001F, 0x7102, (Get-Utf16 'Café ☺'))
        $null = $b.Variable(0x0102, 0x7103, [byte[]](0..255))
        $null = $b.Variable(0x00FB, 0x7104, [byte[]](1, 2, 3))
        $null = $b.Variable(0x0102, 0x7105, [byte[]]@())
        $null = $b.Variable(0x000D, 0x7106, [byte[]](5, 6))
        $doc = Assert-RoundTrip $b.ToArray()
        $p = $doc.Message.Properties
        $p[0].AsString() | Should -Be 'café' -Because 'PtypString8 is read in code page 1252'
        $p[1].AsString() | Should -Be 'Café ☺'
        $p[1].Raw.Length | Should -Be (2 * 7) -Because 'the terminator is kept in the raw value'
        $p[2].AsBytes().Length | Should -Be 256
        $p[3].AsBytes() | Should -Be @(1, 2, 3)
        $p[4].AsBytes().Length | Should -Be 0
        $p[5].TypeInfo.IsObject | Should -BeTrue
        $p[5].BareObject | Should -BeFalse
        $p[5].Raw | Should -Be @(5, 6)
    }
    It 'decodes and re-encodes every multi-valued type, including an empty one and empty items' {
        $g1 = [guid]'11111111-1111-1111-1111-111111111111'; $g2 = [guid]'22222222-2222-2222-2222-222222222222'
        $t = [datetime]::new(2026, 1, 2, 3, 4, 5, [DateTimeKind]::Utc)
        $b = New-Builder
        $null = $b.Multi(0x1002, 0x7201, (Get-Jagged ([BitConverter]::GetBytes([int16]1)) ([BitConverter]::GetBytes([int16]-1))), $false)
        $null = $b.Multi(0x1003, 0x7202, (Get-Jagged ([BitConverter]::GetBytes([int]7)) ([BitConverter]::GetBytes([int]8)) ([BitConverter]::GetBytes([int]9))), $false)
        $null = $b.Multi(0x1004, 0x7203, (Get-Jagged ([BitConverter]::GetBytes([single]0.5))), $false)
        $null = $b.Multi(0x1005, 0x7204, (Get-Jagged ([BitConverter]::GetBytes([double]0.25)) ([BitConverter]::GetBytes([double]4))), $false)
        $null = $b.Multi(0x1006, 0x7205, (Get-Jagged ([BitConverter]::GetBytes([long]10000))), $false)
        $null = $b.Multi(0x1007, 0x7206, (Get-Jagged ([BitConverter]::GetBytes([double]1.5))), $false)
        $null = $b.Multi(0x1014, 0x7207, (Get-Jagged ([BitConverter]::GetBytes([long]1)) ([BitConverter]::GetBytes([long]2))), $false)
        $null = $b.Multi(0x1040, 0x7208, (Get-Jagged ([BitConverter]::GetBytes($t.ToFileTimeUtc()))), $false)
        $null = $b.Multi(0x1048, 0x7209, (Get-Jagged $g1.ToByteArray() $g2.ToByteArray()), $false)
        $null = $b.Multi(0x101E, 0x720A, (Get-Jagged ([byte[]](0x61, 0xE9, 0)) ([byte[]](0x62, 0))), $true)
        $null = $b.Multi(0x101F, 0x720B, (Get-Jagged (Get-Utf16 'un') (Get-Utf16 'deux') (Get-Utf16 '')), $true)
        $null = $b.Multi(0x1102, 0x720C, (Get-Jagged ([byte[]](1, 2, 3)) ([byte[]]@()) ([byte[]](4))), $true)
        $null = $b.Multi(0x1003, 0x720D, (Get-Jagged), $false)
        $doc = Assert-RoundTrip $b.ToArray()
        $p = $doc.Message.Properties
        $p.Count | Should -Be 13
        $p[0].Value | Should -Be @(1, -1)
        $p[1].Items.Count | Should -Be 3
        $p[3].Value | Should -Be @(0.25, 4)
        $p[4].Value | Should -Be @([decimal]1)
        $p[6].Value | Should -Be @(1, 2)
        $p[7].Value[0].UtcDateTime | Should -Be $t
        $p[8].Value | Should -Be @($g1, $g2)
        $p[9].AsStrings() | Should -Be @('aé', 'b')
        $p[10].AsStrings() | Should -Be @('un', 'deux', '')
        $p[11].Items[1].Length | Should -Be 0
        $p[11].ValueLength | Should -Be 4
        $p[12].Items.Count | Should -Be 0
        $p[12].IsMultiValued | Should -BeTrue
    }
    It 'decodes code page strings: 1200 (0x84B0), 1252 (0x84E4), UTF-8 (0xFDE9), KOI8-R (0xD182 = 20866 | 0x8000), and their multi-valued forms (0x94B0, 0x94E4)' {
        $koi = [Text.Encoding]::GetEncoding(20866)
        $b = New-Builder
        $null = $b.Variable(0x84B0, 0x7301, (Get-Utf16 'Unicode'))
        $null = $b.Variable(0x84E4, 0x7302, [byte[]](0x63, 0x61, 0x66, 0xE9, 0))
        $null = $b.Variable(0xFDE9, 0x7303, [Text.Encoding]::UTF8.GetBytes("Zo$([char]0xEB)`0"))
        $null = $b.Variable(0xD182, 0x7304, $koi.GetBytes("$([char]0x041F)$([char]0x0440)$([char]0x0438)`0"))
        $null = $b.Multi(0x94B0, 0x7305, (Get-Jagged (Get-Utf16 'a') (Get-Utf16 'bc')), $true)
        $null = $b.Multi(0x94E4, 0x7306, (Get-Jagged ([byte[]](0x63, 0xE9, 0))), $true)
        $doc = Assert-RoundTrip $b.ToArray()
        $p = $doc.Message.Properties
        $p[0].TypeInfo.CodePage | Should -Be 1200
        $p[0].AsString() | Should -Be 'Unicode'
        $p[1].TypeInfo.CodePage | Should -Be 1252
        $p[1].AsString() | Should -Be 'café'
        $p[2].TypeInfo.CodePage | Should -Be 65001
        $p[2].AsString() | Should -Be "Zo$([char]0xEB)"
        $p[3].TypeInfo.CodePage | Should -Be 20866
        $p[3].TypeInfo.MultiValued | Should -BeFalse -Because 'KOI8-R has the bit 0x1000 itself: 0xD182 is a code page as it stands'
        $p[3].AsString() | Should -Be "$([char]0x041F)$([char]0x0440)$([char]0x0438)"
        $p[4].TypeInfo.MultiValued | Should -BeTrue
        $p[4].TypeInfo.CodePage | Should -Be 1200
        $p[4].AsStrings() | Should -Be @('a', 'bc')
        $p[5].AsStrings() | Should -Be @('cé')
        [MailboxTransfer.FxNames]::TypeName(0x84B0) | Should -Be 'CodePage(1200)'
        [MailboxTransfer.FxNames]::TypeName(0x94B0) | Should -Be 'CodePage(1200)Multiple'
    }
    It 'decodes named properties by LID and by name, single and multi-valued, with the property set' {
        $g = [guid]'00020329-0000-0000-c000-000000000046'
        $b = New-Builder
        $null = $b.NamedLidTag(0x0003, 0x8001, $g, 0x8535).Raw([BitConverter]::GetBytes([int]42))
        $null = $b.NamedNameTag(0x001F, 0x8002, $g, 'Keywords-é').Raw([BitConverter]::GetBytes([uint32]10)).Raw((Get-Utf16 'blue'))
        $null = $b.NamedNameTag(0x101F, 0x8003, $g, 'Categories').Raw([BitConverter]::GetBytes([uint32]2))
        $null = $b.Raw([BitConverter]::GetBytes([uint32]4)).Raw((Get-Utf16 'a')).Raw([BitConverter]::GetBytes([uint32]4)).Raw((Get-Utf16 'b'))
        $null = $b.NamedLidTag(0x0102, 0xFFFE, [guid]::Empty, [uint32]::MaxValue).Raw([BitConverter]::GetBytes([uint32]1)).Raw([byte[]](7))
        $doc = Assert-RoundTrip $b.ToArray()
        $p = $doc.Message.Properties
        $p[0].IsNamed | Should -BeTrue
        $p[0].NamedGuid | Should -Be $g
        $p[0].NamedLid | Should -Be 0x8535
        $p[0].AsInt32() | Should -Be 42
        $p[1].NamedName | Should -Be 'Keywords-é'
        $p[1].NamedLid | Should -BeNullOrEmpty
        $p[1].AsString() | Should -Be 'blue'
        $p[2].NamedName | Should -Be 'Categories'
        $p[2].AsStrings() | Should -Be @('a', 'b')
        $p[3].NamedLid | Should -Be ([uint32]::MaxValue)
        $doc.Message.Find($g, [uint32]0x8535).AsInt32() | Should -Be 42
        $doc.Message.Find($g, 'Categories').AsStrings().Count | Should -Be 2
    }
    It 'keeps the extended types (bit 0x4000 and two bytes) and both forms of PtypObject' {
        $b = New-Builder
        $null = $b.ExtTag(0x0102, 0x010B, 0x0200).Raw([BitConverter]::GetBytes([uint32]3)).Raw([byte[]](1, 2, 3))
        $null = $b.ExtTag(0x0040, 0x0FE3, 0x0200).Raw([BitConverter]::GetBytes([long]131000000000000000))
        $null = $b.ExtTag(0x0002, 0x1714, 0x0007).Raw([byte[]](2, 0))
        $null = $b.Marker(0x40000003).Object(0x3701).Marker(0x40010003).Str(0x0037, 'Inner').Marker(0x40020003).Marker(0x400E0003)
        $null = $b.Marker(0x40000003).Variable(0x000D, 0x3701, [byte[]](1, 2, 3, 4)).Marker(0x400E0003)
        $doc = Assert-RoundTrip $b.ToArray()
        $p = $doc.Message.Properties
        $p[0].Extended | Should -BeTrue
        $p[0].Type | Should -Be 0x0102 -Because 'the type is shown without the bit 0x4000'
        $p[0].ExtendedFlags | Should -Be 0x0200
        $p[2].ExtendedFlags | Should -Be 7
        $p[2].AsInt16() | Should -Be 2
        $doc.Message.Attachments.Count | Should -Be 2
        $doc.Message.Attachments[0].Properties[0].BareObject | Should -BeTrue
        $doc.Message.Attachments[0].Embedded.GetString(0x0037) | Should -Be 'Inner'
        $doc.Message.Attachments[1].Properties[0].BareObject | Should -BeFalse
        $doc.Message.Attachments[1].Properties[0].Raw | Should -Be @(1, 2, 3, 4)
        $doc.Message.Attachments[1].Embedded | Should -BeNullOrEmpty
    }
    It 'refuses an invalid named property and an unknown multi-valued type, with the offset' {
        $g = [guid]'00020329-0000-0000-c000-000000000046'
        $b = New-Builder
        $null = $b.Int32(0x0017, 1).TagOnly(0x0003, 0x8001).Raw($g.ToByteArray()).Raw([byte[]](9, 0, 0, 0, 0))
        $doc = [MailboxTransfer.FxCodec]::Decode($b.ToArray())
        $doc.Error | Should -BeLike 'Unknown kind 9 of named property*'
        $doc.ErrorOffset | Should -Be 8
        $doc.Message.Properties.Count | Should -Be 1
        $doc2 = [MailboxTransfer.FxCodec]::Decode((New-Builder).Raw([byte[]](0x77, 0x17, 1, 0x30, 1, 0, 0, 0)).ToArray())
        $doc2.Error | Should -BeLike 'Unknown multi-valued type 0x1777*'
    }
}

Describe 'syntactical structure: message, recipients, attachments, embedded messages' {
    It 'builds the model of a message with recipients, attachments and two levels of embedded messages' {
        $data = New-NestedMessage
        $doc = Assert-RoundTrip $data
        $m = $doc.Message
        $m.StartMarker | Should -Be 0
        $m.GetString(0x0037) | Should -Be 'Top'
        $m.GetString(0x001A) | Should -Be 'IPM.Note'
        $m.Find([guid]'11111111-2222-3333-4444-555555555555', 'Custom').AsString() | Should -Be 'value'
        $m.DelPropRecipients.AsInt32() | Should -Be 0x0E12000D
        $m.DelPropAttachments.AsInt32() | Should -Be 0x0E13000D
        $m.Recipients.Count | Should -Be 2
        $m.Recipients[1].Properties[1].AsString() | Should -Be 'Bob'
        $m.Attachments.Count | Should -Be 2
        $m.Attachments[0].Embedded | Should -BeNullOrEmpty
        $m.Attachments[0].Properties[1].AsString() | Should -Be 'file.txt'
        $l1 = $m.Attachments[1].Embedded
        $l1.GetString(0x0037) | Should -Be 'Level 1'
        $l1.Recipients[0].Properties[0].AsString() | Should -Be 'Carol'
        $l1.StartMarker | Should -Be 0
        $l2 = $l1.Attachments[0].Embedded
        $l2.GetString(0x0037) | Should -Be 'Level 2'
        $l2.Find(0x1009).AsBytes() | Should -Be @(9, 8, 7)
        $m.EmbeddedDepth() | Should -Be 2
        $l1.EmbeddedDepth() | Should -Be 1
        $l2.EmbeddedDepth() | Should -Be 0
        $flat = [MailboxTransfer.FxReader]::Read([IO.MemoryStream]::new($data))
        $m.CountProperties() | Should -Be ($flat.PropertyCount + @($flat.Elements | Where-Object { $_.IsMarker -and $_.Name -eq 'FXDelProp' }).Count) -Because 'the flat reader shows FXDelProp as a marker'
        $doc.Envelope | Should -BeNullOrEmpty
        $doc.PayloadBytes | Should -Be $data.Length
    }
    It 'keeps the wrapper StartMessage ... EndMessage when there is one' {
        $doc = Assert-RoundTrip (New-NestedMessage -Wrapper)
        $doc.Message.StartMarker | Should -Be 0x400C0003
        $doc.Message.Recipients.Count | Should -Be 2
        $b = New-Builder
        $null = $b.Marker(0x40100003).Str(0x0037, 'FAI').Marker(0x400D0003)
        $fai = Assert-RoundTrip $b.ToArray()
        $fai.Message.StartMarker | Should -Be 0x40100003
    }
    It 'agrees with the flat reader and the summary on the same stream' {
        $data = New-NestedMessage
        $read = [MailboxTransfer.FxReader]::Read([IO.MemoryStream]::new($data))
        $read.Error | Should -BeNullOrEmpty
        $s = [MailboxTransfer.FxMessage]::Summarize($read)
        $s.Subject | Should -Be 'Top'
        $s.Recipients.Count | Should -Be 2
        $s.Attachments.Count | Should -Be 2
        $s.Attachments[1].EmbeddedSubject | Should -Be 'Level 1'
    }
    It 'handles values that cross the boundaries of the records of the envelope, whatever their size' {
        $data = New-NestedMessage -Wrapper -BigBytes 100000
        foreach ($chunk in 1, 2, 5, 13, 100, 4096, 31680, 99999, 200000) {
            $file = [MailboxTransfer.Tests.FxBuilder]::Envelope($data, $chunk, $true)
            $doc = Assert-RoundTrip $file "chunk $chunk"
            $doc.Envelope.DataLength | Should -Be $data.Length
            $doc.Envelope.DataRecords | Should -Be ([math]::Ceiling($data.Length / $chunk))
            $doc.Message.Find(0x1013).Raw.Length | Should -Be 100000
            $doc.Message.Attachments[1].Embedded.Attachments[0].Embedded.GetString(0x0037) | Should -Be 'Level 2'
        }
        # the envelope records and their payloads come back as they were (the GUID record of type 3 included)
        $doc = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.Tests.FxBuilder]::Envelope($data, 31680, $true))
        (@($doc.Envelope.Entries) | ForEach-Object Type | Select-Object -First 3) | Should -Be @(1, 3, 4)
        $doc.Envelope.Entries[1].Payload.Length | Should -Be 24
    }
    It 'cuts the stream again in records of 31,680 bytes when the message changed size' {
        $data = New-NestedMessage -BigBytes 100000
        $doc = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.Tests.FxBuilder]::Envelope($data, 5000, $false))
        $doc.Message.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x1000, ('x' * 90000)))
        $bytes = [MailboxTransfer.FxCodec]::Encode($doc)
        $again = [MailboxTransfer.FxCodec]::Decode($bytes)
        $again.Error | Should -BeNullOrEmpty
        @($again.Warnings) | Should -BeNullOrEmpty
        $again.Message.GetString(0x1000).Length | Should -Be 90000
        $sizes = @($again.Envelope.Entries | Where-Object IsData | ForEach-Object Length)
        ($sizes | Select-Object -SkipLast 1 | Select-Object -Unique) | Should -Be 31680
        ($sizes | Measure-Object -Sum).Sum | Should -Be $again.PayloadBytes
        $read = [MailboxTransfer.FxReader]::Read([IO.MemoryStream]::new($bytes))
        $read.Error | Should -BeNullOrEmpty
        @($read.Elements | Where-Object Id -EQ 0x1000).Count | Should -Be 1
    }
    It 'tells what is not canonical and encodes it in the canonical order (property after the recipients, marker out of place, bytes after the end)' {
        $b = New-Builder
        $null = $b.Str(0x0037, 'S').Marker(0x40030003).Str(0x3001, 'R').Marker(0x40040003).Str(0x001A, 'IPM.Note').Marker(0x40000003).Marker(0x400E0003)
        $data = $b.ToArray()
        $rt = [MailboxTransfer.FxCodec]::RoundTrip($data)
        $rt.Identical | Should -BeFalse
        $rt.Document.Warnings[0] | Should -BeLike '*after the recipients*'
        $rt.FirstDifference | Should -BeGreaterThan -1
        $rt.Document.Message.Properties.Count | Should -Be 2 -Because 'the property is kept, moved before the recipients'
        $rt.Note | Should -BeLike '*after the recipients*'

        $b = New-Builder
        $null = $b.Str(0x0037, 'S').Marker(0x40090003).Str(0x001A, 'IPM.Note')
        $doc = [MailboxTransfer.FxCodec]::Decode($b.ToArray())
        $doc.Warnings[0] | Should -BeLike '*StartTopFld*not allowed*'
        $doc.Message.Properties.Count | Should -Be 2

        $b = New-Builder
        $null = $b.Marker(0x400C0003).Str(0x0037, 'S').Marker(0x400D0003).Str(0x001A, 'after')
        $doc = [MailboxTransfer.FxCodec]::Decode($b.ToArray())
        $doc.Warnings[0] | Should -BeLike '*after the end of the message*'

        $b = New-Builder
        $null = $b.Marker(0x40000003).Str(0x3707, 'a')
        $doc = [MailboxTransfer.FxCodec]::Decode($b.ToArray())
        $doc.Warnings[0] | Should -BeLike '*attachment not closed*'
        $doc.Message.Attachments.Count | Should -Be 1
    }
    It 'detects a difference at the right offset (value changed in the model, stream cut)' {
        $data = New-NestedMessage
        $doc = [MailboxTransfer.FxCodec]::Decode($data)
        [MailboxTransfer.FxCodec]::Encode($doc) | Should -Be $data
        $p = $doc.Message.Find(0x0E07)
        $p.Raw[0] = [byte]($p.Raw[0] -bxor 1)
        $changed = [MailboxTransfer.FxCodec]::Encode($doc)
        $diff = 0; while ($changed[$diff] -eq $data[$diff]) { $diff++ }
        $diff | Should -Be ($p.Offset + 4) -Because 'the value follows the 4 bytes of the tag'
        $trimmed = [byte[]]$data[0..($data.Length - 2)]
        $rt = [MailboxTransfer.FxCodec]::RoundTrip($trimmed)
        $rt.Identical | Should -BeFalse
        $rt.Document.Truncated | Should -BeTrue
        $rt.FirstDifference | Should -BeGreaterThan -1
    }
}

Describe 'encoder: a message built in memory' {
    It 'writes exactly what an independent writer writes for the same properties' {
        $t = [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)
        $g = [guid]'00020329-0000-0000-c000-000000000046'
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x0017, 1))
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x0037, 'Hello'))
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromBool(0x0023, $false))
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromTime(0x0039, $t))
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromBinary(0x003B, [byte[]](1, 2, 3)))
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromInt64(0x674A, 99))
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromStrings(0x7001, [string[]]@('a', 'b')))
        $m.Properties.Add(([MailboxTransfer.FxProperty]::FromInt32(0x8001, 7)).Named($g, [uint32]0x8535))
        $m.Properties.Add(([MailboxTransfer.FxProperty]::FromString(0x8002, 'v')).Named($g, 'Name'))
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromCodePageString(0x001A, 1200, 'IPM.Note'))
        $r = [MailboxTransfer.FxRecipient]::new()
        $r.Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x0C15, 1))
        $r.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x3001, 'Alice'))
        $m.Recipients.Add($r)
        $a = [MailboxTransfer.FxAttachment]::new()
        $a.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x3707, 'a.txt'))
        $m.Attachments.Add($a)

        $expected = (New-Builder)
        $null = $expected.Int32(0x0017, 1).Str(0x0037, 'Hello').Bool(0x0023, $false).Time(0x0039, $t).Bin(0x003B, [byte[]](1, 2, 3)).Int64(0x674A, 99)
        $null = $expected.MvStr(0x7001, [string[]]@('a', 'b')).NamedLidInt32(0x8001, $g, 0x8535, 7).NamedNameStr(0x8002, $g, 'Name', 'v')
        $null = $expected.CodePageStr(0x001A, 0x84B0, 'IPM.Note', [Text.Encoding]::Unicode)
        $null = $expected.Marker(0x40030003).Int32(0x0C15, 1).Str(0x3001, 'Alice').Marker(0x40040003)
        $null = $expected.Marker(0x40000003).Str(0x3707, 'a.txt').Marker(0x400E0003)
        $bytes = [MailboxTransfer.FxCodec]::EncodePayload($m)
        $bytes | Should -Be $expected.ToArray()
        $back = Assert-RoundTrip $bytes
        $back.Message.GetString(0x0037) | Should -Be 'Hello'
        $back.Message.Find(0x001A).TypeInfo.CodePage | Should -Be 1200
    }
    It 'refuses a value that does not fit its type, naming the property' {
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add([MailboxTransfer.FxProperty]::Create(0x7001, 0x0003, [byte[]](1, 2)))
        { [MailboxTransfer.FxCodec]::EncodePayload($m) } | Should -Throw '*0x70010003*the value has 2 bytes, the type needs 4*'
        $m2 = [MailboxTransfer.FxMessage]::new()
        $m2.Properties.Add([MailboxTransfer.FxProperty]::Create(0x8001, 0x0003, [byte[]](1, 0, 0, 0)))
        { [MailboxTransfer.FxCodec]::EncodePayload($m2) } | Should -Throw '*without a property set*'
    }
}

Describe 'robustness of the exact decoder' {
    It 'never throws on a stream cut at any length, with or without the envelope, and says the stream is incomplete' {
        $data = New-NestedMessage -Wrapper
        $env = [MailboxTransfer.Tests.FxBuilder]::Envelope($data, 100, $true)
        foreach ($whole in $data, $env) {
            foreach ($len in 0..($whole.Length - 1)) {
                $cut = [byte[]]::new($len); [Array]::Copy($whole, $cut, $len)
                $doc = $null; $err = $null
                try { $doc = [MailboxTransfer.FxCodec]::Decode($cut) } catch { $err = $_ }
                if ($err) { throw "length ${len}: $err" }
                if ($len -ge 8 -and -not ($doc.Truncated -or $doc.Error -or $doc.Warnings.Count -gt 0)) { throw "length ${len}: a cut stream is reported as complete" }
            }
        }
    }
    It 'never throws on garbage, nor allocates for an absurd length or count' {
        $rng = [Random]::new(11)
        foreach ($i in 1..300) {
            $junk = [byte[]]::new($rng.Next(0, 600)); $rng.NextBytes($junk)
            { $null = [MailboxTransfer.FxCodec]::Decode($junk) } | Should -Not -Throw
        }
        $doc = [MailboxTransfer.FxCodec]::Decode([byte[]](0x1F, 0x00, 0x01, 0x30, 0xFF, 0xFF, 0xFF, 0xFF, 1, 2, 3))
        $doc.Truncated | Should -BeTrue
        $doc2 = [MailboxTransfer.FxCodec]::Decode([byte[]](0x03, 0x10, 0x01, 0x70, 0xFF, 0xFF, 0xFF, 0x7F))
        $doc2.Error | Should -BeLike '*absurd count*'
    }
}

Describe 'corpus of real streams (out\fxics-corpus, git-ignored)' {
    It 'decodes every stream without warning and encodes it back byte for byte' -Skip:(-not $script:HasCorpus) {
        $failures = foreach ($f in Get-ChildItem -LiteralPath $script:Corpus -Filter '*.fxs' -File) {
            $rt = [MailboxTransfer.FxCodec]::RoundTripFile($f.FullName)
            $d = $rt.Document
            if (-not $rt.Identical -or $d.Warnings.Count -gt 0 -or $d.Error -or $d.Truncated) {
                "$($f.Name): first difference $($rt.FirstDifference), error '$($d.Error)', truncated $($d.Truncated), warnings $($d.Warnings.Count) $($d.Warnings | Select-Object -First 1), $($rt.Note)"
            }
        }
        @($failures) | Should -BeNullOrEmpty
    }
    It 'gives the flat reader the same count of properties as the model' -Skip:(-not $script:HasCorpus) {
        foreach ($f in Get-ChildItem -LiteralPath $script:Corpus -Filter '*.fxs' -File | Select-Object -First 4) {
            $read = [MailboxTransfer.FxReader]::ReadFile($f.FullName)
            $read.Error | Should -BeNullOrEmpty -Because $f.Name
            $doc = [MailboxTransfer.FxCodec]::DecodeFile($f.FullName)
            $doc.Message.CountProperties() | Should -Be ($read.PropertyCount + @($read.Elements | Where-Object { $_.IsMarker -and $_.Name -eq 'FXDelProp' }).Count) -Because $f.Name
        }
    }
}
