#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the streaming code: values of a FastTransfer stream kept as a source instead of a byte[]
    (Engine.FxStream.cs: FxBlob, FxPayload, FxMessageSource), data trees of a PST written and read block by block
    (PstNdbWriter.WriteStream, PstDataBlob), and the memory used to move items of up to 1 GiB (tools\Measure-PstMemory.ps1 run in a
    child process: growth of the working set, with a generous limit). Nothing here reaches the network.

    The memory test writes about 3.3 GB in the temporary directory for a while; MBT_MEMORY_TEST_MIB changes the size of the attachment
    (default 1024), 0 skips it.

    Run:  Invoke-Pester -Path .\tests\PstStreaming.Tests.ps1 -Output Detailed
#>

BeforeDiscovery {
    $dir = Join-Path (Split-Path $PSScriptRoot -Parent) 'out\fxics-corpus'
    $script:HasCorpus = (Test-Path -LiteralPath $dir) -and @(Get-ChildItem -LiteralPath $dir -Filter '*.fxs' -File).Count -gt 0
    $script:MemoryMiB = if ($env:MBT_MEMORY_TEST_MIB) { [int]$env:MBT_MEMORY_TEST_MIB } else { 1024 }
}

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    $P = [MailboxTransfer.FxProperty]
    $script:P = $P

    function New-Pattern { param([long]$Length, [uint64]$Seed = 7) [MailboxTransfer.TestSupport.PatternBlob]::new($Length, $Seed) }

    function Get-Bytes { param($Blob) $ms = [IO.MemoryStream]::new(); $Blob.CopyTo($ms); , $ms.ToArray() }

    # a message with a big binary attachment and a big string body, both lazy (nothing of them in memory)
    function New-LazyMessage {
        param([long]$AttachBytes = 300000, [long]$BodyBytes = 200000, [uint64]$Seed = 7)
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($P::FromString(0x001A, 'IPM.Note'))
        $m.Properties.Add($P::FromString(0x0037, 'Lazy'))
        $m.Properties.Add($P::FromInt32(0x0E07, 1))
        $body = [MailboxTransfer.FxConcatBlob]::new((New-Pattern $BodyBytes ($Seed + 1)), [MailboxTransfer.FxBytesBlob]::Terminator)
        $m.Properties.Add($P::CreateLazy(0x1000, 0x001F, $body))
        $r = [MailboxTransfer.FxRecipient]::new()
        $r.Properties.Add($P::FromInt32(0x0C15, 1)); $r.Properties.Add($P::FromString(0x3001, 'Bob'))
        $m.Recipients.Add($r)
        $a = [MailboxTransfer.FxAttachment]::new()
        $a.Properties.Add($P::FromInt32(0x3705, 1)); $a.Properties.Add($P::FromString(0x3704, 'big.bin'))
        $a.Properties.Add($P::CreateLazy(0x3701, 0x0102, (New-Pattern $AttachBytes $Seed)))
        $m.Attachments.Add($a)
        $m
    }

    function New-GraphEnvelope {
        $env = [MailboxTransfer.FxEnvelope]::new()
        foreach ($rec in @(@(1, 8), @(4, 6), @(2, 1))) {
            $e = [MailboxTransfer.FxEnvelopeEntry]::new(); $e.Type = $rec[0]; $e.Length = $rec[1]
            if ($rec[0] -ne 2) { $e.Payload = [byte[]]::new($rec[1]) }
            $env.Entries.Add($e)
        }
        , $env
    }

    function Read-Chunked { param($Stream, [int]$Size) $ms = [IO.MemoryStream]::new(); $buf = [byte[]]::new($Size); while (($n = $Stream.Read($buf, 0, $Size)) -gt 0) { $ms.Write($buf, 0, $n) }; , $ms.ToArray() }

    function New-PstPath { param([string]$Name) Join-Path $TestDrive ($Name + '.pst') }
}

Describe 'lazy values of a decoded FastTransfer stream' {
    It 'keeps a value above the threshold as a source and everything else in memory' {
        $m = New-LazyMessage -AttachBytes 100000 -BodyBytes 70000
        $bytes = [MailboxTransfer.FxCodec]::EncodePayload($m)
        $doc = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromBytes($bytes))
        $doc.Ok | Should -BeTrue
        $att = $doc.Message.Attachments[0].Properties | Where-Object Id -EQ 0x3701
        $att.IsLazy | Should -BeTrue
        $att.RawLength | Should -Be 100000
        $att.Lazy.Length | Should -Be 100000
        ($doc.Message.Find(0x1000)).IsLazy | Should -BeTrue
        ($doc.Message.Find(0x0037)).IsLazy | Should -BeFalse
        [MailboxTransfer.TestSupport.Streams]::Equal($att.Blob, (New-Pattern 100000)) | Should -BeTrue
        # Raw still works for a lazy property (it reads the value)
        [Linq.Enumerable]::SequenceEqual([byte[]]$att.Raw, [byte[]](Get-Bytes (New-Pattern 100000))) | Should -BeTrue
    }
    It 'keeps every value when the threshold is raised or when the decoder is the one of a byte array or a stream' {
        $bytes = [MailboxTransfer.FxCodec]::EncodePayload((New-LazyMessage -AttachBytes 100000 -BodyBytes 10))
        $o = [MailboxTransfer.FxDecodeOptions]::new(); $o.LazyAbove = 1GB
        $d1 = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromBytes($bytes), $o)
        ($d1.Message.Attachments[0].Properties | Where-Object Id -EQ 0x3701).IsLazy | Should -BeFalse
        $d2 = [MailboxTransfer.FxCodec]::Decode($bytes)
        ($d2.Message.Attachments[0].Properties | Where-Object Id -EQ 0x3701).IsLazy | Should -BeFalse
    }
    It 'finds a value that straddles the records of the Graph envelope (31,680 bytes), whatever the offset asked' {
        $m = New-LazyMessage -AttachBytes 200000 -BodyBytes 100000
        $doc0 = [MailboxTransfer.FxDocument]::new(); $doc0.Message = $m; $doc0.Envelope = New-GraphEnvelope
        $file = Join-Path $TestDrive 'env.fxs'
        [IO.File]::WriteAllBytes($file, [MailboxTransfer.FxCodec]::Encode($doc0))
        ([IO.File]::ReadAllBytes($file))[0..7] | Should -Be @(1, 0, 0, 0, 8, 0, 0, 0)
        $doc = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromFile($file))
        $doc.Ok | Should -BeTrue
        $doc.Envelope.DataRecords | Should -BeGreaterThan 8
        $att = ($doc.Message.Attachments[0].Properties | Where-Object Id -EQ 0x3701).Lazy
        $want = Get-Bytes (New-Pattern 200000)
        [Linq.Enumerable]::SequenceEqual([byte[]](Get-Bytes $att), [byte[]]$want) | Should -BeTrue
        foreach ($off in 0, 1, 31679, 31680, 31681, 63359, 63360, 100000, 199990) {
            $n = [Math]::Min(50, 200000 - $off)
            [Linq.Enumerable]::SequenceEqual([byte[]]$att.ReadRange($off, $n), [byte[]]$want[$off..($off + $n - 1)]) | Should -BeTrue -Because "offset $off"
        }
        $att.Slice(31000, 2000).Length | Should -Be 2000
        [Linq.Enumerable]::SequenceEqual([byte[]](Get-Bytes $att.Slice(31000, 2000)), [byte[]]$want[31000..32999]) | Should -BeTrue
    }
    It 'round-trips a file byte for byte with its values left in the file (envelope, bare stream, records of any length)' {
        $m = New-LazyMessage -AttachBytes 150000 -BodyBytes 90000
        $bare = Join-Path $TestDrive 'bare.fxs'
        [IO.File]::WriteAllBytes($bare, [MailboxTransfer.FxCodec]::EncodePayload($m))
        $rt = [MailboxTransfer.FxCodec]::RoundTripFile($bare)
        $rt.Identical | Should -BeTrue
        ($rt.Document.Message.Attachments[0].Properties | Where-Object Id -EQ 0x3701).IsLazy | Should -BeTrue
        $doc0 = [MailboxTransfer.FxDocument]::new(); $doc0.Message = $m; $doc0.Envelope = New-GraphEnvelope
        $env = Join-Path $TestDrive 'env2.fxs'
        [IO.File]::WriteAllBytes($env, [MailboxTransfer.FxCodec]::Encode($doc0))
        $rt2 = [MailboxTransfer.FxCodec]::RoundTripFile($env)
        $rt2.Identical | Should -BeTrue
        $rt2.EncodedBytes | Should -Be (Get-Item $env).Length
    }
    It 'refuses a truncated stream without keeping a value that is not there' {
        $bytes = [MailboxTransfer.FxCodec]::EncodePayload((New-LazyMessage -AttachBytes 100000 -BodyBytes 10))
        $cut = [byte[]]$bytes[0..60000]
        $doc = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromBytes($cut))
        $doc.Truncated | Should -BeTrue
        @($doc.Message.Attachments[0].Properties | Where-Object Id -EQ 0x3701).Count | Should -Be 0
    }
}

Describe 'the encoder with lazy values' {
    It 'writes the same bytes for a lazy value and for the same value in memory, and counts without reading' {
        $lazy = New-LazyMessage -AttachBytes 120000 -BodyBytes 90000
        $eager = [MailboxTransfer.FxMessage]::new()
        foreach ($p in $lazy.Properties) { $eager.Properties.Add($(if ($p.IsLazy) { $P::Create($p.Id, $p.Type, $p.Raw) } else { $p })) }
        $eager.Recipients.AddRange($lazy.Recipients)
        $ea = [MailboxTransfer.FxAttachment]::new()
        foreach ($p in $lazy.Attachments[0].Properties) { $ea.Properties.Add($(if ($p.IsLazy) { $P::Create($p.Id, $p.Type, $p.Raw) } else { $p })) }
        $eager.Attachments.Add($ea)
        $a = [MailboxTransfer.FxCodec]::EncodePayload($lazy)
        $b = [MailboxTransfer.FxCodec]::EncodePayload($eager)
        [Linq.Enumerable]::SequenceEqual([byte[]]$a, [byte[]]$b) | Should -BeTrue
        [MailboxTransfer.FxCodec]::PayloadLength($lazy) | Should -Be $a.Length
    }
    It 'knows the length of a message of any size without reading its values' {
        $m = New-LazyMessage -AttachBytes 3GB -BodyBytes 1GB
        $len = [MailboxTransfer.FxCodec]::PayloadLength($m)
        $len | Should -BeGreaterThan (4GB)
        $len | Should -BeLessThan (4GB + 2000)
        ([MailboxTransfer.FxMessageSource]::new($m)).Length | Should -Be $len
    }
    It 'refuses a value of 4 GiB or more: the length of a value is 32 bits in a FastTransfer stream' {
        $m = New-LazyMessage -AttachBytes 5GB -BodyBytes 10
        { [MailboxTransfer.FxCodec]::PayloadLength($m) } | Should -Throw '*32 bits*'
    }
    It 'accepts a lazy value of 4 GiB - 1, the largest length of a value (length only: nothing is read)' {
        $m = New-LazyMessage -AttachBytes ([long][uint32]::MaxValue) -BodyBytes 10
        [MailboxTransfer.FxCodec]::PayloadLength($m) | Should -BeGreaterThan ([long][uint32]::MaxValue)
    }
    It 'refuses a lazy value for a fixed-size type' {
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($P::CreateLazy(0x6000, 0x0003, (New-Pattern 4)))
        { [MailboxTransfer.FxCodec]::PayloadLength($m) } | Should -Throw '*variable-size*'
    }
}

Describe 'FxMessageSource, the FastTransfer stream of a message produced on demand' {
    It 'gives the bytes of EncodePayload / Encode, whatever the size of the reads, with the length announced' {
        $m = New-LazyMessage -AttachBytes 250000 -BodyBytes 100000
        $want = [MailboxTransfer.FxCodec]::EncodePayload($m)
        $src = [MailboxTransfer.FxMessageSource]::new($m)
        $src.Length | Should -Be $want.Length
        foreach ($size in 7, 4096, 100000) {
            $s = $src.OpenRead()
            try { [Linq.Enumerable]::SequenceEqual([byte[]](Read-Chunked $s $size), [byte[]]$want) | Should -BeTrue -Because "reads of $size" } finally { $s.Dispose() }
        }
        $doc0 = [MailboxTransfer.FxDocument]::new(); $doc0.Message = $m; $doc0.Envelope = New-GraphEnvelope
        $wantEnv = [MailboxTransfer.FxCodec]::Encode($doc0)
        $srcEnv = [MailboxTransfer.FxMessageSource]::new($m, $doc0.Envelope)
        $srcEnv.Length | Should -Be $wantEnv.Length
        $s = $srcEnv.OpenRead()
        try { [Linq.Enumerable]::SequenceEqual([byte[]](Read-Chunked $s 5000), [byte[]]$wantEnv) | Should -BeTrue } finally { $s.Dispose() }
    }
    It 'can be opened several times and decodes back to the same message' {
        $m = New-LazyMessage -AttachBytes 90000 -BodyBytes 80000
        $src = [MailboxTransfer.FxMessageSource]::new($m, (New-GraphEnvelope))
        $out = Join-Path $TestDrive 'source.fxs'
        [void][MailboxTransfer.TestSupport.Streams]::CopyToFile($src, $out)
        [void][MailboxTransfer.TestSupport.Streams]::CopyToFile($src, $out)
        (Get-Item $out).Length | Should -Be $src.Length
        $rt = [MailboxTransfer.FxCodec]::RoundTripFile($out)
        $rt.Identical | Should -BeTrue
        $att = ($rt.Document.Message.Attachments[0].Properties | Where-Object Id -EQ 0x3701).Blob
        [MailboxTransfer.TestSupport.Streams]::Equal($att, (New-Pattern 90000)) | Should -BeTrue
    }
    It 'fails, in the stream and in the PST writer, when the source of a value ends before its announced length' {
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($P::FromString(0x001A, 'IPM.Note'))
        $m.Properties.Add($P::CreateLazy(0x1000, 0x0102, [MailboxTransfer.TestSupport.LyingBlob]::new(100000, 60000)))
        $s = [MailboxTransfer.FxMessageSource]::new($m).OpenRead()
        try { { Read-Chunked $s 4096 } | Should -Throw '*ends*before*' } finally { $s.Dispose() }
        $w = [MailboxTransfer.PstWriter]::Create((New-PstPath 'liar'), 'Liar')
        try { { $w.AddMessage($w.IpmSubtree, $m) } | Should -Throw '*ends*before*' } finally { $w.Dispose() }
    }
    It 'is the body of importItem: ImportContent writes it as base64 without keeping it, and the data decodes to the stream' {
        $m = New-LazyMessage -AttachBytes 3000000 -BodyBytes 200000
        $src = [MailboxTransfer.FxMessageSource]::new($m, (New-GraphEnvelope))
        [MailboxTransfer.TestSupport.ImportBody]::Check($src, 'AAMk-folder-id') | Should -Be ''
        $import = [MailboxTransfer.MailboxApi].GetMethods() | Where-Object { $_.Name -eq 'ImportItem' -and $_.GetParameters().Count -ge 3 -and $_.GetParameters()[2].ParameterType -eq [MailboxTransfer.IFxStreamSource] }
        @($import).Count | Should -Be 1
    }
}

Describe 'strings converted while they are read' {
    It 'turns a code page string into UTF-16 without its terminator, across the pieces it is read in' {
        $text = ('caf' + [char]0xE9 + ' ' + [char]0x20AC + ' ') * 9000        # UTF-8 sequences of 1, 2 and 3 bytes cross the 16 KiB pieces
        $utf8 = [Text.Encoding]::UTF8.GetBytes($text) + [byte[]](0, 0)
        $blob = [MailboxTransfer.FxTranscodedBlob]::new([MailboxTransfer.FxBytesBlob]::new($utf8), [Text.Encoding]::UTF8)
        $blob.Length | Should -Be ([Text.Encoding]::Unicode.GetByteCount($text))
        [Text.Encoding]::Unicode.GetString([byte[]](Get-Bytes $blob)) | Should -BeExactly $text
        [Text.Encoding]::Unicode.GetString([byte[]]$blob.ReadRange(10, 20)) | Should -BeExactly $text.Substring(5, 10)
    }
}

Describe 'data trees written and read block by block' {
    It 'writes a value of <Bytes> bytes as a data tree, reads it back whole and from any offset, and the checker is clean' -ForEach @(
        @{ Bytes = 5000 }, @{ Bytes = 8176 }, @{ Bytes = 8177 }, @{ Bytes = 3MB }, @{ Bytes = 9MB + 17 }
    ) {
        $path = New-PstPath "tree$Bytes"
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Tree')
        $m = New-LazyMessage -AttachBytes $Bytes -BodyBytes 100000 -Seed 11
        $nid = $w.AddMessage($w.IpmSubtree, $m)
        $w.Close(); $w.Dispose()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $copy = $r.ReadMessage($nid)
            $prop = $copy.Attachments[0].Properties | Where-Object Id -EQ 0x3701
            $prop.RawLength | Should -Be $Bytes
            $prop.IsLazy | Should -Be ($Bytes -gt 64KB)
            [MailboxTransfer.TestSupport.Streams]::Equal($prop.Blob, (New-Pattern $Bytes 11)) | Should -BeTrue
            if ($prop.IsLazy) {
                $want = New-Pattern $Bytes 11
                foreach ($off in @(0, 1, 8175, 8176, 8177, (8176 * 1021 - 3), (8176 * 1021), (8176 * 1021 + 8176), ($Bytes - 5))) {
                    if ($off -ge $Bytes) { continue }
                    $n = [int][Math]::Min(40, $Bytes - $off)
                    [Linq.Enumerable]::SequenceEqual([byte[]]$prop.Lazy.ReadRange($off, $n), [byte[]]$want.ReadRange($off, $n)) | Should -BeTrue -Because "offset $off"
                }
            }
            # the same file read with every value in memory gives the same message
            $r.LazyThreshold = [long]::MaxValue
            $eager = $r.ReadMessage($nid)
            ($eager.Attachments[0].Properties | Where-Object Id -EQ 0x3701).IsLazy | Should -BeFalse
            $d = [Collections.Generic.List[string]]::new()
            [void][MailboxTransfer.PstFxCompare]::Compare($m, $copy, $d)
            $d | Should -BeNullOrEmpty
        } finally { $r.Dispose() }
        $res = Test-MbtPst -Path $path
        $res.Errors | Should -Be 0
        $res.Warnings | Should -Be 0
    }
    It 'detects a damaged block in the middle of a big value while it is being read' {
        $path = New-PstPath 'damaged'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Damaged')
        $nid = $w.AddMessage($w.IpmSubtree, (New-LazyMessage -AttachBytes 1MB -BodyBytes 10 -Seed 5))
        $w.Close(); $w.Dispose()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $blob = ($r.ReadMessage($nid).Attachments[0].Properties | Where-Object Id -EQ 0x3701).Lazy
            # a data block of the attachment in the file is found by 16 bytes of the pattern
            $head = [byte[]]$blob.ReadRange(8176 * 40, 16)
        } finally { $r.Dispose() }
        $raw = [IO.File]::ReadAllBytes($path)
        $at = [Text.Encoding]::Latin1.GetString($raw).IndexOf([Text.Encoding]::Latin1.GetString($head), [StringComparison]::Ordinal)
        $at | Should -BeGreaterThan 0
        $raw[$at + 100] = $raw[$at + 100] -bxor 0xFF
        [IO.File]::WriteAllBytes($path, $raw)
        $r2 = [MailboxTransfer.PstReader]::Open($path)
        try {
            $copy = $r2.ReadMessage($nid)      # lazy: the damaged block is not read yet
            $v = ($copy.Attachments[0].Properties | Where-Object Id -EQ 0x3701).Lazy
            { [MailboxTransfer.TestSupport.Streams]::Equal($v, (New-Pattern 1MB 5)) } | Should -Throw '*CRC*'
        } finally { $r2.Dispose() }
    }
    It 'refuses a value that does not fit in a data tree (a 32-bit total): 4 GiB or more' {
        $path = New-PstPath 'huge'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Huge')
        $m = New-LazyMessage -AttachBytes 5GB -BodyBytes 10
        try { { $w.AddMessage($w.IpmSubtree, $m) } | Should -Throw '*32-bit*' } finally { $w.Dispose() }
    }
    It 'reads a string8 body above the threshold as a Unicode string of the PST' {
        $text = ('caf' + [char]0xE9 + ' body ') * 20000     # 140,000 characters in code page 1252
        $bytes = [Text.Encoding]::GetEncoding(1252).GetBytes($text + [char]0)
        $m = New-LazyMessage -AttachBytes 10 -BodyBytes 10
        $m.Properties.RemoveAll({ param($p) $p.Id -eq 0x1000 }) | Out-Null
        $m.Properties.Add($P::CreateLazy(0x1000, 0x001E, [MailboxTransfer.FxBytesBlob]::new($bytes)))
        $path = New-PstPath 'string8'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'S8')
        $nid = $w.AddMessage($w.IpmSubtree, $m); $w.Close(); $w.Dispose()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $copy = $r.ReadMessage($nid)
            $b = $copy.Find(0x1000)
            $b.IsLazy | Should -BeTrue
            $b.Type | Should -Be 0x001F
            [Text.Encoding]::Unicode.GetString([byte[]](Get-Bytes $b.Blob)).TrimEnd([char]0) | Should -BeExactly $text
        } finally { $r.Dispose() }
    }
}

Describe 'a PST message as a FastTransfer stream, bounded memory' {
    It 'turns a message with a big attachment into a stream that decodes to the same attachment, and can be sent as an importItem body' {
        $path = New-PstPath 'toFx'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'ToFx')
        $nid = $w.AddMessage($w.IpmSubtree, (New-LazyMessage -AttachBytes 12MB -BodyBytes 300000 -Seed 21))
        $w.Close(); $w.Dispose()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $env = New-GraphEnvelope
            $src = $r.OpenMessageSource($nid, $env)
            $out = Join-Path $TestDrive 'fromPst.fxs'
            [MailboxTransfer.TestSupport.Streams]::CopyToFile($src, $out) | Should -Be $src.Length
            $rt = [MailboxTransfer.FxCodec]::RoundTripFile($out)
            $rt.Identical | Should -BeTrue
            $rt.Document.Envelope.DataRecords | Should -BeGreaterThan 100
            $att = $rt.Document.Message.Attachments[0].Properties | Where-Object Id -EQ 0x3701
            $att.IsLazy | Should -BeTrue
            [MailboxTransfer.TestSupport.Streams]::Equal($att.Blob, (New-Pattern 12MB 21)) | Should -BeTrue
            $rt.Document.Message.GetString(0x0037) | Should -Be 'Lazy'
            [MailboxTransfer.TestSupport.ImportBody]::Check($r.OpenMessageSource($nid), 'folder') | Should -Be ''
            [MailboxTransfer.TestSupport.ImportBody]::Check($src, 'folder') | Should -Be ''
        } finally { $r.Dispose() }
    }
    It 'keeps working when two streams of the same PST are read in turn' {
        $path = New-PstPath 'two'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Two')
        $n1 = $w.AddMessage($w.IpmSubtree, (New-LazyMessage -AttachBytes 1MB -BodyBytes 10 -Seed 31))
        $n2 = $w.AddMessage($w.IpmSubtree, (New-LazyMessage -AttachBytes 2MB -BodyBytes 10 -Seed 32))
        $w.Close(); $w.Dispose()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $a = ($r.ReadMessage($n1).Attachments[0].Properties | Where-Object Id -EQ 0x3701).Blob
            $b = ($r.ReadMessage($n2).Attachments[0].Properties | Where-Object Id -EQ 0x3701).Blob
            $sa = $a.OpenRead(); $sb = $b.OpenRead()
            try {
                $x = [byte[]]::new(5000); $y = [byte[]]::new(5000)
                $ea = [MailboxTransfer.TestSupport.PatternBlob]::new(1MB, 31).OpenRead(); $eb = [MailboxTransfer.TestSupport.PatternBlob]::new(2MB, 32).OpenRead()
                try {
                    for ($i = 0; $i -lt 100; $i++) {
                        $sa.ReadExactly($x, 0, 5000); $sb.ReadExactly($y, 0, 5000)
                        $ex = [byte[]]::new(5000); $ey = [byte[]]::new(5000); $ea.ReadExactly($ex, 0, 5000); $eb.ReadExactly($ey, 0, 5000)
                        [Linq.Enumerable]::SequenceEqual($x, $ex) | Should -BeTrue
                        [Linq.Enumerable]::SequenceEqual($y, $ey) | Should -BeTrue
                    }
                } finally { $ea.Dispose(); $eb.Dispose() }
            } finally { $sa.Dispose(); $sb.Dispose() }
        } finally { $r.Dispose() }
    }
}

Describe 'cancellation and progress (for a UI: no console output, a token and an IProgress)' {
    It 'reports the bytes written and the messages while a PST message is written' {
        $rec = [MailboxTransfer.TestSupport.ProgressRecorder]::new()
        $w = [MailboxTransfer.PstWriter]::Create((New-PstPath 'prog'), 'Prog')
        try { [void]$w.AddMessage($w.IpmSubtree, (New-LazyMessage -AttachBytes 5MB -BodyBytes 100000), $rec.Token, $rec); $w.Close() } finally { $w.Dispose() }
        $rec.Count | Should -BeGreaterThan 2
        $rec.Last.Items | Should -Be 1
        $rec.Last.BytesDone | Should -BeGreaterThan 5MB
        $rec.Last.CurrentFolder | Should -Not -BeNullOrEmpty
    }
    It 'stops writing a big value between two blocks when the token is cancelled' {
        $rec = [MailboxTransfer.TestSupport.ProgressRecorder]::new(); $rec.CancelAtBytes = 2MB
        $w = [MailboxTransfer.PstWriter]::Create((New-PstPath 'cancelw'), 'Cancel')
        try {
            { $w.AddMessage($w.IpmSubtree, (New-LazyMessage -AttachBytes 40MB -BodyBytes 100000), $rec.Token, $rec) } | Should -Throw '*cancel*'
        } finally { $w.Dispose() }
        $rec.Last.BytesDone | Should -BeLessThan 6MB
    }
    It 'refuses to start when the token is already cancelled (split writer too)' {
        $cts = [Threading.CancellationTokenSource]::new(); $cts.Cancel()
        $sw = [MailboxTransfer.PstSplitWriter]::new((New-PstPath 'cancels'), 'C', 0)
        try { { $sw.AddMessage('Mail', 'IPF.Note', (New-LazyMessage), $cts.Token, $null) } | Should -Throw '*cancel*' } finally { $sw.Dispose() }
    }
    It 'copies a PST message to a stream with progress, and stops when cancelled' {
        $path = New-PstPath 'copyto'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Copy')
        $nid = $w.AddMessage($w.IpmSubtree, (New-LazyMessage -AttachBytes 20MB -BodyBytes 1000)); $w.Close(); $w.Dispose()
        $r = [MailboxTransfer.PstReader]::Open($path)
        try {
            $rec = [MailboxTransfer.TestSupport.ProgressRecorder]::new()
            $n = $r.CopyMessageTo($nid, $null, [IO.Stream]::Null, $rec.Token, $rec)
            $rec.Last.BytesDone | Should -Be $n
            $rec.Last.BytesTotal | Should -Be $n
            $rec2 = [MailboxTransfer.TestSupport.ProgressRecorder]::new(); $rec2.CancelAtBytes = 3MB
            { $r.CopyMessageTo($nid, $null, [IO.Stream]::Null, $rec2.Token, $rec2) } | Should -Throw '*cancel*'
            $rec2.Last.BytesDone | Should -BeLessThan 6MB
        } finally { $r.Dispose() }
    }
    It 'stops a check when cancelled and says so in the report' {
        $path = New-PstPath 'cancelc'
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Chk')
        [void]$w.AddMessage($w.IpmSubtree, (New-LazyMessage -AttachBytes 1MB)); $w.Close(); $w.Dispose()
        $rec = [MailboxTransfer.TestSupport.ProgressRecorder]::new(); $rec.CancelAtBytes = 0
        $o = [MailboxTransfer.PstCheckOptions]::new(); $o.ProgressReport = $rec; $o.Cancellation = $rec.Token
        $rep = [MailboxTransfer.PstChecker]::Check($path, $o)
        $rep.Cancelled | Should -BeTrue
        $o2 = [MailboxTransfer.PstCheckOptions]::new(); $rec3 = [MailboxTransfer.TestSupport.ProgressRecorder]::new(); $o2.ProgressReport = $rec3
        $rep2 = [MailboxTransfer.PstChecker]::Check($path, $o2)
        $rep2.Cancelled | Should -BeFalse
        $rep2.Errors | Should -Be 0
        $rec3.Count | Should -BeGreaterThan 3
    }
    It 'stops a decode when cancelled' {
        $bytes = [MailboxTransfer.FxCodec]::EncodePayload((New-LazyMessage -AttachBytes 300000))
        $cts = [Threading.CancellationTokenSource]::new(); $cts.Cancel()
        $o = [MailboxTransfer.FxDecodeOptions]::new(); $o.Cancellation = $cts.Token
        { [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromBytes($bytes), $o) } | Should -Throw '*cancel*'
    }
}

Describe 'memory used to move big items (tools\Measure-PstMemory.ps1 in a child process)' -Tag 'Memory', 'Slow' {
    BeforeAll {
        $script:Tool = Join-Path $script:RepoRoot 'tools\Measure-PstMemory.ps1'
        function Invoke-Measure {
            param([string]$Scenario, [int]$SizeMiB = 1024)
            $out = & (Get-Process -Id $PID).Path -NoProfile -File $script:Tool -Scenario $Scenario -SizeMiB $SizeMiB -LimitMiB 250 -Json 2>&1
            $text = ($out | Where-Object { $_ -is [string] }) -join "`n"
            $start = $text.IndexOf('[')
            if ($start -lt 0) { throw "No JSON from the measure: $text" }
            $text.Substring($start) | ConvertFrom-Json
        }
    }
    It 'moves a synthetic item with an attachment of <MiB> MiB (FastTransfer -> PST -> read back -> FastTransfer -> importItem body) with less than 250 MiB of growth' -ForEach @(@{ MiB = $script:MemoryMiB }) -Skip:($script:MemoryMiB -le 0) {
        $phases = @(Invoke-Measure -Scenario Synthetic -SizeMiB $MiB)
        $phases.Count | Should -Be 5
        foreach ($p in $phases) {
            $p.Ok | Should -BeTrue -Because "$($p.Phase): $($p.Note)"
            $p.WorkingSetGrowthMiB | Should -BeLessThan 250 -Because $p.Phase
            $p.PrivateGrowthMiB | Should -BeLessThan 250 -Because $p.Phase
            $p.Bytes | Should -BeGreaterThan ($MiB * 1MB)
        }
    }
    It 'writes the real corpus (13 items of 100 MiB) to a PST, reads it back and turns it into streams with less than 250 MiB of growth' -Skip:(-not $script:HasCorpus) {
        $phases = @(Invoke-Measure -Scenario Corpus)
        $phases.Count | Should -Be 3
        foreach ($p in $phases) {
            $p.Ok | Should -BeTrue -Because "$($p.Phase): $($p.Note)"
            $p.WorkingSetGrowthMiB | Should -BeLessThan 250 -Because $p.Phase
        }
    }
}
