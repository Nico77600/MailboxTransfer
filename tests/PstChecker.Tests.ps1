#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the PST integrity checker (Engine.PstChecker.cs, Test-MbtPst).
    A clean PST written by our writer must give no finding; deliberately damaged copies (a flipped byte in a block, a broken CRC,
    a contents row removed, a wrong count, a dangling NID, a B-tree cycle, ...) must each give the right finding and never an
    exception. Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\PstChecker.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport

    function New-Msg {
        param([string]$Subject, [int]$AttachBytes = 0, [switch]$Embedded, [switch]$Unread)
        $P = [MailboxTransfer.FxProperty]
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($P::FromString(0x001A, 'IPM.Note'))
        $m.Properties.Add($P::FromString(0x0037, $Subject))
        $m.Properties.Add($P::FromString(0x1000, ('body ' * 400)))
        $m.Properties.Add($P::FromInt32(0x0E07, $(if ($Unread) { 0 } else { 1 })))
        $m.Properties.Add($P::FromInt32(0x8001, 5).Named([guid]'11111111-2222-3333-4444-555555555555', [uint32]0x8501))
        $r = [MailboxTransfer.FxRecipient]::new()
        $r.Properties.Add($P::FromInt32(0x0C15, 1)); $r.Properties.Add($P::FromString(0x3001, 'Bob'))
        $m.Recipients.Add($r)
        if ($AttachBytes -gt 0) {
            $a = [MailboxTransfer.FxAttachment]::new()
            $b = [byte[]]::new($AttachBytes); [Random]::new(1).NextBytes($b)
            $a.Properties.Add($P::FromInt32(0x3705, 1)); $a.Properties.Add($P::FromBinary(0x3701, $b)); $a.Properties.Add($P::FromString(0x3707, 'f.bin'))
            $m.Attachments.Add($a)
        }
        if ($Embedded) {
            $e = [MailboxTransfer.FxMessage]::new()
            $e.Properties.Add($P::FromString(0x0037, 'Inner'))
            $ea = [MailboxTransfer.FxAttachment]::new()
            $ea.Properties.Add($P::FromInt32(0x3705, 5)); $ea.Properties.Add($P::Create(0x3701, 0x000D, [byte[]]@()))
            $ea.Embedded = $e
            $m.Attachments.Add($ea)
        }
        return $m
    }

    function New-TestPst {
        param([string]$Name, [string[]]$Faults = @())
        $path = Join-Path $TestDrive ($Name + '.pst')
        $o = [MailboxTransfer.PstWriterOptions]::new(); $o.Faults = $Faults
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Checked', $o)
        $a = $w.CreateFolder($null, 'A', 'IPF.Note'); $b = $w.CreateFolder($a, 'B', 'IPF.Note'); $c = $w.CreateFolder($null, 'C', 'IPF.Note')
        foreach ($i in 1..40) { [void]$w.AddMessage($a, (New-Msg "a$i" -AttachBytes $(if ($i % 10 -eq 0) { 30000 } else { 0 }) -Unread:($i % 4 -eq 0))) }
        foreach ($i in 1..15) { [void]$w.AddMessage($b, (New-Msg "b$i" -Embedded:($i % 5 -eq 0))) }
        [void]$w.AddMessage($c, (New-Msg 'big' -AttachBytes 3MB))
        $w.Close()
        return $path
    }

    function Invoke-Check {
        param([string]$Path, [switch]$NoDeep)
        $o = [MailboxTransfer.PstCheckOptions]::new(); $o.Deep = -not $NoDeep
        [MailboxTransfer.PstChecker]::Check($Path, $o)
    }

    function Get-Codes { param($Report) @($Report.Findings | ForEach-Object { $_.Code } | Sort-Object -Unique) }

    function Copy-Pst { param([string]$Source, [string]$Name) $d = Join-Path $TestDrive ($Name + '.pst'); Copy-Item -LiteralPath $Source -Destination $d; $d }

    function Set-PageCrc { param([byte[]]$Raw, [long]$Ib) [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($Raw, [int]$Ib, 496)).CopyTo($Raw, $Ib + 500) }

    function Get-LeafPage {
        param([byte[]]$Raw, [long]$RootIb, [int]$Which = 0)
        $ib = $RootIb
        while ($Raw[$ib + 491] -gt 0) { $ib = [BitConverter]::ToInt64($Raw, $ib + 16 + 24 * $Which) }
        $ib
    }

    $script:Good = New-TestPst 'good'
}

Describe 'a clean PST' {
    It 'is clean: no error, no warning, and the whole structure is counted' {
        $r = Invoke-Check $script:Good
        ($r.Findings | Where-Object { $_.Severity -ne 'Info' }) | Should -BeNullOrEmpty
        $r.Errors | Should -Be 0
        $r.Warnings | Should -Be 0
        $r.ExitCode | Should -Be 0
        $r.Stats['messages'] | Should -Be 56
        $r.Stats['folders'] | Should -Be 7
        $r.Stats['attachments'] | Should -Be 8        # 4 files of 30 KB, the big one, 3 embedded
        $r.Stats['namedProperties'] | Should -Be 1
        $r.Stats['bytesVerified'] | Should -BeGreaterThan 3000000
    }
    It 'is clean through Test-MbtPst (the public function), with the exit code convention' {
        Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
        $x = Test-MbtPst -Path $script:Good
        $x.Ok | Should -BeTrue
        $x.ExitCode | Should -Be 0
        $x.Errors | Should -Be 0
    }
    It 'is clean with a PST written without the spam search folder, and with an empty store' {
        $p = Join-Path $TestDrive 'nospam.pst'
        $o = [MailboxTransfer.PstWriterOptions]::new(); $o.SpamSearchFolder = $false
        $w = [MailboxTransfer.PstWriter]::Create($p, 'x', $o); $w.Close()
        $r = Invoke-Check $p
        $r.Errors | Should -Be 0
        $e = Join-Path $TestDrive 'empty.pst'
        $w = [MailboxTransfer.PstWriter]::Create($e, 'x'); $w.Close()
        (Invoke-Check $e).Errors | Should -Be 0
    }
    It 'is clean on a multi-level B-tree file made of raw nodes (NDB layer only: the messaging layer reports the missing store)' {
        $p = Join-Path $TestDrive 'raw.pst'
        $w = [MailboxTransfer.PstNdbWriter]::new($p)
        for ($i = 1; $i -le 9000; $i++) { $n = [MailboxTransfer.PstNode]::new(); $n.Raw = [byte[]]@(1, 2, 3); $w.WriteTopLevelNode([uint32]($i * 32 + 4), $n, 0) }
        $w.Close()
        $o = [MailboxTransfer.PstCheckOptions]::new(); $o.Messaging = $false
        $r = [MailboxTransfer.PstChecker]::Check($p, $o)
        $r.Errors | Should -Be 0
        $r.Warnings | Should -Be 0
        $r.Stats['btreePages'] | Should -BeGreaterThan 50
    }
}

Describe 'damaged messaging structures' {
    It 'detects <Fault>: <Code>' -ForEach @(
        @{ Fault = 'contents-row-removed'; Code = 'MSG-NOT-LISTED' }
        @{ Fault = 'contents-row-removed'; Code = 'FOLDER-COUNT' }
        @{ Fault = 'dangling-nid'; Code = 'CONTENTS-ROW' }
        @{ Fault = 'wrong-count'; Code = 'FOLDER-COUNT' }
        @{ Fault = 'hierarchy-missing'; Code = 'FOLDER-NOT-LISTED' }
        @{ Fault = 'no-recipient-table'; Code = 'MSG-RECIPIENTS' }
        @{ Fault = 'attach-table-short'; Code = 'MSG-ATTACH-MISMATCH' }
    ) {
        $p = New-TestPst "fault-$Fault" -Faults @($Fault)
        $r = Invoke-Check $p
        Get-Codes $r | Should -Contain $Code
        $r.ExitCode | Should -BeGreaterThan 0
        # the finding says where: folder path or message
        ($r.Findings | Where-Object Code -EQ $Code | Select-Object -First 1).Where + ($r.Findings | Where-Object Code -EQ $Code | Select-Object -First 1).Nid | Should -Not -BeNullOrEmpty
    }
    It 'reports the folder path and the subject of the damaged message' {
        $p = New-TestPst 'fault-norec' -Faults @('no-recipient-table')
        $f = (Invoke-Check $p).Findings | Where-Object Code -EQ 'MSG-RECIPIENTS' | Select-Object -First 1
        $f.Where | Should -Match '^/'
        $f.Where | Should -Match '\| (a|b|big)'
    }
}

Describe 'damaged files (bytes changed in a copy)' {
    It 'detects a flipped byte in a data block with its offset and BID, without failing on the rest' {
        $p = Copy-Pst $script:Good 'flip'
        $f = [MailboxTransfer.PstFile]::Open($p)
        try {
            $nbt = $f.OpenNode(0x122); $e = [MailboxTransfer.PstBbtEntry]::new(); [void]$f.TryFindBlock($nbt.BidData, [ref]$e)
            $ib = [long]$e.Ib; $bid = $e.Bid
        } finally { $f.Dispose() }
        $raw = [IO.File]::ReadAllBytes($p); $raw[$ib + 10] = $raw[$ib + 10] -bxor 0xFF; [IO.File]::WriteAllBytes($p, $raw)
        $r = Invoke-Check $p
        $hit = $r.Findings | Where-Object Code -EQ 'BLOCK-CRC' | Select-Object -First 1
        $hit | Should -Not -BeNullOrEmpty
        $hit.Offset | Should -Be $ib
        $hit.Bid | Should -Be $bid
        $r.ExitCode | Should -Be 1
    }
    It 'detects broken header CRCs' {
        $p = Copy-Pst $script:Good 'hdrcrc'
        $raw = [IO.File]::ReadAllBytes($p); $raw[20] = $raw[20] -bxor 1; [IO.File]::WriteAllBytes($p, $raw)
        $c = Get-Codes (Invoke-Check $p)
        $c | Should -Contain 'HDR-CRC-PARTIAL'
        $c | Should -Contain 'HDR-CRC-FULL'
    }
    It 'detects a broken CRC of a B-tree page, a wrong page type and a wrong signature' {
        $p = Copy-Pst $script:Good 'pagecrc'
        $raw = [IO.File]::ReadAllBytes($p)
        $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 224))     # NBT
        $raw[$leaf + 5] = $raw[$leaf + 5] -bxor 0x55; [IO.File]::WriteAllBytes($p, $raw)
        Get-Codes (Invoke-Check $p) | Should -Contain 'NBT-PAGE'
        $raw = [IO.File]::ReadAllBytes($script:Good)
        $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 240))     # BBT
        $raw[$leaf + 498] = $raw[$leaf + 498] -bxor 1; Set-PageCrc $raw $leaf; [IO.File]::WriteAllBytes($p, $raw)
        (Invoke-Check $p).Findings | Where-Object { $_.Code -eq 'BBT-PAGE' -and $_.Message -match 'signature' } | Should -Not -BeNullOrEmpty
    }
    It 'detects keys out of order in an NBT leaf and a dangling data BID (CRC recomputed so the page is otherwise valid)' {
        $p = Copy-Pst $script:Good 'order'
        $raw = [IO.File]::ReadAllBytes($p)
        $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 224))
        # swap the NIDs of entries 0 and 1
        $a = $raw[$leaf..($leaf + 7)]; $b = $raw[($leaf + 32)..($leaf + 39)]
        $b.CopyTo($raw, $leaf); $a.CopyTo($raw, $leaf + 32); Set-PageCrc $raw $leaf
        [IO.File]::WriteAllBytes($p, $raw)
        Get-Codes (Invoke-Check $p) | Should -Contain 'NBT-KEY-ORDER'
        $raw = [IO.File]::ReadAllBytes($script:Good)
        $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 224))
        [BitConverter]::GetBytes([uint64]0x7777770).CopyTo($raw, $leaf + 8); Set-PageCrc $raw $leaf
        [IO.File]::WriteAllBytes($p, $raw)
        $c = Get-Codes (Invoke-Check $p)
        $c | Should -Contain 'BID-DANGLING'
    }
    It 'detects a header whose NID high-water marks are lower than the NIDs in use (header CRCs recomputed)' {
        $p = Copy-Pst $script:Good 'highwater'
        $raw = [IO.File]::ReadAllBytes($p)
        [BitConverter]::GetBytes([uint32]0x400).CopyTo($raw, 44 + 4 * 4)          # rgnid[NID_TYPE_NORMAL_MESSAGE]
        [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($raw, 8, 471)).CopyTo($raw, 4)
        [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($raw, 8, 516)).CopyTo($raw, 524)
        [IO.File]::WriteAllBytes($p, $raw)
        Get-Codes (Invoke-Check $p) | Should -Contain 'HDR-NID-HIGHWATER'
    }
    It 'detects a cycle in the BBT (a child reference pointing back to the root page)' {
        $p = Join-Path $TestDrive 'cycle.pst'
        $w = [MailboxTransfer.PstNdbWriter]::new($p)
        for ($i = 1; $i -le 9000; $i++) { $n = [MailboxTransfer.PstNode]::new(); $n.Raw = [byte[]]@(1, 2, 3); $w.WriteTopLevelNode([uint32]($i * 32 + 4), $n, 0) }
        $w.Close()
        $raw = [IO.File]::ReadAllBytes($p)
        $rootIb = [BitConverter]::ToInt64($raw, 240); $rootBid = [BitConverter]::ToUInt64($raw, 232)
        $raw[$rootIb + 491] | Should -BeGreaterThan 1
        [BitConverter]::GetBytes($rootBid).CopyTo($raw, $rootIb + 24 + 8)          # entry 1: bid and ib of the root itself
        [BitConverter]::GetBytes($rootIb).CopyTo($raw, $rootIb + 24 + 16)
        Set-PageCrc $raw $rootIb
        [IO.File]::WriteAllBytes($p, $raw)
        $o = [MailboxTransfer.PstCheckOptions]::new(); $o.Messaging = $false
        $r = [MailboxTransfer.PstChecker]::Check($p, $o)
        Get-Codes $r | Should -Contain 'BBT-CYCLE'
    }
    It 'detects a block marked free in the AMap, and overlapping blocks' {
        $p = Copy-Pst $script:Good 'amap'
        $f = [MailboxTransfer.PstFile]::Open($p)
        try { $e = [MailboxTransfer.PstBbtEntry]::new(); [void]$f.TryFindBlock($f.OpenNode(0x122).BidData, [ref]$e); $ib = [long]$e.Ib } finally { $f.Dispose() }
        $raw = [IO.File]::ReadAllBytes($p)
        $base = 0x4400 + [math]::Floor(($ib - 0x4400) / 253952) * 253952
        $unit = [int](($ib - $base) / 64)
        $raw[$base + [math]::Floor($unit / 8)] = $raw[$base + [math]::Floor($unit / 8)] -band ((-bnot (0x80 -shr ($unit % 8))) -band 0xFF)
        Set-PageCrc $raw $base; [IO.File]::WriteAllBytes($p, $raw)
        Get-Codes (Invoke-Check $p) | Should -Contain 'AMAP-UNMARKED'
        # overlap: the second BBT entry of the first leaf points to the offset of the first
        $p2 = Copy-Pst $script:Good 'overlap'
        $raw = [IO.File]::ReadAllBytes($p2)
        $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 240))
        $raw[$leaf..($leaf + 7)] | Out-Null
        $raw[($leaf + 8)..($leaf + 15)].CopyTo($raw, $leaf + 24 + 8); Set-PageCrc $raw $leaf
        [IO.File]::WriteAllBytes($p2, $raw)
        Get-Codes (Invoke-Check $p2) | Should -Contain 'ALLOC-OVERLAP'
    }
    It 'detects an XBLOCK whose total does not match its data blocks' {
        $p = Copy-Pst $script:Good 'xblock'
        # find any internal data block of a top-level node by scanning the BBT leaf pages
        $raw = [IO.File]::ReadAllBytes($p)
        $found = $false
        $f = [MailboxTransfer.PstFile]::Open($p)
        try {
            $script:xb = $null
            function Find-XBlock($view) {
                foreach ($sn in $view.SubnodeNids) {
                    $sub = $view.OpenSubnode([uint32]$sn)
                    if ($sub.BidData -ne 0 -and ($sub.BidData -band 2) -ne 0) {
                        $d = $f.ReadBlock($sub.BidData)
                        if ($d[0] -eq 1 -and $d[1] -eq 1) { return $sub.BidData }
                    }
                    if ($sub.BidSub -ne 0) { $r = Find-XBlock $sub; if ($r) { return $r } }
                }
            }
            foreach ($n in $f.EnumerateNodes()) {
                if ($n.BidSub -eq 0) { continue }
                $bid = Find-XBlock $f.OpenNode($n.Nid)
                if ($bid) {
                    $e = [MailboxTransfer.PstBbtEntry]::new(); [void]$f.TryFindBlock($bid, [ref]$e)
                    $ib = [long]$e.Ib
                    $raw[$ib + 4] = $raw[$ib + 4] + 1       # lcbTotal + 1
                    $total = [MailboxTransfer.PstFormat]::BlockSize($e.Cb)
                    [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($raw, [int]$ib, $e.Cb)).CopyTo($raw, $ib + $total - 16 + 4)
                    $found = $true; break
                }
            }
        } finally { $f.Dispose() }
        $found | Should -BeTrue
        [IO.File]::WriteAllBytes($p, $raw)
        $r = Invoke-Check $p
        Get-Codes $r | Should -Contain 'XBLOCK-TOTAL'
        Get-Codes $r | Should -Not -Contain 'BLOCK-CRC'
    }
    It 'survives a truncated file, a file of zeros, random bytes, a tiny file and an empty file, with findings and no exception' -ForEach @(
        @{ Kind = 'truncated' }, @{ Kind = 'zeros' }, @{ Kind = 'random' }, @{ Kind = 'tiny' }, @{ Kind = 'empty' }
    ) {
        $p = Join-Path $TestDrive "bad-$Kind.pst"
        switch ($Kind) {
            'truncated' { $raw = [IO.File]::ReadAllBytes($script:Good); [IO.File]::WriteAllBytes($p, $raw[0..([int]($raw.Length / 2))]) }
            'zeros' { [IO.File]::WriteAllBytes($p, [byte[]]::new(300000)) }
            'random' { $b = [byte[]]::new(300000); [Random]::new(9).NextBytes($b); [IO.File]::WriteAllBytes($p, $b) }
            'tiny' { [IO.File]::WriteAllBytes($p, [byte[]](1..100)) }
            'empty' { [IO.File]::WriteAllBytes($p, [byte[]]@()) }
        }
        { $script:r = Invoke-Check $p } | Should -Not -Throw
        $script:r.Errors | Should -BeGreaterThan 0
        $script:r.ExitCode | Should -Be 1
        if ($Kind -eq 'truncated') { Get-Codes $script:r | Should -Contain 'HDR-EOF' }
        if ($Kind -in 'zeros', 'random') { Get-Codes $script:r | Should -Contain 'HDR-MAGIC' }
    }
    It 'survives random damage: 40 copies with 30 random bytes changed each never raise an exception' {
        $raw0 = [IO.File]::ReadAllBytes($script:Good)
        $rnd = [Random]::new(2026)
        $withErrors = 0
        for ($i = 0; $i -lt 40; $i++) {
            $raw = $raw0.Clone()
            for ($k = 0; $k -lt 30; $k++) { $raw[$rnd.Next(0, $raw.Length)] = [byte]$rnd.Next(0, 256) }
            $p = Join-Path $TestDrive "rnd$i.pst"; [IO.File]::WriteAllBytes($p, $raw)
            $r = $null
            try { $r = Invoke-Check $p } catch { throw "exception on copy $i : $($_.Exception.Message)" }
            if ($r.Errors -gt 0) { $withErrors++ }
            Remove-Item -LiteralPath $p
        }
        $withErrors | Should -BeGreaterThan 30       # most random bytes land in checked data
    }
}

Describe 'constraints beyond the specification (observed in Outlook''s own validation during development), each a rule of the checker' {
    BeforeAll {
        function Set-HeaderCrc { param([byte[]]$Raw) [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($Raw, 8, 471)).CopyTo($Raw, 4); [BitConverter]::GetBytes([MailboxTransfer.PstFormat]::Crc($Raw, 8, 516)).CopyTo($Raw, 524) }
        function New-DamagedCopy {
            param([string]$Name, [scriptblock]$Damage)
            $p = Copy-Pst $script:Good $Name
            $raw = [IO.File]::ReadAllBytes($p)
            & $Damage $raw
            [IO.File]::WriteAllBytes($p, $raw)
            $p
        }
    }

    It 'the writer never produces any of them: a large folder (heaps and row matrices over several blocks), attachments and embedded messages are clean' {
        $p = Join-Path $TestDrive 'large.pst'
        $w = [MailboxTransfer.PstWriter]::Create($p, 'Large')
        $f = $w.CreateFolder($null, 'Many', 'IPF.Note')
        foreach ($i in 1..1500) { [void]$w.AddMessage($f, (New-Msg "m$i" -Embedded:($i % 300 -eq 0) -AttachBytes $(if ($i % 500 -eq 0) { 20000 } else { 0 }) -Unread:($i % 3 -eq 0))) }
        $w.Close()
        $r = Invoke-Check $p
        ($r.Findings | Where-Object { $_.Severity -ne 'Info' } | Select-Object -First 5 | ForEach-Object { $_.ToString() }) | Should -BeNullOrEmpty
        $r.ExitCode | Should -Be 0
    }

    It 'data that looks like a map page (type bytes 0x83 0x83 at offset 496 of a page) is data, not a PMap: no MAP-BAD, no ALLOC-OVERLAP (found on a 2.4 GB file with ordinary attachment data)' {
        $p = Join-Path $TestDrive 'maplike.pst'
        $w = [MailboxTransfer.PstNdbWriter]::new($p)
        for ($i = 1; $i -le 90; $i++) {
            $n = [MailboxTransfer.PstNode]::new(); $b = [byte[]]::new(8176); $b[496] = 0x83; $b[497] = 0x83; $b[1008] = 0x85; $b[1009] = 0x85; $b[3] = $i
            $n.Raw = $b; $w.WriteTopLevelNode([uint32]($i * 32 + 4), $n, 0)
        }
        $w.Close()
        $o = [MailboxTransfer.PstCheckOptions]::new(); $o.Messaging = $false
        $r = [MailboxTransfer.PstChecker]::Check($p, $o)
        $r.Stats['sections'] | Should -BeGreaterThan 2
        Get-Codes $r | Should -Not -Contain 'MAP-BAD'
        Get-Codes $r | Should -Not -Contain 'ALLOC-OVERLAP'
        $r.Errors | Should -Be 0
        $r.Warnings | Should -Be 0
    }

    It 'AMAP-BIT-ORDER: the bits of an AMap are numbered from the most significant bit' {
        $p = New-DamagedCopy 'bitorder' {
            param($raw)
            for ($i = 0; $i -lt 496; $i++) { $b = $raw[0x4400 + $i]; $r = 0; for ($k = 0; $k -lt 8; $k++) { if ($b -band (1 -shl $k)) { $r = $r -bor (0x80 -shr $k) } }; $raw[0x4400 + $i] = $r }
            Set-PageCrc $raw 0x4400
        }
        Get-Codes (Invoke-Check $p) | Should -Contain 'AMAP-BIT-ORDER'
    }
    It 'DATATREE-NOT-FULL: every block of a data tree but the last holds 8,176 bytes' {
        $p = Join-Path $TestDrive 'short-block.pst'
        $w = [MailboxTransfer.PstNdbWriter]::new($p)
        $n = [MailboxTransfer.PstNode]::new(); $n.Blocks = [Collections.Generic.List[byte[]]]::new(); $n.Blocks.Add([byte[]]::new(100)); $n.Blocks.Add([byte[]]::new(100))
        $w.WriteTopLevelNode([uint32]0x4004, $n, 0); $w.Close()
        $o = [MailboxTransfer.PstCheckOptions]::new(); $o.Messaging = $false
        Get-Codes ([MailboxTransfer.PstChecker]::Check($p, $o)) | Should -Contain 'DATATREE-NOT-FULL'
        $ok = Join-Path $TestDrive 'full-blocks.pst'
        $w = [MailboxTransfer.PstNdbWriter]::new($ok)
        $n = [MailboxTransfer.PstNode]::new(); $n.Blocks = [Collections.Generic.List[byte[]]]::new(); $n.Blocks.Add([byte[]]::new(8176)); $n.Blocks.Add([byte[]]::new(8176)); $n.Blocks.Add([byte[]]::new(100))
        $w.WriteTopLevelNode([uint32]0x4004, $n, 0); $w.Close()
        Get-Codes ([MailboxTransfer.PstChecker]::Check($ok, $o)) | Should -Not -Contain 'DATATREE-NOT-FULL'
    }
    It 'HDR-NID-HIGHWATER: attachment and embedded message NIDs come from the counters of the header (subnodes outside the NBT)' {
        $p = New-TestPst 'fault-attach-highwater' -Faults @('attach-highwater')
        Get-Codes (Invoke-Check $p) | Should -Contain 'HDR-NID-HIGHWATER'
    }
    It 'PMAP-NOT-FULL and HDR-PMAPFREE: the deprecated PMap of a file without free page is all ones and the header says 0' {
        $p = New-DamagedCopy 'pmap' { param($raw) $raw[0x4600 + 3] = 0x7F; Set-PageCrc $raw 0x4600 }
        Get-Codes (Invoke-Check $p) | Should -Contain 'PMAP-NOT-FULL'
        $p2 = New-DamagedCopy 'pmapfree' { param($raw) [BitConverter]::GetBytes([uint64]64).CopyTo($raw, 208); Set-HeaderCrc $raw }
        Get-Codes (Invoke-Check $p2) | Should -Contain 'HDR-PMAPFREE'
    }
    It 'HDR-RGBFM-RGBFP: the deprecated FMap / FPMap of the header are all 0xFF (written so, reported otherwise)' {
        $ok = New-TestPst 'hdr-fm-ok'
        $raw = [IO.File]::ReadAllBytes($ok)
        ($raw[256..511] | Where-Object { $_ -ne 0xFF }).Count | Should -Be 0
        Get-Codes (Invoke-Check $ok) | Should -Not -Contain 'HDR-RGBFM-RGBFP'
        $p = New-DamagedCopy 'rgbfm' { param($raw) $raw[256 + 5] = 0x3C; Set-HeaderCrc $raw }
        Get-Codes (Invoke-Check $p) | Should -Contain 'HDR-RGBFM-RGBFP'
    }
    It 'HDR-UNIQUE and HDR-BIDNEXTP: dwUnique is positive and the page BIDs are below bidNextP' {
        $p = New-DamagedCopy 'unique' { param($raw) [BitConverter]::GetBytes([uint32]0).CopyTo($raw, 40); Set-HeaderCrc $raw }
        Get-Codes (Invoke-Check $p) | Should -Contain 'HDR-UNIQUE'
        $p2 = New-DamagedCopy 'bidnextp' { param($raw) [BitConverter]::GetBytes([uint64]1).CopyTo($raw, 32); Set-HeaderCrc $raw }
        Get-Codes (Invoke-Check $p2) | Should -Contain 'HDR-BIDNEXTP'
    }
    It 'BID-RESERVED-BIT: bit 0 of a BID is reserved (readers ignore it, the reference still resolves)' {
        $p = New-DamagedCopy 'bidbit' {
            param($raw)
            $leaf = Get-LeafPage $raw ([BitConverter]::ToInt64($raw, 224))
            $raw[$leaf + 8] = $raw[$leaf + 8] -bor 1
            Set-PageCrc $raw $leaf
        }
        $c = Get-Codes (Invoke-Check $p)
        $c | Should -Contain 'BID-RESERVED-BIT'
        $c | Should -Not -Contain 'BID-DANGLING'
    }
    It 'detects <Fault>: <Code> (rule from the specification or from the behaviour of Outlook: a warning)' -ForEach @(
        @{ Fault = 'message-size-wrong'; Code = 'MSG-SIZE-COMPUTED' }
        @{ Fault = 'attach-size-wrong'; Code = 'ATTACH-SIZE' }
        @{ Fault = 'embedded-nid-type'; Code = 'ATTACH-EMBEDDED-NID' }
        @{ Fault = 'contents-cell-stale'; Code = 'CONTENTS-CELL' }
        @{ Fault = 'hierarchy-row-value'; Code = 'HIERARCHY-ROW-VALUE' }
        @{ Fault = 'hidden-count-wrong'; Code = 'FOLDER-HIDDEN-COUNT' }
    ) {
        $p = New-TestPst "rule-$Fault" -Faults @($Fault)
        $r = Invoke-Check $p
        Get-Codes $r | Should -Contain $Code
        ($r.Findings | Where-Object Code -EQ $Code | Select-Object -First 1).Severity | Should -Be 'Warning'
        $r.ExitCode | Should -Be 2
    }
    It 'reports <Fault>: <Code> as information only (heuristic of an old scan tool, neither in the specification nor needed by Outlook)' -ForEach @(
        @{ Fault = 'template-column-missing'; Code = 'TEMPLATE-COLUMN' }
        @{ Fault = 'no-outgoing-queue'; Code = 'MSG-OUTGOING-QUEUE' }
        @{ Fault = 'no-receive-row'; Code = 'RECEIVE-TABLE-ROW' }
        @{ Fault = 'no-template-6b6'; Code = 'MSG-TEMPLATE-MISSING' }
        @{ Fault = 'no-ec1'; Code = 'MSG-NODE-EC1' }
        @{ Fault = 'no-hmp'; Code = 'HMP-MISSING' }
        @{ Fault = 'hmp-wrong-entry'; Code = 'HMP-ENTRY' }
    ) {
        $p = New-TestPst "heur-$Fault" -Faults @($Fault)
        $r = Invoke-Check $p
        $f = $r.Findings | Where-Object Code -EQ $Code | Select-Object -First 1
        $f | Should -Not -BeNullOrEmpty
        $f.Severity | Should -Be 'Info'
        $f.Message | Should -Match 'heuristic'
        $r.ExitCode | Should -Be 0
    }
}

Describe 'file name rule: a PST named like a Windows device (FILE-RESERVED-NAME)' {
    It 'warns for aux.pst, con.v2.pst and the other device names (opens in our tools, crashes Outlook), and for them only' {
        $dir = Join-Path $TestDrive 'devnames'
        $null = New-Item -ItemType Directory -Path $dir
        try {
            foreach ($n in 'aux.pst', 'CON.PST', 'nul.v2.pst', 'com3.pst') {
                $p = Join-Path $dir $n
                [IO.File]::Copy($script:Good, $p, $true)
                $r = Invoke-Check $p
                $f = @($r.Findings | Where-Object Code -EQ 'FILE-RESERVED-NAME')
                $f.Count | Should -Be 1 -Because $n
                $f[0].Severity | Should -Be 'Warning'
                $f[0].Message | Should -Match 'Outlook crashes'
                $f[0].Message | Should -Match ([regex]::Escape([MailboxTransfer.FileNaming]::SafeFileName($n)))
                $r.Errors | Should -Be 0
                $r.Warnings | Should -Be 1
                $r.ExitCode | Should -Be 2
                [IO.File]::Delete($p)
            }
            foreach ($n in 'aux_.pst', 'aux-archive.pst', 'com10.pst') {
                $p = Join-Path $dir $n
                [IO.File]::Copy($script:Good, $p, $true)
                Get-Codes (Invoke-Check $p) | Should -Not -Contain 'FILE-RESERVED-NAME'
                [IO.File]::Delete($p)
            }
        } finally {
            foreach ($n in 'aux.pst', 'CON.PST', 'nul.v2.pst', 'com3.pst') { try { [IO.File]::Delete((Join-Path $dir $n)) } catch { } }
        }
    }
    It 'Test-MbtPst reports it with the other findings' {
        Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
        $dir = Join-Path $TestDrive 'devnames2'
        $null = New-Item -ItemType Directory -Path $dir
        $p = Join-Path $dir 'aux.pst'
        try {
            [IO.File]::Copy($script:Good, $p, $true)
            $r = Test-MbtPst -Path $p
            $r.ExitCode | Should -Be 2
            @($r.Findings | Where-Object Code -EQ 'FILE-RESERVED-NAME').Count | Should -Be 1
        } finally { try { [IO.File]::Delete($p) } catch { } }
    }
}
