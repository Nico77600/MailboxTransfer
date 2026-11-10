#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - the PST repair on the real corpora (git-ignored, written from the FastTransfer items of the lab mailbox):
    out\pst-validation\corpus2.pst (23 varied items: mail, contacts, calendar with a recurring series) is damaged in every way of
    tests\support\PstDamage.ps1 and repaired at each level; out\pst-validation\corpus.pst (16 items, 1.3 GB) is repaired once clean and
    twice damaged when MBT_REPAIR_BIG_TEST=1 (it takes minutes and 4 GB of disk). Skipped when the corpus is absent.

    Run:  Invoke-Pester -Path .\tests\PstRepair.Corpus.Tests.ps1 -Output Detailed
#>

BeforeDiscovery {
    $repo = Split-Path $PSScriptRoot -Parent
    $script:HaveSmall = Test-Path (Join-Path $repo 'out\pst-validation\corpus2.pst')
    $script:HaveBig = (Test-Path (Join-Path $repo 'out\pst-validation\corpus.pst')) -and $env:MBT_REPAIR_BIG_TEST -eq '1'
}

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\PstDamage.ps1')
    $script:Small = Join-Path $script:RepoRoot 'out\pst-validation\corpus2.pst'
    $script:Big = Join-Path $script:RepoRoot 'out\pst-validation\corpus.pst'

    function Invoke-Repair {
        param([string]$Path, [string]$Level = 'Recompute', [switch]$NoEscalate, [switch]$NoVerify, [string]$Dir = $TestDrive)
        $o = [MailboxTransfer.PstRepairOptions]::new()
        $o.Level = $Level; $o.Escalate = -not $NoEscalate; $o.Overwrite = $true
        $out = Join-Path $Dir ([IO.Path]::GetFileNameWithoutExtension($Path) + "-fixed-$Level.pst")
        $r = [MailboxTransfer.PstRepair]::Repair($Path, $out, $o)
        $v = $null
        if (-not $r.Error -and -not $r.Cancelled -and -not $NoVerify) { $v = [MailboxTransfer.TestSupport.PstRepairVerifier]::Verify($Path, $out, $r) }
        [pscustomobject]@{ Report = $r; Out = $out; Verify = $v }
    }
}

Describe 'repair of the real corpus2.pst (23 items)' -Skip:(-not $script:HaveSmall) {
    It 'at level <Level> a clean file gives an identical clean file' -ForEach @(
        @{ Level = 'Recompute' }, @{ Level = 'Rebuild' }, @{ Level = 'Salvage' }
    ) {
        $x = Invoke-Repair $script:Small -Level $Level
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsRecovered | Should -Be 23
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
        $x.Verify.Identical | Should -Be 23
        $x.Verify.Problems | Should -BeNullOrEmpty
    }
    It '<Kind> damage: repaired (starting at level <Level>, escalating if needed), clean result, every item that was readable is identical' -ForEach @(
        @{ Kind = 'highwater'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'unique'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'bidnextp'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'pmap'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'amap-bitorder'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'amap-unmarked'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'flip-folder'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'hdr-crc'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'page-sig-bbt'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'bid-dangling'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'overlap'; Level = 'Recompute'; Items = 23 }
        @{ Kind = 'truncated-tail'; Level = 'Recompute'; Items = 23 }
    ) {
        $p = New-DamagedPst $script:Small $TestDrive $Kind
        $x = Invoke-Repair $p -Level $Level
        $x.Report.Error | Should -BeNullOrEmpty
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsLost | Should -Be 0
        $x.Report.ItemsRecovered + $x.Report.ItemsPartial | Should -Be $Items
        $x.Report.After.Errors | Should -Be 0
        $x.Verify.Problems | Should -BeNullOrEmpty
    }
    It 'page-crc-nbt and nbt-order: pages of the node B-tree are damaged: all items found, the nodes of the lost page are lost' -ForEach @(@{ Kind = 'page-crc-nbt' }, @{ Kind = 'nbt-order' }) {
        $p = New-DamagedPst $script:Small $TestDrive $Kind
        $x = Invoke-Repair $p -Level Recompute
        $x.Report.Success | Should -BeTrue
        $x.Report.After.Errors | Should -Be 0
        $x.Report.ItemsRecovered + $x.Report.ItemsPartial | Should -BeGreaterOrEqual 20
    }
    It 'random damage, <Bytes> bytes changed per copy, Salvage, <Copies> copies: never fails, result clean, every readable item kept; recovery rate measured' -ForEach @(
        @{ Bytes = 5; Copies = 60; Min = 0.9 }
        @{ Bytes = 30; Copies = 100; Min = 0.6 }
    ) {
        $kept = [Collections.Generic.List[double]]::new(); $readable = [Collections.Generic.List[double]]::new(); $perfect = 0; $lost = 0
        for ($i = 0; $i -lt $Copies; $i++) {
            $p = Copy-Pst $script:Small $TestDrive "rd-$Bytes-$i"
            $raw = [IO.File]::ReadAllBytes($p); $rnd = [Random]::new(900 + $i)
            for ($k = 0; $k -lt $Bytes; $k++) { $raw[$rnd.Next(0, $raw.Length)] = [byte]$rnd.Next(0, 256) }
            [IO.File]::WriteAllBytes($p, $raw)
            $x = Invoke-Repair $p -Level Salvage
            $x.Report.Error | Should -BeNullOrEmpty
            $x.Report.After.Errors | Should -Be 0
            $x.Verify.Problems | Should -BeNullOrEmpty
            if ($x.Report.ItemsRecovered -eq 23) { $perfect++ }
            $lost += $x.Report.ItemsLost
            $kept.Add(($x.Report.ItemsRecovered + $x.Report.ItemsPartial) / 23.0)
            $readable.Add($x.Verify.Readable / 23.0)
            Remove-Item -LiteralPath $p, $x.Out -Force -ErrorAction SilentlyContinue
        }
        $mk = ($kept | Measure-Object -Average).Average; $mr = ($readable | Measure-Object -Average).Average
        Write-Host ("corpus2, Salvage, {0} copies with {1} random bytes changed: items kept (recovered or partial) {2:P1} on average (minimum {3:P1}), {4} copies fully recovered, items lost {5:P1}; items the ordinary reader can still read: {6:P1}" -f $Copies, $Bytes, $mk, ($kept | Measure-Object -Minimum).Minimum, $perfect, ($lost / (23.0 * $Copies)), $mr)
        $mk | Should -BeGreaterThan $Min
        $mk | Should -BeGreaterOrEqual $mr
    }
    It 'random damage with Recompute and escalation: the same guarantees' {
        $ok = 0
        for ($i = 0; $i -lt 30; $i++) {
            $p = New-DamagedPst $script:Small $TestDrive 'random-damage' -Seed (1900 + $i)
            $x = Invoke-Repair $p -Level Recompute
            $x.Report.Error | Should -BeNullOrEmpty
            $x.Report.After.Errors | Should -Be 0
            $x.Verify.Problems | Should -BeNullOrEmpty
            Remove-Item -LiteralPath $p, $x.Out -Force -ErrorAction SilentlyContinue
        }
    }
}

Describe 'repair of the real corpus.pst (1.3 GB, 16 items with attachments of 100 MiB), opt-in: MBT_REPAIR_BIG_TEST=1' -Skip:(-not $script:HaveBig) {
    It 'clean file, level Recompute: 16 items, identical' {
        $x = Invoke-Repair $script:Big -Level Recompute
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsRecovered | Should -Be 16
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
        $x.Verify.Identical | Should -Be 16
        Remove-Item -LiteralPath $x.Out -Force
    }
    It 'a copy with the header and a node B-tree page damaged and 200 random bytes changed: Salvage keeps what is intact' {
        $p = Join-Path $TestDrive 'big-damaged.pst'
        Copy-Item -LiteralPath $script:Big -Destination $p
        $fs = [IO.File]::Open($p, 'Open', 'ReadWrite')
        try {
            $rnd = [Random]::new(77)
            for ($k = 0; $k -lt 200; $k++) { $pos = $rnd.NextInt64(0x4400, $fs.Length); $fs.Position = $pos; $b = $fs.ReadByte(); $fs.Position = $pos; $fs.WriteByte([byte]($b -bxor 0xA5)) }
            $fs.Position = 20; $fs.WriteByte(0xFF)
        } finally { $fs.Dispose() }
        $x = Invoke-Repair $p -Level Recompute
        $x.Report.Success | Should -BeTrue
        $x.Report.Used | Should -Be 'Salvage'
        $x.Report.After.Errors | Should -Be 0
        $x.Verify.Problems | Should -BeNullOrEmpty
        Write-Host ("corpus.pst damaged: {0} recovered, {1} partial, {2} lost of 16" -f $x.Report.ItemsRecovered, $x.Report.ItemsPartial, $x.Report.ItemsLost)
        Remove-Item -LiteralPath $p, $x.Out -Force
    }
}
