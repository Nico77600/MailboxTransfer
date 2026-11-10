#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the PST repair (Engine.PstRepair.cs, Engine.PstScan.cs, Repair-MbtPst).
    Damaged copies of a clean PST (the generators of tests\support\PstDamage.ps1, the same as the checker tests) are repaired into a
    NEW file: the result opens with our reader, our checker reports no error, and every item the ordinary reader could read from the
    damaged file is still there, in the same folder, identical (PstFxCompare). The source is never modified. Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\PstRepair.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\PstDamage.ps1')

    function Invoke-Repair {
        param([string]$Path, [string]$Level = 'Recompute', [switch]$NoEscalate, [switch]$NoVerify)
        $o = [MailboxTransfer.PstRepairOptions]::new()
        $o.Level = $Level; $o.Escalate = -not $NoEscalate; $o.Overwrite = $true
        $out = Join-Path $TestDrive ([IO.Path]::GetFileNameWithoutExtension($Path) + "-fixed-$Level.pst")
        $r = [MailboxTransfer.PstRepair]::Repair($Path, $out, $o)
        $v = $null
        if (-not $r.Error -and -not $r.Cancelled -and -not $NoVerify) { $v = [MailboxTransfer.TestSupport.PstRepairVerifier]::Verify($Path, $out, $r) }
        [pscustomobject]@{ Report = $r; Out = $out; Verify = $v }
    }

    function Get-ActionCodes { param($Report) @($Report.Actions | ForEach-Object { $_.Code } | Sort-Object -Unique) }

    $script:Good = New-TestPst $TestDrive 'good'
    $script:GoodHash = (Get-FileHash -LiteralPath $script:Good).Hash
}

Describe 'Repair, level Recompute' {
    It 'a clean PST gives an identical clean PST, with all 56 items, and the source is untouched' {
        $x = Invoke-Repair $script:Good
        $x.Report.Success | Should -BeTrue
        $x.Report.Error | Should -BeNullOrEmpty
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.ItemsLost | Should -Be 0
        $x.Report.After.Errors | Should -Be 0
        $x.Report.After.Warnings | Should -Be 0
        $x.Report.Before.Errors | Should -Be 0
        $x.Verify.Readable | Should -Be 56
        $x.Verify.Identical | Should -Be 56
        $x.Verify.Problems | Should -BeNullOrEmpty
        $x.Report.After.Stats['folders'] | Should -Be 7
        $x.Report.After.Stats['attachments'] | Should -Be 8
        (Get-FileHash -LiteralPath $script:Good).Hash | Should -Be $script:GoodHash
        Test-Path "$($x.Out).repair-tmp" | Should -BeFalse
    }

    It '<Fault>: the structure is recomputed, every item is recovered, the checker is clean, the action is reported' -ForEach @(
        @{ Fault = 'wrong-count'; Action = 'RECOMPUTE-COUNTS'; Event = 'RECOUNT-FOLDER' }
        @{ Fault = 'hidden-count-wrong'; Action = 'RECOMPUTE-COUNTS'; Event = $null }
        @{ Fault = 'hierarchy-row-value'; Action = 'RECOMPUTE-COUNTS'; Event = $null }
        @{ Fault = 'contents-cell-stale'; Action = 'RECOMPUTE-COUNTS'; Event = $null }
        @{ Fault = 'message-size-wrong'; Action = 'RECOMPUTE-SIZES'; Event = $null }
        @{ Fault = 'attach-size-wrong'; Action = 'RECOMPUTE-SIZES'; Event = $null }
        @{ Fault = 'embedded-nid-type'; Action = 'RECOVER-PARTIAL'; Event = $null }
        @{ Fault = 'attach-highwater'; Action = 'REBUILD-HEADER'; Event = $null }
    ) {
        $p = New-TestPst $TestDrive "f-$Fault" -Faults @($Fault)
        $x = Invoke-Repair $p
        $x.Report.Success | Should -BeTrue
        $x.Report.Before.ExitCode | Should -BeGreaterThan 0
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.After.Errors | Should -Be 0
        $x.Report.After.Warnings | Should -Be 0
        Get-ActionCodes $x.Report | Should -Contain $Action
        ($x.Report.Actions | Where-Object Code -EQ $Action).Status | Should -Be 'Applied'
        if ($Event) { Get-ActionCodes $x.Report | Should -Contain $Event }
        $x.Verify.Problems | Should -BeNullOrEmpty
        $x.Verify.Identical | Should -Be $x.Verify.Readable
    }

    It 'a message without recipient table or with a short attachment table is read as it is and rewritten consistently' -ForEach @(
        @{ Fault = 'no-recipient-table' }, @{ Fault = 'attach-table-short' }
    ) {
        $p = New-TestPst $TestDrive "g-$Fault" -Faults @($Fault)
        $x = Invoke-Repair $p
        $x.Report.Success | Should -BeTrue
        $x.Report.After.Errors | Should -Be 0
        $x.Verify.Problems | Should -BeNullOrEmpty
        $x.Report.ItemsLost | Should -Be 0
    }

    It 'dangling-nid: the contents row of a node that does not exist is dropped and reported, nothing else changes' {
        $p = New-TestPst $TestDrive 'f-dangling' -Faults @('dangling-nid')
        $x = Invoke-Repair $p
        $x.Report.Success | Should -BeTrue
        $x.Report.DanglingDropped | Should -BeGreaterThan 0
        $x.Report.ItemsLost | Should -Be 0
        $x.Report.ItemsRecovered | Should -Be 56
        Get-ActionCodes $x.Report | Should -Contain 'DANGLING-ROW'
        Get-ActionCodes $x.Report | Should -Contain 'DROP-DANGLING'
        $x.Report.After.Errors | Should -Be 0
    }

    It '<Fault>: a message the tables do not list cannot be found at this level: it is reported as needing Rebuild, and the rest is clean' -ForEach @(
        @{ Fault = 'contents-row-removed'; Items = 55 }
        @{ Fault = 'hierarchy-missing'; Items = 41 }
    ) {
        $p = New-TestPst $TestDrive "f-$Fault" -Faults @($Fault)
        $x = Invoke-Repair $p
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsRecovered | Should -BeLessThan 56
        $x.Report.After.Errors | Should -Be 0
        ($x.Report.Actions | Where-Object { $_.Code -eq 'REBUILD-TABLES' }).Status | Should -Be 'NotApplied'
        $x.Verify.Problems | Should -BeNullOrEmpty
    }

    It '<Kind>: the damage is handled at this level (no escalation), the damaged item is partial or lost and the others identical' -ForEach @(
        @{ Kind = 'flip-message'; Lost = 1; Partial = 0 }
        @{ Kind = 'flip-attach'; Lost = 0; Partial = 1 }
        @{ Kind = 'xblock-total'; Lost = 0; Partial = 1 }
        @{ Kind = 'flip-folder'; Lost = 0; Partial = 0 }
        @{ Kind = 'highwater'; Lost = 0; Partial = 0 }
        @{ Kind = 'unique'; Lost = 0; Partial = 0 }
        @{ Kind = 'bidnextp'; Lost = 0; Partial = 0 }
        @{ Kind = 'pmap'; Lost = 0; Partial = 0 }
        @{ Kind = 'amap-unmarked'; Lost = 0; Partial = 0 }
        @{ Kind = 'amap-bitorder'; Lost = 0; Partial = 0 }
    ) {
        $p = New-DamagedPst $script:Good $TestDrive $Kind
        $x = Invoke-Repair $p -NoEscalate
        $x.Report.Error | Should -BeNullOrEmpty
        $x.Report.Success | Should -BeTrue
        $x.Report.Used | Should -Be 'Recompute'
        $x.Report.ItemsLost | Should -Be $Lost
        $x.Report.ItemsPartial | Should -Be $Partial
        $x.Report.After.Errors | Should -Be 0
        $x.Report.After.Warnings | Should -Be 0
        $x.Verify.Problems | Should -BeNullOrEmpty
        $x.Verify.Identical | Should -Be $x.Verify.Readable
        $x.Report.Before.Errors + $x.Report.Before.Warnings | Should -BeGreaterThan 0
    }

    It 'a partial item says which part is lost (the attachment data) and keeps the rest of the attachment' {
        $p = New-DamagedPst $script:Good $TestDrive 'flip-attach'
        $x = Invoke-Repair $p
        $it = $x.Report.Items | Where-Object Status -EQ 'Partial'
        @($it).Count | Should -Be 1
        $it.Lost[0] | Should -Match 'attachment 0 property 0x37010102'
        $it.Subject | Should -Be 'big'
        $r = [MailboxTransfer.PstReader]::Open($x.Out)
        try {
            $f = $r.AllFolders() | Where-Object Name -EQ 'C'
            $m = $r.ReadMessage(($r.MessageNids($f, $false))[0])
            $m.Attachments.Count | Should -Be 1
            ($m.Attachments[0].Properties | Where-Object { $_.Id -eq 0x3707 }).Raw.Length | Should -BeGreaterThan 0   # the file name is kept
        } finally { $r.Dispose() }
    }

    It 'counts folders again: the stored count of the folder with the lost message is corrected, each change is an action with its folder' {
        $p = New-DamagedPst $script:Good $TestDrive 'flip-message'
        $x = Invoke-Repair $p
        $ev = $x.Report.Actions | Where-Object Code -EQ 'RECOUNT-FOLDER'
        @($ev | Where-Object { $_.Message -match '40 -> 39' -and $_.Where -match '/A$' }).Count | Should -Be 1
        $r = [MailboxTransfer.PstReader]::Open($x.Out)
        try { ($r.AllFolders() | Where-Object Name -EQ 'A').ContentCount | Should -Be 39 } finally { $r.Dispose() }
    }
}

Describe 'Repair, safety' {
    It 'never modifies the source, refuses the source as target, refuses an existing target without Overwrite, leaves no temporary file' {
        $p = New-DamagedPst $script:Good $TestDrive 'flip-message' -Seed 5
        $h = (Get-FileHash -LiteralPath $p).Hash
        $o = [MailboxTransfer.PstRepairOptions]::new()
        $r = [MailboxTransfer.PstRepair]::Repair($p, $p, $o)
        $r.Success | Should -BeFalse; $r.Error | Should -Match 'never modifies the source'
        $t = Join-Path $TestDrive 'exists.pst'; Set-Content $t 'x'
        $r = [MailboxTransfer.PstRepair]::Repair($p, $t, $o)
        $r.Success | Should -BeFalse; $r.Error | Should -Match 'exists'
        (Get-Content $t -Raw).Trim() | Should -Be 'x'
        $r = [MailboxTransfer.PstRepair]::Repair((Join-Path $TestDrive 'nope.pst'), (Join-Path $TestDrive 'o.pst'), $o)
        $r.Success | Should -BeFalse
        (Get-FileHash -LiteralPath $p).Hash | Should -Be $h
        Test-Path "$t.repair-tmp" | Should -BeFalse
    }
    It 'a file that cannot be repaired at the requested level without escalation gives an error report, no target and no exception' -ForEach @(
        @{ Kind = 'hdr-crc' }, @{ Kind = 'page-crc-nbt' }, @{ Kind = 'zeros' }, @{ Kind = 'tiny' }, @{ Kind = 'empty' }
    ) {
        $p = New-DamagedPst $script:Good $TestDrive $Kind
        $x = Invoke-Repair $p -NoEscalate -NoVerify
        $x.Report.Success | Should -BeFalse
        $x.Report.Error | Should -Not -BeNullOrEmpty
        Test-Path $x.Out | Should -BeFalse
        Test-Path "$($x.Out).repair-tmp" | Should -BeFalse
    }
    It 'a cancelled repair stops, reports it and leaves no file' {
        $cts = [Threading.CancellationTokenSource]::new(); $cts.Cancel()
        $out = Join-Path $TestDrive 'cancelled.pst'
        $r = [MailboxTransfer.PstRepair]::Repair($script:Good, $out, [MailboxTransfer.PstRepairOptions]::new(), $null, $cts.Token)
        $r.Cancelled | Should -BeTrue
        Test-Path $out | Should -BeFalse
        Test-Path "$out.repair-tmp" | Should -BeFalse
    }
    It 'reports progress (phase, items, total) and sends every item to the sink' {
        $items = [Collections.Generic.List[object]]::new()
        $o = [MailboxTransfer.PstRepairOptions]::new(); $o.Overwrite = $true; $o.ItemSink = [Action[MailboxTransfer.PstRepairItem]]{ param($i) $items.Add($i) }
        $o.CheckBefore = $false; $o.CheckAfter = $false
        $pr = [MailboxTransfer.TestSupport.SyncProgress[MailboxTransfer.PstProgress]]::new()
        $out = Join-Path $TestDrive 'sink.pst'
        $r = [MailboxTransfer.PstRepair]::Repair($script:Good, $out, $o, $pr, [Threading.CancellationToken]::None)
        $r.Success | Should -BeTrue
        $items.Count | Should -Be 56
        $r.Before | Should -BeNullOrEmpty
        $pr.Reports.Count | Should -BeGreaterThan 1
        ($pr.Reports | Select-Object -Last 1).Phase | Should -Be 'repair'
    }
    It 'random damage (40 copies, 30 bytes each) at level Recompute without escalation: never an exception; when it succeeds the result is clean and keeps what was readable' {
        $raw0 = [IO.File]::ReadAllBytes($script:Good)
        $ok = 0
        for ($i = 0; $i -lt 40; $i++) {
            $p = New-DamagedPst $script:Good $TestDrive 'random-damage' -Seed (100 + $i)
            $x = Invoke-Repair $p -NoEscalate
            if ($x.Report.Success) {
                $ok++
                $x.Report.After.Errors | Should -Be 0
                $x.Verify.Problems | Should -BeNullOrEmpty
            } else {
                $x.Report.Error | Should -Not -BeNullOrEmpty
            }
            Remove-Item -LiteralPath $p, $x.Out -Force -ErrorAction SilentlyContinue
        }
        $ok | Should -BeGreaterThan 20
    }
}

Describe 'Repair plan (what-if)' {
    It 'lists the actions from the check without writing, with the lowest level that deals with every error' -ForEach @(
        @{ Fault = 'wrong-count'; Level = 'Recompute'; Action = 'RECOMPUTE-COUNTS' }
        @{ Fault = 'contents-row-removed'; Level = 'Rebuild'; Action = 'REBUILD-TABLES' }
        @{ Fault = 'hierarchy-missing'; Level = 'Rebuild'; Action = 'REBUILD-TABLES' }
        @{ Fault = 'dangling-nid'; Level = 'Recompute'; Action = 'DROP-DANGLING' }
    ) {
        $p = New-TestPst $TestDrive "plan-$Fault" -Faults @($Fault)
        $r = [MailboxTransfer.PstRepair]::Plan($p, [MailboxTransfer.PstRepairOptions]::new(), $null, [Threading.CancellationToken]::None)
        $r.PlanOnly | Should -BeTrue
        $r.Success | Should -BeTrue
        $r.Recommended | Should -Be $Level
        Get-ActionCodes $r | Should -Contain $Action
        @($r.Actions | Where-Object { $_.Status -notin 'Planned', 'NotApplied' }).Count | Should -Be 0
        $r.TargetPath | Should -BeNullOrEmpty
    }
    It 'recommends Salvage for a damaged header or B-tree' -ForEach @(@{ Kind = 'hdr-crc' }, @{ Kind = 'page-crc-nbt' }, @{ Kind = 'truncated' }) {
        $p = New-DamagedPst $script:Good $TestDrive $Kind
        $r = [MailboxTransfer.PstRepair]::Plan($p, [MailboxTransfer.PstRepairOptions]::new(), $null, [Threading.CancellationToken]::None)
        $r.Recommended | Should -Be 'Salvage'
    }
    It 'a clean file needs no action' {
        $r = [MailboxTransfer.PstRepair]::Plan($script:Good, [MailboxTransfer.PstRepairOptions]::new(), $null, [Threading.CancellationToken]::None)
        $r.Actions.Count | Should -Be 0
        $r.Recommended | Should -Be 'Recompute'
    }
}

Describe 'Repair-MbtPst (the public function)' {
    BeforeAll { Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force }
    It 'repairs into a new file and returns the structured result' {
        $p = New-DamagedPst $script:Good $TestDrive 'flip-message' -Seed 9
        $out = Join-Path $TestDrive 'cmdlet-out.pst'
        $r = Repair-MbtPst -Path $p -Destination $out -Level Recompute
        $r.Success | Should -BeTrue
        $r.ItemsLost | Should -Be 1
        $r.ItemsRecovered | Should -Be 55
        $r.After.Errors | Should -Be 0
        $r.Actions.Count | Should -BeGreaterThan 0
        (Test-MbtPst -Path $out).ExitCode | Should -Be 0
    }
    It '-WhatIf shows the planned actions and writes nothing' {
        $p = New-DamagedPst $script:Good $TestDrive 'page-crc-nbt' -Seed 9
        $out = Join-Path $TestDrive 'cmdlet-whatif.pst'
        $r = Repair-MbtPst -Path $p -Destination $out -WhatIf
        $r.PlanOnly | Should -BeTrue
        $r.Recommended | Should -Be 'Salvage'
        Test-Path $out | Should -BeFalse
    }
    It 'refuses an existing destination unless -Force' {
        $out = Join-Path $TestDrive 'cmdlet-exists.pst'; Set-Content $out 'x'
        (Repair-MbtPst -Path $script:Good -Destination $out).Success | Should -BeFalse
        (Repair-MbtPst -Path $script:Good -Destination $out -Force).Success | Should -BeTrue
    }
}

Describe 'Repair, level Rebuild' {
    BeforeAll {
        $script:NoHier = New-TestPst $TestDrive 'nohier' -Faults @('hierarchy-missing')
        function Invoke-Check { param([string]$Path) [MailboxTransfer.PstChecker]::Check($Path, [MailboxTransfer.PstCheckOptions]::new()) }
        function Get-PstTree {
            param([string]$Path)
            $r = [MailboxTransfer.PstReader]::Open($Path)
            try {
                $h = @{}
                foreach ($f in $r.AllFolders()) { $h["/$($f.Path)"] = @($r.MessageNids($f, $false)).Count }
                $h
            } finally { $r.Dispose() }
        }
        function Get-Subjects {
            param([string]$Path, [string]$Folder)
            $r = [MailboxTransfer.PstReader]::Open($Path)
            try {
                $f = $r.AllFolders() | Where-Object { "/$($_.Path)" -eq $Folder }
                @($r.MessageNids($f, $false) | ForEach-Object { [Text.Encoding]::Unicode.GetString(($r.ReadMessage($_).Properties | Where-Object { $_.Id -eq 0x0037 }).Raw).TrimEnd([char]0) })
            } finally { $r.Dispose() }
        }
    }

    It 'a clean PST rebuilt by the nodes is identical to the source, with the same folders' {
        $x = Invoke-Repair $script:Good -Level Rebuild
        $x.Report.Success | Should -BeTrue
        $x.Report.Used | Should -Be 'Rebuild'
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
        $x.Verify.Identical | Should -Be 56
        $x.Verify.Problems | Should -BeNullOrEmpty
    }
    It '<Fault>: the messages and folders the tables do not list are all there, in their folder, counts correct' -ForEach @(
        @{ Fault = 'contents-row-removed' }, @{ Fault = 'hierarchy-missing' }
    ) {
        $p = New-TestPst $TestDrive "l2-$Fault" -Faults @($Fault)
        $x = Invoke-Repair $p -Level Rebuild
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
        $x.Verify.Problems | Should -BeNullOrEmpty
        $t = Get-PstTree $x.Out
        $t['/Checked/A'] | Should -Be 40
        $t['/Checked/A/B'] | Should -Be 15
        $t['/Checked/C'] | Should -Be 1
        ($x.Report.Actions | Where-Object Code -EQ 'REBUILD-TABLES').Status | Should -Be 'Applied'
        $t.Keys | Where-Object { $_ -match 'Recovered' } | Should -BeNullOrEmpty
    }
    It 'folder-orphan with the hierarchy intact: the table that lists the folder says where it was, it is re-attached there' {
        $p = New-DamagedPst $script:Good $TestDrive 'folder-orphan'
        $x = Invoke-Repair $p -Level Rebuild -NoVerify
        $x.Report.Success | Should -BeTrue
        $t = Get-PstTree $x.Out
        $t['/Checked/A/B'] | Should -Be 15
        $t.Keys | Where-Object { $_ -match 'Recovered' } | Should -BeNullOrEmpty
        Get-ActionCodes $x.Report | Should -Contain 'FOLDER-REATTACHED'
        $x.Report.After.Errors | Should -Be 0
    }
    It 'folder-orphan without any table that lists it: the folder and its messages are kept under Recovered items' {
        $p = New-DamagedPst $script:NoHier $TestDrive 'folder-orphan' -NidSource $script:Good
        $x = Invoke-Repair $p -Level Rebuild -NoVerify
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.FoldersRecovered | Should -Be 1
        $t = Get-PstTree $x.Out
        $t['/Checked/Recovered items'] | Should -Be 0
        $t['/Checked/Recovered items/B'] | Should -Be 15
        $t['/Checked/A'] | Should -Be 40
        Get-ActionCodes $x.Report | Should -Contain 'FOLDER-ORPHAN'
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
    }
    It 'folder-moved: the parent given by the node wins over the table' {
        $p = New-DamagedPst $script:Good $TestDrive 'folder-moved'
        $x = Invoke-Repair $p -Level Rebuild -NoVerify
        $x.Report.Success | Should -BeTrue
        $t = Get-PstTree $x.Out
        $t['/Checked/C/B'] | Should -Be 15
        $t.ContainsKey('/Checked/A/B') | Should -BeFalse
        $x.Report.ItemsRecovered | Should -Be 56
    }
    It 'folder-loop and folder-self: a loop in the hierarchy is cut: with a table that lists the folder it is restored, without one it goes to Recovered items' {
        $p = New-DamagedPst $script:Good $TestDrive 'folder-loop'
        $x = Invoke-Repair $p -Level Rebuild -NoVerify
        $x.Report.Success | Should -BeTrue
        (Get-PstTree $x.Out)['/Checked/A/B'] | Should -Be 15
        $x.Report.ItemsRecovered | Should -Be 56
        $p2 = New-DamagedPst $script:NoHier $TestDrive 'folder-self' -Seed 2 -NidSource $script:Good
        $x2 = Invoke-Repair $p2 -Level Rebuild -NoVerify
        $x2.Report.Success | Should -BeTrue
        $x2.Report.FoldersRecovered | Should -BeGreaterThan 0
        $x2.Report.ItemsRecovered | Should -Be 56
        $x2.Report.After.Errors | Should -Be 0
        @((Get-PstTree $x2.Out).Keys | Where-Object { $_ -like '/Checked/Recovered items/*' }).Count | Should -BeGreaterThan 0
    }
    It '<Kind>: a message whose folder does not exist is kept under Recovered items, grouped by the folder it came from' -ForEach @(
        @{ Kind = 'msg-orphan'; Folder = '/Checked/Recovered items/Lost folder 0x8888' }
        @{ Kind = 'msg-orphan0'; Folder = '/Checked/Recovered items/Unassigned' }
    ) {
        $p = New-DamagedPst $script:Good $TestDrive $Kind
        $x = Invoke-Repair $p -Level Rebuild -NoVerify
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.OrphanItems | Should -Be 1
        $t = Get-PstTree $x.Out
        $t[$Folder] | Should -Be 1
        $t['/Checked/C'] | Should -Be 0
        Get-Subjects $x.Out $Folder | Should -Be @('big')
        Get-ActionCodes $x.Report | Should -Contain 'ORPHAN-RECOVERED'
        ($x.Report.Items | Where-Object Orphan).Subject | Should -Be 'big'
        $x.Report.After.Errors | Should -Be 0
    }
    It 'namemap-empty: properties whose id has no entry in the name map are dropped and listed, the named-property map of the result is consistent' {
        $p = New-TestPst $TestDrive 'l2-namemap' -Faults @('namemap-empty')
        (Invoke-Check $p).Errors | Should -BeGreaterThan 0
        $x = Invoke-Repair $p -Level Rebuild
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsPartial | Should -Be 56
        $x.Report.ItemsLost | Should -Be 0
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
        ($x.Report.Items | Select-Object -First 1).Lost[0] | Should -Match 'named property 0x8000 \(no entry in the name map\)'
        Get-ActionCodes $x.Report | Should -Contain 'REBUILD-NAMEMAP'
    }
    It 'the name map is rebuilt from what the messages use: the named property of a message comes back with its GUID and LID' {
        $x = Invoke-Repair $script:Good -Level Rebuild
        $r = [MailboxTransfer.PstReader]::Open($x.Out)
        try {
            $f = $r.AllFolders() | Where-Object Name -EQ 'A'
            $m = $r.ReadMessage(($r.MessageNids($f, $false))[0])
            $np = $m.Properties | Where-Object { $_.IsNamed }
            $np.NamedGuid | Should -Be ([guid]'11111111-2222-3333-4444-555555555555')
            $np.NamedLid | Should -Be 0x8501
        } finally { $r.Dispose() }
    }
    It 'level Rebuild also deals with what level Recompute does (counts, sizes) on the same file' {
        $p = New-TestPst $TestDrive 'l2-both' -Faults @('wrong-count', 'message-size-wrong', 'contents-row-removed')
        $x = Invoke-Repair $p -Level Rebuild
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
    }
}

Describe 'Repair, level Salvage' {
    It 'a clean PST salvaged from the scan of its blocks and pages is identical to the source' {
        $x = Invoke-Repair $script:Good -Level Salvage
        $x.Report.Success | Should -BeTrue
        $x.Report.Used | Should -Be 'Salvage'
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
        $x.Verify.Identical | Should -Be 56
        $x.Verify.Problems | Should -BeNullOrEmpty
        ($x.Report.Actions | Where-Object Code -EQ 'SALVAGE-SCAN').Message | Should -Match '707 valid blocks'
    }
    It '<Kind>: the B-trees or the header are damaged, the blocks are found by the scan, every item is recovered and identical' -ForEach @(
        @{ Kind = 'hdr-crc'; Recovered = 56 }
        @{ Kind = 'bid-dangling'; Recovered = 56 }
        @{ Kind = 'page-sig-bbt'; Recovered = 56 }
        @{ Kind = 'overlap'; Recovered = 56 }
        @{ Kind = 'truncated-tail'; Recovered = 56 }
        @{ Kind = 'highwater'; Recovered = 56 }
        @{ Kind = 'amap-bitorder'; Recovered = 56 }
    ) {
        $p = New-DamagedPst $script:Good $TestDrive $Kind
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.Success | Should -BeTrue
        $x.Report.Used | Should -Be 'Salvage'
        $x.Report.ItemsRecovered | Should -Be $Recovered
        $x.Report.ItemsLost | Should -Be 0
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
        $x.Verify.Problems | Should -BeNullOrEmpty
    }
    It 'page-crc-nbt: a lost leaf page of the node B-tree loses its nodes (here the name map: the named properties), the messages are all found' {
        $p = New-DamagedPst $script:Good $TestDrive 'page-crc-nbt'
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsRecovered + $x.Report.ItemsPartial | Should -Be 56
        $x.Report.ItemsLost | Should -Be 0
        $x.Report.After.Errors | Should -Be 0
    }
    It 'a damaged data block costs only its own item, as with the other levels' -ForEach @(
        @{ Kind = 'flip-message'; Lost = 1; Partial = 0 }, @{ Kind = 'flip-attach'; Lost = 0; Partial = 1 }
    ) {
        $p = New-DamagedPst $script:Good $TestDrive $Kind
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.ItemsLost | Should -Be $Lost
        $x.Report.ItemsPartial | Should -Be $Partial
        $x.Report.ItemsRecovered | Should -Be (56 - $Lost - $Partial)
        $x.Report.After.Errors | Should -Be 0
        $x.Verify.Problems | Should -BeNullOrEmpty
    }
    It 'escalation: a damaged header stops Recompute and Rebuild, the repair goes on with Salvage and says so' -ForEach @(@{ Kind = 'hdr-crc' }) {
        $p = New-DamagedPst $script:Good $TestDrive $Kind
        $x = Invoke-Repair $p -Level Recompute
        $x.Report.Success | Should -BeTrue
        $x.Report.Escalated | Should -BeTrue
        $x.Report.Requested | Should -Be 'Recompute'
        $x.Report.Used | Should -Be 'Salvage'
        @($x.Report.Actions | Where-Object Code -EQ 'ESCALATE').Count | Should -Be 2
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Verify.Problems | Should -BeNullOrEmpty
    }
    It 'truncated: the node B-tree at the end of the file is gone, the message property contexts are found by the scan and kept without folder, recipients and attachments' {
        $p = New-DamagedPst $script:Good $TestDrive 'truncated'
        $x = Invoke-Repair $p -Level Salvage -NoVerify
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsLost | Should -Be 0
        $x.Report.ItemsPartial | Should -BeGreaterOrEqual 50
        $x.Report.After.Errors | Should -Be 0
        Get-ActionCodes $x.Report | Should -Contain 'UNLINKED-MESSAGES'
        ($x.Report.Items | Where-Object Status -EQ 'Partial' | Select-Object -First 1).Lost[0] | Should -Match 'node entry lost'
        $r = [MailboxTransfer.PstReader]::Open($x.Out)
        try {
            $f = $r.AllFolders() | Where-Object Name -EQ 'Unassigned'
            @($r.MessageNids($f, $false)).Count | Should -Be $x.Report.ItemsPartial
            $subjects = @($r.MessageNids($f, $false) | ForEach-Object { [Text.Encoding]::Unicode.GetString(($r.ReadMessage($_).Properties | Where-Object { $_.Id -eq 0x0037 }).Raw).TrimEnd([char]0) })
            $subjects | Should -Contain 'a1'
            $subjects | Should -Contain 'b15'
        } finally { $r.Dispose() }
    }
    It 'a file of zeros, random bytes, a tiny file and an empty file cannot be salvaged: an error report, no file, no exception' -ForEach @(
        @{ Kind = 'zeros' }, @{ Kind = 'random' }, @{ Kind = 'tiny' }, @{ Kind = 'empty' }
    ) {
        $p = New-DamagedPst $script:Good $TestDrive $Kind
        $x = Invoke-Repair $p -Level Salvage -NoVerify
        $x.Report.Success | Should -BeFalse
        $x.Report.Error | Should -Not -BeNullOrEmpty
        Test-Path $x.Out | Should -BeFalse
    }
    It 'random damage, 60 copies with 30 bytes changed: salvage never fails, the result is clean and keeps every item that was readable; recovery rate' {
        $rates = [Collections.Generic.List[double]]::new()
        $l1ok = 0
        for ($i = 0; $i -lt 60; $i++) {
            $p = New-DamagedPst $script:Good $TestDrive 'random-damage' -Seed (500 + $i)
            $x = Invoke-Repair $p -Level Salvage
            $x.Report.Error | Should -BeNullOrEmpty
            $x.Report.Success | Should -BeTrue
            $x.Report.After.Errors | Should -Be 0
            $x.Verify.Problems | Should -BeNullOrEmpty
            $rates.Add(($x.Report.ItemsRecovered + $x.Report.ItemsPartial) / 56.0)
            Remove-Item -LiteralPath $p, $x.Out -Force -ErrorAction SilentlyContinue
        }
        $mean = ($rates | Measure-Object -Average).Average
        Write-Host ("salvage of 60 randomly damaged copies: mean recovery {0:P1}, minimum {1:P1}" -f $mean, ($rates | Measure-Object -Minimum).Minimum)
        $mean | Should -BeGreaterThan 0.9
    }
}

Describe 'Repair, name-to-id map recovery' {
    BeforeAll {
        # Messages with three named properties (two numbers, one name) so that two maps can be told apart by the ids they resolve.
        function New-NamedPst {
            param([string]$Name, [string[]]$Faults = @(), [int]$LongNames = 0)
            $path = Join-Path $TestDrive ($Name + '.pst')
            $P = [MailboxTransfer.FxProperty]
            $o = [MailboxTransfer.PstWriterOptions]::new(); $o.Faults = $Faults
            $w = [MailboxTransfer.PstWriter]::Create($path, 'Named', $o)
            $a = $w.CreateFolder($null, 'A', 'IPF.Note')
            foreach ($i in 1..12) {
                $m = New-Msg "n$i"
                $m.Properties.Add($P::FromInt32(0x8002, 7).Named([guid]'00062002-0000-0000-C000-000000000046', [uint32]0x820D))
                $m.Properties.Add($P::FromString(0x8003, 'blue').Named([guid]'00020329-0000-0000-C000-000000000046', 'Keywords'))
                if ($LongNames -gt 0) { foreach ($k in 1..$LongNames) { $m.Properties.Add($P::FromInt32(0x8010 + $k, $i).Named([guid]'00020329-0000-0000-C000-000000000046', ('long-name-' + $k + '-' + ('x' * 450)))) } }
                [void]$w.AddMessage($a, $m)
            }
            $w.Close()
            $path
        }
        function Get-NamedProps {
            param([string]$Path)
            $r = [MailboxTransfer.PstReader]::Open($Path)
            try {
                $f = $r.AllFolders() | Where-Object Name -EQ 'A'
                $m = $r.ReadMessage(($r.MessageNids($f, $false))[0])
                @($m.Properties | Where-Object { $_.IsNamed } | ForEach-Object { if ($null -ne $_.NamedLid) { '{0}/0x{1:X}' -f $_.NamedGuid, $_.NamedLid } else { '{0}/{1}' -f $_.NamedGuid, $_.NamedName } } | Sort-Object)
            } finally { $r.Dispose() }
        }
        $script:Named = New-NamedPst 'named'
        $script:NamedDecoy = New-NamedPst 'named-decoy' -Faults @('decoy-namemap')
        $script:AllNamed = @('00020329-0000-0000-c000-000000000046/Keywords', '00062002-0000-0000-c000-000000000046/0x820D', '11111111-2222-3333-4444-555555555555/0x8501')
    }

    It 'a clean file: the map is the original, nothing searched' {
        $x = Invoke-Repair $script:Named -Level Salvage
        $x.Report.NameMapSource | Should -Be 'Original'
        $x.Report.NameMapCandidates | Should -Be 0
        $x.Report.NamedIdsUnresolved | Should -Be 0
        $x.Report.ItemsRecovered | Should -Be 12
        Get-NamedProps $x.Out | Should -Be $script:AllNamed
    }
    It 'page-crc-nbt: the leaf page with the map is lost, Salvage finds the map in the blocks, every message keeps its named properties' {
        $p = New-DamagedPst $script:Good $TestDrive 'page-crc-nbt'
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.Success | Should -BeTrue
        $x.Report.NameMapSource | Should -Be 'RecoveredByScan'
        $x.Report.NameMapCandidates | Should -Be 1
        $x.Report.NamedIdsUnresolved | Should -Be 0
        $x.Report.NameMapDetail | Should -Match 'map found in the blocks of the file'
        Get-ActionCodes $x.Report | Should -Contain 'NAMEMAP-RECOVERED'
        $x.Report.ItemsRecovered | Should -Be 56
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
    }
    It 'namemap-ptr: the entry of the map points nowhere (a map with a subnode tree: the long strings): Salvage finds head and subnode tree, the named properties are all there' {
        $src = New-NamedPst 'named-long' -LongNames 6          # the string stream is long enough for a subnode
        $p = New-DamagedPst $src $TestDrive 'namemap-ptr'
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.NameMapSource | Should -Be 'RecoveredByScan'
        $x.Report.NameMapDetail | Should -Match 'subnodes 0x'
        $x.Report.ItemsRecovered | Should -Be 12
        $x.Report.NamedIdsUnresolved | Should -Be 0
        $x.Report.After.Errors | Should -Be 0
        @(Get-NamedProps $x.Out).Count | Should -Be 9
    }
    It 'namemap-ptr at level Rebuild: the node B-tree is intact but the map is unreachable, the file is scanned for it and the report says so' {
        $p = New-DamagedPst $script:Named $TestDrive 'namemap-ptr'
        $x = Invoke-Repair $p -Level Rebuild
        $x.Report.Used | Should -Be 'Rebuild'
        $x.Report.NameMapSource | Should -Be 'RecoveredByScan'
        Get-ActionCodes $x.Report | Should -Contain 'NAMEMAP-SCAN'
        $x.Report.ItemsRecovered | Should -Be 12
        Get-NamedProps $x.Out | Should -Be $script:AllNamed
    }
    It 'namemap-ptr at level Recompute: not searched at that level, the named properties are dropped and the report says what is lost' {
        $p = New-DamagedPst $script:Named $TestDrive 'namemap-ptr'
        $x = Invoke-Repair $p -Level Recompute -NoEscalate
        $x.Report.NameMapSource | Should -Be 'RebuiltFromUsage'
        $x.Report.NameMapDetail | Should -Match 'not searched at this level'
        $x.Report.NamedIdsUnresolved | Should -Be 3
        $x.Report.NamedPropsDropped | Should -Be 36
        $x.Report.ItemsPartial | Should -Be 12
        Get-NamedProps $x.Out | Should -BeNullOrEmpty
    }
    It 'two maps in the file: the one that resolves the most ids used by the messages is chosen, the choice is reported' {
        $p = New-DamagedPst $script:NamedDecoy $TestDrive 'namemap-ptr-decoy'
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.NameMapSource | Should -Be 'RecoveredByScan'
        $x.Report.NameMapCandidates | Should -Be 2
        $x.Report.NameMapDetail | Should -Match '2 candidate\(s\) \(ids resolved: 3, 1 of 3 used\)'
        $x.Report.NameMapEntries | Should -Be 3
        $x.Report.ItemsRecovered | Should -Be 12
        Get-NamedProps $x.Out | Should -Be $script:AllNamed
    }
    It 'namemap-data: one byte of the map itself is wrong (the CRC fails): the CRC gives the byte, the map is exact, nothing is lost' {
        $p = New-DamagedPst $script:Named $TestDrive 'namemap-data'
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.NameMapSource | Should -Be 'RecoveredByScan'
        $x.Report.NameMapDetail | Should -Match 'correcting a wrong byte in 1 block\(s\)'
        $x.Report.NameMapEntries | Should -Be 3
        $x.Report.NamedIdsUnresolved | Should -Be 0
        $x.Report.ItemsRecovered | Should -Be 12
        Get-NamedProps $x.Out | Should -Be $script:AllNamed
        $x.Report.After.Errors | Should -Be 0
    }
    It 'namemap-data at level Rebuild: the same, the node B-tree being intact' {
        $p = New-DamagedPst $script:Named $TestDrive 'namemap-data'
        $x = Invoke-Repair $p -Level Rebuild
        $x.Report.Used | Should -Be 'Rebuild'
        $x.Report.NameMapSource | Should -Be 'RecoveredByScan'
        $x.Report.NamedIdsUnresolved | Should -Be 0
        Get-NamedProps $x.Out | Should -Be $script:AllNamed
    }
    It 'namemap-byte: a wrong byte anywhere in the heap of the map (header, GUIDs, entries, buckets) is corrected exactly: never a lost or a wrong identity' {
        $cb = (Get-NodeBlock $script:Named 0x61).Cb
        foreach ($off in @(0, 1, 7, 13, 20, 33, 60, 100, 160, 211, [int]($cb / 2), ($cb - 20), ($cb - 1)) | Where-Object { $_ -ge 0 -and $_ -lt $cb } | Select-Object -Unique) {
            $p = New-DamagedPst $script:Named $TestDrive 'namemap-byte' -Seed $off
            $x = Invoke-Repair $p -Level Salvage -NoVerify
            $x.Report.NameMapSource | Should -Be 'RecoveredByScan' -Because "byte $off"
            $x.Report.NamedIdsUnresolved | Should -Be 0 -Because "byte $off"
            Get-NamedProps $x.Out | Should -Be $script:AllNamed -Because "byte $off"
            Remove-Item -LiteralPath $p, $x.Out -Force
        }
    }
    It 'namemap-two: two wrong bytes in the map: the CRC cannot say which, nothing is guessed, the named properties are dropped and counted' {
        $p = New-DamagedPst $script:Named $TestDrive 'namemap-two'
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.NameMapSource | Should -Be 'RebuiltFromUsage'
        $x.Report.NamedIdsUnresolved | Should -Be 3
        Get-NamedProps $x.Out | Should -BeNullOrEmpty
    }
    It 'PstCrcRepair: one wrong byte anywhere in a buffer is found and put back; two wrong bytes and a wrong stored CRC are left alone' {
        $rnd = [Random]::new(5)
        foreach ($len in 1, 2, 8, 13, 100, 5772, 8176) {
            $data = [byte[]]::new($len); $rnd.NextBytes($data)
            $crc = [MailboxTransfer.PstFormat]::Crc($data, 0, $len)
            foreach ($pos in 0, ($len - 1), [int]($len / 2)) {
                $bad = $data.Clone(); $bad[$pos] = $bad[$pos] -bxor 0x81
                [MailboxTransfer.PstCrcRepair]::TrySingleByte($bad, $len, $crc) | Should -BeTrue -Because "length $len byte $pos"
                [Convert]::ToHexString($bad) | Should -Be ([Convert]::ToHexString($data))
            }
            if ($len -ge 8) {
                $two = $data.Clone(); $two[1] = $two[1] -bxor 1; $two[$len - 2] = $two[$len - 2] -bxor 2
                [MailboxTransfer.PstCrcRepair]::TrySingleByte($two, $len, $crc) | Should -BeFalse
                [MailboxTransfer.PstCrcRepair]::TrySingleByte($data.Clone(), $len, $crc -bxor 0x10) | Should -BeFalse -Because 'the data is fine, the stored CRC is wrong'
            }
        }
    }
    It 'namemap-block: the header of the map is destroyed and nothing else in the file holds a map: rebuilt from what is usable, the loss is counted' {
        $p = New-DamagedPst $script:Named $TestDrive 'namemap-block'
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.Success | Should -BeTrue
        $x.Report.NameMapSource | Should -Be 'RebuiltFromUsage'
        $x.Report.NameMapCandidates | Should -Be 0
        $x.Report.NamedIdsUsed | Should -Be 3
        $x.Report.NamedIdsUnresolved | Should -Be 3
        $x.Report.NamedPropsDropped | Should -Be 36
        Get-ActionCodes $x.Report | Should -Contain 'NAMEMAP-REBUILT'
        Get-ActionCodes $x.Report | Should -Contain 'NAMEMAP-IDS-LOST'
        $x.Report.After.Errors | Should -Be 0
    }
    It 'namemap-strings: a byte of the string stream (a subnode, another block) is wrong: corrected exactly as well' {
        $src = New-NamedPst 'named-long2' -LongNames 6
        $p = New-DamagedPst $src $TestDrive 'namemap-strings'
        $x = Invoke-Repair $p -Level Salvage
        $x.Report.NameMapSource | Should -Be 'RecoveredByScan'
        $x.Report.NameMapDetail | Should -Match 'correcting a wrong byte in 1 block\(s\)'
        $x.Report.NamedIdsUnresolved | Should -Be 0
        $x.Report.ItemsRecovered | Should -Be 12
        @(Get-NamedProps $x.Out).Count | Should -Be 9
        $x.Report.After.Errors | Should -Be 0
    }
    It 'the entry stream cannot be read: the entries are rebuilt from the hash buckets' {
        $f = [MailboxTransfer.PstFile]::Open($script:Named)
        try {
            $pc = [MailboxTransfer.PstPcReader]::new($f.OpenNode([uint32]0x61))
            $errs = [Collections.Generic.List[string]]::new()
            $e = $pc.Entries | Where-Object Id -EQ 3
            $e.Hnid = 0x7777 -bor 0x1F       # a subnode that does not exist
            $m = [MailboxTransfer.PstNameMapReader]::Read($pc, $errs, $null)
        } finally { $f.Dispose() }
        $m.Count | Should -Be 3
        $errs -join ';' | Should -Match 'entry stream unreadable: 3 entries rebuilt from the hash buckets'
        @($m.Values | Where-Object Name).Name | Should -Be 'Keywords'
    }
}

Describe 'Repair, KeepDamagedAttachments' {
    BeforeAll {
        function Invoke-RepairKeep {
            param([string]$Path, [string]$Level = 'Recompute', [switch]$Keep)
            $o = [MailboxTransfer.PstRepairOptions]::new()
            $o.Level = $Level; $o.Overwrite = $true; $o.KeepDamagedAttachments = [bool]$Keep
            $out = Join-Path $TestDrive ([IO.Path]::GetFileNameWithoutExtension($Path) + "-keep-$([bool]$Keep)-$Level.pst")
            $r = [MailboxTransfer.PstRepair]::Repair($Path, $out, $o)
            [pscustomobject]@{ Report = $r; Out = $out }
        }
        # a block of the data tree of the attachment of the message in folder C: leaf n (0-based, across the XBLOCKs), a whole XBLOCK child of an XXBLOCK, or the root
        function Break-Tree {
            param([string]$Path, [int[]]$Leaf = @(), [int]$Child = -1, [switch]$Root)
            $t = Get-BigAttachmentTree $Path
            $raw = [IO.File]::ReadAllBytes($Path)
            try {
                $rootBid = $t.Node.BidData
                $x = $t.File.ReadBlock($rootBid)
                $targets = [Collections.Generic.List[uint64]]::new()
                if ($Root) { $targets.Add($rootBid) }
                elseif ($Child -ge 0) { $targets.Add([BitConverter]::ToUInt64($x, 8 + 8 * $Child)) }
                else {
                    foreach ($l in $Leaf) {
                        if ($x[1] -eq 2) { $xb = $t.File.ReadBlock([BitConverter]::ToUInt64($x, 8 + 8 * [math]::Floor($l / 1021))); $targets.Add([BitConverter]::ToUInt64($xb, 8 + 8 * ($l % 1021))) }
                        else { $targets.Add([BitConverter]::ToUInt64($x, 8 + 8 * $l)) }
                    }
                }
                foreach ($bid in $targets) {
                    $e = [MailboxTransfer.PstBbtEntry]::new(); [void]$t.File.TryFindBlock($bid, [ref]$e)
                    $raw[[long]$e.Ib + 77] = $raw[[long]$e.Ib + 77] -bxor 0x5A
                }
            } finally { $t.File.Dispose() }
            [IO.File]::WriteAllBytes($Path, $raw)
        }
        function Get-ResultAttachment {
            param([string]$Path)
            $r = [MailboxTransfer.PstReader]::Open($Path)
            try {
                $f = $r.AllFolders() | Where-Object Name -EQ 'C'
                $m = $r.ReadMessage(($r.MessageNids($f, $false))[0])
                $a = $m.Attachments[0]
                $str = { param($id) $p = $a.Properties | Where-Object { -not $_.IsNamed -and $_.Id -eq $id }; if ($p) { [Text.Encoding]::Unicode.GetString($p.Raw).TrimEnd([char]0) } }
                $d = $a.Properties | Where-Object { -not $_.IsNamed -and $_.Id -eq 0x3701 }
                $bytes = if (-not $d) { $null } elseif ($d.IsLazy) { $d.Lazy.ToArray() } else { $d.Raw }
                [pscustomobject]@{ Long = (& $str 0x3707); Display = (& $str 0x3001); Data = $bytes }
            } finally { $r.Dispose() }
        }
        function New-HugePst {
            $path = Join-Path $TestDrive 'huge.pst'
            $w = [MailboxTransfer.PstWriter]::Create($path, 'Huge')
            $c = $w.CreateFolder($null, 'C', 'IPF.Note')
            $m = New-Msg 'huge' -AttachBytes 9MB
            $a = $m.Attachments[0]
            $a.Properties.Add([MailboxTransfer.FxProperty]::FromString(0x3001, 'f.bin'))       # display name too
            [void]$w.AddMessage($c, $m)
            $w.Close()
            $path
        }
        function Get-Expected { param([int]$Length) $b = [byte[]]::new($Length); [Random]::new(1).NextBytes($b); , $b }
        # the bytes of $Got equal the expected ones outside the ranges (offset, length) and are zeros inside
        function Test-Ranges {
            param([byte[]]$Got, [byte[]]$Expected, $Ranges)
            if ($Got.Length -ne $Expected.Length) { return "length $($Got.Length) instead of $($Expected.Length)" }
            $bad = [bool[]]::new($Got.Length)
            foreach ($r in $Ranges) { for ($i = $r[0]; $i -lt $r[0] + $r[1]; $i++) { $bad[$i] = $true } }
            for ($i = 0; $i -lt $Got.Length; $i++) {
                if ($bad[$i]) { if ($Got[$i] -ne 0) { return "byte $i should be a zero" } }
                elseif ($Got[$i] -ne $Expected[$i]) { return "byte $i differs" }
            }
            ''
        }
        $script:Huge = New-HugePst
    }

    It 'default (off): a bad block drops the whole data, the attachment stays with its name and 0 bytes, as before' {
        $p = New-DamagedPst $script:Good $TestDrive 'flip-attach'
        $x = Invoke-RepairKeep $p
        $x.Report.ItemsPartial | Should -Be 1
        $x.Report.AttachmentsKeptDamaged | Should -Be 0
        @($x.Report.Items | Where-Object { $_.Damaged.Count }).Count | Should -Be 0
        $a = Get-ResultAttachment $x.Out
        $a.Long | Should -Be 'f.bin'
        $a.Data.Length | Should -Be 0
    }
    It 'on, a block of an XBLOCK is bad: the data keeps its length, zeros in the bad block, the name says damaged, the ranges are listed' {
        $p = New-DamagedPst $script:Good $TestDrive 'flip-attach'
        $x = Invoke-RepairKeep $p -Keep
        $x.Report.Success | Should -BeTrue
        $x.Report.ItemsPartial | Should -Be 1
        $x.Report.AttachmentsKeptDamaged | Should -Be 1
        $x.Report.DamagedBytes | Should -Be 8176
        $x.Report.After.Errors | Should -Be 0; $x.Report.After.Warnings | Should -Be 0
        $it = $x.Report.Items | Where-Object Status -EQ 'Partial'
        $it.Lost[0] | Should -Match 'attachment 0 data damaged, kept with zeros in place of the unreadable blocks: .* of .* bytes read; zero-filled '
        $it.Damaged.Count | Should -Be 1
        $it.Damaged[0].OriginalName | Should -Be 'f.bin'
        $it.Damaged[0].NewName | Should -Be 'f (damaged).bin'
        $it.Damaged[0].BadRanges.Count | Should -Be 1
        $it.Damaged[0].BadRanges[0].Offset | Should -Be (100 * 8176)
        $it.Damaged[0].BadRanges[0].Length | Should -Be 8176
        $a = Get-ResultAttachment $x.Out
        $a.Long | Should -Be 'f (damaged).bin'
        $a.Display | Should -Be 'f (damaged).bin'
        $rg = ,@((100 * 8176), 8176)
        Test-Ranges $a.Data (Get-Expected 3MB) $rg | Should -Be ''
    }
    It 'on, other items and clean attachments are untouched: the 55 other items are identical to the source' {
        $p = New-DamagedPst $script:Good $TestDrive 'flip-attach' -Seed 3
        $x = Invoke-RepairKeep $p -Keep
        $x.Report.ItemsRecovered | Should -Be 55
        $v = [MailboxTransfer.TestSupport.PstRepairVerifier]::Verify($p, $x.Out, $x.Report)
        $v.Problems | Should -BeNullOrEmpty
        $y = Invoke-RepairKeep $script:Good -Keep
        $y.Report.ItemsRecovered | Should -Be 56
        $y.Report.AttachmentsKeptDamaged | Should -Be 0
    }
    It 'on, an XXBLOCK (9 MB, 1,100 blocks): bad blocks in two XBLOCKs give two ranges, everything else is where it was' {
        $p = Copy-Pst $script:Huge $TestDrive 'huge-leaves'
        Break-Tree $p -Leaf 10, 1030
        $x = Invoke-RepairKeep $p -Keep
        $x.Report.AttachmentsKeptDamaged | Should -Be 1
        $x.Report.DamagedBytes | Should -Be (2 * 8176)
        $d = ($x.Report.Items | Where-Object Status -EQ 'Partial').Damaged[0]
        @($d.BadRanges | ForEach-Object { "$($_.Offset)+$($_.Length)" }) | Should -Be @("$(10 * 8176)+8176", "$(1030 * 8176)+8176")
        $a = Get-ResultAttachment $x.Out
        $a.Long | Should -Be 'f (damaged).bin'; $a.Display | Should -Be 'f (damaged).bin'
        $rg = @(@((10 * 8176), 8176), @((1030 * 8176), 8176))
        Test-Ranges $a.Data (Get-Expected 9MB) $rg | Should -Be ''
        $x.Report.After.Errors | Should -Be 0
    }
    It 'on, an XBLOCK of the XXBLOCK is destroyed: its 1,021 blocks are zeros, the rest is read' {
        $p = Copy-Pst $script:Huge $TestDrive 'huge-child'
        Break-Tree $p -Child 0
        $x = Invoke-RepairKeep $p -Keep
        $d = ($x.Report.Items | Where-Object Status -EQ 'Partial').Damaged[0]
        $d.BadRanges.Count | Should -Be 1
        $d.BadRanges[0].Offset | Should -Be 0
        $d.BadRanges[0].Length | Should -Be (1021 * 8176)
        $a = Get-ResultAttachment $x.Out
        $rg = ,@(0, (1021 * 8176))
        Test-Ranges $a.Data (Get-Expected 9MB) $rg | Should -Be ''
        $x.Report.After.Errors | Should -Be 0
    }
    It 'on, the root of the data tree is destroyed: no length, nothing to keep, the data is dropped as without the option' {
        $p = Copy-Pst $script:Huge $TestDrive 'huge-root'
        Break-Tree $p -Root
        $x = Invoke-RepairKeep $p -Keep
        $x.Report.AttachmentsKeptDamaged | Should -Be 0
        $a = Get-ResultAttachment $x.Out
        $a.Long | Should -Be 'f.bin'
        $a.Data.Length | Should -Be 0
        ($x.Report.Items | Where-Object Status -EQ 'Partial').Lost -join ';' | Should -Match 'CRC mismatch'
    }
    It 'the same through Repair-MbtPst -KeepDamagedAttachments, and the level Salvage keeps the damaged attachment too' {
        Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
        $p = New-DamagedPst $script:Good $TestDrive 'flip-attach' -Seed 4
        $r = Repair-MbtPst -Path $p -Destination (Join-Path $TestDrive 'cmdlet-keep.pst') -Level Salvage -KeepDamagedAttachments
        $r.Success | Should -BeTrue
        $r.AttachmentsKeptDamaged | Should -Be 1
        $r.ItemsPartial | Should -Be 1
        $r.Items.Where({ $_.Status -eq 'Partial' }).Damaged[0].NewName | Should -Be 'f (damaged).bin'
        $r2 = Repair-MbtPst -Path $p -Destination (Join-Path $TestDrive 'cmdlet-nokeep.pst') -Level Salvage
        $r2.AttachmentsKeptDamaged | Should -Be 0
    }
}
