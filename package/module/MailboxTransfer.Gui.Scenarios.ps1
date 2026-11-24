#Requires -Version 7.6
<#
.SYNOPSIS
    Mailbox Transfer console - scenarios for the documentation images and the tests (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    Each scenario puts the window in one state of one page, on synthetic data (contoso.com) and the fake engine: no network, no mailbox. They are
    the states of a run (Running, PC slept, Needs a decision, Stopped), of a result (Complete, Not finished, Needs a decision, Error) and of a PST check (Damaged, Clean).
    The run model of a scenario is built from events, the way a real job feeds it, with the times set back so that elapsed time and ETA read
    like a job of hours.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

function Get-MbtGuiScenarioName {
    @('Start', 'StartImport', 'ExportFromSignedOut', 'ExportFromCode', 'ExportFrom', 'ExportTo', 'ExportPlan', 'RunExport', 'RunImport', 'RunImportSleep', 'RunImportDecision', 'RunImportStopped', 'RunRepair',
        'DoneExportComplete', 'DoneExportPartial', 'DoneExportError', 'ImportFrom', 'ImportTo', 'ImportPlan', 'DoneImportComplete', 'DoneImportPartial', 'DoneImportDecision', 'DoneImportError',
        'RepairFromDamaged', 'RepairFromClean', 'RepairTo', 'RepairPlan', 'DoneRepair', 'Open',
        'ImportFromChoice', 'ImportPlanChoice', 'DoneImportRefused', 'OpenPreview', 'CopyTo', 'CopyPlan', 'DoneCopy', 'ExportPlanShared', 'Setup', 'SetupApplication', 'TaskWhat', 'ExportToMore', 'ExportPlanExpert', 'ImportPlanTicked', 'ImportPlanChanged', 'ExportPlanChecking', 'RunExportPreparing', 'ExportFromPeriod', 'ExportFromPeriodOpen', 'ExportPlanPeriod', 'RunExportPeriod', 'DoneExportPeriod',
        'CopyMailboxFrom', 'CopyMailboxTo', 'CopyMailboxPlan', 'CopyMailboxPlanTicked', 'RunCopyMailboxReading', 'RunCopyMailbox', 'DoneCopyMailbox', 'DoneCopyMailboxPartial')
}

function Get-MbtGuiEdgeScenarioName {
    <# The states with exactly one of something (one folder, one message, one finding, one step...) or with none, and a plan check that failed: the lists of the pages are built from loops that run once or not at all. Walked by the tests, not drawn for the documentation. #>
    @('StartOneJob', 'StartNoJob', 'ExportFromOneLocation', 'ExportFromNoLocation', 'ExportFromOneMessage', 'ExportPlanOneFolder', 'ExportPlanOneMessage', 'ImportToOneLocation', 'ImportPlanOneFolder',
        'RepairFromOneFinding', 'CopyPlanOneFolder', 'RunOneStep', 'DoneExportOneStep', 'DoneImportOneRefused', 'DoneImportOneDecision', 'DoneRepairOneIncomplete', 'DoneCopyOneRefused', 'ExportPlanFailed', 'ImportFromManyFolders', 'ExportFromPeriodInvalid', 'ExportFromPeriodOpenEnd',
        'CopyMailboxFromSeveralPlaces', 'CopyMailboxToApplication', 'CopyMailboxToSame', 'CopyMailboxToArchive', 'CopyMailboxPlanChecking', 'DoneCopyMailboxError', 'DoneCopyMailboxDecision')
}

function New-MbtGuiEvt {
    param([string]$Kind, [datetime]$At, [hashtable]$Fields = @{})
    $o = [ordered]@{ v = 1; t = $At.ToString('o'); e = $Kind }
    foreach ($k in $Fields.Keys) { $o[$k] = $Fields[$k] }
    return [pscustomobject]$o
}

function New-MbtGuiScenarioRun {
    <# A run model for a task in a state, built from events. -Period: an export of a period (the listing says how many items it left out). #>
    param([string]$Task, [string]$State = 'running', [switch]$Period)
    $run = New-MbtGuiRun -Task $Task -Title $(switch ($Task) { 'export' { 'Export ' + [char]0x00B7 + ' adele.vance@contoso.com' } 'import' { 'Import ' + [char]0x00B7 + ' archive-2019.pst ' + [char]0x2192 + ' adele.vance@contoso.com (Archive)' } 'mailcopy' { 'Copy ' + [char]0x00B7 + ' projects@contoso.com (Primary) ' + [char]0x2192 + ' adele.vance@contoso.com (Primary)' } 'copy' { 'Copy ' + [char]0x00B7 + ' archive-2019.pst ' + [char]0x2192 + ' archive-2019-copy.pst' } default { 'Repair ' + [char]0x00B7 + ' contoso-archive-2019.pst ' + [char]0x2192 + ' contoso-archive-2019.repaired.pst' } })
    $hours = switch ($Task) { 'import' { 6.33 } 'repair' { 0.05 } 'mailcopy' { $(if ($State -eq 'reading') { 0.18 } else { 6.33 }) } default { 0.18 } }
    $t0 = [datetime]::UtcNow.AddHours(-$hours)
    $at = { param($min) $t0.AddMinutes($min) }
    $ev = { param($k, $min, $f) $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt $k (& $at $min) $f) }
    switch ($Task) {
        'mailcopy' {
            & $ev 'plan' 0 @{ steps = @(@{ id = 'tree'; title = 'Read the folder tree' }, @{ id = 'list'; title = 'List the items you chose' }, @{ id = 'download'; title = 'Download the items' }, @{ id = 'verify'; title = 'Check the downloads' }, @{ id = 'check'; title = 'Verify the export store' }, @{ id = 'folders'; title = 'Create the folders' }, @{ id = 'send'; title = 'Send the items' }, @{ id = 'reconcile'; title = 'Check the uncertain items' }, @{ id = 'final'; title = 'Final check' }) }
            & $ev 'step' 0.2 @{ id = 'tree'; state = 'done'; msg = '312 folders' }
            & $ev 'step' 1.9 @{ id = 'list'; state = 'done'; msg = '1,208 items ' + [char]0x00B7 + ' 6.2 GiB'; total = 1208 }
            if ($State -eq 'reading') {
                & $ev 'step' 2 @{ id = 'download'; state = 'running'; done = 744; total = 1208; bytes = [long](3.8GB); bytesTotal = [long](6.2GB) }
                & $ev 'object' 3 @{ step = 'download'; key = 'Inbox'; state = 'done'; done = 410; total = 410 }
                & $ev 'object' 5 @{ step = 'download'; key = 'Inbox/Projects/Contoso Rollout'; state = 'running'; done = 260; total = 410 }
                & $ev 'object' 5 @{ step = 'download'; key = 'Archive/Finance/Budgets'; state = 'planned'; done = 0; total = 212 }
                foreach ($m in 5, 6, 7, 8, 9, 10) { & $ev 'rate' $m @{ window = '5m'; gibPerHour = 38 + $m % 3 } }
                & $ev 'step' 11 @{ id = 'download'; state = 'running'; done = 744; total = 1208; bytes = [long](3.8GB); bytesTotal = [long](6.2GB) }
            } else {
                & $ev 'step' 12 @{ id = 'download'; state = 'done'; done = 1208; total = 1208; bytes = [long](6.2GB); bytesTotal = [long](6.2GB) }
                & $ev 'step' 12.5 @{ id = 'verify'; state = 'done'; msg = 'Each item is checked against its fingerprint as it arrives.' }
                & $ev 'step' 13 @{ id = 'check'; state = 'done'; msg = 'The export store is open.' }
                & $ev 'step' 14 @{ id = 'folders'; state = 'done'; msg = '37 folders created' }
                & $ev 'step' 14.2 @{ id = 'send'; state = 'running'; done = 523; total = 1208; bytes = [long](2.7GB); bytesTotal = [long](6.2GB) }
                & $ev 'step' 14.3 @{ id = 'reconcile'; state = 'running'; done = 0; total = 2; msg = '2 waiting for a check' }
                & $ev 'object' 15 @{ step = 'send'; key = 'Inbox'; state = 'done'; done = 410; total = 410 }
                & $ev 'object' 16 @{ step = 'send'; key = 'Projects/Contoso Rollout'; state = 'running'; done = 113; total = 410 }
                & $ev 'object' 16 @{ step = 'send'; key = 'Finance/Budgets'; state = 'planned'; done = 0; total = 212 }
                & $ev 'object' 16 @{ step = 'send'; key = 'Projects/Fabrikam Pilot'; state = 'check'; msg = '2 to check' }
                foreach ($m in 330..380) {
                    & $ev 'step' $m @{ id = 'send'; state = 'running'; done = 500 + ($m - 330) / 2; total = 1208; bytes = [long](2.6GB + ($m - 330) * 2MB); bytesTotal = [long](6.2GB) }
                    if ($m % 5 -eq 0) { & $ev 'rate' $m @{ window = '5m'; gibPerHour = 3.4 + (($m / 5) % 6) * 0.2 } }
                }
                & $ev 'step' 380 @{ id = 'send'; state = 'running'; done = 523; total = 1208; bytes = [long](2.7GB); bytesTotal = [long](6.2GB) }
            }
        }
        'copy' {
            & $ev 'plan' 0 @{ steps = @(@{ id = 'pst'; title = 'Write the PST files' }, @{ id = 'pstcheck'; title = 'Check every PST file' }) }
            & $ev 'step' 0.5 @{ id = 'pst'; state = 'running'; done = 600; total = 1198; bytes = [long](0.5GB); bytesTotal = [long](1.1GB) }
            & $ev 'object' 0.5 @{ step = 'pst'; key = 'Projects/Contoso Rollout'; state = 'running'; done = 410; total = 780 }
        }
        'export' {
            & $ev 'plan' 0 @{ steps = @(@{ id = 'tree'; title = 'Read the folder tree' }, @{ id = 'list'; title = 'List the items you chose' }, @{ id = 'download'; title = 'Download the items' }, @{ id = 'verify'; title = 'Check the downloads' }, @{ id = 'pst'; title = 'Write the PST files' }, @{ id = 'pstcheck'; title = 'Check every PST file' }) }
            & $ev 'step' 0.2 @{ id = 'tree'; state = 'done'; msg = '4,439 folders' }
            & $ev 'step' 1.9 @{ id = 'list'; state = 'done'; msg = $(if ($Period) { '4,612 items outside the period left out; 3 without a date kept' } else { '1,208 items ' + [char]0x00B7 + ' 6.2 GiB' }); total = 1208 }
            & $ev 'step' 2 @{ id = 'download'; state = 'running'; done = 744; total = 1208; bytes = [long](3.8GB); bytesTotal = [long](6.2GB) }
            & $ev 'object' 3 @{ step = 'download'; key = 'Inbox'; state = 'done'; done = 410; total = 410 }
            & $ev 'object' 5 @{ step = 'download'; key = 'Inbox/Projects/Contoso Rollout'; state = 'running'; done = 260; total = 410 }
            & $ev 'object' 5 @{ step = 'download'; key = 'Archive/Finance/Budgets'; state = 'planned'; done = 0; total = 212 }
            & $ev 'object' 5.5 @{ step = 'download'; key = 'Auxiliary archive 1 / Old Mail/2018'; state = 'notready'; msg = 'not ready' }
            foreach ($m in 5, 6, 7, 8, 9, 10) { & $ev 'rate' $m @{ window = '5m'; gibPerHour = 38 + $m % 3 } }
            & $ev 'step' 11 @{ id = 'download'; state = 'running'; done = 744; total = 1208; bytes = [long](3.8GB); bytesTotal = [long](6.2GB) }
        }
        'import' {
            & $ev 'plan' 0 @{ steps = @(@{ id = 'check'; title = 'Check the PST file' }, @{ id = 'folders'; title = 'Create the folders' }, @{ id = 'send'; title = 'Send the items' }, @{ id = 'reconcile'; title = 'Check the uncertain items' }, @{ id = 'final'; title = 'Final check' }) }
            & $ev 'step' 3.3 @{ id = 'check'; state = 'done'; msg = '4 parts, clean' }
            & $ev 'step' 4.1 @{ id = 'folders'; state = 'done'; msg = '214 folders created' }
            & $ev 'step' 4.2 @{ id = 'send'; state = 'running'; done = 5120; total = 9860; bytes = [long](24.6GB); bytesTotal = [long](47.3GB) }
            & $ev 'step' 4.3 @{ id = 'reconcile'; state = 'running'; done = 1; total = 5; msg = '3 waiting for a check' }
            & $ev 'object' 5 @{ step = 'send'; key = 'Inbox'; state = 'done'; done = 1180; total = 1180 }
            & $ev 'object' 6 @{ step = 'send'; key = 'Projects/Contoso Rollout'; state = 'running'; done = 2010; total = 4300 }
            & $ev 'object' 6 @{ step = 'send'; key = 'Old Mail/2019'; state = 'planned'; done = 0; total = 3120 }
            & $ev 'object' 6 @{ step = 'send'; key = 'Finance/Budgets'; state = 'refused'; msg = '2 refused' }
            & $ev 'object' 6 @{ step = 'send'; key = 'Projects/Fabrikam Pilot'; state = 'check'; msg = '3 to check' }
            & $ev 'item' 6 @{ step = 'send'; key = 'Specification-v3 (31 MiB)'; state = 'check'; detail = 'Projects/Fabrikam Pilot' }
            & $ev 'item' 6 @{ step = 'send'; key = 'Site visit photos (12 MiB)'; state = 'check'; detail = 'Projects/Fabrikam Pilot' }
            & $ev 'wait' 300 @{ reason = 'throttle'; seconds = 38; count = 7; waited = 252 }
            foreach ($m in 312..372) {
                & $ev 'step' $m @{ id = 'send'; state = 'running'; done = 5000 + $m - 312; total = 9860; bytes = [long](24.5GB + ($m - 312) * 1MB); bytesTotal = [long](47.3GB) }
                if ($m % 5 -eq 0) { & $ev 'rate' $m @{ window = '5m'; gibPerHour = 3.2 + (($m / 5) % 6) * 0.3 } }
            }
            & $ev 'step' 372 @{ id = 'send'; state = 'running'; done = 5120; total = 9860; bytes = [long](24.6GB); bytesTotal = [long](47.3GB) }
            $run.Waits.Active = $null
        }
        default {
            & $ev 'plan' 0 @{ steps = @(@{ id = 'before'; title = 'Check the file (before)' }, @{ id = 'scan'; title = 'Scan and read what can be read' }, @{ id = 'write'; title = 'Write the new file' }, @{ id = 'after'; title = 'Check the new file (after)' }) }
            & $ev 'step' 1 @{ id = 'before'; state = 'done'; msg = '6 errors, 5 warnings' }
            & $ev 'step' 1.1 @{ id = 'scan'; state = 'done'; msg = '1,203 messages found'; total = 1203; done = 1203 }
            & $ev 'step' 2 @{ id = 'write'; state = 'running'; done = 740; total = 1203; bytes = [long](2.6GB); bytesTotal = [long](4.2GB) }
            & $ev 'object' 2 @{ step = 'write'; key = 'Inbox'; state = 'done'; msg = '412 recovered' }
            & $ev 'object' 2 @{ step = 'write'; key = 'Projects'; state = 'running'; done = 210; total = 480 }
            & $ev 'object' 2 @{ step = 'write'; key = 'Recovered items/Lost folder 0x8021'; state = 'check'; msg = '7 messages' }
            & $ev 'object' 2 @{ step = 'write'; key = 'Sent Items'; state = 'planned'; done = 0; total = 306 }
        }
    }
    switch ($State) {
        'sleep' { & $ev 'slept' 372 @{ from = (& $at 330).ToString('o'); to = (& $at 366).ToString('o'); seconds = 2160 }; $run.Sleep.From = [datetime]::UtcNow.Date.AddHours(2).AddMinutes(14); $run.Sleep.To = [datetime]::UtcNow.Date.AddHours(6).AddMinutes(48); $run.EpochStart = [datetime]::UtcNow.AddMinutes(-3) }
        'decision' { & $ev 'decision' 372 @{ key = '7114'; title = 'Signed contract (final)'; msg = '"Signed contract (final)" was sent, but Microsoft did not answer, and **two items in the folder could be it** (both 4 MiB, both created after the attempt). We do not guess.'; candidates = @('A', 'B') } }
        'stopped' { $run.StopRequested = $true; & $ev 'result' 372 @{ exit = 2; complete = $false; next = 'continue'; reason = 'stopped'; counts = @{ done = 5120; pending = 4738 } } }
    }
    return $run
}

function New-MbtGuiScenarioPreparingRun {
    <# An export that has not counted an item yet: it reads the folders of the mailbox, and says how far it is. #>
    $run = New-MbtGuiRun -Task export -Title ('Export ' + [char]0x00B7 + ' adele.vance@contoso.com')
    $t0 = [datetime]::UtcNow.AddSeconds(-34)
    $ev = { param($k, $sec, $f) $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt $k $t0.AddSeconds($sec) $f) }
    & $ev 'plan' 0 @{ steps = @(@{ id = 'tree'; title = 'Read the folder tree' }, @{ id = 'list'; title = 'List the items you chose' }, @{ id = 'download'; title = 'Download the items' }, @{ id = 'verify'; title = 'Check the downloads' }, @{ id = 'pst'; title = 'Write the PST files' }, @{ id = 'pstcheck'; title = 'Check every PST file' }) }
    & $ev 'step' 0.5 @{ id = 'tree'; state = 'running'; msg = 'Signing in to Microsoft' }
    & $ev 'step' 3 @{ id = 'tree'; state = 'running'; msg = 'Reading your selection' }
    & $ev 'step' 33 @{ id = 'tree'; state = 'running'; msg = '1,240 folders read' }
    return $run
}

function New-MbtGuiScenarioOneRun {
    <#
        A run with exactly one step, one folder and one warning: every list of the Run page holds one row. -Outcome: running, or how it ended (complete, refused, decision). -Item: one item of that
        state (refused, partial, lost) is listed as well.
    #>
    param([Parameter(Mandatory)][string]$Task, [ValidateSet('running', 'complete', 'refused', 'decision')][string]$Outcome = 'running', [string]$Item)
    $run = New-MbtGuiRun -Task $Task -Title "One step $([char]0x00B7) $Task"
    $t0 = [datetime]::UtcNow.AddMinutes(-20)
    $ev = { param($k, $min, $f) $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt $k $t0.AddMinutes($min) $f) }
    $id = switch ($Task) { 'import' { 'send' } 'repair' { 'write' } 'copy' { 'pst' } default { 'download' } }
    & $ev 'plan' 0 @{ steps = @(@{ id = $id; title = 'The only step' }) }
    & $ev 'step' 1 @{ id = $id; state = 'running'; done = 3; total = 10; bytes = [long](3MB); bytesTotal = [long](10MB) }
    & $ev 'object' 2 @{ step = $id; key = 'Inbox'; state = 'running'; done = 3; total = 10 }
    & $ev 'warn' 3 @{ code = 'throttled'; step = $id; key = 'Inbox'; msg = 'Microsoft asked to slow down.' }
    if ($Item) { & $ev 'item' 4 @{ step = $id; key = 'Signed contract (final)'; state = $Item; msg = $(if ($Item -eq 'refused') { 'refused by the service: ErrorItemTooLarge' } else { 'one attachment dropped' }); detail = 'Inbox' } }
    if ($Outcome -eq 'decision') { & $ev 'decision' 5 @{ key = '7114'; title = 'Signed contract (final)'; msg = 'Two items in the folder could be this one. We do not guess.'; candidates = @('A') } }
    if ($Outcome -ne 'running') {
        $exit = switch ($Outcome) { 'complete' { 0 } 'decision' { 3 } default { 1 } }
        $counts = switch ($Task) { 'import' { @{ done = 9; failed = 0; refused = 1; unknown = 1; pending = 0; foldersCreated = 1 } } 'repair' { @{ recovered = 9; partial = 1; lost = 0 } } 'copy' { @{ itemsWritten = 9; failed = 0; refused = 1; parts = 1; bytes = [long](1MB) } } default { @{ done = 1; failed = 0; bytes = [long](10MB); parts = 1 } } }
        & $ev 'result' 10 @{ exit = $exit; complete = ($exit -eq 0 -or $Outcome -eq 'refused'); counts = $counts; next = 'done'; reason = $(if ($Outcome -eq 'refused') { 'refused-items' } else { '' }) }
        foreach ($step in $run.Steps) { $step.State = 'done' }
    }
    return $run
}

function New-MbtGuiScenarioDone {
    <# A run model at its end: Task, Outcome complete / partial / decision / error. -Period: an export of a period (the result counts the items kept, outside the period and without a date). #>
    param([string]$Task, [string]$Outcome, [switch]$Period)
    $run = New-MbtGuiScenarioRun -Task $Task -Period:($Period)
    $exit = switch ($Outcome) { 'complete' { 0 } 'partial' { 2 } 'decision' { 3 } default { 1 } }   # refused: finished, some items refused for good (exit 1, reason refused-items)
    $counts = switch ("$Task/$Outcome") {
        'export/complete' { @{ done = 1208; failed = 0; bytes = [long](6.2GB); parts = 3 } }
        'export/partial' { @{ done = 1004; pending = 204; failed = 0 } }
        'export/error' { @{ done = 744; failed = 0 } }
        'mailcopy/complete' { @{ exported = 1208; done = 1208; failed = 0; unknown = 0; pending = 0; total = 1208; foldersCreated = 37; bytes = [long](6.2GB) } }
        'mailcopy/partial' { @{ exported = 1208; done = 523; failed = 0; unknown = 0; pending = 685; total = 1208; foldersCreated = 37; bytes = [long](6.2GB) } }
        'mailcopy/decision' { @{ exported = 1208; done = 1207; failed = 0; unknown = 1; pending = 0; total = 1208; foldersCreated = 37; bytes = [long](6.2GB) } }
        'mailcopy/error' { @{ exported = 1208; done = 523; failed = 0; pending = 685; total = 1208 } }
        'import/complete' { @{ done = 9860; failed = 0; unknown = 0; pending = 0; total = 9860; foldersCreated = 214 } }
        'import/partial' { @{ done = 5120; failed = 2; unknown = 0; pending = 4738; total = 9860; foldersCreated = 214 } }
        'import/decision' { @{ done = 9859; failed = 0; unknown = 1; pending = 0; total = 9860; foldersCreated = 214 } }
        'import/error' { @{ done = 5120; failed = 2; unknown = 0; pending = 0; total = 9860 } }
        'import/refused' { @{ done = 9857; failed = 0; refused = 3; unknown = 0; pending = 0; total = 9860; foldersCreated = 214; report = 'D:\MailboxTransfer\jobs\20261109-1402-import-7c1e\work\pstimport-not-imported.csv' } }
        'copy/complete' { @{ itemsWritten = 1198; failed = 0; parts = 1; bytes = [long](1.1GB) } }
        'repair/complete' { @{ recovered = 1198; partial = 4; lost = 1; levelUsed = 'Rebuild'; destination = 'D:\PST\contoso-archive-2019.repaired.pst'; before = @{ errors = 6; warnings = 5 }; after = @{ errors = 0; warnings = 0 }; sizeBytes = [long](3.9GB) } }
        default { @{} }
    }
    if ($Period -and $Task -eq 'export') { $counts = $counts + @{ kept = [long]1208; outsidePeriod = [long]4612; withoutDate = [long]3 } }
    $at = [datetime]::UtcNow
    if ($Outcome -eq 'error') {
        $err = switch ($Task) { 'mailcopy' { @{ code = 'throttled-out'; msg = 'Microsoft refused the requests of the second part again and again. What was read is kept in the store and what was added stays in the mailbox. Press Try again in a while.'; tech = 'HTTP 503 ErrorServerBusy on importItem (7 times in a row)' } } 'export' { @{ code = 'disk-full'; msg = 'The disk D: is full. Free some space (about 5 GiB) and press Try again. The 744 items already downloaded are kept.'; tech = 'IOException: There is not enough space on the disk. (store pack-000003.bin)' } } 'import' { @{ code = 'signin-required'; msg = 'Your sign-in has expired. Sign in again to continue. Nothing is lost.'; tech = 'AADSTS50173: The provided grant has expired due to it being revoked.' } } default { @{ code = 'destination-exists'; msg = 'The destination file already exists. Choose another name. The repair never overwrites a file.'; tech = '' } } }
        $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt 'error' $at $err)
    }
    if ($Task -eq 'repair' -and $Outcome -eq 'complete') {
        foreach ($x in @(@('Signed contract (final)', 'partial', 'one attachment: 3 blocks unreadable (kept as Specification-v3 (damaged).pdf when "keep damaged attachments" is on)', 'Projects/Contoso Rollout'), @('Quarterly report draft', 'partial', 'the recipient list could not be read', 'Inbox'),
                @('Site visit photos', 'partial', 'one attachment dropped', 'Inbox'), @('Budget forecast Q4', 'partial', 'one attachment dropped', 'Finance/Budgets'), @('(no subject)', 'lost', 'its own data blocks are damaged; nothing could be read', 'Recovered items'))) {
            $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt 'item' $at @{ step = 'write'; key = $x[0]; state = $x[1]; msg = $x[2]; detail = $x[3] })
        }
    }
    if ($Task -eq 'import' -and $Outcome -eq 'refused') {
        foreach ($x in @(@('Signed contract (final)', 'refused by the service: ErrorItemTooLarge The item is larger than the limit of the mailbox.', 'Projects/Contoso Rollout'), @('Site visit photos', 'cannot be read from the PST: the data block 0x4F2A is damaged', 'Inbox'), @('(no subject)', 'gave up after 3 attempt(s), nothing was created. Last error: HTTP 404 ErrorItemNotFound', 'Old Mail/2019'))) {
            $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt 'item' $at @{ step = 'send'; key = $x[0]; state = 'refused'; msg = $x[1]; detail = $x[2] })
        }
    }
    if ($Task -in 'import', 'mailcopy' -and $Outcome -eq 'decision') {
        $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt 'decision' $at @{ key = '7114'; title = 'Signed contract (final)'; msg = 'Two items in the folder could be this one. We do not guess.'; candidates = @('A', 'B') })
    }
    $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt 'result' $at @{ exit = $exit; complete = ($exit -eq 0 -or $Outcome -eq 'refused'); counts = $counts; next = $(if ($Outcome -eq 'refused') { 'done' } else { switch ($exit) { 0 { 'done' } 3 { 'decide' } 1 { 'retry' } default { 'continue' } } }); reason = $(if ($Task -eq 'import' -and $Outcome -eq 'error') { 'signin-required' } elseif ($Outcome -eq 'refused') { 'refused-items' } else { '' }) })
    if ($Outcome -eq 'complete') { foreach ($s in $run.Steps) { $s.State = 'done' } }
    if ($Task -eq 'export' -and $Outcome -eq 'partial') { $run.Objects[3].State = 'notready' }
    return $run
}

function Show-MbtGuiRunScenario {
    param([string]$Task, $Run)
    $g = $script:Gui; $s = $g.S
    $s.Task = $Task; $s.Sel = $Task; $s.Signed = $true; $s.Mailbox = 'adele.vance@contoso.com'; $s.Account = $s.Mailbox
    if ($Task -eq 'mailcopy') { Initialize-MbtGuiScenarioCopy }
    if ($Task -eq 'copy') { $s.OpSrc = 'pst'; $s.OpPath = 'D:\PST\archive-2019.pst'; $s.OutFolder = 'D:\PST'; $s.OutName = 'archive-2019-copy.pst' }
    $g.Run = $Run; $g.Job = $null
    $idx = $script:Flows[$Task].IndexOf($(if ($Run.Result -and $Run.State -ne 'stopped') { 'done' } else { 'run' }))
    $s.Step = $idx; $s.Outcome = $Run.State
    Show-MbtGuiPage
}

function Initialize-MbtGuiScenarioExport {
    <#
        Signed in, the three locations, some folders and messages ticked. -Kinds: only the locations of these kinds ('None': no location). -OneFolder: one folder ticked alone, nothing else.
        -OneMessage: the one message of a folder, ticked alone.
    #>
    param([switch]$Ticks, [string[]]$Kinds, [switch]$OneFolder, [switch]$OneMessage, [switch]$OnePlace)
    $g = $script:Gui; $s = $g.S
    $s.Task = 'export'; $s.Sel = 'export'; $s.Mailbox = 'adele.vance@contoso.com'; $s.Signed = $true; $s.Account = $s.Mailbox
    $s.Token = [pscustomobject]@{ Access = 'Delegated (you, with your rights)'; CanReadMail = $true; Missing = @() }
    & $g.Engine.Locations $s.Mailbox '' { param($l) $s.Locations = @($l | Where-Object { -not $Kinds -or $_.Kind -in $Kinds }); foreach ($x in $s.Locations) { $s.LocationChecked[$x.Kind] = $true } } { }
    $browser = Get-MbtGuiBrowser 'export'
    & $g.Engine.Roots $s.Locations { param($r) Set-MbtGuiBrowserRoots -Browser $browser -Roots @($r) } { }
    $inbox = $null
    foreach ($r in $browser.Roots) { foreach ($n in $r.Descendants()) { if ($n.Id -eq 'inbox') { $inbox = $n } } }
    if ($inbox) {
        $inbox.IsExpanded = $true; $inbox.IsSelected = $true
        Open-MbtGuiFolder $inbox
    }
    if ($OneFolder) {
        foreach ($r in $browser.Roots) { foreach ($n in $r.Descendants()) { if ($n.Id -eq 'p-c') { $n.Check = $true } } }
    }
    if ($OneMessage) {
        # a folder that holds one message: the list has one row, and it is the one that is ticked
        foreach ($r in $browser.Roots) { foreach ($n in $r.Descendants()) { if ($n.Id -eq 'drafts') { $n.Count = 1; $n.IsSelected = $true; Open-MbtGuiFolder $n } } }
        $browser.Rows[0].IsChecked = $true
    }
    if ($Ticks) {
        $tickIds = if ($OnePlace) { 'p-c', 'p-f' } else { 'p-c', 'fin-b', 'x-18' }
        foreach ($r in $browser.Roots) { foreach ($n in $r.Descendants()) { if ($n.Id -in $tickIds) { $n.Check = $true } } }
        $rows = @($browser.Rows | Select-Object -First 5)
        foreach ($row in $rows) { $row.IsChecked = $true }
        $browser.Rows[0].IsChecked = $true
    }
    $g.ExportSelection = Get-MbtGuiSelection -Roots @($browser.Roots)
    $s.OutFolder = 'D:\Exports\adele'; $s.OutName = 'adele.pst'
}

function Initialize-MbtGuiScenarioCopy {
    <# A copy to another mailbox: a shared mailbox is read, the signed-in account receives, the folder is named. -Ticks: the ticks, in one place. -Application: the application signs in, the target is typed. -Archive: the target is the archive. #>
    param([switch]$Ticks, [switch]$Application, [switch]$Archive)
    $g = $script:Gui; $s = $g.S
    if ($Ticks) { Initialize-MbtGuiScenarioExport -Ticks -OnePlace }
    $s.Task = 'mailcopy'; $s.Sel = 'mailcopy'; $s.Signed = $true
    $s.Mailbox = 'adele.vance@contoso.com'; $s.Account = $s.Mailbox; $s.Shared = 'projects@contoso.com'; $s.McFromKind = 'Primary'
    $s.TgtLoc = $(if ($Archive) { 'Archive' } else { 'Primary' }); $s.TgtMode = 'new'; $s.TgtName = 'Copy of projects 2026-11-12'
    if ($Application) { $s.AuthMode = 'app'; $s.AppThumb = '5D41402ABC4B2A76B9719D911017C59228F3D1AE'; $s.Mailbox = 'projects@contoso.com'; $s.Account = $s.Mailbox; $s.Shared = ''; $s.McTarget = 'adele.vance@contoso.com' }
}

function Set-MbtGuiScenario {
    <# Puts the window in a state. Names: Get-MbtGuiScenarioName. #>
    param([Parameter(Mandatory)][string]$Name)
    $g = $script:Gui; $s = $g.S
    $s.WorkFolder = 'D:\MailboxTransfer'
    $script:FakeJobCount = $null
    $script:FakePstExtraFolders = 0
    switch ($Name) {
        'Start' { $s.Step = 1; Show-MbtGuiPage }
        'StartImport' { $s.Step = 1; $s.Sel = 'import'; Show-MbtGuiPage }
        'ExportFromSignedOut' { $s.Task = 'export'; $s.Sel = 'export'; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); $s.Mailbox = 'adele.vance@contoso.com'; Show-MbtGuiPage }
        'ExportFromCode' { $s.Task = 'export'; $s.Sel = 'export'; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); $s.Mailbox = 'adele.vance@contoso.com'; $s.SigningIn = $true; $s.CodeShown = $true; $s.Code = @{ Code = 'ABCD1EFGH'; Url = 'https://microsoft.com/devicelogin'; Until = (Get-Date).Date.AddHours(14).AddMinutes(32) }; Show-MbtGuiPage }
        'ExportFrom' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportTo' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exTo'); Show-MbtGuiPage }
        'ExportPlan' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage }
        'RunExport' { Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioRun -Task export) }
        'RunExportPreparing' { Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioPreparingRun) }
        'ExportPlanChecking' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage; $g.PlanRun = New-MbtGuiScenarioPreparingRun; $g.PlanStarted = [datetime]::UtcNow.AddSeconds(-34); $g.PlanShown = ''; Set-MbtGuiPlanState checking }
        'RunImport' { Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioRun -Task import) }
        'RunImportSleep' { Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioRun -Task import -State sleep) }
        'RunImportDecision' { Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioRun -Task import -State decision) }
        'RunImportStopped' { Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioRun -Task import -State stopped) }
        'RunRepair' { Show-MbtGuiRunScenario 'repair' (New-MbtGuiScenarioRun -Task repair) }
        'DoneExportComplete' { Initialize-MbtGuiScenarioExport -Ticks; $s.OutName = 'adele.pst'; $s.OutFolder = 'D:\Exports\adele'; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioDone export complete) }
        'ExportFromPeriod' { Initialize-MbtGuiScenarioExport -Ticks; $s.Period = 'between'; $s.ArrFrom = '2026-01-01'; $s.ArrTo = '2026-03-31'; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportFromPeriodOpen' { Initialize-MbtGuiScenarioExport -Ticks; $s.Period = 'between'; $s.ArrFrom = '2026-01-01'; $s.ArrTo = '2026-03-31'; $s.PeriodOpen = $true; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportFromPeriodOpenEnd' { Initialize-MbtGuiScenarioExport -Ticks; $s.Period = 'between'; $s.ArrFrom = '2025-07-01'; $s.ArrTo = ''; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportFromPeriodInvalid' { Initialize-MbtGuiScenarioExport -Ticks; $s.Period = 'between'; $s.ArrFrom = '2026-04-01'; $s.ArrTo = '2026-03-01'; $s.PeriodOpen = $true; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportPlanPeriod' { Initialize-MbtGuiScenarioExport -Ticks; $s.Period = 'between'; $s.ArrFrom = '2026-01-01'; $s.ArrTo = '2026-03-31'; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage }
        'RunExportPeriod' { Initialize-MbtGuiScenarioExport -Ticks; $s.Period = 'between'; $s.ArrFrom = '2026-01-01'; $s.ArrTo = '2026-03-31'; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioRun -Task export -Period) }
        'DoneExportPeriod' { Initialize-MbtGuiScenarioExport -Ticks; $s.Period = 'between'; $s.ArrFrom = '2026-01-01'; $s.ArrTo = '2026-03-31'; $s.OutName = 'adele.pst'; $s.OutFolder = 'D:\Exports\adele'; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioDone export complete -Period) }
        'DoneExportPartial' { Initialize-MbtGuiScenarioExport -Ticks; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioDone export partial) }
        'DoneExportError' { Initialize-MbtGuiScenarioExport -Ticks; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioDone export error) }
        'ImportFrom' { $s.Task = 'import'; $s.Sel = 'import'; $s.Step = (Get-MbtGuiScenarioStep 'import' 'imFrom'); $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; Open-MbtGuiSource -Use import }
        'ImportFromManyFolders' { $script:FakePstExtraFolders = 400; $s.Task = 'import'; $s.Sel = 'import'; $s.Step = (Get-MbtGuiScenarioStep 'import' 'imFrom'); $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; Open-MbtGuiSource -Use import }
        'ImportTo' { Set-MbtGuiScenario 'ImportFrom'; Initialize-MbtGuiScenarioSignIn; $s.Step = (Get-MbtGuiScenarioStep 'import' 'imTo'); $s.TgtLoc = 'Archive'; $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiPage }
        'ImportPlan' { Set-MbtGuiScenario 'ImportTo'; $s.Step = (Get-MbtGuiScenarioStep 'import' 'imPlan'); Show-MbtGuiPage }
        'DoneImportComplete' { $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; $s.TgtLoc = 'Archive'; $s.TgtMode = 'new'; $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioDone import complete) }
        'DoneImportPartial' { $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; $s.TgtLoc = 'Archive'; $s.TgtMode = 'new'; $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioDone import partial) }
        'DoneImportDecision' { $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; $s.TgtLoc = 'Archive'; $s.TgtMode = 'new'; $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioDone import decision) }
        'DoneImportError' { $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; $s.TgtLoc = 'Archive'; $s.TgtMode = 'new'; $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioDone import error) }
        'RepairFromDamaged' { $script:FakePstState = 'damaged'; $s.Task = 'repair'; $s.Sel = 'repair'; $s.Step = (Get-MbtGuiScenarioStep 'repair' 'rpFrom'); $s.RpPath = 'D:\PST\contoso-archive-2019.pst'; Start-MbtGuiPstCheck }
        'RepairFromClean' { $script:FakePstState = 'clean'; $s.Task = 'repair'; $s.Sel = 'repair'; $s.Step = (Get-MbtGuiScenarioStep 'repair' 'rpFrom'); $s.RpPath = 'D:\PST\contoso-archive-2019.pst'; Start-MbtGuiPstCheck; $script:FakePstState = 'damaged' }
        'RepairTo' { Set-MbtGuiScenario 'RepairFromDamaged'; $s.Step = (Get-MbtGuiScenarioStep 'repair' 'rpTo'); $s.RpDest = 'D:\PST\contoso-archive-2019.repaired.pst'; Show-MbtGuiPage }
        'RepairPlan' { Set-MbtGuiScenario 'RepairFromDamaged'; $s.Step = (Get-MbtGuiScenarioStep 'repair' 'rpPlan'); $s.RpDest = 'D:\PST\contoso-archive-2019.repaired.pst'; Show-MbtGuiPage }
        'DoneRepair' { $s.RpPath = 'D:\PST\contoso-archive-2019.pst'; $s.RpDest = 'D:\PST\contoso-archive-2019.repaired.pst'; Show-MbtGuiRunScenario 'repair' (New-MbtGuiScenarioDone repair complete) }
        'Open' { $s.Task = 'open'; $s.Sel = 'open'; $s.Step = (Get-MbtGuiScenarioStep 'open' 'opFrom'); $s.OpSrc = 'pst'; $s.OpPath = 'D:\PST\archive-2019.pst'; Open-MbtGuiSource -Use open; Open-MbtGuiFirstFolder 'open' }
        'ImportFromChoice' { Set-MbtGuiScenario 'ImportFrom'; Set-MbtGuiScenarioTicks 'import' 'o-p1', 'o-bud'; Show-MbtGuiPage }
        'ImportPlanChoice' { Set-MbtGuiScenario 'ImportTo'; Set-MbtGuiScenarioTicks 'import' 'o-p1', 'o-bud'; $s.Step = (Get-MbtGuiScenarioStep 'import' 'imPlan'); Show-MbtGuiPage }
        'DoneImportRefused' { $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; $s.TgtLoc = 'Archive'; $s.TgtMode = 'new'; $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioDone import refused) }
        'OpenPreview' {
            Set-MbtGuiScenario 'Open'
            $b = Get-MbtGuiBrowser 'open'
            Set-MbtGuiScenarioTicks 'open' 'o-p1', 'o-bud'
            Show-MbtGuiPage
            # let the tree settle (it opens the folder again when its items are realised), then read a message as a person would
            $script:Gui.Window.Content.UpdateLayout(); [Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action] { }, [Windows.Threading.DispatcherPriority]::ContextIdle)
            $row = @($b.Rows | Where-Object HasAttachment | Select-Object -First 1)[0]; if (-not $row -and $b.Rows.Count) { $row = $b.Rows[0] }
            if ($row) { $row.IsChecked = $true; Show-MbtGuiMessage $b $row }
        }
        'CopyTo' { Set-MbtGuiScenario 'OpenPreview'; $s.Task = 'copy'; $s.Sel = 'copy'; $s.Step = (Get-MbtGuiScenarioStep 'copy' 'cpTo'); Initialize-MbtGuiCopy; $s.OutFolder = 'D:\PST'; $s.OutName = 'archive-2019-copy.pst'; Show-MbtGuiPage }
        'CopyPlan' { Set-MbtGuiScenario 'CopyTo'; $s.Step = (Get-MbtGuiScenarioStep 'copy' 'cpPlan'); Show-MbtGuiPage }
        'DoneCopy' { $s.OpPath = 'D:\PST\archive-2019.pst'; $s.OutFolder = 'D:\PST'; $s.OutName = 'archive-2019-copy.pst'; Show-MbtGuiRunScenario 'copy' (New-MbtGuiScenarioDone copy complete) }
        'CopyMailboxFrom' { Initialize-MbtGuiScenarioCopy -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'exFrom'); Show-MbtGuiPage }
        'CopyMailboxFromSeveralPlaces' { Initialize-MbtGuiScenarioExport -Ticks; $s.Task = 'mailcopy'; $s.Sel = 'mailcopy'; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'exFrom'); Show-MbtGuiPage }
        'CopyMailboxTo' { Initialize-MbtGuiScenarioCopy -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'mcTo'); Show-MbtGuiPage }
        'CopyMailboxToApplication' { Initialize-MbtGuiScenarioCopy -Ticks -Application; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'mcTo'); Show-MbtGuiPage }
        'CopyMailboxToSame' { Initialize-MbtGuiScenarioCopy -Ticks; $s.Shared = ''; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'mcTo'); Show-MbtGuiPage }
        'CopyMailboxToArchive' { Initialize-MbtGuiScenarioCopy -Ticks -Archive; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'mcTo'); Show-MbtGuiPage }
        'CopyMailboxPlan' { Initialize-MbtGuiScenarioCopy -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'mcPlan'); Show-MbtGuiPage }
        'CopyMailboxPlanTicked' { Initialize-MbtGuiScenarioCopy -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'mcPlan'); Show-MbtGuiPage; $s.Confirmed = $true; $s.ConfirmKey = (Get-MbtGuiPlanKey); Show-MbtGuiPage }
        'CopyMailboxPlanChecking' { Initialize-MbtGuiScenarioCopy -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'mailcopy' 'mcPlan'); Show-MbtGuiPage; $g.PlanRun = New-MbtGuiScenarioPreparingRun; $g.PlanStarted = [datetime]::UtcNow.AddSeconds(-34); $g.PlanShown = ''; Set-MbtGuiPlanState checking; Show-MbtGuiPage }
        'RunCopyMailboxReading' { Initialize-MbtGuiScenarioCopy -Ticks; Show-MbtGuiRunScenario 'mailcopy' (New-MbtGuiScenarioRun -Task mailcopy -State reading) }
        'RunCopyMailbox' { Initialize-MbtGuiScenarioCopy -Ticks; Show-MbtGuiRunScenario 'mailcopy' (New-MbtGuiScenarioRun -Task mailcopy) }
        'DoneCopyMailbox' { Initialize-MbtGuiScenarioCopy -Ticks; Show-MbtGuiRunScenario 'mailcopy' (New-MbtGuiScenarioDone mailcopy complete) }
        'DoneCopyMailboxPartial' { Initialize-MbtGuiScenarioCopy -Ticks; Show-MbtGuiRunScenario 'mailcopy' (New-MbtGuiScenarioDone mailcopy partial) }
        'DoneCopyMailboxError' { Initialize-MbtGuiScenarioCopy -Ticks; Show-MbtGuiRunScenario 'mailcopy' (New-MbtGuiScenarioDone mailcopy error) }
        'DoneCopyMailboxDecision' { Initialize-MbtGuiScenarioCopy -Ticks; Show-MbtGuiRunScenario 'mailcopy' (New-MbtGuiScenarioDone mailcopy decision) }
        'ExportPlanShared' { Initialize-MbtGuiScenarioExport -Ticks; $s.Shared = 'projects@contoso.com'; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage }
        'Setup' { $s.Step = 0; Show-MbtGuiPage }
        'SetupApplication' { $s.Step = 0; $s.AuthMode = 'app'; $s.AppThumb = '5D41402ABC4B2A76B9719D911017C59228F3D1AE'; Show-MbtGuiPage }
        'TaskWhat' { $s.Step = 1; $s.Sel = 'export'; Show-MbtGuiPage }
        'ExportToMore' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exTo'); Show-MbtGuiPage; Expand-MbtGuiScenarioFold 'More options' }
        'ExportPlanExpert' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage; Expand-MbtGuiScenarioFold 'Expert mode'; $script:Gui.C.PageScroll.ScrollToEnd() }
        'ImportPlanTicked' { Set-MbtGuiScenario 'ImportPlan'; $s.Confirmed = $true; $s.ConfirmKey = (Get-MbtGuiPlanKey); Show-MbtGuiPage }
        'ImportPlanChanged' { Set-MbtGuiScenario 'ImportPlanTicked'; $s.Mailbox = 'Adele.Vance@Contoso.com'; Show-MbtGuiPage; $s.TgtName = 'Imported archive-2019 2026-11-13'; Show-MbtGuiPage }
        'StartOneJob' { $script:FakeJobCount = 1; $s.Step = 1; Show-MbtGuiPage }
        'StartNoJob' { $script:FakeJobCount = 0; $s.Step = 1; Show-MbtGuiPage }
        'ExportFromOneLocation' { Initialize-MbtGuiScenarioExport -Kinds Primary -OneFolder; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportFromNoLocation' { Initialize-MbtGuiScenarioExport -Kinds None; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportFromOneMessage' { Initialize-MbtGuiScenarioExport -OneMessage; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportPlanFailed' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage; Set-MbtGuiPlanState error '**The plan cannot be used.** Microsoft did not answer correctly (a temporary problem). Press Check again.' }
        'ExportPlanOneFolder' { Initialize-MbtGuiScenarioExport -OneFolder; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage }
        'ExportPlanOneMessage' { Initialize-MbtGuiScenarioExport -OneMessage; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage }
        'ImportToOneLocation' { Set-MbtGuiScenario 'ImportFrom'; Initialize-MbtGuiScenarioSignIn -Kinds Primary; $s.Step = (Get-MbtGuiScenarioStep 'import' 'imTo'); $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiPage }
        'ImportPlanOneFolder' { Set-MbtGuiScenario 'ImportTo'; Set-MbtGuiScenarioTicks 'import' 'o-p1'; $s.Step = (Get-MbtGuiScenarioStep 'import' 'imPlan'); Show-MbtGuiPage }
        'RepairFromOneFinding' { $script:FakePstState = 'one'; $s.Task = 'repair'; $s.Sel = 'repair'; $s.Step = (Get-MbtGuiScenarioStep 'repair' 'rpFrom'); $s.RpPath = 'D:\PST\contoso-archive-2019.pst'; Start-MbtGuiPstCheck; $script:FakePstState = 'damaged' }
        'CopyPlanOneFolder' { Set-MbtGuiScenario 'Open'; Set-MbtGuiScenarioTicks 'open' 'o-p1'; $s.Task = 'copy'; $s.Sel = 'copy'; $s.Step = (Get-MbtGuiScenarioStep 'copy' 'cpPlan'); Initialize-MbtGuiCopy; $s.OutFolder = 'D:\PST'; $s.OutName = 'archive-2019-copy.pst'; Show-MbtGuiPage }
        'RunOneStep' { Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioOneRun -Task export) }
        'DoneExportOneStep' { Initialize-MbtGuiScenarioExport -OneFolder; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioOneRun -Task export -Outcome complete) }
        'DoneImportOneRefused' { $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; $s.TgtLoc = 'Archive'; $s.TgtMode = 'new'; $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioOneRun -Task import -Outcome refused -Item refused) }
        'DoneImportOneDecision' { $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; $s.TgtLoc = 'Archive'; $s.TgtMode = 'new'; $s.TgtName = 'Imported archive-2019 2026-11-12'; Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioOneRun -Task import -Outcome decision) }
        'DoneRepairOneIncomplete' { $s.RpPath = 'D:\PST\contoso-archive-2019.pst'; $s.RpDest = 'D:\PST\contoso-archive-2019.repaired.pst'; Show-MbtGuiRunScenario 'repair' (New-MbtGuiScenarioOneRun -Task repair -Outcome complete -Item partial) }
        'DoneCopyOneRefused' { $s.OpPath = 'D:\PST\archive-2019.pst'; $s.OutFolder = 'D:\PST'; $s.OutName = 'archive-2019-copy.pst'; Show-MbtGuiRunScenario 'copy' (New-MbtGuiScenarioOneRun -Task copy -Outcome refused -Item refused) }
        default { throw "Unknown scenario '$Name'. Known: $((Get-MbtGuiScenarioName) -join ', '), $((Get-MbtGuiEdgeScenarioName) -join ', ')." }
    }
    if ($Name -in 'ImportFrom') { Open-MbtGuiFirstFolder 'import' }
    if ($Name -in 'ImportTo', 'ImportPlan') { }
}

function Get-MbtGuiScenarioStep {
    <# The index of a page in the flow of a task. #>
    param([string]$Task, [string]$Screen)
    return $script:Flows[$Task].IndexOf($Screen)
}

function Expand-MbtGuiScenarioFold {
    <# Opens the fold ("More options", "Expert mode") whose header starts with the text, in the page that is shown. #>
    param([string]$Summary)
    $find = $null
    $find = {
        param($node)
        if ($node -is [Windows.Controls.Expander] -and $node.Header -is [Windows.Controls.TextBlock] -and $node.Header.Text -like "$Summary*") { $node.IsExpanded = $true }
        foreach ($c in [Windows.LogicalTreeHelper]::GetChildren($node)) { if ($c -is [Windows.DependencyObject]) { & $find $c } }
    }
    & $find $script:Gui.C.Page.Content
    $script:Gui.Window.Content.UpdateLayout()
}

function Get-MbtGuiFlowShots {
    <# The screens of the flow of the console, in order: the file name and the scenario that draws it. #>
    @(
        @{ File = '01-setup'; Scenario = 'Setup' }, @{ File = '02-task-what'; Scenario = 'TaskWhat' }, @{ File = '03-task-from'; Scenario = 'ExportFrom' }, @{ File = '04-task-to'; Scenario = 'ExportTo' }
        @{ File = '05-verify-export'; Scenario = 'ExportPlan' }, @{ File = '06-verify-import-unticked'; Scenario = 'ImportPlan' }, @{ File = '07-verify-import-ticked'; Scenario = 'ImportPlanTicked' }
        @{ File = '08-run-export'; Scenario = 'RunExport' }, @{ File = '09-next-export-complete'; Scenario = 'DoneExportComplete' }, @{ File = '10-next-import-partial'; Scenario = 'DoneImportPartial' }
        @{ File = '11-setup-application'; Scenario = 'SetupApplication' }, @{ File = '12-task-from-signed-out'; Scenario = 'ExportFromSignedOut' }, @{ File = '13-task-to-more-options'; Scenario = 'ExportToMore' }
        @{ File = '14-verify-export-expert'; Scenario = 'ExportPlanExpert' }, @{ File = '15-repair-from-damaged'; Scenario = 'RepairFromDamaged' }, @{ File = '16-verify-repair'; Scenario = 'RepairPlan' }
        @{ File = '17-next-repair-complete'; Scenario = 'DoneRepair' }, @{ File = '18-task-from-copy'; Scenario = 'CopyMailboxFrom' }, @{ File = '19-task-to-copy'; Scenario = 'CopyMailboxTo' }, @{ File = '20-verify-copy'; Scenario = 'CopyMailboxPlanTicked' }
        @{ File = '21-run-copy'; Scenario = 'RunCopyMailbox' }, @{ File = '22-next-copy-complete'; Scenario = 'DoneCopyMailbox' }
    )
}
function Set-MbtGuiScenarioTicks {
    <# Ticks folders of a browser by node id (a whole folder, with what is below it). #>
    param([string]$Use, [string[]]$Ids)
    $b = Get-MbtGuiBrowser $Use
    foreach ($r in $b.Roots) { foreach ($n in $r.Descendants()) { if ($n.Id -in $Ids) { $n.Check = $true } } }
    Update-MbtGuiTray $b
}

function Open-MbtGuiFirstFolder {
    param([string]$Use)
    $b = Get-MbtGuiBrowser $Use
    foreach ($r in $b.Roots) { foreach ($n in $r.Descendants()) { if ($n.Id -eq 'o-in') { $n.IsSelected = $true; Open-MbtGuiFolder $n; return } } }
}

function Initialize-MbtGuiScenarioSignIn {
    <# Signed in, the locations (-Kinds: only these kinds), the quota of the archive. #>
    param([string[]]$Kinds)
    $g = $script:Gui; $s = $g.S
    $s.Mailbox = 'adele.vance@contoso.com'; $s.Signed = $true; $s.Account = $s.Mailbox
    $s.Token = [pscustomobject]@{ Access = 'Delegated (you, with your rights)'; CanReadMail = $true; Missing = @() }
    & $g.Engine.Locations $s.Mailbox '' { param($l) $s.Locations = @($l | Where-Object { -not $Kinds -or $_.Kind -in $Kinds } | ForEach-Object { if ($_.Kind -eq 'Archive') { $_ | Add-Member Quota ([long](110GB)) -PassThru | Add-Member Used ([long](35.3GB)) -PassThru } else { $_ } }) } { }
}

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDayxaAfguEtx6a
# eVpzB7grKO9pJ7ha0+ruZgJQwTkpnaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYETCCGA0CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAQrmKajAod
# tgQmhpPZGwGNddR4M3kMWgyOHRL+RlsnBjANBgkqhkiG9w0BAQEFAASCAgAnI5F9
# mEc+5ApKa1By0V24ZsKLJjNSi3RB1MM/a9La2WKdiSYzPP3K8UhXFlbK9mfjtH1b
# 8W6+syNNkOPTuZ2TSLGTgmQyfSL7f9LmZEZMfEMM7PpOlIRRKGvQ/uAkL1GETw1B
# FaMDARypYvqBrxzFB0T5t9XZbObhnndrt20QJ8lufzPf8Dj/71a0Q21b/AcFvAQa
# ADR/2Gpucer1UO8pR7/0zwa9VJY3kaheOOsKIyhG3ic9hsJEkQW3WM+XU4FNBkr3
# D6mOosLVQyjeM2b1/o+jMupwy1LHTOaymqYNBp/pS/g3DHWwsIhMnyTEmawGRLT5
# d8fgwHvpRYWe50cF4AIvmI44ukendHB1r9mB6l5nDHW1CQrobEiWNJUjk/3HIbY8
# bMavxGOfC+CC/dlcUeNnka3XeetjiLSFi7/3lI1bqM6XdCvwqniYwP6qblSoivSG
# ulZ9DSFTYr6d7epDMn+trsp+S0R8DPV9jC1YYh/UPpAWg9MYC58YgeSN1wr8lmU8
# 6JVokYgnWcFxtxg6R8GCGt06ih/JajZ+eEERcHZfbYhNCmJ2ZkTuBiFCETuU9aQF
# E9hE057gNdQNR57st9nPKQVveOsMyhA/AmaKfHthIn9AFHh1WXqhaLkQqP2S0Gdw
# ekJe6EMKqmjKNcp2Mqbgw5Ryp6nJu5oLLWm0L6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEICmbsZf9ae1aa5j339ZqVboXEUqAbtDGz67JYG9ulfmzAghX17gxh4OQ
# GhgPMjAyNjExMTUxNDM1NDRaMAMCAQECEG0Ajotd/EMFJU1oex16kr6gghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlYwggJSAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTE0MzU0NFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIJ6B
# IgvCYRLzDg8KqC0f5SVpFooMnn6f3zSzXXgYTMAWMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiBkfJ99dRjjqw2poNMtZ0HY+20lCAiZ7vNAz9xNEOgyMgIhAKyD
# jGoi8ttY1ToA9XaepXIGr3QYdlvvPHKDL4Fe7+Ws
# SIG # End signature block
