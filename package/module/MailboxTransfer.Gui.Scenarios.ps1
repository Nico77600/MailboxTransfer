#Requires -Version 7.4
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
        'ImportFromChoice', 'ImportPlanChoice', 'DoneImportRefused', 'OpenPreview', 'CopyTo', 'CopyPlan', 'DoneCopy', 'ExportPlanShared', 'Setup', 'SetupApplication', 'TaskWhat', 'ExportToMore', 'ExportPlanExpert', 'ImportPlanTicked', 'ImportPlanChanged')
}

function New-MbtGuiEvt {
    param([string]$Kind, [datetime]$At, [hashtable]$Fields = @{})
    $o = [ordered]@{ v = 1; t = $At.ToString('o'); e = $Kind }
    foreach ($k in $Fields.Keys) { $o[$k] = $Fields[$k] }
    return [pscustomobject]$o
}

function New-MbtGuiScenarioRun {
    <# A run model for a task in a state, built from events. #>
    param([string]$Task, [string]$State = 'running')
    $run = New-MbtGuiRun -Task $Task -Title $(switch ($Task) { 'export' { 'Export ' + [char]0x00B7 + ' adele.vance@contoso.com' } 'import' { 'Import ' + [char]0x00B7 + ' archive-2019.pst ' + [char]0x2192 + ' adele.vance@contoso.com (Archive)' } 'copy' { 'Copy ' + [char]0x00B7 + ' archive-2019.pst ' + [char]0x2192 + ' archive-2019-copy.pst' } default { 'Repair ' + [char]0x00B7 + ' contoso-archive-2019.pst ' + [char]0x2192 + ' contoso-archive-2019.repaired.pst' } })
    $hours = switch ($Task) { 'import' { 6.33 } 'repair' { 0.05 } default { 0.18 } }
    $t0 = [datetime]::UtcNow.AddHours(-$hours)
    $at = { param($min) $t0.AddMinutes($min) }
    $ev = { param($k, $min, $f) $null = Add-MbtGuiEvent -Run $run -Event (New-MbtGuiEvt $k (& $at $min) $f) }
    switch ($Task) {
        'copy' {
            & $ev 'plan' 0 @{ steps = @(@{ id = 'pst'; title = 'Write the PST files' }, @{ id = 'pstcheck'; title = 'Check every PST file' }) }
            & $ev 'step' 0.5 @{ id = 'pst'; state = 'running'; done = 600; total = 1198; bytes = [long](0.5GB); bytesTotal = [long](1.1GB) }
            & $ev 'object' 0.5 @{ step = 'pst'; key = 'Projects/Contoso Rollout'; state = 'running'; done = 410; total = 780 }
        }
        'export' {
            & $ev 'plan' 0 @{ steps = @(@{ id = 'tree'; title = 'Read the folder tree' }, @{ id = 'list'; title = 'List the items you chose' }, @{ id = 'download'; title = 'Download the items' }, @{ id = 'verify'; title = 'Check the downloads' }, @{ id = 'pst'; title = 'Write the PST files' }, @{ id = 'pstcheck'; title = 'Check every PST file' }) }
            & $ev 'step' 0.2 @{ id = 'tree'; state = 'done'; msg = '4,439 folders' }
            & $ev 'step' 1.9 @{ id = 'list'; state = 'done'; msg = '1,208 items ' + [char]0x00B7 + ' 6.2 GiB'; total = 1208 }
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

function New-MbtGuiScenarioDone {
    <# A run model at its end: Task, Outcome complete / partial / decision / error. #>
    param([string]$Task, [string]$Outcome)
    $run = New-MbtGuiScenarioRun -Task $Task
    $exit = switch ($Outcome) { 'complete' { 0 } 'partial' { 2 } 'decision' { 3 } default { 1 } }   # refused: finished, some items refused for good (exit 1, reason refused-items)
    $counts = switch ("$Task/$Outcome") {
        'export/complete' { @{ done = 1208; failed = 0; bytes = [long](6.2GB); parts = 3 } }
        'export/partial' { @{ done = 1004; pending = 204; failed = 0 } }
        'export/error' { @{ done = 744; failed = 0 } }
        'import/complete' { @{ done = 9860; failed = 0; unknown = 0; pending = 0; total = 9860; foldersCreated = 214 } }
        'import/partial' { @{ done = 5120; failed = 2; unknown = 0; pending = 4738; total = 9860; foldersCreated = 214 } }
        'import/decision' { @{ done = 9859; failed = 0; unknown = 1; pending = 0; total = 9860; foldersCreated = 214 } }
        'import/error' { @{ done = 5120; failed = 2; unknown = 0; pending = 0; total = 9860 } }
        'import/refused' { @{ done = 9857; failed = 0; refused = 3; unknown = 0; pending = 0; total = 9860; foldersCreated = 214; report = 'D:\MailboxTransfer\jobs\20261109-1402-import-7c1e\work\pstimport-not-imported.csv' } }
        'copy/complete' { @{ itemsWritten = 1198; failed = 0; parts = 1; bytes = [long](1.1GB) } }
        'repair/complete' { @{ recovered = 1198; partial = 4; lost = 1; levelUsed = 'Rebuild'; destination = 'D:\PST\contoso-archive-2019.repaired.pst'; before = @{ errors = 6; warnings = 5 }; after = @{ errors = 0; warnings = 0 }; sizeBytes = [long](3.9GB) } }
        default { @{} }
    }
    $at = [datetime]::UtcNow
    if ($Outcome -eq 'error') {
        $err = switch ($Task) { 'export' { @{ code = 'disk-full'; msg = 'The disk D: is full. Free some space (about 5 GiB) and press Try again. The 744 items already downloaded are kept.'; tech = 'IOException: There is not enough space on the disk. (store pack-000003.bin)' } } 'import' { @{ code = 'signin-required'; msg = 'Your sign-in has expired. Sign in again to continue. Nothing is lost.'; tech = 'AADSTS50173: The provided grant has expired due to it being revoked.' } } default { @{ code = 'destination-exists'; msg = 'The destination file already exists. Choose another name. The repair never overwrites a file.'; tech = '' } } }
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
    if ($Task -eq 'import' -and $Outcome -eq 'decision') {
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
    if ($Task -eq 'copy') { $s.OpSrc = 'pst'; $s.OpPath = 'D:\PST\archive-2019.pst'; $s.OutFolder = 'D:\PST'; $s.OutName = 'archive-2019-copy.pst' }
    $g.Run = $Run; $g.Job = $null
    $idx = $script:Flows[$Task].IndexOf($(if ($Run.Result -and $Run.State -ne 'stopped') { 'done' } else { 'run' }))
    $s.Step = $idx; $s.Outcome = $Run.State
    Show-MbtGuiPage
}

function Initialize-MbtGuiScenarioExport {
    <# Signed in, the three locations, some folders and messages ticked. #>
    param([switch]$Ticks)
    $g = $script:Gui; $s = $g.S
    $s.Task = 'export'; $s.Sel = 'export'; $s.Mailbox = 'adele.vance@contoso.com'; $s.Signed = $true; $s.Account = $s.Mailbox
    $s.Token = [pscustomobject]@{ Access = 'Delegated (you, with your rights)'; CanReadMail = $true; Missing = @() }
    & $g.Engine.Locations $s.Mailbox '' { param($l) $s.Locations = @($l); foreach ($x in $l) { $s.LocationChecked[$x.Kind] = $true } } { }
    $browser = Get-MbtGuiBrowser 'export'
    & $g.Engine.Roots $s.Locations { param($r) Set-MbtGuiBrowserRoots -Browser $browser -Roots @($r) } { }
    $inbox = $null
    foreach ($r in $browser.Roots) { foreach ($n in $r.Descendants()) { if ($n.Id -eq 'inbox') { $inbox = $n } } }
    $inbox.IsExpanded = $true; $inbox.IsSelected = $true
    Open-MbtGuiFolder $inbox
    if ($Ticks) {
        foreach ($r in $browser.Roots) { foreach ($n in $r.Descendants()) { if ($n.Id -in 'p-c', 'fin-b', 'x-18') { $n.Check = $true } } }
        $rows = @($browser.Rows | Select-Object -First 5)
        foreach ($row in $rows) { $row.IsChecked = $true }
        $browser.Rows[0].IsChecked = $true
    }
    $g.ExportSelection = Get-MbtGuiSelection -Roots @($browser.Roots)
    $s.OutFolder = 'D:\Exports\adele'; $s.OutName = 'adele.pst'
}

function Set-MbtGuiScenario {
    <# Puts the window in a state. Names: Get-MbtGuiScenarioName. #>
    param([Parameter(Mandatory)][string]$Name)
    $g = $script:Gui; $s = $g.S
    $s.WorkFolder = 'D:\MailboxTransfer'
    switch ($Name) {
        'Start' { $s.Step = 1; Show-MbtGuiPage }
        'StartImport' { $s.Step = 1; $s.Sel = 'import'; Show-MbtGuiPage }
        'ExportFromSignedOut' { $s.Task = 'export'; $s.Sel = 'export'; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); $s.Mailbox = 'adele.vance@contoso.com'; Show-MbtGuiPage }
        'ExportFromCode' { $s.Task = 'export'; $s.Sel = 'export'; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); $s.Mailbox = 'adele.vance@contoso.com'; $s.SigningIn = $true; $s.CodeShown = $true; $s.Code = @{ Code = 'ABCD1EFGH'; Url = 'https://microsoft.com/devicelogin'; Until = (Get-Date).Date.AddHours(14).AddMinutes(32) }; Show-MbtGuiPage }
        'ExportFrom' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exFrom'); Show-MbtGuiPage }
        'ExportTo' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exTo'); Show-MbtGuiPage }
        'ExportPlan' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage }
        'RunExport' { Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioRun -Task export) }
        'RunImport' { Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioRun -Task import) }
        'RunImportSleep' { Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioRun -Task import -State sleep) }
        'RunImportDecision' { Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioRun -Task import -State decision) }
        'RunImportStopped' { Show-MbtGuiRunScenario 'import' (New-MbtGuiScenarioRun -Task import -State stopped) }
        'RunRepair' { Show-MbtGuiRunScenario 'repair' (New-MbtGuiScenarioRun -Task repair) }
        'DoneExportComplete' { Initialize-MbtGuiScenarioExport -Ticks; $s.OutName = 'adele.pst'; $s.OutFolder = 'D:\Exports\adele'; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioDone export complete) }
        'DoneExportPartial' { Initialize-MbtGuiScenarioExport -Ticks; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioDone export partial) }
        'DoneExportError' { Initialize-MbtGuiScenarioExport -Ticks; Show-MbtGuiRunScenario 'export' (New-MbtGuiScenarioDone export error) }
        'ImportFrom' { $s.Task = 'import'; $s.Sel = 'import'; $s.Step = (Get-MbtGuiScenarioStep 'import' 'imFrom'); $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\archive-2019.pst'; Open-MbtGuiSource -Use import }
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
        'ExportPlanShared' { Initialize-MbtGuiScenarioExport -Ticks; $s.Shared = 'projects@contoso.com'; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage }
        'Setup' { $s.Step = 0; Show-MbtGuiPage }
        'SetupApplication' { $s.Step = 0; $s.AuthMode = 'app'; $s.AppThumb = '5D41402ABC4B2A76B9719D911017C59228F3D1AE'; Show-MbtGuiPage }
        'TaskWhat' { $s.Step = 1; $s.Sel = 'export'; Show-MbtGuiPage }
        'ExportToMore' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exTo'); Show-MbtGuiPage; Expand-MbtGuiScenarioFold 'More options' }
        'ExportPlanExpert' { Initialize-MbtGuiScenarioExport -Ticks; $s.Step = (Get-MbtGuiScenarioStep 'export' 'exPlan'); Show-MbtGuiPage; Expand-MbtGuiScenarioFold 'Expert mode'; $script:Gui.C.PageScroll.ScrollToEnd() }
        'ImportPlanTicked' { Set-MbtGuiScenario 'ImportPlan'; $s.Confirmed = $true; $s.ConfirmKey = (Get-MbtGuiPlanKey); Show-MbtGuiPage }
        'ImportPlanChanged' { Set-MbtGuiScenario 'ImportPlanTicked'; $s.Mailbox = 'Adele.Vance@Contoso.com'; Show-MbtGuiPage; $s.TgtName = 'Imported archive-2019 2026-11-13'; Show-MbtGuiPage }
        default { throw "Unknown scenario '$Name'. Known: $((Get-MbtGuiScenarioName) -join ', ')." }
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
        @{ File = '17-next-repair-complete'; Scenario = 'DoneRepair' }
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
    $g = $script:Gui; $s = $g.S
    $s.Mailbox = 'adele.vance@contoso.com'; $s.Signed = $true; $s.Account = $s.Mailbox
    $s.Token = [pscustomobject]@{ Access = 'Delegated (you, with your rights)'; CanReadMail = $true; Missing = @() }
    & $g.Engine.Locations $s.Mailbox '' { param($l) $s.Locations = @($l | ForEach-Object { if ($_.Kind -eq 'Archive') { $_ | Add-Member Quota ([long](110GB)) -PassThru | Add-Member Used ([long](35.3GB)) -PassThru } else { $_ } }) } { }
}

