#Requires -Version 7.6
<#
.SYNOPSIS
    Mailbox Transfer console - starting a job, the Run page and the Done page (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    Run: three columns (steps, what the step works on, details) under a header card with the status, the overall bar, the counters and the honest
    ETA. Everything on it comes from the run model (MailboxTransfer.Gui.Run.ps1), which the events of the job feed. Done: one card, one headline,
    one pill, tiles, facts and the next step, by the exit code of the engine.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

#region Plan check and start -----------------------------------------------------------------------------

function Get-MbtGuiTaskSelection {
    <# What is ticked for the current task: the export tree, the tree of the PST or of the .eml folder to import, the tree of the PST that is open (copy). #>
    $g = $script:Gui; $s = $g.S
    $use = switch ($s.Task) { 'import' { 'import' } 'copy' { 'open' } default { $null } }
    if (-not $use) { return $g.ExportSelection }
    # an import takes everything unless "Only what I tick" is chosen: the ticks are then not looked at
    if ($use -eq 'import' -and [string]$s.ImSelMode -ne 'ticks') { return [pscustomobject]@{ Entries = @(); Items = 0L; Bytes = 0L; Folders = 0; UnknownSize = 0; SubtreeUnread = $false } }
    if ($g.Browsers.ContainsKey($use)) { return (Get-MbtGuiSelection -Roots @($g.Browsers[$use].Roots)) }
    return $null
}

function Write-MbtGuiChoiceFiles {
    <# The folders (and messages) ticked in a PST or in .eml files, written in the job folder for the entry script (folders.txt, items.txt). #>
    param([Parameter(Mandatory)][hashtable]$Job, $Selection)
    $s = $script:Gui.S
    $src = if ($s.Task -eq 'copy') { 'pst' } elseif ($s.Task -eq 'import' -and $s.ImSrc -in 'pst', 'eml') { $s.ImSrc } else { return }
    $choice = if ($Selection -and @($Selection.Entries).Count) { Get-MbtGuiFolderChoice -Selection $Selection -Source $src } else { [pscustomobject]@{ Include = @(); Exclude = @(); Nids = @() } }
    Write-MbtGuiFolderFiles -Job $Job -Choice $choice
}

function Get-MbtGuiPlanKey {
    <# The plan is valid for exactly this request: selection, destination, effective settings (any change voids it, as in PRA). #>
    $g = $script:Gui; $s = $g.S
    $sb = [Text.StringBuilder]::new()
    # what is typed or picked (addresses, names, paths) does not depend on the case; the ids of the folders ticked are machine ids and stay exact
    [void]$sb.Append("$($s.Task)|$($s.Dest)|$($s.OutFolder)|$($s.OutName)|$($s.SplitGB)|$($s.CheckEach)|$($s.SepAux)|$($s.KeepStore)|$($s.Conc)|$($s.Retry)|$($s.Tries)|$($s.Reconcile)|$($s.ImSrc)|$($s.ImPath)|$($s.TgtLoc)|$($s.TgtMode)|$($s.TgtName)|$($s.Mapping)|$($s.Shared)|$($s.ImShared)|$($s.McTarget)|$($s.RecoverableItems)|$($s.RpPath)|$($s.RpDest)|$($s.RpLevel)|$($s.RpKeep)|$($s.OpPath)|$($s.Mailbox)|$($s.AuthMode)|$($s.TenantId)|$($s.ClientId)|$($s.AppTenant)|$($s.AppClient)|$($s.AppThumb)|$($s.Awake)".ToLowerInvariant())
    $period = Get-MbtGuiPeriod $s
    [void]$sb.Append("|$($period.Active)|$($period.From)|$($period.To)|$($s.SelMode)|$($s.ImSelMode)|$($s.McLayout)|$($s.TgtMode2)|$($s.TgtName2)|$($s.PrimaryIdValue)|$($s.ArchiveIdValue)".ToLowerInvariant())
    $sel = Get-MbtGuiTaskSelection
    if ($sel) {
        foreach ($e in $sel.Entries) { [void]$sb.Append("|$($e.Kind):$($e.Node.Id):$($e.Items)") }
        if ($sel.PSObject.Properties['Places']) { foreach ($p in @($sel.Places)) { [void]$sb.Append("|all:$($p.Kind):$($p.Id)") } }
    }
    $typedArchive = Get-MbtGuiProp $s 'McArchive' $null
    if ($typedArchive -is [Collections.IDictionary]) { foreach ($k in @($typedArchive.Keys | Sort-Object)) { [void]$sb.Append("|archive:$k=$($typedArchive[$k])") } }
    $bytes = [Text.Encoding]::UTF8.GetBytes($sb.ToString())
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}

function Get-MbtGuiSelectionKey {
    <# What was chosen for the operation (the source, the mailbox, the ticks): the selection file and the folder files are named after it, so the same choice is the same file. #>
    $g = $script:Gui; $s = $g.S
    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append("$($s.Task)|$($s.Mailbox)|$($s.Shared)|$($s.ImSrc)|$($s.ImPath)|$($s.OpPath)|$($s.RecoverableItems)|$($s.SelMode)|$($s.ImSelMode)".ToLowerInvariant())
    $sel = Get-MbtGuiTaskSelection
    if ($sel) {
        foreach ($e in $sel.Entries) { [void]$sb.Append("|$($e.Kind):$($e.Node.Id):$($e.Items)") }
        if ($sel.PSObject.Properties['Places']) { foreach ($p in @($sel.Places)) { [void]$sb.Append("|all:$($p.Kind):$($p.Id)") } }
    }
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($sb.ToString())))
}

function Get-MbtGuiJobName {
    <# The name of the import job of this plan (the same in the line that is shown and in the run); a new plan gets a new one. #>
    $s = $script:Gui.S
    if (-not $s.JobName) { $s.JobName = "import-$(New-MbtGuiJobId)" }
    return $s.JobName
}

function New-MbtGuiPlanStateMessage {
    <# The first check of a Verify page: what the service said about this exact request. #>
    $s = $script:Gui.S
    $name = if ($s.Task -in 'import', 'mailcopy') { 'The plan is valid' } else { 'Every folder matched' }
    switch ($s.PlanState) {
        'checking' {
            $g = $script:Gui
            $run = if ($g.ContainsKey('PlanRun')) { $g.PlanRun } else { $null }
            $started = if ($g.ContainsKey('PlanStarted')) { $g.PlanStarted } else { $null }
            return (New-MbtGuiCheckRow -State wait -Text 'Checking the plan with the service' -Detail (Get-MbtGuiPlanProgressText -Run $run -Started $started))
        }
        'ok' {
            $atMost = (Get-MbtGuiPeriod $s).Active
            $counted = if ($s.PlanTotals) { "counted by the service: $(Format-MbtGuiPlanTotals $s.PlanTotals -AtMost:$atMost)" + $(if ($atMost) { ' (the period is applied to the received date while the items are listed)' }) + $(if ($s.PlanTotals.FoldersToCreate -gt 0) { "; $(Format-MbtGuiQuantity $s.PlanTotals.FoldersToCreate 'folder') will be created" } else { '' }) } else { '' }
            return (New-MbtGuiCheckRow -State ok -Text $name -Detail $counted)
        }
        'error' {
            $row = New-MbtGuiCheckRow -State bad -Text 'The plan cannot be used' -Detail (($s.PlanMsg -replace '^\*\*[^*]+\*\*\s*', '') -replace '\*\*', '')
            $again = New-GuiButton -Text 'Check again' -Small -Icon refresh -OnClick { Start-MbtGuiPlanCheck }
            $again.Margin = [Windows.Thickness]::new(26, 4, 0, 0); $again.HorizontalAlignment = 'Left'
            $panel = [Windows.Controls.StackPanel]::new()
            $panel.Children.Add($row) | Out-Null; $panel.Children.Add($again) | Out-Null
            return $panel
        }
        default { return (New-MbtGuiCheckRow -State wait -Text 'The plan is not checked yet') }
    }
}
function Get-MbtGuiPlanFooter {
    <# The text beside Start on a Verify page: why Start waits while the plan is checked, why it stays grey when the plan is refused, else what the page asks. #>
    $g = $script:Gui; $s = $g.S
    switch ($s.PlanState) {
        'ok' {
            switch (Get-MbtGuiScreen) {
                { $_ -in 'imPlan', 'mcPlan' } { return $(if ($s.Confirmed) { 'Ready. Nothing is written before you press Start.' } else { 'Tick the box to unlock Start.' }) }
                'cpPlan' { return 'Read the plan. Nothing starts before you press Start.' }
                default { return 'Read the checks and the plan. Nothing starts before you press Start.' }
            }
        }
        'error' { return 'The plan cannot be used. Press Check again, or go back and change your choice.' }
        default { return 'Start unlocks when the plan is checked.' }
    }
}

function Get-MbtGuiPlanFailureText {
    <# One sentence for a plan check that failed: a temporary problem of the service or of the network says to press Check again; any other known cause is said in its own words; else the sentence the engine gave. #>
    param([string]$Raw, [string]$Message = '')
    $m = ConvertTo-MbtGuiMessage $Raw
    switch ($m.Temporary) {
        'service' { return 'Microsoft did not answer correctly (a temporary problem). Press Check again.' }
        'network' { return 'We cannot reach Microsoft. Check your internet connection, then press Check again.' }
    }
    if ($m.Known) { return $m.Text }
    if ($Message) { return $Message }
    return 'The service did not accept the request.'
}

function Set-MbtGuiPlanState {
    param([ValidateSet('idle', 'checking', 'ok', 'error')][string]$State, [string]$Message = '')
    $g = $script:Gui; $s = $g.S
    $s.PlanState = $State; $s.PlanMsg = $Message
    if ((Get-MbtGuiScreen) -notin 'exPlan', 'imPlan', 'cpPlan', 'mcPlan') { return }
    if ($g.PlanHolder) { $g.PlanHolder.Content = New-MbtGuiPlanStateMessage }
    $ok = ($State -eq 'ok')
    if ((Get-MbtGuiScreen) -in 'imPlan', 'mcPlan') { $ok = $ok -and $s.Confirmed }
    if ($State -eq 'ok' -and (Get-MbtGuiScreen) -eq 'exPlan') { Update-MbtGuiExportPlanSizes }
    $footer = Get-MbtGuiPlanFooter
    Update-MbtGuiPrimary -Enabled $ok -Footer $footer
    $g.C.Footer.Tag = $footer
}

function Set-MbtGuiPrimaryBusy {
    <# The primary button says at once that its action is under way (disabled, with the text), and is painted before the handler goes on with the work. #>
    param([Parameter(Mandatory)][string]$Text, [string]$Footer = '')
    $g = $script:Gui; $c = $g.C
    $c.Primary.IsEnabled = $false
    $c.PrimaryLabel.Text = $Text
    $c.PrimaryIcon.Visibility = 'Collapsed'
    ConvertTo-MbtGuiRuns $c.Footer $Footer
    $c.Footer.Tag = $Footer
    try { $g.Window.Dispatcher.Invoke([Action] { }, [Windows.Threading.DispatcherPriority]::Render) } catch { Write-Verbose "not painted: $($_.Exception.Message)" }
}

function New-MbtGuiRequest {
    <# Everything a job needs, from the state: selection files (one per location), store, config. #>
    param([Parameter(Mandatory)][string]$Task)
    $g = $script:Gui; $s = $g.S
    $storeRoot = [IO.Path]::Combine($s.WorkFolder, 'stores')
    $store = switch ($Task) {
        'export' { [IO.Path]::Combine($storeRoot, (($s.Mailbox + (Get-MbtGuiPeriodStoreSuffix $s)) -replace '[^\w\.\-@]', '_')) }
        'import' { if ($s.ImSrc -eq 'store') { $s.ImPath } else { [IO.Path]::Combine($storeRoot, 'import') } }
        # a copy: one store for what is read (the mailbox, the ticks, the days), whatever the target; another choice has another store
        'mailcopy' { $key = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes((Get-MbtGuiSelectionKey) + (Get-MbtGuiPeriodStoreSuffix $s)))).Substring(0, 8).ToLowerInvariant(); [IO.Path]::Combine($storeRoot, ('copy-' + (($(if ($s.Shared) { $s.Shared } else { $s.Mailbox }) -replace '[^\w\.\-@]', '_') + '-' + $key))) }
        default { $null }
    }
    [pscustomobject]@{ Task = $Task; StorePath = $store }
}

function New-MbtGuiPreviewJob {
    <# The job the Start button would begin, without making its folder: the paths the command line names (the selection files are the ones the job will write). #>
    param([Parameter(Mandatory)][string]$Task)
    $g = $script:Gui; $s = $g.S
    $req = New-MbtGuiRequest -Task $Task
    $folder = [IO.Path]::Combine($s.WorkFolder, 'jobs', ((Get-Date).ToString('yyyyMMdd-HHmm') + "-$Task"))
    $base = Get-MbtGuiSelectionBase -WorkFolder $s.WorkFolder -Task $Task -Key (Get-MbtGuiSelectionKey)
    $sel = Get-MbtGuiTaskSelection
    $hasSel = [bool]($sel -and @($sel.Entries).Count)
    $files = @(); $folderFile = $null; $itemsFile = $null
    if ($hasSel -and ($Task -in 'export', 'mailcopy' -or ($Task -eq 'import' -and $s.ImSrc -eq 'store'))) { $files = @("$base.json") }
    $source = if ($Task -eq 'copy') { 'pst' } elseif ($Task -eq 'import' -and $s.ImSrc -in 'pst', 'eml') { $s.ImSrc } else { $null }
    if ($hasSel -and $source) {
        $choice = Get-MbtGuiFolderChoice -Selection $sel -Source $source
        if (@($choice.Include).Count -or @($choice.Exclude).Count) { $folderFile = "$base.folders.txt" }
        if (@($choice.Nids).Count) { $itemsFile = "$base.items.txt" }
    }
    return @{ Id = "$Task-job"; Task = $Task; Folder = $folder; EventsPath = [IO.Path]::Combine($folder, 'events.jsonl'); StopPath = [IO.Path]::Combine($folder, 'stop.request')
        StorePath = $req.StorePath; SelectionFiles = $files; JobName = (Get-MbtGuiJobName); FolderFile = $folderFile; ItemsFile = $itemsFile; SelectionBase = $base }
}

function Save-MbtGuiSelectionFile {
    <# The selection of the export (one selection per location, in one file). Calls OnDone with the path in an array. #>
    param([Parameter(Mandatory)]$Selection, [Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][scriptblock]$OnDone, [scriptblock]$OnError)
    $g = $script:Gui
    $groups = @($Selection.Entries | Group-Object { (Get-MbtGuiRootNode $_.Node).Id } | ForEach-Object { @{ Root = (Get-MbtGuiRootNode $_.Group[0].Node); Entries = @($_.Group) } })
    if (-not $groups.Count) { & $OnDone @(); return }
    & $g.Engine.SaveSelections $groups $Path { param($p) & $OnDone @([string]$p) }.GetNewClosure() $OnError
}

function Start-MbtGuiPlanCheck {
    <# Plan page: the read-only run of the engine (-WhatIf) of exactly this request. Start unlocks only when it is valid. #>
    $g = $script:Gui; $s = $g.S
    $s.PlanTotals = $null; $s.JobName = ''
    $g.PlanRun = $null; $g.PlanShown = ''; $g.PlanStarted = [datetime]::UtcNow
    Set-MbtGuiPlanState checking
    $req = New-MbtGuiRequest -Task $s.Task
    $base = Get-MbtGuiSelectionBase -WorkFolder $s.WorkFolder -Task $s.Task -Key (Get-MbtGuiSelectionKey)
    $fail = { param($msg) Set-MbtGuiPlanState error ("**The plan cannot be used.** $(Get-MbtGuiPlanFailureText -Raw $msg -Message $msg)") }
    if ($g.Engine.Name -eq 'Fake') {
        & $g.Engine.Plan $req { param($r) if ($r.Ok) { if ($r.PSObject.Properties['Totals']) { $s.PlanTotals = $r.Totals }; Set-MbtGuiPlanState ok } else { & $fail $r.Message } }.GetNewClosure() $fail
        return
    }
    $sel = Get-MbtGuiTaskSelection
    $job = New-MbtGuiJob -WorkFolder $s.WorkFolder -Task "$($s.Task)-plan" -EntryScript $g.EntryScript -StorePath $req.StorePath -JobName (Get-MbtGuiJobName) -SelectionBase $base
    $begin = {
        param($files)
        $job.SelectionFiles = @($files)
        Write-MbtGuiChoiceFiles -Job $job -Selection $sel
        foreach ($c in (Get-MbtGuiCommand -Task $s.Task -S $s -Job $job -Plan)) { $job.Queue.Enqueue($c) }
        $g.PlanRun = New-MbtGuiRun -Task $s.Task; $g.PlanJob = $job
        $null = Start-MbtGuiProcess -Job $job
    }.GetNewClosure()
    if ($sel -and $sel.Entries.Count -and $s.Task -in 'export', 'mailcopy') { Save-MbtGuiSelectionFile -Selection $sel -Path "$base.json" -OnDone $begin -OnError $fail }
    else { & $begin @() }
}

function Step-MbtGuiPlanJob {
    $g = $script:Gui
    if (-not $g.ContainsKey('PlanJob') -or -not $g.PlanJob) { return }
    $null = Step-MbtGuiJob -Job $g.PlanJob -Run $g.PlanRun
    # while the check runs the line says what it does and for how long (it is drawn again about once a second)
    if (-not $g.PlanRun.Result -and $g.S.PlanState -eq 'checking' -and $g.PlanHolder -and ((Get-MbtGuiScreen) -in 'exPlan', 'imPlan', 'cpPlan', 'mcPlan')) {
        $text = Get-MbtGuiPlanProgressText -Run $g.PlanRun -Started $g.PlanStarted
        if ($text -ne $g.PlanShown) { $g.PlanShown = $text; $g.PlanHolder.Content = New-MbtGuiPlanStateMessage }
    }
    if ($g.PlanRun.Result) {
        $r = $g.PlanRun.Result; $g.PlanJob = $null
        if ($r.Exit -eq 0) { $g.S.PlanTotals = Get-MbtGuiPlanTotals $g.PlanRun; Set-MbtGuiPlanState ok }
        else {
            $msg = if ($g.PlanRun.Error) { Get-MbtGuiPlanFailureText -Raw "$($g.PlanRun.Error.Code) $($g.PlanRun.Error.Msg) $($g.PlanRun.Error.Tech)" -Message $g.PlanRun.Error.Msg }
                elseif ($g.PlanRun.Warnings.Count) { $g.PlanRun.Warnings[0].Msg } else { 'The service did not accept the request.' }
            Set-MbtGuiPlanState error "**The plan cannot be used.** $msg"
        }
    }
}

function Start-MbtGuiRun {
    <# Start: the job folder, the command(s), the child process; the Run page opens. The button says at once that it is starting. #>
    $g = $script:Gui; $s = $g.S
    Set-MbtGuiPrimaryBusy -Text 'Starting...' -Footer 'Starting the task. This takes a few seconds.'
    $req = New-MbtGuiRequest -Task $s.Task
    $base = Get-MbtGuiSelectionBase -WorkFolder $s.WorkFolder -Task $s.Task -Key (Get-MbtGuiSelectionKey)
    $job = New-MbtGuiJob -WorkFolder $s.WorkFolder -Task $s.Task -EntryScript $g.EntryScript -StorePath $req.StorePath -JobName (Get-MbtGuiJobName) -SelectionBase $base
    $run = New-MbtGuiRun -Task $s.Task -Title (Get-MbtGuiRunTitle)
    $begin = {
        param($files)
        $job.SelectionFiles = @($files)
        Write-MbtGuiChoiceFiles -Job $job -Selection $sel
        if ($g.Engine.Name -eq 'Fake') { $job.Fake = $false; $job.EntryScript = $g.EntryScript }
        foreach ($c in (Get-MbtGuiCommand -Task $s.Task -S $s -Job $job)) { $job.Queue.Enqueue($c) }
        Write-MbtGuiRequestFile -Job $job
        $g.Job = $job; $g.Run = $run; $s.RStep = $null; $s.RObj = $null
        $null = Start-MbtGuiProcess -Job $job
        Move-MbtGuiStep 1
    }.GetNewClosure()
    $sel = Get-MbtGuiTaskSelection
    # ticks choose folders of a mailbox (export) or of an export store (import); the folders ticked in a PST or in .eml files go to the engine as a list of folders (Write-MbtGuiChoiceFiles)
    $needsSel = $s.Task -in 'export', 'mailcopy' -or ($s.Task -eq 'import' -and $s.ImSrc -eq 'store' -and $sel -and $sel.Entries.Count)
    if ($needsSel) { Save-MbtGuiSelectionFile -Selection $sel -Path "$base.json" -OnDone $begin -OnError { param($m) Set-MbtGuiPlanState error ((ConvertTo-MbtGuiMessage $m).Text); Show-MbtGuiPage } }
    else { & $begin @() }
}

function Write-MbtGuiRequestFile {
    <# request.json: what was asked, no secrets; the Continue list and a resumed job read it. #>
    param([Parameter(Mandatory)][hashtable]$Job)
    $s = $script:Gui.S
    $keys = 'Task', 'OpSrc', 'OpPath', 'Mailbox', 'Shared', 'ImShared', 'McTarget', 'McFromKind', 'AuthMode', 'RecoverableItems', 'Period', 'ArrFrom', 'ArrTo', 'Dest', 'KeepStore', 'SplitGB', 'CheckEach', 'SepAux', 'OutFolder', 'OutName', 'Conc', 'Retry', 'Tries', 'Reconcile', 'Awake', 'ImSrc', 'ImPath', 'TgtLoc', 'TgtMode', 'TgtName', 'TgtMode2', 'TgtName2', 'McLayout', 'McArchive', 'SelMode', 'ImSelMode', 'EvPlaces', 'PrimaryIdValue', 'ArchiveIdValue', 'Mapping', 'RpPath', 'RpLevel', 'RpDest', 'RpKeep'
    $req = [ordered]@{ Title = (Get-MbtGuiRunTitle); Created = [datetime]::UtcNow.ToString('o'); JobId = $Job.Id; JobName = $Job.JobName; StorePath = $Job.StorePath; SelectionFiles = @($Job.SelectionFiles); Commands = @($Job.Queue | ForEach-Object { , $_.Arguments }) }
    foreach ($k in $keys) { $req[$k] = $s[$k] }
    $req | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $Job.Folder 'request.json') -Encoding utf8
}

function Get-MbtGuiRunTitle {
    $s = $script:Gui.S
    $mid = [char]0x00B7; $arrow = [char]0x2192
    switch ($s.Task) {
        'export' { return "Export $mid $($s.Mailbox)" }
        'import' { return "Import $mid $([IO.Path]::GetFileName($s.ImPath)) $arrow $($s.Mailbox) ($($s.TgtLoc))" }
        'repair' { return "Repair $mid $([IO.Path]::GetFileName($s.RpPath)) $arrow $([IO.Path]::GetFileName($s.RpDest))" }
        'copy' { return "Copy $mid $([IO.Path]::GetFileName($s.OpPath)) $arrow $($s.OutName)" }
        'mailcopy' { $ends = Get-MbtGuiMailCopyEnds; return "Copy $mid $($ends.From) ($($ends.FromKind)) $arrow $($ends.To) ($($ends.ToKind))" }
        default { return $script:Tasks[$s.Task].Title }
    }
}

function Open-MbtGuiBatchReport {
    <# The report of a batch: the file in its folder, written again by the worker when it is not there (the batch was made by an older version, or the file was removed). #>
    param([Parameter(Mandatory)]$Entry)
    $g = $script:Gui
    $say = { param([string]$Title, [string]$Body) Show-MbtGuiDialog -Title $Title -Body $Body -Buttons @(@{ Text = 'Close'; Accent = $true }) }
    $open = { param($path) if ($path -and (Test-Path -LiteralPath ([string]$path))) { Start-Process -FilePath ([string]$path) } else { & $say 'Nothing to open here' 'No file was written: this is the demonstration.' } }.GetNewClosure()
    $report = [string](Get-MbtGuiProp $Entry 'Report' '')
    if ($report -and (Test-Path -LiteralPath $report)) { & $open $report; return }
    $fail = { param($message) & $say 'The report of the batch could not be written' ([string]$message) }.GetNewClosure()
    try { & $g.Engine.BatchReport $g.S.WorkFolder ([string]$Entry.Batch) $open $fail } catch { & $fail $_.Exception.Message }
}

function Open-MbtGuiJobReport {
    <# The report of a job of the list: the file in its folder, written again by the worker when it is not there. #>
    param([Parameter(Mandatory)]$Entry)
    $g = $script:Gui
    $say = { param([string]$Title, [string]$Body) Show-MbtGuiDialog -Title $Title -Body $Body -Buttons @(@{ Text = 'Close'; Accent = $true }) }
    $open = { param($path) if ($path -and (Test-Path -LiteralPath ([string]$path))) { Start-Process -FilePath ([string]$path) } else { & $say 'Nothing to open here' 'No file was written: this is the demonstration.' } }.GetNewClosure()
    $folder = [string](Get-MbtGuiProp $Entry 'Folder' '')
    if (-not $folder) { & $open $null; return }
    $report = Join-Path $folder 'report.html'
    if (Test-Path -LiteralPath $report) { & $open $report; return }
    $fail = { param($message) & $say 'The report of the job could not be written' ([string]$message) }.GetNewClosure()
    try { & $g.Engine.JobReport $folder $open $fail } catch { & $fail $_.Exception.Message }
}

function Show-MbtGuiJobDelete {
    <# Delete on a job of the list: the engine says what the deletion would do (and refuse), then the dialog names the job and says exactly what goes. #>
    param([Parameter(Mandatory)]$Entry)
    $g = $script:Gui
    $name = [string](Get-MbtGuiProp $Entry 'Name' '')
    $fail = { param($message) Show-MbtGuiDialog -Title 'The job cannot be deleted' -Body ([string]$message) -Buttons @(@{ Text = 'Close'; Accent = $true }) }
    $ask = { param($info) Show-MbtGuiJobDeleteQuestion -Entry $Entry -Info @($info)[0] }.GetNewClosure()
    try { & $g.Engine.JobDeleteInfo $g.S.WorkFolder $name $ask $fail } catch { & $fail $_.Exception.Message }
}

function Show-MbtGuiJobDeleteQuestion {
    param([Parameter(Mandatory)]$Entry, [Parameter(Mandatory)]$Info)
    $g = $script:Gui
    $name = [string](Get-MbtGuiProp $Entry 'Name' '')
    if (-not [bool](Get-MbtGuiProp $Info 'Allowed' $false)) {
        Show-MbtGuiDialog -Title 'This job cannot be deleted' -Body ([string](Get-MbtGuiProp $Info 'Reason' 'The job is kept.')) -Buttons @(@{ Text = 'Close'; Accent = $true })
        return
    }
    $when = { param($v) if ($v -is [datetime]) { $v.ToLocalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture) } else { '' } }
    $lines = [Collections.Generic.List[object]]::new()
    $lines.Add((New-GuiText ('**' + [string](Get-MbtGuiProp $Info 'Title' $name) + '**') 'Body'))
    $source = [string](Get-MbtGuiProp $Info 'Source' ''); $destination = [string](Get-MbtGuiProp $Info 'Destination' '')
    if ($source -or $destination) { $lines.Add((New-GuiText ((@($source, $destination) | Where-Object { $_ }) -join (' ' + [char]0x2192 + ' ')) 'Small')) }
    $dates = @(); $started = & $when (Get-MbtGuiProp $Info 'Started'); $updated = & $when (Get-MbtGuiProp $Info 'Updated')
    if ($started) { $dates += "started $started" }
    if ($updated) { $dates += "last activity $updated" }
    if ($dates.Count) { $lines.Add((New-GuiText ($dates -join (' ' + [char]0x00B7 + ' ')) 'Small')) }
    $lines.Add((New-GuiText ('This deletes the folder of the job on this PC: its log, its report and its events. ' + [string](Get-MbtGuiProp $Info 'Folder' '')) 'Body' -Margin '0,10,0,0'))
    $lines.Add((New-GuiText 'Your data is never deleted: the export store you chose, PST files, source files and the mailboxes are not touched.' 'Body' -Margin '0,6,0,0'))
    foreach ($w in @(Get-MbtGuiProp $Info 'Warnings' @())) { $lines.Add((New-GuiMessage -Kind warn -Icon warn -Text ([string]$w))) }
    $store = Get-MbtGuiProp $Info 'Store'
    $check = $null
    if ($store -and [bool](Get-MbtGuiProp $store 'Offered' $false)) {
        $bytes = Get-MbtGuiProp $store 'Bytes'
        $size = $(if ($null -ne $bytes) { Format-MbtGuiBytes ([double]$bytes) } else { 'size unknown' })
        $allowed = [bool](Get-MbtGuiProp $store 'Allowed' $false)
        $hint = if ($allowed) { [string]$store.Path + ': the items read from the source mailbox. The copy is complete, so nothing reads them any more.' } else { [string]$store.Reason }
        $check = New-GuiOption -Type Check -Title "Also delete the copy store ($size)" -Hint $hint -Checked $false -Disabled:(-not $allowed) -Margin '0,12,0,0'
        [Windows.Automation.AutomationProperties]::SetName($check, 'Also delete the copy store')
        $lines.Add($check)
    } elseif ($store -and (Get-MbtGuiProp $store 'Reason')) { $lines.Add((New-GuiText ([string]$store.Reason) 'Hint')) }
    $body = New-GuiStack $lines.ToArray()
    Show-MbtGuiDialog -Title 'Delete this job?' -Body $body -Buttons @(
        @{ Text = 'Cancel' }
        @{ Text = 'Delete the job'; Danger = $true; Action = { Invoke-MbtGuiJobDelete -Entry $Entry -WithStore ([bool]($check -and $check.IsChecked)) }.GetNewClosure() })
}

function Invoke-MbtGuiJobDelete {
    param([Parameter(Mandatory)]$Entry, [bool]$WithStore)
    $g = $script:Gui
    $name = [string](Get-MbtGuiProp $Entry 'Name' '')
    $fail = { param($message) Show-MbtGuiDialog -Title 'The job could not be deleted' -Body ([string]$message) -Buttons @(@{ Text = 'Close'; Accent = $true }); Show-MbtGuiPage }
    $done = { param($result) Show-MbtGuiPage }
    try { & $g.Engine.JobDelete $g.S.WorkFolder $name $WithStore $done $fail } catch { & $fail $_.Exception.Message }
}

function Resume-MbtGuiJob {
    <# Continue on a job of the list: the Resume view first (what is done, what is checked again, where it restarts); the job starts when the button of that view is pressed. #>
    param($Entry)
    $g = $script:Gui; $s = $g.S
    if ($g.Engine.Name -eq 'Fake' -and $Entry.PSObject.Properties['Demo'] -and $Entry.Demo) { & $Entry.Demo; return }
    # a batch is run from the command line: the console shows its report, read only
    if ($Entry.PSObject.Properties['Task'] -and $Entry.Task -eq 'batch') { Open-MbtGuiBatchReport $Entry; return }
    $folder = [string](Get-MbtGuiProp $Entry 'Folder' '')
    $name = [string](Get-MbtGuiProp $Entry 'Name' '')
    $info = $null
    try { if ($name) { $info = @(& $g.Engine.JobResume $s.WorkFolder $name)[0] } } catch { Write-Verbose "resume: $_" }
    if (-not $info) { Show-MbtGuiDialog -Title 'This job cannot be read' -Body 'Its folder or its status file is not there any more.' -Buttons @(@{ Text = 'Close'; Accent = $true }); return }
    $req = $null
    $reqFile = if ($folder) { Join-Path $folder 'request.json' } else { '' }
    if ($reqFile -and (Test-Path -LiteralPath $reqFile)) {
        $req = Get-Content -LiteralPath $reqFile -Raw | ConvertFrom-Json
        foreach ($p in $req.PSObject.Properties) { if ($s.ContainsKey($p.Name) -and $p.Name -notin 'Task') { $s[$p.Name] = $p.Value } }
        $s.Task = [string]$req.Task; $s.Sel = $s.Task
    }
    if (-not $s.Task -or -not $script:Flows.ContainsKey($s.Task)) {
        Show-MbtGuiDialog -Title 'This job cannot be continued here' -Body ([string](Get-MbtGuiProp $info 'Problem' 'Its request says nothing the console can run again. Run its command again from the command line.')) -Buttons @(@{ Text = 'Close'; Accent = $true })
        return
    }
    $s.Resume = @{ Entry = $Entry; Info = $info; Request = $req }
    $s.Step = [math]::Max(0, $script:Flows[$s.Task].IndexOf('run') - 1)
    Show-MbtGuiPage
}

function New-MbtGuiResumeList {
    <# A list of sentences under a small heading, each with the icon of its kind. #>
    param([string]$Heading, [string]$Icon, [string]$Brush, [string[]]$Lines)
    $stack = [Windows.Controls.StackPanel]::new()
    $stack.Children.Add((New-GuiText $Heading 'Label')) | Out-Null
    foreach ($line in @($Lines)) {
        $row = [Windows.Controls.Grid]::new(); $row.Margin = [Windows.Thickness]::new(0, 3, 0, 3)
        $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::new(24)
        $c2 = [Windows.Controls.ColumnDefinition]::new(); $c2.Width = [Windows.GridLength]::new(1, 'Star')
        $row.ColumnDefinitions.Add($c1); $row.ColumnDefinitions.Add($c2)
        $ic = New-GuiIcon $Icon 13 $Brush; $ic.VerticalAlignment = 'Top'; $ic.Margin = [Windows.Thickness]::new(0, 2, 0, 0)
        $tx = New-GuiText $line 'Body'; $tx.FontSize = 13; [Windows.Controls.Grid]::SetColumn($tx, 1)
        $row.Children.Add($ic) | Out-Null; $row.Children.Add($tx) | Out-Null
        $stack.Children.Add($row) | Out-Null
    }
    return $stack
}

function Build-MbtGuiResumePage {
    <# The Resume view: what happened to the job, and for its mode what is already done and not done again, what is checked again, where it restarts. Nothing starts before Continue is pressed. #>
    $g = $script:Gui; $s = $g.S
    $info = $s.Resume.Info
    $problem = [string](Get-MbtGuiProp $info 'Problem' '')
    $kind = [string](Get-MbtGuiProp $info 'Kind' '')
    $label = switch ($kind) { 'interrupted' { 'Interrupted' } 'stopped' { 'Stopped by you' } 'error' { 'Stopped by an error' } 'decision' { 'Needs a decision' } 'signin' { 'Sign-in expired' } default { 'Not finished' } }
    $pill = New-GuiPill $label $(if ($kind -eq 'error') { 'bad' } else { 'warn' })
    $left = [Collections.Generic.List[object]]::new()
    $left.Add((New-GuiText ([string](Get-MbtGuiProp $info 'Title' '')) 'Title'))
    $why = [string](Get-MbtGuiProp $info 'Why' ''); if ($why) { $left.Add((New-GuiText $why 'Body' -Margin '0,6,0,0')) }
    $progress = [string](Get-MbtGuiProp $info 'Progress' ''); if ($progress) { $left.Add((New-GuiText $progress 'Muted' -Margin '0,4,0,0')) }
    $facts = @()
    $started = Get-MbtGuiProp $info 'Started'; $updated = Get-MbtGuiProp $info 'Updated'
    if ($started -is [datetime]) { $facts += , @('Started', $started.ToLocalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)) }
    if ($updated -is [datetime]) { $facts += , @('Last sign of life', $updated.ToLocalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)) }
    if ($facts.Count) { $left.Add((New-GuiFacts $facts)) }
    $leftCard = New-GuiCard -Title 'The job' -Right $pill -Content $left.ToArray()
    $right = [Collections.Generic.List[object]]::new()
    $right.Add((New-GuiText ([string](Get-MbtGuiProp $info 'Heading' 'Continue the job')) 'Body'))
    if ($problem) { $right.Add((New-GuiMessage -Kind bad -Icon x -Text $problem)) }
    else {
        $right.Add((New-MbtGuiResumeList -Heading 'Already done, and not done again' -Icon check -Brush MbtSuccess -Lines @(Get-MbtGuiProp $info 'Already' @())))
        $right.Add((New-MbtGuiResumeList -Heading 'Checked again' -Icon refresh -Brush MbtBrandText -Lines @(Get-MbtGuiProp $info 'Again' @())))
        $from = [string](Get-MbtGuiProp $info 'From' '')
        if ($from) { $right.Add((New-MbtGuiResumeList -Heading 'Where it restarts' -Icon play -Brush MbtBrandText -Lines @($from))) }
        foreach ($note in @(Get-MbtGuiProp $info 'Notes' @())) { $right.Add((New-GuiMessage -Kind info -Icon info -Text ([string]$note))) }
    }
    $rightCard = New-GuiCard -Title 'When you press Continue' -Content $right.ToArray()
    $back = New-GuiButton -Text 'Back to the jobs' -Icon refresh -OnClick { $s = $script:Gui.S; $s.Resume = $null; $s.Task = $null; $s.Sel = $null; $s.Step = 1; Show-MbtGuiPage }
    Set-MbtGuiActions -Back $false -Left $back -Footer $(if ($problem) { 'This job cannot be continued from here.' } else { 'Nothing starts before you press Continue. Everything done so far is kept.' }) -Label 'Continue' -Icon play -Enabled (-not $problem) -OnClick { Start-MbtGuiResumeRun }
    return (New-GuiStack @((New-GuiText 'Resume this job' 'H2'), (New-GuiTwoColumns -Left $leftCard -Right $rightCard -LeftWidth 380)))
}

function Start-MbtGuiResumeRun {
    <# Continue on the Resume view: the recorded command again, in the same job folder, so only what is left is done. #>
    $g = $script:Gui; $s = $g.S
    $resume = $s.Resume
    if (-not $resume) { return }
    $Entry = $resume.Entry; $req = $resume.Request
    $s.Resume = $null
    if ($g.Engine.Name -eq 'Fake' -and (Get-MbtGuiProp $Entry 'ResumeDemo')) { & $Entry.ResumeDemo; return }
    if (-not $req) { return }
    $job = New-MbtGuiJob -WorkFolder $s.WorkFolder -Task $s.Task -EntryScript $g.EntryScript -StorePath ([string](Get-MbtGuiProp $req 'StorePath' '')) -SelectionFiles @(Get-MbtGuiProp $req 'SelectionFiles' @()) -JobName ([string](Get-MbtGuiProp $req 'JobName' ''))
    $job.Folder = $Entry.Folder
    foreach ($k in 'EventsPath', 'StopPath', 'StatusPath', 'StdOut', 'StdErr') { $job[$k] = Join-Path $Entry.Folder (Split-Path $job[$k] -Leaf) }
    if (Test-Path -LiteralPath $job.StopPath) { Remove-Item -LiteralPath $job.StopPath -Force }
    # the events of the earlier runs stay in the file (with their result): the new run reads only what the new process writes
    $job.Offset = $(if (Test-Path -LiteralPath $job.EventsPath) { (Get-Item -LiteralPath $job.EventsPath).Length } else { 0L })
    if ([string](Get-MbtGuiProp $req 'Source' '') -eq 'cli') {
        # a job of the command line runs again by its name: the recorded command, in the folder it ran in
        $job.ResumeArguments = @('-ResumeJob', [string]$Entry.Name, '-WorkFolder', $s.WorkFolder, '-StopFile', $job.StopPath, '-NoPrompt')
        $job.Queue.Enqueue([pscustomobject]@{ Title = 'Continue'; Arguments = $job.ResumeArguments })
    } else {
        foreach ($c in @($req.Commands)) { $job.Queue.Enqueue([pscustomobject]@{ Title = 'Continue'; Arguments = @($c) }) }
    }
    $g.Job = $job; $g.Run = New-MbtGuiRun -Task $s.Task -Title ([string]$req.Title)
    $g.Run.Resumed = $true
    $g.Run.Dirty = $true
    $null = Start-MbtGuiProcess -Job $job
    Set-MbtGuiStep ($script:Flows[$s.Task].IndexOf('run'))
}

#endregion

#region Run page ------------------------------------------------------------------------------------------------

function Get-MbtGuiSelectedStep {
    <# The step shown in the middle column: the one clicked, else the first that runs, else the first. #>
    param($Run, $S)
    if ($null -ne $S.RStep) { return [int]$S.RStep }
    for ($i = 0; $i -lt $Run.Steps.Count; $i++) { if ($Run.Steps[$i].State -eq 'running') { return $i } }
    return 0
}

function Get-MbtGuiStepTone {
    param($Step, [string]$RunState)
    switch ($Step.State) {
        'done' { return @('ok', 'check') }
        'failed' { return @('bad', 'x') }
        'running' { if ($RunState -in 'sleep', 'stopped') { return @('warn', 'pause') } else { return @('run', 'play') } }
        default { return @('', 'clock') }
    }
}

function Get-MbtGuiRunStepControl {
    <# The button the icon of the running step becomes: 'pause' while the job runs (when the Stop button of the action bar is shown), 'resume' once it is stopped (when Continue is shown), $null otherwise. #>
    param($Run)
    if ($Run.State -eq 'stopped') { return 'resume' }
    if ($Run.State -in 'running', 'sleep', 'decision' -and -not $Run.StopRequested -and -not $Run.Result) { return 'pause' }
    return $null
}

function New-MbtGuiStepControlButton {
    <# The pause or resume button in the place of the icon of the running step: a small icon button, reachable with the keyboard. Pause opens the dialog of Stop and keep progress; Resume does what Continue does. #>
    param([ValidateSet('pause', 'resume')][string]$Kind, [string]$Brush)
    $btn = [Windows.Controls.Button]::new()
    $btn.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtIconButton')
    $text = if ($Kind -eq 'pause') { 'Pause (stop and keep progress)' } else { 'Resume (continue where it stopped)' }
    $btn.Content = New-GuiIcon $(if ($Kind -eq 'pause') { 'pause' } else { 'play' }) 13 $Brush
    $btn.ToolTip = $text
    $btn.VerticalAlignment = 'Top'; $btn.HorizontalAlignment = 'Left'
    $btn.Margin = [Windows.Thickness]::new(-4, -3, 0, 0)
    [Windows.Automation.AutomationProperties]::SetName($btn, $text)
    [Windows.Automation.AutomationProperties]::SetAutomationId($btn, $(if ($Kind -eq 'pause') { 'RunStepPause' } else { 'RunStepResume' }))
    if ($Kind -eq 'pause') { Add-MbtGuiHandler $btn 'Click' { Show-MbtGuiStopDialog } }
    else { Add-MbtGuiHandler $btn 'Click' { Invoke-MbtGuiContinue } }
    return $btn
}

function Get-MbtGuiStepSub {
    param($Step)
    if ($Step.State -eq 'waiting') { return $Step.Msg }
    if ($Step.Total -gt 0 -and $Step.State -eq 'running') { return (Format-MbtGuiFraction $Step.Done $Step.Total -Unit 'item') }
    if ($Step.Msg) { return $Step.Msg }
    if ($Step.Total -gt 0) { return (Format-MbtGuiFraction $Step.Done $Step.Total -Unit 'item' -Finished) }
    return ''
}

function Get-MbtGuiStepHelp {
    param([string]$Task, [string]$Id, [string]$Title)
    switch -Regex ("$Id $Title") {
        '^resume|Resume: reading' { return 'Reads what the earlier run of this job already did (the store, the journal or the manifest it keeps), so that only what is left is done. Nothing is changed.' }
        '^tree|Read the folder tree' { return 'Reads the folders of the mailbox, 4 requests at a time. A mailbox with thousands of folders takes a minute or two. Nothing is changed.' }
        '^list|List the items' { return 'Lists the items of the folders you chose, so that the number and the size are known before they are downloaded. With a period, only the messages received in it are kept (the arrival date for what has no received date).' }
        '^verify|Check the downloads' { return 'Each item was checked against its fingerprint when it arrived in the store. Nothing else to do.' }
        '^check|Verify the export store' { return 'The second part begins: the store that was just filled is opened and verified before anything is written into the other mailbox.' }
        '^folders|Create the folders' { return 'Creates the folders that are missing in the mailbox that receives the copy. A folder that exists is used as it is; none is renamed or deleted.' }
        '^final|Final check' { return 'Counts what is done in the job, the earlier runs included. The job is complete when every item is done.' }
        'send|Send the items' { return 'Reads each message and sends it to the mailbox. Only creates; never deletes or replaces.' }
        'reconcile|uncertain' { return 'When Microsoft did not answer, we cannot know if an item arrived. We look in the folder and decide for each one.' }
        'download|Download' { return 'Downloads the items, 20 at a time, and checks each against its fingerprint. An item that fails is tried again.' }
        'write|Write the new file' { return $(if ($Task -eq 'repair') { 'Writes a new, clean PST with what could be read. The damaged original is not touched.' } else { 'Writes the PST files with the real folder tree. A message is never cut in two.' }) }
        default { return 'Click a step to see what it does.' }
    }
}

function New-MbtGuiRunHead {
    $g = $script:Gui; $run = $g.Run; $s = $g.S
    $st = $run.State
    $pill = switch ($st) { 'running' { New-GuiPill 'Running' 'run' } 'sleep' { New-GuiPill 'Resumed after sleep' 'warn' } 'stopped' { New-GuiPill 'Stopped' 'warn' } 'decision' { New-GuiPill 'Needs a decision' 'warn' } default { New-GuiPill 'Finishing' 'run' } }
    $frac = if ($run.BytesTotal -gt 0) { $run.Bytes / $run.BytesTotal } elseif ($run.Total -gt 0) { $run.Done / $run.Total } else { 0 }
    $bar = New-GuiProgressBar -Fraction $frac -Tone $(if ($st -in 'running', 'finishing') { 'run' } else { 'warn' }) -Height 8
    $dock = [Windows.Controls.DockPanel]::new()
    [Windows.Controls.DockPanel]::SetDock($pill, 'Right'); $pill.Margin = [Windows.Thickness]::new(12, 0, 0, 0); $dock.Children.Add($pill) | Out-Null
    $title = New-GuiText $run.Title 'Big' -NoWrap; $title.TextTrimming = 'CharacterEllipsis'; $dock.Children.Add($title) | Out-Null
    $kp = [Windows.Controls.WrapPanel]::new(); $kp.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
    $kpi = { param($text) $t = New-GuiText $text 'Muted' -NoWrap; $t.FontSize = 13; $t.Margin = [Windows.Thickness]::new(0, 0, 22, 2); $kp.Children.Add($t) | Out-Null }
    $preparing = Get-MbtGuiPreparingText $run
    if ($preparing) { & $kpi **Preparing** } else { & $kpi ((Format-MbtGuiProgressOf $run.Done $run.Total $run.Unit) -replace '^([\d,]+)', '**$1**') }
    if ($run.BytesTotal -gt 0) { & $kpi ((Format-MbtGuiBytesOf $run.Bytes $run.BytesTotal) -replace '^(\S+ \S+)', '**$1**') }
    $elapsed = if ($run.Started) { Format-MbtGuiDuration (([datetime]::UtcNow - $run.Started).TotalSeconds) } else { '0 s' }
    if ($run.LastEvent -and $run.Started -and $run.State -in 'stopped') { $elapsed = Format-MbtGuiDuration (($run.LastEvent - $run.Started).TotalSeconds) }
    & $kpi ("Elapsed **$elapsed**")
    if ($run.Waits.Count -gt 0) { & $kpi ("Microsoft slowed us down **$(Format-MbtGuiQuantity $run.Waits.Count 'time')** ($(Format-MbtGuiDuration $run.Waits.Seconds -WithSeconds) waited)") }
    $items = [Collections.Generic.List[object]]::new()
    foreach ($x in @($dock, $bar, $kp)) { $items.Add($x) }
    $bar.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    if ($run.Waits.Active -and $run.Waits.Active.Reason -eq 'throttle') {
        $items.Add((New-GuiMessage -Kind info -Icon info -Text "Microsoft asks us to slow down. Waiting $(Format-MbtGuiQuantity ([math]::Round($run.Waits.Active.Seconds)) 'second'). This is normal."))
    }
    if ($preparing) { $items.Add((New-GuiMessage -Kind info -Icon info -Text $preparing)) }
    switch ($st) {
        'sleep' {
            $from = if ($run.Sleep.From) { $run.Sleep.From.ToLocalTime().ToString('HH:mm') } else { '?' }; $to = if ($run.Sleep.To) { $run.Sleep.To.ToLocalTime().ToString('HH:mm') } else { '?' }
            $items.Add((New-GuiMessage -Kind warn -Icon moon -Text "**This PC slept from $from to $to.** The job paused and has now started again. The time left is being measured again; it shows after about 10 minutes."))
        }
        'stopped' {
            $items.Add((New-GuiMessage -Kind info -Icon pause -Text "**Stopped by you.** $(Format-MbtGuiProgressOf $run.Done $run.Total $run.Unit) $(Get-MbtGuiPlural $(if ($run.Done -gt $run.Total) { $run.Done } else { $run.Total }) 'is' 'are') done. Press **Continue** to do only what is left."))
        }
        default {
            $eta = Get-MbtGuiEta -Run $run -Task $(if ($s.Task -eq 'mailcopy') { if ($run.ImportHalf) { 'import' } else { 'export' } } else { $s.Task })
            if ($eta.State -eq 'ready') { $e = New-GuiText ("**$($eta.Text)**") 'Body'; $e.FontSize = 13; $e.Margin = [Windows.Thickness]::new(0, 8, 0, 0); $b2 = New-GuiText $eta.Basis 'Small'; $b2.Margin = [Windows.Thickness]::new(0, 2, 0, 0); $items.Add((New-GuiStack @($e, $b2))) }
            elseif ($eta.State -eq 'measuring') { $e = New-GuiText "**Measuring...** $($eta.Basis)" 'Body'; $e.FontSize = 13; $e.Margin = [Windows.Thickness]::new(0, 8, 0, 0); $items.Add($e) }
        }
    }
    if ($run.Waits.Active -and $run.Waits.Active.Reason -eq 'network') { $items.Add((New-GuiMessage -Kind warn -Icon warn -Text 'We cannot reach Microsoft. Trying again every 30 seconds. Check your internet connection.')) }
    $head = New-GuiCard -Content $items.ToArray()
    $head.Margin = [Windows.Thickness]::new(0, 0, 0, 12)
    return $head
}

function New-MbtGuiDecisionPanel {
    $run = $script:Gui.Run
    if (-not $run.Decisions.Count) { return $null }
    $d = $run.Decisions[0]
    $title = New-GuiStack @((New-GuiIcon warn 14 'MbtCaution'), (New-GuiText (Format-MbtGuiDecisionTitle $run.Decisions.Count) 'Title')) -Orientation Horizontal
    $title.Children[0].Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $btns = [Windows.Controls.WrapPanel]::new(); $btns.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
    $btns.Children.Add((New-GuiButton -Text 'It arrived: mark done' -Small -OnClick { Show-MbtGuiDecisionDialog $d 'arrived' }.GetNewClosure())) | Out-Null
    $btns.Children.Add((New-GuiButton -Text 'It did not arrive: send again' -Small -OnClick { Show-MbtGuiDecisionDialog $d 'again' }.GetNewClosure())) | Out-Null
    $btns.Children.Add((New-GuiButton -Text 'Decide later' -Small -OnClick { })) | Out-Null
    $card = New-GuiCard -Content @($title, (New-GuiText $d.Msg 'Body'), $btns, (New-GuiText 'The job goes on with the other items. It finishes as "needs a decision" until this is settled.' 'Hint')) -BorderKey 'MbtCaution'
    $card.Margin = [Windows.Thickness]::new(0, 0, 0, 12)
    return $card
}

function Show-MbtGuiDecisionDialog {
    param($Decision, [ValidateSet('arrived', 'again')][string]$Choice)
    $g = $script:Gui
    if ($Choice -eq 'arrived') {
        Show-MbtGuiDialog -Title 'Mark it as done?' -Body 'You are telling us the item is already in the folder. We will not send it again.' -Buttons @(@{ Text = 'Cancel' }, @{ Text = 'It arrived: mark done'; Accent = $true; Action = { Resolve-MbtGuiDecision $Decision 'Imported' }.GetNewClosure() })
    } else {
        Show-MbtGuiDialog -Title 'Send it again?' -Body 'If the item did arrive after all, the mailbox will hold it twice. Use this only when you looked in the folder and it is not there.' -Buttons @(@{ Text = 'Cancel' }, @{ Text = 'Send it again'; Accent = $true; Action = { Resolve-MbtGuiDecision $Decision 'NotImported' }.GetNewClosure() })
    }
}

function Resolve-MbtGuiDecision {
    <# Settles one uncertain item (engine: -Mode Status -ResolveUnknown <importKey> -As Imported|NotImported). The job keeps going; the row leaves the list. #>
    param($Decision, [string]$As)
    $g = $script:Gui; $run = $g.Run
    $null = $run.Decisions.Remove($Decision)
    if (-not $run.Decisions.Count -and $run.State -eq 'decision') { $run.State = 'running' }
    $run.Dirty = $true
    if ($g.Engine.Name -ne 'Fake' -and $g.Job -and $g.Job.Exited) {
        $args = @('-NoProfile', '-NonInteractive', '-File', $g.EntryScript, '-Mode', 'Status', '-StorePath', $g.Job.StorePath, '-ResolveUnknown', [string]$Decision.Key, '-As', $As, '-UserPrincipalName', $g.S.Mailbox) + @(Get-MbtGuiAuthArguments -S $g.S)
        Start-MbtGuiDetached -Argument $args
    }
    Sync-MbtGuiRunPage
}

function New-MbtGuiStepsColumn {
    $g = $script:Gui; $run = $g.Run; $s = $g.S
    $sel = Get-MbtGuiSelectedStep $run $s
    $stack = [Windows.Controls.StackPanel]::new()
    $g.RunStepControl = $null
    $i = 0
    if (-not $run.Steps.Count -and -not $run.Result) {
        $first = New-GuiStack @((New-GuiIcon play 13 'MbtBrandText'), (New-GuiText $(if ($run.Resumed) { 'Resume: reading what is already done' } else { 'Starting the task...' }) 'Body')) -Orientation Horizontal
        $first.Children[0].Margin = [Windows.Thickness]::new(0, 2, 8, 0)
        $stack.Children.Add($first) | Out-Null
    }
    foreach ($step in $run.Steps) {
        $tone = Get-MbtGuiStepTone $step $run.State
        $btn = [Windows.Controls.Primitives.ToggleButton]::new(); $btn.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtPick')
        $btn.IsChecked = ($i -eq $sel); $btn.Tag = $i; $btn.Margin = [Windows.Thickness]::new(0, 0, 0, 4)
        $grid = [Windows.Controls.Grid]::new()
        foreach ($w in 26, '*', 'Auto') { $cd = [Windows.Controls.ColumnDefinition]::new(); $cd.Width = if ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } elseif ($w -eq 'Auto') { [Windows.GridLength]::Auto } else { [Windows.GridLength]::new($w) }; $grid.ColumnDefinitions.Add($cd) }
        foreach ($r in 1..3) { $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new()) }
        $brush = switch ($tone[0]) { 'ok' { 'MbtSuccess' } 'warn' { 'MbtCaution' } 'bad' { 'MbtCritical' } 'run' { 'MbtBrandText' } default { 'TextFillColorTertiaryBrush' } }
        $control = if ($step.State -eq 'running' -and -not $g.RunStepControl) { Get-MbtGuiRunStepControl $run } else { $null }
        if ($control) { $ic = New-MbtGuiStepControlButton -Kind $control -Brush $brush; $g.RunStepControl = $ic }
        else { $ic = New-GuiIcon $tone[1] 13 $brush; $ic.VerticalAlignment = 'Top'; $ic.Margin = [Windows.Thickness]::new(0, 2, 0, 0) }
        [Windows.Controls.Grid]::SetRowSpan($ic, 2)
        $tt = New-GuiText ("$($i + 1). $($step.Title)") 'Body'; $tt.FontSize = 13; $tt.FontWeight = [Windows.FontWeights]::SemiBold; [Windows.Controls.Grid]::SetColumn($tt, 1)
        $tm = ''
        if ($step.Started) { $end = if ($step.Finished) { $step.Finished } else { [datetime]::UtcNow }; if ($step.State -in 'done', 'failed') { $tm = Format-MbtGuiDuration (($end - $step.Started).TotalSeconds) } }
        $tmb = New-GuiText $tm 'Small' -NoWrap; [Windows.Controls.Grid]::SetColumn($tmb, 2)
        $sub = New-GuiText (Get-MbtGuiStepSub $step) 'Small'; $sub.Margin = [Windows.Thickness]::new(0); [Windows.Controls.Grid]::SetRow($sub, 1); [Windows.Controls.Grid]::SetColumn($sub, 1); [Windows.Controls.Grid]::SetColumnSpan($sub, 2)
        $grid.Children.Add($ic) | Out-Null; $grid.Children.Add($tt) | Out-Null; $grid.Children.Add($tmb) | Out-Null; $grid.Children.Add($sub) | Out-Null
        if ($step.State -eq 'running') {
            $frac = if ($step.Total -gt 0) { $step.Done / $step.Total } else { 0 }
            $pb = New-GuiProgressBar -Fraction $frac -Tone $(if ($run.State -in 'sleep', 'stopped') { 'warn' } else { 'run' }); $pb.Margin = [Windows.Thickness]::new(0, 4, 0, 0)
            [Windows.Controls.Grid]::SetRow($pb, 2); [Windows.Controls.Grid]::SetColumn($pb, 1); [Windows.Controls.Grid]::SetColumnSpan($pb, 2)
            $grid.Children.Add($pb) | Out-Null
        }
        $btn.Content = $grid
        Add-MbtGuiHandler $btn 'Click' { param($sender, $e) $script:Gui.S.RStep = [int]$sender.Tag; $script:Gui.S.RObj = $null; Sync-MbtGuiRunPage -Force }
        $stack.Children.Add($btn) | Out-Null
        $i++
    }
    return (New-GuiCard -Title 'Steps' -Info 'all known before the start' -Content @($stack))
}

function New-MbtGuiObjectRows {
    param($Objects, [int]$Selected)
    $g = $script:Gui
    $stack = [Windows.Controls.StackPanel]::new()
    $i = 0
    foreach ($o in $Objects) {
        $btn = [Windows.Controls.Primitives.ToggleButton]::new(); $btn.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtPick')
        $btn.Padding = [Windows.Thickness]::new(4, 7, 4, 7); $btn.IsChecked = ($i -eq $Selected); $btn.Tag = $i
        $grid = [Windows.Controls.Grid]::new()
        foreach ($w in 24, '*', 'Auto') { $cd = [Windows.Controls.ColumnDefinition]::new(); $cd.Width = if ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } elseif ($w -eq 'Auto') { [Windows.GridLength]::Auto } else { [Windows.GridLength]::new($w) }; $grid.ColumnDefinitions.Add($cd) }
        $gl = switch ($o.Tone) { 'ok' { 'check' } 'warn' { 'warn' } 'bad' { 'x' } 'run' { 'play' } default { 'clock' } }
        $brush = switch ($o.Tone) { 'ok' { 'MbtSuccess' } 'warn' { 'MbtCaution' } 'bad' { 'MbtCritical' } 'run' { 'MbtBrandText' } default { 'TextFillColorTertiaryBrush' } }
        $ic = New-GuiIcon $gl 12 $brush
        $name = New-GuiText $o.Key 'Body' -NoWrap; $name.FontSize = 13; $name.TextTrimming = 'CharacterEllipsis'; [Windows.Controls.Grid]::SetColumn($name, 1)
        $right = if ($o.Msg) { $o.Msg } elseif ($o.Total -gt 0) { Format-MbtGuiFraction $o.Done $o.Total -Finished:($o.Tone -eq 'ok') } else { '' }
        $rt = New-GuiText $right 'Small' -NoWrap; $rt.Margin = [Windows.Thickness]::new(8, 0, 0, 0); [Windows.Controls.Grid]::SetColumn($rt, 2)
        $grid.Children.Add($ic) | Out-Null; $grid.Children.Add($name) | Out-Null; $grid.Children.Add($rt) | Out-Null
        $btn.Content = $grid
        Add-MbtGuiHandler $btn 'Click' { param($sender, $e) $script:Gui.S.RObj = [int]$sender.Tag; Sync-MbtGuiRunPage -Force }
        $stack.Children.Add($btn) | Out-Null
        $i++
    }
    return $stack
}

function New-MbtGuiStepDetailsColumn {
    $g = $script:Gui; $run = $g.Run; $s = $g.S
    $selIdx = Get-MbtGuiSelectedStep $run $s
    $step = if ($run.Steps.Count -gt $selIdx) { $run.Steps[$selIdx] } else { $null }
    $pill = if ($step) { switch ($step.State) { 'done' { New-GuiPill 'Done' 'ok' } 'running' { New-GuiPill 'Running' 'run' } 'failed' { New-GuiPill 'Failed' 'bad' } default { New-GuiPill 'Waiting' '' } } } else { $null }
    $objs = @($run.Objects | Where-Object { -not $step -or $_.Step -eq $step.Id -or -not $_.Step })
    if (-not $objs.Count) { $objs = @($run.Objects) }
    $content = [Collections.Generic.List[object]]::new()
    $help = New-GuiText (Get-MbtGuiStepHelp $s.Task $(if ($step) { $step.Id } else { '' }) $(if ($step) { $step.Title } else { '' })) 'Hint'; $help.Margin = [Windows.Thickness]::new(0, 0, 0, 8)
    $content.Add($help)
    $lbl = New-GuiText ('FOLDERS ' + [char]0x00B7 + ' click one for details') 'Small'; $lbl.Margin = [Windows.Thickness]::new(0, 6, 0, 2)
    $content.Add($lbl)
    $selObj = if ($null -ne $s.RObj) { [int]$s.RObj } else { -1 }
    $content.Add((New-MbtGuiObjectRows -Objects $objs -Selected $selObj))
    if ($s.Task -eq 'import' -or ($s.Task -eq 'mailcopy' -and $run.ImportHalf)) { $content.Add((New-GuiMessage -Kind info -Icon info -Text '**Why is it slow?** Microsoft accepts only a few GiB per hour into one mailbox. The tool waits exactly as long as Microsoft asks. This is normal.')) }
    $logStack = [Windows.Controls.StackPanel]::new()
    foreach ($l in ($run.Log | Select-Object -First 40)) {
        $row = [Windows.Controls.Grid]::new()
        foreach ($w in 48, 20, '*') { $cd = [Windows.Controls.ColumnDefinition]::new(); $cd.Width = if ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } else { [Windows.GridLength]::new($w) }; $row.ColumnDefinitions.Add($cd) }
        $tm = New-GuiText $(if ($l.T) { ([datetime]$l.T).ToLocalTime().ToString('HH:mm') } else { '' }) 'Small' -NoWrap; $tm.Margin = [Windows.Thickness]::new(0); Set-MbtGuiResource $tm ([Windows.Controls.TextBlock]::ForegroundProperty) 'TextFillColorTertiaryBrush'
        $gn = switch ($l.Kind) { 'ok' { 'check' } 'warn' { 'warn' } 'bad' { 'x' } default { 'play' } }
        $gb = switch ($l.Kind) { 'ok' { 'MbtSuccess' } 'warn' { 'MbtCaution' } 'bad' { 'MbtCritical' } default { 'MbtBrandText' } }
        $gi = New-GuiIcon $gn 11 $gb; [Windows.Controls.Grid]::SetColumn($gi, 1)
        $tx = New-GuiText $l.Text 'Body'; $tx.FontSize = 12; [Windows.Controls.Grid]::SetColumn($tx, 2)
        $row.Children.Add($tm) | Out-Null; $row.Children.Add($gi) | Out-Null; $row.Children.Add($tx) | Out-Null
        $row.Margin = [Windows.Thickness]::new(0, 0, 0, 3)
        $logStack.Children.Add($row) | Out-Null
    }
    if (-not $run.Log.Count) { $logStack.Children.Add((New-GuiText 'Nothing yet.' 'Hint')) | Out-Null }
    $more = New-GuiMore -Summary 'Activity log' -Content @($logStack) -Open:([bool]$s.LogOpen)
    Add-MbtGuiHandler $more 'Expanded' { $script:Gui.S.LogOpen = $true }
    Add-MbtGuiHandler $more 'Collapsed' { $script:Gui.S.LogOpen = $false }
    $content.Add($more)
    return (New-GuiCard -Title $(if ($step) { $step.Title } else { 'Details' }) -Right $pill -Content $content.ToArray())
}

function New-MbtGuiObjectDetailsColumn {
    $g = $script:Gui; $run = $g.Run; $s = $g.S
    $items = [Collections.Generic.List[object]]::new()
    $selIdx = Get-MbtGuiSelectedStep $run $s
    $step = if ($run.Steps.Count -gt $selIdx) { $run.Steps[$selIdx] } else { $null }
    $objs = @($run.Objects | Where-Object { -not $step -or $_.Step -eq $step.Id -or -not $_.Step })
    if (-not $objs.Count) { $objs = @($run.Objects) }
    $o = $null
    if ($null -ne $s.RObj -and $s.RObj -lt $objs.Count) { $o = $objs[[int]$s.RObj] }
    elseif ($null -eq $s.RObj) { $o = $objs | Where-Object { $_.Tone -in 'warn', 'bad' } | Select-Object -First 1 }
    if (-not $o) { $items.Add((New-GuiText 'Click a step or a folder to read what it does.' 'Muted')); return (New-GuiCard -Title 'Details' -Content $items.ToArray()) }
    $h = New-GuiText $o.Key 'Body'; $h.FontWeight = [Windows.FontWeights]::SemiBold; $items.Add($h)
    $rel = @($run.Items | Where-Object { $_.Step -eq $o.Step -and ($_.Detail -eq $o.Key -or $_.Key -like "$($o.Key)*") -or $_.Detail -eq $o.Key })
    if ($o.State -in 'notready', 'not-ready') {
        $items.Add((New-GuiMessage -Kind warn -Icon warn -Text '**Not ready.** Microsoft is still moving this folder inside the archive, so it cannot be read yet. It is kept on the list. Press **Continue** later (the same day is usually enough).'))
    } elseif ($o.Tone -eq 'bad' -or $o.State -eq 'refused') {
        $items.Add((New-GuiText $(if ($o.Msg) { $o.Msg } else { 'Microsoft refused it. Nothing was created.' }) 'Hint'))
    } elseif ($rel.Count) {
        $items.Add((New-GuiText "$(Format-MbtGuiQuantity $rel.Count 'item') $(Get-MbtGuiPlural $rel.Count 'was' 'were') sent but Microsoft did not answer. We are looking in the folder to see if $(Get-MbtGuiPlural $rel.Count 'it' 'they') arrived." 'Hint'))
        $rows = @(foreach ($r in $rel) { , @($r.Key, (New-GuiPill $(if ($r.State -eq 'decision') { 'Two matches' } else { 'Checking' }) 'warn')) })
        $items.Add((New-GuiTable -Headers @('Item', 'State') -Rows @($rows) -Widths @('*', 'Auto')))
    } else {
        $stateWord = switch ($o.Tone) { 'ok' { 'Done' } 'run' { 'Running' } 'warn' { 'Needs attention' } default { 'Waiting' } }
        $items.Add((New-GuiFacts @(@('Progress', $(if ($o.Msg) { $o.Msg } else { Format-MbtGuiFraction $o.Done $o.Total -Finished:($o.Tone -eq 'ok') })), @('State', $stateWord))))
        if ($o.Msg -and $o.Tone -eq 'warn') { $items.Add((New-GuiMessage -Kind warn -Icon warn -Text $o.Msg)) }
    }
    return (New-GuiCard -Title 'Details' -Content $items.ToArray())
}

function Build-MbtGuiRunPage {
    $g = $script:Gui; $s = $g.S; $run = $g.Run
    if (-not $run) { $run = New-MbtGuiRun -Task $s.Task; $g.Run = $run }
    $g.RunHeadHolder = [Windows.Controls.ContentControl]::new()
    $g.RunDecisionHolder = [Windows.Controls.ContentControl]::new()
    $g.RunCols = [Windows.Controls.Grid]::new()
    foreach ($w in '264', '12', '*', '12', '300') { $cd = [Windows.Controls.ColumnDefinition]::new(); $cd.Width = if ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } else { [Windows.GridLength]::new([double]$w) }; $g.RunCols.ColumnDefinitions.Add($cd) }
    $g.RunStepsHolder = [Windows.Controls.ContentControl]::new(); $g.RunMidHolder = [Windows.Controls.ContentControl]::new(); $g.RunDetailHolder = [Windows.Controls.ContentControl]::new()
    [Windows.Controls.Grid]::SetColumn($g.RunMidHolder, 2); [Windows.Controls.Grid]::SetColumn($g.RunDetailHolder, 4)
    $g.RunStepsHolder.VerticalAlignment = 'Top'; $g.RunMidHolder.VerticalAlignment = 'Top'; $g.RunDetailHolder.VerticalAlignment = 'Top'
    foreach ($h in $g.RunStepsHolder, $g.RunMidHolder, $g.RunDetailHolder) { $g.RunCols.Children.Add($h) | Out-Null }
    $page = New-GuiStack @($g.RunHeadHolder, $g.RunDecisionHolder, $g.RunCols)
    Sync-MbtGuiRunPage -Force
    return $page
}

function Sync-MbtGuiRunPage {
    <# The run page follows the model: each region is built again (cheap), the page itself is kept, so scroll and open sections stay. #>
    param([switch]$Force)
    $g = $script:Gui
    if ((Get-MbtGuiScreen) -ne 'run' -or -not $g.Run -or -not $g.ContainsKey('RunHeadHolder')) { return }
    $now = [datetime]::UtcNow
    if (-not $Force -and ($now - $g.LastEtaPaint).TotalMilliseconds -lt 400) { return }
    # a press of the mouse button is not interrupted by building the buttons again
    if (-not $Force -and [Windows.Input.Mouse]::LeftButton -eq 'Pressed') { return }
    $g.LastEtaPaint = $now
    $run = $g.Run
    $controlHadFocus = [bool]($g.ContainsKey('RunStepControl') -and $g.RunStepControl -and $g.RunStepControl.IsKeyboardFocused)
    $g.RunHeadHolder.Content = New-MbtGuiRunHead
    $g.RunDecisionHolder.Content = New-MbtGuiDecisionPanel
    $g.RunStepsHolder.Content = New-MbtGuiStepsColumn
    if ($controlHadFocus -and $g.RunStepControl) {
        $control = $g.RunStepControl
        [void]$control.Dispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Input, [Action] { [void]$control.Focus() }.GetNewClosure())
    }
    $g.RunMidHolder.Content = New-MbtGuiStepDetailsColumn
    $g.RunDetailHolder.Content = New-MbtGuiObjectDetailsColumn
    $run.Dirty = $false
    Update-MbtGuiRunActions
    # the taskbar button follows the run
    $info = $g.Window.TaskbarItemInfo
    if ($info) {
        $frac = if ($run.BytesTotal -gt 0) { $run.Bytes / $run.BytesTotal } elseif ($run.Total -gt 0) { $run.Done / $run.Total } else { 0 }
        $info.ProgressState = if ($run.Result) { 'None' } elseif ($run.State -in 'sleep', 'stopped') { 'Paused' } else { 'Normal' }
        $info.ProgressValue = [math]::Max(0.0, [math]::Min(1.0, [double]$frac))
    }
}

function Update-MbtGuiRunClock { if ($script:Gui.Run -and -not $script:Gui.Run.Result) { Sync-MbtGuiRunPage } }

function Update-MbtGuiRunActions {
    <# The action bar of the run: Stop and keep progress at the left while it runs, Continue when it is stopped. #>
    $g = $script:Gui; $run = $g.Run; $s = $g.S
    $active = $run.State -in 'running', 'sleep', 'decision'
    $left = $null
    if ($active -and -not $run.StopRequested) {
        $left = New-GuiButton -Text 'Stop and keep progress' -Danger -Icon stop -OnClick { Show-MbtGuiStopDialog }
    } elseif ($run.StopRequested -and -not $run.Result) {
        $left = New-GuiText 'Stopping... the items in flight are finished first (seconds to a minute).' 'Small'
        $wait = Get-MbtGuiStopWait $g.Job
        if ($wait -ge ($script:Throughput.ForceStopMinutes * 60)) { $left = New-GuiButton -Text 'Force stop' -Danger -OnClick { Stop-MbtGuiProcessTree -Job $g.Job } }
    }
    $footer = if ($run.State -eq 'stopped') { 'Stopped. Nothing is lost.' } elseif ($active -and $s.Awake) { 'Keeping this PC awake. You can leave this window open and do other things.' } else { '' }
    if ($run.State -eq 'stopped') { Set-MbtGuiActions -Back $false -Footer $footer -Left $left -Label 'Continue' -Icon play -OnClick { Invoke-MbtGuiContinue } }
    else { Set-MbtGuiActions -Back $false -Footer $footer -Left $left }
    $g.C.Primary.Visibility = $(if ($run.State -eq 'stopped') { 'Visible' } else { 'Collapsed' })
}

function Show-MbtGuiStopDialog {
    Show-MbtGuiDialog -Title 'Stop this job?' -Body 'The items being sent right now are finished first (a few seconds to a minute). Everything done so far is kept. You can continue later: only what is left will be done.' -Buttons @(
        @{ Text = 'Keep running' }
        @{ Text = 'Stop and keep progress'; Danger = $true; Action = { $g = $script:Gui; Request-MbtGuiStop -Job $g.Job; $g.Run.StopRequested = $true; Sync-MbtGuiRunPage -Force } })
}

function Invoke-MbtGuiContinue {
    <# Continue: the same command again with the same store and job name: it does only what is left (Done items are skipped, items in flight are checked first). #>
    $g = $script:Gui; $s = $g.S
    if (-not $g.Job) { return }
    $job = $g.Job
    if (Test-Path -LiteralPath $job.StopPath) { Remove-Item -LiteralPath $job.StopPath -Force }
    $job.StopAsked = $null; $job.PartExit = $null
    $req = New-MbtGuiRequest -Task $s.Task
    # a job of the command line runs again by its name; the others are built again from the state of the console
    if ($job.ContainsKey('ResumeArguments') -and $job.ResumeArguments) { $job.Queue.Enqueue([pscustomobject]@{ Title = 'Continue'; Arguments = $job.ResumeArguments }) }
    else { foreach ($c in (Get-MbtGuiCommand -Task $s.Task -S $s -Job $job)) { $job.Queue.Enqueue($c) } }
    $g.Run = New-MbtGuiRun -Task $s.Task -Title (Get-MbtGuiRunTitle)
    $g.Run.Resumed = $true
    $s.RStep = $null; $s.RObj = $null
    $null = Start-MbtGuiProcess -Job $job
    Set-MbtGuiStep ($script:Flows[$s.Task].IndexOf('run'))
}

function Complete-MbtGuiRun {
    <# The result arrived: the Done page opens by itself (a run stopped by the user stays on the Run page, with Continue). #>
    $g = $script:Gui; $run = $g.Run; $s = $g.S
    if ($run.State -eq 'stopped') { $run.Dirty = $true; Sync-MbtGuiRunPage -Force; return }
    if ($s.Task -eq 'export' -and $s.Dest -eq 'pst' -and -not $s.KeepStore -and $run.Result.Exit -eq 0 -and $g.Job.StorePath -and (Test-Path -LiteralPath $g.Job.StorePath)) {
        try { Remove-Item -LiteralPath $g.Job.StorePath -Recurse -Force; Add-MbtGuiLog $run 'ok' 'The working store was removed, as you chose.' ([datetime]::UtcNow) } catch { Add-MbtGuiLog $run 'warn' "The working store could not be removed: $($_.Exception.Message)" ([datetime]::UtcNow) }
    }
    $s.Outcome = $run.State
    # report.json and report.csv are small and are written here; report.html of a big job takes about a minute, so the worker writes it while the Done page is usable ("Open the report" says so until it is there)
    try { $g.Reports = @(Write-MbtGuiRunReport -Run $run -Folder $g.Job.Folder -Task $s.Task -RepairDestination $s.RpDest -NoHtml) } catch { Add-MbtGuiLog $run 'warn' "The report could not be written: $($_.Exception.Message)" ([datetime]::UtcNow) }
    $g.ReportState = 'writing'
    $work = $s.WorkFolder
    $reportDone = {
        param($path)
        $g.ReportState = 'ready'
        if ($path) { $g.Reports = @($g.Reports) + [string]$path }
        if ($g.ReportHint) { try { $g.ReportHint.Visibility = 'Collapsed' } catch { Write-Verbose "hint: $_" } }
        # this job is one more row of the overview of all the jobs (written by the worker, so the page does not wait for it)
        try { & $g.Engine.JobOverview $work { param($overview) } { param($message) Write-MbtGuiFault 'job overview' $message } } catch { Write-MbtGuiFault 'job overview' $_ }
    }.GetNewClosure()
    $reportFailed = { param($message) $g.ReportState = 'failed'; Write-MbtGuiFault 'job report' $message; try { $g.ReportHint.Text = "The report could not be written: $message"; $g.ReportHint.Visibility = 'Visible' } catch { Write-Verbose "hint: $_" } }.GetNewClosure()
    try { & $g.Engine.JobReport $g.Job.Folder $reportDone $reportFailed } catch { & $reportFailed $_.Exception.Message }
    Set-MbtGuiStep ($script:Flows[$s.Task].IndexOf('done'))
}

#endregion

#region Done page -----------------------------------------------------------------------------------------------

function Get-MbtGuiCount {
    param($Run, [string]$Name, $Default = 0)
    $c = $Run.Result.Counts
    if ($null -ne $c) { return (Get-MbtGuiProp $c $Name $Default) }
    return $Default
}

function Build-MbtGuiDonePage {
    $g = $script:Gui; $s = $g.S; $run = $g.Run
    if (-not $run) { $run = New-MbtGuiRun -Task $s.Task; $run.State = 'complete'; $run.Result = @{ Exit = 0; Complete = $true; Counts = $null; Next = ''; Reason = '' }; $g.Run = $run }
    $k = $s.Task
    $o = $run.State
    $outcome = switch ($o) { 'complete' { 'complete' } 'decision-final' { 'decision' } 'error' { 'error' } 'refused' { 'refused' } default { 'partial' } }
    $verb = switch ($k) { 'import' { 'imported' } 'repair' { 'repaired' } 'copy' { 'copied' } 'mailcopy' { 'copied' } default { 'exported' } }
    $head = switch ($outcome) {
        'complete' { @{ Pill = 'ok'; Text = "Done. Everything was $verb."; Icon = 'check'; Label = 'Complete' } }
        'refused' { $n = Get-MbtGuiCount $run 'refused' (Get-MbtGuiCount $run 'failed' 0); @{ Pill = 'warn'; Text = "Finished. $(Format-MbtGuiQuantity $n 'item') could not be $verb."; Icon = 'warn'; Label = 'Finished, with refused items' } }
        'decision' { @{ Pill = 'warn'; Text = 'Almost done. One item needs your decision.'; Icon = 'warn'; Label = 'Needs a decision' } }
        'error' { @{ Pill = 'bad'; Text = 'Stopped by an error.'; Icon = 'x'; Label = 'Error' } }
        default { @{ Pill = 'warn'; Text = 'Not finished. Nothing is lost.'; Icon = 'warn'; Label = $(if ($k -eq 'repair') { 'Stopped' } else { 'Resumable' }) } }
    }
    if ($k -eq 'repair' -and $outcome -eq 'complete') {
        $kept = Get-MbtGuiCount $run 'recovered' 0; $part = Get-MbtGuiCount $run 'partial' 0; $lost = Get-MbtGuiCount $run 'lost' 0
        $head.Text = "Repair finished. $(Format-MbtGuiCount ($kept + $part)) of $(Format-MbtGuiQuantity ($kept + $part + $lost) 'message') kept."; $head.Label = 'Finished'
    }
    if ($k -eq 'repair' -and $outcome -eq 'partial') { $head.Text = 'Stopped. No file was written.' }
    $body = switch ($k) { 'repair' { New-MbtGuiRepairDone $outcome } 'import' { New-MbtGuiImportDone $outcome } 'copy' { New-MbtGuiCopyDone $outcome } 'mailcopy' { New-MbtGuiMailCopyDone $outcome } default { New-MbtGuiExportDone $outcome } }
    $title = New-GuiStack @((New-GuiIcon $head.Icon 18 $(switch ($head.Pill) { 'ok' { 'MbtSuccess' } 'bad' { 'MbtCritical' } default { 'MbtCaution' } })), (New-GuiText $head.Text 'Big' -NoWrap)) -Orientation Horizontal
    $title.Children[0].Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    $pill = New-GuiPill $head.Label $head.Pill
    $headRow = [Windows.Controls.DockPanel]::new()
    [Windows.Controls.DockPanel]::SetDock($pill, 'Right'); $pill.Margin = [Windows.Thickness]::new(12, 0, 0, 0); $headRow.Children.Add($pill) | Out-Null
    $headRow.Children.Add($title) | Out-Null
    $card = New-GuiCard -Content (@($headRow) + @($body))
    Set-MbtGuiActions -Back $false -Label 'Start another task' -Secondary -OnClick { $g = $script:Gui; $g.S.Task = $null; $g.S.Sel = $null; $g.S.Step = 1; $g.Run = $null; $g.Job = $null; Show-MbtGuiPage }
    return (New-GuiStack @($card, (New-MbtGuiNextCard -Task $k -Outcome $outcome), (New-MbtGuiCommandLineCard -Task $k -Job $g.Job), (New-MbtGuiReportsCard)))
}
function Start-MbtGuiOpenOf {
    <# An export store goes to the Open task: its source is already chosen and read. #>
    param([Parameter(Mandatory)][string]$Path)
    $g = $script:Gui; $s = $g.S
    $s.OpSrc = 'store'; $s.OpPath = $Path; $s.OpInfo = $null
    $s.Task = 'open'; $s.Sel = 'open'
    $s.Step = $script:Flows.open.IndexOf('opFrom')
    Initialize-MbtGuiTask
    Open-MbtGuiSource -Use open
}

function Start-MbtGuiImportOf {
    <# What a task made (a PST, an export store) goes to the Import task: its source is already chosen and checked. #>
    param([Parameter(Mandatory)][ValidateSet('pst', 'store')][string]$Kind, [Parameter(Mandatory)][string]$Path)
    $g = $script:Gui; $s = $g.S
    $s.ImSrc = $Kind; $s.ImPath = $Path; $s.ImInfo = $null
    $s.Task = 'import'; $s.Sel = 'import'
    $s.Step = $script:Flows.import.IndexOf('imFrom')
    Initialize-MbtGuiTask
    Open-MbtGuiSource -Use import
}

function Get-MbtGuiNextSteps {
    <# What to do next, in one sentence, and the buttons to the logical next task. Task: export, import, repair, copy. Outcome: complete, partial, decision, error, refused. #>
    param([string]$Task, [string]$Outcome)
    $g = $script:Gui; $s = $g.S
    $howTo = @{ Text = 'How to open the PST in Outlook'; Icon = 'open'; Action = { Open-MbtGuiHelp -Target Guide } }
    $again = @{ Text = 'Continue'; Icon = 'play'; Accent = $true; Action = { Invoke-MbtGuiContinue } }
    $retry = @{ Text = 'Try again'; Icon = 'refresh'; Accent = $true; Action = { Invoke-MbtGuiContinue } }
    $pstPath = switch ($Task) { 'repair' { $s.RpDest } default { [IO.Path]::Combine($s.OutFolder, $s.OutName) } }
    $importPst = @{ Text = 'Import into a mailbox'; Icon = 'up'; Accent = $true; Action = { Start-MbtGuiImportOf -Kind pst -Path $pstPath }.GetNewClosure() }
    switch ("$Task/$Outcome") {
        'export/complete' {
            if ($s.Dest -eq 'pst') { return @{ Text = 'The export is finished and every PST file was checked. Most people now put it into a mailbox, or open the PST in Outlook.'; Buttons = @($importPst, $howTo) } }
            $store = (New-MbtGuiRequest -Task export).StorePath
            return @{ Text = 'The export store is finished and checked. You can put it into a mailbox now.'; Buttons = @(@{ Text = 'Import into a mailbox'; Icon = 'up'; Accent = $true; Action = { Start-MbtGuiImportOf -Kind store -Path $store }.GetNewClosure() }) }
        }
        'export/partial' { return @{ Text = 'Press Continue (later, for a folder that is not ready): only what is left is done.'; Buttons = @($again) } }
        'import/complete' { return @{ Text = 'The items are in the mailbox. Open Outlook and look in the new folder; nothing that was there before was changed.'; Buttons = @(@{ Text = 'Import another file'; Icon = 'up'; Action = { $g = $script:Gui; $g.S.ImPath = ''; $g.S.ImInfo = $null; $g.S.Task = 'import'; $g.S.Sel = 'import'; $g.S.Step = $script:Flows.import.IndexOf('imFrom'); Initialize-MbtGuiTask; Show-MbtGuiPage } }) } }
        'import/partial' { return @{ Text = 'Press Continue: only what is left is done.'; Buttons = @($again) } }
        'import/decision' { return @{ Text = 'Decide for each uncertain item above. When every one is decided, the job is finished.'; Buttons = @() } }
        'import/refused' { return @{ Text = 'The job is finished. The refused items cannot be imported as they are: the list above says why, item by item.'; Buttons = @() } }
        'repair/complete' { return @{ Text = 'The repaired file is a new PST; your original was not touched. Open it in Outlook, or put what it holds into a mailbox.'; Buttons = @($importPst, $howTo) } }
        'repair/partial' { return @{ Text = 'Start the repair again whenever you want.'; Buttons = @($again) } }
        'mailcopy/complete' {
            $ends = Get-MbtGuiMailCopyEnds
            $store = (New-MbtGuiRequest -Task mailcopy).StorePath
            $open = @{ Text = 'Open the store'; Icon = 'open'; Action = { Start-MbtGuiOpenOf -Path $store }.GetNewClosure() }
            $lookIn = if ($ends.Both) { "the folders **$(if ($s.TgtMode -eq 'new') { $s.TgtName } else { 'of the same names' })** (Primary part) and **$(if ($s.TgtMode2 -eq 'new') { $s.TgtName2 } else { 'of the same names' })** (Archive part)" } elseif ($s.TgtMode -eq 'new') { "the folder **$($s.TgtName)**" } else { 'the folders of the same names' }
            return @{ Text = "The items are in the mailbox $($ends.To). Open Outlook, or Outlook on the web, and look in ${lookIn}: nothing that was there before was changed. The store of the copy stays on this PC as proof and as a resume point; you can look inside it, or delete its folder when you no longer need it."; Buttons = @($open) }
        }
        'mailcopy/partial' { return @{ Text = 'Press Continue: a part that is finished is not done again, and no item is copied twice.'; Buttons = @($again) } }
        'mailcopy/decision' { return @{ Text = 'Decide for each uncertain item above. When every one is decided, the job is finished.'; Buttons = @() } }
        'mailcopy/refused' { return @{ Text = 'The job is finished. The refused items cannot be added as they are: the list above says why, item by item.'; Buttons = @() } }
        'copy/complete' { return @{ Text = 'The new PST is ready; your original was only read. Open it in Outlook, or put what it holds into a mailbox.'; Buttons = @($importPst, $howTo) } }
        'copy/refused' { return @{ Text = 'The new PST is ready except for the items listed above.'; Buttons = @($howTo) } }
        'copy/partial' { return @{ Text = 'Start the copy again whenever you want.'; Buttons = @() } }
        default {
            if ($Outcome -eq 'error') { return @{ Text = 'Fix what the message above says, then try again. Nothing that was done is lost.'; Buttons = @($retry) } }
            return @{ Text = 'Nothing else is needed.'; Buttons = @() }
        }
    }
}

function New-MbtGuiNextCard {
    param([string]$Task, [string]$Outcome)
    $info = Get-MbtGuiNextSteps -Task $Task -Outcome $Outcome
    $bar = [Windows.Controls.WrapPanel]::new(); $bar.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    foreach ($b in @($info.Buttons)) { $bar.Children.Add((New-GuiButton -Text $b.Text -Icon $b.Icon -Accent:([bool](Get-MbtGuiProp $b 'Accent')) -OnClick $b.Action)) | Out-Null }
    return (New-GuiCard -Title 'What to do next' -Content $(if (@($info.Buttons).Count) { @((New-GuiText $info.Text 'Body'), $bar) } else { @((New-GuiText $info.Text 'Body')) }))
}

function New-MbtGuiReportsCard {
    <# The report, the log and the folder of the job. In the demo no file exists: the buttons say so. #>
    $g = $script:Gui; $s = $g.S
    $open = {
        param($path)
        if ($path -and (Test-Path -LiteralPath $path)) { Start-Process -FilePath $path }
        else { Show-MbtGuiDialog -Title 'Nothing to open here' -Body 'No file was written: this is the demonstration.' -Buttons @(@{ Text = 'Close'; Accent = $true }) }
    }
    $folder = if ($g.Job) { $g.Job.Folder } else { $null }
    $logPath = if ($folder) { [IO.Path]::Combine($folder, 'events.jsonl') } else { $null }
    # the report is looked for when the button is pressed: it may have been written since the page was drawn
    $openReport = {
        if ($g.ReportState -eq 'writing') { Show-MbtGuiDialog -Title 'The report is being written' -Body 'The report of a big job takes about a minute. Press the button again in a moment: the line under the buttons goes away when it is ready.' -Buttons @(@{ Text = 'Close'; Accent = $true }); return }
        & $open (Get-MbtGuiReportPath -Reports @($g.Reports) -Folder $folder)
    }.GetNewClosure()
    $dir = switch ($s.Task) { 'repair' { Split-Path $s.RpDest -Parent } { $_ -in 'import', 'mailcopy' } { $folder } default { $s.OutFolder } }
    $bar = [Windows.Controls.WrapPanel]::new()
    $bar.Children.Add((New-GuiButton -Text 'Open the report' -Small -Icon doc -OnClick { & $openReport }.GetNewClosure())) | Out-Null
    $bar.Children.Add((New-GuiButton -Text 'Open the log' -Small -Icon clip -OnClick { & $open $logPath }.GetNewClosure())) | Out-Null
    $bar.Children.Add((New-GuiButton -Text 'Open the folder' -Small -Icon open -OnClick { & $open $dir }.GetNewClosure())) | Out-Null
    $bar.Children.Add((New-GuiButton -Text 'All jobs' -Small -Icon doc -OnClick { Open-MbtGuiJobOverview })) | Out-Null
    $hint = New-GuiText 'The report is being written: it takes about a minute for a big job. You can keep working.' 'Small'
    $hint.Visibility = $(if ($g.ReportState -eq 'writing') { 'Visible' } else { 'Collapsed' })
    $g.ReportHint = $hint
    return (New-GuiCard -Title 'Reports and log' -Content @($bar, $hint))
}

function New-MbtGuiTiles {
    param([object[]]$Tiles)
    $p = [Windows.Controls.WrapPanel]::new(); $p.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
    foreach ($t in $Tiles) { $p.Children.Add((New-GuiTile $t[0] $t[1] $t[2])) | Out-Null }
    return $p
}

function New-MbtGuiErrorBlock {
    param($Run, [string]$Fallback)
    $g = $script:Gui
    $raw = if ($Run.Error) { $Run.Error.Msg + $(if ($Run.Error.Tech) { "`n" + $Run.Error.Tech } else { '' }) } else { $Fallback }
    $m = ConvertTo-MbtGuiMessage $(if ($Run.Error) { "$($Run.Error.Code) $($Run.Error.Msg) $($Run.Error.Tech)" } else { $Fallback })
    $text = if ($Run.Error -and $Run.Error.Msg) { $Run.Error.Msg } else { $m.Text }
    if ($Run.Result -and $Run.Result.Reason -eq 'signin-required') { $text = 'Your sign-in has expired. Sign in again to continue. Nothing is lost.' }
    $tech = if ($Run.Error -and $Run.Error.Tech) { New-GuiMore -Summary 'Technical message' -Content @((New-GuiText $Run.Error.Tech 'Small')) } else { $null }
    return @((New-GuiMessage -Kind bad -Icon x -Text $text -After @($tech)))
}

function New-MbtGuiPeriodNote {
    <# What the period did, from the result of the job: the items kept, the ones outside it (skipped, not an error), the ones without a date that were kept. $null for an export without a period. #>
    param($Run)
    $s = $script:Gui.S
    $outside = [long](Get-MbtGuiCount $Run 'outsidePeriod' -1)
    $period = Get-MbtGuiPeriod $s
    if ($outside -lt 0 -and -not $period.Active) { return $null }
    $kept = [long](Get-MbtGuiCount $Run 'kept' (Get-MbtGuiCount $Run 'done' 0))
    $undated = [long](Get-MbtGuiCount $Run 'withoutDate' 0)
    $text = "**Period: items $($period.Text).** $(Format-MbtGuiQuantity $kept 'item') kept"
    if ($outside -ge 0) { $text += ", $(Format-MbtGuiCount $outside) outside the period (skipped, not an error)" }
    if ($undated -gt 0) { $text += ". $(Format-MbtGuiQuantity $undated 'item') had no date and $(Get-MbtGuiPlural $undated 'was' 'were') kept" }
    $date = [string](Get-MbtGuiCount $Run 'periodDate' '')
    $dateNote = [string](Get-MbtGuiCount $Run 'periodNote' '')
    if ($dateNote) { $text += ". Date used: $dateNote" }
    $kind = $(if ($date -and $date -ne 'received') { 'warn' } else { 'info' })
    return (New-GuiMessage -Kind $kind -Icon $(if ($kind -eq 'warn') { 'warn' } else { 'info' }) -Text ($text + '.'))
}

function New-MbtGuiExportDone {
    param([string]$Outcome)
    $g = $script:Gui; $s = $g.S; $run = $g.Run
    $periodNote = New-MbtGuiPeriodNote $run
    $done = Get-MbtGuiCount $run 'done' (Get-MbtGuiCount $run 'exported' $run.Done)
    $failed = Get-MbtGuiCount $run 'failed' 0; $pending = Get-MbtGuiCount $run 'pending' 0
    $pst = $s.Dest -eq 'pst'
    $parts = Get-MbtGuiCount $run 'parts' 0
    switch ($Outcome) {
        'complete' {
            $real = [double](Get-MbtGuiCount $run 'bytes' 0); if ($real -le 0) { $real = [double]$run.Bytes }; if ($real -le 0) { $real = [double]$run.BytesTotal }
            $size = Format-MbtGuiBytes $real
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' exported'), 'ok'), @($(if ($pst -and $parts) { "$size in $(Format-MbtGuiQuantity $parts 'PST file')" } else { $size }), $(if ($pst) { 'written' } else { 'in the store' }), 'ok'), @((Format-MbtGuiCount $failed), 'failed', $(if ($failed) { 'bad' } else { 'ok' })))
            $facts = New-GuiFacts @(@('Saved to', $s.OutFolder), @('Files', $(if ($pst) { "$($s.OutName) and its parts + manifest" } else { 'the export store' })), @('Checked', $(if ($pst -and -not $s.CheckEach) { 'not checked (you turned it off)' } else { 'every file checked: clean' })))
            return @($t, $periodNote, $facts)
        }
        'partial' {
            $notReady = @($run.Objects | Where-Object { $_.State -in 'notready', 'not-ready' })
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' exported'), 'ok'), @((Format-MbtGuiCount $pending), 'waiting', 'warn'), @((Format-MbtGuiCount $failed), 'failed', $(if ($failed) { 'bad' } else { 'ok' })))
            $msg = if ($notReady.Count) { "**$($notReady[0].Key)** is not ready at Microsoft yet. It was not exported as empty, and it stays on the list." } else { "**$(Format-MbtGuiStillToDo $pending)**" }
            $extra = @()
            if ($run.Result -and $run.Result.Reason -eq 'signin-required') { $extra = New-MbtGuiErrorBlock $run '' }
            return @($t, $periodNote, (New-GuiMessage -Kind warn -Icon warn -Text $msg)) + @($extra)
        }
        'decision' { return @((New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' exported'), 'ok'), @('1', 'needs a decision', 'warn'), @((Format-MbtGuiCount $failed), 'failed', 'ok'))), (New-GuiMessage -Kind warn -Icon warn -Text 'One item could not be read twice. See the list and choose.')) }
        default {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' exported'), 'ok'), @('0', 'waiting', ''), @('1', 'error', 'bad'))
            return @($t) + @(New-MbtGuiErrorBlock $run 'The export stopped.')
        }
    }
}

function New-MbtGuiImportDone {
    param([string]$Outcome)
    $g = $script:Gui; $s = $g.S; $run = $g.Run
    $done = Get-MbtGuiCount $run 'done' $run.Done; $pending = Get-MbtGuiCount $run 'pending' 0; $failed = Get-MbtGuiCount $run 'failed' 0; $unknown = Get-MbtGuiCount $run 'unknown' 0; $folders = Get-MbtGuiCount $run 'foldersCreated' 0
    switch ($Outcome) {
        'complete' {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @((Format-MbtGuiCount $folders), ((Get-MbtGuiPlural $folders 'folder') + ' created'), 'ok'), @((Format-MbtGuiCount $failed), 'refused', 'ok'), @((Format-MbtGuiCount $unknown), 'uncertain', 'ok'))
            $facts = New-GuiFacts @(@('Mailbox', "$($s.Mailbox) $([char]0x00B7) $($s.TgtLoc)$(if ($s.TgtMode -eq 'new') { " $([char]0x00B7) folder **$($s.TgtName)**" })"), @('Final check', 'every item of the job is done'))
            return @($t, $facts)
        }
        'refused' {
            $refused = Get-MbtGuiCount $run 'refused' $failed
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @((Format-MbtGuiCount $folders), ((Get-MbtGuiPlural $folders 'folder') + ' created'), 'ok'), @((Format-MbtGuiCount $refused), 'refused for good', 'bad'), @((Format-MbtGuiCount $unknown), 'uncertain', 'ok'))
            return @($t) + @(New-MbtGuiRefusedBlock -Verb 'imported')
        }
        'partial' {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @((Format-MbtGuiCount $pending), 'still to do', 'warn'), @((Format-MbtGuiCount $failed), 'refused', $(if ($failed) { 'bad' } else { 'ok' })), @((Format-MbtGuiCount $unknown), 'uncertain', $(if ($unknown) { 'warn' } else { 'ok' })))
            $hours = ''
            if ($pending -gt 0 -and $run.Total -gt 0) { $left = $run.BytesTotal - $run.Bytes; if ($left -gt 0) { $e = Get-MbtGuiPlanEta -Task import -Bytes $left; $span = ($e.Text -split " \(")[0] -replace "^about ", ""; $hours = " About $span of work remain at Microsoft's speed." } }
            $extra = @(); if ($run.Result -and $run.Result.Reason -eq 'signin-required') { $extra = New-MbtGuiErrorBlock $run '' }
            return @($t, (New-GuiMessage -Kind warn -Icon warn -Text "**$(Format-MbtGuiStillToDo $pending)** $(if ($done -eq 1) { 'The item already done is skipped.' } else { 'The items already done are skipped.' })$hours")) + @($extra)
        }
        'decision' {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @('0', 'still to do', 'ok'), @((Format-MbtGuiCount $failed), 'refused', 'ok'), @((Format-MbtGuiCount ([math]::Max(1, $unknown))), 'needs your decision', 'warn'))
            $list = @($run.Decisions | ForEach-Object {
                    $d = $_
                    $btns = [Windows.Controls.WrapPanel]::new(); $btns.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
                    $btns.Children.Add((New-GuiButton -Text 'It arrived' -Small -OnClick { Show-MbtGuiDecisionDialog $d 'arrived' }.GetNewClosure())) | Out-Null
                    $btns.Children.Add((New-GuiButton -Text 'Send it again' -Small -OnClick { Show-MbtGuiDecisionDialog $d 'again' }.GetNewClosure())) | Out-Null
                    New-GuiMessage -Kind warn -Icon warn -Text "**$($d.Title)**: $($d.Msg) Is it already in the mailbox?" -After @($btns)
                })
            $bulk = New-GuiButton -Text 'Send again every item that was not found' -Small -OnClick { Show-MbtGuiBulkDialog }
            $bulk.Margin = [Windows.Thickness]::new(0, 10, 0, 0); $bulk.HorizontalAlignment = 'Left'
            return @($t) + $list + @($bulk)
        }
        default {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @((Format-MbtGuiCount $pending), 'still to do', ''), @((Format-MbtGuiCount $failed), 'refused', $(if ($failed) { 'bad' } else { '' })), @((Format-MbtGuiCount $unknown), 'uncertain', ''))
            return @($t) + @(New-MbtGuiErrorBlock $run 'The import stopped.')
        }
    }
}

function New-MbtGuiRefusedBlock {
    <# The items refused for good, in plain words: why they are not retried, the list with the reason of each one, the CSV with all of them. #>
    param([string]$Verb = 'imported')
    $g = $script:Gui; $run = $g.Run
    $parts = [Collections.Generic.List[object]]::new()
    $n = Get-MbtGuiCount $run 'refused' 0
    $parts.Add((New-GuiMessage -Kind warn -Icon warn -Text ("**$(Format-MbtGuiQuantity $n 'item') could not be $Verb, and trying again would not help.** $(if ($n -eq 1) { 'It was' } else { 'Each one was' }) refused by Microsoft or could not be read from the file; the job tried as many times as allowed. " +
                    "Everything else is done: the job is finished, so **Continue** has nothing left to do. The list below says why$(if ($n -eq 1) { '.' } else { ', item by item.' })")))
    $items = @($run.Items | Where-Object { $_.State -eq 'refused' })
    if ($items.Count) {
        $rows = @(foreach ($i in ($items | Select-Object -First 25)) { , @((New-GuiStack @((New-GuiText $(if ($i.Key) { $i.Key } else { '(no subject)' }) 'Body'), (New-GuiText $i.Detail 'Small'))), (ConvertTo-MbtGuiRefusalReason $i.Msg)) })
        $parts.Add((New-GuiMore -Summary (Format-MbtGuiRefusedSummary $rows.Count $n) -Open -Content @((New-GuiTable -Headers @('Item', 'Why it was refused') -Rows @($rows) -Widths @('*', '*')))))
    }
    $csv = [string](Get-MbtGuiCount $run 'report' '')
    if ($csv) {
        $open = New-GuiButton -Text 'Open the list of refused items (CSV)' -Small -Icon open -OnClick { if (Test-Path -LiteralPath $csv) { Start-Process -FilePath $csv } }.GetNewClosure()
        $open.Margin = [Windows.Thickness]::new(0, 8, 0, 0); $open.HorizontalAlignment = 'Left'
        $parts.Add($open)
        $parts.Add((New-GuiText "One line per item with the folder, the reason and the technical message: $csv" 'Hint'))
    }
    return $parts.ToArray()
}

function New-MbtGuiCopyDone {
    param([string]$Outcome)
    $g = $script:Gui; $s = $g.S; $run = $g.Run
    $done = Get-MbtGuiCount $run 'itemsWritten' $run.Done; $failed = Get-MbtGuiCount $run 'failed' 0; $parts = Get-MbtGuiCount $run 'parts' 0; $bytes = Get-MbtGuiCount $run 'bytes' 0
    switch ($Outcome) {
        { $_ -in 'complete', 'refused' } {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'message') + ' copied'), 'ok'), @($(if ($parts) { "$(Format-MbtGuiBytes $bytes) in $(Format-MbtGuiQuantity $parts 'PST file')" } else { Format-MbtGuiBytes $bytes }), 'written', 'ok'), @((Format-MbtGuiCount $failed), 'not copied', $(if ($failed) { 'bad' } else { 'ok' })))
            $facts = New-GuiFacts @(@('Source', "$($s.OpPath) $([char]0x00B7) **not changed**"), @('Saved to', "$($s.OutFolder) ($($s.OutName) and its parts)"), @('Checked', $(if ($s.CheckEach) { 'every file checked: clean' } else { 'not checked (you turned it off)' })))
            $out = @($t, $facts)
            if ($Outcome -eq 'refused') { $out += @(New-MbtGuiRefusedBlock -Verb 'copied') }
            return $out
        }
        'partial' { return @((New-GuiMessage -Kind warn -Icon warn -Text 'A stopped copy leaves no file, and your original was never touched.')) }
        default { return @((New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'message') + ' copied'), 'ok'), @('1', 'error', 'bad'))) + @(New-MbtGuiErrorBlock $run 'The copy stopped. No file was kept.')) }
    }
}

function New-MbtGuiMailCopyDone {
    <# The end of a copy to another mailbox: what was read, what was added, what is left; both parts in one card. #>
    param([string]$Outcome)
    $g = $script:Gui; $s = $g.S; $run = $g.Run
    $mid = [char]0x00B7
    $read = Get-MbtGuiCount $run 'exported' 0; $done = Get-MbtGuiCount $run 'done' $run.Done; $pending = Get-MbtGuiCount $run 'pending' 0; $failed = Get-MbtGuiCount $run 'failed' 0; $folders = Get-MbtGuiCount $run 'foldersCreated' 0
    $ends = Get-MbtGuiMailCopyEnds
    $periodNote = New-MbtGuiPeriodNote $run
    $readTile = @((Format-MbtGuiCount $read), ((Get-MbtGuiPlural $read 'item') + ' read'), 'ok')
    switch ($Outcome) {
        'complete' {
            $t = New-MbtGuiTiles @($readTile, @((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @((Format-MbtGuiCount $folders), ((Get-MbtGuiPlural $folders 'folder') + ' created'), 'ok'), @((Format-MbtGuiCount $failed), 'refused', 'ok'))
            $into = if ($ends.Both) { "$($ends.To) $mid Primary part: $(if ($s.TgtMode -eq 'new') { "folder **$($s.TgtName)**" } else { 'same folder names' }) $mid Archive part: $(if ($s.TgtMode2 -eq 'new') { "folder **$($s.TgtName2)**" } else { 'same folder names' }) ($($ends.ToKind))" } else { "$($ends.To) $mid $($ends.ToKind)$(if ($s.TgtMode -eq 'new') { " $mid folder **$($s.TgtName)**" })" }
            $facts = New-GuiFacts @(@('From', "$($ends.From) $mid $($ends.FromKind) $mid **not changed**"), @('Into', $into), @('Store', "$((New-MbtGuiRequest -Task mailcopy).StorePath) (kept)"), @('Final check', 'every item of the job is done'))
            return @($t, $periodNote, $facts)
        }
        'refused' {
            $refused = Get-MbtGuiCount $run 'refused' $failed
            $t = New-MbtGuiTiles @($readTile, @((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @((Format-MbtGuiCount $folders), ((Get-MbtGuiPlural $folders 'folder') + ' created'), 'ok'), @((Format-MbtGuiCount $refused), 'refused for good', 'bad'))
            return @($t) + @(New-MbtGuiRefusedBlock -Verb 'copied')
        }
        'partial' {
            $t = New-MbtGuiTiles @($readTile, @((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @((Format-MbtGuiCount $pending), 'still to do', 'warn'), @((Format-MbtGuiCount $failed), 'failed', $(if ($failed) { 'bad' } else { 'ok' })))
            $extra = @(); if ($run.Result -and $run.Result.Reason -eq 'signin-required') { $extra = New-MbtGuiErrorBlock $run '' }
            $left = if ($run.ImportHalf) { "**$(Format-MbtGuiStillToDo $pending)** The part that was read is kept in the store." } else { '**The reading is not finished.** What was read is kept in the store; the part that adds the items to the mailbox follows it.' }
            return @($t, $periodNote, (New-GuiMessage -Kind warn -Icon warn -Text "$left Press **Continue**: a part that is finished is not done again, and no item is added twice.")) + @($extra)
        }
        'decision' { return (New-MbtGuiImportDone 'decision') }
        default {
            $t = New-MbtGuiTiles @($readTile, @((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' in the mailbox'), 'ok'), @('0', 'still to do', ''), @('1', 'error', 'bad'))
            return @($t) + @(New-MbtGuiErrorBlock $run 'The copy stopped.')
        }
    }
}

function Show-MbtGuiBulkDialog {
    Show-MbtGuiDialog -Title 'Send again every item that was not found?' -Body 'For each uncertain item that is not in the folder, the item is sent again. An item for which the folder holds a candidate stays as it is. If an item did arrive after all, the mailbox will hold it twice.' -Buttons @(
        @{ Text = 'Cancel' }
        @{ Text = 'Send again'; Accent = $true; Action = {
                $g = $script:Gui
                if ($g.Engine.Name -ne 'Fake') { Start-MbtGuiDetached -Argument (@('-NoProfile', '-NonInteractive', '-File', $g.EntryScript, '-Mode', 'Status', '-StorePath', $g.Job.StorePath, '-Unknown', '-Check', '-ResolveAbsent', '-JobName', $g.Job.JobName, '-UserPrincipalName', $g.S.Mailbox) + @(Get-MbtGuiAuthArguments -S $g.S)) }
            } })
}

function New-MbtGuiRepairDone {
    param([string]$Outcome)
    $g = $script:Gui; $s = $g.S; $run = $g.Run
    switch ($Outcome) {
        'partial' { return @((New-GuiMessage -Kind warn -Icon warn -Text 'The new file is written under a temporary name and moved only when complete, so no half-written file is left. Your original was never touched.')) }
        'error' { return @(New-MbtGuiErrorBlock $run 'The repair stopped.') }
    }
    $c = $run.Result.Counts
    $rec = Get-MbtGuiCount $run 'recovered' 0; $part = Get-MbtGuiCount $run 'partial' 0; $lost = Get-MbtGuiCount $run 'lost' 0
    $total = [math]::Max(1, $rec + $part + $lost)
    $pct = ('{0}%' -f ([math]::Round(100 * ($rec + $part) / $total, 1)).ToString([cultureinfo]::InvariantCulture))
    $tiles = New-MbtGuiTiles @(@((Format-MbtGuiCount $rec), 'recovered as they were', 'ok'), @((Format-MbtGuiCount $part), 'partly recovered', 'warn'), @((Format-MbtGuiCount $lost), 'lost', $(if ($lost) { 'bad' } else { 'ok' })), @($pct, 'kept (recovered + partial)', ''))
    $bef = Get-MbtGuiCount $run 'before' $null; $aft = Get-MbtGuiCount $run 'after' $null
    $be = [int](Get-MbtGuiProp $bef 'errors' 0); $bw = [int](Get-MbtGuiProp $bef 'warnings' 0); $ae = [int](Get-MbtGuiProp $aft 'errors' 0); $aw = [int](Get-MbtGuiProp $aft 'warnings' 0)
    $mkpills = { param($e, $w) $p = [Windows.Controls.StackPanel]::new(); $p.Orientation = 'Horizontal'; $a = New-GuiPill (Format-MbtGuiQuantity $e 'error') $(if ($e) { 'bad' } else { 'ok' }); $a.Margin = [Windows.Thickness]::new(0, 0, 8, 0); $b = New-GuiPill (Format-MbtGuiQuantity $w 'warning') $(if ($w) { 'warn' } else { 'ok' }); $p.Children.Add($a) | Out-Null; $p.Children.Add($b) | Out-Null; $p }
    $before = New-GuiStack @((New-GuiText 'BEFORE (original, untouched)' 'Small'), (& $mkpills $be $bw)); $after = New-GuiStack @((New-GuiText 'AFTER (new file)' 'Small'), (& $mkpills $ae $aw))
    $ba = New-GuiGrid -Children @($before, $after); $ba.Margin = [Windows.Thickness]::new(0, 12, 0, 0)
    $level = [string](Get-MbtGuiProp $c 'levelUsed' $s.RpLevel)
    $destFile = [string](Get-MbtGuiProp $c 'destination' $s.RpDest)
    $facts = New-GuiFacts @(@('New file', $destFile), @('Level used', "$level$(if ($s.RpLevel -eq 'Automatic') { ' (chosen automatically)' })"), @('Named properties', [string](Get-MbtGuiProp $c 'nameMap' 'recovered from the file (calendar and contact fields kept)')))
    $parts = @($tiles, $ba, $facts)
    $incomplete = @($run.Items | Where-Object { $_.State -in 'partial', 'lost' })
    if ($incomplete.Count) {
        $rows = @(foreach ($w in $incomplete) { $name = New-GuiStack @((New-GuiText $(if ($w.Key) { $w.Key } else { '(no subject)' }) 'Body'), (New-GuiText $w.Detail 'Small')); , @($name, (New-GuiPill $(if ($w.State -eq 'lost') { 'Lost' } else { 'Partial' }) $(if ($w.State -eq 'lost') { 'bad' } else { 'warn' })), $w.Msg) })
        $parts += New-GuiMore -Summary (Format-MbtGuiIncompleteSummary $incomplete.Count) -Content @((New-GuiTable -Headers @('Message', 'Result', 'What was lost') -Rows @($rows) -Widths @('*', '90', '*')))
    }
    $did = @($run.Steps | Where-Object { $_.State -eq 'done' } | ForEach-Object { $x = New-GuiText ([char]0x2022 + '  ' + $_.Title + $(if ($_.Msg) { ': ' + $_.Msg } else { '' })) 'Body'; $x.FontSize = 13; $x })
    if ($did.Count) { $parts += New-GuiMore -Summary 'What the repair did' -Content (@($did) + @(New-GuiText 'Counters, sizes and indexes are recomputed; messages no folder listed are put back in their folders.' 'Hint')) }
    $parts += New-GuiText "The report is saved next to the new file: $([IO.Path]::GetFileNameWithoutExtension($s.RpDest)).repair-report.csv and .json." 'Hint'
    return $parts
}

#endregion

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA21trLV37bvkGY
# B77QvRjw6oSKBgvGoz4zcv6IKig1JaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYEDCCGAwCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCn+5gEYd8O
# hiDSguVeojz4AngiRRdySbPCpedjSh9MlzANBgkqhkiG9w0BAQEFAASCAgCJIliR
# L9yJX9PGAPg+x0oiv1eBaQzfIv0VkyuSesMyhcvNYb87KfOAeNzLhMi4qT0PmTPy
# 0L+P3KdzcqYJeMmVLwmAA1ctH2EZJKQrMqilZIvciHQllVvgpFow9RbThAzWM9Uw
# plK84ttL1xjK+w9BqZhFjEDTMIvLsCkHFbaEtd3eR2zsY2jINLMChYVcNcPLEnmk
# 8mJA3HnLHhdXsiBNRmOTrGhvS1zlKXgUcTjGy0NKpgz9LCtuia8a6Qv7mQKsmv0y
# ZnuWQuRTzeT5kkI4V2QWPp/kMIZQdAYB0Ki9JUQn8qNrPyaAWIH0+vNNShZHrMcA
# Fupf+FQ+3Bj4OqpJe72LsAyGTSPNv0OIjIJUrvtmacmMRDX6rVa99nMu0GQQ+IC3
# zufwrYcTAILKwtslccWujsy0ExkFR3jHb9EWXH/5RTs3pzTcbMC4bcF9HTzpglyd
# Fig9R2H6KDNMuF3B2qVGs5dxslXWDs0PAnGD1g1YF/ILwX9HcLwIKOLw9yb2CJgp
# F+AFUZAys9RK7LS/R06br28xMcJjEThZ0WbSfwcB6QlkIZW+jrf0COt7aWI/76Qn
# DenaY1tx9AanRAvGHKG9dWuB016ovmh9tkRsa58TVT0iOI7Ecq9a1PVmaqCzfR1z
# Do2WtRCclKb2u4HtvlQMB8ohmWnbGvvi7HvmpaGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIMF2wdtm3tAyjJLZmPNwdQXeo1jzS8l2uRz3GUr1/pTvAghWfNYqmqia
# FhgPMjAyNjExMjcwNTI4MDlaMAMCAQECEFVqQO4dY+utCuzEgnVRzo6gghGnMIIE
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
# wzSPOzGCAlUwggJRAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTEyNzA1MjgwOVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIFDl
# m/ADbuDfojJcbQWlLUZBuKABS41/Zy3+4VBVonBCMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiBN3gahGCs9lTcvodRYGG4gu7sy9Uf3TQk+E1o2FA2ByAIgN/LP
# RvLwWqKmwOf0irOhLSJbRbvE3KNmj6mn/qnbG0Y=
# SIG # End signature block
