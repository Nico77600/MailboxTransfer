#Requires -Version 7.4
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
    [void]$sb.Append("$($s.Task)|$($s.Dest)|$($s.OutFolder)|$($s.OutName)|$($s.SplitGB)|$($s.CheckEach)|$($s.SepAux)|$($s.KeepStore)|$($s.Conc)|$($s.Retry)|$($s.Tries)|$($s.Reconcile)|$($s.ImSrc)|$($s.ImPath)|$($s.TgtLoc)|$($s.TgtMode)|$($s.TgtName)|$($s.Mapping)|$($s.Shared)|$($s.ImShared)|$($s.RecoverableItems)|$($s.RpPath)|$($s.RpDest)|$($s.RpLevel)|$($s.RpKeep)|$($s.OpPath)|$($s.Mailbox)|$($s.AuthMode)|$($s.TenantId)|$($s.ClientId)|$($s.AppTenant)|$($s.AppClient)|$($s.AppThumb)|$($s.Awake)".ToLowerInvariant())
    $sel = Get-MbtGuiTaskSelection
    if ($sel) { foreach ($e in $sel.Entries) { [void]$sb.Append("|$($e.Kind):$($e.Node.Id):$($e.Items)") } }
    $bytes = [Text.Encoding]::UTF8.GetBytes($sb.ToString())
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}

function Get-MbtGuiSelectionKey {
    <# What was chosen for the operation (the source, the mailbox, the ticks): the selection file and the folder files are named after it, so the same choice is the same file. #>
    $g = $script:Gui; $s = $g.S
    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append("$($s.Task)|$($s.Mailbox)|$($s.Shared)|$($s.ImSrc)|$($s.ImPath)|$($s.OpPath)|$($s.RecoverableItems)".ToLowerInvariant())
    $sel = Get-MbtGuiTaskSelection
    if ($sel) { foreach ($e in $sel.Entries) { [void]$sb.Append("|$($e.Kind):$($e.Node.Id):$($e.Items)") } }
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
    $name = if ($s.Task -eq 'import') { 'The plan is valid' } else { 'Every folder matched' }
    switch ($s.PlanState) {
        'checking' {
            $g = $script:Gui
            $run = if ($g.ContainsKey('PlanRun')) { $g.PlanRun } else { $null }
            $started = if ($g.ContainsKey('PlanStarted')) { $g.PlanStarted } else { $null }
            return (New-MbtGuiCheckRow -State wait -Text 'Checking the plan with the service' -Detail (Get-MbtGuiPlanProgressText -Run $run -Started $started))
        }
        'ok' {
            $counted = if ($s.PlanTotals) { "counted by the service: $(Format-MbtGuiPlanTotals $s.PlanTotals)" + $(if ($s.PlanTotals.FoldersToCreate -gt 0) { "; $(Format-MbtGuiQuantity $s.PlanTotals.FoldersToCreate 'folder') will be created" } else { '' }) } else { '' }
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
                'imPlan' { return $(if ($s.Confirmed) { 'Ready. Nothing is written before you press Start.' } else { 'Tick the box to unlock Start.' }) }
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
    if ((Get-MbtGuiScreen) -notin 'exPlan', 'imPlan', 'cpPlan') { return }
    if ($g.PlanHolder) { $g.PlanHolder.Content = New-MbtGuiPlanStateMessage }
    $ok = ($State -eq 'ok')
    if ((Get-MbtGuiScreen) -eq 'imPlan') { $ok = $ok -and $s.Confirmed }
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
        'export' { [IO.Path]::Combine($storeRoot, ($s.Mailbox -replace '[^\w\.\-@]', '_')) }
        'import' { if ($s.ImSrc -eq 'store') { $s.ImPath } else { [IO.Path]::Combine($storeRoot, 'import') } }
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
    if ($hasSel -and ($Task -eq 'export' -or ($Task -eq 'import' -and $s.ImSrc -eq 'store'))) { $files = @("$base.json") }
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
    <# The selection of the export (one selection per location, in one file). Calls OnDone with the path in an array, as the earlier one-file-per-location version did. #>
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
    if ($sel -and $sel.Entries.Count -and $s.Task -eq 'export') { Save-MbtGuiSelectionFile -Selection $sel -Path "$base.json" -OnDone $begin -OnError $fail }
    else { & $begin @() }
}

function Step-MbtGuiPlanJob {
    $g = $script:Gui
    if (-not $g.ContainsKey('PlanJob') -or -not $g.PlanJob) { return }
    $null = Step-MbtGuiJob -Job $g.PlanJob -Run $g.PlanRun
    # while the check runs the line says what it does and for how long (it is drawn again about once a second)
    if (-not $g.PlanRun.Result -and $g.S.PlanState -eq 'checking' -and $g.PlanHolder -and ((Get-MbtGuiScreen) -in 'exPlan', 'imPlan', 'cpPlan')) {
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
    $needsSel = $s.Task -eq 'export' -or ($s.Task -eq 'import' -and $s.ImSrc -eq 'store' -and $sel -and $sel.Entries.Count)
    if ($needsSel) { Save-MbtGuiSelectionFile -Selection $sel -Path "$base.json" -OnDone $begin -OnError { param($m) Set-MbtGuiPlanState error ((ConvertTo-MbtGuiMessage $m).Text); Show-MbtGuiPage } }
    else { & $begin @() }
}

function Write-MbtGuiRequestFile {
    <# request.json: what was asked, no secrets; the Continue list and a resumed job read it. #>
    param([Parameter(Mandatory)][hashtable]$Job)
    $s = $script:Gui.S
    $keys = 'Task', 'OpSrc', 'OpPath', 'Mailbox', 'Shared', 'ImShared', 'RecoverableItems', 'Dest', 'KeepStore', 'SplitGB', 'CheckEach', 'SepAux', 'OutFolder', 'OutName', 'Conc', 'Retry', 'Tries', 'Reconcile', 'Awake', 'ImSrc', 'ImPath', 'TgtLoc', 'TgtMode', 'TgtName', 'Mapping', 'RpPath', 'RpLevel', 'RpDest', 'RpKeep'
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
        default { return $script:Tasks[$s.Task].Title }
    }
}

function Resume-MbtGuiJob {
    <# Continue: the same request again, in the same job folder, so only what is left is done. #>
    param($Entry)
    $g = $script:Gui; $s = $g.S
    if ($g.Engine.Name -eq 'Fake' -and $Entry.PSObject.Properties['Demo'] -and $Entry.Demo) { & $Entry.Demo; return }
    $reqFile = Join-Path $Entry.Folder 'request.json'
    if (-not (Test-Path -LiteralPath $reqFile)) { return }
    $req = Get-Content -LiteralPath $reqFile -Raw | ConvertFrom-Json
    foreach ($p in $req.PSObject.Properties) { if ($s.ContainsKey($p.Name) -and $p.Name -notin 'Task') { $s[$p.Name] = $p.Value } }
    $s.Task = [string]$req.Task; $s.Sel = $s.Task
    $job = New-MbtGuiJob -WorkFolder $s.WorkFolder -Task $s.Task -EntryScript $g.EntryScript -StorePath ([string]$req.StorePath) -SelectionFiles @($req.SelectionFiles) -JobName ([string]$req.JobName)
    $job.Folder = $Entry.Folder
    foreach ($k in 'EventsPath', 'StopPath', 'StatusPath', 'StdOut', 'StdErr') { $job[$k] = Join-Path $Entry.Folder (Split-Path $job[$k] -Leaf) }
    if (Test-Path -LiteralPath $job.StopPath) { Remove-Item -LiteralPath $job.StopPath -Force }
    # the events of the earlier runs stay in the file (with their result): the new run reads only what the new process writes
    $job.Offset = $(if (Test-Path -LiteralPath $job.EventsPath) { (Get-Item -LiteralPath $job.EventsPath).Length } else { 0L })
    foreach ($c in @($req.Commands)) { $job.Queue.Enqueue([pscustomobject]@{ Title = 'Continue'; Arguments = @($c) }) }
    $g.Job = $job; $g.Run = New-MbtGuiRun -Task $s.Task -Title ([string]$req.Title)
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
        '^tree|Read the folder tree' { return 'Reads the folders of the mailbox, 4 requests at a time. A mailbox with thousands of folders takes a minute or two. Nothing is changed.' }
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
            $eta = Get-MbtGuiEta -Run $run -Task $s.Task
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
    $i = 0
    if (-not $run.Steps.Count -and -not $run.Result) {
        $first = New-GuiStack @((New-GuiIcon play 13 'MbtBrandText'), (New-GuiText 'Starting the task...' 'Body')) -Orientation Horizontal
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
        $ic = New-GuiIcon $tone[1] 13 $brush; $ic.VerticalAlignment = 'Top'; $ic.Margin = [Windows.Thickness]::new(0, 2, 0, 0)
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
    if ($s.Task -eq 'import') { $content.Add((New-GuiMessage -Kind info -Icon info -Text '**Why is it slow?** Microsoft accepts only a few GiB per hour into one mailbox. The tool waits exactly as long as Microsoft asks. This is normal.')) }
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
        $items.Add((New-GuiMessage -Kind warn -Icon warn -Text '**Not ready.** Microsoft is still moving this folder into the auxiliary archive, so it cannot be read yet. It is kept on the list. Press **Continue** later (the same day is usually enough).'))
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
    $g.LastEtaPaint = $now
    $run = $g.Run
    $g.RunHeadHolder.Content = New-MbtGuiRunHead
    $g.RunDecisionHolder.Content = New-MbtGuiDecisionPanel
    $g.RunStepsHolder.Content = New-MbtGuiStepsColumn
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
    foreach ($c in (Get-MbtGuiCommand -Task $s.Task -S $s -Job $job)) { $job.Queue.Enqueue($c) }
    $g.Run = New-MbtGuiRun -Task $s.Task -Title (Get-MbtGuiRunTitle)
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
    $verb = switch ($k) { 'import' { 'imported' } 'repair' { 'repaired' } 'copy' { 'copied' } default { 'exported' } }
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
    $body = switch ($k) { 'repair' { New-MbtGuiRepairDone $outcome } 'import' { New-MbtGuiImportDone $outcome } 'copy' { New-MbtGuiCopyDone $outcome } default { New-MbtGuiExportDone $outcome } }
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
    $dir = switch ($s.Task) { 'repair' { Split-Path $s.RpDest -Parent } 'import' { $folder } default { $s.OutFolder } }
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

function New-MbtGuiExportDone {
    param([string]$Outcome)
    $g = $script:Gui; $s = $g.S; $run = $g.Run
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
            return @($t, $facts)
        }
        'partial' {
            $notReady = @($run.Objects | Where-Object { $_.State -in 'notready', 'not-ready' })
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), ((Get-MbtGuiPlural $done 'item') + ' exported'), 'ok'), @((Format-MbtGuiCount $pending), 'waiting', 'warn'), @((Format-MbtGuiCount $failed), 'failed', $(if ($failed) { 'bad' } else { 'ok' })))
            $msg = if ($notReady.Count) { "**$($notReady[0].Key)** is not ready at Microsoft yet. It was not exported as empty, and it stays on the list." } else { "**$(Format-MbtGuiStillToDo $pending)**" }
            $extra = @()
            if ($run.Result -and $run.Result.Reason -eq 'signin-required') { $extra = New-MbtGuiErrorBlock $run '' }
            return @($t, (New-GuiMessage -Kind warn -Icon warn -Text $msg)) + @($extra)
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

