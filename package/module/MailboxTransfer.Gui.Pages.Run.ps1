#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - starting a job, the Run page and the Done page (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    Run: three columns (steps, what the step works on, details) under a header card with the status, the overall bar, the counters and the honest
    ETA. Everything on it comes from the run model (MailboxTransfer.Gui.Run.ps1), which the events of the job feed. Done: one card, one headline,
    one pill, tiles, facts and the next step, by the exit code of the engine (CONSOLE-SPEC 3.6).

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
    [void]$sb.Append("$($s.Task)|$($s.Dest)|$($s.OutFolder)|$($s.OutName)|$($s.SplitGB)|$($s.CheckEach)|$($s.SepAux)|$($s.KeepStore)|$($s.Conc)|$($s.Retry)|$($s.Tries)|$($s.Reconcile)|$($s.ImSrc)|$($s.ImPath)|$($s.TgtLoc)|$($s.TgtMode)|$($s.TgtName)|$($s.Mapping)|$($s.Shared)|$($s.ImShared)|$($s.RecoverableItems)|$($s.RpPath)|$($s.RpDest)|$($s.RpLevel)|$($s.RpKeep)|$($s.OpPath)")
    $sel = Get-MbtGuiTaskSelection
    if ($sel) { foreach ($e in $sel.Entries) { [void]$sb.Append("|$($e.Kind):$($e.Node.Id):$($e.Items)") } }
    $bytes = [Text.Encoding]::UTF8.GetBytes($sb.ToString())
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
}

function New-MbtGuiPlanStateMessage {
    $s = $script:Gui.S
    switch ($s.PlanState) {
        'checking' { return (New-GuiMessage -Kind info -Icon clock -Text 'Checking the plan with the service. Nothing is changed.') }
        'ok' {
            $counted = if ($s.PlanTotals) { " **Counted by the service: $(Format-MbtGuiPlanTotals $s.PlanTotals).**" + $(if ($s.PlanTotals.FoldersToCreate -gt 0) { " $(Format-MbtGuiCount $s.PlanTotals.FoldersToCreate) folders will be created." } else { '' }) } else { '' }
            return (New-GuiMessage -Kind ok -Icon check -Text ('**The plan is valid.** Every folder was matched.' + $counted + ' Nothing starts before you press Start.'))
        }
        'error' { return (New-GuiMessage -Kind bad -Icon x -Text ($s.PlanMsg)) }
        default { return $null }
    }
}

function Set-MbtGuiPlanState {
    param([ValidateSet('idle', 'checking', 'ok', 'error')][string]$State, [string]$Message = '')
    $g = $script:Gui; $s = $g.S
    $s.PlanState = $State; $s.PlanMsg = $Message
    if ((Get-MbtGuiScreen) -notin 'exPlan', 'imPlan', 'cpPlan') { return }
    if ($g.PlanHolder) { $g.PlanHolder.Content = New-MbtGuiPlanStateMessage }
    $ok = ($State -eq 'ok')
    if ((Get-MbtGuiScreen) -eq 'imPlan') { $ok = $ok -and ($s.Confirm.Trim().ToLowerInvariant() -eq 'import') }
    $g.C.Primary.IsEnabled = $ok
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
    [pscustomobject]@{ Task = $Task; StorePath = $store; Config = (Get-MbtGuiEffectiveConfig -State $s -WorkFolder $s.WorkFolder) }
}

function Save-MbtGuiSelectionFiles {
    <# One selection file per location (the engine exports one mailbox per selection). Calls OnDone with the file paths. #>
    param([Parameter(Mandatory)]$Selection, [Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][scriptblock]$OnDone, [scriptblock]$OnError)
    $g = $script:Gui; $s = $g.S
    $groups = @($Selection.Entries | Group-Object { (Get-MbtGuiRootNode $_.Node).Id })
    if (-not $groups.Count) { & $OnDone @(); return }
    $paths = [Collections.Generic.List[string]]::new()
    $pending = @{ N = $groups.Count; Failed = $false }
    $i = 0
    foreach ($grp in $groups) {
        $i++
        $path = Join-Path $Folder ("selection-{0}.json" -f $i)
        $root = Get-MbtGuiRootNode $grp.Group[0].Node
        & $g.Engine.SaveSelection $root @($grp.Group) $path {
            param($p)
            $paths.Add([string]$p); $pending.N--
            if ($pending.N -eq 0 -and -not $pending.Failed) { & $OnDone $paths.ToArray() }
        }.GetNewClosure() {
            param($msg)
            if (-not $pending.Failed) { $pending.Failed = $true; if ($OnError) { & $OnError $msg } }
        }.GetNewClosure()
    }
}

function Start-MbtGuiPlanCheck {
    <# Plan page: the read-only run of the engine (-WhatIf) of exactly this request. Start unlocks only when it is valid. #>
    $g = $script:Gui; $s = $g.S
    $s.PlanTotals = $null
    Set-MbtGuiPlanState checking
    $req = New-MbtGuiRequest -Task $s.Task
    $fail = { param($msg) $m = ConvertTo-MbtGuiMessage $msg; Set-MbtGuiPlanState error ("**The plan cannot be used.** $($m.Text)") }
    if ($g.Engine.Name -eq 'Fake') {
        & $g.Engine.Plan $req { param($r) if ($r.Ok) { if ($r.PSObject.Properties['Totals']) { $s.PlanTotals = $r.Totals }; Set-MbtGuiPlanState ok } else { & $fail $r.Message } }.GetNewClosure() $fail
        return
    }
    $sel = Get-MbtGuiTaskSelection
    $job = New-MbtGuiJob -WorkFolder $s.WorkFolder -Task "$($s.Task)-plan" -EntryScript $g.EntryScript -StorePath $req.StorePath -ConfigText $req.Config
    $begin = {
        param($files)
        $job.SelectionFiles = @($files)
        Write-MbtGuiChoiceFiles -Job $job -Selection $sel
        foreach ($c in (Get-MbtGuiCommand -Task $s.Task -S $s -Job $job -Plan)) { $job.Queue.Enqueue($c) }
        $g.PlanRun = New-MbtGuiRun -Task $s.Task; $g.PlanJob = $job
        $null = Start-MbtGuiProcess -Job $job
    }.GetNewClosure()
    if ($sel -and $sel.Entries.Count -and $s.Task -eq 'export') { Save-MbtGuiSelectionFiles -Selection $sel -Folder $job.Folder -OnDone $begin -OnError $fail }
    else { & $begin @() }
}

function Step-MbtGuiPlanJob {
    $g = $script:Gui
    if (-not $g.ContainsKey('PlanJob') -or -not $g.PlanJob) { return }
    $null = Step-MbtGuiJob -Job $g.PlanJob -Run $g.PlanRun
    if ($g.PlanRun.Result) {
        $r = $g.PlanRun.Result; $g.PlanJob = $null
        if ($r.Exit -eq 0) { $g.S.PlanTotals = Get-MbtGuiPlanTotals $g.PlanRun; Set-MbtGuiPlanState ok }
        else {
            $msg = if ($g.PlanRun.Error) { (ConvertTo-MbtGuiMessage $g.PlanRun.Error.Msg).Text } elseif ($g.PlanRun.Warnings.Count) { $g.PlanRun.Warnings[0].Msg } else { 'The service did not accept the request.' }
            Set-MbtGuiPlanState error "**The plan cannot be used.** $msg"
        }
    }
}

function Start-MbtGuiRun {
    <# Start: the job folder, the command(s), the child process; the Run page opens. #>
    $g = $script:Gui; $s = $g.S
    $req = New-MbtGuiRequest -Task $s.Task
    $job = New-MbtGuiJob -WorkFolder $s.WorkFolder -Task $s.Task -EntryScript $g.EntryScript -StorePath $req.StorePath -ConfigText $req.Config
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
    if ($needsSel) { Save-MbtGuiSelectionFiles -Selection $sel -Folder $job.Folder -OnDone $begin -OnError { param($m) Set-MbtGuiPlanState error ((ConvertTo-MbtGuiMessage $m).Text) } }
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
    $job.ConfigPath = $(if (Test-Path -LiteralPath (Join-Path $Entry.Folder 'effective.config.psd1')) { Join-Path $Entry.Folder 'effective.config.psd1' } else { $null })
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
    if ($Step.Total -gt 0 -and $Step.State -eq 'running') { return ('{0} / {1} items' -f (Format-MbtGuiCount $Step.Done), (Format-MbtGuiCount $Step.Total)) }
    if ($Step.Msg) { return $Step.Msg }
    if ($Step.Total -gt 0) { return ('{0} items' -f (Format-MbtGuiCount $Step.Total)) }
    return ''
}

function Get-MbtGuiStepHelp {
    param([string]$Task, [string]$Id, [string]$Title)
    switch -Regex ("$Id $Title") {
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
    & $kpi ("**$(Format-MbtGuiCount $run.Done)** of $(Format-MbtGuiCount $run.Total) $($run.Unit)")
    if ($run.BytesTotal -gt 0) { & $kpi ("**$(Format-MbtGuiBytes $run.Bytes)** of $(Format-MbtGuiBytes $run.BytesTotal)") }
    $elapsed = if ($run.Started) { Format-MbtGuiDuration (([datetime]::UtcNow - $run.Started).TotalSeconds) } else { '0 s' }
    if ($run.LastEvent -and $run.Started -and $run.State -in 'stopped') { $elapsed = Format-MbtGuiDuration (($run.LastEvent - $run.Started).TotalSeconds) }
    & $kpi ("Elapsed **$elapsed**")
    if ($run.Waits.Count -gt 0) { & $kpi ("Microsoft slowed us down **$($run.Waits.Count) times** ($(Format-MbtGuiDuration $run.Waits.Seconds -WithSeconds) waited)") }
    $items = [Collections.Generic.List[object]]::new()
    foreach ($x in @($dock, $bar, $kp)) { $items.Add($x) }
    $bar.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    if ($run.Waits.Active -and $run.Waits.Active.Reason -eq 'throttle') {
        $items.Add((New-GuiMessage -Kind info -Icon info -Text "Microsoft asks us to slow down. Waiting $([math]::Round($run.Waits.Active.Seconds)) seconds. This is normal."))
    }
    switch ($st) {
        'sleep' {
            $from = if ($run.Sleep.From) { $run.Sleep.From.ToLocalTime().ToString('HH:mm') } else { '?' }; $to = if ($run.Sleep.To) { $run.Sleep.To.ToLocalTime().ToString('HH:mm') } else { '?' }
            $items.Add((New-GuiMessage -Kind warn -Icon moon -Text "**This PC slept from $from to $to.** The job paused and has now started again. The time left is being measured again; it shows after about 10 minutes."))
        }
        'stopped' {
            $items.Add((New-GuiMessage -Kind info -Icon pause -Text "**Stopped by you.** $(Format-MbtGuiCount $run.Done) of $(Format-MbtGuiCount $run.Total) $($run.Unit) are done. Press **Continue** to do only what is left."))
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
    $title = New-GuiStack @((New-GuiIcon warn 14 'MbtCaution'), (New-GuiText "$($run.Decisions.Count) item$(if ($run.Decisions.Count -gt 1) { 's' }) need$(if ($run.Decisions.Count -eq 1) { 's' }) your decision" 'Title')) -Orientation Horizontal
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
        $btn.Add_Click({ param($b, $e) $g.S.RStep = [int]$b.Tag; $g.S.RObj = $null; Sync-MbtGuiRunPage -Force })
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
        $right = if ($o.Msg) { $o.Msg } elseif ($o.Total -gt 0) { '{0} / {1}' -f (Format-MbtGuiCount $o.Done), (Format-MbtGuiCount $o.Total) } else { '' }
        $rt = New-GuiText $right 'Small' -NoWrap; $rt.Margin = [Windows.Thickness]::new(8, 0, 0, 0); [Windows.Controls.Grid]::SetColumn($rt, 2)
        $grid.Children.Add($ic) | Out-Null; $grid.Children.Add($name) | Out-Null; $grid.Children.Add($rt) | Out-Null
        $btn.Content = $grid
        $btn.Add_Click({ param($b, $e) $g.S.RObj = [int]$b.Tag; Sync-MbtGuiRunPage -Force })
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
    $more.Add_Expanded({ $g.S.LogOpen = $true }); $more.Add_Collapsed({ $g.S.LogOpen = $false })
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
        $items.Add((New-GuiText "$($rel.Count) item$(if ($rel.Count -gt 1) { 's were' } else { ' was' }) sent but Microsoft did not answer. We are looking in the folder to see if they arrived." 'Hint'))
        $rows = foreach ($r in $rel) { , @($r.Key, (New-GuiPill $(if ($r.State -eq 'decision') { 'Two matches' } else { 'Checking' }) 'warn')) }
        $items.Add((New-GuiTable -Headers @('Item', 'State') -Rows @($rows) -Widths @('*', 'Auto')))
    } else {
        $stateWord = switch ($o.Tone) { 'ok' { 'Done' } 'run' { 'Running' } 'warn' { 'Needs attention' } default { 'Waiting' } }
        $items.Add((New-GuiFacts @(@('Progress', $(if ($o.Msg) { $o.Msg } else { '{0} / {1}' -f (Format-MbtGuiCount $o.Done), (Format-MbtGuiCount $o.Total) })), @('State', $stateWord))))
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
    try { $g.Reports = @(Write-MbtGuiRunReport -Run $run -Folder $g.Job.Folder -Task $s.Task -RepairDestination $s.RpDest) } catch { Add-MbtGuiLog $run 'warn' "The report could not be written: $($_.Exception.Message)" ([datetime]::UtcNow) }
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
        'refused' { $n = Get-MbtGuiCount $run 'refused' (Get-MbtGuiCount $run 'failed' 0); @{ Pill = 'warn'; Text = "Finished. $(Format-MbtGuiCount $n) item$(if ($n -ne 1) { 's' }) could not be $verb."; Icon = 'warn'; Label = 'Finished, with refused items' } }
        'decision' { @{ Pill = 'warn'; Text = 'Almost done. One item needs your decision.'; Icon = 'warn'; Label = 'Needs a decision' } }
        'error' { @{ Pill = 'bad'; Text = 'Stopped by an error.'; Icon = 'x'; Label = 'Error' } }
        default { @{ Pill = 'warn'; Text = 'Not finished. Nothing is lost.'; Icon = 'warn'; Label = $(if ($k -eq 'repair') { 'Stopped' } else { 'Resumable' }) } }
    }
    if ($k -eq 'repair' -and $outcome -eq 'complete') {
        $kept = Get-MbtGuiCount $run 'recovered' 0; $part = Get-MbtGuiCount $run 'partial' 0; $lost = Get-MbtGuiCount $run 'lost' 0
        $head.Text = "Repair finished. $(Format-MbtGuiCount ($kept + $part)) of $(Format-MbtGuiCount ($kept + $part + $lost)) messages kept."; $head.Label = 'Finished'
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
    $open = New-GuiButton -Text 'Open the folder' -Icon open -OnClick {
        $dir = switch ($s.Task) { 'repair' { Split-Path $s.RpDest -Parent } 'import' { $g.Job.Folder } default { $s.OutFolder } }
        if ($dir -and (Test-Path -LiteralPath $dir)) { Start-Process -FilePath $dir }
    }
    $extra = if ($outcome -in 'complete', 'partial', 'refused') { $open } else { $null }
    switch ($outcome) {
        { $_ -in 'complete', 'refused' } { Set-MbtGuiActions -Back $false -Extra $extra -Label 'Start a new task' -OnClick { $g.S.Task = $null; $g.S.Sel = $null; $g.S.Step = 0; $g.Run = $null; $g.Job = $null; Show-MbtGuiPage } }
        'error' { Set-MbtGuiActions -Back $false -Label 'Try again' -Icon refresh -OnClick { Invoke-MbtGuiContinue } }
        default { Set-MbtGuiActions -Back $false -Extra $extra -Label 'Continue' -Icon play -OnClick { Invoke-MbtGuiContinue } }
    }
    return $card
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
            $size = if ($run.BytesTotal -gt 0) { Format-MbtGuiBytes $run.BytesTotal } else { Format-MbtGuiBytes (Get-MbtGuiCount $run 'bytes' 0) }
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items exported', 'ok'), @($(if ($pst -and $parts) { "$size in $parts PST files" } else { $size }), $(if ($pst) { 'written' } else { 'in the store' }), 'ok'), @((Format-MbtGuiCount $failed), 'failed', $(if ($failed) { 'bad' } else { 'ok' })))
            $facts = New-GuiFacts @(@('Saved to', $s.OutFolder), @('Files', $(if ($pst) { "$($s.OutName) and its parts + manifest" } else { 'the export store' })), @('Checked', $(if ($pst -and -not $s.CheckEach) { 'not checked (you turned it off)' } else { 'every file checked: clean' })))
            return @($t, $facts, (New-GuiMessage -Kind info -Icon info -Text $(if ($pst) { '**What next?** Open a PST in Outlook, or choose **Import** to put it into a mailbox.' } else { '**What next?** Choose **Import** to put the store into a mailbox, or write PST files from it.' })))
        }
        'partial' {
            $notReady = @($run.Objects | Where-Object { $_.State -in 'notready', 'not-ready' })
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items exported', 'ok'), @((Format-MbtGuiCount $pending), 'waiting', 'warn'), @((Format-MbtGuiCount $failed), 'failed', $(if ($failed) { 'bad' } else { 'ok' })))
            $msg = if ($notReady.Count) { "**$($notReady[0].Key)** is not ready at Microsoft yet. It was not exported as empty. Press **Continue** later: only these folders will be tried again." } else { 'The export is not complete. Nothing is lost. Press **Continue**: only what is left will be done.' }
            $extra = @()
            if ($run.Result -and $run.Result.Reason -eq 'signin-required') { $extra = New-MbtGuiErrorBlock $run '' }
            return @($t, (New-GuiMessage -Kind warn -Icon warn -Text $msg)) + @($extra)
        }
        'decision' { return @((New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items exported', 'ok'), @('1', 'needs a decision', 'warn'), @((Format-MbtGuiCount $failed), 'failed', 'ok'))), (New-GuiMessage -Kind warn -Icon warn -Text 'One item could not be read twice. See the list and choose.')) }
        default {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items exported', 'ok'), @('0', 'waiting', ''), @('1', 'error', 'bad'))
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
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items in the mailbox', 'ok'), @((Format-MbtGuiCount $folders), 'folders created', 'ok'), @((Format-MbtGuiCount $failed), 'refused', 'ok'), @((Format-MbtGuiCount $unknown), 'uncertain', 'ok'))
            $facts = New-GuiFacts @(@('Mailbox', "$($s.Mailbox) $([char]0x00B7) $($s.TgtLoc)$(if ($s.TgtMode -eq 'new') { " $([char]0x00B7) folder **$($s.TgtName)**" })"), @('Final check', 'every item of the job is done'))
            return @($t, $facts, (New-GuiMessage -Kind info -Icon info -Text 'Open Outlook and look in the new folder. Nothing existing was changed.'))
        }
        'refused' {
            $refused = Get-MbtGuiCount $run 'refused' $failed
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items in the mailbox', 'ok'), @((Format-MbtGuiCount $folders), 'folders created', 'ok'), @((Format-MbtGuiCount $refused), 'refused for good', 'bad'), @((Format-MbtGuiCount $unknown), 'uncertain', 'ok'))
            return @($t) + @(New-MbtGuiRefusedBlock -Verb 'imported')
        }
        'partial' {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items in the mailbox', 'ok'), @((Format-MbtGuiCount $pending), 'still to do', 'warn'), @((Format-MbtGuiCount $failed), 'refused', $(if ($failed) { 'bad' } else { 'ok' })), @((Format-MbtGuiCount $unknown), 'uncertain', $(if ($unknown) { 'warn' } else { 'ok' })))
            $hours = ''
            if ($pending -gt 0 -and $run.Total -gt 0) { $left = $run.BytesTotal - $run.Bytes; if ($left -gt 0) { $e = Get-MbtGuiPlanEta -Task import -Bytes $left; $span = ($e.Text -split " \(")[0] -replace "^about ", ""; $hours = " About $span of work remain at Microsoft's speed." } }
            $extra = @(); if ($run.Result -and $run.Result.Reason -eq 'signin-required') { $extra = New-MbtGuiErrorBlock $run '' }
            return @($t, (New-GuiMessage -Kind warn -Icon warn -Text "**Not finished, nothing lost.** Press **Continue**: the items already done are skipped.$hours")) + @($extra)
        }
        'decision' {
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items in the mailbox', 'ok'), @('0', 'still to do', 'ok'), @((Format-MbtGuiCount $failed), 'refused', 'ok'), @((Format-MbtGuiCount ([math]::Max(1, $unknown))), 'needs your decision', 'warn'))
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
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'items in the mailbox', 'ok'), @((Format-MbtGuiCount $pending), 'still to do', ''), @((Format-MbtGuiCount $failed), 'refused', $(if ($failed) { 'bad' } else { '' })), @((Format-MbtGuiCount $unknown), 'uncertain', ''))
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
    $parts.Add((New-GuiMessage -Kind warn -Icon warn -Text ("**$(Format-MbtGuiCount $n) item$(if ($n -ne 1) { 's' }) could not be $Verb, and trying again would not help.** Each one was refused by Microsoft or could not be read from the file; the job tried as many times as allowed. " +
                    "Everything else is done: the job is finished, so **Continue** has nothing left to do. The list below says why, item by item.")))
    $items = @($run.Items | Where-Object { $_.State -eq 'refused' })
    if ($items.Count) {
        $rows = foreach ($i in ($items | Select-Object -First 25)) { , @((New-GuiStack @((New-GuiText $(if ($i.Key) { $i.Key } else { '(no subject)' }) 'Body'), (New-GuiText $i.Detail 'Small'))), (ConvertTo-MbtGuiRefusalReason $i.Msg)) }
        $parts.Add((New-GuiMore -Summary "The first $($rows.Count) of $(Format-MbtGuiCount $n) refused items" -Open -Content @((New-GuiTable -Headers @('Item', 'Why it was refused') -Rows @($rows) -Widths @('*', '*')))))
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
            $t = New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'messages copied', 'ok'), @($(if ($parts) { "$(Format-MbtGuiBytes $bytes) in $parts PST file$(if ($parts -ne 1) { 's' })" } else { Format-MbtGuiBytes $bytes }), 'written', 'ok'), @((Format-MbtGuiCount $failed), 'not copied', $(if ($failed) { 'bad' } else { 'ok' })))
            $facts = New-GuiFacts @(@('Source', "$($s.OpPath) $([char]0x00B7) **not changed**"), @('Saved to', "$($s.OutFolder) ($($s.OutName) and its parts)"), @('Checked', $(if ($s.CheckEach) { 'every file checked: clean' } else { 'not checked (you turned it off)' })))
            $out = @($t, $facts)
            if ($Outcome -eq 'refused') { $out += @(New-MbtGuiRefusedBlock -Verb 'copied') } else { $out += New-GuiMessage -Kind info -Icon info -Text '**What next?** Open the new PST in Outlook, or choose **Open** to look inside it.' }
            return $out
        }
        'partial' { return @((New-GuiMessage -Kind warn -Icon warn -Text '**The copy was stopped.** A stopped copy leaves no file, and your original was never touched. Start it again whenever you want.')) }
        default { return @((New-MbtGuiTiles @(@((Format-MbtGuiCount $done), 'messages copied', 'ok'), @('1', 'error', 'bad'))) + @(New-MbtGuiErrorBlock $run 'The copy stopped. No file was kept.')) }
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
        'partial' { return @((New-GuiMessage -Kind warn -Icon warn -Text '**You stopped the repair.** The new file is written under a temporary name and moved only when complete, so no half-written file is left. Your original was never touched. Start the repair again whenever you want.')) }
        'error' { return @(New-MbtGuiErrorBlock $run 'The repair stopped.') }
    }
    $c = $run.Result.Counts
    $rec = Get-MbtGuiCount $run 'recovered' 0; $part = Get-MbtGuiCount $run 'partial' 0; $lost = Get-MbtGuiCount $run 'lost' 0
    $total = [math]::Max(1, $rec + $part + $lost)
    $pct = ('{0}%' -f ([math]::Round(100 * ($rec + $part) / $total, 1)).ToString([cultureinfo]::InvariantCulture))
    $tiles = New-MbtGuiTiles @(@((Format-MbtGuiCount $rec), 'recovered as they were', 'ok'), @((Format-MbtGuiCount $part), 'partly recovered', 'warn'), @((Format-MbtGuiCount $lost), 'lost', $(if ($lost) { 'bad' } else { 'ok' })), @($pct, 'kept (recovered + partial)', ''))
    $bef = Get-MbtGuiCount $run 'before' $null; $aft = Get-MbtGuiCount $run 'after' $null
    $be = [int](Get-MbtGuiProp $bef 'errors' 0); $bw = [int](Get-MbtGuiProp $bef 'warnings' 0); $ae = [int](Get-MbtGuiProp $aft 'errors' 0); $aw = [int](Get-MbtGuiProp $aft 'warnings' 0)
    $mkpills = { param($e, $w) $p = [Windows.Controls.StackPanel]::new(); $p.Orientation = 'Horizontal'; $a = New-GuiPill "$e errors" $(if ($e) { 'bad' } else { 'ok' }); $a.Margin = [Windows.Thickness]::new(0, 0, 8, 0); $b = New-GuiPill "$w warnings" $(if ($w) { 'warn' } else { 'ok' }); $p.Children.Add($a) | Out-Null; $p.Children.Add($b) | Out-Null; $p }
    $before = New-GuiStack @((New-GuiText 'BEFORE (original, untouched)' 'Small'), (& $mkpills $be $bw)); $after = New-GuiStack @((New-GuiText 'AFTER (new file)' 'Small'), (& $mkpills $ae $aw))
    $ba = New-GuiGrid -Children @($before, $after); $ba.Margin = [Windows.Thickness]::new(0, 12, 0, 0)
    $level = [string](Get-MbtGuiProp $c 'levelUsed' $s.RpLevel)
    $destFile = [string](Get-MbtGuiProp $c 'destination' $s.RpDest)
    $facts = New-GuiFacts @(@('New file', $destFile), @('Level used', "$level$(if ($s.RpLevel -eq 'Automatic') { ' (chosen automatically)' })"), @('Named properties', [string](Get-MbtGuiProp $c 'nameMap' 'recovered from the file (calendar and contact fields kept)')))
    $parts = @($tiles, $ba, $facts)
    $incomplete = @($run.Items | Where-Object { $_.State -in 'partial', 'lost' })
    if ($incomplete.Count) {
        $rows = foreach ($w in $incomplete) { $name = New-GuiStack @((New-GuiText $(if ($w.Key) { $w.Key } else { '(no subject)' }) 'Body'), (New-GuiText $w.Detail 'Small')); , @($name, (New-GuiPill $(if ($w.State -eq 'lost') { 'Lost' } else { 'Partial' }) $(if ($w.State -eq 'lost') { 'bad' } else { 'warn' })), $w.Msg) }
        $parts += New-GuiMore -Summary "The $($incomplete.Count) messages that are not complete" -Open -Content @((New-GuiTable -Headers @('Message', 'Result', 'What was lost') -Rows @($rows) -Widths @('*', '90', '*')))
    }
    $did = @($run.Steps | Where-Object { $_.State -eq 'done' } | ForEach-Object { $x = New-GuiText ([char]0x2022 + '  ' + $_.Title + $(if ($_.Msg) { ': ' + $_.Msg } else { '' })) 'Body'; $x.FontSize = 13; $x })
    if ($did.Count) { $parts += New-GuiMore -Summary 'What the repair did' -Content (@($did) + @(New-GuiText 'Counters, sizes and indexes are recomputed; messages no folder listed are put back in their folders.' 'Hint')) }
    $parts += New-GuiText "The report is saved next to the new file: $([IO.Path]::GetFileNameWithoutExtension($s.RpDest)).repair-report.csv and .json." 'Hint'
    return $parts
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAxVmhuFSflqIF4
# DCCLyTt5fL4TmAKWRffjkx0Dlr9+vqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBpLIBqfYMQ
# FW/X74wrdjKjPN6VuiCDis4y3NenGuIwRjANBgkqhkiG9w0BAQEFAASCAgANGLBL
# 2iPzQNgIcEYrXWltcpRf6+IJfZe61kGS2ogY3/jBxfdWAZI89uoZ4t/258qGNuqW
# /OBwkVmeqTdOE+dD4L66QRHqqXtYqCl4suGK7l3P3/U0wPpLCU/+e9qE1K9ezeKc
# lEgwgG6DQ1eivMYZ0k1ejRaF/pwv8daJPuf7wfAI7crdNVTZjb657Evt/YGkSByU
# ei9TG+zNobOaYf+cM8wVvw7JzX/FcseaAHTHMKhAgLtUFGh3OOviww7a5e/bO7Ns
# 9ZdGQEqku/N19I7eGfMbCY/ZCKLnE8avxCtLTci3wF4aTAN2+Gc7T7rnaB6r5sDn
# 9vII++8YF9nsscdKVjs/vesuNh0dBMfKEEaNHz3Bl7wbKZMWVfhjQSqTAGOQsZ4d
# HFcnzPt6ve73joBGz4IjNIJnytsjqwfHdto/JTgnFJAvJwZL0Ynspy5vCLXO9hbo
# iGJ9iOwvx5T7T7YdETTO2nWYkAJH+0yrrT75lC2Y3HrOPCgXrMDdMtb7SJgl5xOn
# kKyaQpnm1WT3FkW3CDIE/wp0zZhee+HYGTc+JwI8l4Le2+pOs4lxP4+CLKKmGH8G
# vZ+Gm9SMHgUoSK/jrNZCnKYkxCN1wHrwbxkdMYZreyDeQi+0u47pkkaLmv9u/4s/
# r8zMtUGRM+4bhKBUfUyYz9N0ERJNhMFPoLTh2aGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIBk4GRN3QLdxH6aoWdqR6sN8eyAu5nFOkrVxxh35UFhOAggDwXELvdiP
# OBgPMjAyNjExMTAwODIyMzZaMAMCAQECEE1SY8OKc+t8ocKv0jp4TjSgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjIzNlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIAa2
# ydzSMnqElUwhat935kIAcMfenk2BVpG1HyQ6AjNbMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEAhB/9FvE3i1P3O8Ei7vHxFlYoPTiP0e4U3Knzy7uTo6QCIGcJ
# ayjr1SD6nj0EywsoSrW4Nry9OwyCN82ZsnAHVinY
# SIG # End signature block
