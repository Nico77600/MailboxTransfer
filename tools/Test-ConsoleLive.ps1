#Requires -Version 7.4
<#
.SYNOPSIS
    Live, READ-ONLY check of the console's real adapter (lab script, run by the owner; nothing in the automated tests contacts the cloud).

.DESCRIPTION
    Drives the functions the window uses, without a window, in a single-threaded pwsh (run it with  pwsh -STA -File tools\Test-ConsoleLive.ps1):
      1. silent sign-in with the cached token (never a browser, never a device code: if a sign-in is needed the script stops and says so);
      2. the locations (Primary, Archive, the auxiliary archives given with -AuxArchiveId) and a lazy browse of a few folders of each;
      3. the browse of a shared mailbox by its address (-SharedMailbox);
      4. the plan of an export of one small folder to a PST (-WhatIf run of the engine, as the Plan page does);
      5. one real small child job started by the console's job launcher (export of -ExportFolder of the primary mailbox to a PST in -OutDir):
         events, steps, the folders of the run, status.json, the report, the check of the PST;
      6. Stop and Continue of a second job of the same kind.
    Nothing is written to any mailbox: every request to Microsoft is a read (folder trees, item lists, exportItems). Files are written below -OutDir.
    Prints one line per check and writes <OutDir>\live-report.json.
#>
[CmdletBinding()]
param(
    [string]$UserPrincipalName = 'adele.vance@contoso.com',
    [string[]]$AuxArchiveId = @(),
    [string]$SharedMailbox = 'MBT-Shared-Projects@contoso.com',
    [string]$ExportFolder = 'MBT-PstImport-Test-20261109/v2-Mail',
    [string]$OutDir,
    [switch]$SkipJob,
    [switch]$SkipStop
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') { throw 'Run this script with pwsh -STA (the view models of the console are WPF types).' }
$package = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
if (-not $OutDir) { $OutDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'out\console-live' }
$null = New-Item -ItemType Directory -Force -Path $OutDir
$work = Join-Path $OutDir 'work'
$null = New-Item -ItemType Directory -Force -Path $work

foreach ($f in 'Core', 'Run', 'Jobs', 'Browse', 'State', 'Pages', 'Pages.Export', 'Pages.Run', 'Pages.Source', 'Fake', 'Engine', 'Scenarios') { . (Join-Path $package "module\MailboxTransfer.Gui.$f.ps1") }
$script:GuiRoot = $package
Initialize-MbtGuiTypes
$entry = Join-Path $package 'Invoke-MailboxTransfer.ps1'
$engine = New-MbtGuiRealEngine -PackageRoot $package -EntryScript $entry -WorkFolder $work
# the functions of the adapter and of the pages read $script:Gui (state, engine); no window here
$script:Gui = @{ S = (New-MbtGuiState -WorkFolder $work); Engine = $engine; EntryScript = $entry; Pollers = [Collections.Generic.List[scriptblock]]::new(); Browsers = @{}; Run = $null; Job = $null }
$script:Gui.S.Mailbox = $UserPrincipalName

$script:Report = [Collections.Generic.List[object]]::new()
function Add-Check {
    param([string]$Name, [bool]$Ok, [string]$Detail = '')
    $script:Report.Add([pscustomobject]@{ Name = $Name; Ok = $Ok; Detail = $Detail })
    Write-Host ('{0} {1}{2}' -f $(if ($Ok) { '[PASS]' } else { '[FAIL]' }), $Name, $(if ($Detail) { " - $Detail" } else { '' }))
}
function Wait-Adapter {
    <# Calls a member of the adapter with its two callbacks and pumps the worker (the window does it with a timer) until one of them ran. #>
    param([scriptblock]$Start, [int]$TimeoutSeconds = 240)
    $box = @{ Done = $false; Result = $null; Error = $null }
    $ok = { param($r) $box.Result = $r; $box.Done = $true }.GetNewClosure()
    $err = { param($m) $box.Error = [string]$m; $box.Done = $true }.GetNewClosure()
    & $Start $ok $err
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not $box.Done) {
        Step-MbtGuiWork $engine.Worker
        Start-Sleep -Milliseconds 40
        if ($watch.Elapsed.TotalSeconds -gt $TimeoutSeconds) { throw "no answer from the adapter after $TimeoutSeconds s" }
    }
    if ($box.Error) { throw $box.Error }
    return $box.Result
}
function Find-Node {
    param($Parent, [string]$Name)
    @($Parent.Children | Where-Object { $_.Name -eq $Name } | Select-Object -First 1)[0]
}
function Open-Node {
    <# The lazy read of the children of a node (as the window does when a folder is opened). #>
    param($Node)
    if ($Node.ChildrenNotRead -or $Node.HasPlaceholder) {
        $kids = @(Wait-Adapter { param($ok, $err) & $engine.Children $Node $ok $err })
        $Node.SetChildren([MailboxTransferGui.FolderNode[]]@($kids | ForEach-Object { $_ }))
    }
    return @($Node.Children)
}
function Wait-Job {
    param([hashtable]$Job, $Run, [int]$TimeoutSeconds = 900, [scriptblock]$OnTick)
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while (-not $Run.Result) {
        $null = Step-MbtGuiJob -Job $Job -Run $Run
        if ($OnTick) { & $OnTick }
        Start-Sleep -Milliseconds 300
        if ($watch.Elapsed.TotalSeconds -gt $TimeoutSeconds) { throw "the job did not end within $TimeoutSeconds s" }
    }
}

try {
    # ------------------------------------------------------------------------------------------------------------------------------------------
    #  1. Silent sign-in
    # ------------------------------------------------------------------------------------------------------------------------------------------
    Write-Host "== 1. silent sign-in as $UserPrincipalName"
    $cap = $engine.Capability
    Add-Check 'capabilities of the merged engine' ($cap.PstImport -and $cap.EmlExport -and $cap.EmlImport -and $cap.SharedMailbox -and $cap.Events) (($cap.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join ' ')
    $codeFile = Join-Path $work 'device-code.txt'
    if (Test-Path -LiteralPath $codeFile) { Remove-Item -LiteralPath $codeFile -Force }
    try {
        $info = Wait-Adapter { param($ok, $err) & $engine.SignIn $UserPrincipalName $false $codeFile $ok $err $true }
    } catch {
        Add-Check 'silent sign-in with the cached token' $false $_.Exception.Message
        Write-Host 'STOP: no valid cached sign-in (a browser or a device code would be needed). Nothing was done.'
        $script:Report | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutDir 'live-report.json') -Encoding utf8
        exit 3
    }
    Add-Check 'silent sign-in with the cached token' ((Test-Path -LiteralPath $codeFile) -eq $false) "account $($info.Account); no device code was asked"
    Add-Check 'token carries the permissions of the console' (@($info.Missing).Count -eq 0) $(if (@($info.Missing).Count) { 'missing: ' + (@($info.Missing) -join ', ') } else { 'scopes: ' + (@($info.Scopes) -join ' ') })
    $script:Gui.S.Signed = $true; $script:Gui.S.Account = $info.Account; $script:Gui.S.Token = $info

    # ------------------------------------------------------------------------------------------------------------------------------------------
    #  2. Locations and the lazy browse
    # ------------------------------------------------------------------------------------------------------------------------------------------
    Write-Host '== 2. locations and lazy browse'
    $locs = @(Wait-Adapter { param($ok, $err) & $engine.Locations $UserPrincipalName $null $ok $err $null $AuxArchiveId })
    Add-Check 'locations' ($locs.Count -ge 1) (($locs | ForEach-Object { "$($_.Kind) $($_.MailboxId)" }) -join '; ')
    foreach ($want in 'Primary', 'Archive') { Add-Check "location $want is listed" (@($locs | Where-Object Kind -EQ $want).Count -eq 1) }
    if ($AuxArchiveId.Count) { Add-Check 'the known auxiliary archive is listed' (@($locs | Where-Object { $_.Kind -eq 'AuxArchive' }).Count -ge 1) }
    $roots = @(Wait-Adapter { param($ok, $err) & $engine.Roots $locs $ok $err })
    Add-Check 'roots of the tree' ($roots.Count -eq $locs.Count) (($roots | ForEach-Object { "$($_.Name) [$($_.Kind)]" }) -join '; ')
    foreach ($root in $roots) {
        $watch = [Diagnostics.Stopwatch]::StartNew()
        try {
            $top = @(Open-Node $root)
            $detail = "$($top.Count) top folders in $([int]$watch.Elapsed.TotalMilliseconds) ms: " + (($top | Select-Object -First 6 | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', ')
            Add-Check "browse $($root.Name): top folders (lazy)" ($top.Count -gt 0) $detail
            # the first folders: its children, then its first messages
            $opened = 0
            foreach ($f in @($top | Where-Object { $_.Count -gt 0 } | Select-Object -First 3)) {
                $kids = @(Open-Node $f)
                $page = Wait-Adapter { param($ok, $err) & $engine.Items $f $null $ok $err }
                $opened++
                Add-Check "browse $($root.Name) / $($f.Name)" $true "$($kids.Count) subfolder(s), $(@($page.Rows).Count) of $($page.Total) message(s) listed$(if ($page.Note) { ' - ' + $page.Note })$(if (@($page.Rows).Count) { ' - first: ' + $page.Rows[0].Subject })"
            }
            if (-not $opened) { Add-Check "browse $($root.Name): a folder with messages" $false 'no top folder holds messages' }
        } catch { Add-Check "browse $($root.Name)" $false $_.Exception.Message }
    }

    # ------------------------------------------------------------------------------------------------------------------------------------------
    #  3. A shared mailbox by its address
    # ------------------------------------------------------------------------------------------------------------------------------------------
    Write-Host "== 3. shared mailbox $SharedMailbox"
    try {
        $sl = @(Wait-Adapter { param($ok, $err) & $engine.Locations $UserPrincipalName $null $ok $err $SharedMailbox $null })
        Add-Check 'shared mailbox: resolved by its address' ($sl.Count -eq 1) "$($sl[0].Label) -> $($sl[0].MailboxId)"
        $script:Gui.S.Shared = $SharedMailbox
        $sr = @(Wait-Adapter { param($ok, $err) & $engine.Roots $sl $ok $err })
        $stop = @(Open-Node $sr[0])
        Add-Check 'shared mailbox: top folders' ($stop.Count -gt 0) ("$($stop.Count): " + (($stop | Select-Object -First 8 | ForEach-Object { "$($_.Name) ($($_.Count))" }) -join ', '))
        foreach ($f in @($stop | Where-Object { $_.Count -gt 0 } | Select-Object -First 2)) {
            $page = Wait-Adapter { param($ok, $err) & $engine.Items $f $null $ok $err }
            Add-Check "shared mailbox: $($f.Name)" $true "$(@($page.Rows).Count) of $($page.Total) message(s) listed"
        }
        $script:Gui.S.Shared = ''
    } catch { Add-Check 'shared mailbox' $false $_.Exception.Message }

    if ($SkipJob) { Write-Host 'jobs skipped (-SkipJob)'; throw [OperationCanceledException]::new('skip') }

    # ------------------------------------------------------------------------------------------------------------------------------------------
    #  4. The folder to export, its plan, a real job
    # ------------------------------------------------------------------------------------------------------------------------------------------
    Write-Host "== 4. plan and job: $ExportFolder"
    $primary = @($roots | Where-Object Kind -EQ 'primary')[0]
    $node = $primary
    foreach ($part in ($ExportFolder -split '/')) {
        $kids = @(Open-Node $node)
        $node = Find-Node $node $part
        if (-not $node) { throw "folder '$part' of '$ExportFolder' not found in the primary mailbox (found: $(($kids | ForEach-Object Name) -join ', '))" }
    }
    Add-Check "folder $ExportFolder found" $true "$($node.Count) item(s), $(Format-MbtGuiBytes $node.SizeBytes)"

    $s = $script:Gui.S
    $s.Task = 'export'; $s.Dest = 'pst'; $s.OutFolder = $OutDir; $s.CheckEach = $true; $s.SplitGB = 50; $s.Conc = 2; $s.Awake = $true
    $startJob = {
        param([string]$Tag, [string]$PstName, [switch]$Plan)
        $s.OutName = $PstName
        $jobStore = Join-Path $work "stores\$Tag"
        $job = New-MbtGuiJob -WorkFolder $work -Task $(if ($Plan) { "$Tag-plan" } else { $Tag }) -EntryScript $entry -StorePath $jobStore
        $selFile = Join-Path $job.Folder 'selection-1.json'
        $entries = @([pscustomobject]@{ Node = $node; Kind = 'tree'; Items = 0 })
        $null = Wait-Adapter { param($ok, $err) & $engine.SaveSelection $primary $entries $selFile $ok $err }
        $job.SelectionFiles = @($selFile)
        foreach ($c in (Get-MbtGuiCommand -Task export -S $s -Job $job -Plan:$Plan)) { $job.Queue.Enqueue($c) }
        if (-not $Plan) { Write-MbtGuiRequestFile -Job $job }
        $run = New-MbtGuiRun -Task export -Title "Live check $Tag"
        $null = Start-MbtGuiProcess -Job $job
        [pscustomobject]@{ Job = $job; Run = $run }
    }

    # the plan: the read-only run of the engine, as the Plan page does
    $p = & $startJob 'plan' 'live-plan.pst' -Plan
    Wait-Job $p.Job $p.Run 300
    $planObjects = @($p.Run.Objects | Where-Object Step -EQ 'plan')
    Add-Check 'plan: the read-only run ends with exit 0 and writes no PST' ($p.Run.Result.Exit -eq 0 -and -not (Test-Path -LiteralPath (Join-Path $OutDir 'live-plan.pst'))) ("exit $($p.Run.Result.Exit); $($planObjects.Count) planned folder(s): " + (($planObjects | ForEach-Object { "$($_.Key) ($($_.Total))" }) -join ', '))

    # the real job
    $j = & $startJob 'export1' 'live-export.pst'
    Wait-Job $j.Job $j.Run 900
    $run = $j.Run; $folder = $j.Job.Folder
    Add-Check 'job: ends with exit 0 and a result event' ($run.Result.Exit -eq 0) "exit $($run.Result.Exit), counts: $(($run.Result.Counts | ConvertTo-Json -Compress -Depth 3))"
    Add-Check 'job: the plan event gave the steps before the start' ($run.Planned -and $run.Steps.Count -ge 4) (($run.Steps | ForEach-Object { "$($_.Id)=$($_.State)" }) -join ' ')
    Add-Check 'job: the Run page folder list is filled (object events of the download and of the PST)' (@($run.Objects | Where-Object { $_.Step -eq 'download' -and $_.State -eq 'done' }).Count -ge 1 -and @($run.Objects | Where-Object { $_.Step -eq 'pst' -and $_.State -eq 'done' }).Count -ge 1) (($run.Objects | Where-Object Step -In 'download', 'pst' | ForEach-Object { "$($_.Step):$($_.Key)=$($_.State) $($_.Done)/$($_.Total)" }) -join '; ')
    $status = Get-Content -LiteralPath (Join-Path $folder 'status.json') -Raw | ConvertFrom-Json
    Add-Check 'job: status.json says complete' ($status.state -eq 'complete' -and $status.exitCode -eq 0) "state $($status.state), counts $($status.counts | ConvertTo-Json -Compress)"
    $listed = @(Get-MbtJob -WorkFolder $work | Where-Object { $_.JobId -eq $j.Job.Id })
    Add-Check 'job: the Continue list finds it (Get-MbtJob), complete' ($listed.Count -eq 1 -and $listed[0].EffectiveState -eq 'complete') ''
    $reports = @(Write-MbtGuiRunReport -Run $run -Folder $folder -Task export)
    Add-Check 'job: the Done report (CSV and JSON) is written' ((@($reports | Where-Object { Test-Path -LiteralPath $_ })).Count -eq 2) (($reports | ForEach-Object { Split-Path $_ -Leaf }) -join ', ')
    $pstFile = Join-Path $OutDir 'live-export.pst'
    if (Test-Path -LiteralPath $pstFile) {
        Import-Module (Join-Path $package 'MailboxTransfer.psd1') -Force
        $chk = Test-MbtPst -Path $pstFile
        Add-Check 'job: the PST passes the checker' ($chk.ExitCode -eq 0) "$((Get-Item $pstFile).Length) bytes, errors $($chk.Errors), warnings $($chk.Warnings)"
    } else { Add-Check 'job: the PST was written' $false 'live-export.pst is missing' }

    # ------------------------------------------------------------------------------------------------------------------------------------------
    #  6. Stop and Continue
    # ------------------------------------------------------------------------------------------------------------------------------------------
    if ($SkipStop) { Write-Host 'stop and continue skipped (-SkipStop)'; throw [OperationCanceledException]::new('skip') }
    Write-Host '== 6. stop and continue'
    $k = & $startJob 'export2' 'live-stop.pst'
    $asked = $false
    Wait-Job $k.Job $k.Run 600 -OnTick {
        # the stop request goes in as soon as the job has told its plan
        if (-not $asked -and $k.Run.Planned) { Request-MbtGuiStop -Job $k.Job; $script:asked = $true; $k.Run.StopRequested = $true }
    }
    $stopRes = $k.Run.Result
    $stopped = ($stopRes.Exit -eq 2 -and $stopRes.Reason -eq 'stopped')
    $finishedBeforeStop = ($stopRes.Exit -eq 0)
    Add-Check 'stop: "Stop and keep progress" ends the job resumable (exit 2, reason stopped, next continue)' ($stopped -or $finishedBeforeStop) "exit $($stopRes.Exit), reason '$($stopRes.Reason)', next '$($stopRes.Next)'$(if ($finishedBeforeStop) { ' (the small job was already finished when the request came)' })"
    $st2 = Get-Content -LiteralPath (Join-Path $k.Job.Folder 'status.json') -Raw | ConvertFrom-Json
    Add-Check 'stop: status.json says stopped (or complete)' ($st2.state -in 'stopped', 'complete') "state $($st2.state)"
    $entryJob = @(Get-MbtJob -WorkFolder $work | Where-Object { $_.JobId -eq $k.Job.Id })[0]
    Add-Check 'stop: the Continue list offers it' ($null -ne $entryJob -and ($entryJob.EffectiveState -in 'stopped', 'complete')) "state $($entryJob.EffectiveState)"

    # Continue: the same request again in the same job folder (what Resume-MbtGuiJob does)
    $req = Get-Content -LiteralPath (Join-Path $k.Job.Folder 'request.json') -Raw | ConvertFrom-Json
    $job2 = New-MbtGuiJob -WorkFolder $work -Task export -EntryScript $entry -StorePath ([string]$req.StorePath) -SelectionFiles @($req.SelectionFiles) -JobName ([string]$req.JobName)
    $job2.Folder = $k.Job.Folder
    foreach ($key in 'EventsPath', 'StopPath', 'StatusPath', 'StdOut', 'StdErr') { $job2[$key] = Join-Path $k.Job.Folder (Split-Path $job2[$key] -Leaf) }
    if (Test-Path -LiteralPath $job2.StopPath) { Remove-Item -LiteralPath $job2.StopPath -Force }
    $job2.Offset = $(if (Test-Path -LiteralPath $job2.EventsPath) { (Get-Item -LiteralPath $job2.EventsPath).Length } else { 0L })
    foreach ($c in @($req.Commands)) { $job2.Queue.Enqueue([pscustomobject]@{ Title = 'Continue'; Arguments = @($c) }) }
    $run2 = New-MbtGuiRun -Task export -Title 'Live check: continue'
    $null = Start-MbtGuiProcess -Job $job2
    Wait-Job $job2 $run2 900
    Add-Check 'continue: the same request again finishes the job (exit 0)' ($run2.Result.Exit -eq 0) "exit $($run2.Result.Exit), counts: $(($run2.Result.Counts | ConvertTo-Json -Compress -Depth 3))"
    $pst2 = Join-Path $OutDir 'live-stop.pst'
    if (Test-Path -LiteralPath $pst2) {
        Import-Module (Join-Path $package 'MailboxTransfer.psd1') -Force
        $chk2 = Test-MbtPst -Path $pst2
        Add-Check 'continue: the PST passes the checker' ($chk2.ExitCode -eq 0) "$((Get-Item $pst2).Length) bytes, errors $($chk2.Errors), warnings $($chk2.Warnings)"
    } else { Add-Check 'continue: the PST was written' $false 'live-stop.pst is missing' }
} catch [OperationCanceledException] {
    # a skip
} catch {
    Add-Check 'the check itself' $false ($_.Exception.Message + ' at ' + ($_.ScriptStackTrace -split "`n")[0])
} finally {
    $script:Report | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutDir 'live-report.json') -Encoding utf8
    try { Stop-MbtGuiWorker $engine.Worker } catch { }
}
$failed = @($script:Report | Where-Object { -not $_.Ok })
Write-Host ("{0} check(s), {1} failed. Report: {2}" -f $script:Report.Count, $failed.Count, (Join-Path $OutDir 'live-report.json'))
exit $(if ($failed.Count) { 1 } else { 0 })
