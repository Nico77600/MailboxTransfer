<#
.SYNOPSIS
    Mailbox Transfer - events of a job, stop request, keep awake, job status (dot-sourced by MailboxTransfer.psm1).

.DESCRIPTION
    The console reads what a job does from events.jsonl in the job folder (format: docs\MailboxTransfer-Events.md). This file holds
    the module side of it:

        Open-MbtEventLog / Write-MbtEvent / Close-MbtEventLog    the JSON-lines log of the current job (a no-op when none is open)
        Write-MbtJobPlan / Write-MbtJobStep / Write-MbtJobResult / Write-MbtJobError    helpers of the entry script and of the module
        Write-MbtJobStatus / Get-MbtJob                          status.json of a job folder, and the list of the jobs of a work folder
        Start-MbtStopWatch / Stop-MbtStopWatch / Test-MbtStopRequested / Get-MbtStopToken    a stop-request file stops a job gracefully
        Start-MbtKeepAwake / Stop-MbtKeepAwake                   Windows does not sleep by idle while a long job runs
        Get-MbtCapability                                        what the engine can do (the console shows the rest as "coming")

.NOTES
    Author  : Nicolas Fabert
#>

$script:EventLog = $null
$script:StopWatcher = $null
$script:KeepAwake = $null
$script:JobInfo = $null
$script:JobProgress = $null

#region Event log ------------------------------------------------------------------------------

function Open-MbtEventLog {
    <#
    .SYNOPSIS
        Opens the events file of the job (appends to it) and writes status.json (state running) in the same folder.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$JobId,
        [string]$Task,
        [string]$Title,
        [string]$Mode,
        [string]$Store,
        # period of the heartbeat that detects a sleep (milliseconds, 0 = none)
        [int]$HeartbeatMs = 5000
    )
    Initialize-MbtEngine
    Close-MbtEventLog
    $full = [IO.Path]::GetFullPath($Path)
    $script:EventLog = [MailboxTransfer.EventLog]::new($full, $null, $HeartbeatMs, 120)
    if (-not $JobId) { $JobId = (Split-Path (Split-Path $full -Parent) -Leaf) }
    $script:JobInfo = @{ Folder = (Split-Path $full -Parent); Id = $JobId; Task = $Task; Title = $Title; Mode = $Mode; Store = $Store; Started = [DateTime]::UtcNow; Steps = @{}; Counts = @{}; StatusAt = [DateTime]::UtcNow; Finished = $false }
    $script:JobProgress = @{ Step = ''; ThrottleSeen = 0; Waited = 0.0; WindowsSeen = 0; Warns = 0; Last = [DateTime]::MinValue }
    Write-MbtJobStatus -JobFolder $script:JobInfo.Folder -State running -JobId $JobId -Task $Task -Title $Title -Mode $Mode -Store $Store
}

function Close-MbtEventLog {
    [CmdletBinding()]
    param()
    if ($script:EventLog) { $script:EventLog.Dispose(); $script:EventLog = $null }
}

function Write-MbtEvent {
    <# One line of the events file: {"v":1,"t":...,"e":Kind, ...Data}. Does nothing when no log is open. #>
    [CmdletBinding()]
    param([Parameter(Mandatory, Position = 0)][string]$Kind, [Parameter(Position = 1)][System.Collections.IDictionary]$Data)
    if (-not $script:EventLog) { return }
    $script:EventLog.Write($Kind, $Data)
}

function Write-MbtJobPlan {
    <# The plan event: the steps of the job, all known before the start. -Steps: @( @{ Id = 'tree'; Title = 'Read the folder tree' }, ... ) #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][hashtable[]]$Steps)
    if (-not $script:EventLog) { return }
    $list = @($Steps | ForEach-Object { [ordered]@{ id = $_.Id; title = $_.Title } })
    foreach ($s in $Steps) { $script:JobInfo.Steps[$s.Id] = 'waiting' }
    Write-MbtEvent plan ([ordered]@{ job = $script:JobInfo.Id; task = $script:JobInfo.Task; mode = $script:JobInfo.Mode; steps = $list })
}

function Update-MbtJobStatusTick {
    <#
    .SYNOPSIS
        Keeps status.json of a running job alive: at most every 15 seconds it writes state running with the time of the last event and, when the
        step that is counting gives them, counts.done and counts.total (items). The Continue list of the console reads it for a job whose process
        is gone ("stopped <time>, N items done"): without this the file keeps the state of the start of the job for hours.
    #>
    [CmdletBinding()]
    param([long]$Done = -1, [long]$Total = -1)
    if (-not $script:EventLog -or -not $script:JobInfo -or $script:JobInfo.Finished) { return }
    $now = [DateTime]::UtcNow
    if (($now - $script:JobInfo.StatusAt).TotalSeconds -lt 15) { return }
    $script:JobInfo.StatusAt = $now
    $counts = $null
    if ($Done -ge 0) { $counts = @{ done = $Done }; if ($Total -ge 0) { $counts.total = $Total } }
    try {
        if ($counts) { Write-MbtJobStatus -JobFolder $script:JobInfo.Folder -State running -Counts $counts } else { Write-MbtJobStatus -JobFolder $script:JobInfo.Folder -State running }
    } catch { Write-Verbose "status.json not refreshed: $($_.Exception.Message)" }
}

function Write-MbtJobStep {
    <# A step event (state waiting | running | done | failed) with its counters. Counters left out are unknown, not zero. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Id,
        [Parameter(Mandatory)][ValidateSet('waiting', 'running', 'done', 'failed')][string]$State,
        [long]$Done = -1, [long]$Total = -1, [long]$Bytes = -1, [long]$BytesTotal = -1, [string]$Msg
    )
    if (-not $script:EventLog) { return }
    $script:JobInfo.Steps[$Id] = $State
    $script:EventLog.Step($Id, $State, $Done, $Total, $Bytes, $BytesTotal, $(if ($Msg) { $Msg }))
    if ($State -eq 'running') { Update-MbtJobStatusTick -Done $Done -Total $Total }
}

function Write-MbtJobObject {
    <# An object event: a folder or a group of items with its state and counters (state e.g. planned, running, done, check, refused, notready). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Step, [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][string]$State, [long]$Done = -1, [long]$Total = -1, [long]$Bytes = -1, [string]$Msg)
    if (-not $script:EventLog) { return }
    $script:EventLog.Object($Step, $Key, $State, $Done, $Total, $(if ($Msg) { $Msg }), $Bytes)
    Update-MbtJobStatusTick
}

function Write-MbtJobResult {
    <#
    .SYNOPSIS
        The result event (exit code of the script, counts, what to do next) and the final status.json.
    .PARAMETER Next
        continue (run the same command again), decide (rows to settle), retry (after an error), done, signin (sign in again, then continue).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][int]$ExitCode, [hashtable]$Counts, [string]$Next, [string]$Reason, [switch]$Complete)
    if (-not $script:EventLog) { return }
    if (-not $Next) { $Next = switch ($ExitCode) { 0 { 'done' } 2 { 'continue' } 3 { 'decide' } default { 'retry' } } }
    $isComplete = [bool]$Complete -or ($ExitCode -eq 0)
    $script:EventLog.Result($ExitCode, $isComplete, $(if ($Counts) { $Counts } else { @{} }), $Next, $(if ($Reason) { $Reason }))
    $state = if ($ExitCode -eq 0 -or ($ExitCode -eq 1 -and $Reason -eq 'refused-items')) { 'complete' } elseif ($ExitCode -eq 3) { 'needs-decision' } elseif ($ExitCode -eq 2) { if ($Reason -eq 'stopped') { 'stopped' } else { 'resumable' } } else { 'error' }
    $script:JobInfo.Finished = $true
    Write-MbtJobStatus -JobFolder $script:JobInfo.Folder -State $state -Counts $Counts -ExitCode $ExitCode -Next $Next -Reason $Reason
}

function ConvertTo-MbtPlainError {
    <# The plain sentence for an engine message (CONSOLE-SPEC 11): what happened, what is safe. The raw message goes to "tech". Returns Code and Message. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Message)
    $map = @(
        @{ Re = 'Not signed in|AADSTS(50173|70043|700082|50076|50079)|interaction_required|invalid_grant'; Code = 'signin-required'; Text = 'Your sign-in has expired. Sign in again to continue. Nothing is lost.' }
        @{ Re = 'not enough space on the disk|disk is full|No space left'; Code = 'disk-full'; Text = 'The disk is full. Free some space and press Try again. What was downloaded is kept.' }
        @{ Re = 'Windows device name|reserved name|FileNaming'; Code = 'reserved-name'; Text = 'This file name is a name Windows keeps for itself (or has a character Windows refuses). Choose another file name.' }
        @{ Re = 'target exists|already exists|Overwrite'; Code = 'destination-exists'; Text = 'This file already exists. The tool never overwrites a file. Choose another name.' }
        @{ Re = 'ErrorMailboxNotFound|MailboxNotFound'; Code = 'mailbox-not-found'; Text = 'We cannot find this mailbox. Check the address. Mailboxes that are not in Exchange Online cannot be used.' }
        @{ Re = '403|ErrorAccessDenied|Access is denied|Forbidden'; Code = 'access-denied'; Text = 'This sign-in is not allowed to do this. Ask your administrator for the permission named in the technical message.' }
        @{ Re = 'database is locked|SQLITE_BUSY|being used by another process'; Code = 'store-busy'; Text = 'Another MailboxTransfer is working on this folder. Wait for it, or open it from Continue.' }
    )
    foreach ($m in $map) { if ($Message -match $m.Re) { return [pscustomobject]@{ Code = $m.Code; Message = $m.Text } } }
    [pscustomobject]@{ Code = 'error'; Message = 'The task stopped because of an error. What was done is kept.' }
}

function Write-MbtJobError {
    <# The error event for an exception or a message (code and plain sentence from ConvertTo-MbtPlainError, the raw text in "tech"). Returns the plain error. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Message, [string]$Code, [string]$Plain)
    $p = ConvertTo-MbtPlainError -Message $Message
    if (-not $Code) { $Code = $p.Code }
    if (-not $Plain) { $Plain = $p.Message }
    if ($script:EventLog) { $script:EventLog.Error($Code, $Plain, $Message) }
    [pscustomobject]@{ Code = $Code; Message = $Plain }
}

#endregion

#region Job status -----------------------------------------------------------------------------

function Read-MbtJsonFile {
    <# A JSON file as an object (or a hashtable): dates stay text (ConvertFrom-Json turns ISO strings into dates, from which -DateKind String is the way back, PowerShell 7.5). #>
    param([Parameter(Mandatory)][string]$Path, [switch]$AsHashtable)
    $text = Get-Content -LiteralPath $Path -Raw
    $args = @{}
    if ($AsHashtable) { $args.AsHashtable = $true }
    if ((Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { $args.DateKind = 'String' }
    $text | ConvertFrom-Json @args
}

function Write-MbtJobStatus {
    <#
    .SYNOPSIS
        Writes (or updates) status.json in a job folder: what the Continue list of the console reads.
    .DESCRIPTION
        Fields: v, jobId, task, title, mode, store, state (running | stopped | resumable | needs-decision | complete | error), exitCode,
        counts, next, reason, started, updated, lastEvent, pid. A call updates the fields it is given and keeps the others.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$JobFolder,
        [Parameter(Mandatory)][ValidateSet('running', 'stopped', 'resumable', 'needs-decision', 'complete', 'error')][string]$State,
        [string]$JobId, [string]$Task, [string]$Title, [string]$Mode, [string]$Store,
        [hashtable]$Counts, [Nullable[int]]$ExitCode, [string]$Next, [string]$Reason
    )
    if (-not (Test-Path -LiteralPath $JobFolder)) { New-Item -ItemType Directory -Path $JobFolder -Force | Out-Null }
    $file = Join-Path $JobFolder 'status.json'
    $s = [ordered]@{ v = 1; jobId = $null; task = $null; title = $null; mode = $null; store = $null; state = $null; exitCode = $null; counts = @{}; next = $null; reason = $null; started = $null; updated = $null; lastEvent = $null; pid = $PID }
    if (Test-Path -LiteralPath $file) {
        try { $old = Read-MbtJsonFile -Path $file -AsHashtable; foreach ($k in $old.Keys) { $s[$k] = $(if ($old[$k] -is [datetime]) { $old[$k].ToUniversalTime().ToString('yyyy-MM-dd\THH:mm:ss.fff\Z') } else { $old[$k] }) } } catch { }
    }
    $now = [DateTime]::UtcNow.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z')
    if (-not $s.started) { $s.started = $now }
    $s.state = $State; $s.updated = $now; $s.pid = $PID
    $s.lastEvent = if ($script:EventLog) { [MailboxTransfer.EventLog]::Iso($script:EventLog.LastEventUtc) } else { $now }
    foreach ($pair in @(@('jobId', $JobId), @('task', $Task), @('title', $Title), @('mode', $Mode), @('store', $Store), @('next', $Next), @('reason', $Reason))) { if ($pair[1]) { $s[$pair[0]] = $pair[1] } }
    if ($Counts) { $s.counts = $Counts }
    if ($PSBoundParameters.ContainsKey('ExitCode') -and $null -ne $ExitCode) { $s.exitCode = [int]$ExitCode }
    $json = $s | ConvertTo-Json -Depth 6
    $tmp = "$file.$PID.tmp"
    [IO.File]::WriteAllText($tmp, $json, [Text.UTF8Encoding]::new($false))
    try { Move-Item -LiteralPath $tmp -Destination $file -Force } catch { [IO.File]::WriteAllText($file, $json, [Text.UTF8Encoding]::new($false)); Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
}

function Get-MbtJob {
    <#
    .SYNOPSIS
        The jobs found in a work folder (<WorkFolder>\jobs\<job>\status.json), newest first: the list of "Continue a job".
    .DESCRIPTION
        A job whose state is running but whose process is gone (crash, power loss, killed) is reported with EffectiveState resumable and
        Interrupted = $true; if the process is still there, Alive is $true ("running elsewhere").
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkFolder, [int]$Limit = 50)
    $jobs = Join-Path $WorkFolder 'jobs'
    if (-not (Test-Path -LiteralPath $jobs)) { return }
    $found = foreach ($dir in Get-ChildItem -LiteralPath $jobs -Directory -ErrorAction SilentlyContinue) {
        $file = Join-Path $dir.FullName 'status.json'
        if (-not (Test-Path -LiteralPath $file)) { continue }
        try { $s = Read-MbtJsonFile -Path $file } catch { continue }
        $get = { param($n) $p = $s.PSObject.Properties[$n]; if ($p) { $p.Value } }
        $parse = { param($v, $fallback) if ($v -is [datetime]) { $v.ToUniversalTime() } elseif ($v) { [DateTime]::Parse([string]$v, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal) } else { $fallback } }
        $started = & $parse (& $get 'started') $dir.CreationTimeUtc
        $updated = & $parse (& $get 'updated') $dir.LastWriteTimeUtc
        $state = [string](& $get 'state'); $jobPid = & $get 'pid'
        $alive = $false
        if ($state -eq 'running' -and $jobPid) {
            $proc = Get-Process -Id ([int]$jobPid) -ErrorAction SilentlyContinue
            if ($proc) { try { $alive = $proc.StartTime.ToUniversalTime() -le $started.AddSeconds(30) } catch { $alive = $true } }
        }
        $effective = if ($state -eq 'running' -and -not $alive) { 'resumable' } else { $state }
        [pscustomobject]@{
            JobId = & $get 'jobId'; Task = & $get 'task'; Title = & $get 'title'; Mode = & $get 'mode'; Store = & $get 'store'; State = $state; EffectiveState = $effective
            Alive = $alive; Interrupted = ($state -eq 'running' -and -not $alive); ExitCode = & $get 'exitCode'; Counts = & $get 'counts'; Next = & $get 'next'; Reason = & $get 'reason'
            Started = $started; Updated = $updated; Folder = $dir.FullName; EventsFile = (Join-Path $dir.FullName 'events.jsonl')
        }
    }
    @($found | Sort-Object Updated -Descending | Select-Object -First $Limit)
}

#endregion

#region Stop request, keep awake -----------------------------------------------------------------

function Start-MbtStopWatch {
    <# Watches a stop-request file: when it exists, Get-MbtStopToken is cancelled (as Ctrl+C does) and Test-MbtStopRequested is true. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [int]$PollMilliseconds = 1000)
    Initialize-MbtEngine
    Stop-MbtStopWatch
    $script:StopWatcher = [MailboxTransfer.StopWatcher]::new([IO.Path]::GetFullPath($Path), $PollMilliseconds)
}

function Stop-MbtStopWatch {
    [CmdletBinding()]
    param()
    if ($script:StopWatcher) { $script:StopWatcher.Dispose(); $script:StopWatcher = $null }
}

function Test-MbtStopRequested {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    return [bool]($script:StopWatcher -and $script:StopWatcher.IsRequested)
}

function Get-MbtStopToken {
    <# The token cancelled by a stop request ([Threading.CancellationToken]::None without a watcher). #>
    [CmdletBinding()]
    [OutputType([Threading.CancellationToken])]
    param()
    if ($script:StopWatcher) { return $script:StopWatcher.Token }
    return [Threading.CancellationToken]::None
}

function Start-MbtKeepAwake {
    <# Windows does not go to sleep by idle until Stop-MbtKeepAwake (a closed lid, Sleep chosen by hand or a restart still stop the PC). #>
    [CmdletBinding()]
    param()
    Initialize-MbtEngine
    if ($script:KeepAwake) { return $script:KeepAwake }
    $script:KeepAwake = [MailboxTransfer.AwakeKeeper]::Begin()
    return $script:KeepAwake
}

function Stop-MbtKeepAwake {
    [CmdletBinding()]
    param()
    if ($script:KeepAwake) { $script:KeepAwake.Dispose(); $script:KeepAwake = $null }
}

#endregion

#region Capabilities ----------------------------------------------------------------------------

function Get-MbtCapability {
    <#
    .SYNOPSIS
        What this engine can do, decided from the entry script's parameters and the module's commands: the console shows what is missing as
        "coming in this version" and lights it up when it is there.
    .DESCRIPTION
        Events        parameter -EventsFile of the entry script
        PstExport     Mode ExportPst          PstCheck  Mode CheckPst          PstRepair  Mode RepairPst
        PstImport     Mode ImportPst (an import from a PST file: the entry script takes -PstPath, -TargetRoot)
        FolderChoice  parameters -IncludeFolder / -ExcludeFolder: a PST or a folder of .eml files is imported (or a PST copied) in part
        PstCopy       Mode ExportPst with -PstPath (a PST copied into new PST files: -IncludeFolder, -FolderFile, -IncludeItemFile)
        PstPreview    command Get-MbtPstPreview (the text and attachments of a message of a PST)
        EmlExport     Mode ExportEml          EmlImport  Mode ImportEml
        AppOnly       parameter -CertificateThumbprint of the entry script (sign-in as the application with a certificate, for every mode)
        SharedMailbox parameter -MailboxId of the entry script that takes an SMTP address / UPN (command Resolve-MbtMailboxAddress): a shared
                      mailbox or another user's mailbox is given by its address
    #>
    [CmdletBinding()]
    param([string]$EntryScript = (Join-Path $script:ToolRoot 'Invoke-MailboxTransfer.ps1'))
    $cmd = Get-Command -Name $EntryScript -CommandType ExternalScript -ErrorAction Stop
    $modes = @()
    if ($cmd.Parameters.ContainsKey('Mode')) { $modes = @($cmd.Parameters['Mode'].Attributes | Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } | ForEach-Object { $_.ValidValues }) }
    $has = { param($n) $cmd.Parameters.ContainsKey($n) }
    [pscustomobject]@{
        Events        = [bool](& $has 'EventsFile')
        StopFile      = [bool](& $has 'StopFile')
        KeepAwake     = [bool](& $has 'KeepAwake')
        PstExport     = $modes -contains 'ExportPst'
        PstCheck      = $modes -contains 'CheckPst'
        PstRepair     = $modes -contains 'RepairPst'
        PstImport     = ($modes -contains 'ImportPst') -and [bool](& $has 'PstPath')
        FolderChoice  = [bool](& $has 'IncludeFolder')
        PstCopy       = ($modes -contains 'ExportPst') -and [bool](& $has 'IncludeItemFile') -and [bool](& $has 'FolderFile')
        PstPreview    = [bool](Get-Command -Name Get-MbtPstPreview -ErrorAction SilentlyContinue)
        EmlExport     = $modes -contains 'ExportEml'
        EmlImport     = $modes -contains 'ImportEml'
        AppOnly       = [bool](& $has 'CertificateThumbprint')
        SharedMailbox = [bool]((& $has 'MailboxId') -and (Get-Command -Name Resolve-MbtMailboxAddress -ErrorAction SilentlyContinue))
        Modes         = $modes
    }
}

#endregion

#region Progress of the long jobs -----------------------------------------------------------------

function Get-MbtJobStepOrder {
    switch ($script:JobInfo.Task) {
        'export' { @('tree', 'list', 'download', 'verify', 'pst', 'pstcheck') }
        'import' { @('check', 'folders', 'send', 'reconcile', 'final') }
        default { @() }
    }
}

function Update-MbtJobProgress {
    <#
    .SYNOPSIS
        Called by Wait-MbtTransferJob on every poll: the step events of the job phases and counters, the wait events of the throttling
        answers, the rate events of the windows, a warn event per engine message that is a warning. Does nothing without an event log.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Job, [string[]]$Messages = @(), [switch]$Final)
    if (-not $script:EventLog -or -not $script:JobInfo) { return }
    $p = $script:JobProgress
    $snap = $Job.Snapshot()
    $order = @(Get-MbtJobStepOrder)
    $export = $script:JobInfo.Task -eq 'export'
    $current = switch -Regex ($snap.Phase) {
        '^tree$' { 'tree' } '^inventory$' { 'list' } '^export$' { 'download' }
        '^check$' { 'check' } '^folders$|^plan$' { 'folders' } '^import$' { 'send' } '^reconcile$' { 'reconcile' }
        default { '' }
    }
    if ($current -and $order -contains $current -and $current -ne $p.Step) {
        $i = [array]::IndexOf($order, $current)
        # the steps before the current one are done (a reconcile pass is a side step of the send step)
        if ($current -ne 'reconcile' -and $i -gt 0) { foreach ($s in @($order[0..($i - 1)])) { if ($script:JobInfo.Steps[$s] -in 'waiting', 'running') { Write-MbtJobStep -Id $s -State done } } }
        Write-MbtJobStep -Id $current -State running
        $p.Step = $current
    }
    $counting = if ($export) { 'download' } else { 'send' }
    # the folders that changed since the last poll (the Run page lists them)
    try { if ($Job.Metrics -and $Job.Metrics.Folders) { [void][MailboxTransfer.FolderEventPump]::Write($script:EventLog, $Job.Metrics.Folders, $counting) } } catch { }
    if ($p.Step -eq $counting -or ($p.Step -eq 'reconcile' -and -not $export)) {
        Write-MbtJobStep -Id $counting -State running -Done $snap.Items -Total $snap.TotalItems -Bytes $snap.Bytes -BytesTotal $snap.TotalBytes
    }
    # throttling answers (429 / 503) since the last poll
    try {
        $events = @($Job.GetThrottleEvents())
        if ($events.Count -gt $p.ThrottleSeen) {
            foreach ($e in $events[$p.ThrottleSeen..($events.Count - 1)]) {
                $p.Waited += $e.WaitSeconds
                $script:EventLog.Wait('throttle', [math]::Round($e.WaitSeconds, 1), ($p.ThrottleSeen + 1), [math]::Round($p.Waited, 1))
                $p.ThrottleSeen++
            }
        }
    } catch { }
    $windows = @($Job.GetWindows())
    if ($windows.Count -gt $p.WindowsSeen) {
        foreach ($w in $windows[$p.WindowsSeen..($windows.Count - 1)]) { $script:EventLog.Rate('5m', $w.GiBPerHour) }
        $p.WindowsSeen = $windows.Count
    }
    foreach ($m in $Messages) {
        if ($m -match 'failed|not listed|not created|not imported|refused|Unknown|lowered' -and $m -notmatch '^throttled' -and $p.Warns -lt 500) {
            $p.Warns++
            $script:EventLog.Warn('engine', $p.Step, $null, $m)
        }
    }
    if ($Final) {
        $done = if ($Job.Cancelled -or $Job.Error) { $null } else { 'done' }
        # the job ended well: every step of the job is done, also the ones that went by between two polls
        foreach ($s in $order) {
            if ($done -and $s -in 'tree', 'list', 'download', 'check', 'folders', 'send', 'reconcile' -and $script:JobInfo.Steps[$s] -ne 'done') {
                if ($s -eq $counting) { Write-MbtJobStep -Id $s -State done -Done $snap.Items -Total $snap.TotalItems -Bytes $snap.Bytes -BytesTotal $snap.TotalBytes } else { Write-MbtJobStep -Id $s -State done }
            }
        }
        if ($export -and $done) { Write-MbtJobStep -Id verify -State done -Msg 'Each item is checked against its fingerprint as it arrives.' }
        if ($Job.Error) { if ($p.Step) { Write-MbtJobStep -Id $p.Step -State failed -Msg $Job.Error.Message } }
    }
}

function Get-MbtPstProgressHook {
    <#
    .SYNOPSIS
        The progress callback of a PST check (Kind Check) or of a repair (Kind Repair) that writes step events, or $null when no event log
        is open for such a job. It receives a [MailboxTransfer.PstProgress].
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Check', 'Repair')][string]$Kind)
    if (-not $script:EventLog -or -not $script:JobInfo) { return $null }
    if ($Kind -eq 'Check' -and $script:JobInfo.Task -ne 'check') { return $null }
    if ($Kind -eq 'Repair' -and $script:JobInfo.Task -ne 'repair') { return $null }
    $state = @{ Stage = 'before'; Phase = ''; Watch = [Diagnostics.Stopwatch]::StartNew() }
    if ($Kind -eq 'Check') {
        return {
            param($e)
            if ($e.Phase -ne $state.Phase -or $state.Watch.ElapsedMilliseconds -ge 500) {
                $state.Phase = $e.Phase; $state.Watch.Restart()
                Write-MbtJobStep -Id check -State running -Bytes $e.BytesDone -BytesTotal $e.BytesTotal -Done $e.Items -Total $e.ItemsTotal -Msg $e.Phase
            }
        }.GetNewClosure()
    }
    return {
        param($e)
        if ($e.Phase -eq 'repair') {
            if ($state.Stage -eq 'before') {
                Write-MbtJobStep -Id before -State done
                $state.Stage = 'repair'
            }
            if ($state.Phase -ne 'repair' -or $state.Watch.ElapsedMilliseconds -ge 500) {
                $state.Phase = 'repair'; $state.Watch.Restart()
                Write-MbtJobStep -Id scan -State running -Done $e.Items -Total $e.ItemsTotal
                Write-MbtJobStep -Id write -State running -Done $e.Items -Total $e.ItemsTotal
            }
            return
        }
        if ($state.Stage -eq 'repair') {
            $state.Stage = 'after'
            Write-MbtJobStep -Id scan -State done
            Write-MbtJobStep -Id write -State done
        }
        if ($e.Phase -ne $state.Phase -or $state.Watch.ElapsedMilliseconds -ge 500) {
            $state.Phase = $e.Phase; $state.Watch.Restart()
            Write-MbtJobStep -Id $(if ($state.Stage -eq 'after') { 'after' } else { 'before' }) -State running -Bytes $e.BytesDone -BytesTotal $e.BytesTotal -Msg $e.Phase
        }
    }.GetNewClosure()
}

function Start-MbtFolderEvents {
    <#
    .SYNOPSIS
        The progress per folder of a job that runs in the calling thread (PST export, EML export and import): a tracker to hand to the engine
        and a timer that writes the folders that changed as object events of -Step (and, with -StepEvents, the step event of the same id with the sums).
        Returns $null when no event log is open (the engine is then not asked to follow the folders).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Step, [switch]$StepEvents)
    if (-not $script:EventLog) { return $null }
    Initialize-MbtEngine
    $tracker = [MailboxTransfer.FolderTracker]::new()
    $pump = [MailboxTransfer.FolderEventPump]::new($script:EventLog, $tracker, $Step, 1000, $(if ($StepEvents) { $Step }))
    [pscustomobject]@{ Tracker = $tracker; Pump = $pump }
}

function Stop-MbtFolderEvents {
    <# Writes the last changes and stops the timer (safe on $null). #>
    [CmdletBinding()]
    param($Handle)
    if ($Handle -and $Handle.Pump) { $Handle.Pump.Dispose() }
}

function Write-MbtPstExportEvent {
    <# Maps a progress of the PST export (phase plan, write, check, manifest, done) to the steps pst and pstcheck. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Progress)
    if (-not $script:EventLog -or -not $script:JobInfo -or -not $script:JobInfo.Steps.ContainsKey('pst')) { return }
    switch ($Progress.Phase) {
        'write' { Write-MbtJobStep -Id pst -State running -Done $Progress.ItemsDone -Total $Progress.ItemsTotal -Bytes $Progress.BytesDone -BytesTotal $Progress.BytesTotal -Msg ("part {0}" -f $Progress.PartIndex) }
        'check' { if ($script:JobInfo.Steps.ContainsKey('pstcheck')) { Write-MbtJobStep -Id pstcheck -State running -Bytes $Progress.BytesDone -BytesTotal $Progress.BytesTotal -Msg ("part {0}" -f $Progress.PartIndex) } }
        'done' {
            Write-MbtJobStep -Id pst -State done -Done $Progress.ItemsDone -Total $Progress.ItemsTotal -Bytes $Progress.BytesDone -BytesTotal $Progress.BytesTotal
            if ($script:JobInfo.Steps.ContainsKey('pstcheck')) { Write-MbtJobStep -Id pstcheck -State done }
        }
    }
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA1Te8LThY0Rqu6
# Hd+NO4isWh+TLiq08otxxfuMDre8lqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA790mNJVBd
# wdjj2QMWm9LbODbSFyOI0U9E2qgvzRQ7dzANBgkqhkiG9w0BAQEFAASCAgAXEOPl
# CclTM1sbZZjIod1YuttBr4jThR+Dstp7lLBOHvwLN34gPzQa/OnMnHHQc7sW6T8b
# ZYlg1TtIi3dKw+7nm9wAkwHw+Cge5e8Knjtxp2sTwl437SwxQa5DTWaVSZt4c1+1
# 5B/uKVgkBAB9Uc4/Mt7hYa219tPh/fNWas2ML+VsEnIV80bDo8YLAsmRE0NUObcb
# eaBXeQbFhLz58D5w1fRwagQuSDLP8fuyCVrQxN+jmKx7pemLWW9oCFGNkNonOngr
# PCJQU8R8oOjt5FnSdxhA9CrOoDMwiXvUX/x7XdAGkuXMsDRmrGVJ4bVCMRDfn5bQ
# +fYJ772yM4Zm5Ce5BAEhAo1Rn96ezRCyQYuC4DHk514RMXb7qJhNzeVBwQuNeB41
# YsdHjkAuRh70vCagKXY0X8cp0LhgdoBVuu6x0rzjSippHenjaQBGL1P8dj9ZE+8z
# aJhCs+4++fU4VqIRuqgwqtO4nZB4G+7A0bKRm1e1XRSL2YXGFEJTPHK3+tc0oked
# ywhsdLK+xGEKYnFrLRyKx59uS1ZoUCGXbxDybKHgVrmW93pDr4Ye5fJssOFF2zry
# SbH5yS62vDYwLcaE455OJxhXyjOkiNpbV6WWGnKt11V53WH1hcNWBafjNsb7Qp8H
# gYBlDgiNXQsnPIgTKGKL3y2Ma8PDHann8ZKu66GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIBzdyQOvnnW0IEhxvrbfK64xHMYCwKzWrBfTYHnMT5+FAghi30SfjFg5
# axgPMjAyNjExMTAwODIxNTVaMAMCAQECEDN6srsTstzIf/hLcLMxRUCgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjE1NVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIIeB
# Hz8nVcJ4sCUmNJJjSjDUUIZqs/t2RxgiHyOJ0NDDMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEAwSBW0wPRFb8c3RYPerl5/JbkEJ3/Rcmf2gMbdNXthxACIAU+
# pchmbo49TlXDEDtKD3B6i1NXULsjpKt+U6ynx58n
# SIG # End signature block
