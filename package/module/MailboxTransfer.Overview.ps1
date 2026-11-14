#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer - the overview of all the jobs of a work folder (dot-sourced by MailboxTransfer.psm1 and by MailboxTransfer.Gui.Jobs.ps1).

.DESCRIPTION
    <work folder>\jobs\overview.html: one row for each job (started, task, mailbox or source, destination, result, items done and failed, data, duration, a link to the
    report of the job), tiles with the totals, the jobs that failed, stopped or wait for a decision first, filters by result, task and period, a search, sortable columns.
    It is built from the status.json of every job folder only (and the counts of report.json when status.json has none): the events of a job are never read, so a folder of
    thousands of jobs is read in seconds. The jobs of a plan check (task "...-plan") are left out. The page is one self-contained file (templates\Overview.template.html):
    no network, no other file; the data is JSON in the page and every text of a job is put in the page as text, never as markup.

        Get-MbtJobOverviewModel   what the overview says, as data
        Write-MbtJobOverview      overview.html in the jobs folder; returns the path. Two jobs that end at the same time never corrupt it

    The command line and the console rewrite it at the end of every job and after the job folders were purged (Invoke-MbtJobRetention).

.NOTES
    Author  : Nicolas Fabert
#>

# The result of a job, in the order the page shows them when it starts: what needs a look first.
$script:MbtOverviewResults = @(
    @{ Code = 0; Name = 'Complete'; Tone = 'ok' }
    @{ Code = 1; Name = 'Failed'; Tone = 'fail' }
    @{ Code = 2; Name = 'Not complete'; Tone = 'warn' }
    @{ Code = 3; Name = 'Needs a decision'; Tone = 'warn' }
    @{ Code = 4; Name = 'Interrupted'; Tone = 'fail' }
    @{ Code = 5; Name = 'Running'; Tone = 'info' }
    @{ Code = 6; Name = 'Items refused'; Tone = 'warn' }
    @{ Code = 7; Name = 'Unknown'; Tone = 'wait' }
)
$script:MbtOverviewRank = @{ 1 = 0; 3 = 1; 4 = 2; 2 = 3; 6 = 4; 5 = 5; 7 = 6; 0 = 7 }
$script:MbtOverviewColumns = @('id', 'started', 'seconds', 'task', 'source', 'destination', 'result', 'code', 'done', 'failed', 'bytes', 'link', 'batch')

function Read-MbtOverviewStatus {
    <#
        status.json (or the counts of report.json) as a hashtable of the few fields the overview needs, or $null when the file is missing, empty or not JSON.
        Read with System.Text.Json: thousands of files in a second or two.
    #>
    param([Parameter(Mandatory)][string]$Path, [switch]$CountsOnly)
    $text = $null
    try { $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8) } catch { return $null }
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $doc = $null
    try { $doc = [Text.Json.JsonDocument]::Parse($text) } catch { return $null }
    try {
        $root = $doc.RootElement
        if ($root.ValueKind -ne [Text.Json.JsonValueKind]::Object) { return $null }
        $s = @{ counts = @{} }
        foreach ($p in $root.EnumerateObject()) {
            $kind = $p.Value.ValueKind
            switch ($p.Name) {
                'counts' {
                    if ($kind -eq [Text.Json.JsonValueKind]::Object) {
                        foreach ($c in $p.Value.EnumerateObject()) {
                            if ($c.Name -in 'done', 'failed', 'bytes', 'fileBytes') { $n = [long]0; if ($c.Value.ValueKind -eq [Text.Json.JsonValueKind]::Number -and $c.Value.TryGetInt64([ref]$n)) { $s.counts[$c.Name] = $n } }
                        }
                    }
                }
                { $_ -in 'jobId', 'task', 'title', 'mode', 'store', 'source', 'destination', 'state', 'next', 'reason', 'started', 'updated', 'lastEvent', 'batch' } {
                    if ($CountsOnly) { break }
                    if ($kind -eq [Text.Json.JsonValueKind]::String) { $s[$p.Name] = $p.Value.GetString() }
                }
                'exitCode' { $n = [long]0; if ($kind -eq [Text.Json.JsonValueKind]::Number -and $p.Value.TryGetInt64([ref]$n)) { $s.exitCode = [int]$n } }
                'pid' { $n = [long]0; if ($kind -eq [Text.Json.JsonValueKind]::Number -and $p.Value.TryGetInt64([ref]$n)) { $s.pid = [int]$n } }
            }
        }
        return $s
    } finally { $doc.Dispose() }
}

function ConvertTo-MbtOverviewSeconds {
    <# A time of status.json (UTC text) as seconds since 1970, or $null. #>
    param($Value)
    $t = ConvertTo-MbtReportTime $Value
    if ($null -eq $t) { return $null }
    return [long][math]::Floor(([datetime]$t - [datetime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)).TotalSeconds)
}

function ConvertFrom-MbtOverviewJsonText {
    <# The JSON texts [Start, Start + Count) of a list, as hashtables, in one call; when one of them is not JSON the range is cut in two, so a damaged file costs a few small parses and not a slow one for each file. #>
    param([Parameter(Mandatory)][Collections.Generic.List[string]]$Texts, [Parameter(Mandatory)][int]$Start, [Parameter(Mandatory)][int]$Count, [Parameter(Mandatory)][hashtable]$Into, [Parameter(Mandatory)][int[]]$At)
    if ($Count -lt 1) { return }
    $parsed = $null
    try { $parsed = ('[' + [string]::Join(',', $Texts, $Start, $Count) + ']') | ConvertFrom-Json -AsHashtable -Depth 10 -NoEnumerate } catch { $parsed = $null }
    if ($null -ne $parsed -and $parsed.Count -eq $Count) {
        for ($k = 0; $k -lt $Count; $k++) { if ($parsed[$k] -is [hashtable]) { $Into[$At[$Start + $k]] = $parsed[$k] } }
        return
    }
    if ($Count -eq 1) { return }
    $half = [int][math]::Floor($Count / 2)
    ConvertFrom-MbtOverviewJsonText -Texts $Texts -Start $Start -Count $half -Into $Into -At $At
    ConvertFrom-MbtOverviewJsonText -Texts $Texts -Start ($Start + $half) -Count ($Count - $half) -Into $Into -At $At
}

function Read-MbtOverviewStatusBlock {
    <# The status.json of a block of job folders as hashtables, by position in the block (a file that is missing, empty or not JSON has none). #>
    param([Parameter(Mandatory)][string[]]$Folders)
    $texts = [Collections.Generic.List[string]]::new($Folders.Count)
    $at = [Collections.Generic.List[int]]::new($Folders.Count)
    for ($i = 0; $i -lt $Folders.Count; $i++) {
        $text = $null
        try { $text = [IO.File]::ReadAllText($Folders[$i] + '\status.json') } catch { continue }
        if ($text.Trim().Length -lt 2) { continue }
        $texts.Add($text); $at.Add($i)
    }
    $out = @{}
    if ($texts.Count) { ConvertFrom-MbtOverviewJsonText -Texts $texts -Start 0 -Count $texts.Count -Into $out -At $at.ToArray() }
    return $out
}

function ConvertTo-MbtOverviewRowJson {
    <# One row as a JSON array: the strings are escaped by System.Text.Json (< > & ' and every character above ASCII become \uXXXX, so a text can never close a script element). #>
    param([Parameter(Mandatory)][object[]]$Row, [Parameter(Mandatory)][int[]]$TextColumns)
    $copy = [object[]]$Row.Clone()
    foreach ($c in $TextColumns) { $copy[$c] = '"' + [Text.Json.JsonEncodedText]::Encode([string]$Row[$c]).ToString() + '"' }
    return '[' + [string]::Join(',', $copy) + ']'
}

function Get-MbtJobOverviewModel {
    <#
    .SYNOPSIS
        What the overview of a work folder says, as data: the rows (newest first), the strings they share, the totals.
    .DESCRIPTION
        A job folder that has not changed since the last time (its last write time is the same) is not read again: its row comes from jobs\overview.cache.json, which
        Write-MbtJobOverview keeps. A job that is running is always read. Without the cache every job is read, and the result is the same.
    .PARAMETER WorkFolder
        The work folder: its jobs are in <WorkFolder>\jobs.
    .PARAMETER Batch
        Only the jobs of this batch (the "batch" field of status.json). A job without one is never listed when a batch is asked for.
    .PARAMETER JobFolderName
        Only these job folders (the names, as in the jobs folder): a batch knows its jobs, so the other jobs of the work folder are not looked at.
    .PARAMETER MaxJobs
        The page lists the newest jobs only; Left says how many older ones were left out.
    .PARAMETER TimeBudgetSeconds
        Stops reading job folders after this many seconds (0 = no limit). What was read is kept for the next build (Incomplete is true and Unread says how many folders are left), so a folder that
        is slow to read (a virus scanner looks at every file the first time) is finished by several short builds.
    .PARAMETER NoCache
        Reads every job folder.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkFolder, [string]$Batch, [string[]]$JobFolderName, [ValidateRange(1, 1000000)][int]$MaxJobs = 20000, [ValidateRange(0, 86400)][int]$TimeBudgetSeconds = 0, [switch]$NoCache)
    Set-StrictMode -Version Latest
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $jobsFolder = Join-Path $WorkFolder 'jobs'
    $strings = [Collections.Generic.List[string]]::new()
    $index = [Collections.Hashtable]::new([StringComparer]::Ordinal)
    $live = $null
    $nowTicks = [datetime]::UtcNow.Ticks
    $epochTicks = 621355968000000000L
    $nameStart = [regex]::new('^(\d{4})(\d{2})(\d{2})-(\d{2})(\d{2})-(.+)-[0-9a-fA-F]{4}$')
    $batchText = if ($Batch) { $Batch.Trim() } else { '' }
    $resultMemo = @{}
    $taskMemo = @{}
    $none = @{ counts = @{} }

    # what the last build knew: a row for each job folder (the first value is the last write time of the folder), kept for the folders that have not changed since
    $cache = @{}
    $cacheFile = Join-Path $jobsFolder 'overview.cache.json'
    if (-not $NoCache -and [IO.File]::Exists($cacheFile)) {
        try {
            $loaded = [IO.File]::ReadAllText($cacheFile) | ConvertFrom-Json -AsHashtable -Depth 4 -NoEnumerate
            if ($loaded -is [hashtable] -and $loaded['format'] -eq 1) { foreach ($r in $loaded['rows']) { $cache[[string]$r[1]] = $r } }
        } catch { $cache = @{} }
    }
    $cacheRows = [Collections.Generic.List[object]]::new()
    $cacheDirty = $false
    $fresh = 0

    $dirs = [Collections.Generic.List[IO.DirectoryInfo]]::new()
    $only = $null
    if ($PSBoundParameters.ContainsKey('JobFolderName') -and $null -ne $JobFolderName) {
        $only = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($n in $JobFolderName) { if ($n) { [void]$only.Add([string]$n) } }
        foreach ($n in $only) { $d = [IO.DirectoryInfo]::new([IO.Path]::Combine($jobsFolder, $n)); if ($d.Exists) { $dirs.Add($d) } }
    }
    elseif (Test-Path -LiteralPath $jobsFolder -PathType Container) { foreach ($d in [IO.DirectoryInfo]::new($jobsFolder).EnumerateDirectories()) { $dirs.Add($d) } }
    $todo = [Collections.Generic.List[IO.DirectoryInfo]]::new()
    $known = @{}
    foreach ($d in $dirs) {
        $c = $cache[$d.Name]
        if ($null -ne $c -and $c[0] -eq $d.LastWriteTimeUtc.Ticks) { $known[$d.Name] = $c; $cacheRows.Add($c) } else { $todo.Add($d) }
    }
    if ($only) {
        # the other folders were not looked at: what the last build knew of them stays in the cache
        foreach ($k in $cache.Keys) { if (-not $only.Contains([string]$k)) { $cacheRows.Add($cache[$k]) } }
        if ($todo.Count -ne 0) { $cacheDirty = $true }
    }
    elseif ($todo.Count -ne 0 -or $cache.Count -ne $known.Count) { $cacheDirty = $true }
    $incomplete = $false
    $unread = 0
    $blockSize = if ($TimeBudgetSeconds -gt 0) { 50 } else { 250 }

    for ($block = 0; $block -lt $todo.Count; $block += $blockSize) {
        if ($TimeBudgetSeconds -gt 0 -and $clock.Elapsed.TotalSeconds -ge $TimeBudgetSeconds) { $incomplete = $true; $unread = $todo.Count - $block; break }
        $last = [math]::Min($block + $blockSize - 1, $todo.Count - 1)
        $names = [string[]]@($todo[$block..$last] | ForEach-Object { $_.FullName })
        $statuses = Read-MbtOverviewStatusBlock -Folders $names
        for ($n = 0; $n -lt $names.Length; $n++) {
            $dir = $names[$n]
            $name = $todo[$block + $n].Name
            $stamp = $todo[$block + $n].LastWriteTimeUtc.Ticks
            $status = $statuses[$n]
            $hasStatus = $null -ne $status
            if (-not $hasStatus) { $status = $none }
            $fromName = $nameStart.Match($name)
            $task = $status['task']
            if ($null -eq $task) { $task = if ($fromName.Success) { $fromName.Groups[6].Value } else { '' } }
            $task = [string]$task
            # the check of a plan is a job of its own that is not a transfer: its task says so, or the name of its folder, or its id
            $isPlan = $task.EndsWith('-plan', [StringComparison]::OrdinalIgnoreCase) -or ($fromName.Success -and $fromName.Groups[6].Value.EndsWith('-plan', [StringComparison]::OrdinalIgnoreCase)) -or ([string]$status['jobId']) -like '*-plan-*'
            if ($isPlan) { $cacheRows.Add(@($stamp, $name)); continue }
            $mode = [string]$status['mode']

            # the times, in seconds since 1970: a date of status.json is read as a date or as text
            $startTicks = 0L
            $v = $status['started']
            if ($v -is [datetime]) { $startTicks = $v.ToUniversalTime().Ticks } elseif ($v -is [string] -and $v) { $t = ConvertTo-MbtReportTime $v; if ($t) { $startTicks = $t.Ticks } }
            if (-not $startTicks) {
                if ($fromName.Success) { try { $startTicks = [datetime]::new([int]$fromName.Groups[1].Value, [int]$fromName.Groups[2].Value, [int]$fromName.Groups[3].Value, [int]$fromName.Groups[4].Value, [int]$fromName.Groups[5].Value, 0, [DateTimeKind]::Local).ToUniversalTime().Ticks } catch { $startTicks = 0L } }
                if (-not $startTicks) { $startTicks = [IO.Directory]::GetCreationTimeUtc($dir).Ticks }
            }
            $endTicks = 0L
            $v = $status['updated']; if ($null -eq $v) { $v = $status['lastEvent'] }
            if ($v -is [datetime]) { $endTicks = $v.ToUniversalTime().Ticks } elseif ($v -is [string] -and $v) { $t = ConvertTo-MbtReportTime $v; if ($t) { $endTicks = $t.Ticks } }
            if (-not $endTicks) { $endTicks = $startTicks }
            $started = [long][math]::Floor(($startTicks - $epochTicks) / 10000000)

            # the result: what status.json says; a job that says "running" with no process behind it never wrote its end
            $code = 7; $text = 'Unknown'
            $state = [string]$status['state']
            if ($state -eq 'running') {
                $alive = $false
                $processId = $status['pid']
                if ($null -ne $processId) {
                    if ($null -eq $live) { $live = @{}; foreach ($p in [Diagnostics.Process]::GetProcesses()) { $live[[int]$p.Id] = $p } }
                    $proc = $live[[int]$processId]
                    if ($proc) { try { $alive = $proc.StartTime.ToUniversalTime().Ticks -le $startTicks + 300000000L } catch { $alive = $true } }
                }
                if ($alive) { $code = 5; $text = 'Running'; $endTicks = $nowTicks } else { $code = 4; $text = 'Interrupted' }
            } elseif ($state) {
                $exit = $status['exitCode']
                if ($null -eq $exit) { $exit = switch ($state) { 'complete' { 0 } 'error' { 1 } 'needs-decision' { 3 } 'resumable' { 2 } 'stopped' { 2 } default { -1 } } }
                $exit = [int]$exit
                if ($exit -ge 0) {
                    $reason = [string]$status['reason']
                    $memo = "$exit|$reason|$($status['next'])"
                    $result = $resultMemo[$memo]
                    if ($null -eq $result) {
                        $verdict = Get-MbtReportVerdict -Result @{ Exit = $exit; Reason = $reason; Next = [string]$status['next'] } -Interrupted $false
                        $result = @{ Text = [string]$verdict.Text; Code = $(switch ($exit) { 0 { 0 } 1 { if ($reason -eq 'refused-items') { 6 } else { 1 } } 2 { 2 } 3 { 3 } default { 1 } }) }
                        $resultMemo[$memo] = $result
                    }
                    $code = $result.Code; $text = $result.Text
                }
            }
            $seconds = if ($hasStatus -and $code -ne 7) { [math]::Max(0L, [long][math]::Floor(($endTicks - $startTicks) / 10000000)) } else { 0L }

            $done = -1L; $failed = -1L; $bytes = -1L
            $counts = $status['counts']
            if ($counts -isnot [hashtable]) { $counts = @{} }
            if ($hasStatus -and $code -ne 7 -and -not $counts.ContainsKey('done')) {
                $reportFile = $dir + '\report.json'
                if ([IO.File]::Exists($reportFile) -and ([IO.FileInfo]::new($reportFile)).Length -lt 2MB) {
                    $fromReport = Read-MbtReportJsonFile -Path $reportFile
                    if ($fromReport -and $fromReport['counts'] -is [hashtable]) { $counts = $fromReport['counts'] }
                }
            }
            $v = $counts['done']; if ($null -ne $v) { $done = [long]$v }
            $v = $counts['failed']; if ($null -ne $v) { $failed = [long]$v }
            $v = $counts['bytes']; if ($null -eq $v) { $v = $counts['fileBytes'] }
            if ($null -ne $v) { $bytes = [long]$v }

            $link = if ([IO.File]::Exists($dir + '\report.html')) { 1 } elseif ([IO.File]::Exists($dir + '\report.csv')) { 2 } else { 0 }
            $source = [string]$status['source']; $destination = [string]$status['destination']
            if (-not $source -and -not $destination) { $ends = Get-MbtJobEndpoint -Title ([string]$status['title']) -Mode $mode -Store ([string]$status['store']) -Source '' -Destination ''; $source = [string]$ends.Source; $destination = [string]$ends.Destination }
            $taskKey = "$task|$mode"
            $taskName = $taskMemo[$taskKey]
            if ($null -eq $taskName) { $taskName = Get-MbtReportTaskName -Task $task -Mode $mode; $taskMemo[$taskKey] = $taskName }
            $row = @($stamp, $name, $started, $seconds, $taskName, $source, $destination, $text, $code, $done, $failed, $bytes, $link, ([string]$status['batch']))
            # a job that is running changes without its folder changing: it is read again every time
            if ($code -ne 5) { $cacheRows.Add($row) }
            $known[$name] = $row
        }
    }
    if ($live) { foreach ($p in $live.Values) { $p.Dispose() } }

    # the rows of the page: the strings that rows share are written once
    $rows = [Collections.Generic.List[object]]::new()
    $keys = [Collections.Generic.List[long]]::new()
    foreach ($d in $dirs) {
        $r = $known[$d.Name]
        if ($null -eq $r -or $r.Count -lt 3) { continue }
        $jobBatch = [string]$r[13]
        if ($batchText -and $jobBatch -ine $batchText) { continue }
        $row = [object[]]::new(13)
        $row[0] = $r[1]; $row[1] = [long]$r[2]; $row[2] = [long]$r[3]
        $i = 3
        foreach ($s in $r[4], $r[5], $r[6], $r[7]) {
            $at = $index[[string]$s]
            if ($null -eq $at) { $at = $strings.Count; $strings.Add([string]$s); $index[[string]$s] = $at }
            $row[$i++] = $at
        }
        $row[7] = [int]$r[8]; $row[8] = [long]$r[9]; $row[9] = [long]$r[10]; $row[10] = [long]$r[11]; $row[11] = [int]$r[12]
        $row[12] = -1
        if ($jobBatch) { $at = $index[$jobBatch]; if ($null -eq $at) { $at = $strings.Count; $strings.Add($jobBatch); $index[$jobBatch] = $at }; $row[12] = $at }
        $rows.Add($row)
        $keys.Add([long]$r[2])
    }
    # newest first (the index is sorted: Array.Sort does not carry rows that are arrays along with their keys)
    $n = $rows.Count
    $order = $keys.ToArray(); $index = [int[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) { $index[$i] = $i }
    [Array]::Sort($order, $index)
    $items = [object[]]::new($n)
    for ($i = 0; $i -lt $n; $i++) { $items[$i] = $rows[$index[$n - 1 - $i]] }
    $left = [math]::Max(0, $items.Count - $MaxJobs)
    if ($left -gt 0) { $items = $items[0..($MaxJobs - 1)] }

    $totals = [ordered]@{ Jobs = $items.Count; Done = 0L; Failed = 0L; Bytes = 0L; ByResult = @{} }
    foreach ($r in $script:MbtOverviewResults) { $totals.ByResult[[int]$r.Code] = 0 }
    foreach ($r in $items) {
        $totals.ByResult[[int]$r[7]]++
        if ($r[8] -gt 0) { $totals.Done += $r[8] }
        if ($r[9] -gt 0) { $totals.Failed += $r[9] }
        if ($r[10] -gt 0) { $totals.Bytes += $r[10] }
    }
    return [ordered]@{
        JobsFolder = $jobsFolder; Batch = $batchText; Columns = $script:MbtOverviewColumns; Strings = $strings.ToArray(); Rows = $items; Left = $left; MaxJobs = $MaxJobs; Totals = $totals
        Incomplete = $incomplete; Unread = $unread; Folders = $dirs.Count; LinkBase = './'; Caption = ''
        CacheRows = $cacheRows; CacheDirty = ($cacheDirty -and -not $NoCache)
    }
}

function Write-MbtOverviewFile {
    <# Writes a text file beside its name and moves it into place (replacing the older one); a reader that holds the file for a moment is waited for. #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Text)
    $temporary = '{0}.{1}.{2}.tmp' -f $Path, $PID, [guid]::NewGuid().ToString('N').Substring(0, 8)
    [IO.File]::WriteAllText($temporary, $Text, [Text.UTF8Encoding]::new($false))
    $attempt = 0
    while ($true) {
        $attempt++
        try { [IO.File]::Move($temporary, $Path, $true); return }
        catch [System.IO.IOException], [System.UnauthorizedAccessException] {
            if ($attempt -ge 8) { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue; throw }
            Start-Sleep -Milliseconds (100 * $attempt)
        }
        catch { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue; throw }
    }
}

function ConvertTo-MbtOverviewJson {
    <# The model as compact JSON for the page: strings are escaped by System.Text.Json (< > & ' and every character above ASCII become \uXXXX), so the text can never close the script element. #>
    param([Parameter(Mandatory)]$Model)
    $q = { param([string]$S) '"' + [Text.Json.JsonEncodedText]::Encode($S).ToString() + '"' }
    $sb = [Text.StringBuilder]::new(256 + $Model.Rows.Count * 96)
    [void]$sb.Append('{"format":1,"columns":[' + (($Model.Columns | ForEach-Object { & $q $_ }) -join ',') + '],"results":[')
    [void]$sb.Append((($script:MbtOverviewResults | ForEach-Object { '{"code":' + $_.Code + ',"name":' + (& $q $_.Name) + ',"tone":' + (& $q $_.Tone) + ',"rank":' + $script:MbtOverviewRank[[int]$_.Code] + '}' }) -join ','))
    [void]$sb.Append('],"strings":[' + (($Model.Strings | ForEach-Object { & $q $_ }) -join ',') + '],"left":' + $Model.Left + ',"batch":' + (& $q $Model.Batch) + ',"base":' + (& $q $Model.LinkBase) + ',"rows":[')
    $lines = [string[]]::new($Model.Rows.Count)
    for ($n = 0; $n -lt $lines.Length; $n++) {
        $r = $Model.Rows[$n]
        $lines[$n] = '["' + [Text.Json.JsonEncodedText]::Encode([string]$r[0]).ToString() + '",' + ($r[1..12] -join ',') + ']'
    }
    [void]$sb.Append([string]::Join(',', $lines))
    [void]$sb.Append(']}')
    return $sb.ToString()
}

function ConvertTo-MbtOverviewCacheJson {
    <# What the next build may keep: one row for each job folder that is not running (the first value is the last write time of the folder). #>
    param([Parameter(Mandatory)]$Model)
    $lines = [string[]]::new($Model.CacheRows.Count)
    for ($n = 0; $n -lt $lines.Length; $n++) {
        $r = $Model.CacheRows[$n]
        $lines[$n] = if ($r.Count -lt 3) { '[' + $r[0] + ',"' + [Text.Json.JsonEncodedText]::Encode([string]$r[1]).ToString() + '"]' } else { ConvertTo-MbtOverviewRowJson -Row $r -TextColumns 1, 4, 5, 6, 7, 13 }
    }
    return '{"format":1,"rows":[' + [string]::Join(',', $lines) + ']}'
}

function ConvertTo-MbtJobOverviewHtml {
    <# The page for a model: the template with its markers replaced, in one pass (a text that comes from a job is never taken for a marker). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Model, [string]$TemplatePath, [string]$Version = '')
    Set-StrictMode -Version Latest
    $templateFile = if ($TemplatePath) { $TemplatePath } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'templates\Overview.template.html' }
    $template = if (Test-Path -LiteralPath $templateFile -PathType Leaf) { [IO.File]::ReadAllText($templateFile, [Text.Encoding]::UTF8) }
        else { '<!doctype html><html lang="en"><head><meta charset="UTF-8"><title>{{TITLE}}</title></head><body><h1>{{TITLE}}</h1><p>{{SUBTITLE}}</p><div>{{TILES}}</div>{{CAP}}{{DATA}}<p>{{GENERATED}}</p></body></html>' }
    $enc = { param($V) ConvertTo-MbtHtmlText $V }
    $totals = $Model.Totals
    $tiles = [Text.StringBuilder]::new()
    $addTile = {
        param([string]$Label, [string]$Value, [string]$Class = '', [string]$Note = '')
        [void]$tiles.Append('<div class="tile ' + $Class + '"><div class="tile-label">' + (& $enc $Label) + '</div><div class="tile-value">' + (& $enc $Value) + '</div>' + $(if ($Note) { '<div class="tile-note">' + (& $enc $Note) + '</div>' } else { '' }) + '</div>')
    }
    $by = $totals.ByResult
    $attention = [int]$by[1] + [int]$by[2] + [int]$by[3] + [int]$by[4] + [int]$by[6]
    & $addTile 'Jobs' (Format-MbtReportNumber $totals.Jobs) 'main' $(if ($Model.Batch) { 'batch ' + $Model.Batch } else { '' })
    & $addTile 'Complete' (Format-MbtReportNumber $by[0]) 'ok'
    & $addTile 'Need a look' (Format-MbtReportNumber $attention) $(if ($attention -gt 0) { 'warn' } else { '' }) 'failed, not complete, to decide, interrupted, items refused'
    & $addTile 'Failed' (Format-MbtReportNumber $by[1]) $(if ($by[1] -gt 0) { 'fail' } else { '' })
    if ($by[5] -gt 0) { & $addTile 'Running' (Format-MbtReportNumber $by[5]) 'info' }
    & $addTile 'Items done' (Format-MbtReportNumber $totals.Done) 'ok'
    & $addTile 'Items failed' (Format-MbtReportNumber $totals.Failed) $(if ($totals.Failed -gt 0) { 'fail' } else { '' })
    & $addTile 'Data' (Format-MbtReportBytes ([double]$totals.Bytes))

    $cap = ''
    if ($Model.Caption) { $cap = '<p class="caption">' + (& $enc $Model.Caption) + '</p>' }
    if ($Model.Left -gt 0) { $cap += '<p class="caption">This page lists the ' + (Format-MbtReportNumber $Model.MaxJobs) + ' newest jobs. ' + $(if ($Model.Left -eq 1) { '1 older job is' } else { (Format-MbtReportNumber $Model.Left) + ' older jobs are' }) + ' not listed; their folders are in the jobs folder.</p>' }
    $json = ConvertTo-MbtOverviewJson -Model $Model
    $generated = 'Written ' + (Format-MbtReportTime ([datetime]::UtcNow)) + ' from the status.json of every job folder' + $(if ($Version) { ' by Mailbox Transfer ' + $Version } else { '' }) + '. The times in the table are in the time zone of the browser. The page needs no network.'
    $subtitle = (Format-MbtReportNumber $Model.Totals.Jobs) + $(if ($Model.Totals.Jobs -eq 1) { ' job' } else { ' jobs' }) + ' in ' + $Model.JobsFolder
    $values = @{
        TITLE = $(if ($Model.Batch) { 'Mailbox Transfer: jobs of batch ' + (& $enc $Model.Batch) } else { 'Mailbox Transfer: all jobs' }); HEADING = $(if ($Model.Batch) { 'Jobs of batch ' + (& $enc $Model.Batch) } else { 'All jobs' }); VERSION = (& $enc $Version)
        SUBTITLE = (& $enc $subtitle); TILES = $tiles.ToString(); CAP = $cap; GENERATED = (& $enc $generated)
        DATA = ('<script type="application/json" id="mbt-overview-data">' + $json + '</script>')
    }
    $evaluator = [Text.RegularExpressions.MatchEvaluator] { param($m) $name = $m.Groups[1].Value; if ($values.ContainsKey($name)) { [string]$values[$name] } else { $m.Value } }
    return [regex]::Replace($template, '\{\{([A-Z_]+)\}\}', $evaluator)
}

function Write-MbtJobOverview {
    <#
    .SYNOPSIS
        Writes <WorkFolder>\jobs\overview.html (or -FileName in the jobs folder), from the status.json of every job folder. Returns the path.
        -JobFolder names the folder of one job instead of the work folder (the entry script knows only that one): the work folder is the one that holds its jobs folder; a job that is not in a folder named jobs gives nothing ($null).
    .DESCRIPTION
        Safe to call from several jobs at once: the build is serialised by a named mutex, the page is written beside a temporary name and moved into place, and a reader that
        holds the file for a moment is waited for. The links to the reports of the jobs are relative to the jobs folder, so the page works where it is.
    .PARAMETER Batch
        Only the jobs of this batch (status.json: batch).
    .PARAMETER JobFolderName
        Only these job folders of the jobs folder (the report of a batch lists the jobs it started, whatever the number of other jobs).
    .PARAMETER OutputPath
        The page is written to this file instead of <jobs folder>\<FileName> (the report of a batch is in the folder of the batch); -LinkBase is then the path from that file to the jobs folder, with a slash at its end, for the links to the reports of the jobs.
    .PARAMETER Caption
        A sentence above the table (the batch report says how many mailboxes have not started).
    .PARAMETER TimeBudgetSeconds
        The build stops after this many seconds (0 = no limit). The page is then not written (the one that is there stays), what was read is kept for the next build, and $null is returned; -Info says how many job folders are left.
    .PARAMETER Info
        A hashtable that receives Complete, Unread, Folders and Seconds of the build.
    .PARAMETER NoCache
        Reads every job folder instead of the ones that changed since the last build (jobs\overview.cache.json, which holds what the last build learned of the folders that were not running).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Work')][string]$WorkFolder,
        [Parameter(Mandatory, ParameterSetName = 'Job')][string]$JobFolder,
        [string]$Batch, [string[]]$JobFolderName, [string]$FileName = 'overview.html', [string]$OutputPath, [string]$LinkBase = './', [string]$Caption,
        [string]$TemplatePath, [string]$Version, [ValidateRange(1, 1000000)][int]$MaxJobs = 20000, [ValidateRange(0, 86400)][int]$TimeBudgetSeconds = 0, [hashtable]$Info, [switch]$NoCache
    )
    Set-StrictMode -Version Latest
    if ($PSCmdlet.ParameterSetName -eq 'Job') {
        $parent = Split-Path ([IO.Path]::GetFullPath($JobFolder)) -Parent
        if (-not $parent -or (Split-Path $parent -Leaf) -ine 'jobs') { return $null }
        $WorkFolder = Split-Path $parent -Parent
    }
    if ($FileName -ne [IO.Path]::GetFileName($FileName) -or -not $FileName) { throw "The file name of the overview cannot hold a folder: $FileName" }
    $jobsFolder = [IO.Path]::GetFullPath((Join-Path $WorkFolder 'jobs'))
    [void][IO.Directory]::CreateDirectory($jobsFolder)
    $root = Split-Path $PSScriptRoot -Parent
    if (-not $Version) {
        $productFile = Join-Path $root 'product.json'
        if (Test-Path -LiteralPath $productFile -PathType Leaf) { try { $Version = [string](Get-MbtReportValue ([IO.File]::ReadAllText($productFile) | ConvertFrom-Json -AsHashtable) 'version' '') } catch { $Version = '' } }
    }
    $hash = [BitConverter]::ToString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($jobsFolder.ToLowerInvariant()))).Replace('-', '').Substring(0, 16)
    $mutex = [Threading.Mutex]::new($false, "Local\MailboxTransfer.Overview.$hash")
    $held = $false
    try {
        try { $held = $mutex.WaitOne(30000) } catch [Threading.AbandonedMutexException] { $held = $true }
        $clock = [Diagnostics.Stopwatch]::StartNew()
        $modelArgs = @{ WorkFolder = $WorkFolder; Batch = $Batch; MaxJobs = $MaxJobs; TimeBudgetSeconds = $TimeBudgetSeconds; NoCache = $NoCache }
        if ($PSBoundParameters.ContainsKey('JobFolderName')) { $modelArgs.JobFolderName = $JobFolderName }
        $model = Get-MbtJobOverviewModel @modelArgs
        if ($Info) { $Info.Complete = -not $model.Incomplete; $Info.Unread = $model.Unread; $Info.Folders = $model.Folders; $Info.Seconds = [math]::Round($clock.Elapsed.TotalSeconds, 1) }
        if ($model.CacheDirty) { try { Write-MbtOverviewFile -Path (Join-Path $jobsFolder 'overview.cache.json') -Text (ConvertTo-MbtOverviewCacheJson -Model $model) } catch { Write-Verbose "overview cache not written: $($_.Exception.Message)" } }
        if ($model.Incomplete) {
            Write-MbtLog WARN ('The overview was not finished in {0} s: {1} of {2} job folders are left for the next build' -f $TimeBudgetSeconds, $model.Unread, $model.Folders)
            return $null
        }
        $model.LinkBase = $LinkBase
        $model.Caption = $Caption
        $html = ConvertTo-MbtJobOverviewHtml -Model $model -TemplatePath $TemplatePath -Version $Version
        $target = if ($OutputPath) { [IO.Path]::GetFullPath($OutputPath) } else { Join-Path $jobsFolder $FileName }
        [void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($target))
        Write-MbtOverviewFile -Path $target -Text $html
        return $target
    } finally {
        if ($held) { try { $mutex.ReleaseMutex() } catch { Write-Verbose "overview mutex: $($_.Exception.Message)" } }
        $mutex.Dispose()
    }
}
