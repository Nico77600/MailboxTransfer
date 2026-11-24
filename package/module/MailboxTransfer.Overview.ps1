#Requires -Version 7.6
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
$script:MbtOverviewColumns = @('id', 'started', 'seconds', 'task', 'source', 'destination', 'result', 'code', 'done', 'failed', 'bytes', 'link', 'batch', 'outside')

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
                            if ($c.Name -in 'done', 'failed', 'bytes', 'fileBytes', 'outsidePeriod') { $n = [long]0; if ($c.Value.ValueKind -eq [Text.Json.JsonValueKind]::Number -and $c.Value.TryGetInt64([ref]$n)) { $s.counts[$c.Name] = $n } }
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
            if ($loaded -is [hashtable] -and $loaded['format'] -eq 2) { foreach ($r in $loaded['rows']) { $cache[[string]$r[1]] = $r } }
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

            $done = -1L; $failed = -1L; $bytes = -1L; $outside = -1L
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
            $v = $counts['outsidePeriod']; if ($null -ne $v) { $outside = [long]$v }
            $v = $counts['bytes']; if ($null -eq $v) { $v = $counts['fileBytes'] }
            if ($null -ne $v) { $bytes = [long]$v }

            $link = if ([IO.File]::Exists($dir + '\report.html')) { 1 } elseif ([IO.File]::Exists($dir + '\report.csv')) { 2 } elseif ($hasStatus) { 3 } else { 0 }
            $source = [string]$status['source']; $destination = [string]$status['destination']
            if (-not $source -and -not $destination) { $ends = Get-MbtJobEndpoint -Title ([string]$status['title']) -Mode $mode -Store ([string]$status['store']) -Source '' -Destination ''; $source = [string]$ends.Source; $destination = [string]$ends.Destination }
            $taskKey = "$task|$mode"
            $taskName = $taskMemo[$taskKey]
            if ($null -eq $taskName) { $taskName = Get-MbtReportTaskName -Task $task -Mode $mode; $taskMemo[$taskKey] = $taskName }
            $row = @($stamp, $name, $started, $seconds, $taskName, $source, $destination, $text, $code, $done, $failed, $bytes, $link, ([string]$status['batch']), $outside)
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
        $row = [object[]]::new(14)
        $row[0] = $r[1]; $row[1] = [long]$r[2]; $row[2] = [long]$r[3]
        $i = 3
        foreach ($s in $r[4], $r[5], $r[6], $r[7]) {
            $at = $index[[string]$s]
            if ($null -eq $at) { $at = $strings.Count; $strings.Add([string]$s); $index[[string]$s] = $at }
            $row[$i++] = $at
        }
        $row[7] = [int]$r[8]; $row[8] = [long]$r[9]; $row[9] = [long]$r[10]; $row[10] = [long]$r[11]; $row[11] = [int]$r[12]
        $row[12] = -1
        $row[13] = if ($r.Count -gt 14) { [long]$r[14] } else { -1L }
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
        $lines[$n] = '["' + [Text.Json.JsonEncodedText]::Encode([string]$r[0]).ToString() + '",' + ($r[1..13] -join ',') + ']'
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
    return '{"format":2,"rows":[' + [string]::Join(',', $lines) + ']}'
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

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDkYWO7S7wzKV4J
# J0PTpX53eGSygEt+XeiPYvdM0mOQrqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAwMZJFlJ9C
# hXcqtuNIZfLyp1JsSBGlQT9at9wfPPkVoDANBgkqhkiG9w0BAQEFAASCAgA6+5T3
# LcCFusZSc9AAsR1OaLPV2aIdAZs7exm/wDYd0mKQdlyFI/M//C0JGn/Gl+kILaSB
# jNqwF8I+ym/6xJGEe6RfZ+PIHXr9M7EzUTrpK1bVJHGaRRaKjPupLXd/Btq70V5n
# c1nk6m6wE7y2X6LS50E24iyy6y5WSIXSqvnQBm9aF2ux+60QUijeAzUeC8tw+rIy
# GMwHzzzqk6Ewi7t6WaTFpZ/omcyVe441VgKYLOjSHbyPgmmk2AEcQmNqubX6bYuI
# MrPGTLZCYQsPW/HpSrUiH/dEAybgv4OK5zrnwtEhB+TKoUF2X1DpC250R0NdFRXK
# jfMDlKVNuXTOm3AWZMRisSw8sAwT0IuAymCldQRyH+nuIEmarmRB4zS3WQle7U/D
# Mq2U5yIgDQR5TnMD68hpOLTO5TdfGAuJx1hP7Tdfj24UMK7/o9RzCoXWc4j9pSWv
# jtPfbfECjeiSk0p5JsK4m0WkZSs+52bhDj5CHs6GkqnmRfeF14MGwQ83gLksA54S
# RAyBy5bQ6Bb5cS7av9fui2hEW/88N+aNlU/+y4JN6uc7T1sK5yhduV3dbXztfBbY
# aIryREfoHCrEjc+kjdFTVWIKhrYo+pXQ9qu5DV+O+N7OoW14X/SxDqEAMZO6jmeL
# 8zm44E34v+dai94ePRvPpVtEZ4GwLSaIci7yJaGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIEPf8ITB/7+FbU0VvKTvqkPMOWxfyiZUjH8Do1u8uv7LAgguclW6YLi7
# 8hgPMjAyNjExMTUxNDM0NDFaMAMCAQECEHOQmHc+a7OU/tiXHmbv9figghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTE0MzQ0MVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIA0W
# uEJfbisqPyB0kmXkZwhw9fu/pPWPo4x6u51puSg2MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiBvGWyvcYWHxn2MhF6foWUuH9sLBjtr3ws04X2KdoyHnAIgHFjV
# SEtTEZn51ealjQ7vGRihvmXoV+7qTU4/e92D/Fo=
# SIG # End signature block
