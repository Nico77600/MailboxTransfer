#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer - the HTML report of a job (dot-sourced by MailboxTransfer.psm1 and by MailboxTransfer.Gui.Jobs.ps1).

.DESCRIPTION
    Every job leaves report.html in its job folder: what was asked, the result, the numbers, the steps and the chronology of the job (every
    step with its start, end, duration and state, the waits for the service, the sleeps of the PC, the warnings and the errors, in time order),
    one row for each folder, and what to do next. It is built from the files the job already has: events.jsonl (package\docs\MailboxTransfer-Events.md)
    and status.json. The page is one self-contained file (templates\Report.template.html): no network, no other file; every text that comes from
    a job (a folder name, a message of the service) is HTML-encoded.

        Read-MbtJobEvents          the events of a job folder, parsed
        Get-MbtJobReportModel      what the report says, as data
        ConvertTo-MbtJobReportHtml the page for a model
        Write-MbtJobHtmlReport     report.html in the job folder (the other files of the folder are linked from it)

.NOTES
    Author  : Nicolas Fabert
#>

function ConvertTo-MbtHtmlText {
    <# A value as text that is safe inside HTML (element text and attribute values). #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-MbtReportTime {
    <# A time of an event as UTC, or $null. A reader of JSON may hand a date as text or as a date. #>
    param($Value)
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
        return $Value.ToUniversalTime()
    }
    if ($Value -is [string] -and $Value) {
        $parsed = [datetime]::MinValue
        $styles = [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal
        if ([datetime]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture, $styles, [ref]$parsed)) { return $parsed }
    }
    return $null
}

function Format-MbtReportTime {
    <# A UTC time as the local time of this PC. -Short leaves the date out. #>
    param($Utc, [switch]$Short)
    if ($null -eq $Utc) { return '' }
    $local = ([datetime]$Utc).ToLocalTime()
    return $local.ToString($(if ($Short) { 'HH:mm:ss' } else { 'yyyy-MM-dd HH:mm:ss' }), [Globalization.CultureInfo]::InvariantCulture)
}

function Format-MbtReportDuration {
    <# 9.3 s, 5 min 12 s, 2 h 04 min, 1 d 03 h. #>
    param([double]$Seconds)
    if ($Seconds -lt 0) { $Seconds = 0 }
    $en = [Globalization.CultureInfo]::InvariantCulture
    if ($Seconds -lt 10) { return [string]::Format($en, '{0:0.0} s', $Seconds) }
    $total = [long][math]::Round($Seconds)
    if ($total -lt 60) { return "$total s" }
    if ($total -lt 3600) { return ('{0} min {1:00} s' -f [math]::Floor($total / 60), ($total % 60)) }
    if ($total -lt 86400) { return ('{0} h {1:00} min' -f [math]::Floor($total / 3600), [math]::Floor(($total % 3600) / 60)) }
    return ('{0} d {1:00} h' -f [math]::Floor($total / 86400), [math]::Floor(($total % 86400) / 3600))
}

function Format-MbtReportBytes {
    param([double]$Bytes)
    if ($Bytes -lt 0) { return '' }
    $units = 'B', 'KiB', 'MiB', 'GiB', 'TiB'
    $i = 0
    while ($Bytes -ge 1024 -and $i -lt $units.Count - 1) { $Bytes /= 1024; $i++ }
    $en = [Globalization.CultureInfo]::InvariantCulture
    if ($i -eq 0) { return [string]::Format($en, '{0:0} B', $Bytes) }
    return [string]::Format($en, '{0:0.0} {1}', $Bytes, $units[$i])
}

function Format-MbtReportNumber {
    param($Value)
    if ($null -eq $Value -or "$Value" -eq '') { return '' }
    return ([long]$Value).ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function Get-MbtReportValue {
    <# One field of a parsed JSON object (hashtable or object), or $Default. Safe under strict mode. #>
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] } else { return $Default } }
    $p = $Object.PSObject.Properties[$Name]
    if ($p -and $null -ne $p.Value) { return $p.Value }
    return $Default
}

function ConvertTo-MbtReportLong {
    param($Value, [long]$Default = 0)
    if ($null -eq $Value) { return $Default }
    try { return [long][math]::Round([double]$Value) } catch { return $Default }
}

function Read-MbtReportJsonFile {
    <# A JSON file of the job folder as a hashtable, or $null when it is missing or unreadable. #>
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    try {
        $text = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
        if (-not $text.Trim()) { return $null }
        return ($text | ConvertFrom-Json -AsHashtable -Depth 20)
    } catch { return $null }
}

function Read-MbtJobEvents {
    <# The events of events.jsonl in the order they were written. A line that is not JSON is left out. The file may still be open for writing. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    Set-StrictMode -Version Latest
    $list = [Collections.Generic.List[object]]::new()
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return , $list }
    $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
    try {
        $reader = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false), $true)
        $text = $reader.ReadToEnd()
    } finally { $stream.Dispose() }
    $lines = @($text -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ })
    if (-not $lines.Count) { return , $list }
    $parsed = $null
    try { $parsed = @(('[' + ($lines -join ',') + ']') | ConvertFrom-Json -AsHashtable -Depth 20) } catch { $parsed = $null }
    if ($null -ne $parsed -and $parsed.Count -eq 1 -and $parsed[0] -is [System.Collections.IEnumerable] -and $parsed[0] -isnot [System.Collections.IDictionary] -and $parsed[0] -isnot [string]) { $parsed = @($parsed[0]) }
    if ($null -eq $parsed) {
        $parsed = foreach ($line in $lines) { try { $line | ConvertFrom-Json -AsHashtable -Depth 20 } catch { } }
    }
    foreach ($e in @($parsed)) { if ($e -is [System.Collections.IDictionary]) { $list.Add($e) } }
    return , $list
}

function Get-MbtReportVerdict {
    <# The result of a job in a few words: Text, Tone (ok, warn, fail, info), Next (what to do, in plain words). #>
    param($Result, [bool]$Interrupted)
    if ($Interrupted -or $null -eq $Result) {
        return @{ Text = 'Interrupted'; Tone = 'fail'; Sentence = 'The job did not write its end: its process was stopped or the PC lost power.'; Next = 'Run the same job again: it resumes and does only what is left.'; NextTone = 'warn' }
    }
    $exit = [int](Get-MbtReportValue $Result 'Exit' 1)
    $reason = [string](Get-MbtReportValue $Result 'Reason' '')
    $next = [string](Get-MbtReportValue $Result 'Next' '')
    switch ($exit) {
        0 {
            $text = switch ($reason) { 'clean' { 'Clean' } default { 'Complete' } }
            return @{ Text = $text; Tone = 'ok'; Sentence = 'The job finished and nothing is left to do.'; Next = 'Nothing is left to do.'; NextTone = 'ok' }
        }
        1 {
            if ($reason -eq 'refused-items') {
                return @{ Text = 'Finished, items refused'; Tone = 'warn'; Sentence = 'The job finished, but some items were refused for good: trying again cannot change the answer of the service.'; Next = 'Nothing can be run again. The items that were refused are listed below and in the list of items not imported.'; NextTone = 'warn' }
            }
            if ($reason -eq 'errors') {
                return @{ Text = 'Errors found'; Tone = 'fail'; Sentence = 'The job found errors in the file.'; Next = 'Repair the file (console: Repair a PST file; command line: -Mode RepairPst), then check the new file.'; NextTone = 'warn' }
            }
            return @{ Text = 'Failed'; Tone = 'fail'; Sentence = 'An error stopped the job. What was done before it is kept.'; Next = 'Read the error above, fix its cause, then run the same job again. It does only what is left.'; NextTone = 'warn' }
        }
        2 {
            $text = switch ($reason) { 'stopped' { 'Stopped' } 'signin-required' { 'Sign-in needed' } 'warnings' { 'Finished with warnings' } 'errors-remain' { 'Errors remain' } default { 'Not complete' } }
            $sentence = switch ($reason) {
                'stopped' { 'The job was stopped. Nothing was lost.' }
                'signin-required' { 'The sign-in to Microsoft expired. Nothing was lost.' }
                'warnings' { 'The job finished and found warnings.' }
                'errors-remain' { 'The file was repaired as far as possible, but errors remain.' }
                default { 'The job is not complete. Nothing was lost.' }
            }
            $nextText = switch ($next) {
                'signin' { 'Sign in again, then continue the job: it does only what is left.' }
                'repair' { 'Repair the file (console: Repair a PST file; command line: -Mode RepairPst).' }
                'done' { 'Nothing is left to do.' }
                default { 'Continue the job (console: Continue; command line: run the same command again). It does only what is left.' }
            }
            return @{ Text = $text; Tone = 'warn'; Sentence = $sentence; Next = $nextText; NextTone = $(if ($next -eq 'done') { 'ok' } else { 'warn' }) }
        }
        3 {
            return @{ Text = 'Needs your decision'; Tone = 'warn'; Sentence = 'Some items may already exist in the mailbox. They are never sent again by themselves.'; Next = 'Settle the uncertain items (console: Decide; command line: -Mode Status -Unknown -Check), then run the job again.'; NextTone = 'warn' }
        }
        default {
            return @{ Text = "Exit code $exit"; Tone = 'fail'; Sentence = "The job ended with exit code $exit."; Next = 'Read the messages of the job, then run it again.'; NextTone = 'warn' }
        }
    }
}

function Get-MbtJobReportModel {
    <#
    .SYNOPSIS
        What the report of a job folder says, as data: header, result, counts, steps, chronology, folders, items, warnings, errors, decisions.
    .DESCRIPTION
        Reads events.jsonl and status.json of the folder. A job that has no result event (killed, power loss) is reported as interrupted.
        Lists are capped (MaxChronology, MaxFolders, MaxItems, MaxWarnings); Left* says how many rows were left out.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Folder,
        [int]$MaxChronology = 2000, [int]$MaxFolders = 3000, [int]$MaxItems = 500, [int]$MaxWarnings = 500
    )
    Set-StrictMode -Version Latest
    $events = Read-MbtJobEvents -Path (Join-Path $Folder 'events.jsonl')
    $status = Read-MbtReportJsonFile -Path (Join-Path $Folder 'status.json')
    if ($null -eq $status) { $status = @{} }

    $steps = [ordered]@{}
    $stepTitles = @{}
    $chrono = [Collections.Generic.List[hashtable]]::new()
    $folders = [ordered]@{}
    $items = [ordered]@{}
    $warnings = [Collections.Generic.List[hashtable]]::new()
    $errors = [Collections.Generic.List[hashtable]]::new()
    $decisions = [Collections.Generic.List[hashtable]]::new()
    $rates = [Collections.Generic.List[double]]::new()
    $phase = @{}
    $result = $null
    $first = $null; $last = $null
    $waitCount = 0L; $waitSeconds = 0.0
    $job = [string](Get-MbtReportValue $status 'jobId' ''); $task = [string](Get-MbtReportValue $status 'task' ''); $mode = [string](Get-MbtReportValue $status 'mode' '')
    $tickWords = '\d[\d,\.]*'

    $addStep = {
        param([string]$Id, [string]$Title)
        if (-not $steps.Contains($Id)) { $steps[$Id] = [ordered]@{ Id = $Id; Title = $(if ($Title) { $Title } else { $Id }); State = 'waiting'; Started = $null; Finished = $null; Done = 0L; Total = 0L; Bytes = 0L; BytesTotal = 0L; Msg = '' } }
        elseif ($Title -and $steps[$Id].Title -eq $Id) { $steps[$Id].Title = $Title }
    }
    $titleOf = { param([string]$Id) if ($steps.Contains($Id)) { $steps[$Id].Title } else { $Id } }
    $closePhase = {
        param([string]$Id, $At)
        if ($phase.ContainsKey($Id) -and $null -ne $phase[$Id].Row -and $null -eq $phase[$Id].Row.Until) { $phase[$Id].Row.Until = $At }
    }

    foreach ($e in $events) {
        $kind = [string](Get-MbtReportValue $e 'e' '')
        $at = ConvertTo-MbtReportTime (Get-MbtReportValue $e 't')
        if ($null -eq $at) { $at = if ($last) { $last } else { [datetime]::UtcNow } }
        if (-not $first) { $first = $at }
        $last = $at
        switch ($kind) {
            'plan' {
                if (-not $job) { $job = [string](Get-MbtReportValue $e 'job' '') }
                if (-not $task) { $task = [string](Get-MbtReportValue $e 'task' '') }
                if (-not $mode) { $mode = [string](Get-MbtReportValue $e 'mode' '') }
                foreach ($s in @(Get-MbtReportValue $e 'steps' @())) { & $addStep ([string](Get-MbtReportValue $s 'id' '')) ([string](Get-MbtReportValue $s 'title' '')) }
            }
            'step' {
                $id = [string](Get-MbtReportValue $e 'id' '')
                & $addStep $id ''
                $s = $steps[$id]
                $state = [string](Get-MbtReportValue $e 'state' 'running')
                $msg = [string](Get-MbtReportValue $e 'msg' '')
                if ($state -eq 'running' -and -not $s.Started) {
                    $s.Started = $at
                    $chrono.Add(@{ T = $at; Kind = 'step'; Tone = 'info'; Step = $s.Title; Text = "$($s.Title) started"; Until = $null; Count = 0 })
                }
                if ($state -in 'done', 'failed') {
                    if (-not $s.Started) { $s.Started = $at }
                    $s.Finished = $at
                    & $closePhase $id $at
                    $text = "$($s.Title) " + $(if ($state -eq 'done') { 'done' } else { 'failed' }) + $(if ($msg) { ": $msg" } else { '' })
                    $chrono.Add(@{ T = $at; Kind = 'step'; Tone = $(if ($state -eq 'done') { 'ok' } else { 'fail' }); Step = $s.Title; Text = $text; Until = $null; Count = 0 })
                }
                $s.State = $state
                $s.Done = ConvertTo-MbtReportLong (Get-MbtReportValue $e 'done' $s.Done)
                $s.Total = ConvertTo-MbtReportLong (Get-MbtReportValue $e 'total' $s.Total)
                $s.Bytes = ConvertTo-MbtReportLong (Get-MbtReportValue $e 'bytes' $s.Bytes)
                $s.BytesTotal = ConvertTo-MbtReportLong (Get-MbtReportValue $e 'bytesTotal' $s.BytesTotal)
                if ($msg) {
                    $s.Msg = $msg
                    if ($state -eq 'running') {
                        $norm = [regex]::Replace($msg, $tickWords, '#')
                        if ($phase.ContainsKey($id) -and $phase[$id].Norm -eq $norm -and $null -ne $phase[$id].Row) { $phase[$id].Row.Text = $msg }
                        else {
                            & $closePhase $id $at
                            $row = @{ T = $at; Kind = 'phase'; Tone = 'info'; Step = $s.Title; Text = $msg; Until = $null; Count = 0 }
                            $chrono.Add($row)
                            $phase[$id] = @{ Norm = $norm; Row = $row }
                        }
                    }
                }
            }
            { $_ -in 'object', 'item' } {
                $stepId = [string](Get-MbtReportValue $e 'step' '')
                $key = [string](Get-MbtReportValue $e 'key' '')
                $table = if ($kind -eq 'object') { $folders } else { $items }
                $id = "$stepId|$key"
                if (-not $table.Contains($id)) { $table[$id] = [ordered]@{ Step = $stepId; Key = $key; State = ''; Done = 0L; Total = 0L; Bytes = -1L; Msg = ''; Detail = ''; Count = 0L; At = $at } }
                $o = $table[$id]
                $o.State = [string](Get-MbtReportValue $e 'state' 'running')
                $o.Done = ConvertTo-MbtReportLong (Get-MbtReportValue $e 'done' $o.Done)
                $o.Total = ConvertTo-MbtReportLong (Get-MbtReportValue $e 'total' $o.Total)
                $o.Bytes = ConvertTo-MbtReportLong (Get-MbtReportValue $e 'bytes' $o.Bytes) -1
                $o.Msg = [string](Get-MbtReportValue $e 'msg' $o.Msg)
                $o.Detail = [string](Get-MbtReportValue $e 'detail' $o.Detail)
                $o.Count = ConvertTo-MbtReportLong (Get-MbtReportValue $e 'count' $o.Count)
            }
            'wait' {
                $secs = [double](Get-MbtReportValue $e 'seconds' 0)
                $reason = [string](Get-MbtReportValue $e 'reason' 'throttle')
                $waitCount = [math]::Max($waitCount + 0, (ConvertTo-MbtReportLong (Get-MbtReportValue $e 'count' ($waitCount + 1))))
                $waitSeconds = [math]::Max($waitSeconds, [double](Get-MbtReportValue $e 'waited' ($waitSeconds + $secs)))
                $prev = if ($chrono.Count) { $chrono[$chrono.Count - 1] } else { $null }
                if ($prev -and $prev.Kind -eq 'wait' -and $prev.Reason -eq $reason) { $prev.Count++; $prev.Seconds += $secs; $prev.Until = $at.AddSeconds($secs); $prev.Spans.Add(@($at, $at.AddSeconds($secs))) }
                else {
                    $spans = [Collections.Generic.List[object]]::new()
                    $spans.Add(@($at, $at.AddSeconds($secs)))
                    $chrono.Add(@{ T = $at; Kind = 'wait'; Tone = 'warn'; Step = ''; Reason = $reason; Text = ''; Until = $at.AddSeconds($secs); Count = 1; Seconds = $secs; Spans = $spans })
                }
            }
            'rate' { $rates.Add([double](Get-MbtReportValue $e 'gibPerHour' 0)) }
            'warn' {
                $w = @{ T = $at; Code = [string](Get-MbtReportValue $e 'code' ''); Step = [string](Get-MbtReportValue $e 'step' ''); Key = [string](Get-MbtReportValue $e 'key' ''); Msg = [string](Get-MbtReportValue $e 'msg' '') }
                $warnings.Add($w)
                $chrono.Add(@{ T = $at; Kind = 'warn'; Tone = 'warn'; Step = (& $titleOf $w.Step); Text = $w.Msg; Until = $null; Count = 0 })
            }
            'decision' {
                $decisions.Add(@{ T = $at; Key = [string](Get-MbtReportValue $e 'key' ''); Msg = [string](Get-MbtReportValue $e 'msg' ''); Folder = [string](Get-MbtReportValue $e 'folder' '') })
                $chrono.Add(@{ T = $at; Kind = 'decision'; Tone = 'warn'; Step = ''; Text = [string](Get-MbtReportValue $e 'msg' ''); Until = $null; Count = 0 })
            }
            'slept' {
                $from = ConvertTo-MbtReportTime (Get-MbtReportValue $e 'from'); $to = ConvertTo-MbtReportTime (Get-MbtReportValue $e 'to')
                $secs = [double](Get-MbtReportValue $e 'seconds' 0)
                $chrono.Add(@{ T = $(if ($from) { $from } else { $at }); Kind = 'sleep'; Tone = 'info'; Step = ''; Text = ('This PC was asleep for ' + (Format-MbtReportDuration $secs) + '. The job went on when it woke up.'); Until = $(if ($to) { $to } else { $at }); Count = 1; Seconds = $secs })
            }
            'info' {
                $text = [string](Get-MbtReportValue $e 'msg' '')
                if ($text) { $chrono.Add(@{ T = $at; Kind = 'info'; Tone = 'info'; Step = ''; Text = $text; Until = $null; Count = 0 }) }
            }
            'error' {
                $err = @{ T = $at; Code = [string](Get-MbtReportValue $e 'code' ''); Msg = [string](Get-MbtReportValue $e 'msg' ''); Tech = [string](Get-MbtReportValue $e 'tech' '') }
                $errors.Add($err)
                $chrono.Add(@{ T = $at; Kind = 'error'; Tone = 'fail'; Step = ''; Text = $err.Msg; Until = $null; Count = 0 })
            }
            'result' {
                $result = @{ Exit = [int](ConvertTo-MbtReportLong (Get-MbtReportValue $e 'exit' 1) 1); Complete = [bool](Get-MbtReportValue $e 'complete' $false); Counts = (Get-MbtReportValue $e 'counts' @{}); Next = [string](Get-MbtReportValue $e 'next' ''); Reason = [string](Get-MbtReportValue $e 'reason' ''); At = $at }
            }
        }
    }
    foreach ($id in @($phase.Keys)) { if ($phase[$id].Row -and $null -eq $phase[$id].Row.Until) { $phase[$id].Row.Until = $last } }
    foreach ($row in $chrono) { if ($row.Kind -eq 'wait') { $row.Text = if ($row.Reason -eq 'network') { "Microsoft could not be reached; the job tried again ($($row.Count) times, $(Format-MbtReportDuration $row.Seconds) in all)." } elseif ($row.Count -gt 1) { "Microsoft asked to wait because of too many requests: $($row.Count) waits, $(Format-MbtReportDuration $row.Seconds) in all." } else { "Microsoft asked to wait $(Format-MbtReportDuration $row.Seconds) because of too many requests." } } }

    $interrupted = $null -eq $result
    if ($null -eq $first) { $first = ConvertTo-MbtReportTime (Get-MbtReportValue $status 'started') }
    $finished = if ($result) { $result.At } elseif ($last) { $last } else { ConvertTo-MbtReportTime (Get-MbtReportValue $status 'updated') }
    if ($null -eq $first) { $first = $finished }
    if ($null -eq $first) { $first = [datetime]::UtcNow; $finished = $first }
    $seconds = if ($first -and $finished) { [math]::Max(0.0, ($finished - $first).TotalSeconds) } else { 0.0 }
    $verdict = Get-MbtReportVerdict -Result $result -Interrupted $interrupted

    $counts = [ordered]@{}
    $rawCounts = if ($result) { $result.Counts } else { Get-MbtReportValue $status 'counts' @{} }
    if ($rawCounts -is [System.Collections.IDictionary]) { foreach ($k in $rawCounts.Keys) { $counts[[string]$k] = $rawCounts[$k] } }

    # the plan of a selection is shown only when the job has nothing better: the folders it then worked on say the same with their progress
    $folderRows = @($folders.Values)
    if (@($folderRows | Where-Object { $_.Step -ne 'plan' }).Count) { $folderRows = @($folderRows | Where-Object { $_.Step -ne 'plan' }) }
    $stepList = @($steps.Values)
    # a job without counts: the folders say how far it got
    if (-not $counts.Contains('done')) {
        $work = @($folderRows | Where-Object { $_.Step -notin 'plan', '' })
        $sum = 0L; foreach ($w in $work) { $sum += [long]$w.Done }
        $most = 0L; foreach ($s in $stepList) { $most = [math]::Max($most, [long]$s.Done) }
        if ($work.Count) { $counts['done'] = $sum } elseif ($stepList.Count) { $counts['done'] = $most }
    }

    $title = [string](Get-MbtReportValue $status 'title' '')
    if (-not $title) { $title = 'Mailbox Transfer job' }
    $tail = if ($title -match ' - (.+)$') { $Matches[1] } else { '' }
    $store = [string](Get-MbtReportValue $status 'store' '')
    $source = [string](Get-MbtReportValue $status 'source' '')
    $destination = [string](Get-MbtReportValue $status 'destination' '')
    if (-not $source -and -not $destination) {
        switch ($mode) {
            'Export' { $source = $tail; $destination = $store }
            'ExportPst' { $source = $store; $destination = $tail }
            'Import' { $source = $store; $destination = $tail }
            'ImportPst' { $source = 'PST files'; $destination = $tail }
            'ExportEml' { $destination = $tail }
            'ImportEml' { $source = 'EML files'; $destination = $tail }
            'CheckPst' { $source = $tail }
            'RepairPst' { $source = $tail }
        }
    }

    $model = [ordered]@{
        Title = $title; Task = $task; Mode = $mode; JobId = $job; Source = $source; Destination = $destination; Store = $store
        Started = $first; Finished = $finished; Seconds = $seconds; Interrupted = $interrupted
        Result = $(if ($result) { @{ Exit = $result.Exit; Complete = $result.Complete; Next = $result.Next; Reason = $result.Reason } } else { $null })
        Verdict = $verdict
        Counts = $counts
        Waits = @{ Count = $waitCount; Seconds = $waitSeconds }
        Rates = @{ Samples = $rates.Count; MaxGiBPerHour = $(if ($rates.Count) { [math]::Round(($rates | Measure-Object -Maximum).Maximum, 2) } else { 0 }) }
        Steps = $stepList
        Chronology = @($chrono | Select-Object -First $MaxChronology)
        LeftChronology = [math]::Max(0, $chrono.Count - $MaxChronology)
        Folders = @($folderRows | Select-Object -First $MaxFolders)
        LeftFolders = [math]::Max(0, $folderRows.Count - $MaxFolders)
        Items = @($items.Values | Select-Object -First $MaxItems)
        LeftItems = [math]::Max(0, $items.Count - $MaxItems)
        Warnings = @($warnings | Select-Object -First $MaxWarnings)
        LeftWarnings = [math]::Max(0, $warnings.Count - $MaxWarnings)
        Errors = $errors.ToArray()
        Decisions = $decisions.ToArray()
        EventCount = @($events).Count
    }
    return $model
}

function New-MbtReportTimeline {
    <# The picture of the chronology: one bar for each step, one lane for the waits, one for the warnings and errors; a sleep of the PC is a band over every lane. #>
    param([Parameter(Mandatory)]$Model)
    Set-StrictMode -Version Latest
    $en = [Globalization.CultureInfo]::InvariantCulture
    $t0 = [datetime]$Model.Started; $t1 = [datetime]$Model.Finished
    $span = [math]::Max(1.0, ($t1 - $t0).TotalSeconds)
    $width = 1000; $labelWidth = 200; $x0 = $labelWidth; $x1 = 985; $laneHeight = 18; $gap = 8; $top = 6
    $lanes = [Collections.Generic.List[hashtable]]::new()
    foreach ($s in $Model.Steps) { if ($s.Started) { $lanes.Add(@{ Kind = 'step'; Label = $s.Title; Step = $s }) } }
    $waits = @($Model.Chronology | Where-Object { $_.Kind -eq 'wait' })
    $sleeps = @($Model.Chronology | Where-Object { $_.Kind -eq 'sleep' })
    $marks = @($Model.Chronology | Where-Object { $_.Kind -in 'warn', 'error' })
    if ($waits.Count) { $lanes.Add(@{ Kind = 'waits'; Label = 'Waits for the service' }) }
    if ($marks.Count) { $lanes.Add(@{ Kind = 'marks'; Label = 'Warnings and errors' }) }
    if (-not $lanes.Count) { return '<p class="empty">The job wrote no step.</p>' }
    $height = $top + $lanes.Count * ($laneHeight + $gap) + 26
    $xOf = { param([datetime]$T) $f = [math]::Min(1.0, [math]::Max(0.0, ($T - $t0).TotalSeconds / $span)); [math]::Round($x0 + $f * ($x1 - $x0), 1) }
    $num = { param($V) ([double]$V).ToString('0.#', $en) }
    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append('<svg class="timeline" viewBox="0 0 ' + $width + ' ' + $height + '" role="img" aria-label="Chronology of the job: one bar for each step">')
    for ($i = 0; $i -le 4; $i++) {
        $gx = & $num ($x0 + $i * ($x1 - $x0) / 4)
        $at = $t0.AddSeconds($span * $i / 4)
        $anchor = if ($i -eq 0) { 'start' } elseif ($i -eq 4) { 'end' } else { 'middle' }
        [void]$sb.Append('<line class="grid" x1="' + $gx + '" x2="' + $gx + '" y1="' + $top + '" y2="' + ($height - 22) + '"/>')
        [void]$sb.Append('<text x="' + $gx + '" y="' + ($height - 6) + '" text-anchor="' + $anchor + '">' + (ConvertTo-MbtHtmlText (Format-MbtReportTime $at -Short)) + '</text>')
    }
    for ($i = 0; $i -lt $lanes.Count; $i++) {
        $y = $top + $i * ($laneHeight + $gap)
        $lane = $lanes[$i]
        $label = [string]$lane.Label
        if ($label.Length -gt 30) { $label = $label.Substring(0, 29) + '...' }
        [void]$sb.Append('<rect class="lane" x="' + $x0 + '" y="' + $y + '" width="' + ($x1 - $x0) + '" height="' + $laneHeight + '" rx="4"/>')
        [void]$sb.Append('<text x="' + ($labelWidth - 10) + '" y="' + ($y + 13) + '" text-anchor="end">' + (ConvertTo-MbtHtmlText $label) + '</text>')
        switch ($lane.Kind) {
            'step' {
                $s = $lane.Step
                $a = [datetime]$s.Started
                $b = if ($s.Finished) { [datetime]$s.Finished } else { $t1 }
                $bx = & $xOf $a
                $bw = [math]::Max(3.0, (& $xOf $b) - $bx)
                $class = if ($s.State -eq 'done') { 'done' } elseif ($s.State -eq 'failed') { 'failed' } else { 'open' }
                $tip = "$($s.Title): " + (Format-MbtReportTime $a -Short) + ' to ' + $(if ($s.Finished) { Format-MbtReportTime $b -Short } else { 'not finished' }) + ', ' + (Format-MbtReportDuration ($b - $a).TotalSeconds)
                [void]$sb.Append('<rect class="bar ' + $class + '" x="' + (& $num $bx) + '" y="' + $y + '" width="' + (& $num $bw) + '" height="' + $laneHeight + '" rx="4"><title>' + (ConvertTo-MbtHtmlText $tip) + '</title></rect>')
            }
            'waits' {
                foreach ($w in $waits) {
                    foreach ($part in $w.Spans) {
                        $a = [datetime]$part[0]; $b = [datetime]$part[1]
                        $bx = & $xOf $a
                        $bw = [math]::Max(2.0, (& $xOf $b) - $bx)
                        [void]$sb.Append('<rect class="bar wait" x="' + (& $num $bx) + '" y="' + $y + '" width="' + (& $num $bw) + '" height="' + $laneHeight + '" rx="2"><title>' + (ConvertTo-MbtHtmlText ((Format-MbtReportTime $a -Short) + ': waited ' + (Format-MbtReportDuration ($b - $a).TotalSeconds))) + '</title></rect>')
                    }
                }
            }
            'marks' {
                foreach ($m in ($marks | Select-Object -First 300)) {
                    $mx = & $xOf ([datetime]$m.T)
                    $cls = if ($m.Kind -eq 'error') { 'mark-error' } else { 'mark-warn' }
                    $tip = (Format-MbtReportTime $m.T -Short) + ' ' + $m.Text
                    [void]$sb.Append('<polygon class="' + $cls + '" points="' + (& $num ($mx - 5)) + ',' + ($y + 9) + ' ' + (& $num $mx) + ',' + ($y + 1) + ' ' + (& $num ($mx + 5)) + ',' + ($y + 9) + ' ' + (& $num $mx) + ',' + ($y + 17) + '"><title>' + (ConvertTo-MbtHtmlText $tip) + '</title></polygon>')
                }
            }
        }
    }
    foreach ($s in $sleeps) {
        $sx = & $xOf ([datetime]$s.T); $ex = & $xOf ([datetime]$s.Until)
        [void]$sb.Append('<rect class="sleep" x="' + (& $num $sx) + '" y="' + ($top - 2) + '" width="' + (& $num ([math]::Max(3.0, $ex - $sx))) + '" height="' + ($height - 22 - $top + 2) + '"><title>' + (ConvertTo-MbtHtmlText $s.Text) + '</title></rect>')
    }
    [void]$sb.Append('</svg>')
    $legend = '<div class="legend"><span><i style="background:var(--cp-success)"></i>step done</span><span><i style="background:var(--cp-danger)"></i>step failed</span><span><i style="background:var(--cp-warning)"></i>step not finished, wait, warning</span><span><i style="background:var(--cp-violet)"></i>the PC was asleep</span></div>'
    return $sb.ToString() + $legend
}

function Get-MbtReportBadge {
    param([string]$State)
    $tone = switch -Regex ($State) {
        '^(done|ok|complete|clean|exported|imported)$' { 'ok' }
        '^(failed|error|lost|bad)$' { 'fail' }
        '^(running|partial|check|checking|notready|not-ready|refused|warning|warn|decision|unknown)$' { 'warn' }
        '^(planned|waiting|wait)$' { 'wait' }
        default { 'info' }
    }
    $text = switch ($State) { 'running' { 'not finished' } 'waiting' { 'not started' } 'notready' { 'not ready' } 'not-ready' { 'not ready' } default { $State } }
    return '<span class="badge ' + $tone + '">' + (ConvertTo-MbtHtmlText $text) + '</span>'
}

function ConvertTo-MbtJobReportHtml {
    <#
    .SYNOPSIS
        The HTML page of a report model: the template with its markers replaced (one pass; a text that comes from a job is never taken for a marker).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Model, [string]$TemplatePath, [string]$Version = '')
    Set-StrictMode -Version Latest
    $enc = { param($V) ConvertTo-MbtHtmlText $V }
    $en = [Globalization.CultureInfo]::InvariantCulture
    $templateFile = if ($TemplatePath) { $TemplatePath } else { Join-Path (Split-Path $PSScriptRoot -Parent) 'templates\Report.template.html' }
    $template = if (Test-Path -LiteralPath $templateFile -PathType Leaf) { [IO.File]::ReadAllText($templateFile, [Text.Encoding]::UTF8) }
        else { '<!doctype html><html lang="en"><head><meta charset="UTF-8"><title>{{TITLE}}</title></head><body><h1>{{TITLE}}</h1><p>{{SUBTITLE}}</p><div>{{META}}</div><div>{{TILES}}</div>{{NEXT}}{{ERRORS}}{{STEPS}}{{CHRONOLOGY}}{{FOLDERS}}{{WARNINGS}}{{FILES}}<p>{{GENERATED}}</p></body></html>' }
    $verdict = $Model.Verdict
    $started = [datetime]$Model.Started; $finished = [datetime]$Model.Finished
    $offset = [TimeZoneInfo]::Local.GetUtcOffset($started.ToLocalTime())
    $zone = 'UTC' + $(if ($offset -lt [TimeSpan]::Zero) { '-' } else { '+' }) + ([math]::Abs($offset.Hours)).ToString('00') + ':' + ([math]::Abs($offset.Minutes)).ToString('00')

    # header
    $meta = [Text.StringBuilder]::new()
    $metaItems = [ordered]@{
        'Task' = $Model.Title; 'Source' = $Model.Source; 'Destination' = $Model.Destination
        'Store' = $(if ($Model.Store -and $Model.Store -ne $Model.Source -and $Model.Store -ne $Model.Destination) { $Model.Store } else { '' })
        'Started' = (Format-MbtReportTime $started); 'Finished' = $(if ($Model.Interrupted) { 'not written (last event ' + (Format-MbtReportTime $finished) + ')' } else { Format-MbtReportTime $finished })
        'Duration' = (Format-MbtReportDuration $Model.Seconds)
        'Result' = $(if ($Model.Result) { $verdict.Text + ' (exit code ' + $Model.Result.Exit + ')' } else { $verdict.Text })
        'Job' = $Model.JobId
    }
    foreach ($k in $metaItems.Keys) { if ($metaItems[$k]) { [void]$meta.Append('<div><span>' + (& $enc $k) + '</span>' + (& $enc $metaItems[$k]) + '</div>') } }

    # tiles
    $tiles = [Text.StringBuilder]::new()
    $addTile = {
        param([string]$Label, [string]$Value, [string]$Class = '', [string]$Note = '')
        [void]$tiles.Append('<div class="tile ' + $Class + '"><div class="tile-label">' + (& $enc $Label) + '</div><div class="tile-value">' + (& $enc $Value) + '</div>' + $(if ($Note) { '<div class="tile-note">' + (& $enc $Note) + '</div>' } else { '' }) + '</div>')
    }
    & $addTile 'Result' $verdict.Text 'main' $(if ($Model.Result) { 'exit code ' + $Model.Result.Exit } else { 'no end was written' })
    $counts = $Model.Counts
    $spec = @(
        @('done', 'Items done', 'ok', $true), @('recovered', 'Recovered', 'ok', $false), @('itemsWritten', 'Items written', 'ok', $false), @('foldersCreated', 'Folders created', '', $false), @('folders', 'Folders', '', $false),
        @('failed', 'Failed', 'fail', $true), @('refused', 'Refused', 'warn', $false), @('deferred', 'Deferred', 'warn', $false), @('unknown', 'To settle', 'warn', $false), @('pending', 'Left to do', 'warn', $false),
        @('partial', 'Partly recovered', 'warn', $false), @('lost', 'Lost', 'fail', $false), @('orphan', 'Orphans', 'warn', $false), @('errors', 'Errors', 'fail', $false), @('warnings', 'Warnings', 'warn', $false), @('infos', 'Notes', '', $false),
        @('skipped', 'Skipped', '', $false), @('resumed', 'Resumed', '', $false), @('bytes', 'Data', 'bytes', $false), @('fileBytes', 'File size', 'bytes', $false), @('parts', 'PST files', '', $false)
    )
    foreach ($s in $spec) {
        $has = $counts.Contains($s[0])
        if (-not $has -and -not $s[3]) { continue }
        $value = if ($has) { $counts[$s[0]] } else { 0 }
        if ($s[2] -eq 'bytes') { & $addTile $s[1] (Format-MbtReportBytes ([double]$value)) '' ; continue }
        $number = ConvertTo-MbtReportLong $value
        $class = switch ($s[2]) { 'ok' { 'ok' } 'fail' { if ($number -gt 0) { 'fail' } else { '' } } 'warn' { if ($number -gt 0) { 'warn' } else { '' } } default { '' } }
        & $addTile $s[1] (Format-MbtReportNumber $number) $class
    }
    $waitsNote = if ($Model.Waits.Count -gt 0) { (Format-MbtReportDuration $Model.Waits.Seconds) + ' waited' } else { 'the service never asked to wait' }
    & $addTile 'Throttling waits' (Format-MbtReportNumber $Model.Waits.Count) $(if ($Model.Waits.Count -gt 0) { 'warn' } else { '' }) $waitsNote
    & $addTile 'Duration' (Format-MbtReportDuration $Model.Seconds) 'info'

    # next step and errors
    $next = '<section class="card next ' + $verdict.NextTone + '"><h2>What to do next</h2><p>' + (& $enc $verdict.Sentence) + '</p><p><strong>' + (& $enc $verdict.Next) + '</strong></p></section>'
    $errorHtml = ''
    if ($Model.Errors.Count) {
        $sb = [Text.StringBuilder]::new()
        [void]$sb.Append('<section class="card problem"><h2>What went wrong</h2>')
        foreach ($err in $Model.Errors) {
            [void]$sb.Append('<p>' + (& $enc $err.Msg) + $(if ($err.Code) { ' <span class="badge fail">' + (& $enc $err.Code) + '</span>' } else { '' }) + ' <small>' + (& $enc (Format-MbtReportTime $err.T)) + '</small></p>')
            if ($err.Tech) { [void]$sb.Append('<details><summary>Technical message</summary><pre>' + (& $enc $err.Tech) + '</pre></details>') }
        }
        [void]$sb.Append('</section>')
        $errorHtml = $sb.ToString()
    }
    $decisionHtml = ''
    if ($Model.Decisions.Count) {
        $sb = [Text.StringBuilder]::new()
        [void]$sb.Append('<section class="card attention"><h2>Items waiting for a decision<span class="count">' + $Model.Decisions.Count + '</span></h2><ul>')
        foreach ($d in ($Model.Decisions | Select-Object -First 200)) { [void]$sb.Append('<li>' + (& $enc $d.Msg) + $(if ($d.Folder) { ' <small>' + (& $enc $d.Folder) + '</small>' } else { '' }) + '</li>') }
        [void]$sb.Append('</ul></section>')
        $decisionHtml = $sb.ToString()
    }

    # steps
    $sb = [Text.StringBuilder]::new()
    if ($Model.Steps.Count) {
        [void]$sb.Append('<div class="scroll"><table><thead><tr><th>Step</th><th>State</th><th>Started</th><th>Finished</th><th class="num">Duration</th><th class="num">Items</th><th class="num">Data</th><th>Message</th></tr></thead><tbody>')
        foreach ($s in $Model.Steps) {
            $dur = if ($s.Started -and $s.Finished) { Format-MbtReportDuration (([datetime]$s.Finished) - ([datetime]$s.Started)).TotalSeconds } elseif ($s.Started) { 'not finished' } else { '' }
            $itemsText = if ($s.Total -gt 0) { (Format-MbtReportNumber $s.Done) + ' of ' + (Format-MbtReportNumber $s.Total) } elseif ($s.Done -gt 0) { Format-MbtReportNumber $s.Done } else { '' }
            $dataText = if ($s.BytesTotal -gt 0) { (Format-MbtReportBytes $s.Bytes) + ' of ' + (Format-MbtReportBytes $s.BytesTotal) } elseif ($s.Bytes -gt 0) { Format-MbtReportBytes $s.Bytes } else { '' }
            [void]$sb.Append('<tr><td>' + (& $enc $s.Title) + '</td><td>' + (Get-MbtReportBadge $s.State) + '</td><td class="when">' + (& $enc $(if ($s.Started) { Format-MbtReportTime $s.Started -Short } else { '' })) + '</td><td class="when">' + (& $enc $(if ($s.Finished) { Format-MbtReportTime $s.Finished -Short } else { '' })) + '</td><td class="num">' + (& $enc $dur) + '</td><td class="num">' + (& $enc $itemsText) + '</td><td class="num">' + (& $enc $dataText) + '</td><td>' + (& $enc $s.Msg) + '</td></tr>')
        }
        [void]$sb.Append('</tbody></table></div>')
    } else { [void]$sb.Append('<p class="empty">The job wrote no step.</p>') }
    $stepsHtml = $sb.ToString()

    # chronology
    $sb = [Text.StringBuilder]::new()
    $firstDay = $started.ToLocalTime().Date
    if ($Model.Chronology.Count) {
        [void]$sb.Append('<div class="bar"><div class="chips" role="group" aria-label="Filter the events"><button type="button" class="chip" data-filter="all" aria-pressed="true">All</button><button type="button" class="chip" data-filter="step phase" aria-pressed="false">Steps</button><button type="button" class="chip" data-filter="wait sleep" aria-pressed="false">Waits and sleeps</button><button type="button" class="chip" data-filter="warn decision" aria-pressed="false">Warnings</button><button type="button" class="chip" data-filter="error" aria-pressed="false">Errors</button></div></div>')
        [void]$sb.Append('<div class="scroll"><table id="chrono"><thead><tr><th>Time</th><th class="num">Since the start</th><th>Step</th><th>What happened</th></tr></thead><tbody>')
        foreach ($row in $Model.Chronology) {
            $t = ([datetime]$row.T).ToLocalTime()
            $when = if ($t.Date -ne $firstDay) { Format-MbtReportTime $row.T } else { Format-MbtReportTime $row.T -Short }
            $since = Format-MbtReportDuration (([datetime]$row.T) - $started).TotalSeconds
            $text = [string]$row.Text
            if ($row.Kind -in 'phase', 'wait', 'sleep' -and $row.Until -and ([datetime]$row.Until - [datetime]$row.T).TotalSeconds -ge 1) { $text += ' (' + (Format-MbtReportDuration (([datetime]$row.Until) - ([datetime]$row.T)).TotalSeconds) + ')' }
            $cls = switch ($row.Kind) { 'error' { 'k-error' } { $_ -in 'warn', 'decision' } { 'k-warn' } 'wait' { 'k-wait' } 'sleep' { 'k-sleep' } default { '' } }
            [void]$sb.Append('<tr class="' + $cls + '" data-kind="' + (& $enc $row.Kind) + '"><td class="when">' + (& $enc $when) + '</td><td class="num">' + (& $enc $since) + '</td><td>' + (& $enc $row.Step) + '</td><td>' + (& $enc $text) + '</td></tr>')
        }
        [void]$sb.Append('</tbody></table></div>')
        if ($Model.LeftChronology -gt 0) { [void]$sb.Append('<p class="caption">' + (Format-MbtReportNumber $Model.LeftChronology) + ' more events are in events.jsonl.</p>') }
    } else { [void]$sb.Append('<p class="empty">The job wrote no event.</p>') }
    $chronoHtml = $sb.ToString()

    # folders
    $foldersHtml = ''
    if ($Model.Folders.Count) {
        $sb = [Text.StringBuilder]::new()
        [void]$sb.Append('<section class="card" aria-labelledby="folders-heading"><h2 id="folders-heading">Folders<span class="count">' + (Format-MbtReportNumber ($Model.Folders.Count + $Model.LeftFolders)) + '</span></h2>')
        [void]$sb.Append('<div class="bar"><input type="search" data-search="folders" placeholder="Find a folder or a state" aria-label="Find a folder or a state"></div>')
        [void]$sb.Append('<div class="scroll"><table id="folders"><thead><tr><th>Folder</th><th>Step</th><th>State</th><th class="num">Items</th><th class="num">Size</th><th>Message</th></tr></thead><tbody>')
        foreach ($f in $Model.Folders) {
            $titleOfStep = $Model.Steps | Where-Object { $_.Id -eq $f.Step } | Select-Object -First 1
            $stepText = if ($titleOfStep) { $titleOfStep.Title } elseif ($f.Step -eq 'plan') { 'Plan' } else { $f.Step }
            $itemsText = if ($f.Total -gt 0) { (Format-MbtReportNumber $f.Done) + ' of ' + (Format-MbtReportNumber $f.Total) } elseif ($f.Done -gt 0) { Format-MbtReportNumber $f.Done } else { '0' }
            [void]$sb.Append('<tr><td>' + (& $enc $f.Key) + '</td><td>' + (& $enc $stepText) + '</td><td>' + (Get-MbtReportBadge $f.State) + '</td><td class="num">' + (& $enc $itemsText) + '</td><td class="num">' + (& $enc $(if ($f.Bytes -ge 0) { Format-MbtReportBytes $f.Bytes } else { '' })) + '</td><td>' + (& $enc $f.Msg) + '</td></tr>')
        }
        [void]$sb.Append('</tbody></table></div>')
        if ($Model.LeftFolders -gt 0) { [void]$sb.Append('<p class="caption">' + (Format-MbtReportNumber $Model.LeftFolders) + ' more folders are in events.jsonl.</p>') }
        [void]$sb.Append('</section>')
        $foldersHtml = $sb.ToString()
    }

    # items that need a look
    $itemsHtml = ''
    if ($Model.Items.Count) {
        $sb = [Text.StringBuilder]::new()
        [void]$sb.Append('<section class="card attention" aria-labelledby="items-heading"><h2 id="items-heading">Items and findings that need a look<span class="count">' + (Format-MbtReportNumber ($Model.Items.Count + $Model.LeftItems)) + '</span></h2><div class="scroll"><table><thead><tr><th>Item</th><th>State</th><th>What</th><th>Where</th><th class="num">Count</th></tr></thead><tbody>')
        foreach ($i in $Model.Items) {
            [void]$sb.Append('<tr><td>' + (& $enc $i.Key) + '</td><td>' + (Get-MbtReportBadge $i.State) + '</td><td>' + (& $enc $i.Msg) + '</td><td>' + (& $enc $i.Detail) + '</td><td class="num">' + (& $enc $(if ($i.Count -gt 0) { Format-MbtReportNumber $i.Count } else { '' })) + '</td></tr>')
        }
        [void]$sb.Append('</tbody></table></div>')
        if ($Model.LeftItems -gt 0) { [void]$sb.Append('<p class="caption">' + (Format-MbtReportNumber $Model.LeftItems) + ' more are in the CSV report.</p>') }
        [void]$sb.Append('</section>')
        $itemsHtml = $sb.ToString()
    }

    # warnings
    $warningsHtml = ''
    if ($Model.Warnings.Count) {
        $sb = [Text.StringBuilder]::new()
        [void]$sb.Append('<section class="card attention" aria-labelledby="warnings-heading"><h2 id="warnings-heading">Warnings<span class="count">' + (Format-MbtReportNumber ($Model.Warnings.Count + $Model.LeftWarnings)) + '</span></h2><div class="scroll"><table><thead><tr><th>Time</th><th>Step</th><th>Code</th><th>Item</th><th>Message</th></tr></thead><tbody>')
        foreach ($w in $Model.Warnings) {
            $titleOfStep = $Model.Steps | Where-Object { $_.Id -eq $w.Step } | Select-Object -First 1
            [void]$sb.Append('<tr><td class="when">' + (& $enc (Format-MbtReportTime $w.T -Short)) + '</td><td>' + (& $enc $(if ($titleOfStep) { $titleOfStep.Title } else { $w.Step })) + '</td><td class="mono">' + (& $enc $w.Code) + '</td><td>' + (& $enc $w.Key) + '</td><td>' + (& $enc $w.Msg) + '</td></tr>')
        }
        [void]$sb.Append('</tbody></table></div>')
        if ($Model.LeftWarnings -gt 0) { [void]$sb.Append('<p class="caption">' + (Format-MbtReportNumber $Model.LeftWarnings) + ' more warnings are in events.jsonl.</p>') }
        [void]$sb.Append('</section>')
        $warningsHtml = $sb.ToString()
    }

    # files
    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append('<ul class="files">')
    $folderOfReport = if ($Model.Contains('Folder') -and $Model.Folder) { [string]$Model.Folder } else { $null }
    foreach ($f in @(@('report.csv', 'Report (CSV)'), @('report.json', 'Report (JSON)'), @('events.jsonl', 'Events of the job'), @('status.json', 'Status of the job'), @('stdout.txt', 'Text of the job'))) {
        if ($folderOfReport -and -not (Test-Path -LiteralPath (Join-Path $folderOfReport $f[0]) -PathType Leaf)) { continue }
        [void]$sb.Append('<li><a href="' + (& $enc $f[0]) + '">' + (& $enc $f[1]) + '</a> <small>' + (& $enc $f[0]) + '</small></li>')
    }
    [void]$sb.Append('</ul>')
    $listFile = if ($counts.Contains('report')) { [string]$counts['report'] } else { '' }
    if ($listFile) { [void]$sb.Append('<p class="caption">The list of the items not imported, with the reason of each: <code>' + (& $enc $listFile) + '</code></p>') }
    $filesHtml = $sb.ToString()

    # data for tools
    $data = [ordered]@{
        format = 1; title = $Model.Title; task = $Model.Task; mode = $Model.Mode; job = $Model.JobId; source = $Model.Source; destination = $Model.Destination; store = $Model.Store
        started = $started.ToString('o'); finished = $finished.ToString('o'); seconds = [math]::Round($Model.Seconds, 1); interrupted = $Model.Interrupted
        result = $Model.Result; verdict = $verdict.Text; counts = $Model.Counts; waits = $Model.Waits
        steps = @($Model.Steps | ForEach-Object { [ordered]@{ id = $_.Id; title = $_.Title; state = $_.State; started = $(if ($_.Started) { ([datetime]$_.Started).ToString('o') } else { $null }); finished = $(if ($_.Finished) { ([datetime]$_.Finished).ToString('o') } else { $null }); done = $_.Done; total = $_.Total; bytes = $_.Bytes; msg = $_.Msg } })
        chronology = @($Model.Chronology | ForEach-Object { [ordered]@{ t = ([datetime]$_.T).ToString('o'); kind = $_.Kind; step = $_.Step; text = $_.Text; until = $(if ($_.Until) { ([datetime]$_.Until).ToString('o') } else { $null }) } })
        folders = @($Model.Folders | ForEach-Object { [ordered]@{ step = $_.Step; key = $_.Key; state = $_.State; done = $_.Done; total = $_.Total; bytes = $_.Bytes; msg = $_.Msg } })
        warnings = @($Model.Warnings | ForEach-Object { [ordered]@{ t = ([datetime]$_.T).ToString('o'); code = $_.Code; step = $_.Step; key = $_.Key; msg = $_.Msg } })
        errors = @($Model.Errors | ForEach-Object { [ordered]@{ t = ([datetime]$_.T).ToString('o'); code = $_.Code; msg = $_.Msg; tech = $_.Tech } })
    }
    $json = ConvertTo-Json -InputObject $data -Depth 8 -Compress
    $escape = [string][char]92 + 'u'
    $json = $json.Replace('<', $escape + '003c').Replace('>', $escape + '003e').Replace('&', $escape + '0026').Replace([string][char]0x2028, $escape + '2028').Replace([string][char]0x2029, $escape + '2029')
    $dataHtml = '<script type="application/json" id="mbt-report-data">' + $json + '</script>'

    $generated = 'Written ' + (Format-MbtReportTime ([datetime]::UtcNow)) + ' from events.jsonl and status.json of the job' + $(if ($Version) { ' by Mailbox Transfer ' + $Version } else { '' }) + '. Times are in the time zone of the PC that wrote this page (' + $zone + '). The page needs no network.'
    $subtitle = $verdict.Sentence + ' ' + (Format-MbtReportTime $started)
    $values = @{
        TITLE = (& $enc $Model.Title); VERSION = (& $enc $Version); SUBTITLE = (& $enc $subtitle); META = $meta.ToString(); TILES = $tiles.ToString(); NEXT = $next; ERRORS = $errorHtml; DECISIONS = $decisionHtml
        TIMELINE = (New-MbtReportTimeline -Model $Model); STEPS = $stepsHtml; CHRONOLOGY = $chronoHtml; FOLDERS = $foldersHtml; ITEMS = $itemsHtml; WARNINGS = $warningsHtml
        FILES = $filesHtml; DATA = $dataHtml; GENERATED = (& $enc $generated)
    }
    $evaluator = [Text.RegularExpressions.MatchEvaluator] { param($m) $name = $m.Groups[1].Value; if ($values.ContainsKey($name)) { [string]$values[$name] } else { $m.Value } }
    return [regex]::Replace($template, '\{\{([A-Z_]+)\}\}', $evaluator)
}

function Write-MbtJobHtmlReport {
    <#
    .SYNOPSIS
        Writes report.html in a job folder, from events.jsonl and status.json of that folder. Returns the path of the file.
    .DESCRIPTION
        The page is written beside a temporary name and moved into place, so a reader never meets half a file. The other reports of the folder
        (report.csv, report.json) are linked when they exist.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Folder, [string]$OutputPath, [string]$TemplatePath, [string]$Version)
    Set-StrictMode -Version Latest
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { throw "The job folder does not exist: $Folder" }
    $root = Split-Path $PSScriptRoot -Parent
    if (-not $Version) {
        $productFile = Join-Path $root 'product.json'
        if (Test-Path -LiteralPath $productFile -PathType Leaf) { try { $Version = [string](Get-MbtReportValue ([IO.File]::ReadAllText($productFile) | ConvertFrom-Json -AsHashtable) 'version' '') } catch { $Version = '' } }
    }
    $model = Get-MbtJobReportModel -Folder $Folder
    $model['Folder'] = (Resolve-Path -LiteralPath $Folder).Path
    $html = ConvertTo-MbtJobReportHtml -Model $model -TemplatePath $TemplatePath -Version $Version
    $target = if ($OutputPath) { $OutputPath } else { Join-Path $Folder 'report.html' }
    $temporary = "$target.$PID.tmp"
    [IO.File]::WriteAllText($temporary, $html, [Text.UTF8Encoding]::new($false))
    try { Move-Item -LiteralPath $temporary -Destination $target -Force }
    catch { Remove-Item -LiteralPath $temporary -Force -ErrorAction SilentlyContinue; throw }
    return $target
}
