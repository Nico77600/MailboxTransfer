#Requires -Version 7.6
<#
.SYNOPSIS
    Mailbox Transfer - the HTML report of a job (dot-sourced by MailboxTransfer.psm1 and by MailboxTransfer.Gui.Jobs.ps1).

.DESCRIPTION
    Every job leaves report.html in its job folder: what was asked, the result, the numbers, the steps and the chronology of the job (every
    step with its start, end, duration and state, the waits for the service, the sleeps of the PC, the warnings and the errors, in time order),
    one row for each folder, and what to do next. It is built from the files the job already has: events.jsonl (package\docs\MailboxTransfer-Events.md)
    and status.json. The page is one self-contained file (templates\Report.template.html): no network, no other file; every text that comes from
    a job (a folder name, a message of the service) is HTML-encoded.

        Get-MbtJobReportModel      what the report says, as data (events.jsonl is read once, a block of lines at a time; the lists are capped as they are built)
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

function Format-MbtReportQuantity {
    <# A count and its noun: "1 time", "3 times". #>
    param($Count, [string]$Singular, [string]$Plural = '')
    $n = ConvertTo-MbtReportLong $Count
    $word = if ($n -eq 1) { $Singular } elseif ($Plural) { $Plural } else { $Singular + 's' }
    return (Format-MbtReportNumber $n) + ' ' + $word
}

function Format-MbtReportProgress {
    <#
        What a step, a folder or a total has done, in words. -Finished: only what was really done, never "8.2 MiB of 7.9 MiB". Not finished: "x of y", and "x of about y"
        for -About (the sizes that Microsoft announces are only about right); more than the total says: x alone. Nothing done and no total: ''.
    #>
    param([double]$Done, [double]$Total, [switch]$Bytes, [switch]$About, [switch]$Finished)
    $say = { param([double]$n) if ($Bytes) { Format-MbtReportBytes $n } else { Format-MbtReportNumber $n } }
    if ($Finished) {
        $real = if ($Done -gt 0) { $Done } else { $Total }
        return $(if ($real -gt 0) { & $say $real } else { '' })
    }
    if ($Total -gt 0 -and $Done -le $Total) { return ((& $say $Done) + ' of ' + $(if ($About) { 'about ' } else { '' }) + (& $say $Total)) }
    if ($Done -gt 0) { return (& $say $Done) }
    return ''
}

function Get-MbtReportTaskName {
    <# The task alone, never the title: "Export", "Import (PST)", "Copy (EML files)". #>
    param([string]$Task, [string]$Mode)
    $name = ($Task -replace '-plan$', '').Trim()
    if ($name -eq 'exportpst') { $name = 'export' }
    if (-not $name) {
        $name = switch -Regex ($Mode) { '^Export' { 'export' } '^Import' { 'import' } '^Check' { 'check' } '^Repair' { 'repair' } default { '' } }
    }
    if (-not $name) { return '' }
    $text = $name.Substring(0, 1).ToUpperInvariant() + $name.Substring(1).ToLowerInvariant()
    $where = switch -Regex ($Mode) { '^(Export|Import)Pst$' { ' (PST)' } '^(Export|Import)Eml$' { ' (EML files)' } default { '' } }
    if ($name -notin 'export', 'import', 'copy') { $where = '' }
    if ($Task -match '-plan$') { $where += ' (plan check)' }
    return $text + $where
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
            return @{ Text = $text; Tone = 'ok'; Sentence = 'The job finished and nothing is left to do.'; Next = ''; NextTone = 'ok' }
        }
        1 {
            if ($reason -eq 'refused-items') {
                return @{ Text = 'Finished, items refused'; Tone = 'warn'; Sentence = 'The job finished, but some items were refused for good.'; Next = 'Nothing can be run again: trying again cannot change the answer of the service. The items that were refused are listed below and in the list of items not imported.'; NextTone = 'warn' }
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
                'done' { '' }
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

function Read-MbtJobEventChunk {
    <#
        The events of lines [Start, Start + Count) of events.jsonl, parsed in one call, in the order they were written (a block at a time, so that a file of hundreds of
        thousands of events never sits in memory as parsed objects). A line that is not JSON (the one still being written) is left out.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Lines, [Parameter(Mandatory)][int]$Start, [Parameter(Mandatory)][int]$Count)
    $parsed = $null
    try { $parsed = ('[' + [string]::Join(',', $Lines, $Start, $Count) + ']') | ConvertFrom-Json -AsHashtable -Depth 20 -NoEnumerate } catch { $parsed = $null }
    if ($null -ne $parsed) { return , $parsed }
    $each = [Collections.Generic.List[object]]::new($Count)
    for ($i = $Start; $i -lt $Start + $Count; $i++) {
        if ($Lines[$i].Trim().Length -lt 2) { continue }
        try { $each.Add(($Lines[$i] | ConvertFrom-Json -AsHashtable -Depth 20)) } catch { Write-Verbose 'An event line that is not JSON is left out.' }
    }
    return , $each
}

function Get-MbtJobEndpoint {
    <# What a job reads from and writes to, in a few words: status.json says it, else the title and the mode do. Returns Source and Destination. #>
    param([string]$Title, [string]$Mode, [string]$Store, [string]$Source, [string]$Destination)
    if (-not $Source -and -not $Destination) {
        $tail = if ($Title -match ' - (.+)$') { $Matches[1] } else { '' }
        switch ($Mode) {
            'Export' { $Source = $tail; $Destination = $Store }
            'ExportPst' { $Source = $Store; $Destination = $tail }
            'Import' { $Source = $Store; $Destination = $tail }
            'ImportPst' { $Source = 'PST files'; $Destination = $tail }
            'ExportEml' { $Destination = $tail }
            'ImportEml' { $Source = 'EML files'; $Destination = $tail }
            'CheckPst' { $Source = $tail }
            'RepairPst' { $Source = $tail }
        }
    }
    return @{ Source = $Source; Destination = $Destination }
}

function Get-MbtJobReportModel {
    <#
    .SYNOPSIS
        What the report of a job folder says, as data: header, result, counts, steps, chronology, folders, items, warnings, errors, decisions.
    .DESCRIPTION
        Reads events.jsonl and status.json of the folder. A job that has no result event (killed, power loss) is reported as interrupted.
        The events are read once, a block at a time, and the lists are capped as they are built (MaxChronology, MaxFolders, MaxItems, MaxWarnings, MaxDecisions): a job of
        50,000 folders and 200,000 events gives the same small page as a job of a hundred. Left* says how many rows were left out. When there are more folders than the cap,
        the ones that are not done come first (a folder that failed or is not finished is what a reader looks for). The numbers (counts, waits, steps) are always complete.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Folder,
        [int]$MaxChronology = 2000, [int]$MaxFolders = 2000, [int]$MaxItems = 500, [int]$MaxWarnings = 500, [int]$MaxDecisions = 200, [int]$MaxText = 2000
    )
    Set-StrictMode -Version Latest
    $status = Read-MbtReportJsonFile -Path (Join-Path $Folder 'status.json')
    if ($null -eq $status) { $status = @{} }

    $steps = [ordered]@{}
    $chrono = [Collections.Generic.List[hashtable]]::new()
    $chronoLeft = [ref]0L
    $tail = @{ Kind = ''; Reason = ''; Row = $null }
    $folders = @{}
    $folderList = [Collections.Generic.List[hashtable]]::new()
    $items = [ordered]@{}
    $warnings = [Collections.Generic.List[hashtable]]::new()
    $warningCount = 0L
    $errors = [Collections.Generic.List[hashtable]]::new()
    $decisions = [Collections.Generic.List[hashtable]]::new()
    $decisionCount = 0L
    $waitSpans = [Collections.Generic.List[datetime[]]]::new()
    $sleepSpans = [Collections.Generic.List[hashtable]]::new()
    $marks = [Collections.Generic.List[hashtable]]::new()
    $rates = [Collections.Generic.List[double]]::new()
    $phase = @{}
    $result = $null
    $first = $null; $last = $null; $lastRaw = $null
    $waitCount = 0L; $waitSeconds = 0.0
    $eventCount = 0L
    $job = [string](Get-MbtReportValue $status 'jobId' ''); $task = [string](Get-MbtReportValue $status 'task' ''); $mode = [string](Get-MbtReportValue $status 'mode' '')
    $period = [string](Get-MbtReportValue (Get-MbtReportValue $status 'period' $null) 'text' '')
    $tickWords = [regex]::new('\d[\d,\.]*')
    $oneText = { param([string]$S) if ($S.Length -gt $MaxText) { $S.Substring(0, $MaxText) + '...' } else { $S } }

    $addStep = {
        param([string]$Id, [string]$Title)
        if (-not $steps.Contains($Id)) { $steps[$Id] = [ordered]@{ Id = $Id; Title = $(if ($Title) { $Title } else { $Id }); State = 'waiting'; Started = $null; Finished = $null; Done = 0L; Total = 0L; Bytes = 0L; BytesTotal = 0L; Msg = '' } }
        elseif ($Title -and $steps[$Id].Title -eq $Id) { $steps[$Id].Title = $Title }
    }
    $titleOf = { param([string]$Id) if ($steps.Contains($Id)) { $steps[$Id].Title } else { $Id } }
    # a row of the chronology: kept while there is room, else only counted (the line "N more events" says how many)
    $addRow = {
        param([hashtable]$Row)
        $tail.Kind = $Row.Kind; $tail.Reason = $(if ($Row.ContainsKey('Reason')) { $Row.Reason } else { '' })
        if ($chrono.Count -lt $MaxChronology) { $chrono.Add($Row); $tail.Row = $Row; return $true }
        $tail.Row = $null
        $chronoLeft.Value++
        return $false
    }

    $eventsFile = Join-Path $Folder 'events.jsonl'
    $lines = [string[]]@()
    if (Test-Path -LiteralPath $eventsFile -PathType Leaf) {
        $stream = [IO.FileStream]::new($eventsFile, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete)
        try { $lines = [IO.StreamReader]::new($stream, [Text.UTF8Encoding]::new($false), $true).ReadToEnd().Split("`n") } finally { $stream.Dispose() }
    }
    $lineCount = $lines.Length
    while ($lineCount -gt 0 -and $lines[$lineCount - 1].Trim().Length -lt 2) { $lineCount-- }
    for ($start = 0; $start -lt $lineCount; $start += 2000) {
        $block = Read-MbtJobEventChunk -Lines $lines -Start $start -Count ([math]::Min(2000, $lineCount - $start))
        foreach ($e in $block) {
            if ($e -isnot [hashtable]) { continue }
            $eventCount++
            $kind = [string]$e['e']
            $at = $e['t']
            # the folders and the items are the bulk of a large job: they need no time, only the order
            if ($kind -eq 'object' -or $kind -eq 'item') {
                if (-not $first) { $first = ConvertTo-MbtReportTime $at }
                $lastRaw = $at
                $id = [string]$e['step'] + '|' + [string]$e['key']
                if ($kind -eq 'object') {
                    $o = $folders[$id]
                    if ($null -eq $o) { $o = @{ Step = [string]$e['step']; Key = [string]$e['key']; State = ''; Done = 0L; Total = 0L; Bytes = -1L; Msg = ''; Detail = ''; Count = 0L }; $folders[$id] = $o; $folderList.Add($o) }
                } else {
                    if ($items.Contains($id)) { $o = $items[$id] } else { $o = @{ Step = [string]$e['step']; Key = [string]$e['key']; State = ''; Done = 0L; Total = 0L; Bytes = -1L; Msg = ''; Detail = ''; Count = 0L }; $items[$id] = $o }
                }
                $v = $e['state']; $o.State = if ($null -ne $v) { [string]$v } else { 'running' }
                $v = $e['done']; if ($null -ne $v) { $o.Done = if ($v -is [long] -or $v -is [int]) { [long]$v } else { ConvertTo-MbtReportLong $v } }
                $v = $e['total']; if ($null -ne $v) { $o.Total = if ($v -is [long] -or $v -is [int]) { [long]$v } else { ConvertTo-MbtReportLong $v } }
                $v = $e['bytes']; if ($null -ne $v) { $o.Bytes = if ($v -is [long] -or $v -is [int]) { [long]$v } else { ConvertTo-MbtReportLong $v -1 } }
                $v = $e['msg']; if ($null -ne $v) { $o.Msg = [string]$v }
                $v = $e['detail']; if ($null -ne $v) { $o.Detail = [string]$v }
                $v = $e['count']; if ($null -ne $v) { $o.Count = ConvertTo-MbtReportLong $v }
                continue
            }
            if ($at -is [datetime]) { $at = if ($at.Kind -eq [DateTimeKind]::Unspecified) { [datetime]::SpecifyKind($at, [DateTimeKind]::Utc) } else { $at.ToUniversalTime() } }
            else { $at = ConvertTo-MbtReportTime $at }
            if ($null -eq $at) { $at = if ($last) { $last } else { [datetime]::UtcNow } }
            if (-not $first) { $first = $at }
            $last = $at
            $lastRaw = $null
            switch ($kind) {
                'step' {
                    $id = [string]$e['id']
                    & $addStep $id ''
                    $s = $steps[$id]
                    $v = $e['state']; $state = if ($null -ne $v) { [string]$v } else { 'running' }
                    $v = $e['msg']; $msg = if ($null -ne $v) { & $oneText ([string]$v) } else { '' }
                    if ($state -eq 'running' -and -not $s.Started) {
                        $s.Started = $at
                        [void](& $addRow @{ T = $at; Kind = 'step'; Tone = 'info'; Step = $s.Title; Text = "$($s.Title) started"; Until = $null; Count = 0 })
                    }
                    if ($state -in 'done', 'failed') {
                        if (-not $s.Started) { $s.Started = $at }
                        $s.Finished = $at
                        if ($phase.ContainsKey($id) -and $null -ne $phase[$id].Row -and $null -eq $phase[$id].Row.Until) { $phase[$id].Row.Until = $at }
                        [void](& $addRow @{ T = $at; Kind = 'step'; Tone = $(if ($state -eq 'done') { 'ok' } else { 'fail' }); Step = $s.Title; Text = ("$($s.Title) " + $(if ($state -eq 'done') { 'done' } else { 'failed' }) + $(if ($msg) { ": $msg" } else { '' })); Until = $null; Count = 0 })
                    }
                    $s.State = $state
                    if ($state -eq 'done') { $s.Msg = '' }
                    $v = $e['done']; if ($null -ne $v) { $s.Done = ConvertTo-MbtReportLong $v }
                    $v = $e['total']; if ($null -ne $v) { $s.Total = ConvertTo-MbtReportLong $v }
                    $v = $e['bytes']; if ($null -ne $v) { $s.Bytes = ConvertTo-MbtReportLong $v }
                    $v = $e['bytesTotal']; if ($null -ne $v) { $s.BytesTotal = ConvertTo-MbtReportLong $v }
                    if ($msg) {
                        $s.Msg = $msg
                        if ($state -eq 'running') {
                            $norm = $tickWords.Replace($msg, '#')
                            if ($phase.ContainsKey($id) -and $phase[$id].Norm -eq $norm -and ($null -ne $phase[$id].Row -or $phase[$id].Over)) { if ($phase[$id].Row) { $phase[$id].Row.Text = $msg } }
                            else {
                                if ($phase.ContainsKey($id) -and $null -ne $phase[$id].Row -and $null -eq $phase[$id].Row.Until) { $phase[$id].Row.Until = $at }
                                $row = @{ T = $at; Kind = 'phase'; Tone = 'info'; Step = $s.Title; Text = $msg; Until = $null; Count = 0 }
                                $kept = & $addRow $row
                                $phase[$id] = @{ Norm = $norm; Row = $(if ($kept) { $row } else { $null }); Over = (-not $kept) }
                            }
                        }
                    }
                }
                'wait' {
                    $v = $e['seconds']; $secs = if ($null -ne $v) { [double]$v } else { 0.0 }
                    $v = $e['reason']; $reason = if ($null -ne $v) { [string]$v } else { 'throttle' }
                    $v = $e['count']; $waitCount = [math]::Max($waitCount, $(if ($null -ne $v) { ConvertTo-MbtReportLong $v } else { $waitCount + 1 }))
                    $v = $e['waited']; $waitSeconds = [math]::Max($waitSeconds, $(if ($null -ne $v) { [double]$v } else { $waitSeconds + $secs }))
                    $waitSpans.Add([datetime[]]@($at, $at.AddSeconds($secs)))
                    if ($tail.Kind -eq 'wait' -and $tail.Reason -eq $reason) {
                        $prev = $tail.Row
                        if ($prev) { $prev.Count++; $prev.Seconds += $secs; $prev.Until = $at.AddSeconds($secs) }
                    } else { [void](& $addRow @{ T = $at; Kind = 'wait'; Tone = 'warn'; Step = ''; Reason = $reason; Text = ''; Until = $at.AddSeconds($secs); Count = 1; Seconds = $secs }) }
                }
                'rate' { $v = $e['gibPerHour']; $rates.Add($(if ($null -ne $v) { [double]$v } else { 0.0 })) }
                'warn' {
                    $warningCount++
                    $text = & $oneText ([string]$e['msg'])
                    if ($warnings.Count -lt $MaxWarnings) { $warnings.Add(@{ T = $at; Code = [string]$e['code']; Step = [string]$e['step']; Key = [string]$e['key']; Msg = $text }) }
                    if ($marks.Count -lt 300) { $marks.Add(@{ T = $at; Kind = 'warn'; Text = $text }) }
                    [void](& $addRow @{ T = $at; Kind = 'warn'; Tone = 'warn'; Step = (& $titleOf ([string]$e['step'])); Text = $text; Until = $null; Count = 0 })
                }
                'decision' {
                    $decisionCount++
                    $text = & $oneText ([string]$e['msg'])
                    if ($decisions.Count -lt $MaxDecisions) { $decisions.Add(@{ T = $at; Key = [string]$e['key']; Msg = $text; Folder = [string]$e['folder'] }) }
                    [void](& $addRow @{ T = $at; Kind = 'decision'; Tone = 'warn'; Step = ''; Text = $text; Until = $null; Count = 0 })
                }
                'slept' {
                    $from = ConvertTo-MbtReportTime $e['from']; $to = ConvertTo-MbtReportTime $e['to']
                    $v = $e['seconds']; $secs = if ($null -ne $v) { [double]$v } else { 0.0 }
                    $row = @{ T = $(if ($from) { $from } else { $at }); Kind = 'sleep'; Tone = 'info'; Step = ''; Text = ('This PC was asleep for ' + (Format-MbtReportDuration $secs) + '. The job went on when it woke up.'); Until = $(if ($to) { $to } else { $at }); Count = 1; Seconds = $secs }
                    if ($sleepSpans.Count -lt 200) { $sleepSpans.Add($row) }
                    [void](& $addRow $row)
                }
                'info' {
                    $text = & $oneText ([string]$e['msg'])
                    if ($text) { [void](& $addRow @{ T = $at; Kind = 'info'; Tone = 'info'; Step = ''; Text = $text; Until = $null; Count = 0 }) }
                }
                'error' {
                    $err = @{ T = $at; Code = [string]$e['code']; Msg = (& $oneText ([string]$e['msg'])); Tech = (& $oneText ([string]$e['tech'])) }
                    if ($errors.Count -lt 100) { $errors.Add($err) }
                    if ($marks.Count -lt 300) { $marks.Add(@{ T = $at; Kind = 'error'; Text = $err.Msg }) }
                    [void](& $addRow @{ T = $at; Kind = 'error'; Tone = 'fail'; Step = ''; Text = $err.Msg; Until = $null; Count = 0 })
                }
                'plan' {
                    if (-not $job) { $job = [string]$e['job'] }
                    if (-not $task) { $task = [string]$e['task'] }
                    if (-not $mode) { $mode = [string]$e['mode'] }
                    $planPeriod = [string](Get-MbtReportValue (Get-MbtReportValue $e 'period' $null) 'text' '')
                    if ($planPeriod) { $period = $planPeriod }
                    foreach ($s in @(Get-MbtReportValue $e 'steps' @())) { & $addStep ([string](Get-MbtReportValue $s 'id' '')) ([string](Get-MbtReportValue $s 'title' '')) }
                }
                'result' {
                    $result = @{ Exit = [int](ConvertTo-MbtReportLong (Get-MbtReportValue $e 'exit' 1) 1); Complete = [bool](Get-MbtReportValue $e 'complete' $false); Counts = (Get-MbtReportValue $e 'counts' @{}); Next = [string](Get-MbtReportValue $e 'next' ''); Reason = [string](Get-MbtReportValue $e 'reason' ''); At = $at }
                }
            }
        }
    }
    # the time of the last event, when the last events were folders (their time was not read)
    if ($null -ne $lastRaw) { $lastFolder = ConvertTo-MbtReportTime $lastRaw; if ($lastFolder -and (-not $last -or $lastFolder -gt $last)) { $last = $lastFolder } }
    foreach ($id in @($phase.Keys)) { if ($phase[$id].Row -and $null -eq $phase[$id].Row.Until) { $phase[$id].Row.Until = $last } }
    foreach ($row in $chrono) { if ($row.Kind -eq 'wait') { $row.Text = if ($row.Reason -eq 'network') { "Microsoft could not be reached; the job tried again ($(Format-MbtReportQuantity $row.Count 'time'), $(Format-MbtReportDuration $row.Seconds) in all)." } elseif ($row.Count -gt 1) { "Microsoft asked to wait because of too many requests: $($row.Count) waits, $(Format-MbtReportDuration $row.Seconds) in all." } else { "Microsoft asked to wait $(Format-MbtReportDuration $row.Seconds) because of too many requests." } } }

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
    $worked = [Collections.Generic.List[hashtable]]::new()
    foreach ($f in $folderList) { if ($f.Step -ne 'plan') { $worked.Add($f) } }
    $folderRows = $folderList
    if ($worked.Count) { $folderRows = $worked }
    $stepList = @($steps.Values)
    # a job without counts: the folders say how far it got
    if (-not $counts.Contains('done')) {
        $sum = 0L; $nWork = 0
        foreach ($w in $folderRows) { if ($w.Step -notin 'plan', '') { $sum += [long]$w.Done; $nWork++ } }
        $most = 0L; foreach ($s in $stepList) { $most = [math]::Max($most, [long]$s.Done) }
        if ($nWork) { $counts['done'] = $sum } elseif ($stepList.Count) { $counts['done'] = $most }
    }
    # more folders than the page holds: the ones that are not done first (in the order of the events), then the others
    $folderShown = $folderRows
    if ($folderRows.Count -gt $MaxFolders) {
        $take = [bool[]]::new($folderRows.Count); $n = 0
        for ($i = 0; $i -lt $folderRows.Count -and $n -lt $MaxFolders; $i++) { if ($folderRows[$i].State -notin 'done', 'ok', 'complete') { $take[$i] = $true; $n++ } }
        for ($i = 0; $i -lt $folderRows.Count -and $n -lt $MaxFolders; $i++) { if (-not $take[$i]) { $take[$i] = $true; $n++ } }
        $folderShown = [Collections.Generic.List[hashtable]]::new($MaxFolders)
        for ($i = 0; $i -lt $folderRows.Count; $i++) { if ($take[$i]) { $folderShown.Add($folderRows[$i]) } }
    }
    $itemShown = @($items.Values | Select-Object -First $MaxItems)

    $title = [string](Get-MbtReportValue $status 'title' '')
    if (-not $title) { $title = 'Mailbox Transfer job' }
    $store = [string](Get-MbtReportValue $status 'store' '')
    $ends = Get-MbtJobEndpoint -Title $title -Mode $mode -Store $store -Source ([string](Get-MbtReportValue $status 'source' '')) -Destination ([string](Get-MbtReportValue $status 'destination' ''))
    $source = $ends.Source; $destination = $ends.Destination

    $model = [ordered]@{
        Title = $title; Task = $task; TaskName = (Get-MbtReportTaskName -Task $task -Mode $mode); Mode = $mode; JobId = $job; Source = $source; Destination = $destination; Store = $store; Period = $period
        Started = $first; Finished = $finished; Seconds = $seconds; Interrupted = $interrupted
        Result = $(if ($result) { @{ Exit = $result.Exit; Complete = $result.Complete; Next = $result.Next; Reason = $result.Reason } } else { $null })
        Verdict = $verdict
        Counts = $counts
        Waits = @{ Count = $waitCount; Seconds = $waitSeconds }
        Rates = @{ Samples = $rates.Count; MaxGiBPerHour = $(if ($rates.Count) { [math]::Round(($rates | Measure-Object -Maximum).Maximum, 2) } else { 0 }) }
        Steps = $stepList
        Chronology = $chrono.ToArray()
        LeftChronology = $chronoLeft.Value
        WaitSpans = $waitSpans
        SleepSpans = $sleepSpans.ToArray()
        Marks = $marks.ToArray()
        Folders = $folderShown.ToArray()
        LeftFolders = [math]::Max(0, $folderRows.Count - $folderShown.Count)
        Items = $itemShown
        LeftItems = [math]::Max(0, $items.Count - $itemShown.Count)
        Warnings = $warnings.ToArray()
        LeftWarnings = [math]::Max(0, $warningCount - $warnings.Count)
        Errors = $errors.ToArray()
        Decisions = $decisions.ToArray()
        LeftDecisions = [math]::Max(0, $decisionCount - $decisions.Count)
        EventCount = $eventCount
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
    $sleeps = @($Model.SleepSpans)
    $marks = @($Model.Marks)
    # the waits of a long job are many: the ones that fall on the same pixel are drawn as one bar
    $waits = [Collections.Generic.List[object]]::new()
    $pixel = $span / ($x1 - $x0)
    foreach ($w in $Model.WaitSpans) {
        $cur = if ($waits.Count) { $waits[$waits.Count - 1] } else { $null }
        if ($cur -and ($w[0] - $cur.To).TotalSeconds -le $pixel) { if ($w[1] -gt $cur.To) { $cur.To = $w[1] }; $cur.Count++; $cur.Seconds += ($w[1] - $w[0]).TotalSeconds }
        else { $waits.Add(@{ From = $w[0]; To = $w[1]; Count = 1; Seconds = ($w[1] - $w[0]).TotalSeconds }) }
    }
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
                    $a = [datetime]$w.From; $b = [datetime]$w.To
                    $bx = & $xOf $a
                    $bw = [math]::Max(2.0, (& $xOf $b) - $bx)
                    $tip = (Format-MbtReportTime $a -Short) + ': waited ' + (Format-MbtReportDuration $w.Seconds) + $(if ($w.Count -gt 1) { ' in ' + (Format-MbtReportQuantity $w.Count 'wait') } else { '' })
                    [void]$sb.Append('<rect class="bar wait" x="' + (& $num $bx) + '" y="' + $y + '" width="' + (& $num $bw) + '" height="' + $laneHeight + '" rx="2"><title>' + (ConvertTo-MbtHtmlText $tip) + '</title></rect>')
                }
            }
            'marks' {
                foreach ($m in $marks) {
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
        'Task' = $Model.TaskName; 'Source' = $Model.Source; 'Destination' = $Model.Destination
        'Store' = $(if ($Model.Store -and $Model.Store -ne $Model.Source -and $Model.Store -ne $Model.Destination) { $Model.Store } else { '' })
        'Period' = $(if ($Model.Contains('Period') -and $Model.Period) { 'Items that ' + $Model.Period + ' in the mailbox' } else { '' })
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
        @('skipped', 'Skipped', '', $false), @('resumed', 'Resumed', '', $false), @('kept', 'Items kept', '', $false), @('outsidePeriod', 'Outside the period', '', $false), @('withoutDate', 'Kept without a date', '', $false), @('bytes', 'Data', 'bytes', $false), @('fileBytes', 'File size', 'bytes', $false), @('parts', 'PST files', '', $false)
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
    $folderOfReport = if ($Model.Contains('Folder') -and $Model.Folder) { [string]$Model.Folder } else { $null }
    $moreIn = if ($folderOfReport -and (Test-Path -LiteralPath (Join-Path $folderOfReport 'report.csv') -PathType Leaf)) { 'report.csv and events.jsonl' } else { 'events.jsonl' }
    $nextBody = '<p>' + (& $enc $verdict.Sentence) + '</p>'
    if ($verdict.Next -and $verdict.Next -ne $verdict.Sentence) { $nextBody += '<p><strong>' + (& $enc $verdict.Next) + '</strong></p>' }
    $next = '<section class="card next ' + $verdict.NextTone + '"><h2>What to do next</h2>' + $nextBody + '</section>'
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
        [void]$sb.Append('<section class="card attention"><h2>Items waiting for a decision<span class="count">' + (Format-MbtReportNumber ($Model.Decisions.Count + $Model.LeftDecisions)) + '</span></h2><ul>')
        foreach ($d in $Model.Decisions) { [void]$sb.Append('<li>' + (& $enc $d.Msg) + $(if ($d.Folder) { ' <small>' + (& $enc $d.Folder) + '</small>' } else { '' }) + '</li>') }
        [void]$sb.Append('</ul>')
        if ($Model.LeftDecisions -gt 0) { [void]$sb.Append('<p class="caption">' + $(if ($Model.LeftDecisions -eq 1) { '1 more item is' } else { (Format-MbtReportNumber $Model.LeftDecisions) + ' more items are' }) + ' in events.jsonl.</p>') }
        [void]$sb.Append('</section>')
        $decisionHtml = $sb.ToString()
    }

    # steps
    $sb = [Text.StringBuilder]::new()
    if ($Model.Steps.Count) {
        [void]$sb.Append('<div class="scroll"><table><thead><tr><th>Step</th><th>State</th><th>Started</th><th>Finished</th><th class="num">Duration</th><th class="num">Items</th><th class="num">Data</th><th>Message</th></tr></thead><tbody>')
        foreach ($s in $Model.Steps) {
            $dur = if ($s.Started -and $s.Finished) { Format-MbtReportDuration (([datetime]$s.Finished) - ([datetime]$s.Started)).TotalSeconds } elseif ($s.Started) { 'not finished' } else { '' }
            $finishedStep = $s.State -eq 'done'
            $itemsText = Format-MbtReportProgress $s.Done $s.Total -Finished:$finishedStep
            $dataText = Format-MbtReportProgress $s.Bytes $s.BytesTotal -Bytes -About -Finished:$finishedStep
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
        if ($Model.LeftChronology -gt 0) { [void]$sb.Append('<p class="caption">' + $(if ($Model.LeftChronology -eq 1) { '1 more event is' } else { (Format-MbtReportNumber $Model.LeftChronology) + ' more events are' }) + ' in events.jsonl. This table shows the first ' + (Format-MbtReportNumber $Model.Chronology.Count) + '.</p>') }
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
            $itemsText = Format-MbtReportProgress $f.Done $f.Total -Finished:($f.State -in 'done', 'ok', 'complete')
            if (-not $itemsText) { $itemsText = '0' }
            [void]$sb.Append('<tr><td>' + (& $enc $f.Key) + '</td><td>' + (& $enc $stepText) + '</td><td>' + (Get-MbtReportBadge $f.State) + '</td><td class="num">' + (& $enc $itemsText) + '</td><td class="num">' + (& $enc $(if ($f.Bytes -ge 0) { Format-MbtReportBytes $f.Bytes } else { '' })) + '</td><td>' + (& $enc $f.Msg) + '</td></tr>')
        }
        [void]$sb.Append('</tbody></table></div>')
        if ($Model.LeftFolders -gt 0) { [void]$sb.Append('<p class="caption">' + $(if ($Model.LeftFolders -eq 1) { '1 more folder is' } else { (Format-MbtReportNumber $Model.LeftFolders) + ' more folders are' }) + ' in ' + $moreIn + '. The folders that are not done come first.</p>') }
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
        if ($Model.LeftItems -gt 0) { [void]$sb.Append('<p class="caption">' + $(if ($Model.LeftItems -eq 1) { '1 more is' } else { (Format-MbtReportNumber $Model.LeftItems) + ' more are' }) + ' in ' + $moreIn + '.</p>') }
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
        if ($Model.LeftWarnings -gt 0) { [void]$sb.Append('<p class="caption">' + $(if ($Model.LeftWarnings -eq 1) { '1 more warning is' } else { (Format-MbtReportNumber $Model.LeftWarnings) + ' more warnings are' }) + ' in events.jsonl.</p>') }
        [void]$sb.Append('</section>')
        $warningsHtml = $sb.ToString()
    }

    # files
    $sb = [Text.StringBuilder]::new()
    [void]$sb.Append('<ul class="files">')
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
        period = $(if ($Model.Contains('Period') -and $Model.Period) { [string]$Model.Period } else { $null })
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
    $subtitle = 'Started ' + (Format-MbtReportTime $started) + ' ' + [char]0x00B7 + ' ' + (Format-MbtReportDuration $Model.Seconds)
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

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBl1DBLHrWE8qSM
# hiMfMZiZh9VUXERKMi2e8QwvSqr4CKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCLmA0opaDz
# 9bVUnbFoNojbZ9KprhcPqL5QMKCqm88KnjANBgkqhkiG9w0BAQEFAASCAgCnwcEt
# nngHPCvGQIsDHOTs4aZlo8h63jrPoWh0udMHmCLjqjAAz6qVx6/cWHGH1/BVbrZb
# rIM3Rkqo4MFDAf5S3aiKsfNleg4JWtJ/ftvJGH0bX4wqHuWX7pn6YnTiHhzWoi8R
# knNKJ3mdHj4kQRkf1NvUsE0tMpGJtU2QJbnRFcWumXA6NIijrDZwrKCNuX4JQn5X
# o90q+fg5p6UsGh/F32ZXcxAOqaA9QGOVfBZF53rriCn9bKTBQhK9DdhZp6T6ErBc
# uMl17/IQTqst/Wr/faVJajk4GGqLAcb7Fw+OnL3pk5qSZV7BoO/5e4qHJkgpry2Q
# OWHK6p0cPlB/nNG4tFNh1PND1gD0l4Qq/ruPsO7NFVCqunb8YqjyZcuX/Cysiqp9
# qMy/3c9ovOZcuVytxqFrAXVMVvCibtsAthE5S+xXVyDaSZko8LZsFIz4MZRd/xml
# QzJQ/cQBM4Bhx7xJ4VP6svODNFuNmEwIzAcnGmAQKK6m2ysXO6FLraaKw1DxCfXH
# +om1UiR7Pei8apoK1LnXBPNmzMxhHEdVyjZhiUieMrucVLvhUeeTaZV1w8NNbE5f
# N93jUp80SMIoWOzVHQoeQwpKI1mC7uX+JwyC+qF59C6qsHqyk0O30EVYZ5+ctvWT
# Rp/V1DUt3Y3DsaTet7zXX+sUrWzhncP9P7qWW6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIEkkfs9n2h9/3LhnhH+pNk+g3KsQWOImZU6uqms8yBmBAggkG/QRFTxr
# 9RgPMjAyNjExMTUxNDM0MzVaMAMCAQECECGkRS62TCCQ4VcpqWmG4o+gghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTE0MzQzNVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIKUm
# Tplp3cpHNt6XORfwkhxLLzr5pyzpAbIwTiqKPp/mMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiBwzIzzazdGJYRxOaZhoytjHM3spmSt8xGnF4wiAddQewIhAImX
# wBcLFa/dsx//YsY7nFqmnxkd/yVL9RMBDiE7z0iw
# SIG # End signature block
