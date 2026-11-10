#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - the model of a run (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    Pure logic, no WPF: the events of a job (package\docs\MailboxTransfer-Events.md, JSON lines) are applied to a run model; the Run and Done
    pages draw that model. The console never parses text lines of the engine. Unknown event kinds and unknown fields are ignored.

    Also here: the honest ETA (a range from the sustained rate, nothing during the first minutes and after a sleep) and the ranges of the Plan
    page (constants in $script:Throughput, Gui.Core.ps1).

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

function Get-MbtGuiProp {
    <# A property of a parsed event, or $Default (strict mode safe). #>
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] } else { return $Default } }
    $p = $Object.PSObject.Properties[$Name]
    if ($p -and $null -ne $p.Value) { return $p.Value }
    return $Default
}

function New-MbtGuiRun {
    param([string]$Task = 'export', [string]$Title = '', [string]$Unit = 'items')
    [ordered]@{
        Task = $Task; Title = $Title; Unit = $Unit; State = 'running'; Planned = $false
        Steps = [Collections.Generic.List[hashtable]]::new()
        Objects = [Collections.Generic.List[hashtable]]::new()
        Log = [Collections.Generic.List[hashtable]]::new()
        Items = [Collections.Generic.List[hashtable]]::new()
        Decisions = [Collections.Generic.List[hashtable]]::new()
        Warnings = [Collections.Generic.List[hashtable]]::new()
        Waits = @{ Count = 0; Seconds = 0.0; Active = $null }
        Rates = [Collections.Generic.List[hashtable]]::new()
        Samples = [Collections.Generic.List[hashtable]]::new()
        Started = $null; LastEvent = $null; EpochStart = $null
        Done = 0L; Total = 0L; Bytes = 0L; BytesTotal = 0L
        Sleep = $null; Error = $null; Result = $null; StopRequested = $false; Dirty = $true; EventCount = 0
    }
}

function Add-MbtGuiLog {
    param($Run, [string]$Kind, [string]$Text, $At)
    $Run.Log.Insert(0, @{ T = $At; Kind = $Kind; Text = $Text })
    while ($Run.Log.Count -gt 300) { $Run.Log.RemoveAt($Run.Log.Count - 1) }
}

function ConvertTo-MbtGuiTime {
    param($Value)
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    if ($Value -is [string] -and $Value) { try { return [datetimeoffset]::Parse($Value, [cultureinfo]::InvariantCulture).UtcDateTime } catch { return $null } }
    return $null
}

function Find-MbtGuiStep {
    param($Run, [string]$Id)
    foreach ($s in $Run.Steps) { if ($s.Id -eq $Id) { return $s } }
    return $null
}

function Update-MbtGuiRunTotals {
    <# The overall bar follows the step that works on items (bytes when known). #>
    param($Run)
    $active = $null
    foreach ($s in $Run.Steps) { if ($s.BytesTotal -gt 0 -and ($s.State -eq 'running' -or -not $active)) { $active = $s } }
    if (-not $active) { foreach ($s in $Run.Steps) { if ($s.State -eq 'running' -and $s.Total -gt 0 -and (-not $active -or $s.Total -gt $active.Total)) { $active = $s } } }
    if (-not $active) { foreach ($s in $Run.Steps) { if ($s.Total -gt 0 -and (-not $active -or $s.Total -ge $active.Total)) { $active = $s } } }
    if ($active) {
        $Run.Done = $active.Done; $Run.Total = $active.Total
        if ($active.BytesTotal -gt 0) { $Run.Bytes = $active.Bytes; $Run.BytesTotal = $active.BytesTotal }
    }
}

function Add-MbtGuiEvent {
    <# Applies one parsed event to the run. Returns $true when something visible changed. #>
    param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)]$Event)
    $kind = [string](Get-MbtGuiProp $Event 'e' '')
    $at = ConvertTo-MbtGuiTime (Get-MbtGuiProp $Event 't')
    if (-not $at) { $at = [datetime]::UtcNow }
    if (-not $Run.Started) { $Run.Started = $at; $Run.EpochStart = $at }
    $gap = if ($Run.LastEvent) { ($at - $Run.LastEvent).TotalSeconds } else { 0 }
    $Run.LastEvent = $at
    $Run.EventCount++
    $Run.Dirty = $true
    switch ($kind) {
        'plan' {
            $Run.Planned = $true
            $Run.Steps.Clear()
            foreach ($s in @(Get-MbtGuiProp $Event 'steps' @())) {
                $Run.Steps.Add(@{ Id = [string](Get-MbtGuiProp $s 'id'); Title = [string](Get-MbtGuiProp $s 'title'); State = 'waiting'; Done = 0L; Total = 0L; Bytes = 0L; BytesTotal = 0L; Msg = ''; Started = $null; Finished = $null })
            }
        }
        'step' {
            $id = [string](Get-MbtGuiProp $Event 'id')
            $s = Find-MbtGuiStep $Run $id
            if (-not $s) { $s = @{ Id = $id; Title = $id; State = 'waiting'; Done = 0L; Total = 0L; Bytes = 0L; BytesTotal = 0L; Msg = ''; Started = $null; Finished = $null }; $Run.Steps.Add($s) }
            $state = [string](Get-MbtGuiProp $Event 'state' 'running')
            if ($state -eq 'running' -and -not $s.Started) { $s.Started = $at }
            if ($state -in 'done', 'failed') { $s.Finished = $at; if (-not $s.Started) { $s.Started = $at } }
            $s.State = $state
            $s.Done = [long](Get-MbtGuiProp $Event 'done' $s.Done)
            $s.Total = [long](Get-MbtGuiProp $Event 'total' $s.Total)
            $s.Bytes = [long](Get-MbtGuiProp $Event 'bytes' $s.Bytes)
            $s.BytesTotal = [long](Get-MbtGuiProp $Event 'bytesTotal' $s.BytesTotal)
            $s.Msg = [string](Get-MbtGuiProp $Event 'msg' $s.Msg)
            if ($state -eq 'done') { Add-MbtGuiLog $Run 'ok' "$($s.Title) done." $at }
            elseif ($state -eq 'failed') { Add-MbtGuiLog $Run 'bad' "$($s.Title) failed. $($s.Msg)" $at }
            elseif ($state -eq 'running' -and $s.Done -eq 0 -and $Run.Log.Count -lt 400) { Add-MbtGuiLog $Run 'run' "$($s.Title) started." $at }
            if ($s.Bytes -gt 0) {
                $Run.Samples.Add(@{ T = $at; Bytes = $s.Bytes })
                while ($Run.Samples.Count -gt 400) { $Run.Samples.RemoveAt(0) }
            }
            if ($Run.Waits.Active) { $Run.Waits.Active = $null }
            Update-MbtGuiRunTotals $Run
        }
        { $_ -in 'object', 'item' } {
            $key = [string](Get-MbtGuiProp $Event 'key')
            $stepId = [string](Get-MbtGuiProp $Event 'step')
            $state = [string](Get-MbtGuiProp $Event 'state' 'running')
            $tone = switch ($state) { { $_ -in 'done', 'ok' } { 'ok' } { $_ -in 'running', 'run', 'sending' } { 'run' } { $_ -in 'check', 'checking', 'warn', 'warning', 'partial', 'notready', 'not-ready', 'refused', 'decision' } { 'warn' } { $_ -in 'failed', 'bad', 'error', 'lost' } { 'bad' } default { 'wait' } }
            $list = $Run.Objects; if ($kind -eq 'item') { $list = $Run.Items }
            $o = $null
            foreach ($x in $list) { if ($x.Key -eq $key -and $x.Step -eq $stepId) { $o = $x } }
            if (-not $o) { $o = @{ Key = $key; Step = $stepId; State = $state; Tone = $tone; Done = 0L; Total = 0L; Bytes = -1L; Msg = ''; Detail = '' }; $list.Add($o) }
            $o.State = $state; $o.Tone = $tone
            $o.Done = [long](Get-MbtGuiProp $Event 'done' $o.Done); $o.Total = [long](Get-MbtGuiProp $Event 'total' $o.Total)
            $o.Bytes = [long](Get-MbtGuiProp $Event 'bytes' $o.Bytes)
            $o.Msg = [string](Get-MbtGuiProp $Event 'msg' $o.Msg); $o.Detail = [string](Get-MbtGuiProp $Event 'detail' $o.Detail)
            if ($kind -eq 'object' -and $state -in 'done', 'ok' -and $o.Total -gt 0) { Add-MbtGuiLog $Run 'ok' "Folder $key done: $(Format-MbtGuiCount $o.Total) items." $at }
            if ($kind -eq 'object' -and $tone -eq 'warn' -and $state -in 'notready', 'not-ready') { Add-MbtGuiLog $Run 'warn' "$key is not ready at Microsoft. Will try again later." $at }
        }
        'wait' {
            $secs = [double](Get-MbtGuiProp $Event 'seconds' 0)
            $Run.Waits.Count = [int](Get-MbtGuiProp $Event 'count' ($Run.Waits.Count + 1))
            $Run.Waits.Seconds = [double](Get-MbtGuiProp $Event 'waited' $Run.Waits.Seconds)
            $reason = [string](Get-MbtGuiProp $Event 'reason' 'throttle')
            $Run.Waits.Active = @{ Reason = $reason; Seconds = $secs; At = $at }
            $text = switch ($reason) {
                'throttle' { "Microsoft asked to wait $([math]::Round($secs)) s (too many requests). Waiting." }
                'network' { 'We cannot reach Microsoft. Trying again every 30 s.' }
                default { "Waiting $([math]::Round($secs)) s." }
            }
            Add-MbtGuiLog $Run 'warn' $text $at
        }
        'rate' {
            $g = [double](Get-MbtGuiProp $Event 'gibPerHour' 0)
            $Run.Rates.Add(@{ T = $at; GiBPerHour = $g; Window = [string](Get-MbtGuiProp $Event 'window' '5m'); Gap = ($gap -gt $script:Throughput.SleepGapSeconds) })
            while ($Run.Rates.Count -gt 24) { $Run.Rates.RemoveAt(0) }
        }
        'warn' {
            $msg = [string](Get-MbtGuiProp $Event 'msg' '')
            $Run.Warnings.Add(@{ Code = [string](Get-MbtGuiProp $Event 'code'); Step = [string](Get-MbtGuiProp $Event 'step'); Key = [string](Get-MbtGuiProp $Event 'key'); Msg = $msg })
            Add-MbtGuiLog $Run 'warn' $msg $at
        }
        'decision' {
            $key = [string](Get-MbtGuiProp $Event 'key')
            $Run.Decisions.Add(@{ Key = $key; Msg = [string](Get-MbtGuiProp $Event 'msg'); Candidates = @(Get-MbtGuiProp $Event 'candidates' @()); Title = [string](Get-MbtGuiProp $Event 'title' $key) })
            Add-MbtGuiLog $Run 'warn' "$key needs your decision." $at
            if ($Run.State -eq 'running') { $Run.State = 'decision' }
        }
        'slept' {
            $from = ConvertTo-MbtGuiTime (Get-MbtGuiProp $Event 'from'); $to = ConvertTo-MbtGuiTime (Get-MbtGuiProp $Event 'to')
            $Run.Sleep = @{ From = $from; To = $to; Seconds = [double](Get-MbtGuiProp $Event 'seconds' 0) }
            $Run.EpochStart = $at
            $Run.Rates.Clear()
            if ($Run.State -in 'running', 'decision') { $Run.State = 'sleep' }
            $fromText = if ($from) { $from.ToLocalTime().ToString('HH:mm') } else { '?' }; $toText = if ($to) { $to.ToLocalTime().ToString('HH:mm') } else { '?' }
            Add-MbtGuiLog $Run 'run' "Resumed after sleep. This PC slept from $fromText to $toText. Looking at the items that were in flight." $at
        }
        'info' { Add-MbtGuiLog $Run 'run' ([string](Get-MbtGuiProp $Event 'msg' '')) $at }
        'error' {
            $Run.Error = @{ Code = [string](Get-MbtGuiProp $Event 'code'); Msg = [string](Get-MbtGuiProp $Event 'msg'); Tech = [string](Get-MbtGuiProp $Event 'tech') }
            Add-MbtGuiLog $Run 'bad' $Run.Error.Msg $at
        }
        'result' {
            $exit = [int](Get-MbtGuiProp $Event 'exit' 1)
            $Run.Result = @{ Exit = $exit; Complete = [bool](Get-MbtGuiProp $Event 'complete' ($exit -eq 0)); Counts = Get-MbtGuiProp $Event 'counts' $null; Next = [string](Get-MbtGuiProp $Event 'next'); Reason = [string](Get-MbtGuiProp $Event 'reason') }
            $Run.State = switch ($exit) {
                0 { 'complete' }
                2 { if ($Run.StopRequested) { 'stopped' } else { 'partial' } }
                3 { 'decision-final' }
                # finished, but some items were refused for good (nothing left to run again): not an error
                1 { if ($Run.Result.Reason -eq 'refused-items') { 'refused' } else { 'error' } }
                default { 'error' }
            }
            if ($Run.Result.Reason -eq 'signin-required') { $Run.State = 'partial' }
            $Run.Waits.Active = $null
        }
        default { $Run.Dirty = $false }
    }
    return $true
}

function ConvertTo-MbtGuiRefusalReason {
    <# The reason the engine gives for a refused item, in plain words (the technical text stays in the CSV). #>
    param([string]$Reason)
    switch -Regex ($Reason) {
        '^cannot be read from the PST' { return 'This message cannot be read from the PST file (it is damaged at this message). It is the only one lost.' }
        '^cannot be expressed|^not sent, the stream built' { return 'Its content cannot be turned into what Microsoft accepts. It was not sent.' }
        '^refused by the service' { return 'Microsoft refused the content of this message (for instance it is too large, or the mailbox is full).' }
        '^gave up after' { return 'Every try failed and nothing was created in the mailbox. Microsoft''s last answer is in the list (CSV).' }
        '^no matching item' { return 'It was not found in the folder after the last try, and the tries allowed are used up.' }
        default { return $(if ($Reason) { $Reason } else { 'Refused.' }) }
    }
}

function Get-MbtGuiPlanTotals {
    <#
        What the engine says the plan holds, from the "plan" objects of a -WhatIf run (one per folder): items, size, folders. A folder whose size is not given (bytes -1: an auxiliary
        archive, a folder of an unknown size) is counted apart, never summed. Returns $null when the run has no plan object.
    #>
    param([Parameter(Mandatory)]$Run)
    $objs = @($Run.Objects | Where-Object { $_.Step -eq 'plan' -and $_.State -ne 'refused' -and $_.Key -ne '(folders)' })
    if (-not $objs.Count) { return $null }
    $items = 0L; $bytes = 0L; $unknown = 0; $toCreate = 0L
    foreach ($o in $objs) {
        if ($o.Total -gt 0) { $items += $o.Total }
        if ($o.Bytes -ge 0) { $bytes += $o.Bytes } elseif ($o.Total -ne 0) { $unknown++ }
    }
    $create = @($Run.Objects | Where-Object { $_.Step -eq 'plan' -and $_.Key -eq '(folders)' })
    if ($create.Count) { $toCreate = $create[0].Total }
    [pscustomobject]@{ Items = $items; Bytes = $bytes; Folders = $objs.Count; UnknownSize = $unknown; FoldersToCreate = $toCreate }
}

function Format-MbtGuiPlanTotals {
    <# "The service counted 1,208 items in 12 folders, 6.2 GiB." with the honest words when a size is not given. #>
    param([Parameter(Mandatory)]$Totals)
    $size = if ($Totals.UnknownSize -le 0) { Format-MbtGuiBytes $Totals.Bytes } elseif ($Totals.Bytes -le 0) { 'size not known' } else { '{0} + {1} folder{2} of unknown size' -f (Format-MbtGuiBytes $Totals.Bytes), $Totals.UnknownSize, $(if ($Totals.UnknownSize -gt 1) { 's' } else { '' }) }
    return ('{0} item{1} in {2} folder{3}, {4}' -f (Format-MbtGuiCount $Totals.Items), $(if ($Totals.Items -ne 1) { 's' } else { '' }), $Totals.Folders, $(if ($Totals.Folders -ne 1) { 's' } else { '' }), $size)
}

function Read-MbtGuiEventTail {
    <# Reads the complete new lines of the events file from an offset, with a second read-only stream (the engine keeps writing). Returns Lines and Offset. #>
    param([Parameter(Mandatory)][string]$Path, [long]$Offset = 0, [int]$MaxLines = 400)
    $lines = [Collections.Generic.List[string]]::new()
    if (-not (Test-Path -LiteralPath $Path)) { return [pscustomobject]@{ Lines = $lines; Offset = $Offset } }
    $fs = [IO.FileStream]::new($Path, 'Open', 'Read', 'ReadWrite,Delete')
    try {
        if ($fs.Length -lt $Offset) { $Offset = 0 }
        $fs.Seek($Offset, 'Begin') | Out-Null
        $buf = [byte[]]::new([math]::Min($fs.Length - $Offset, 4MB))
        $n = $fs.Read($buf, 0, $buf.Length)
        $last = [Array]::LastIndexOf($buf, [byte]10, [math]::Max(0, $n - 1))
        if ($n -gt 0 -and $last -ge 0) {
            $text = [Text.Encoding]::UTF8.GetString($buf, 0, $last + 1)
            foreach ($l in ($text -split "`n")) { $t = $l.Trim(); if ($t) { $lines.Add($t) } }
            $Offset += $last + 1
        }
    } finally { $fs.Dispose() }
    return [pscustomobject]@{ Lines = $lines; Offset = $Offset }
}

function ConvertFrom-MbtGuiEventLine {
    param([string]$Line)
    try { return ($Line | ConvertFrom-Json -ErrorAction Stop) } catch { return $null }
}

#region ETA -----------------------------------------------------------------------------------

function Format-MbtGuiEtaMinutes {
    <# Rounds as the spec says: 5 min under 2 h, 10 min under 12 h, 30 min above. #>
    param([double]$Minutes)
    $step = if ($Minutes -lt 120) { 5 } elseif ($Minutes -lt 720) { 10 } else { 30 }
    $m = [math]::Max($step, [math]::Round($Minutes / $step) * $step)
    return $m
}

function Format-MbtGuiSpan {
    <# 12 min, 1.5 h, 9 hours: minutes under 90, tenths of an hour under 3 hours, whole hours above. #>
    param([double]$Minutes, [switch]$Short)
    if ($Minutes -lt 90) { return ('{0} min' -f [math]::Max(1, [math]::Round($Minutes))) }
    if ($Minutes -lt 180) { return ('{0} h' -f ([math]::Round($Minutes / 60, 1)).ToString([cultureinfo]::InvariantCulture)) }
    $h = [math]::Round($Minutes / 60)
    if ($Short) { return "$h" }
    return "$h hours"
}

function Get-MbtGuiPlanEta {
    <# The range shown before the run: from ceilings, never one number (constants in $script:Throughput). #>
    param([ValidateSet('import', 'export')][string]$Task, [double]$Bytes)
    $gib = $Bytes / 1GB
    $range = if ($Task -eq 'import') { $script:Throughput.ImportGiBPerHour } else { $script:Throughput.ExportGiBPerHour }
    if ($gib -lt 0.5) { return [pscustomobject]@{ Text = 'a few minutes'; Low = 2; High = 6 } }
    $lo = $gib / $range[1] * 60; $hi = $gib / $range[0] * 60
    if ($Task -eq 'export') { $lo = [math]::Max($lo, $script:Throughput.ExportMinimumMinutes[0]); $hi = [math]::Max($hi, $script:Throughput.ExportMinimumMinutes[1]) }
    $mid = [math]::Sqrt($lo * $hi)
    $text = if ($hi -ge 180) { 'about {0} (between {1} and {2})' -f (Format-MbtGuiSpan $mid), (Format-MbtGuiSpan $lo -Short), (Format-MbtGuiSpan $hi) }
    elseif ($Task -eq 'import') { 'about {0} (between {1} and {2})' -f (Format-MbtGuiSpan $mid), (Format-MbtGuiSpan $lo), (Format-MbtGuiSpan $hi) }
    else { 'between {0} and {1}' -f (Format-MbtGuiSpan $lo), (Format-MbtGuiSpan $hi) }
    [pscustomobject]@{ Text = $text; Low = $lo; High = $hi }
}
function Get-MbtGuiEta {
    <#
        The time left during a run. Nothing for the first minutes (and after a sleep): 'Measuring...'. Then a range from the sustained rate of the last
        windows (P80 to P20), rounded. Windows with a pause (gap over 2 minutes) are dropped. Returns State (measuring, ready, none), Text, Basis.
    #>
    param([Parameter(Mandatory)]$Run, [datetime]$Now = [datetime]::UtcNow, [string]$Task = 'import')
    $remaining = $Run.BytesTotal - $Run.Bytes
    $epoch = if ($Run.EpochStart) { $Run.EpochStart } else { $Run.Started }
    $age = if ($epoch) { ($Now - $epoch).TotalMinutes } else { 0 }
    if ($Run.BytesTotal -le 0 -or $remaining -le 0) { return [pscustomobject]@{ State = 'none'; Text = ''; Basis = '' } }
    $good = @($Run.Rates | Where-Object { -not $_.Gap -and $_.GiBPerHour -gt 0 } | Select-Object -Last $script:Throughput.EtaWindows)
    if ($age -lt $script:Throughput.EtaWarmupMinutes -or $good.Count -eq 0) {
        return [pscustomobject]@{ State = 'measuring'; Text = 'Measuring...'; Basis = 'The time left shows after about 10 minutes.' }
    }
    $sorted = @($good | ForEach-Object { $_.GiBPerHour } | Sort-Object)
    $p20 = $sorted[[math]::Min($sorted.Count - 1, [int][math]::Floor(0.2 * ($sorted.Count - 1)))]
    $p80 = $sorted[[math]::Min($sorted.Count - 1, [int][math]::Ceiling(0.8 * ($sorted.Count - 1)))]
    $remGiB = $remaining / 1GB
    $loMin = Format-MbtGuiEtaMinutes ($remGiB / $p80 * 60); $hiMin = Format-MbtGuiEtaMinutes ($remGiB / $p20 * 60)
    $midMin = Format-MbtGuiEtaMinutes ($remGiB / (($p20 + $p80) / 2) * 60)
    $mid = Format-MbtGuiDuration ($midMin * 60)
    $range = if ($loMin -ne $hiMin) { " (between $(Format-MbtGuiDuration ($loMin * 60)) and $(Format-MbtGuiDuration ($hiMin * 60)))" } else { '' }
    $mean = ($good | Measure-Object -Property GiBPerHour -Average).Average
    $basis = ('Based on the last {0} minutes: {1} GiB per hour.' -f ($good.Count * 5), ([math]::Round($mean, 1)).ToString([cultureinfo]::InvariantCulture))
    if ($Task -eq 'import') { $basis += ' Microsoft limits how fast one mailbox accepts items; we cannot go faster.' }
    else { $basis += ' It varies with Microsoft''s load.' }
    [pscustomobject]@{ State = 'ready'; Text = "About $mid left$range"; Basis = $basis }
}

#endregion

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCn0yAyd7d0DouF
# IDKoeqrUTBYOAeew9ZjNGtvtjidfv6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC1N5dDFseO
# QE2ltMFPBweGoQzo/GO5zdVdVjSli5huCzANBgkqhkiG9w0BAQEFAASCAgBAOHGa
# wBw6Ppn3qw9jnc4+AvvyJo7RdgZudmfOjWwUBHiqh7xkLbarReBL9yasxobWOaL1
# 1FNGEmOMOm66ZjFoYi90y+4tdsUnSKO8iUWjCCnxtMBISO6itHQyWNK7V/bM8XCN
# +j+A/fWDvG0KGBFhwGLzIKX6utIfxT8HIWWLYNkiySY1P6HT301HczuHwfriuocX
# LQu/PH4vu0YuZ5T5ye1fG0GzaPxVNQeIrYcyk8YdOA/uWu1f+mM44Q6t4bdBjCFT
# GWWWHvpb6TpKARBtaSWmcJbNkGwIm9jMUSJETaY04kfPFw3kMiixMz9pf5v9gglj
# /bJZSQgRxaXMW3x0BJ59nGXZIJKaXzUkil3pqDjss1Y6LQ3zg4CFw1PtbIHjurF0
# 7PLxPPfG7lrN2xCbyjX9yFsDPWOfTUkpl3yO88+1xFaWA92MeuRQBJY3hOxdFXP7
# 17rOsLD3tFmCkAL7xoK/Ryf/nHhyiJhTIomYnNX4619+IpWMBAAtbTNLTQ+uG8SD
# QzQZxCrWcUbSrTui2w6Pt3HG63l2n9vgSJzBab2sDxIQ4vUSLU8zPEM/MqB9XmgG
# J8bBKZVH535rkhX+rvWOleHuW7l/SJ0MVmi3/QOaqYE0VRwq3cTVin8hJjmmrkPM
# gwSUGkkTUP2hnJ+ApF9jTWhSzBbdKp2bn3d3T6GCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIPLAhxgybwWYq4qEnrDuFfQZmeT+gWrhEdShhwuZb4vGAggWm703PeXm
# OBgPMjAyNjExMTAwODIyNDJaMAMCAQECECF67MLLEYTLbD9bjmtnBfygghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjI0MlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEICV+
# AfTvUi6Zn6seD19ia6faDQaCi2ez50319FAseHKKMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiBoAXAICQPkY9gNiJtneQx9woKjNDhEl1L3PjNcm/BEvgIgdyRs
# wfk/4JpAGeqH9U24twD/yuvHgX3r2PFsNqwzrDw=
# SIG # End signature block
