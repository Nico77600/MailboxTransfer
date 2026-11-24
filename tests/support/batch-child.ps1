#Requires -Version 7.6
<#
    A stand-in for Invoke-MailboxTransfer.ps1 in the tests of a batch: it takes the same arguments, behaves as the local part of the mailbox address says, writes the status.json, the events and the exit code
    of a job, and records what it was given (args.json) and when it ran (run-<n>.txt). Nothing reaches a network.

        ok-...       complete at once (exit 0)
        slow-...     runs until its stop request appears, then stops (exit 2, reason stopped); with "quick" in the name it ends alone after a second
        partial-...  not complete the first time (exit 2), complete the next time
        fail-...     an error (exit 1)
        signin-...   needs a sign-in (exit 2, reason signin-required)
        crash-...    ends with exit 1 before it writes anything
        asker-...    writes the file named by MBT_STUB_STOPFILE (a stop request of the batch), then completes
#>
$named = @{}
for ($i = 0; $i -lt $args.Count; $i++) {
    $t = [string]$args[$i]
    if ($t -notmatch '^-[A-Za-z]+$') { continue }
    if (($i + 1) -lt $args.Count -and [string]$args[$i + 1] -notmatch '^-[A-Za-z]+$') { $named[$t.Substring(1)] = [string]$args[$i + 1]; $i++ } else { $named[$t.Substring(1)] = $true }
}
$events = [string]$named.EventsFile
$folder = Split-Path $events -Parent
$upn = if ($named.ContainsKey('MailboxId')) { [string]$named.MailboxId } else { [string]$named.UserPrincipalName }
$local = ($upn -split '@')[0]
$utf8 = [Text.UTF8Encoding]::new($false)
$now = { [DateTime]::UtcNow.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z') }
[IO.File]::WriteAllText((Join-Path $folder 'args.json'), (($args | ForEach-Object { [string]$_ }) | ConvertTo-Json), $utf8)
$runs = @(Get-ChildItem -LiteralPath $folder -Filter 'run-*.txt' -ErrorAction SilentlyContinue).Count + 1
$began = [DateTime]::UtcNow
[IO.File]::WriteAllText((Join-Path $folder "run-$runs.txt"), $began.ToString('o'), $utf8)

$write = {
    param([string]$State, [int]$Exit, [string]$Reason, [hashtable]$Counts)
    $s = [ordered]@{ v = 1; jobId = [string]$named.JobId; task = 'export'; title = "Export - $upn"; mode = 'Export'; source = $upn; destination = [string]$named.StorePath; batch = $(if ($named.ContainsKey('BatchId')) { [string]$named.BatchId } else { $null })
        state = $State; exitCode = $Exit; counts = $Counts; next = $(if ($Exit -eq 0) { 'done' } else { 'continue' }); reason = $Reason; started = $began.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z'); updated = (& $now); lastEvent = (& $now); pid = $PID }
    [IO.File]::WriteAllText((Join-Path $folder 'status.json'), ($s | ConvertTo-Json -Depth 5), $utf8)
    [IO.File]::AppendAllText($events, ((@{ v = 1; t = (& $now); e = 'result'; exit = $Exit; complete = ($Exit -eq 0); counts = $Counts; reason = $Reason } | ConvertTo-Json -Compress -Depth 5) + "`n"), $utf8)
}
[IO.File]::AppendAllText($events, ((@{ v = 1; t = (& $now); e = 'plan'; job = [string]$named.JobId; task = 'export'; mode = 'Export'; steps = @(@{ id = 'tree'; title = 'Read the folder tree' }) } | ConvertTo-Json -Compress -Depth 5) + "`n"), $utf8)

if ($local -like 'crash-*') { exit 1 }
if ($local -like 'asker-*') {
    # asks for the stop of the batch (the file the test names), then ends
    Start-Sleep -Milliseconds 500
    if ($env:MBT_STUB_STOPFILE) { [IO.File]::WriteAllText($env:MBT_STUB_STOPFILE, 'stop') }
    & $write 'complete' 0 $null @{ done = 7; failed = 0; bytes = 70 }; exit 0
}
if ($local -like 'slow-*') {
    $until = [DateTime]::UtcNow.AddSeconds(60)
    while ([DateTime]::UtcNow -lt $until) {
        if ([IO.File]::Exists([string]$named.StopFile)) { & $write 'stopped' 2 'stopped' @{ done = 5; failed = 0 }; exit 2 }
        if ($local -like 'slow-quick*' -and ([DateTime]::UtcNow - $began).TotalSeconds -ge 1.2) { & $write 'complete' 0 $null @{ done = 10; failed = 0; bytes = 1000 }; exit 0 }
        Start-Sleep -Milliseconds 100
    }
    & $write 'resumable' 2 $null @{ done = 5; failed = 0 }; exit 2
}
if ($local -like 'fail-*') {
    [IO.File]::AppendAllText($events, ((@{ v = 1; t = (& $now); e = 'error'; code = 'mailbox-not-found'; msg = 'We cannot find this mailbox. Check the address.'; tech = 'ErrorMailboxNotFound' } | ConvertTo-Json -Compress) + "`n"), $utf8)
    & $write 'error' 1 $null @{ done = 0; failed = 0 }; exit 1
}
if ($local -like 'signin-*') { & $write 'resumable' 2 'signin-required' @{}; exit 2 }
if ($local -like 'partial-*' -and $runs -eq 1) { & $write 'resumable' 2 $null @{ done = 3; failed = 1 }; exit 2 }
Start-Sleep -Milliseconds ([int]$(if ($local -like 'ok-pause*') { 700 } else { 50 }))
& $write 'complete' 0 $null @{ done = 100 * $runs; failed = 0; bytes = 5000 }
exit 0
