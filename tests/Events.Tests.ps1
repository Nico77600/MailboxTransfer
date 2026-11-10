#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the contract with the console: the events of a job (JSON lines), the stop-request file, keep awake,
    status.json and the list of jobs, the capabilities of the engine, the progress of the PST check, and the entry script modes
    CheckPst / RepairPst / Status run in a child pwsh with -EventsFile. Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\Events.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\PstDamage.ps1')

    # ConvertFrom-Json turns ISO strings into dates; the events are checked as text (-DateKind String exists from PowerShell 7.5)
    $script:StringDates = (Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')
    function ConvertFrom-JsonText { param([Parameter(ValueFromPipeline)][string]$Text) process { if ($script:StringDates) { $Text | ConvertFrom-Json -DateKind String } else { $Text | ConvertFrom-Json } } }
    function Read-Events { param([string]$Path) @(Get-Content -LiteralPath $Path | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-JsonText }) }

    # Runs the entry script in a fresh pwsh; returns the exit code, the output and the events (when -EventsFile is among the arguments).
    function Invoke-EntryChild {
        param([string[]]$Arguments)
        $out = & pwsh -NoProfile -NonInteractive -File $script:Entry -DevelopmentMode @Arguments 2>&1 | Out-String
        $code = $LASTEXITCODE
        $ix = [array]::IndexOf($Arguments, '-EventsFile')
        $events = if ($ix -ge 0 -and (Test-Path -LiteralPath $Arguments[$ix + 1])) { Read-Events $Arguments[$ix + 1] } else { @() }
        [pscustomobject]@{ Exit = $code; Output = $out; Events = $events; EventsFile = $(if ($ix -ge 0) { $Arguments[$ix + 1] }) }
    }

    $script:Good = New-TestPst $TestDrive 'good'
    $script:Wrong = New-TestPst $TestDrive 'wrongcount' -Faults @('wrong-count')
}

Describe 'EventLog: the JSON lines' {
    It 'writes every kind as one valid JSON object per line, with v, t and e' {
        $path = Join-Path $TestDrive 'kinds.jsonl'
        $log = [MailboxTransfer.EventLog]::new($path, $null, 0)
        $log.Plan('imp-1', [Collections.Generic.KeyValuePair[string, string][]]@([Collections.Generic.KeyValuePair[string, string]]::new('check', 'Check the file'), [Collections.Generic.KeyValuePair[string, string]]::new('send', 'Send the items')))
        $log.Step('send', 'running', 5120, 9860, 26391000000, 50784000000, $null)
        $log.Object('send', 'Projects/Fabrikam Pilot', 'check', 777, 780, '3 uncertain', -1)
        $log.Item('reconcile', 'item 7,114', 'checking', 'Looking again in 3 min', $null)
        $log.Wait('throttle', 38, 7, 252)
        $log.Rate('5m', 3.9)
        $log.Warn('UNKNOWN_ROW', 'send', 'item 7,114', 'no answer')
        $log.Decision('42', 'Two items could be this one.', [string[]]@('id1', 'id2'))
        $log.Error('disk-full', 'The disk is full.', 'IOException ...')
        $log.Info('hello')
        $log.Result(2, $false, @{ done = 5120; failed = 2 }, 'continue', 'stopped')
        $log.Dispose()
        $lines = @(Get-Content -LiteralPath $path)
        $lines.Count | Should -Be 11
        $ev = @($lines | ForEach-Object { $_ | ConvertFrom-JsonText })
        $ev.e | Should -Be @('plan', 'step', 'object', 'item', 'wait', 'rate', 'warn', 'decision', 'error', 'info', 'result')
        $ev | ForEach-Object { $_.v | Should -Be 1; $_.t | Should -Match '^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$' }
        $ev[0].job | Should -Be 'imp-1'
        $ev[0].steps.Count | Should -Be 2
        $ev[0].steps[1].id | Should -Be 'send'
        $ev[0].steps[1].title | Should -Be 'Send the items'
        $ev[1].id | Should -Be 'send'; $ev[1].state | Should -Be 'running'; $ev[1].done | Should -Be 5120; $ev[1].total | Should -Be 9860; $ev[1].bytesTotal | Should -Be 50784000000
        $ev[2].key | Should -Be 'Projects/Fabrikam Pilot'; $ev[2].msg | Should -Be '3 uncertain'
        $ev[4].seconds | Should -Be 38; $ev[4].count | Should -Be 7; $ev[4].waited | Should -Be 252
        $ev[5].gibPerHour | Should -Be 3.9
        $ev[7].candidates | Should -Be @('id1', 'id2')
        $ev[8].code | Should -Be 'disk-full'; $ev[8].tech | Should -Be 'IOException ...'
        $ev[10].exit | Should -Be 2; $ev[10].complete | Should -BeFalse; $ev[10].counts.done | Should -Be 5120; $ev[10].next | Should -Be 'continue'; $ev[10].reason | Should -Be 'stopped'
    }

    It 'leaves out the counters that are unknown: unknown is not zero' {
        $path = Join-Path $TestDrive 'unknown.jsonl'
        $log = [MailboxTransfer.EventLog]::new($path, $null, 0)
        $log.Step('list', 'running', -1, -1, -1, -1, $null)
        $log.Dispose()
        $e = Get-Content -LiteralPath $path | ConvertFrom-JsonText
        $e.PSObject.Properties.Name | Should -Not -Contain 'done'
        $e.PSObject.Properties.Name | Should -Not -Contain 'total'
        $e.PSObject.Properties.Name | Should -Not -Contain 'bytes'
    }

    It 'takes lines from several threads without mixing them' {
        $path = Join-Path $TestDrive 'threads.jsonl'
        $log = [MailboxTransfer.EventLog]::new($path, $null, 0)
        1..8 | ForEach-Object -Parallel { $l = $using:log; foreach ($i in 1..100) { $l.Info("thread $_ line $i " + ('x' * 200)) } } -ThrottleLimit 8
        $log.Dispose()
        $lines = @(Get-Content -LiteralPath $path)
        $lines.Count | Should -Be 800
        $bad = @($lines | Where-Object { try { $null = $_ | ConvertFrom-Json; $false } catch { $true } })
        $bad.Count | Should -Be 0
    }

    It 'is read by several readers while it is written, and appends to an existing file' {
        $path = Join-Path $TestDrive 'append.jsonl'
        $a = [MailboxTransfer.EventLog]::new($path, $null, 0); $a.Info('one')
        $first = [MailboxTransfer.EventTail]::Read($path, 0)
        $first.Lines.Count | Should -Be 1
        $a.Info('two')
        $b = [MailboxTransfer.EventLog]::new($path, $null, 0); $b.Info('three')
        $a.Dispose(); $b.Dispose()
        $more = [MailboxTransfer.EventTail]::Read($path, $first.Offset)
        $more.Lines.Count | Should -Be 2
        ($more.Lines[1] | ConvertFrom-Json).msg | Should -Be 'three'
    }

    It 'writes nothing after Dispose and does not throw' {
        $path = Join-Path $TestDrive 'disposed.jsonl'
        $log = [MailboxTransfer.EventLog]::new($path, $null, 0)
        $log.Info('a'); $log.Dispose(); $log.Dispose()
        { $log.Info('b') } | Should -Not -Throw
        @(Get-Content -LiteralPath $path).Count | Should -Be 1
    }

    It 'reports a sleep when the clock jumps between two heartbeats, not for a small gap' {
        $path = Join-Path $TestDrive 'slept.jsonl'
        $script:fake = [datetime]::new(2026, 11, 9, 2, 0, 0, [DateTimeKind]::Utc)
        $log = [MailboxTransfer.EventLog]::new($path, [Func[datetime]]{ $script:fake }, 0, 120)
        $log.Info('before')
        $script:fake = $script:fake.AddSeconds(60); $log.Heartbeat()
        $script:fake = $script:fake.AddSeconds(30); $log.Info('still awake')
        $script:fake = $script:fake.AddHours(4).AddMinutes(34); $log.Heartbeat()
        $log.Dispose()
        $ev = Read-Events $path
        $ev.e | Should -Be @('info', 'info', 'slept')
        $ev[2].seconds | Should -Be 16440
        $ev[2].from | Should -Match '^2026-11-09T02:01:30'
        $ev[2].to | Should -Match '^2026-11-09T06:35:30'
    }

    It 'reports a sleep before the next line when a line is written after the wake-up' {
        $path = Join-Path $TestDrive 'slept2.jsonl'
        $script:fake = [datetime]::new(2026, 11, 9, 2, 0, 0, [DateTimeKind]::Utc)
        $log = [MailboxTransfer.EventLog]::new($path, [Func[datetime]]{ $script:fake }, 0, 120)
        $log.Info('a')
        $script:fake = $script:fake.AddMinutes(30); $log.Step('send', 'running', 1, 2, -1, -1, $null)
        $log.Dispose()
        (Read-Events $path).e | Should -Be @('info', 'slept', 'step')
    }
}

Describe 'EventTail' {
    It 'ignores a last line that is not complete and gives it when it is' {
        $path = Join-Path $TestDrive 'tail.jsonl'
        [IO.File]::WriteAllText($path, "{`"a`":1}`n{`"a`":2}`n{`"a`":")
        $r = [MailboxTransfer.EventTail]::Read($path, 0)
        $r.Lines | Should -Be @('{"a":1}', '{"a":2}')
        $r.Offset | Should -Be 16
        [IO.File]::AppendAllText($path, "3}`r`n")
        $r2 = [MailboxTransfer.EventTail]::Read($path, $r.Offset)
        $r2.Lines | Should -Be @('{"a":3}')
        ([MailboxTransfer.EventTail]::Read($path, $r2.Offset)).Lines.Count | Should -Be 0
    }
    It 'gives nothing for a missing file and starts again from zero when the file was replaced by a shorter one' {
        ([MailboxTransfer.EventTail]::Read((Join-Path $TestDrive 'nope.jsonl'), 5)).Lines.Count | Should -Be 0
        $path = Join-Path $TestDrive 'tail2.jsonl'
        [IO.File]::WriteAllText($path, "{`"a`":1}`n")
        $r = [MailboxTransfer.EventTail]::Read($path, 500)
        $r.Lines | Should -Be @('{"a":1}')
    }
    It 'reads UTF-8 text' {
        $path = Join-Path $TestDrive 'tail3.jsonl'
        $log = [MailboxTransfer.EventLog]::new($path, $null, 0); $log.Info("caf$([char]0xE9) $([char]0x2014) ok"); $log.Dispose()
        (([MailboxTransfer.EventTail]::Read($path, 0)).Lines[0] | ConvertFrom-Json).msg | Should -Be "caf$([char]0xE9) $([char]0x2014) ok"
    }
}

Describe 'StopWatcher' {
    It 'cancels the token within a few seconds when the file appears' {
        $f = Join-Path $TestDrive 'stop.request'
        $w = [MailboxTransfer.StopWatcher]::new($f, 200)
        try {
            $w.IsRequested | Should -BeFalse
            Set-Content -LiteralPath $f -Value 'stop'
            $w.Token.WaitHandle.WaitOne(3000) | Should -BeTrue
            $w.IsRequested | Should -BeTrue
        } finally { $w.Dispose() }
    }
    It 'sees a file that exists already when it starts' {
        $f = Join-Path $TestDrive 'stop2.request'
        Set-Content -LiteralPath $f -Value 'stop'
        $w = [MailboxTransfer.StopWatcher]::new($f)
        try { $w.IsRequested | Should -BeTrue } finally { $w.Dispose() }
    }
    It 'is driven by Start-MbtStopWatch / Test-MbtStopRequested / Get-MbtStopToken' {
        Test-MbtStopRequested | Should -BeFalse
        (Get-MbtStopToken).CanBeCanceled | Should -BeFalse
        $f = Join-Path $TestDrive 'stop3.request'
        Start-MbtStopWatch -Path $f -PollMilliseconds 100
        try {
            Test-MbtStopRequested | Should -BeFalse
            (Get-MbtStopToken).CanBeCanceled | Should -BeTrue
            Set-Content -LiteralPath $f -Value 'stop'
            (Get-MbtStopToken).WaitHandle.WaitOne(3000) | Should -BeTrue
            Test-MbtStopRequested | Should -BeTrue
        } finally { Stop-MbtStopWatch }
        Test-MbtStopRequested | Should -BeFalse
    }
}

Describe 'AwakeKeeper' {
    It 'holds the execution state until it is disposed, and releases it' {
        $k = [MailboxTransfer.AwakeKeeper]::Begin()
        $k.Succeeded | Should -BeTrue
        $k.IsActive | Should -BeTrue
        $k.Dispose()
        $k.IsActive | Should -BeFalse
        # the state the thread had when it let go: ES_CONTINUOUS | ES_SYSTEM_REQUIRED, so it was held
        ([uint32]$k.StateBeforeRelease -band [uint32]2147483649) | Should -Be ([uint32]2147483649)
        { $k.Dispose() } | Should -Not -Throw
    }
    It 'Start-MbtKeepAwake is idempotent and Stop-MbtKeepAwake releases it' {
        $a = Start-MbtKeepAwake
        $b = Start-MbtKeepAwake
        $b | Should -Be $a
        $a.IsActive | Should -BeTrue
        Stop-MbtKeepAwake
        $a.IsActive | Should -BeFalse
        { Stop-MbtKeepAwake } | Should -Not -Throw
    }
}

Describe 'Event log of the module, status.json and the job list' {
    AfterEach { Close-MbtEventLog }
    It 'Write-MbtEvent does nothing when no log is open' {
        { Write-MbtEvent info @{ msg = 'x' } } | Should -Not -Throw
        { Write-MbtJobStep -Id a -State running } | Should -Not -Throw
        { Write-MbtJobResult -ExitCode 0 } | Should -Not -Throw
    }
    It 'status.json of a running job follows the job: updated, last event and items done are refreshed (at most every 15 seconds), and nothing changes it after the result' {
        $dir = Join-Path $TestDrive 'jobs\20261109-1900-import-cd34'
        Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -JobId 'imp-cd34' -Task import -Title 'Import - adele' -Mode ImportPst -Store 'D:\store'
        $read = { Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-JsonText }
        $start = & $read
        $start.counts.done | Should -BeNullOrEmpty
        Write-MbtJobPlan -Steps @(@{ Id = 'send'; Title = 'Send' })
        # a step event right after the start does not rewrite the file (throttled)
        Write-MbtJobStep -Id send -State running -Done 10 -Total 100
        (& $read).updated | Should -Be $start.updated
        # fifteen seconds later the next step event refreshes it
        InModuleScope MailboxTransfer { $script:JobInfo.StatusAt = [DateTime]::UtcNow.AddSeconds(-20) }
        Start-Sleep -Milliseconds 20
        Write-MbtJobStep -Id send -State running -Done 4200 -Total 9860
        $s = & $read
        $s.state | Should -Be 'running'
        $s.counts.done | Should -Be 4200; $s.counts.total | Should -Be 9860
        [DateTime]::Parse($s.updated).ToUniversalTime() | Should -BeGreaterThan ([DateTime]::Parse($start.updated).ToUniversalTime())
        $s.jobId | Should -Be 'imp-cd34'; $s.title | Should -Be 'Import - adele'; $s.started | Should -Be $start.started
        # an object event also refreshes it, without counts (the counts of the step stay)
        InModuleScope MailboxTransfer { $script:JobInfo.StatusAt = [DateTime]::UtcNow.AddSeconds(-20) }
        Write-MbtJobObject -Step send -Key 'Inbox' -State running -Done 5 -Total 10
        (& $read).counts.done | Should -Be 4200
        # the result closes the job: a late step event never turns it back to running
        Write-MbtJobResult -ExitCode 0 -Counts @{ done = 9860; failed = 0 }
        InModuleScope MailboxTransfer { $script:JobInfo.StatusAt = [DateTime]::UtcNow.AddSeconds(-20) }
        Write-MbtJobStep -Id send -State running -Done 1 -Total 2
        $end = & $read
        $end.state | Should -Be 'complete'; $end.counts.done | Should -Be 9860
    }
    It 'Open-MbtEventLog writes status.json (running) and the result updates it' {
        $dir = Join-Path $TestDrive 'jobs\20261109-1402-export-ab12'
        Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -JobId 'exp-ab12' -Task export -Title 'Export - adele' -Mode Export -Store 'D:\store'
        $s = Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-JsonText
        $s.state | Should -Be 'running'; $s.jobId | Should -Be 'exp-ab12'; $s.task | Should -Be 'export'; $s.store | Should -Be 'D:\store'; $s.pid | Should -Be $PID
        Write-MbtJobPlan -Steps @(@{ Id = 'tree'; Title = 'Read the folder tree' }, @{ Id = 'list'; Title = 'List the items you chose' })
        Write-MbtJobStep -Id tree -State done -Done 10 -Total 10
        Write-MbtJobResult -ExitCode 2 -Counts @{ done = 5; deferred = 2 } -Reason deferred
        $s = Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-JsonText
        $s.state | Should -Be 'resumable'; $s.exitCode | Should -Be 2; $s.counts.done | Should -Be 5; $s.next | Should -Be 'continue'; $s.reason | Should -Be 'deferred'
        $s.jobId | Should -Be 'exp-ab12'
        $s.lastEvent | Should -Match '^\d{4}-'
        Close-MbtEventLog
        $ev = Read-Events (Join-Path $dir 'events.jsonl')
        $ev.e | Should -Be @('plan', 'step', 'result')
        $ev[0].task | Should -Be 'export'
        $ev[2].next | Should -Be 'continue'
    }
    It 'the result maps the exit code to a state and a next action' {
        $cases = @(@(0, 'complete', 'done'), @(3, 'needs-decision', 'decide'), @(1, 'error', 'retry'))
        foreach ($c in $cases) {
            $dir = Join-Path $TestDrive "jobs\map-$($c[0])"
            Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -Task import
            Write-MbtJobResult -ExitCode $c[0]
            Close-MbtEventLog
            $s = Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json
            $s.state | Should -Be $c[1]
            $s.next | Should -Be $c[2]
        }
        $dir = Join-Path $TestDrive 'jobs\map-stop'
        Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -Task import
        Write-MbtJobResult -ExitCode 2 -Reason stopped
        Close-MbtEventLog
        (Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json).state | Should -Be 'stopped'
    }
    It 'Get-MbtJob lists the jobs of a work folder, newest first, and finds the interrupted ones' {
        $work = Join-Path $TestDrive 'work'
        $old = Join-Path $work 'jobs\a-old'; New-Item -ItemType Directory $old -Force | Out-Null
        @{ v = 1; jobId = 'a-old'; task = 'import'; title = 'Import'; state = 'running'; pid = 2147483000; started = '2026-11-08T22:00:00.000Z'; updated = '2026-11-08T22:41:00.000Z'; counts = @{ done = 5120 } } | ConvertTo-Json | Set-Content (Join-Path $old 'status.json')
        $fin = Join-Path $work 'jobs\b-finished'; New-Item -ItemType Directory $fin -Force | Out-Null
        @{ v = 1; jobId = 'b-finished'; task = 'export'; title = 'Export'; state = 'complete'; exitCode = 0; started = '2026-11-02T10:00:00.000Z'; updated = '2026-11-02T11:00:00.000Z' } | ConvertTo-Json | Set-Content (Join-Path $fin 'status.json')
        Open-MbtEventLog -Path (Join-Path $work 'jobs\c-live\events.jsonl') -JobId 'c-live' -Task check -Title 'Check'
        New-Item -ItemType Directory (Join-Path $work 'jobs\d-nostatus') -Force | Out-Null
        $jobs = @(Get-MbtJob -WorkFolder $work)
        $jobs.Count | Should -Be 3
        $jobs[0].JobId | Should -Be 'c-live'
        $jobs[0].Alive | Should -BeTrue
        $jobs[0].EffectiveState | Should -Be 'running'
        $jobs[1].JobId | Should -Be 'a-old'
        $jobs[1].State | Should -Be 'running'
        $jobs[1].EffectiveState | Should -Be 'resumable'
        $jobs[1].Interrupted | Should -BeTrue
        $jobs[1].Counts.done | Should -Be 5120
        $jobs[2].JobId | Should -Be 'b-finished'; $jobs[2].EffectiveState | Should -Be 'complete'
        @(Get-MbtJob -WorkFolder (Join-Path $TestDrive 'no-work')).Count | Should -Be 0
    }
}

Describe 'Plain error sentences' {
    It '<Code> for "<Text>"' -ForEach @(
        @{ Text = 'Not signed in: run the script interactively once.'; Code = 'signin-required' }
        @{ Text = 'AADSTS50173: The provided grant has expired'; Code = 'signin-required' }
        @{ Text = 'There is not enough space on the disk. (store pack-000003.bin)'; Code = 'disk-full' }
        @{ Text = "Repair target 'aux.pst' cannot be used: a Windows device name"; Code = 'reserved-name' }
        @{ Text = 'The target exists (use Overwrite).'; Code = 'destination-exists' }
        @{ Text = 'Graph answered 403 ErrorAccessDenied'; Code = 'access-denied' }
        @{ Text = 'something odd'; Code = 'error' }
    ) {
        $p = ConvertTo-MbtPlainError -Message $Text
        $p.Code | Should -Be $Code
        $p.Message | Should -Not -BeNullOrEmpty
        $p.Message | Should -Not -Match 'Exception|AADSTS'
    }
}

Describe 'Progress of the long jobs' {
    AfterEach { Close-MbtEventLog }
    It 'Update-MbtJobProgress writes steps, throttling waits, rate windows and warnings once each' {
        $dir = Join-Path $TestDrive 'jobs\fakejob'
        Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -Task import
        Write-MbtJobPlan -Steps @(@{ Id = 'check'; Title = 'Verify' }, @{ Id = 'folders'; Title = 'Create the folders' }, @{ Id = 'send'; Title = 'Send the items' }, @{ Id = 'reconcile'; Title = 'Check the uncertain items' }, @{ Id = 'final'; Title = 'Final check' })
        Write-MbtJobStep -Id check -State done
        $job = [pscustomobject]@{ Phase = 'folders'; Cancelled = $false; Error = $null; Throttles = @(); Windows = @() }
        $job | Add-Member ScriptMethod Snapshot { [pscustomobject]@{ Phase = $this.Phase; Items = 5; TotalItems = 10; Bytes = 100; TotalBytes = 1000 } }
        $job | Add-Member ScriptMethod GetThrottleEvents { , $this.Throttles }
        $job | Add-Member ScriptMethod GetWindows { , $this.Windows }
        InModuleScope MailboxTransfer -Parameters @{ Job = $job } { Update-MbtJobProgress -Job $Job }
        $job.Phase = 'import'
        $job.Throttles = @([pscustomobject]@{ WaitSeconds = 38.0 }, [pscustomobject]@{ WaitSeconds = 12.0 })
        $job.Windows = @([pscustomobject]@{ GiBPerHour = 3.94 })
        InModuleScope MailboxTransfer -Parameters @{ Job = $job } { Update-MbtJobProgress -Job $Job -Messages @('item 7114 failed: 504', 'folder done') }
        InModuleScope MailboxTransfer -Parameters @{ Job = $job } { Update-MbtJobProgress -Job $Job }
        InModuleScope MailboxTransfer -Parameters @{ Job = $job } { Update-MbtJobProgress -Job $Job -Final }
        Close-MbtEventLog
        $ev = Read-Events (Join-Path $dir 'events.jsonl')
        $steps = @($ev | Where-Object e -EQ 'step')
        ($steps | Where-Object { $_.id -eq 'folders' -and $_.state -eq 'running' }).Count | Should -Be 1
        ($steps | Where-Object { $_.id -eq 'folders' -and $_.state -eq 'done' }).Count | Should -BeGreaterOrEqual 1
        ($steps | Where-Object { $_.id -eq 'send' -and $_.state -eq 'running' -and $_.done -eq 5 -and $_.total -eq 10 }).Count | Should -BeGreaterOrEqual 1
        ($steps | Where-Object { $_.id -eq 'send' -and $_.state -eq 'done' }).Count | Should -Be 1
        $waits = @($ev | Where-Object e -EQ 'wait')
        $waits.Count | Should -Be 2
        $waits[1].count | Should -Be 2; $waits[1].waited | Should -Be 50; $waits[0].reason | Should -Be 'throttle'
        @($ev | Where-Object e -EQ 'rate').Count | Should -Be 1
        (@($ev | Where-Object e -EQ 'rate')[0]).gibPerHour | Should -Be 3.94
        $warns = @($ev | Where-Object e -EQ 'warn')
        $warns.Count | Should -Be 1
        $warns[0].msg | Should -Match '504'
    }
}

Describe 'Get-MbtCapability' {
    It 'reads the real entry script' {
        $c = Get-MbtCapability
        $c.Events | Should -BeTrue
        $c.StopFile | Should -BeTrue
        $c.KeepAwake | Should -BeTrue
        $c.PstExport | Should -BeTrue
        $c.PstCheck | Should -BeTrue
        $c.PstRepair | Should -BeTrue
        $c.Modes | Should -Contain 'Export'
        # the merged engine has all of it: the console lights up the PST import, the EML export and import and the shared mailboxes
        foreach ($n in 'PstImport', 'EmlExport', 'EmlImport', 'SharedMailbox') { $c.$n | Should -BeTrue }
        $c.Modes | Should -Contain 'ImportPst'; $c.Modes | Should -Contain 'ExportEml'; $c.Modes | Should -Contain 'ImportEml'
    }
    It 'decides from the modes and parameters of a script: ImportPst with -PstPath lights up the PST import, ExportEml / ImportEml the EML, -MailboxId the shared mailboxes' {
        $bare = Join-Path $TestDrive 'bare.ps1'
        Set-Content $bare "param([ValidateSet('Status','Export')][string]`$Mode)"
        $full = Join-Path $TestDrive 'full.ps1'
        Set-Content $full "param([ValidateSet('Status','ExportEml','CheckPst','ImportPst')][string]`$Mode, [string]`$PstPath, [string]`$MailboxId, [string]`$EventsFile)"
        # a script with the mode but not the parameter that carries the file cannot import a PST
        $noPath = Join-Path $TestDrive 'nopath.ps1'
        Set-Content $noPath "param([ValidateSet('Status','ImportPst')][string]`$Mode)"
        $b = Get-MbtCapability -EntryScript $bare
        $b.Events | Should -BeFalse; $b.PstCheck | Should -BeFalse; $b.PstImport | Should -BeFalse; $b.EmlExport | Should -BeFalse; $b.EmlImport | Should -BeFalse; $b.SharedMailbox | Should -BeFalse
        $f = Get-MbtCapability -EntryScript $full
        $f.Events | Should -BeTrue; $f.PstCheck | Should -BeTrue; $f.PstImport | Should -BeTrue; $f.EmlExport | Should -BeTrue; $f.EmlImport | Should -BeFalse; $f.SharedMailbox | Should -BeTrue; $f.PstRepair | Should -BeFalse
        (Get-MbtCapability -EntryScript $noPath).PstImport | Should -BeFalse
    }
}

Describe 'Test-MbtPst: progress and cancellation' {
    It '-OnProgress receives the phases, with the bytes of the file' {
        $script:seen = [Collections.Generic.List[object]]::new()
        $r = Test-MbtPst -Path $script:Good -OnProgress { param($e) $script:seen.Add([pscustomobject]@{ Phase = $e.Phase; Total = $e.BytesTotal; Done = $e.BytesDone }) }
        $r.ExitCode | Should -Be 0
        $r.Cancelled | Should -BeFalse
        $script:seen.Count | Should -BeGreaterThan 3
        ($script:seen | Where-Object Total -GT 0).Count | Should -BeGreaterThan 0
        ($script:seen.Phase | Select-Object -Unique).Count | Should -BeGreaterThan 3
    }
    It 'a cancelled token stops the check: Cancelled, and no exception' {
        $cts = [Threading.CancellationTokenSource]::new(); $cts.Cancel()
        $r = Test-MbtPst -Path $script:Good -CancellationToken $cts.Token
        $r.Cancelled | Should -BeTrue
        $cts.Dispose()
    }
    It 'with an event log open for a check job, the default hook writes step events' {
        $dir = Join-Path $TestDrive 'jobs\hook'
        Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -Task check -Title 'Check'
        try {
            Write-MbtJobPlan -Steps @(@{ Id = 'check'; Title = 'Check the file' })
            $null = Test-MbtPst -Path $script:Good
        } finally { Close-MbtEventLog }
        $steps = @(Read-Events (Join-Path $dir 'events.jsonl') | Where-Object e -EQ 'step')
        $steps.Count | Should -BeGreaterThan 2
        $steps | ForEach-Object { $_.id | Should -Be 'check'; $_.state | Should -Be 'running' }
        ($steps | Where-Object { $_.bytesTotal -gt 0 }).Count | Should -BeGreaterThan 0
    }
}

Describe 'Repair-MbtPst -Level Automatic' {
    It 'checks once, repairs at the recommended level and reports before and after' {
        $out = Join-Path $TestDrive 'auto-out.pst'
        $r = Repair-MbtPst -Path $script:Wrong -Destination $out -Level Automatic
        $r.Success | Should -BeTrue
        $r.Recommended | Should -Be 'Recompute'
        $r.Before.Errors + $r.Before.Warnings | Should -BeGreaterThan 0
        $r.After.Errors | Should -Be 0
        $r.ItemsRecovered | Should -Be 56
        (Test-MbtPst -Path $out).ExitCode | Should -Be 0
    }
}

Describe 'Invoke-MailboxTransfer.ps1 -EventsFile (child pwsh)' {
    It 'Mode Status writes a plan and a result with exit 0, and status.json' {
        $ev = Join-Path $TestDrive 'status-job\events.jsonl'
        $x = Invoke-EntryChild @('-Mode', 'Status', '-StorePath', (Join-Path $TestDrive 'nostore'), '-EventsFile', $ev, '-JobId', 'st-1')
        $x.Exit | Should -Be 0
        $x.Events.e | Should -Be @('plan', 'result')
        $x.Events[0].job | Should -Be 'st-1'
        $x.Events[0].steps.Count | Should -Be 1
        $x.Events[1].exit | Should -Be 0; $x.Events[1].complete | Should -BeTrue; $x.Events[1].next | Should -Be 'done'
        $s = Get-Content (Join-Path (Split-Path $ev) 'status.json') -Raw | ConvertFrom-Json
        $s.state | Should -Be 'complete'; $s.jobId | Should -Be 'st-1'; $s.exitCode | Should -Be 0
        $x.Output | Should -Match 'the C# engine compiles'
    }

    It 'Mode CheckPst on a clean PST: plan, step events with bytes, result exit 0' {
        $ev = Join-Path $TestDrive 'check-clean\events.jsonl'
        $x = Invoke-EntryChild @('-Mode', 'CheckPst', '-Path', $script:Good, '-EventsFile', $ev, '-KeepAwake')
        $x.Exit | Should -Be 0
        $x.Events[0].e | Should -Be 'plan'
        $x.Events[0].steps.Count | Should -Be 1
        $x.Events[0].steps[0].id | Should -Be 'check'
        $steps = @($x.Events | Where-Object e -EQ 'step')
        $steps.Count | Should -BeGreaterThan 2
        ($steps | Where-Object { $_.bytesTotal -gt 0 }).Count | Should -BeGreaterThan 0
        $steps[-1].state | Should -Be 'done'
        $res = $x.Events[-1]
        $res.e | Should -Be 'result'; $res.exit | Should -Be 0; $res.complete | Should -BeTrue; $res.reason | Should -Be 'clean'
        $res.counts.errors | Should -Be 0
        (Get-Content (Join-Path (Split-Path $ev) 'status.json') -Raw | ConvertFrom-Json).state | Should -Be 'complete'
    }

    It 'Mode CheckPst on a damaged PST: exit code of the checker, one item event per finding code, next = repair, the file untouched' {
        $before = (Get-FileHash -LiteralPath $script:Wrong).Hash
        $ev = Join-Path $TestDrive 'check-bad\events.jsonl'
        $x = Invoke-EntryChild @('-Mode', 'CheckPst', '-Path', $script:Wrong, '-EventsFile', $ev)
        $x.Exit | Should -BeIn 1, 2
        $items = @($x.Events | Where-Object e -EQ 'item')
        $items.Count | Should -BeGreaterThan 0
        $items[0].step | Should -Be 'check'; $items[0].key | Should -Not -BeNullOrEmpty; $items[0].state | Should -BeIn 'error', 'warning'; $items[0].count | Should -BeGreaterThan 0
        $res = $x.Events[-1]
        $res.exit | Should -Be $x.Exit
        $res.next | Should -Be 'repair'
        $res.reason | Should -BeIn 'errors', 'warnings'
        (Get-FileHash -LiteralPath $script:Wrong).Hash | Should -Be $before
    }

    It 'a stop-request file that exists makes a check stop: exit 2, reason stopped, next continue' {
        $stop = Join-Path $TestDrive 'stop-now.request'; Set-Content $stop 'stop'
        $ev = Join-Path $TestDrive 'check-stop\events.jsonl'
        $x = Invoke-EntryChild @('-Mode', 'CheckPst', '-Path', $script:Good, '-EventsFile', $ev, '-StopFile', $stop)
        $x.Exit | Should -Be 2
        $res = $x.Events[-1]
        $res.exit | Should -Be 2; $res.complete | Should -BeFalse; $res.reason | Should -Be 'stopped'; $res.next | Should -Be 'continue'
        (Get-Content (Join-Path (Split-Path $ev) 'status.json') -Raw | ConvertFrom-Json).state | Should -Be 'stopped'
    }

    It 'Mode RepairPst (Automatic) writes the new file, never touches the original, and reports before and after' {
        $before = (Get-FileHash -LiteralPath $script:Wrong).Hash
        $dest = Join-Path $TestDrive 'repaired-entry.pst'
        $ev = Join-Path $TestDrive 'repair\events.jsonl'
        $x = Invoke-EntryChild @('-Mode', 'RepairPst', '-Path', $script:Wrong, '-Destination', $dest, '-EventsFile', $ev)
        $x.Exit | Should -Be 0
        Test-Path $dest | Should -BeTrue
        (Get-FileHash -LiteralPath $script:Wrong).Hash | Should -Be $before
        (Test-MbtPst -Path $dest).ExitCode | Should -Be 0
        $x.Events[0].steps.id | Should -Be @('before', 'scan', 'write', 'after')
        $steps = @($x.Events | Where-Object e -EQ 'step')
        foreach ($id in 'before', 'scan', 'write', 'after') { ($steps | Where-Object id -EQ $id).Count | Should -BeGreaterThan 0 }
        ($steps | Where-Object { $_.id -eq 'before' })[-1].state | Should -Be 'done'
        ($steps | Where-Object { $_.id -eq 'after' })[-1].state | Should -Be 'done'
        $res = $x.Events[-1]
        $res.exit | Should -Be 0; $res.counts.recovered | Should -Be 56; $res.counts.lost | Should -Be 0
        $res.counts.after.errors | Should -Be 0
        ($res.counts.before.errors + $res.counts.before.warnings) | Should -BeGreaterThan 0
    }

    It 'Mode RepairPst -WhatIf writes nothing; a reserved file name is refused with an error event and exit 1' {
        $dest = Join-Path $TestDrive 'whatif.pst'
        $x = Invoke-EntryChild @('-Mode', 'RepairPst', '-Path', $script:Wrong, '-Destination', $dest, '-WhatIf', '-EventsFile', (Join-Path $TestDrive 'repair-whatif\events.jsonl'))
        $x.Exit | Should -Be 0
        Test-Path $dest | Should -BeFalse
        $bad = Invoke-EntryChild @('-Mode', 'RepairPst', '-Path', $script:Wrong, '-Destination', (Join-Path $TestDrive 'aux.pst'), '-EventsFile', (Join-Path $TestDrive 'repair-aux\events.jsonl'))
        $bad.Exit | Should -Be 1
        Test-Path (Join-Path $TestDrive 'aux.pst') | Should -BeFalse
        $err = @($bad.Events | Where-Object e -EQ 'error')
        $err.Count | Should -Be 1
        $err[0].code | Should -Be 'reserved-name'
        $err[0].msg | Should -Not -BeNullOrEmpty
        $err[0].tech | Should -Match 'aux'
        $bad.Events[-1].e | Should -Be 'result'; $bad.Events[-1].exit | Should -Be 1; $bad.Events[-1].next | Should -Be 'retry'
        (Get-Content (Join-Path $TestDrive 'repair-aux\status.json') -Raw | ConvertFrom-Json).state | Should -Be 'error'
    }

    It 'a missing file: error event, result exit 1' {
        $x = Invoke-EntryChild @('-Mode', 'CheckPst', '-Path', (Join-Path $TestDrive 'missing.pst'), '-EventsFile', (Join-Path $TestDrive 'check-missing\events.jsonl'))
        $x.Exit | Should -Be 1
        ($x.Events | Where-Object e -EQ 'error').Count | Should -Be 1
        $x.Events[-1].exit | Should -Be 1
    }

    It '-NoPrompt without a cached sign-in: exit 2, result reason signin-required (no browser opens)' -Skip:(-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        $cfg = Join-Path $TestDrive 'nocache.psd1'
        Set-Content $cfg "@{ Auth = @{ CacheFile = '$((Join-Path $TestDrive 'empty-token.cache').Replace("'", "''"))' }; Output = @{ Root = '$($TestDrive.Replace("'", "''"))' } }"
        $ev = Join-Path $TestDrive 'noprompt\events.jsonl'
        # the child sees an empty profile: no token cache of the operator can be found, so nothing can sign in or reach the cloud
        $saved = $env:LOCALAPPDATA
        try { $env:LOCALAPPDATA = Join-Path $TestDrive 'empty-localappdata'; $x = Invoke-EntryChild @('-Mode', 'Export', '-ConfigPath', $cfg, '-StorePath', (Join-Path $TestDrive 'np-store'), '-NoPrompt', '-EventsFile', $ev) } finally { $env:LOCALAPPDATA = $saved }
        $x.Exit | Should -Be 2
        $x.Events[-1].e | Should -Be 'result'
        $x.Events[-1].reason | Should -Be 'signin-required'
        $x.Events[-1].next | Should -Be 'signin'
        (@($x.Events | Where-Object e -EQ 'error')[0]).code | Should -Be 'signin-required'
    }
}

Describe 'Wait-MbtTransferJob and a stop request' {
    AfterEach { Stop-MbtStopWatch; Close-MbtEventLog }
    It 'cancels the job when the stop file appears, says so, and goes on until the job ends' {
        $stop = Join-Path $TestDrive 'wait-stop.request'
        Start-MbtStopWatch -Path $stop -PollMilliseconds 50
        Set-Content $stop 'stop'
        $null = (Get-MbtStopToken).WaitHandle.WaitOne(3000)
        $job = [pscustomobject]@{ Cancelled = $false; CancelCalls = 0; Error = $null; Waits = 0; IsCompleted = $true }
        $job | Add-Member ScriptMethod Wait { param($ms) $this.Waits++; if ($this.CancelCalls -gt 0) { return $true } else { Start-Sleep -Milliseconds 20; return $false } }
        $job | Add-Member ScriptMethod Cancel { $this.CancelCalls++; $this.Cancelled = $true }
        $job | Add-Member ScriptMethod TakeMessages { , @() }
        $job | Add-Member ScriptMethod Snapshot { [pscustomobject]@{ Phase = 'import'; Items = 1; TotalItems = 2; Bytes = 1; TotalBytes = 2 } }
        $job | Add-Member ScriptMethod GetThrottleEvents { , @() }
        $job | Add-Member ScriptMethod GetWindows { , @() }
        $out = Wait-MbtTransferJob -Job $job -IntervalSeconds 1 6>&1 | Out-String
        $job.CancelCalls | Should -Be 1
        $out | Should -Match 'stop requested'
    }
}

Describe 'Entry script against the fake service: events of an export and an import' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
        Import-TestFixtures -TestsRoot $PSScriptRoot
        $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-events-tests-" + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $script:Tmp
        $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $script:Src = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
        $script:Dst = "MBX:99999999-8888-7777-6666-555555555555@$($script:Tenant)"
        function New-Message([string]$Subject) {
            $b = [MailboxTransfer.Tests.FxBuilder]::new()
            $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $Subject)
            return $b.ToArray()
        }
        function New-Service {
            $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
            $null = $svc.AddMailbox($script:Src)
            $null = $svc.AddFolder($script:Src, 'F-INBOX', 'Inbox')
            $null = $svc.AddFolder($script:Src, 'F-PROJ', 'Projects', 'F-INBOX')
            foreach ($i in 1..6) { $null = $svc.AddItem($script:Src, 'F-INBOX', "I$i", (New-Message "Inbox $i"), "ck-I$i") }
            foreach ($i in 1..4) { $null = $svc.AddItem($script:Src, 'F-PROJ', "P$i", (New-Message "Project $i"), "ck-P$i") }
            $null = $svc.AddMailbox($script:Dst)
            $null = $svc.AddFolder($script:Dst, 'D-INBOX', 'Inbox')
            return $svc
        }
        function Invoke-InProcess {
            param($Service, [hashtable]$Parameters)
            $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
            $global:MbtEvClient = [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
            $global:LASTEXITCODE = $null
            $text = & $script:Entry @Parameters *>&1 | Out-String
            [pscustomobject]@{ Text = $text; Exit = $LASTEXITCODE }
        }
    }
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph $global:MbtEvClient
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtEvClient -Scope Global -ErrorAction SilentlyContinue; Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Export: plan, steps tree, list, download, verify, a result with the counts; then Import: plan objects (-WhatIf), steps, final step, result' {
        $svc = New-Service
        $store = Join-Path $script:Tmp 'store'
        $evx = Join-Path $script:Tmp 'job-export\events.jsonl'
        $x = Invoke-InProcess $svc @{ Mode = 'Export'; StorePath = $store; TempPath = (Join-Path $script:Tmp 'tmp'); PrimaryMailboxId = $script:Src; Concurrency = 2; EventsFile = $evx; JobId = 'exp-1' }
        $ev = Read-Events $evx
        $ev[0].e | Should -Be 'plan'; $ev[0].job | Should -Be 'exp-1'
        $ev[0].steps.id | Should -Be @('tree', 'list', 'download', 'verify')
        $steps = @($ev | Where-Object e -EQ 'step')
        foreach ($id in 'tree', 'list', 'download', 'verify') { ($steps | Where-Object { $_.id -eq $id -and $_.state -eq 'done' }).Count | Should -BeGreaterOrEqual 1 }
        ($steps | Where-Object { $_.id -eq 'download' -and $_.state -eq 'done' })[-1].done | Should -Be 10
        $res = $ev[-1]
        $res.e | Should -Be 'result'; $res.exit | Should -Be 0; $res.complete | Should -BeTrue; $res.counts.done | Should -Be 10; $res.counts.failed | Should -Be 0
        (Get-Content (Join-Path (Split-Path $evx) 'status.json') -Raw | ConvertFrom-Json).state | Should -Be 'complete'

        $evp = Join-Path $script:Tmp 'job-plan\events.jsonl'
        $p = Invoke-InProcess $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; TargetRoot = 'R'; WhatIf = $true; EventsFile = $evp }
        $pe = Read-Events $evp
        $objs = @($pe | Where-Object { $_.e -eq 'object' -and $_.step -eq 'plan' })
        $objs.Count | Should -BeGreaterOrEqual 2
        ($objs | Where-Object key -EQ 'R/Inbox').total | Should -Be 6
        $pe[-1].exit | Should -Be 0

        $evi = Join-Path $script:Tmp 'job-import\events.jsonl'
        $i = Invoke-InProcess $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; TargetRoot = 'R'; Concurrency = 2; EventsFile = $evi }
        $ie = Read-Events $evi
        $ie[0].steps.id | Should -Be @('check', 'folders', 'send', 'reconcile', 'final')
        $isteps = @($ie | Where-Object e -EQ 'step')
        ($isteps | Where-Object { $_.id -eq 'check' -and $_.state -eq 'done' }).Count | Should -Be 1
        ($isteps | Where-Object { $_.id -eq 'send' -and $_.state -eq 'done' })[-1].done | Should -Be 10
        ($isteps | Where-Object { $_.id -eq 'final' -and $_.state -eq 'done' }).Count | Should -Be 1
        $ie[-1].exit | Should -Be 0; $ie[-1].counts.done | Should -Be 10; $ie[-1].counts.unknown | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
    }
}
