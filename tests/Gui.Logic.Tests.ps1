#Requires -Version 7.4
# Console: the logic that needs no window (run model, ETA, commands, names, job folder). Nothing contacts the cloud.
BeforeAll {
    $script:Package = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    foreach ($f in 'Core', 'Run', 'Jobs') { . (Join-Path $script:Package "module\MailboxTransfer.Gui.$f.ps1") }
    function New-Evt([string]$Kind, [datetime]$At = [datetime]::UtcNow, [hashtable]$Fields = @{}) { $o = [ordered]@{ v = 1; t = $At.ToString('o'); e = $Kind }; foreach ($k in $Fields.Keys) { $o[$k] = $Fields[$k] }; [pscustomobject]$o }
    function Add-Evt($Run, [string]$Kind, [datetime]$At = [datetime]::UtcNow, [hashtable]$Fields = @{}) { $null = Add-MbtGuiEvent -Run $Run -Event (New-Evt $Kind $At $Fields) }
}

Describe 'Run model' {
    It 'knows the steps from the plan event and follows step and object events' {
        $r = New-MbtGuiRun -Task import
        Add-Evt $r plan -Fields @{ steps = @([pscustomobject]@{ id = 'a'; title = 'Check' }, [pscustomobject]@{ id = 'b'; title = 'Send' }) }
        Add-Evt $r step -Fields @{ id = 'a'; state = 'done' }
        Add-Evt $r step -Fields @{ id = 'b'; state = 'running'; done = 5; total = 10; bytes = 50; bytesTotal = 100 }
        Add-Evt $r object -Fields @{ step = 'b'; key = 'Inbox'; state = 'running'; done = 5; total = 8 }
        $r.Steps.Count | Should -Be 2
        $r.Steps[0].State | Should -Be 'done'
        $r.Done | Should -Be 5
        $r.Total | Should -Be 10
        $r.Objects[0].Tone | Should -Be 'run'
    }
    It 'ignores unknown kinds and unknown fields' {
        $r = New-MbtGuiRun
        { Add-Evt $r 'whatever' -Fields @{ x = 1 } } | Should -Not -Throw
        { Add-Evt $r step -Fields @{ id = 'z'; state = 'running'; shiny = $true } } | Should -Not -Throw
        $r.Steps.Count | Should -Be 1
    }
    It 'counts the waits and shows the active one' {
        $r = New-MbtGuiRun
        Add-Evt $r wait -Fields @{ reason = 'throttle'; seconds = 38; count = 7; waited = 252 }
        $r.Waits.Count | Should -Be 7
        $r.Waits.Seconds | Should -Be 252
        $r.Waits.Active.Seconds | Should -Be 38
    }
    It 'a slept event marks the sleep and measures the rate again' {
        $r = New-MbtGuiRun
        Add-Evt $r rate -Fields @{ gibPerHour = 3.9 }
        Add-Evt $r slept -Fields @{ from = [datetime]::UtcNow.AddHours(-3).ToString('o'); to = [datetime]::UtcNow.ToString('o'); seconds = 10800 }
        $r.State | Should -Be 'sleep'
        $r.Rates.Count | Should -Be 0
    }
    It 'a decision needs the person; the result 3 says so' {
        $r = New-MbtGuiRun
        Add-Evt $r decision -Fields @{ key = '7'; msg = 'Two items.'; candidates = @('a', 'b') }
        $r.State | Should -Be 'decision'
        Add-Evt $r result -Fields @{ exit = 3; complete = $false }
        $r.State | Should -Be 'decision-final'
    }
    It 'maps the exit code of the engine to a state' -ForEach @(
        @{ Exit = 0; Stop = $false; State = 'complete' }, @{ Exit = 2; Stop = $false; State = 'partial' }, @{ Exit = 2; Stop = $true; State = 'stopped' }, @{ Exit = 1; Stop = $false; State = 'error' }
    ) {
        $r = New-MbtGuiRun; $r.StopRequested = $Stop
        Add-Evt $r result -Fields @{ exit = $Exit }
        $r.State | Should -Be $State
    }
    It 'an expired sign-in ends resumable' {
        $r = New-MbtGuiRun
        Add-Evt $r error -Fields @{ code = 'signin-required'; msg = 'x' }
        Add-Evt $r result -Fields @{ exit = 2; reason = 'signin-required' }
        $r.State | Should -Be 'partial'
        $r.Result.Reason | Should -Be 'signin-required'
    }
    It 'reads item events of a check and of a repair' {
        $r = New-MbtGuiRun
        Add-Evt $r item -Fields @{ step = 'check'; key = 'MSG-NOT-LISTED'; state = 'error'; msg = 'x'; count = 3 }
        Add-Evt $r item -Fields @{ step = 'write'; key = 'Quarterly'; state = 'partial'; msg = 'recipients'; detail = 'Inbox' }
        Add-Evt $r item -Fields @{ step = 'write'; key = '(no subject)'; state = 'lost' }
        ($r.Items | ForEach-Object Tone) | Should -Be @('bad', 'warn', 'bad')
    }
}

Describe 'Reading the events file' {
    It 'leaves a last line that is still being written' {
        $f = Join-Path $TestDrive 'ev.jsonl'
        Set-Content -LiteralPath $f -Value ('{"e":"info","msg":"a"}' + "`n" + '{"e":"info","msg":"b"') -NoNewline
        $t = Read-MbtGuiEventTail -Path $f -Offset 0
        $t.Lines.Count | Should -Be 1
        Add-Content -LiteralPath $f -Value '}' -NoNewline
        Add-Content -LiteralPath $f -Value "`n" -NoNewline
        $t2 = Read-MbtGuiEventTail -Path $f -Offset $t.Offset
        $t2.Lines.Count | Should -Be 1
        $t2.Lines[0] | Should -Match '"b"'
    }
    It 'returns nothing for a file that is not there and starts again when the file is shorter' {
        (Read-MbtGuiEventTail -Path (Join-Path $TestDrive 'none.jsonl')).Lines.Count | Should -Be 0
        $f = Join-Path $TestDrive 'short.jsonl'
        Set-Content -LiteralPath $f -Value '{"e":"info"}'
        (Read-MbtGuiEventTail -Path $f -Offset 5000).Lines.Count | Should -Be 1
    }
    It 'skips a line that is not JSON' {
        ConvertFrom-MbtGuiEventLine 'not json' | Should -BeNullOrEmpty
    }
}

Describe 'Time left' {
    It 'shows ranges before the run, never one number (import: about 4 GiB per hour)' {
        (Get-MbtGuiPlanEta -Task import -Bytes 47.3GB).Text | Should -Be 'about 12 hours (between 9 and 16 hours)'
        (Get-MbtGuiPlanEta -Task export -Bytes 6.2GB).Text | Should -Match '^between '
        (Get-MbtGuiPlanEta -Task export -Bytes 100MB).Text | Should -Be 'a few minutes'
    }
    It 'keeps the throughput figures in one place' {
        $script:Throughput.ImportGiBPerHour | Should -Be @(3, 5)
        $script:Throughput.ExportGiBPerHour | Should -Be @(30, 150)
    }
    It 'says Measuring for the first minutes and after a sleep, then a rounded range with its basis' {
        $r = New-MbtGuiRun
        $t0 = [datetime]::UtcNow.AddHours(-2)
        Add-Evt $r step $t0 @{ id = 's'; state = 'running'; done = 1; total = 100; bytes = 1GB; bytesTotal = 50GB }
        foreach ($m in 5, 10, 15, 20, 25, 30) { Add-Evt $r step $t0.AddMinutes($m - 1) @{ id = 's'; state = 'running'; done = $m; total = 100; bytes = 1GB; bytesTotal = 50GB }; Add-Evt $r rate $t0.AddMinutes($m) @{ gibPerHour = 3.5 + $m / 30 } }
        (Get-MbtGuiEta -Run $r -Now $t0.AddMinutes(5)).State | Should -Be 'measuring'
        $eta = Get-MbtGuiEta -Run $r -Now $t0.AddMinutes(31)
        $eta.State | Should -Be 'ready'
        $eta.Text | Should -Match '^About .* left'
        $eta.Basis | Should -Match 'Based on the last 30 minutes'
        Add-Evt $r slept $t0.AddMinutes(40) @{ seconds = 600 }
        (Get-MbtGuiEta -Run $r -Now $t0.AddMinutes(41)).State | Should -Be 'measuring'
    }
    It 'drops a window that contains a pause' {
        $r = New-MbtGuiRun
        $t0 = [datetime]::UtcNow.AddHours(-1)
        Add-Evt $r rate $t0 @{ gibPerHour = 3 }
        Add-Evt $r rate $t0.AddMinutes(30) @{ gibPerHour = 50 }
        $r.Rates[1].Gap | Should -BeTrue
    }
}

Describe 'Names and settings' {
    It 'refuses the Windows device names as a PST file name, with any extension, and nothing else' -ForEach @(
        @{ Name = 'aux.pst'; Bad = $true }, @{ Name = 'AUX.pst'; Bad = $true }, @{ Name = 'Aux.backup.pst'; Bad = $true }, @{ Name = 'nul'; Bad = $true }, @{ Name = 'com1.pst'; Bad = $true }, @{ Name = 'LPT9.pst'; Bad = $true },
        @{ Name = 'auxiliary.pst'; Bad = $false }, @{ Name = 'adele.pst'; Bad = $false }, @{ Name = 'com10.pst'; Bad = $false }, @{ Name = 'a|b.pst'; Bad = $true }, @{ Name = 'x.pst.'; Bad = $true }
    ) {
        [bool](Test-MbtGuiPstName $Name) | Should -Be $Bad
    }
    It 'says the name in plain words' {
        Test-MbtGuiPstName 'aux.pst' | Should -Be '"AUX" is a name Windows keeps for itself. Choose another file name.'
    }
    It 'the settings changed for one run go to the command line as parameters, and only when they were changed; no configuration file is written' {
        $s = @{ Task = 'export'; Mailbox = 'adele@contoso.com'; Dest = 'store'; OutFolder = 'D:\Out'; Conc = 4; Tries = 3; Retry = 2; Reconcile = 5; Awake = $false; Changed = @{} }
        $j = @{ Id = 'job-1'; Folder = 'D:\J'; EventsPath = 'D:\J\events.jsonl'; StopPath = 'D:\J\stop.request'; StorePath = 'D:\Store'; SelectionFiles = @('D:\Sel\a.json') }
        (Get-MbtGuiCommand -Task export -S $s -Job $j)[0].Arguments | Should -Not -Contain '-RetryWaitSeconds'
        $s.Changed['Retry'] = $true; $s.Retry = 5; $s.Tries = 4
        $a = (Get-MbtGuiCommand -Task export -S $s -Job $j)[0].Arguments
        $a[($a.IndexOf('-RetryWaitSeconds') + 1)] | Should -Be '5'
        $a[($a.IndexOf('-MaxAttempts') + 1)] | Should -Be '4'
        $a | Should -Not -Contain '-ConfigPath'
        Get-Command Get-MbtGuiEffectiveConfig -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
    }
}

Describe 'Commands of a job' {
    BeforeAll {
        $script:S = @{ Task = 'export'; Mailbox = 'adele@contoso.com'; Shared = ''; ImShared = ''; Dest = 'pst'; OutFolder = 'D:\Out'; OutName = 'adele.pst'; SplitGB = 20; CheckEach = $true; SepAux = $false; Conc = 4; Tries = 3; Reconcile = 5; Awake = $true; ImSrc = 'store'; ImPath = 'D:\Store'; TgtLoc = 'Archive'; TgtMode = 'new'; TgtName = 'Imported x'; Mapping = ''; RpPath = 'D:\a.pst'; RpDest = 'D:\a.repaired.pst'; RpLevel = 'Automatic'; RpKeep = $false; Changed = @{} }
        $script:Job = @{ Id = 'job-1'; Folder = 'D:\J'; EventsPath = 'D:\J\events.jsonl'; StopPath = 'D:\J\stop.request'; ConfigPath = $null; StorePath = 'D:\Store'; SelectionFiles = @('D:\J\selection-1.json'); JobName = 'import-1' }
    }
    It 'builds the export command from controls: selection file, store, PST size, events, stop, keep awake' {
        $a = (Get-MbtGuiCommand -Task export -S $script:S -Job $script:Job)[0].Arguments
        $a | Should -Contain '-SelectionFile'; $a | Should -Contain 'D:\J\selection-1.json'
        $a[($a.IndexOf('-OutputPst') + 1)] | Should -Be 'D:\Out\adele.pst'
        $a[($a.IndexOf('-PstMaxSizeGB') + 1)] | Should -Be '20'
        $a | Should -Contain '-EventsFile'; $a | Should -Contain '-StopFile'; $a | Should -Contain '-KeepAwake'; $a | Should -Contain '-NoPrompt'
        $a | Should -Not -Contain '-WhatIf'
    }
    It 'signing in as "me" adds no certificate argument' {
        $a = (Get-MbtGuiCommand -Task export -S $script:S -Job $script:Job)[0].Arguments
        $a | Should -Not -Contain '-CertificateThumbprint'; $a | Should -Not -Contain '-ClientId'; $a | Should -Not -Contain '-TenantId'
        @(Get-MbtGuiAuthArguments -S $script:S).Count | Should -Be 0
    }
    It 'signing in as the application passes the thumbprint, tenant and application id to the export, the import and the Status calls (run-only: nothing in the settings file)' {
        $s = $script:S.Clone(); $s.AuthMode = 'app'; $s.AppTenant = 'contoso.onmicrosoft.com'; $s.AppClient = '11111111-2222-3333-4444-555555555555'; $s.AppThumb = ('AB' * 20)
        foreach ($task in 'export', 'import') {
            $a = (Get-MbtGuiCommand -Task $task -S $s -Job $script:Job)[0].Arguments
            $a[($a.IndexOf('-CertificateThumbprint') + 1)] | Should -Be ('AB' * 20)
            $a[($a.IndexOf('-TenantId') + 1)] | Should -Be 'contoso.onmicrosoft.com'
            $a[($a.IndexOf('-ClientId') + 1)] | Should -Be '11111111-2222-3333-4444-555555555555'
            $a[($a.IndexOf('-UserPrincipalName') + 1)] | Should -Be 'adele@contoso.com' -Because 'the mailbox of the card is the target mailbox'
        }
    }
    It 'the plan is the same request with -WhatIf' {
        (Get-MbtGuiCommand -Task export -S $script:S -Job $script:Job -Plan)[0].Arguments | Should -Contain '-WhatIf'
    }
    It 'several locations: one command, with one selection file for all of them and the PST written by the same command' {
        $c = @(Get-MbtGuiCommand -Task export -S $script:S -Job $script:Job)
        $c.Count | Should -Be 1
        $c[0].Arguments | Should -Contain '-OutputPst'
        $c[0].Arguments | Should -Not -Contain 'ExportPst'
        @($c[0].Arguments | Where-Object { $_ -eq '-SelectionFile' }).Count | Should -Be 1
    }
    It 'builds the import command with the job name, the new folder and the reconcile delay in seconds' {
        $a = (Get-MbtGuiCommand -Task import -S $script:S -Job $script:Job)[0].Arguments
        $a[($a.IndexOf('-Mode') + 1)] | Should -Be 'Import'
        $a[($a.IndexOf('-TargetRoot') + 1)] | Should -Be 'Imported x'
        $a[($a.IndexOf('-ReconcileDelaySeconds') + 1)] | Should -Be '300'
        $a[($a.IndexOf('-Mailbox') + 1)] | Should -Be 'Archive'
        $a | Should -Contain '-StorePath'
    }
    It 'imports a PST with the merged engine: -Mode ImportPst -PstPath, the journal in the job folder, the new folder, no selection file' {
        $s = $script:S.Clone(); $s.ImSrc = 'pst'; $s.ImPath = 'D:\PST\a.pst'
        $j = $script:Job.Clone(); $j.Folder = 'D:\J'
        $a = (Get-MbtGuiCommand -Task import -S $s -Job $j)[0].Arguments
        $a[($a.IndexOf('-Mode') + 1)] | Should -Be 'ImportPst'
        $a[($a.IndexOf('-PstPath') + 1)] | Should -Be 'D:\PST\a.pst'
        $a[($a.IndexOf('-TempPath') + 1)] | Should -Be 'D:\J\work'
        $a[($a.IndexOf('-TargetRoot') + 1)] | Should -Be 'Imported x'
        $a | Should -Not -Contain '-PstSource'; $a | Should -Not -Contain '-SelectionFile'; $a | Should -Not -Contain '-StorePath'; $a | Should -Contain '-NoPrompt'
        # the same folder names: the standard folders of the PST are merged into the real ones
        $s.TgtMode = 'same'
        (Get-MbtGuiCommand -Task import -S $s -Job $j)[0].Arguments | Should -Contain '-MergeWellKnown'
    }
    It 'imports a folder of EML files: -Mode ImportEml -EmlPath -TargetRoot' {
        $s = $script:S.Clone(); $s.ImSrc = 'eml'; $s.ImPath = 'D:\eml'; $s.Mapping = 'D:\map.psd1'
        $a = (Get-MbtGuiCommand -Task import -S $s -Job $script:Job)[0].Arguments
        $a[($a.IndexOf('-Mode') + 1)] | Should -Be 'ImportEml'
        $a[($a.IndexOf('-EmlPath') + 1)] | Should -Be 'D:\eml'
        $a[($a.IndexOf('-TargetRoot') + 1)] | Should -Be 'Imported x'
        $a | Should -Not -Contain '-MappingFile'; $a | Should -Not -Contain '-SelectionFile'
    }
    It 'an import into a shared mailbox gives its address as -MailboxId, for every source' -ForEach @('store', 'pst', 'eml') {
        $s = $script:S.Clone(); $s.ImSrc = $_; $s.ImPath = 'D:\x'; $s.ImShared = 'projects@contoso.com'; $s.TgtLoc = 'Primary'
        $j = $script:Job.Clone(); $j.Folder = 'D:\J'
        $a = (Get-MbtGuiCommand -Task import -S $s -Job $j)[0].Arguments
        $a[($a.IndexOf('-MailboxId') + 1)] | Should -Be 'projects@contoso.com'
        $a[($a.IndexOf('-UserPrincipalName') + 1)] | Should -Be 'adele@contoso.com'
        $a[($a.IndexOf('-Mailbox') + 1)] | Should -Be 'Primary'
    }
    It 'an import into the own mailbox carries no -MailboxId' {
        (Get-MbtGuiCommand -Task import -S $script:S -Job $script:Job)[0].Arguments | Should -Not -Contain '-MailboxId'
    }
    It 'EML files: the items come to the store first, then -Mode ExportEml writes them from the store, with events and stop' {
        $s = $script:S.Clone(); $s.Dest = 'eml'
        $c = @(Get-MbtGuiCommand -Task export -S $s -Job $script:Job)
        $c.Count | Should -Be 2
        $c[0].Arguments | Should -Not -Contain '-OutputPst'
        $e = $c[1].Arguments
        $e[($e.IndexOf('-Mode') + 1)] | Should -Be 'ExportEml'
        $e[($e.IndexOf('-StorePath') + 1)] | Should -Be 'D:\Store'
        $e[($e.IndexOf('-EmlPath') + 1)] | Should -Be 'D:\Out'
        $e | Should -Contain '-EventsFile'; $e | Should -Contain '-StopFile'; $e | Should -Contain '-KeepAwake'
    }
    It 'builds repair and check commands; the check never carries a destination' {
        $r = (Get-MbtGuiCommand -Task repair -S $script:S -Job $script:Job)[0].Arguments
        $r[($r.IndexOf('-Level') + 1)] | Should -Be 'Automatic'
        $r | Should -Not -Contain '-KeepDamagedAttachments'
        $c = (Get-MbtGuiCommand -Task check -S $script:S -Job $script:Job)[0].Arguments
        $c | Should -Not -Contain '-Destination'
        $c[($c.IndexOf('-Mode') + 1)] | Should -Be 'CheckPst'
    }
}

Describe 'Job folder and the pump' {
    It 'creates the private job folder under jobs\' {
        $d = New-MbtGuiJobFolder -WorkFolder $TestDrive -Task export
        $d | Should -Match 'jobs\\\d{8}-\d{4}-export-'
        Test-Path -LiteralPath $d | Should -BeTrue
    }
    It 'a stop request is a file in the job folder' {
        $j = New-MbtGuiJob -WorkFolder $TestDrive -Task import -EntryScript 'x.ps1'
        Request-MbtGuiStop -Job $j
        Test-Path -LiteralPath $j.StopPath | Should -BeTrue
        (Get-MbtGuiStopWait $j) | Should -BeLessThan 5
    }
    It 'only the last result of a chain of commands is the result of the job' {
        $j = New-MbtGuiJob -WorkFolder $TestDrive -Task export -EntryScript 'x.ps1'
        $j.Queue.Enqueue([pscustomobject]@{ Arguments = @() })
        $run = New-MbtGuiRun
        Set-Content -LiteralPath $j.EventsPath -Value ('{"v":1,"e":"result","exit":0}')
        $null = Step-MbtGuiJob -Job $j -Run $run
        $run.Result | Should -BeNullOrEmpty
    }
    It 'a chain keeps the worst exit code: a first command that ended resumable (2) makes the job resumable although the last one ended 0' {
        $j = New-MbtGuiJob -WorkFolder $TestDrive -Task export -EntryScript 'x.ps1'
        $j.Queue.Enqueue([pscustomobject]@{ Arguments = @() })
        $run = New-MbtGuiRun
        Set-Content -LiteralPath $j.EventsPath -Value ('{"v":1,"e":"result","exit":2,"counts":{"done":3}}')
        $null = Step-MbtGuiJob -Job $j -Run $run
        $run.Result | Should -BeNullOrEmpty
        $j.Queue.Clear()
        Add-Content -LiteralPath $j.EventsPath -Value ('{"v":1,"e":"result","exit":0,"complete":true,"counts":{"done":9}}')
        $null = Step-MbtGuiJob -Job $j -Run $run
        $run.Result.Exit | Should -Be 2
        $run.Result.Complete | Should -BeFalse
        $run.Result.Next | Should -Be 'continue'
        $run.State | Should -Be 'partial'
        # a chain whose first command is complete keeps the exit of the last one
        $j2 = New-MbtGuiJob -WorkFolder $TestDrive -Task export -EntryScript 'x.ps1'
        $j2.Queue.Enqueue([pscustomobject]@{ Arguments = @() })
        $run2 = New-MbtGuiRun
        Set-Content -LiteralPath $j2.EventsPath -Value ('{"v":1,"e":"result","exit":0}')
        $null = Step-MbtGuiJob -Job $j2 -Run $run2
        $j2.Queue.Clear()
        Add-Content -LiteralPath $j2.EventsPath -Value ('{"v":1,"e":"result","exit":0,"complete":true}')
        $null = Step-MbtGuiJob -Job $j2 -Run $run2
        $run2.Result.Exit | Should -Be 0
    }
    It 'a status of a failed child is an error result, a stopped one is resumable' {
        $j = New-MbtGuiJob -WorkFolder $TestDrive -Task export -EntryScript 'x.ps1'
        $j.Process = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList '-NoProfile', '-Command', 'exit 9' -PassThru -WindowStyle Hidden
        $null = $j.Process.WaitForExit(30000)
        $run = New-MbtGuiRun
        $null = Step-MbtGuiJob -Job $j -Run $run
        $run.State | Should -Be 'error'
    }
}

Describe 'Capabilities of the engine adapter' {
    It 'the real adapter takes its capabilities from the merged engine, and every argument of every command of the console is a parameter of the entry script' {
        Import-Module (Join-Path $script:Package 'MailboxTransfer.psd1') -Force
        $c = Get-MbtCapability -EntryScript (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1')
        foreach ($n in 'PstImport', 'EmlExport', 'EmlImport', 'SharedMailbox', 'PstExport', 'PstCheck', 'PstRepair', 'Events', 'FolderChoice', 'PstCopy', 'PstPreview') { $c.$n | Should -BeTrue -Because $n }
        $entry = Get-Command (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1')
        foreach ($p in 'PstPath', 'EmlPath', 'MailboxId', 'TargetRoot', 'EventsFile', 'StopFile', 'KeepAwake', 'NoPrompt', 'TempPath') { $entry.Parameters.ContainsKey($p) | Should -BeTrue -Because $p }
        $s = @{ Task = 'export'; OpPath = 'D:\open.pst'; Mailbox = 'adele@contoso.com'; Dest = 'pst'; OutFolder = 'D:\Out'; OutName = 'a.pst'; SplitGB = 20; CheckEach = $true; SepAux = $true; Conc = 4; Tries = 3; Reconcile = 5; Awake = $true; ImSrc = 'pst'; ImPath = 'D:\a.pst'; ImShared = 'p@contoso.com'; TgtLoc = 'Primary'; TgtMode = 'new'; TgtName = 'X'; Mapping = 'D:\m.psd1'; RpPath = 'D:\a.pst'; RpDest = 'D:\b.pst'; RpLevel = 'Automatic'; RpKeep = $true; Changed = @{} }
        $j = @{ Id = 'job-1'; Folder = 'D:\J'; EventsPath = 'D:\J\e.jsonl'; StopPath = 'D:\J\s'; ConfigPath = 'D:\J\c.psd1'; StorePath = 'D:\Store'; SelectionFiles = @('D:\J\sel.json'); JobName = 'j'; FolderFile = 'D:\J\folders.txt'; ItemsFile = 'D:\J\items.txt' }
        $modes = @($entry.Parameters['Mode'].Attributes.Where({ $_ -is [Management.Automation.ValidateSetAttribute] }).ValidValues)
        foreach ($case in @(@('export', 'pst'), @('export', 'eml'), @('export', 'store'), @('import', 'pst'), @('import', 'eml'), @('import', 'store'), @('repair', ''), @('check', ''), @('copy', 'pst'))) {
            $t = $s.Clone(); $t.Dest = $case[1]; if ($case[0] -eq 'import') { $t.ImSrc = $case[1] }
            foreach ($cmd in @(Get-MbtGuiCommand -Task $case[0] -S $t -Job $j)) {
                foreach ($a in $cmd.Arguments) {
                    if ($a -like '-*') { $entry.Parameters.ContainsKey($a.TrimStart('-')) | Should -BeTrue -Because "$($case -join '/'): $a" }
                }
                $mode = $cmd.Arguments[($cmd.Arguments.IndexOf('-Mode') + 1)]
                $modes | Should -Contain $mode
            }
        }
    }
    It 'the fake engine says what it can do' {
        . (Join-Path $script:Package 'module\MailboxTransfer.Gui.Fake.ps1')
        $e = New-MbtGuiFakeEngine -Capability @{ PstImport = $true }
        $e.Capability.PstImport | Should -BeTrue
        $e.Capability.EmlExport | Should -BeFalse
        $e.Name | Should -Be 'Fake'
    }
}

Describe 'Folders of a PST or of .eml files, numbers of the plan, items refused for good' {
    BeforeAll {
        function New-FakeSel([object[]]$Entries) { [pscustomobject]@{ Entries = $Entries } }
        function New-FakeNode([hashtable]$Tag, [hashtable]$Picks = @{}) { [pscustomobject]@{ Tag = $Tag; PickRows = $Picks } }
    }
    It 'the ticks of a PST become paths below its top folder: a folder and what is below it, a folder alone, messages by node id' {
        $sel = New-FakeSel @(
            @{ Kind = 'tree'; Node = (New-FakeNode @{ FolderPath = 'Inbox/Projects' }) },
            @{ Kind = 'own'; Node = (New-FakeNode @{ FolderPath = 'Sent Items' }) },
            @{ Kind = 'picks'; Node = (New-FakeNode @{ FolderPath = 'Calendar' } @{ a = @{ Nid = 2097252 }; b = @{ Nid = 2097284 } }) })
        $c = Get-MbtGuiFolderChoice -Selection $sel -Source pst
        @($c.Include) | Should -Be @('Inbox/Projects', 'Sent Items')
        @($c.Exclude) | Should -Be @('Sent Items/*')
        @($c.Nids | Sort-Object) | Should -Be @(2097252, 2097284)
        $eml = Get-MbtGuiFolderChoice -Selection (New-FakeSel @(@{ Kind = 'tree'; Node = (New-FakeNode @{ Rel = '2019/Q1' }) })) -Source eml
        @($eml.Include) | Should -Be @('2019/Q1')
        (Get-MbtGuiFolderChoice -Selection (New-FakeSel @()) -Source pst).Include.Count | Should -Be 0
    }
    It 'a star, a question mark or a backtick in a folder name is written literally' {
        ConvertTo-MbtGuiFolderPattern 'Q1*2020/What?' | Should -Be 'Q1`*2020/What`?'
        ConvertTo-MbtGuiFolderPattern 'a`b' | Should -Be 'a``b'
        ConvertTo-MbtGuiFolderPattern 'Inbox/Projets été' | Should -Be 'Inbox/Projets été'
    }
    It 'writes folders.txt and items.txt in the job folder, and the import command passes them (only for a PST or .eml source)' {
        $j = New-MbtGuiJob -WorkFolder $TestDrive -Task import -EntryScript 'x.ps1'
        Write-MbtGuiFolderFiles -Job $j -Choice ([pscustomobject]@{ Include = @('Inbox', 'Old, mail'); Exclude = @('Inbox/*'); Nids = @([uint32]7) })
        Get-Content -LiteralPath $j.FolderFile -Encoding utf8 | Should -Be @('+Inbox', '+Old, mail', '-Inbox/*')
        Get-Content -LiteralPath $j.ItemsFile | Should -Be @('7')
        $s = @{ Task = 'import'; Mailbox = 'a@contoso.com'; ImSrc = 'pst'; ImPath = 'D:\a.pst'; TgtLoc = 'Primary'; TgtMode = 'new'; TgtName = 'X'; Conc = 4; Tries = 3; Reconcile = 5; Awake = $true; Changed = @{}; Mapping = ''; ImShared = '' }
        $a = (Get-MbtGuiCommand -Task import -S $s -Job $j)[0].Arguments
        $a[($a.IndexOf('-FolderFile') + 1)] | Should -Be $j.FolderFile
        $s.ImSrc = 'store'
        (Get-MbtGuiCommand -Task import -S $s -Job $j)[0].Arguments | Should -Not -Contain '-FolderFile'
        Write-MbtGuiFolderFiles -Job $j -Choice ([pscustomobject]@{ Include = @(); Exclude = @(); Nids = @() })
        $j.FolderFile | Should -BeNullOrEmpty
    }
    It 'the copy command reads the open PST and writes new PST files, with the folders and messages ticked' {
        $j = New-MbtGuiJob -WorkFolder $TestDrive -Task copy -EntryScript 'x.ps1'
        Write-MbtGuiFolderFiles -Job $j -Choice ([pscustomobject]@{ Include = @('Inbox'); Exclude = @(); Nids = @([uint32]5) })
        $s = @{ Task = 'copy'; OpPath = 'D:\open.pst'; OutFolder = 'D:\Out'; OutName = 'copy.pst'; SplitGB = 20; CheckEach = $false; Conc = 4; Tries = 3; Awake = $true; Changed = @{} }
        $a = (Get-MbtGuiCommand -Task copy -S $s -Job $j)[0].Arguments
        $a[($a.IndexOf('-Mode') + 1)] | Should -Be 'ExportPst'
        $a[($a.IndexOf('-PstPath') + 1)] | Should -Be 'D:\open.pst'
        $a[($a.IndexOf('-OutputPst') + 1)] | Should -Be 'D:\Out\copy.pst'
        $a[($a.IndexOf('-PstMaxSizeGB') + 1)] | Should -Be '20'
        $a | Should -Contain '-SkipCheck'
        $a[($a.IndexOf('-IncludeItemFile') + 1)] | Should -Be $j.ItemsFile
        (Get-MbtGuiCommand -Task copy -S $s -Job $j -Plan)[0].Arguments | Should -Contain '-WhatIf'
    }
    It 'the numbers of the plan come from the plan objects of the engine; a size that is not given is never counted as 0 B' {
        $r = New-MbtGuiRun -Task export
        Add-Evt $r object -Fields @{ step = 'plan'; key = 'Primary/Inbox'; state = 'planned'; total = 1284; bytes = [long](2.1GB) }
        Add-Evt $r object -Fields @{ step = 'plan'; key = 'Archive/Finance'; state = 'planned'; total = 860; bytes = [long](3.9GB) }
        $t = Get-MbtGuiPlanTotals $r
        $t.Items | Should -Be 2144
        $t.UnknownSize | Should -Be 0
        Format-MbtGuiPlanTotals $t | Should -Be '2,144 items in 2 folders, 6 GiB'
        Add-Evt $r object -Fields @{ step = 'plan'; key = 'AuxArchive/Old Mail'; state = 'planned'; total = 278; bytes = -1 }
        Add-Evt $r object -Fields @{ step = 'plan'; key = '(folders)'; state = 'planned'; total = 12 }
        $t2 = Get-MbtGuiPlanTotals $r
        $t2.Items | Should -Be 2422
        $t2.UnknownSize | Should -Be 1
        $t2.FoldersToCreate | Should -Be 12
        Format-MbtGuiPlanTotals $t2 | Should -BeLike '2,422 items in 3 folders, 6 GiB + 1 folder of unknown size'
        $only = New-MbtGuiRun
        Add-Evt $only object -Fields @{ step = 'plan'; key = 'AuxArchive/Old Mail'; state = 'planned'; total = 278; bytes = -1 }
        Format-MbtGuiPlanTotals (Get-MbtGuiPlanTotals $only) | Should -Be '278 items in 1 folder, size not known'
        Get-MbtGuiPlanTotals (New-MbtGuiRun) | Should -BeNullOrEmpty
    }
    It 'a selection with a folder of unknown size says so: "size unknown" or "+ n folders of unknown size", and "at least" while subfolders are not read' {
        Format-MbtGuiSelectionSize ([pscustomobject]@{ Bytes = 0; UnknownSize = 1 }) | Should -Be 'size unknown'
        Format-MbtGuiSelectionSize ([pscustomobject]@{ Bytes = [long](1GB); UnknownSize = 2 }) | Should -Be '1 GiB + 2 folders of unknown size'
        Format-MbtGuiSelectionSize ([pscustomobject]@{ Bytes = [long](1GB); UnknownSize = 0 }) | Should -Be '1 GiB'
        Format-MbtGuiSelectionItems ([pscustomobject]@{ Items = 1500; SubtreeUnread = $true }) | Should -Be 'at least 1,500'
        Format-MbtGuiSelectionItems ([pscustomobject]@{ Items = 1500; SubtreeUnread = $false }) | Should -Be '1,500'
    }
    It 'exit 1 with the reason refused-items is a finished job (not an error); without it, it is an error' {
        $r = New-MbtGuiRun -Task import
        Add-Evt $r result -Fields @{ exit = 1; complete = $true; reason = 'refused-items'; next = 'done'; counts = [pscustomobject]@{ done = 9; refused = 3 } }
        $r.State | Should -Be 'refused'
        $r.Result.Complete | Should -BeTrue
        $e = New-MbtGuiRun -Task import
        Add-Evt $e result -Fields @{ exit = 1; complete = $false; next = 'retry' }
        $e.State | Should -Be 'error'
    }
    It 'says in plain words why an item was refused' {
        ConvertTo-MbtGuiRefusalReason 'refused by the service: ErrorItemTooLarge x' | Should -BeLike '*too large*mailbox is full*'
        ConvertTo-MbtGuiRefusalReason 'cannot be read from the PST: block 4 damaged' | Should -BeLike '*cannot be read from the PST file*'
        ConvertTo-MbtGuiRefusalReason 'gave up after 3 attempt(s), nothing was created. Last error: HTTP 404' | Should -BeLike '*nothing was created*'
        ConvertTo-MbtGuiRefusalReason 'something else' | Should -Be 'something else'
    }
}

Describe 'Singular and plural in the words of the console' {
    It 'a count with its noun: 0 and 2 are plural, 1 is singular, the thousands are grouped, an irregular plural is given' {
        Format-MbtGuiQuantity 0 'location' | Should -Be '0 locations'
        Format-MbtGuiQuantity 1 'location' | Should -Be '1 location'
        Format-MbtGuiQuantity 2 'location' | Should -Be '2 locations'
        Format-MbtGuiQuantity 1234 'item' | Should -Be '1,234 items'
        Format-MbtGuiQuantity 1 'entry' 'entries' | Should -Be '1 entry'
        Format-MbtGuiQuantity 2 'entry' 'entries' | Should -Be '2 entries'
        Format-MbtGuiCount 1234 | Should -Be '1,234'
    }
    It 'the word alone, for a noun or a verb' {
        Get-MbtGuiPlural 1 'item' | Should -Be 'item'
        Get-MbtGuiPlural 0 'item' | Should -Be 'items'
        Get-MbtGuiPlural 1 'is' 'are' | Should -Be 'is'
        Get-MbtGuiPlural 3 'is' 'are' | Should -Be 'are'
    }
    It 'the totals of a plan and the size of a selection agree with the count' {
        Format-MbtGuiPlanTotals ([pscustomobject]@{ Items = 1; Folders = 1; Bytes = [long](2GB); UnknownSize = 0 }) | Should -Be '1 item in 1 folder, 2 GiB'
        Format-MbtGuiPlanTotals ([pscustomobject]@{ Items = 2; Folders = 2; Bytes = [long](2GB); UnknownSize = 1 }) | Should -Be '2 items in 2 folders, 2 GiB + 1 folder of unknown size'
        Format-MbtGuiSelectionSize ([pscustomobject]@{ Bytes = [long](1GB); UnknownSize = 1 }) | Should -Be '1 GiB + 1 folder of unknown size'
        Format-MbtGuiSelectionItems ([pscustomobject]@{ Items = 1; SubtreeUnread = $false }) -WithNoun | Should -Be '1 item'
        Format-MbtGuiSelectionItems ([pscustomobject]@{ Items = 2; SubtreeUnread = $true }) -WithNoun | Should -Be 'at least 2 items'
        Format-MbtGuiSelectionItems ([pscustomobject]@{ Items = 1500; SubtreeUnread = $false }) -WithNoun | Should -Be '1,500 items'
    }
    It 'the title of the refused items: the item, all of them, or the first of many' {
        Format-MbtGuiRefusedSummary 1 1 | Should -Be 'The refused item'
        Format-MbtGuiRefusedSummary 3 3 | Should -Be 'The 3 refused items'
        Format-MbtGuiRefusedSummary 25 1234 | Should -Be 'The first 25 of 1,234 refused items'
        Format-MbtGuiRefusedSummary 1 5 | Should -Be 'The first of 5 refused items'
    }
    It 'the title of the messages a repair could not recover whole' {
        Format-MbtGuiIncompleteSummary 1 | Should -Be 'The message that is not complete'
        Format-MbtGuiIncompleteSummary 3 | Should -Be 'The 3 messages that are not complete'
    }
    It 'the sentences with a verb agree with the count' {
        Format-MbtGuiDecisionTitle 1 | Should -Be '1 item needs your decision'
        Format-MbtGuiDecisionTitle 2 | Should -Be '2 items need your decision'
        Format-MbtGuiStillToDo 1 | Should -Be '1 item is still to do.'
        Format-MbtGuiStillToDo 2000 | Should -Be '2,000 items are still to do.'
        Format-MbtGuiProgressOf 0 1 | Should -Be '0 of 1 item'
        Format-MbtGuiProgressOf 5 1208 | Should -Be '5 of 1,208 items'
    }
}

Describe 'Report' {
    It 'writes CSV and JSON of a finished run, and next to the new file for a repair' {
        $r = New-MbtGuiRun -Task repair -Title 'Repair x'
        Add-Evt $r plan -Fields @{ steps = @([pscustomobject]@{ id = 'w'; title = 'Write the new file' }) }
        Add-Evt $r item -Fields @{ step = 'w'; key = 'Quarterly'; state = 'partial'; msg = 'recipients'; detail = 'Inbox' }
        Add-Evt $r result -Fields @{ exit = 0; counts = [pscustomobject]@{ recovered = 5 } }
        $dest = Join-Path $TestDrive 'a.repaired.pst'
        $files = Write-MbtGuiRunReport -Run $r -Folder (Join-Path $TestDrive 'job') -Task repair -RepairDestination $dest
        $files.Count | Should -Be 4
        (Join-Path $TestDrive 'a.repaired.repair-report.csv') | Should -Exist
        $j = Get-Content -LiteralPath (Join-Path $TestDrive 'job\report.json') -Raw | ConvertFrom-Json
        $j.exitCode | Should -Be 0
        @($j.rows | Where-Object Kind -EQ 'item').Count | Should -Be 1
        (Import-Csv -LiteralPath (Join-Path $TestDrive 'job\report.csv')).Count | Should -Be 2
    }
}

Describe 'The report of a run is written in the background' {
    It 'report.json and report.csv are written at once; report.html only when it is not left to the worker' {
        $r = New-MbtGuiRun -Task export -Title 'Export x'
        Add-Evt $r plan -Fields @{ steps = @([pscustomobject]@{ id = 'd'; title = 'Download the items' }) }
        Add-Evt $r result -Fields @{ exit = 0; counts = [pscustomobject]@{ done = 5 } }
        $folder = Join-Path $TestDrive 'bg-job'
        $null = New-Item -ItemType Directory -Force -Path $folder
        Set-Content -LiteralPath (Join-Path $folder 'events.jsonl') -Value '{"v":1,"t":"2026-11-12T10:00:00.000Z","e":"plan","steps":[{"id":"d","title":"Download the items"}]}'
        $files = @(Write-MbtGuiRunReport -Run $r -Folder $folder -Task export -NoHtml)
        $files.Count | Should -Be 2
        (Join-Path $folder 'report.json') | Should -Exist
        (Join-Path $folder 'report.csv') | Should -Exist
        (Join-Path $folder 'report.html') | Should -Not -Exist
        @(Write-MbtGuiRunReport -Run $r -Folder $folder -Task export).Count | Should -Be 3
        (Join-Path $folder 'report.html') | Should -Exist
    }
    It 'the end of a run leaves report.html to the worker, says so on the Done page, and writes the overview when the report is there' {
        $run = Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Run.ps1') -Raw
        $run | Should -Match '(?s)function Complete-MbtGuiRun.*?Write-MbtGuiRunReport [^\r\n]*-NoHtml.*?\$g\.ReportState = ''writing''.*?\$reportDone = \{.*?\$g\.Engine\.JobOverview \$work.*?& \$g\.Engine\.JobReport \$g\.Job\.Folder \$reportDone'
        $run | Should -Match '(?s)\$openReport = \{.*?ReportState -eq ''writing''.*?The report is being written.*?Get-MbtGuiReportPath'
        (Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Engine.ps1') -Raw) | Should -Match '(?s)\$e\.JobReport = .*?-OnDone \{ param\(\$res\) & \$OnDone @\(\$res\)\[0\] \}.*?Write-MbtJobHtmlReport -Folder \$folder'
        (Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Fake.ps1') -Raw) | Should -Match '\$engine\.JobReport = \{ param\(\$Folder, \$OnDone, \$OnError\) & \$OnDone \$null \}'
    }
}

Describe 'Results of the engine handed to the pages' {
    It 'the worker hands an array: an engine call whose page expects one object unwraps it (the sign-in gives one object, never an array)' {
        $file = Join-Path $script:Package 'module\MailboxTransfer.Gui.Engine.ps1'
        $tokens = $null; $errors = $null
        $ast = [Management.Automation.Language.Parser]::ParseFile($file, [ref]$tokens, [ref]$errors)
        $raw = foreach ($call in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'Start-MbtGuiWork' }, $true)) {
            $i = [array]::FindIndex([object[]]$call.CommandElements, [Predicate[object]] { param($x) $x -is [Management.Automation.Language.CommandParameterAst] -and $x.ParameterName -eq 'OnDone' })
            if ($i -ge 0 -and $call.CommandElements[$i + 1].Extent.Text -eq '$OnDone') {
                $assign = $call.Parent; while ($assign -and $assign -isnot [Management.Automation.Language.AssignmentStatementAst]) { $assign = $assign.Parent }
                $assign.Left.Extent.Text
            }
        }
        @($raw | Sort-Object) | Should -Be @('$e.SaveSelection', '$e.SaveSelections') -Because 'these two hand a path that the page reads as an array; every other call gives @($res)[0]'
        (Get-Content -LiteralPath $file -Raw) | Should -Match '(?s)\$e\.SignIn = .*?-OnDone \{ param\(\$res\) & \$OnDone @\(\$res\)\[0\] \}'
    }
    It 'a page step that fails after an engine call is written down, not raised' {
        $record = try { throw 'boom' } catch { $_ }
        { Write-MbtGuiFault 'OnDone' $record } | Should -Not -Throw
    }
}

Describe 'Rows built in a loop' {
    BeforeAll {
        # The rows of a table are written "foreach (...) { , @(cell, cell, cell) }". A loop that runs once gives its one row as the three cells, not as a list of one row, unless the
        # assignment wraps it: $rows = @(foreach ...). This finds every assignment of the result of a loop (or of ForEach-Object, or of a script block) that emits a row with the comma
        # and does not wrap it.
        function Find-UnwrappedRow {
            param([string]$Path, [string]$Code)
            $tokens = $null; $errors = $null
            $ast = if ($Code) { [Management.Automation.Language.Parser]::ParseInput($Code, [ref]$tokens, [ref]$errors) } else { [Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors) }
            $loops = @('ForEachStatementAst', 'ForStatementAst', 'WhileStatementAst', 'DoWhileStatementAst', 'DoUntilStatementAst')
            foreach ($assign in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and $n.Operator -eq 'Equals' }, $true)) {
                $rhs = $assign.Right
                foreach ($comma in $rhs.FindAll({ param($n) $n -is [Management.Automation.Language.ArrayLiteralAst] -and $n.Elements.Count -eq 1 -and $n.Extent.Text.StartsWith(',') }, $true)) {
                    # emitted: the comma is the expression of a statement (alone, or in parentheses)
                    $at = $comma.Parent
                    if ($at -is [Management.Automation.Language.ParenExpressionAst]) { $at = $at.Parent }
                    if ($at -isnot [Management.Automation.Language.CommandExpressionAst]) { continue }
                    # and the statement is in a loop, a ForEach-Object block or a script block that the right side runs, with no @( ) around it
                    $inside = $false; $wrapped = $false; $up = $comma.Parent
                    while ($up -and -not [object]::ReferenceEquals($up, $rhs.Parent)) {
                        if ($up.GetType().Name -in $loops -or $up -is [Management.Automation.Language.ScriptBlockExpressionAst]) { $inside = $true }
                        if ($up -is [Management.Automation.Language.ArrayExpressionAst]) { $wrapped = $true }
                        $up = $up.Parent
                    }
                    if ($inside -and -not $wrapped) { "$(if ($Path) { Split-Path $Path -Leaf }):$($assign.Extent.StartLineNumber): $($assign.Left.Extent.Text)"; break }
                }
            }
        }
    }
    It 'finds the assignments whose single row would be read as its cells, and leaves the wrapped ones' {
        $code = @'
$a = foreach ($x in 1) { , @($x, 2) }
$b = @(foreach ($x in 1) { , @($x, 2) })
$c = $list | ForEach-Object { , @($_, 2) }
$d = for ($i = 0; $i -lt 1; $i++) { , @($i, 2) }
$e = foreach ($x in 1) { if ($x) { , @(1, 2) } }
$f = & { foreach ($x in 1) { (, @(1, 2)) } }
$g = , @(1, 2)
$h = foreach ($x in 1) { @(1, 2) }
$i = @($list | ForEach-Object { , @($_, 2) })
'@
        $found = @(Find-UnwrappedRow -Code $code)
        @($found | ForEach-Object { ($_ -split ': ')[1] }) | Should -Be @('$a', '$c', '$d', '$e', '$f')
    }
    It 'no table of the console is built from a loop that gives a bare row: every list of rows is $rows = @(...)' {
        $found = foreach ($file in Get-ChildItem -LiteralPath (Join-Path $script:Package 'module') -Filter '*.ps1') { Find-UnwrappedRow -Path $file.FullName }
        @($found) | Should -BeNullOrEmpty -Because 'a loop that runs once would hand New-GuiTable its one row as three cells'
    }
}

Describe 'What a job that has not counted an item yet shows' {
    It 'the plan check says what the engine does and for how long' {
        $r = New-MbtGuiRun -Task export
        Add-Evt $r plan -Fields @{ steps = @([pscustomobject]@{ id = 'tree'; title = 'Read the folder tree' }, [pscustomobject]@{ id = 'list'; title = 'List the items' }) }
        $t0 = [datetime]::UtcNow.AddSeconds(-34)
        Get-MbtGuiPlanProgressText -Run $r -Started $t0 | Should -Be 'nothing is changed, 34 s' -Because 'before the first word of the engine'
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'running'; msg = 'Signing in to Microsoft' }
        Get-MbtGuiPlanProgressText -Run $r -Started $t0 -Now $t0.AddSeconds(3) | Should -Be 'Signing in to Microsoft, 3 s'
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'running'; msg = '1,240 folders read' }
        Get-MbtGuiPlanProgressText -Run $r -Started $t0 -Now $t0.AddSeconds(95) | Should -Be '1,240 folders read, 1 min 35 s'
        Get-MbtGuiPlanProgressText -Run $null -Started $null | Should -Be 'nothing is changed'
    }
    It 'says nothing about a step that is not running' {
        $r = New-MbtGuiRun -Task export
        Add-Evt $r plan -Fields @{ steps = @([pscustomobject]@{ id = 'tree'; title = 'Read the folder tree' }) }
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'done'; msg = '4,443 folders read' }
        Get-MbtGuiRunningNote $r | Should -Be ''
    }
    It 'the head of the Run page says Preparing, with what is read, until the first item is counted' {
        $r = New-MbtGuiRun -Task export
        Get-MbtGuiPreparingText $r | Should -Be '**Preparing.** Connecting to Microsoft and reading what you chose.' -Because 'no event yet'
        Add-Evt $r plan -Fields @{ steps = @([pscustomobject]@{ id = 'tree'; title = 'Read the folder tree' }, [pscustomobject]@{ id = 'list'; title = 'List the items you chose' }) }
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'running' }
        Get-MbtGuiPreparingText $r | Should -Be '**Preparing.** Read the folder tree.'
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'running'; msg = '1,240 folders read' }
        Get-MbtGuiPreparingText $r | Should -Be '**Preparing.** 1,240 folders read.'
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'done'; msg = '4,443 folders read' }
        Add-Evt $r step -Fields @{ id = 'list'; state = 'running'; total = 15; done = 0 }
        Get-MbtGuiPreparingText $r | Should -Be '' -Because 'the items are counted'
    }
    It 'no Preparing hint once the job ended or was stopped' {
        $r = New-MbtGuiRun -Task export
        Add-Evt $r result -Fields @{ exit = 1 }
        Get-MbtGuiPreparingText $r | Should -Be ''
        $s = New-MbtGuiRun -Task export
        Add-Evt $s slept -Fields @{ from = [datetime]::UtcNow.AddHours(-1).ToString('o'); to = [datetime]::UtcNow.ToString('o'); seconds = 3600 }
        Get-MbtGuiPreparingText $s | Should -Be ''
    }
    It 'a step that runs is drawn with the play icon and one that waits with the clock' {
        . (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Run.ps1')
        (Get-MbtGuiStepTone @{ State = 'running' } 'running')[1] | Should -Be 'play'
        (Get-MbtGuiStepTone @{ State = 'waiting' } 'running')[1] | Should -Be 'clock'
        (Get-MbtGuiStepTone @{ State = 'done' } 'running')[1] | Should -Be 'check'
    }
}

Describe 'The text beside Start on a Verify page' {
    BeforeAll {
        . (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Run.ps1')
        function Get-MbtGuiScreen { return $script:Screen }
    }
    BeforeEach { $script:Gui = @{ S = @{ PlanState = 'idle'; Confirmed = $false } } }
    It 'says why Start waits while the plan is being checked, and before the check begins' -ForEach @('exPlan', 'imPlan', 'cpPlan') {
        $script:Screen = $_
        foreach ($state in 'idle', 'checking') {
            $script:Gui.S.PlanState = $state
            Get-MbtGuiPlanFooter | Should -Be 'Start unlocks when the plan is checked.'
        }
    }
    It 'says the plan cannot be used when the service refused it' {
        $script:Screen = 'exPlan'; $script:Gui.S.PlanState = 'error'
        Get-MbtGuiPlanFooter | Should -Be 'The plan cannot be used. Press Check again, or go back and change your choice.'
    }
    It 'asks for the tick of an import, then says it is ready' {
        $script:Screen = 'imPlan'; $script:Gui.S.PlanState = 'ok'
        Get-MbtGuiPlanFooter | Should -Be 'Tick the box to unlock Start.'
        $script:Gui.S.Confirmed = $true
        Get-MbtGuiPlanFooter | Should -Be 'Ready. Nothing is written before you press Start.'
    }
    It 'keeps the sentence of the page once the plan is valid' {
        $script:Screen = 'exPlan'; $script:Gui.S.PlanState = 'ok'
        Get-MbtGuiPlanFooter | Should -Be 'Read the checks and the plan. Nothing starts before you press Start.'
        $script:Screen = 'cpPlan'
        Get-MbtGuiPlanFooter | Should -Be 'Read the plan. Nothing starts before you press Start.'
    }
}

Describe 'A plan check that fails' {
    BeforeAll {
        . (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Export.ps1')
        . (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Run.ps1')
        $script:Tech = 'Exception calling "ResolvePlan" with "5" argument(s): "GET https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/MBX:x@contoso.com/folders/AAMk%3D -> {0} UnknownError: "'
    }
    It 'says a 5xx of the service in plain words, with Check again, not the raw exception' -ForEach @(500, 502, 503, 504) {
        $raw = 'error The task stopped because of an error. What was done is kept. ' + ($script:Tech -f $_)
        $text = Get-MbtGuiPlanFailureText -Raw $raw -Message 'The task stopped because of an error. What was done is kept.'
        $text | Should -Be 'Microsoft did not answer correctly (a temporary problem). Press Check again.'
        (ConvertTo-MbtGuiMessage $raw).Temporary | Should -Be 'service'
    }
    It 'says a network failure in plain words, with Check again' {
        $text = Get-MbtGuiPlanFailureText -Raw 'NetworkError GET https://graph.microsoft.com/v1.0/x failed: No such host is known.' -Message 'x'
        $text | Should -Be 'We cannot reach Microsoft. Check your internet connection, then press Check again.'
    }
    It 'keeps the sentence of the engine when the cause is not known, and says a known cause in its own words' {
        Get-MbtGuiPlanFailureText -Raw 'error stopped GET x -> 404 ErrorItemNotFound: gone' -Message 'The task stopped because of an error.' | Should -Be 'The task stopped because of an error.'
        Get-MbtGuiPlanFailureText -Raw 'error stopped GET x -> 400 Bad' | Should -Be 'The service did not accept the request.'
        Get-MbtGuiPlanFailureText -Raw 'ErrorAccessDenied: Access is denied' -Message 'stopped' | Should -BeLike 'This sign-in is not allowed*'
    }
    It 'does not take a number that is only inside an identifier for a status' {
        (ConvertTo-MbtGuiMessage 'GET https://graph.microsoft.com/v1.0/x/folders/AAMk5025Yw -> 404 ErrorItemNotFound').Temporary | Should -Be ''
    }
}

Describe 'The plan page once the service has counted the sizes' {
    BeforeAll {
        . (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Export.ps1')
        function Get-Sel([long]$Bytes, [int]$Unknown) { [pscustomobject]@{ Entries = @(); Items = 12; Bytes = $Bytes; Folders = 3; UnknownSize = $Unknown; SubtreeUnread = $false } }
        function Get-Obj([string]$Key, [long]$Bytes, [string]$State = 'planned', [string]$Step = 'plan') { @{ Key = $Key; Step = $Step; State = $State; Total = 2L; Bytes = $Bytes } }
    }
    BeforeEach { $script:Gui = @{ S = @{ Task = 'export'; PlanState = 'ok'; PlanTotals = [pscustomobject]@{ Items = 12; Bytes = 9000L; Folders = 3; UnknownSize = 0; FoldersToCreate = 0 } }; ExportSelection = (Get-Sel 5000 1) } }
    It 'takes the size the service counted when the browse view knew none for some folder and the service knows every size' {
        $sel = Get-MbtGuiPlanSelection
        $sel.Bytes | Should -Be 9000
        $sel.UnknownSize | Should -Be 0
        $sel.Items | Should -Be 12
    }
    It 'keeps the selection of the browse view while the plan is not counted, when the service left a size out, and when none was unknown' {
        $script:Gui.S.PlanState = 'checking'
        (Get-MbtGuiPlanSelection).UnknownSize | Should -Be 1
        $script:Gui.S.PlanState = 'ok'; $script:Gui.S.PlanTotals.UnknownSize = 1
        (Get-MbtGuiPlanSelection).Bytes | Should -Be 5000
        $script:Gui.S.PlanTotals.UnknownSize = 0; $script:Gui.ExportSelection = Get-Sel 5000 0
        (Get-MbtGuiPlanSelection).Bytes | Should -Be 5000
        $script:Gui.S.PlanTotals = $null; $script:Gui.ExportSelection = Get-Sel 5000 1
        (Get-MbtGuiPlanSelection).UnknownSize | Should -Be 1
    }
    It 'knows the size counted for a folder by its path in the mailbox, never for a path counted twice or without a size' {
        $script:Gui.PlanRun = @{ Objects = @((Get-Obj 'Calendar' 73597), (Get-Obj 'Inbox' -1), (Get-Obj 'Archive/Same' 10), (Get-Obj 'Archive/Same' 20), (Get-Obj 'Refused' 5 'refused'), (Get-Obj '(folders)' 3), (Get-Obj 'Other' 7 'planned' 'download')) }
        $sizes = Get-MbtGuiCountedSizes
        $sizes['Calendar'] | Should -Be 73597
        $sizes['calendar'] | Should -Be 73597 -Because 'a path is not case sensitive'
        $sizes.ContainsKey('Inbox') | Should -BeFalse
        $sizes.ContainsKey('Archive/Same') | Should -BeFalse
        $sizes.ContainsKey('Refused') | Should -BeFalse
        $sizes.ContainsKey('(folders)') | Should -BeFalse
        $sizes.ContainsKey('Other') | Should -BeFalse
        $sizes.Count | Should -Be 1
    }
    It 'counts nothing before the check is done, for an import, or without a run' {
        $script:Gui.PlanRun = @{ Objects = @((Get-Obj 'Calendar' 100)) }
        $script:Gui.S.PlanState = 'checking'
        (Get-MbtGuiCountedSizes).Count | Should -Be 0
        $script:Gui.S.PlanState = 'ok'; $script:Gui.S.Task = 'import'
        (Get-MbtGuiCountedSizes).Count | Should -Be 0
        $script:Gui.S.Task = 'export'; $script:Gui.Remove('PlanRun')
        (Get-MbtGuiCountedSizes).Count | Should -Be 0
    }
}

Describe 'Disk space wording' {
    It 'says disk space, not room, in every text of the console, the guides and the tests' {
        $root = Split-Path $script:Package -Parent
        $hits = foreach ($file in Get-ChildItem -LiteralPath $root -Recurse -File -Include '*.ps1', '*.psm1', '*.md', '*.html', '*.cs', '*.txt' | Where-Object { $_.FullName -notmatch '[\\/](\.git|out)[\\/]' -and $_.Name -ne 'Gui.Logic.Tests.ps1' }) {
            Select-String -LiteralPath $file.FullName -Pattern 'enough room' -SimpleMatch | ForEach-Object { "$($file.Name):$($_.LineNumber)" }
        }
        @($hits) | Should -BeNullOrEmpty
    }
}

Describe 'Disk check sentences of the To page' {
    It 'the check says there is enough disk space, or that there is not enough disk space on the drive' {
        $text = Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Export.ps1') -Raw
        $text | Should -Match 'There is enough disk space\.'
        $text | Should -Match 'Not enough disk space on \$drive\.'
    }
}

Describe 'What a step that is done shows' {
    BeforeAll { . (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Run.ps1') }

    It 'a step that is done without a message of its own keeps none of its progress text, and keeps a real final message' {
        $r = New-MbtGuiRun -Task export
        Add-Evt $r plan -Fields @{ steps = @([pscustomobject]@{ id = 'tree'; title = 'Read the folder tree' }, [pscustomobject]@{ id = 'list'; title = 'List the items' }, [pscustomobject]@{ id = 'bad'; title = 'Write' }) }
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'running'; msg = 'Measuring 1 folder the service gives no size for: 1 done' }
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'running'; done = 1 }
        $r.Steps[0].Msg | Should -Be 'Measuring 1 folder the service gives no size for: 1 done' -Because 'a running step without a message keeps what it said'
        Add-Evt $r step -Fields @{ id = 'tree'; state = 'done' }
        $r.Steps[0].Msg | Should -Be ''
        Get-MbtGuiStepSub $r.Steps[0] | Should -Be '' -Because 'nothing is left to say'
        Add-Evt $r step -Fields @{ id = 'list'; state = 'running'; msg = 'Counting' }
        Add-Evt $r step -Fields @{ id = 'list'; state = 'done'; msg = '4,443 folders read' }
        $r.Steps[1].Msg | Should -Be '4,443 folders read'
        Get-MbtGuiStepSub $r.Steps[1] | Should -Be '4,443 folders read'
        Add-Evt $r step -Fields @{ id = 'bad'; state = 'running'; msg = 'Writing part 2' }
        Add-Evt $r step -Fields @{ id = 'bad'; state = 'failed' }
        $r.Steps[2].Msg | Should -Be 'Writing part 2' -Because 'a step that failed keeps where it was'
    }

    It 'the line of a step never says more done than the total, and a step that is done says what it did' {
        Get-MbtGuiStepSub @{ State = 'running'; Msg = ''; Done = 5; Total = 8 } | Should -Be '5 / 8 items'
        Get-MbtGuiStepSub @{ State = 'running'; Msg = ''; Done = 9; Total = 8 } | Should -Be '9 items'
        Get-MbtGuiStepSub @{ State = 'running'; Msg = ''; Done = 1; Total = 1 } | Should -Be '1 / 1 item'
        Get-MbtGuiStepSub @{ State = 'done'; Msg = ''; Done = 53; Total = 53 } | Should -Be '53 items'
        Get-MbtGuiStepSub @{ State = 'done'; Msg = ''; Done = 60; Total = 53 } | Should -Be '60 items'
        Get-MbtGuiStepSub @{ State = 'done'; Msg = ''; Done = 0; Total = 4 } | Should -Be '4 items'
        Get-MbtGuiStepSub @{ State = 'done'; Msg = ''; Done = 0; Total = 0 } | Should -Be ''
    }

    It 'a number that is done is shown with its total only while it is below it; a size is only about right' {
        Format-MbtGuiProgressOf 5 8 | Should -Be '5 of 8 items'
        Format-MbtGuiProgressOf 9 8 | Should -Be '9 items'
        Format-MbtGuiProgressOf 1 0 | Should -Be '1 item'
        Format-MbtGuiProgressOf 8 8 -Finished | Should -Be '8 items'
        Format-MbtGuiFraction 5 8 | Should -Be '5 / 8'
        Format-MbtGuiFraction 12 10 | Should -Be '12'
        Format-MbtGuiFraction 10 10 -Finished | Should -Be '10'
        Format-MbtGuiFraction 0 10 -Finished | Should -Be '10'
        Format-MbtGuiBytesOf 5MB 8MB | Should -Be '5 MiB of about 8 MiB'
        Format-MbtGuiBytesOf 9MB 8MB | Should -Be '9 MiB'
        Format-MbtGuiBytesOf 8.2MB 7.9MB -Finished | Should -Be '8 MiB'
        Format-MbtGuiBytesOf 3MB 0 | Should -Be '3 MiB'
    }

    It 'the run model keeps what a download really did when the announced size was smaller' {
        $r = New-MbtGuiRun -Task export
        Add-Evt $r plan -Fields @{ steps = @([pscustomobject]@{ id = 'download'; title = 'Download the items' }) }
        Add-Evt $r step -Fields @{ id = 'download'; state = 'running'; done = 0; total = 53; bytes = 0; bytesTotal = 8320848 }
        Add-Evt $r step -Fields @{ id = 'download'; state = 'done'; done = 53; total = 53; bytes = 8620875; bytesTotal = 8320848 }
        Format-MbtGuiBytesOf $r.Bytes $r.BytesTotal -Finished | Should -Be '8 MiB'
        Format-MbtGuiBytesOf $r.Bytes $r.BytesTotal | Should -Be '8 MiB' -Because 'more than announced'
        Get-MbtGuiStepSub $r.Steps[0] | Should -Be '53 items'
    }
}

Describe 'The All jobs button' {
    BeforeAll {
        # the callbacks of the button run in a scope of their own: the stand-ins of the dialog and of the browser are global for the length of these tests
        function global:Show-MbtGuiDialog { param($Title, $Body, $Buttons) $global:MbtTestSaid = @{ Title = $Title; Body = $Body } }
        function global:Start-Process { param($FilePath) $global:MbtTestStarted = $FilePath }
        function New-OverviewGui($Engine) { $script:Gui = @{ S = @{ WorkFolder = 'W:\work' }; Engine = $Engine } }
    }
    AfterAll {
        Remove-Item -Path Function:\Show-MbtGuiDialog, Function:\Start-Process -ErrorAction SilentlyContinue
        Remove-Variable -Name MbtTestSaid, MbtTestStarted -Scope Global -ErrorAction SilentlyContinue
    }
    It 'asks the engine for the overview of the work folder and opens the page it gives' {
        $global:MbtTestStarted = $null
        $page = Join-Path $TestDrive 'overview.html'
        Set-Content -LiteralPath $page -Value '<html></html>'
        $seen = @{}
        New-OverviewGui ([pscustomobject]@{ JobOverview = { param($Work, $OnDone, $OnError) $seen.Work = $Work; & $OnDone $page }.GetNewClosure() })
        Open-MbtGuiJobOverview
        $seen.Work | Should -Be 'W:\work'
        $global:MbtTestStarted | Should -Be $page
    }
    It 'says so when no file was written (the demonstration), and when the page could not be written' {
        $global:MbtTestStarted = $null
        $global:MbtTestSaid = $null
        New-OverviewGui ([pscustomobject]@{ JobOverview = { param($Work, $OnDone, $OnError) & $OnDone $null } })
        Open-MbtGuiJobOverview
        $global:MbtTestSaid.Title | Should -Be 'Nothing to open here'
        $global:MbtTestSaid = $null
        New-OverviewGui ([pscustomobject]@{ JobOverview = { param($Work, $OnDone, $OnError) & $OnError 'The disk is full.' } })
        Open-MbtGuiJobOverview
        $global:MbtTestSaid.Title | Should -Be 'The overview of all the jobs could not be written'
        $global:MbtTestSaid.Body | Should -Be 'The disk is full.'
        $global:MbtTestSaid = $null
        New-OverviewGui ([pscustomobject]@{ JobOverview = { param($Work, $OnDone, $OnError) throw 'boom' } })
        Open-MbtGuiJobOverview
        $global:MbtTestSaid.Body | Should -Be 'boom'
        $global:MbtTestStarted | Should -BeNullOrEmpty
    }
    It 'is on the page of the jobs to continue and in the card Reports and log of the Next steps page, and the engines answer it' {
        $pages = Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.ps1') -Raw
        $pages | Should -Match "(?s)-Title 'Continue a job'.*?-Right \(New-GuiButton -Text 'All jobs'.*?Open-MbtGuiJobOverview"
        $run = Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Run.ps1') -Raw
        $run | Should -Match "(?s)function New-MbtGuiReportsCard.*?New-GuiButton -Text 'All jobs'.*?Open-MbtGuiJobOverview.*?-Title 'Reports and log'"
        (Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Engine.ps1') -Raw) | Should -Match '(?s)\$e\.JobOverview = .*?Write-MbtJobOverview -WorkFolder \$work'
        (Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Fake.ps1') -Raw) | Should -Match '\$engine\.JobOverview = \{ param\(\$Work, \$OnDone, \$OnError\) & \$OnDone \$null \}'
    }
    It 'the overview is written again when a run ends' {
        $run = Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Run.ps1') -Raw
        $run | Should -Match '(?s)function Complete-MbtGuiRun.*?Write-MbtGuiRunReport.*?\$reportDone = \{.*?\$g\.Engine\.JobOverview \$work'
    }
    It 'a batch row of the jobs to continue opens its report and never continues anything' {
        $run = Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Pages.Run.ps1') -Raw
        $run | Should -Match '(?s)function Resume-MbtGuiJob.*?\$Entry\.Task -eq ''batch''\) \{ Open-MbtGuiBatchReport \$Entry; return \}.*?\$reqFile'
        $run | Should -Match '(?s)function Open-MbtGuiBatchReport.*?\$g\.Engine\.BatchReport \$g\.S\.WorkFolder'
    }
}
