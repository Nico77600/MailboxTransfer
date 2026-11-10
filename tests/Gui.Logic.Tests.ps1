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
    It 'writes an effective configuration of the changes only, nothing when nothing changed' {
        Get-MbtGuiEffectiveConfig -State @{ Changed = @{} } | Should -BeNullOrEmpty
        $t = Get-MbtGuiEffectiveConfig -State @{ Changed = @{ Tries = $true; Retry = $true }; Tries = 3; Retry = 5 }
        $t | Should -Match 'MaxAttempts = 3'
        $t | Should -Match 'BaseBackoffSeconds = 5'
        $f = Join-Path $TestDrive 'eff.psd1'; Set-Content -LiteralPath $f -Value $t
        { Import-PowerShellDataFile -LiteralPath $f } | Should -Not -Throw
    }
}

Describe 'Commands of a job' {
    BeforeAll {
        $script:S = @{ Task = 'export'; Mailbox = 'adele@contoso.com'; Shared = ''; ImShared = ''; Dest = 'pst'; OutFolder = 'D:\Out'; OutName = 'adele.pst'; SplitGB = 20; CheckEach = $true; SepAux = $false; Conc = 4; Tries = 6; Reconcile = 5; Awake = $true; ImSrc = 'store'; ImPath = 'D:\Store'; TgtLoc = 'Archive'; TgtMode = 'new'; TgtName = 'Imported x'; Mapping = ''; RpPath = 'D:\a.pst'; RpDest = 'D:\a.repaired.pst'; RpLevel = 'Automatic'; RpKeep = $false; Changed = @{} }
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
        Get-MbtGuiEffectiveConfig -State $s | Should -BeNullOrEmpty
    }
    It 'the plan is the same request with -WhatIf' {
        (Get-MbtGuiCommand -Task export -S $script:S -Job $script:Job -Plan)[0].Arguments | Should -Contain '-WhatIf'
    }
    It 'several locations: one export each, then one PST write' {
        $j = $script:Job.Clone(); $j.SelectionFiles = @('a.json', 'b.json')
        $c = @(Get-MbtGuiCommand -Task export -S $script:S -Job $j)
        $c.Count | Should -Be 3
        $c[2].Arguments | Should -Contain 'ExportPst'
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
        $s = @{ Task = 'export'; OpPath = 'D:\open.pst'; Mailbox = 'adele@contoso.com'; Dest = 'pst'; OutFolder = 'D:\Out'; OutName = 'a.pst'; SplitGB = 20; CheckEach = $true; SepAux = $true; Conc = 4; Tries = 6; Reconcile = 5; Awake = $true; ImSrc = 'pst'; ImPath = 'D:\a.pst'; ImShared = 'p@contoso.com'; TgtLoc = 'Primary'; TgtMode = 'new'; TgtName = 'X'; Mapping = 'D:\m.psd1'; RpPath = 'D:\a.pst'; RpDest = 'D:\b.pst'; RpLevel = 'Automatic'; RpKeep = $true; Changed = @{} }
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
        $s = @{ Task = 'import'; Mailbox = 'a@contoso.com'; ImSrc = 'pst'; ImPath = 'D:\a.pst'; TgtLoc = 'Primary'; TgtMode = 'new'; TgtName = 'X'; Conc = 4; Tries = 6; Reconcile = 5; Awake = $true; Changed = @{}; Mapping = ''; ImShared = '' }
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
        $s = @{ Task = 'copy'; OpPath = 'D:\open.pst'; OutFolder = 'D:\Out'; OutName = 'copy.pst'; SplitGB = 20; CheckEach = $false; Conc = 4; Tries = 6; Awake = $true; Changed = @{} }
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
