#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - the progress per folder of the jobs, as the console reads it: the FolderTracker of the engine, and the object events
    the entry script writes for Export, Import, ExportPst, ImportPst, ExportEml and ImportEml (the Run page lists these folders), with the
    step events, the counts of the result, and a stop request for each mode. Nothing here reaches the network: the entry script runs in
    this process against FakeMailboxService (tests\FakeMailboxService.cs); the PST and EML files are made by the engine.

    Run:  Invoke-Pester -Path .\tests\Events.Folders.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot

    $script:StringDates = (Get-Command ConvertFrom-Json).Parameters.ContainsKey('DateKind')
    function ConvertFrom-JsonText { param([Parameter(ValueFromPipeline)][string]$Text) process { if ($script:StringDates) { $Text | ConvertFrom-Json -DateKind String } else { $Text | ConvertFrom-Json } } }
    function Read-Events { param([string]$Path) @(Get-Content -LiteralPath $Path | Where-Object { $_ } | ForEach-Object { $_ | ConvertFrom-JsonText }) }
}

Describe 'FolderTracker' {
    It 'tells planned, running, done for a folder, with the counts' {
        $t = [MailboxTransfer.FolderTracker]::new()
        $t.Expect('Inbox', 3)
        $first = $t.TakeChanged()
        $first.Count | Should -Be 1
        $first[0].State | Should -Be 'planned'; $first[0].Total | Should -Be 3; $first[0].Done | Should -Be 0
        $t.Done('Inbox', 100)
        $t.TakeChanged()[0].State | Should -Be 'running'
        $t.Done('Inbox', 50); $t.Done('Inbox', 25)
        $last = $t.TakeChanged()
        $last[0].State | Should -Be 'done'; $last[0].Done | Should -Be 3; $last[0].Bytes | Should -Be 175
        $t.TakeChanged().Count | Should -Be 0
    }
    It 'gives only the folders that changed, in the order they were first seen' {
        $t = [MailboxTransfer.FolderTracker]::new()
        $t.Expect('A', 2); $t.Expect('B', 2); $t.Expect('C', 2)
        [void]$t.TakeChanged()
        $t.Done('C'); $t.Done('A')
        $t.TakeChanged().Key | Should -Be @('A', 'C')
    }
    It 'adds up the totals of a folder told about twice and never shows a total below what was done' {
        $t = [MailboxTransfer.FolderTracker]::new()
        $t.Expect('F', 2); $t.Expect('F', 3)
        $t.Snapshot()[0].Total | Should -Be 5
        $t2 = [MailboxTransfer.FolderTracker]::new()
        $t2.Expect('G', 1); $t2.Done('G'); $t2.Done('G')
        $t2.Snapshot()[0].Total | Should -Be 2
    }
    It 'a folder with failures is partial, one with nothing but failures is failed, one with Unknown rows is check' {
        $t = [MailboxTransfer.FolderTracker]::new()
        $t.Expect('P', 3); $t.Done('P'); $t.Done('P'); $t.Fail('P')
        $t.Expect('F', 2); $t.Fail('F'); $t.Fail('F')
        $t.Expect('U', 2); $t.Done('U'); $t.Unknown('U')
        $s = @{}; foreach ($f in $t.Snapshot()) { $s[$f.Key] = $f }
        $s.P.State | Should -Be 'partial'; $s.P.Msg | Should -Match '1 failed'
        $s.F.State | Should -Be 'failed'
        $s.U.State | Should -Be 'check'; $s.U.Msg | Should -Match '1 to check'
    }
    It 'Settled turns an Unknown row into a done one (created) or lets it be sent again (not created), SettledFailed makes it failed' {
        $t = [MailboxTransfer.FolderTracker]::new()
        $t.Expect('F', 3); $t.Done('F'); $t.Unknown('F'); $t.Unknown('F')
        $t.Settled('F', $true, 10)
        $t.Settled('F', $false)
        $x = $t.Snapshot()[0]
        $x.Unknown | Should -Be 0; $x.Done | Should -Be 2
        $t.Done('F')
        $t.Snapshot()[0].State | Should -Be 'done'
        $u = [MailboxTransfer.FolderTracker]::new()
        $u.Expect('G', 1); $u.Unknown('G'); $u.SettledFailed('G')
        $u.Snapshot()[0].State | Should -Be 'failed'
    }
    It 'a folder whose total is not known is running until it is finished; Skip makes it partial; Mark gives an explicit state' {
        $t = [MailboxTransfer.FolderTracker]::new()
        $t.Done('K'); $t.Done('K')
        $k = $t.Snapshot()[0]; $k.State | Should -Be 'running'; $k.Total | Should -Be -1
        $t.Skip('K'); $t.Finish('K')
        $k = $t.Snapshot()[0]; $k.State | Should -Be 'partial'; $k.Total | Should -Be 3; $k.Msg | Should -Match '1 not exported'
        $t.Mark('Calendar', 'refused', 'folder class IPF.Appointment is not mail')
        $m = $t.Snapshot() | Where-Object Key -EQ 'Calendar'
        $m.State | Should -Be 'refused'; $m.Msg | Should -Match 'not mail'
    }
    It 'takes calls from several threads without losing a count' {
        $t = [MailboxTransfer.FolderTracker]::new()
        $t.Expect('X', 4000)
        1..4 | ForEach-Object -Parallel { $tr = $using:t; foreach ($i in 1..1000) { $tr.Done('X', 1) } } -ThrottleLimit 4
        $x = $t.Snapshot()[0]
        $x.Done | Should -Be 4000; $x.Bytes | Should -Be 4000; $x.State | Should -Be 'done'
    }
    It 'Sum adds up the folders' {
        $t = [MailboxTransfer.FolderTracker]::new()
        $t.Expect('A', 2); $t.Expect('B', 3); $t.Done('A', 5); $t.Fail('B')
        $s = $t.Sum()
        $s.Total | Should -Be 5; $s.Done | Should -Be 1; $s.Failed | Should -Be 1; $s.Bytes | Should -Be 5
    }
}

Describe 'FolderEventPump' {
    It 'writes the folders that changed as object events and a step event with the sums, and once more at Dispose' {
        $path = Join-Path $TestDrive 'pump.jsonl'
        $log = [MailboxTransfer.EventLog]::new($path, $null, 0)
        $t = [MailboxTransfer.FolderTracker]::new()
        $pump = [MailboxTransfer.FolderEventPump]::new($log, $t, 'export', 0, 'export')
        $t.Expect('Inbox', 2); $t.Done('Inbox', 10)
        $pump.Flush() | Should -Be 2
        $pump.Flush() | Should -Be 0
        $t.Done('Inbox', 10)
        $pump.Dispose()
        $log.Dispose()
        $ev = Read-Events $path
        $objs = @($ev | Where-Object e -EQ 'object')
        $objs.Count | Should -Be 2
        $objs[0].step | Should -Be 'export'; $objs[0].key | Should -Be 'Inbox'; $objs[0].state | Should -Be 'running'; $objs[0].done | Should -Be 1; $objs[0].total | Should -Be 2
        $objs[-1].state | Should -Be 'done'
        $steps = @($ev | Where-Object e -EQ 'step')
        $steps[-1].id | Should -Be 'export'; $steps[-1].done | Should -Be 2; $steps[-1].total | Should -Be 2
    }
}

Describe 'Entry script against the fake service: the folders of every mode' {
    BeforeAll {
        $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-folders-tests-" + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $script:Tmp
        $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $script:Src = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
        $script:Dst = "MBX:99999999-8888-7777-6666-555555555555@$($script:Tenant)"
        function New-Message([string]$Subject) {
            $b = [MailboxTransfer.Tests.FxBuilder]::new()
            $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $Subject).Str(0x1000, "Body of $Subject")
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
            $global:MbtFoClient = [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
            $global:LASTEXITCODE = $null
            $text = & $script:Entry @Parameters *>&1 | Out-String
            [pscustomobject]@{ Text = $text; Exit = $LASTEXITCODE }
        }
        function Get-Objects($Events, [string]$Step) { @($Events | Where-Object { $_.e -eq 'object' -and $_.step -eq $Step }) }
        # the last event of each folder of a step
        function Get-LastByKey($Events, [string]$Step) { $h = @{}; foreach ($o in (Get-Objects $Events $Step)) { $h[$o.key] = $o }; $h }

        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph $global:MbtFoClient
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport', 'Mail.Read'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }

        $script:Svc = New-Service
        $script:Store = Join-Path $script:Tmp 'store'
        # 1. export of the live mailbox into a store
        $script:EvExport = Join-Path $script:Tmp 'job-export\events.jsonl'
        $script:RunExport = Invoke-InProcess $script:Svc @{ Mode = 'Export'; StorePath = $script:Store; TempPath = (Join-Path $script:Tmp 'tmp'); PrimaryMailboxId = $script:Src; Concurrency = 2; EventsFile = $script:EvExport; JobId = 'exp-1' }
        # 2. import of that store into the other mailbox
        $script:EvImport = Join-Path $script:Tmp 'job-import\events.jsonl'
        $script:RunImport = Invoke-InProcess $script:Svc @{ Mode = 'Import'; StorePath = $script:Store; PrimaryMailboxId = $script:Dst; TargetRoot = 'R'; Concurrency = 2; EventsFile = $script:EvImport; JobId = 'imp-1' }
        # 3. the store written to a PST
        $script:Pst = Join-Path $script:Tmp 'out\mail.pst'
        $script:EvPst = Join-Path $script:Tmp 'job-pst\events.jsonl'
        $script:RunPst = Invoke-InProcess $script:Svc @{ Mode = 'ExportPst'; StorePath = $script:Store; OutputPst = $script:Pst; EventsFile = $script:EvPst; JobId = 'pst-1' }
        # 4. the store written as EML files
        $script:EmlDir = Join-Path $script:Tmp 'eml'
        $script:EvEml = Join-Path $script:Tmp 'job-eml\events.jsonl'
        $script:RunEml = Invoke-InProcess $script:Svc @{ Mode = 'ExportEml'; StorePath = $script:Store; EmlPath = $script:EmlDir; EventsFile = $script:EvEml; JobId = 'eml-1' }
    }
    AfterAll { Remove-Variable -Name MbtFoClient -Scope Global -ErrorAction SilentlyContinue; Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }

    It 'Export lists each folder: planned, then running, then done with its items (step download)' {
        $script:RunExport.Exit | Should -Be 0
        $ev = Read-Events $script:EvExport
        $last = Get-LastByKey $ev 'download'
        $last.Keys.Count | Should -Be 2
        $last['Inbox'].state | Should -Be 'done'; $last['Inbox'].done | Should -Be 6; $last['Inbox'].total | Should -Be 6
        $last['Inbox/Projects'].state | Should -Be 'done'; $last['Inbox/Projects'].done | Should -Be 4; $last['Inbox/Projects'].total | Should -Be 4
        ((Get-Objects $ev 'download') | Where-Object { $_.key -eq 'Inbox' } | Select-Object -First 1).state | Should -BeIn 'planned', 'running', 'done'
        $ev[-1].counts.done | Should -Be 10
    }
    It 'Import lists the folders of the store, with the items sent (step send)' {
        $script:RunImport.Exit | Should -Be 0
        $ev = Read-Events $script:EvImport
        $last = Get-LastByKey $ev 'send'
        $last['Inbox'].state | Should -Be 'done'; $last['Inbox'].done | Should -Be 6
        $last['Inbox/Projects'].state | Should -Be 'done'; $last['Inbox/Projects'].done | Should -Be 4
        $script:Svc.ItemsOf($script:Dst).Count | Should -Be 10
    }
    It 'ExportPst lists the folders of the PST (step pst) and writes the step events' {
        $script:RunPst.Exit | Should -Be 0
        $ev = Read-Events $script:EvPst
        $ev[0].steps.id | Should -Be @('pst', 'pstcheck')
        $last = Get-LastByKey $ev 'pst'
        $last.Keys.Count | Should -BeGreaterOrEqual 2
        $inbox = $last.Keys | Where-Object { $_ -like '*Inbox' -and $_ -notlike '*Projects*' } | Select-Object -First 1
        $inbox | Should -Not -BeNullOrEmpty
        $last[$inbox].state | Should -Be 'done'; $last[$inbox].done | Should -Be 6
        $proj = $last.Keys | Where-Object { $_ -like '*Projects' } | Select-Object -First 1
        $last[$proj].done | Should -Be 4
        $ev[-1].counts.itemsWritten | Should -Be 10
        Test-Path -LiteralPath $script:Pst | Should -BeTrue
    }
    It 'ExportEml lists the folders (step export), with the step event of the sums and the counts' {
        $script:RunEml.Exit | Should -Be 0
        $ev = Read-Events $script:EvEml
        $ev[0].steps.id | Should -Be @('export')
        $last = Get-LastByKey $ev 'export'
        $last['Inbox'].state | Should -Be 'done'; $last['Inbox'].done | Should -Be 6
        $last['Inbox/Projects'].done | Should -Be 4
        $steps = @($ev | Where-Object { $_.e -eq 'step' -and $_.id -eq 'export' })
        $steps[-1].state | Should -Be 'done'; $steps[-1].done | Should -Be 10
        $ev[-1].e | Should -Be 'result'; $ev[-1].counts.done | Should -Be 10; $ev[-1].counts.failed | Should -Be 0
        @(Get-ChildItem -LiteralPath $script:EmlDir -Filter *.eml -Recurse).Count | Should -Be 10
        (Get-Content (Join-Path (Split-Path $script:EvEml) 'status.json') -Raw | ConvertFrom-Json).state | Should -Be 'complete'
    }
    It 'ImportEml lists the folders below the target folder (step send) and the counts' {
        $svc = New-Service
        $ev = Join-Path $script:Tmp 'job-importeml\events.jsonl'
        $r = Invoke-InProcess $svc @{ Mode = 'ImportEml'; EmlPath = $script:EmlDir; PrimaryMailboxId = $script:Dst; TargetRoot = 'FromEml'; EventsFile = $ev; JobId = 'ieml-1' }
        $r.Exit | Should -Be 0
        $e = Read-Events $ev
        $e[0].steps.id | Should -Be @('send')
        $last = Get-LastByKey $e 'send'
        $last['FromEml/Inbox'].state | Should -Be 'done'; $last['FromEml/Inbox'].done | Should -Be 6
        $last['FromEml/Inbox/Projects'].done | Should -Be 4
        $e[-1].exit | Should -Be 0; $e[-1].counts.done | Should -Be 10; $e[-1].counts.failed | Should -Be 0
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
    }
    It 'ImportPst: plan with the five steps, the folders of the target (step send), the step final and the counts of the journal' {
        $svc = New-Service
        $ev = Join-Path $script:Tmp 'job-importpst\events.jsonl'
        $r = Invoke-InProcess $svc @{ Mode = 'ImportPst'; PstPath = $script:Pst; PrimaryMailboxId = $script:Dst; TargetRoot = 'FromPst'; TempPath = (Join-Path $script:Tmp 'pst-work'); Concurrency = 2; EventsFile = $ev; JobId = 'ipst-1' }
        $r.Exit | Should -Be 0
        $e = Read-Events $ev
        $e[0].steps.id | Should -Be @('check', 'folders', 'send', 'reconcile', 'final')
        $steps = @($e | Where-Object e -EQ 'step')
        ($steps | Where-Object { $_.id -eq 'check' -and $_.state -eq 'done' }).Count | Should -BeGreaterOrEqual 1
        ($steps | Where-Object { $_.id -eq 'send' -and $_.state -eq 'done' }).Count | Should -BeGreaterOrEqual 1
        ($steps | Where-Object { $_.id -eq 'final' -and $_.state -eq 'done' }).Count | Should -Be 1
        $last = Get-LastByKey $e 'send'
        $keys = @($last.Keys | Where-Object { $_ -like 'FromPst*' })
        $keys.Count | Should -BeGreaterOrEqual 2
        $inbox = $keys | Where-Object { $_ -like '*/Inbox' } | Select-Object -First 1
        $last[$inbox].state | Should -Be 'done'; $last[$inbox].done | Should -Be 6
        ($keys | ForEach-Object { $last[$_].done } | Measure-Object -Sum).Sum | Should -Be 10
        $e[-1].exit | Should -Be 0; $e[-1].counts.done | Should -Be 10; $e[-1].counts.failed | Should -Be 0; $e[-1].counts.total | Should -Be 10
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
    }
    It 'ImportPst -WhatIf lists the folders as plan objects and sends nothing' {
        $svc = New-Service
        $ev = Join-Path $script:Tmp 'job-importpst-plan\events.jsonl'
        $r = Invoke-InProcess $svc @{ Mode = 'ImportPst'; PstPath = $script:Pst; PrimaryMailboxId = $script:Dst; TargetRoot = 'FromPst'; TempPath = (Join-Path $script:Tmp 'pst-work2'); WhatIf = $true; EventsFile = $ev }
        $r.Exit | Should -Be 0
        $e = Read-Events $ev
        $plan = Get-Objects $e 'plan'
        $plan.Count | Should -BeGreaterOrEqual 2
        ($plan | Where-Object { $_.key -like '*Inbox' -and $_.key -notlike '*Projects*' } | Select-Object -First 1).total | Should -Be 6
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
    }
    It 'a stopped Export with -OutputPst ends resumable (exit 2): the half-made store is not converted and there is no error' {
        $svc = New-Service
        $dir = Join-Path $script:Tmp 'job-stop-exportpst-live'
        $null = New-Item -ItemType Directory -Force -Path $dir
        $stop = Join-Path $dir 'stop.request'
        Set-Content -LiteralPath $stop -Value 'stop'
        $pst = Join-Path $dir 'live.pst'
        $r = Invoke-InProcess $svc @{ Mode = 'Export'; StorePath = (Join-Path $dir 'store'); TempPath = (Join-Path $dir 'tmp'); PrimaryMailboxId = $script:Src; OutputPst = $pst; Concurrency = 2; EventsFile = (Join-Path $dir 'events.jsonl'); StopFile = $stop }
        $r.Exit | Should -Be 2
        $e = Read-Events (Join-Path $dir 'events.jsonl')
        @($e | Where-Object e -EQ 'error').Count | Should -Be 0
        $e[-1].exit | Should -Be 2; $e[-1].reason | Should -Be 'stopped'
        Test-Path -LiteralPath $pst | Should -BeFalse
        $r.Text | Should -Match 'was not converted'
    }
    It 'a stop request that exists at the start stops <Mode>: exit 2, reason stopped, next continue' -ForEach @(
        @{ Mode = 'ExportEml' }, @{ Mode = 'ImportEml' }, @{ Mode = 'ImportPst' }, @{ Mode = 'Import' }, @{ Mode = 'ExportPst' }
    ) {
        $svc = New-Service
        $dir = Join-Path $script:Tmp ("job-stop-$Mode")
        $null = New-Item -ItemType Directory -Force -Path $dir
        $stop = Join-Path $dir 'stop.request'
        Set-Content -LiteralPath $stop -Value 'stop'
        $p = switch ($Mode) {
            'ExportEml' { @{ Mode = 'ExportEml'; StorePath = $script:Store; EmlPath = (Join-Path $script:Tmp 'eml-stop'); NoResume = $true } }
            'ImportEml' { @{ Mode = 'ImportEml'; EmlPath = $script:EmlDir; PrimaryMailboxId = $script:Dst; TargetRoot = 'Stopped'; NoResume = $true } }
            'ImportPst' { @{ Mode = 'ImportPst'; PstPath = $script:Pst; PrimaryMailboxId = $script:Dst; TargetRoot = 'Stopped'; TempPath = (Join-Path $dir 'work') } }
            'Import' { @{ Mode = 'Import'; StorePath = $script:Store; PrimaryMailboxId = $script:Dst; TargetRoot = 'Stopped'; JobName = 'stopjob' } }
            'ExportPst' { @{ Mode = 'ExportPst'; StorePath = $script:Store; OutputPst = (Join-Path $dir 'stop.pst') } }
        }
        $p.EventsFile = Join-Path $dir 'events.jsonl'; $p.StopFile = $stop
        $r = Invoke-InProcess $svc $p
        $r.Exit | Should -Be 2
        $e = Read-Events $p.EventsFile
        $e[-1].e | Should -Be 'result'; $e[-1].exit | Should -Be 2; $e[-1].reason | Should -Be 'stopped'; $e[-1].next | Should -Be 'continue'
        (Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json).state | Should -Be 'stopped'
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
    }
}
