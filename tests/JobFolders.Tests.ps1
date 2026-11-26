#Requires -Version 7.6
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the job folders: what a job says about itself in plain words (why it stopped, when, the step it reached), the list of the jobs (the ones that are not
    finished first, none cut), finding a job by name whatever the case, deleting a job (what is refused, what is deleted and what never is), what Continue does for each mode, and the
    job folder of a single job started from the command line (-ResumeJob, -RemoveJob, -Mode Status). Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\JobFolders.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport

    $script:Inv = [Globalization.CultureInfo]::InvariantCulture
    function Get-TestIso([datetime]$Utc) { $Utc.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", $script:Inv) }
    # the local time the plain words give for an instant (the same whatever the language of Windows)
    function Get-TestWhen([string]$Iso) { [datetime]::Parse($Iso, $script:Inv, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal).ToLocalTime().ToString('yyyy-MM-dd HH:mm', $script:Inv) }
    # a job folder with a status.json (and a request.json when given); the fields of $Status are written as they are
    function New-TestJob {
        param([Parameter(Mandatory)][string]$Work, [Parameter(Mandatory)][string]$Name, [hashtable]$Status = @{}, [hashtable]$Request)
        $dir = Join-Path (Join-Path $Work 'jobs') $Name
        $null = New-Item -ItemType Directory -Force -Path $dir
        $s = [ordered]@{ v = 1; jobId = $Name; task = 'export'; title = "Export $Name"; mode = 'Export'; state = 'stopped'; updated = '2026-11-24T21:41:00.000Z'; started = '2026-11-24T20:00:00.000Z' }
        foreach ($k in $Status.Keys) { $s[$k] = $Status[$k] }
        $s | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $dir 'status.json') -Encoding utf8
        if ($Request) { $Request | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath (Join-Path $dir 'request.json') -Encoding utf8 }
        return $dir
    }
    function Get-TestJobOf([string]$Work, [string]$Name) { @(Get-MbtJob -WorkFolder $Work -Name $Name)[0] }
    function New-TestWork { $w = Join-Path $TestDrive ('work-' + [guid]::NewGuid().ToString('N').Substring(0, 8)); $null = New-Item -ItemType Directory -Force -Path $w; return $w }
}

Describe 'The plain words about a job' {
    BeforeEach { $script:W = New-TestWork }
    It 'a job stopped by you, from a status.json of an older version (no step, no error, no counts), says when and that nothing is lost' {
        $null = New-TestJob $script:W 'old-stopped' @{ state = 'stopped'; reason = 'stopped' }
        $d = Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'old-stopped')
        $d.Kind | Should -Be 'stopped'
        $d.Why | Should -Be "Stopped by you on $(Get-TestWhen '2026-11-24T21:41:00.000Z') (Stop and keep progress). Nothing is lost."
        $d.Progress | Should -Be ''
    }
    It 'a job that stopped (state running, its process gone) says that the PC was shut down or the tool was closed' {
        $null = New-TestJob $script:W 'gone' @{ state = 'running'; pid = 2147483000; counts = @{ done = 5120; total = 9860 } }
        $j = Get-TestJobOf $script:W 'gone'
        $j.Interrupted | Should -BeTrue
        $d = Get-MbtJobDescription -Job $j
        $d.Kind | Should -Be 'interrupted'
        $d.Why | Should -Be "Interrupted on $(Get-TestWhen '2026-11-24T21:41:00.000Z'): the PC was shut down or restarted, or the tool was closed, while it was running. What was done is kept."
        $d.Progress | Should -Be '5,120 of 9,860 items done'
    }
    It 'the step reached and its counts are said: step 3 of 5, the title, done of total' {
        $null = New-TestJob $script:W 'step' @{ state = 'running'; pid = 2147483000; step = @{ id = 'send'; title = 'Send the items'; index = 3; of = 5; done = 5120; total = 9860 } }
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'step')).Progress | Should -Be 'Step 3 of 5, Send the items, 5,120 of 9,860 done'
    }
    It 'a step without a total says how many are done, a step without counts says only the step' {
        $null = New-TestJob $script:W 'a' @{ step = @{ id = 'tree'; title = 'Read the folder tree'; index = 1; of = 4; done = 1240 } }
        $null = New-TestJob $script:W 'b' @{ step = @{ id = 'tree'; title = 'Read the folder tree'; index = 1; of = 4 } }
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'a')).Progress | Should -Be 'Step 1 of 4, Read the folder tree, 1,240 done'
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'b')).Progress | Should -Be 'Step 1 of 4, Read the folder tree'
    }
    It 'a copy says which of its two parts the step belongs to' {
        $null = New-TestJob $script:W 'reading' @{ mode = 'Copy'; task = 'mailcopy'; step = @{ id = 'download'; title = 'Download the items'; index = 3; of = 9 } }
        $null = New-TestJob $script:W 'adding' @{ mode = 'Copy'; task = 'mailcopy'; step = @{ id = 'send'; title = 'Send the items'; index = 7; of = 9 } }
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'reading')).Progress | Should -Be 'Step 3 of 9, Download the items (reading the source mailbox)'
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'adding')).Progress | Should -Be 'Step 7 of 9, Send the items (adding to the target mailbox)'
    }
    It 'an error says its short plain reason when the file has it, and only that it was an error when it has not' {
        $null = New-TestJob $script:W 'new-error' @{ state = 'error'; error = @{ code = 'disk-full'; message = 'The disk is full. Free some space and press Try again. What was downloaded is kept.' } }
        $null = New-TestJob $script:W 'old-error' @{ state = 'error' }
        $null = New-TestJob $script:W 'generic-error' @{ state = 'error'; error = @{ code = 'error'; message = 'The task stopped because of an error. What was done is kept.' } }
        $when = Get-TestWhen '2026-11-24T21:41:00.000Z'
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'new-error')).Why | Should -Be "Stopped by an error on ${when}: The disk is full. Free some space and press Try again. What was downloaded is kept."
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'old-error')).Why | Should -Be "Stopped by an error on $when. What was done is kept."
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'generic-error')).Why | Should -Be "Stopped by an error on $when. What was done is kept."
    }
    It 'a sign-in that expired, a job that waits for a decision, a job that is finished, a job that runs now' {
        $null = New-TestJob $script:W 'signin' @{ state = 'resumable'; reason = 'signin-required'; next = 'signin' }
        $null = New-TestJob $script:W 'decision' @{ state = 'needs-decision'; counts = @{ unknown = 3 } }
        $null = New-TestJob $script:W 'one' @{ state = 'needs-decision' }
        $null = New-TestJob $script:W 'done' @{ state = 'complete' }
        $null = New-TestJob $script:W 'refused' @{ state = 'complete'; reason = 'refused-items' }
        $null = New-TestJob $script:W 'partial' @{ state = 'resumable' }
        $null = New-TestJob $script:W 'live' @{ state = 'running'; pid = $PID; started = (Get-TestIso ([datetime]::UtcNow)); runStarted = (Get-TestIso ([datetime]::UtcNow)) }
        $when = Get-TestWhen '2026-11-24T21:41:00.000Z'
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'signin')).Why | Should -Be "Stopped on $when because the sign-in expired. Sign in again, then continue. Nothing is lost."
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'decision')).Why | Should -Be "Almost done on ${when}: 3 items need your decision."
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'one')).Why | Should -Be "Almost done on ${when}: an item needs your decision."
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'done')).Why | Should -Be "Finished on $when."
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'refused')).Why | Should -Be "Finished on $when. Some items were refused for good: the report lists them."
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'partial')).Why | Should -Be "Not finished on $when`: what is left is done when you continue. Nothing is lost."
        $live = Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'live')
        $live.Kind | Should -Be 'running'
        $live.Why | Should -Be 'Running now, in another window or process.'
    }
    It 'a job with nothing in its status but a state never raises under strict mode' {
        $null = New-TestJob $script:W 'bare' @{}
        { Set-StrictMode -Version Latest; Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'bare') } | Should -Not -Throw
        { Set-StrictMode -Version Latest; Get-MbtJobDescription -Job ([pscustomobject]@{ State = 'error' }) } | Should -Not -Throw
    }
}

Describe 'What a job writes in status.json' {
    AfterEach { Close-MbtEventLog }
    It 'the step that works (index, of, title, counts) is written at once, and stays when the job ends' {
        $dir = Join-Path $TestDrive 'status-step'
        Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -JobId 'j1' -Task export -Title 'Export' -Mode Export
        Write-MbtJobPlan -Steps @(@{ Id = 'tree'; Title = 'Read the folder tree' }, @{ Id = 'list'; Title = 'List the items' }, @{ Id = 'download'; Title = 'Download the items' })
        Write-MbtJobStep -Id tree -State running
        Write-MbtJobStep -Id tree -State done
        Write-MbtJobStep -Id download -State running -Done 120 -Total 900
        $s = Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json
        $s.step.id | Should -Be 'download'; $s.step.index | Should -Be 3; $s.step.of | Should -Be 3; $s.step.title | Should -Be 'Download the items'
        $s.step.done | Should -Be 120; $s.step.total | Should -Be 900
        Write-MbtJobResult -ExitCode 2 -Reason stopped -Counts @{ done = 120 }
        $s = Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json
        $s.state | Should -Be 'stopped'; $s.reason | Should -Be 'stopped'; $s.step.index | Should -Be 3
    }
    It 'the plain reason of the error that ended the job is written with the state' {
        $dir = Join-Path $TestDrive 'status-error'
        Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -JobId 'j2' -Task import -Title 'Import' -Mode Import
        $null = Write-MbtJobError -Message 'There is not enough space on the disk'
        Write-MbtJobResult -ExitCode 1
        $s = Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json
        $s.state | Should -Be 'error'; $s.error.code | Should -Be 'disk-full'; $s.error.message | Should -BeLike 'The disk is full*'
    }
    It 'a job that runs again begins with the step that reads what is already done, and the new run clears the reason of the one before' {
        $dir = Join-Path $TestDrive 'status-resume'
        $file = Join-Path $dir 'events.jsonl'
        Open-MbtEventLog -Path $file -JobId 'j3' -Task export -Title 'Export' -Mode Export
        Write-MbtJobPlan -Steps @(@{ Id = 'tree'; Title = 'Read the folder tree' })
        Write-MbtJobResult -ExitCode 2 -Reason stopped
        Close-MbtEventLog
        $first = @(Get-Content $file | ForEach-Object { $_ | ConvertFrom-Json }) | Where-Object e -EQ 'plan'
        @($first.steps).Count | Should -Be 1
        Open-MbtEventLog -Path $file -JobId 'j3' -Task export -Title 'Export' -Mode Export
        (Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json).reason | Should -BeNullOrEmpty
        Write-MbtJobPlan -Steps @(@{ Id = 'tree'; Title = 'Read the folder tree' })
        Write-MbtJobStep -Id tree -State running
        Close-MbtEventLog
        $events = @(Get-Content $file | ForEach-Object { $_ | ConvertFrom-Json })
        $plans = @($events | Where-Object e -EQ 'plan')
        $plans.Count | Should -Be 2
        $plans[1].steps[0].id | Should -Be 'resume'
        $plans[1].steps[0].title | Should -Be 'Resume: reading what is already done'
        $plans[1].steps[1].id | Should -Be 'tree'
        $resume = @($events | Where-Object { $_.e -eq 'step' -and $_.id -eq 'resume' })
        $resume[0].state | Should -Be 'running'
        $resume[-1].state | Should -Be 'done'
        (Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json).step.of | Should -Be 1
    }
    It 'a first run has no resume step' {
        $dir = Join-Path $TestDrive 'status-first'
        Open-MbtEventLog -Path (Join-Path $dir 'events.jsonl') -JobId 'j4' -Task export -Title 'Export' -Mode Export
        Write-MbtJobPlan -Steps @(@{ Id = 'tree'; Title = 'Read the folder tree' })
        Close-MbtEventLog
        $plan = @(Get-Content (Join-Path $dir 'events.jsonl') | ForEach-Object { $_ | ConvertFrom-Json }) | Where-Object e -EQ 'plan'
        @($plan.steps | Where-Object id -EQ 'resume').Count | Should -Be 0
    }
    It 'a job that runs again is alive from the start of its new process, not from its first start' {
        $w = New-TestWork
        $file = Join-Path $w 'jobs\again\events.jsonl'
        $null = New-TestJob $w 'again' @{ state = 'stopped'; started = '2026-01-01T00:00:00.000Z' }
        Set-Content -LiteralPath $file -Value '{"v":1,"e":"plan"}'
        Open-MbtEventLog -Path $file -JobId 'again' -Task export -Title 'Export' -Mode Export
        $j = Get-TestJobOf $w 'again'
        $j.Alive | Should -BeTrue
    }
}

Describe 'The list of the jobs' {
    It 'every job that is not finished is listed, newest first, before the finished ones; only the finished ones are cut; a plan check is not a job' {
        $w = New-TestWork
        foreach ($n in 1..12) { $null = New-TestJob $w ('20261101-10{0:00}-open' -f $n) @{ state = 'stopped'; updated = ('2026-11-01T10:{0:00}:00.000Z' -f $n) } }
        foreach ($n in 1..12) { $null = New-TestJob $w ('20261102-10{0:00}-done' -f $n) @{ state = 'complete'; updated = ('2026-11-02T10:{0:00}:00.000Z' -f $n) } }
        $null = New-TestJob $w '20261103-1000-export-plan-ab12' @{ state = 'error'; updated = '2026-11-03T10:00:00.000Z' }
        $l = Get-MbtJobListing -WorkFolder $w -FinishedLimit 5
        @($l.Unfinished).Count | Should -Be 12
        @($l.Unfinished)[0].Name | Should -Be '20261101-1012-open'
        @($l.Unfinished)[-1].Name | Should -Be '20261101-1001-open'
        @($l.Finished).Count | Should -Be 5
        $l.FinishedTotal | Should -Be 12
        @($l.Finished)[0].Name | Should -Be '20261102-1012-done'
        foreach ($j in @($l.Unfinished) + @($l.Finished)) { $j.PSObject.Properties.Name | Should -Contain 'Why'; $j.PSObject.Properties.Name | Should -Contain 'Progress' }
        @(@($l.Unfinished) + @($l.Finished) | Where-Object Name -Like '*-plan-*').Count | Should -Be 0
    }
    It 'an interrupted job and an error are unfinished, a stopped one too' {
        $w = New-TestWork
        $null = New-TestJob $w 'a-gone' @{ state = 'running'; pid = 2147483000 }
        $null = New-TestJob $w 'b-error' @{ state = 'error' }
        $null = New-TestJob $w 'c-decision' @{ state = 'needs-decision' }
        $null = New-TestJob $w 'd-stopped' @{ state = 'stopped' }
        $null = New-TestJob $w 'e-resumable' @{ state = 'resumable' }
        $null = New-TestJob $w 'f-done' @{ state = 'complete' }
        $l = Get-MbtJobListing -WorkFolder $w
        @($l.Unfinished).Count | Should -Be 5
        @($l.Finished).Count | Should -Be 1
    }
}

Describe 'Finding a job by its name' {
    It 'the name of the folder or the job id, whatever the case; never a path' {
        $w = New-TestWork
        $null = New-TestJob $w '20261124-2241-export-3fa2' @{ jobId = 'export-3fa2' }
        (Find-MbtJob -WorkFolder $w -Name '20261124-2241-EXPORT-3FA2').Name | Should -Be '20261124-2241-export-3fa2'
        (Find-MbtJob -WorkFolder $w -Name 'Export-3FA2').Name | Should -Be '20261124-2241-export-3fa2'
        Find-MbtJob -WorkFolder $w -Name 'nothing' | Should -BeNullOrEmpty
        Find-MbtJob -WorkFolder $w -Name '..\work' | Should -BeNullOrEmpty
        Find-MbtJob -WorkFolder $w -Name '' | Should -BeNullOrEmpty
    }
}

Describe 'Deleting a job' {
    BeforeEach { $script:W = New-TestWork }
    It 'is refused for a job that runs in another process' {
        $null = New-TestJob $script:W 'live' @{ state = 'running'; pid = $PID; started = (Get-TestIso ([datetime]::UtcNow)); runStarted = (Get-TestIso ([datetime]::UtcNow)) }
        $plan = Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'live')
        $plan.Allowed | Should -BeFalse
        $plan.Reason | Should -BeLike "*is running now*"
        { Remove-MbtJob -WorkFolder $script:W -Name 'LIVE' -Confirm:$false } | Should -Throw '*is running now*'
        Test-Path (Join-Path $script:W 'jobs\live') | Should -BeTrue
    }
    It 'is refused for a job of a batch that is not complete, and allowed for a job of a batch that is complete' {
        $null = New-TestJob $script:W 'in-open' @{ state = 'stopped'; batch = 'Open' }
        $null = New-TestJob $script:W 'in-done' @{ state = 'complete'; batch = 'Done' }
        $row = { param($Id, $State, $Job) [ordered]@{ n = 1; upn = "$Id@contoso.com"; location = 'Primary'; selection = ''; output = ''; store = ''; state = $State; job = $Job; attempts = 1 } }
        $batch = { param($Id, $Rows) [ordered]@{ v = 1; id = $Id; name = $Id; created = '2026-11-10T09:00:00.000Z'; updated = '2026-11-10T09:00:00.000Z'; state = ''; mode = 'Export'; csv = ''; parallel = 1; options = @{}; counts = $null; rows = $Rows } }
        Save-MbtBatch -Batch (& $batch 'Open' @((& $row 'a' 'partial' 'in-open'), (& $row 'b' 'pending' ''))) -Folder (Join-Path $script:W 'batches\Open')
        Save-MbtBatch -Batch (& $batch 'Done' @((& $row 'c' 'complete' 'in-done'))) -Folder (Join-Path $script:W 'batches\Done')
        $open = Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'in-open')
        $open.Allowed | Should -BeFalse
        $open.Reason | Should -BeLike '*batch Open*not complete*-ResumeBatch*'
        (Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'in-done')).Allowed | Should -BeTrue
    }
    It 'offers the copy store only in the stores folder, refuses it while the import part of the copy is not complete, and refuses it while another job that is not finished uses it' {
        $store = Join-Path $script:W 'stores\copy-adele-primary-1a2b'
        $null = New-Item -ItemType Directory -Force -Path $store
        Set-Content -LiteralPath (Join-Path $store 'catalog.db') -Value 'x'
        $null = New-TestJob $script:W 'copy-stopped' @{ mode = 'Copy'; task = 'mailcopy'; state = 'stopped'; store = $store; step = @{ id = 'send'; title = 'Send the items'; index = 7; of = 9 } }
        $p = Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'copy-stopped')
        $p.Allowed | Should -BeTrue
        $p.Store.Offered | Should -BeTrue
        $p.Store.Allowed | Should -BeFalse
        $p.Store.Reason | Should -BeLike 'The copy is not complete: its second part, adding the items to the target mailbox, is not finished*'
        { Remove-MbtJob -WorkFolder $script:W -Name 'copy-stopped' -IncludeStore -Confirm:$false } | Should -Throw '*not complete*'
        Test-Path $store | Should -BeTrue
        Test-Path (Join-Path $script:W 'jobs\copy-stopped') | Should -BeTrue
        $null = New-TestJob $script:W 'copy-done' @{ mode = 'Copy'; task = 'mailcopy'; state = 'complete'; store = $store }
        $null = New-TestJob $script:W 'copy-other' @{ mode = 'Copy'; task = 'mailcopy'; state = 'stopped'; store = $store }
        $shared = Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'copy-done')
        $shared.Store.Allowed | Should -BeFalse
        $shared.Store.Reason | Should -BeLike "*used by the job 'copy-other'*"
        Remove-Item -LiteralPath (Join-Path $script:W 'jobs\copy-other'), (Join-Path $script:W 'jobs\copy-stopped') -Recurse -Force
        $alone = Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'copy-done') -WithSize
        $alone.Store.Allowed | Should -BeTrue
        $alone.Store.Bytes | Should -BeGreaterThan 0
        $outside = Join-Path $TestDrive 'user-store'
        $null = New-Item -ItemType Directory -Force -Path $outside
        $null = New-TestJob $script:W 'copy-outside' @{ mode = 'Copy'; task = 'mailcopy'; state = 'complete'; store = $outside }
        $o = Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'copy-outside')
        $o.Store.Offered | Should -BeFalse
        $o.Store.Reason | Should -BeLike '*outside the stores folder*'
    }
    It 'deletes the folder of the job and nothing else: the store of the user, the PST files and the other jobs stay' {
        $userStore = Join-Path $TestDrive ('user-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        $null = New-Item -ItemType Directory -Force -Path $userStore
        Set-Content -LiteralPath (Join-Path $userStore 'catalog.db') -Value 'x'
        $pst = Join-Path $TestDrive ('mail-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.pst')
        Set-Content -LiteralPath $pst -Value 'pst'
        $dir = New-TestJob $script:W 'to-delete' @{ state = 'error'; store = $userStore } @{ StorePath = $userStore; OutFolder = (Split-Path $pst -Parent) }
        Set-Content -LiteralPath (Join-Path $dir 'events.jsonl') -Value '{}'
        Set-Content -LiteralPath (Join-Path $dir 'report.html') -Value '<html/>'
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $dir 'temp\deep')
        Set-Content -LiteralPath (Join-Path $dir 'temp\deep\x.bin') -Value 'x'
        $keep = New-TestJob $script:W 'kept' @{ state = 'stopped' }
        $r = Remove-MbtJob -WorkFolder $script:W -Name 'TO-DELETE' -Confirm:$false
        $r.Removed | Should -BeTrue
        $r.StoreRemoved | Should -BeFalse
        Test-Path $dir | Should -BeFalse
        Test-Path $keep | Should -BeTrue
        Test-Path (Join-Path $userStore 'catalog.db') | Should -BeTrue
        Test-Path $pst | Should -BeTrue
    }
    It 'never follows a link inside the folder of the job' {
        $target = Join-Path $TestDrive ('linked-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        $null = New-Item -ItemType Directory -Force -Path $target
        Set-Content -LiteralPath (Join-Path $target 'precious.txt') -Value 'keep'
        $dir = New-TestJob $script:W 'with-link' @{ state = 'stopped' }
        $null = New-Item -ItemType Junction -Path (Join-Path $dir 'link') -Target $target
        $null = Remove-MbtJob -WorkFolder $script:W -Name 'with-link' -Confirm:$false
        Test-Path $dir | Should -BeFalse
        Test-Path (Join-Path $target 'precious.txt') | Should -BeTrue
    }
    It 'deletes the copy store of a complete copy only when it is asked, and only that store' {
        $store = Join-Path $script:W 'stores\copy-a'
        $other = Join-Path $script:W 'stores\copy-b'
        foreach ($d in $store, $other) { $null = New-Item -ItemType Directory -Force -Path (Join-Path $d 'packs'); Set-Content -LiteralPath (Join-Path $d 'packs\p.bin') -Value 'x' }
        $null = New-TestJob $script:W 'c1' @{ mode = 'Copy'; task = 'mailcopy'; state = 'complete'; store = $store }
        $null = New-TestJob $script:W 'c2' @{ mode = 'Copy'; task = 'mailcopy'; state = 'complete'; store = $store }
        $first = Remove-MbtJob -WorkFolder $script:W -Name 'c1' -Confirm:$false
        $first.StoreRemoved | Should -BeFalse
        Test-Path $store | Should -BeTrue
        $second = Remove-MbtJob -WorkFolder $script:W -Name 'c2' -IncludeStore -Confirm:$false
        $second.StoreRemoved | Should -BeTrue
        Test-Path $store | Should -BeFalse
        Test-Path (Join-Path $other 'packs\p.bin') | Should -BeTrue
        Test-Path (Join-Path $script:W 'stores') | Should -BeTrue
    }
    It 'with -WhatIf nothing is deleted' {
        $dir = New-TestJob $script:W 'whatif' @{ state = 'stopped' }
        $r = Remove-MbtJob -WorkFolder $script:W -Name 'whatif' -WhatIf
        $r.Removed | Should -BeFalse
        Test-Path $dir | Should -BeTrue
    }
    It 'warns that an unfinished job cannot be continued afterwards, and that the journal of a PST import goes with it' {
        $dir = New-TestJob $script:W 'pst-job' @{ state = 'stopped'; mode = 'ImportPst'; task = 'import' }
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $dir 'work')
        Set-Content -LiteralPath (Join-Path $dir 'work\pstimport.db') -Value 'x'
        $p = Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'pst-job')
        @($p.Warnings).Count | Should -Be 2
        $p.Warnings[0] | Should -BeLike 'This job is not finished*can no longer continue*'
        $p.Warnings[1] | Should -BeLike 'The record of what was already imported from the PST file is in this folder*a second time*'
    }
    It 'is refused for a job whose folder holds the store of the user, and for a job whose files another job uses' {
        $dir = New-TestJob $script:W 'holds-store' @{ state = 'stopped'; mode = 'Export'; store = $null } @{ StorePath = (Join-Path $script:W 'jobs\holds-store\store') }
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $dir 'store')
        (Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'holds-store')).Reason | Should -BeLike '*is inside its folder*'
        $null = New-TestJob $script:W 'a-owner' @{ state = 'stopped' }
        $null = New-TestJob $script:W 'b-user' @{ state = 'stopped' } @{ SelectionFiles = @((Join-Path $script:W 'jobs\a-owner\sel.json')) }
        (Get-MbtJobRemovalPlan -WorkFolder $script:W -Job (Get-TestJobOf $script:W 'a-owner')).Reason | Should -BeLike "*the job 'b-user' uses*"
    }
}

Describe 'What Continue does, for each mode' {
    BeforeEach { $script:W = New-TestWork }
    It 'an export downloads only what is left and reads the tree and the list again' {
        $null = New-TestJob $script:W 'x' @{ mode = 'Export'; state = 'running'; pid = 2147483000; step = @{ id = 'download'; title = 'Download the items'; index = 3; of = 4 } } @{ Tries = 3 }
        $t = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'x') -Request ([pscustomobject]@{ Tries = 3 })
        $t.Already | Should -Be @('The items already downloaded and checked in the export store are not downloaded again.')
        $t.Again[0] | Should -BeLike 'The folder tree and the list of the items are read again from the mailbox*'
        $t.Again[1] | Should -Be 'Items that failed are tried again, up to 3 tries in all.'
        $t.From | Should -Be 'It restarts from step 1, Read the folder tree, and downloads only what is left. Before it stopped it had reached step 3 of 4, Download the items.'
        $t.Notes[0] | Should -BeLike 'The PC was shut down*'
    }
    It 'an export with PST files keeps the parts that are complete' {
        $null = New-TestJob $script:W 'p' @{ mode = 'Export' }
        $t = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'p') -Request ([pscustomobject]@{ Dest = 'pst' })
        ($t.Again -join ' ') | Should -BeLike '*the parts already written and checked are kept; the part that was being written when it stopped is written again from its start*'
    }
    It 'a PST conversion keeps the parts that are complete and writes the unfinished part again from its start' {
        $null = New-TestJob $script:W 'e' @{ mode = 'ExportPst' }
        $t = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'e')
        $t.Already[0] | Should -BeLike 'The PST parts that were completed are kept*'
        $t.Again[0] | Should -Be 'The part that was being written when it stopped is deleted and written again from its start.'
    }
    It 'an import from a store does not send what is done, looks up what was in flight and tries the failures again' {
        $null = New-TestJob $script:W 'i' @{ mode = 'Import' }
        $t = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'i') -Request ([pscustomobject]@{ Tries = 5 })
        $t.Already[0] | Should -BeLike 'The items already added to the mailbox*are not sent again.'
        ($t.Again -join ' ') | Should -BeLike '*They are looked up in the target folder first*an item found there counts as done, an item that is not found is sent again*'
        $t.Again[-1] | Should -Be 'Items that failed are tried again, up to 5 tries in all.'
        $t.From | Should -BeLike 'It restarts from step 1, Verify the export store*'
    }
    It 'an import from a PST file speaks of its journal and of messages' {
        $null = New-TestJob $script:W 'ip' @{ mode = 'ImportPst' }
        $t = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'ip') -Request ([pscustomobject]@{ JournalPath = 'D:\work\pstimport.db' })
        $t.Already[0] | Should -BeLike '*journal of the import records them (D:\work\pstimport.db)*are not added again*'
        ($t.Again -join ' ') | Should -BeLike '*Messages that were being sent*a message found there counts as done*'
        $t.From | Should -BeLike 'It restarts from step 1, Check the PST files*'
    }
    It 'an import of .eml files does not send again what is done and leaves the uncertain files to you' {
        $null = New-TestJob $script:W 'em' @{ mode = 'ImportEml' }
        $t = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'em')
        $t.Already[0] | Should -BeLike 'The files recorded as done in import-state.jsonl*'
        ($t.Again -join ' ') | Should -BeLike '*is not sent again by itself*-ResolveEml Imported or NotImported*'
    }
    It 'an export of .eml files does not write again the files that are in the manifest and in place' {
        $null = New-TestJob $script:W 'ee' @{ mode = 'ExportEml' }
        (Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'ee')).Already[0] | Should -BeLike 'The files already written, listed as exported in the manifest*are not written again.'
    }
    It 'a copy says which of its two parts it resumes' {
        $null = New-TestJob $script:W 'c-read' @{ mode = 'Copy'; task = 'mailcopy'; step = @{ id = 'download'; title = 'Download the items'; index = 3; of = 9 } }
        $null = New-TestJob $script:W 'c-add' @{ mode = 'Copy'; task = 'mailcopy'; step = @{ id = 'send'; title = 'Send the items'; index = 7; of = 9 } }
        $null = New-TestJob $script:W 'c-old' @{ mode = 'Copy'; task = 'mailcopy' }
        $read = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'c-read')
        $read.From | Should -BeLike 'It resumes the first part, reading the source mailbox: from step 1 of 9, Read the folder tree.*'
        $read.Already[0] | Should -Be 'The items already read into the store are not read again.'
        $add = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'c-add')
        $add.From | Should -BeLike 'It resumes the second part, adding the items to the target mailbox: from step 5 of 9, Verify the export store.*'
        $add.Already[0] | Should -Be 'The first part is complete: the source mailbox is not read again, and the store is used as it is.'
        $old = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'c-old')
        $old.From | Should -BeLike 'It resumes the part where it stopped*'
    }
    It 'a copy of the Primary and the Archive counts 14 steps and says which part the step belongs to' {
        $null = New-TestJob $script:W 'b-first' @{ mode = 'Copy'; task = 'mailcopy'; step = @{ id = 'send'; title = 'Send the items (Primary part)'; index = 7; of = 14 } }
        $null = New-TestJob $script:W 'b-second' @{ mode = 'Copy'; task = 'mailcopy'; step = @{ id = 'send-2'; title = 'Send the items (Archive part)'; index = 12; of = 14 } }
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'b-first')).Progress | Should -Be 'Step 7 of 14, Send the items (Primary part) (adding to the target mailbox)'
        (Get-MbtJobDescription -Job (Get-TestJobOf $script:W 'b-second')).Progress | Should -Be 'Step 12 of 14, Send the items (Archive part) (adding to the target mailbox, part 2)'
        $text = Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'b-second')
        $text.From | Should -BeLike 'It resumes the second part, adding the items to the target mailbox: from step 5 of 14, Verify the export store.*'
        $text.Already | Should -Contain 'A part of the copy that is complete (the Primary part, or the Archive part) is not done again: the next run continues with the part that is not.'
    }
    It 'the console takes a copy of both places of the command line with its layout and the folder of each part' {
        $c = New-MbtCliJob -WorkFolder $script:W -Mode Copy -Task mailcopy -Parameter @{ Mode = 'Copy'; UserPrincipalName = 'alex@contoso.com'; Mailbox = 'All'; TargetLayout = 'Archive'; TargetRoot = 'A'; TargetArchiveRoot = 'B'; AppOnly = [switch]$true; TargetMailboxId = 'adele@contoso.com' } -WorkingDirectory $TestDrive -ScriptPath 'x' -Title 't'
        $r = Read-MbtJobRequest -Folder $c.Folder
        $r.McFromKind | Should -Be 'All'; $r.McLayout | Should -Be 'Archive'; $r.TgtName | Should -Be 'A'; $r.TgtName2 | Should -Be 'B'; $r.TgtMode2 | Should -Be 'new'
    }
    It 'says why a job cannot be continued: it runs, it belongs to a batch, it has no request, a command-line job needs its folder or its secret' {
        $folder = Join-Path $TestDrive 'ran-here'
        $null = New-Item -ItemType Directory -Force -Path $folder
        $null = New-TestJob $script:W 'live' @{ state = 'running'; pid = $PID; started = (Get-TestIso ([datetime]::UtcNow)); runStarted = (Get-TestIso ([datetime]::UtcNow)) }
        $null = New-TestJob $script:W 'batched' @{ batch = 'Wave1' }
        $null = New-TestJob $script:W 'bare' @{}
        $null = New-TestJob $script:W 'cli' @{}
        (Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'live') -Request ([pscustomobject]@{ Source = 'cli' })).Problem | Should -BeLike 'This job is running now*'
        (Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'batched') -Request ([pscustomobject]@{})).Problem | Should -BeLike 'This job belongs to the batch Wave1*-ResumeBatch Wave1*'
        (Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'bare')).Problem | Should -BeLike 'This job has no request.json*'
        (Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'cli') -Request ([pscustomobject]@{ Source = 'cli'; WorkingDirectory = $folder })).Problem | Should -BeNullOrEmpty
        (Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'cli') -Request ([pscustomobject]@{ Source = 'cli'; WorkingDirectory = (Join-Path $TestDrive 'gone') })).Problem | Should -BeLike "*is not there any more*"
        (Get-MbtJobResumeText -Job (Get-TestJobOf $script:W 'cli') -Request ([pscustomobject]@{ Source = 'cli'; WorkingDirectory = $folder; NeedsSecret = $true })).Problem | Should -BeLike '*certificate password*'
    }
}

Describe 'The job folder of a command-line run' {
    BeforeEach { $script:W = New-TestWork; $script:Cwd = Join-Path $TestDrive 'cwd'; $null = New-Item -ItemType Directory -Force -Path $script:Cwd }
    It 'records the command, the folder it ran in and the paths, never a password' {
        $params = @{ Mode = 'ImportPst'; UserPrincipalName = 'adele.vance@contoso.com'; PstPath = @('D:\pst\a.pst', 'D:\pst\b.pst'); TargetRoot = 'Imported'; MergeWellKnown = [switch]$true; CertificatePassword = (ConvertTo-SecureString 'Zq9-unique-pw' -AsPlainText -Force); ArrivedFrom = [datetime]'2026-01-01' }
        $j = New-MbtCliJob -WorkFolder $script:W -Mode ImportPst -Task import -Parameter $params -WorkingDirectory $script:Cwd -ScriptPath 'C:\tool\Invoke-MailboxTransfer.ps1' -Resolved @{ StorePath = 'C:\x\store'; TempPath = 'C:\x\transfer'; JournalPath = 'C:\x\transfer\pstimport.db' } -Title 'Import from PST'
        $j.Reused | Should -BeFalse
        $j.Name | Should -Match '^\d{8}-\d{4}-import-[0-9a-f]{4}$'
        $r = Read-MbtJobRequest -Folder $j.Folder
        $r.Source | Should -Be 'cli'
        $r.WorkingDirectory | Should -Be $script:Cwd
        $r.JournalPath | Should -Be 'C:\x\transfer\pstimport.db'
        $r.StorePath | Should -Be 'C:\x\store'
        $r.NeedsSecret | Should -BeTrue
        (Get-Content (Join-Path $j.Folder 'request.json') -Raw) | Should -Not -BeLike '*Zq9-unique-pw*'
        $r.CommandLine | Should -BeLike '.\Invoke-MailboxTransfer.ps1 -Mode ImportPst*-PstPath D:\pst\a.pst,D:\pst\b.pst*-TargetRoot Imported*'
        $r.Task | Should -Be 'import'; $r.ImSrc | Should -Be 'pst'; $r.ImPath | Should -Be 'D:\pst\a.pst'; $r.TgtMode | Should -Be 'new'; $r.TgtName | Should -Be 'Imported'
        $r.Parameters.PSObject.Properties.Name | Should -Not -Contain 'CertificatePassword'
        $splat = Get-MbtCliJobParameter -Request $r
        $splat['MergeWellKnown'].IsPresent | Should -BeTrue
        @($splat['PstPath']).Count | Should -Be 2
        $splat['ArrivedFrom'] | Should -BeLike '2026-01-01*'
    }
    It 'the same command in the same folder continues the job while it is not complete, another command or another folder makes another job' {
        $params = @{ Mode = 'Export'; UserPrincipalName = 'adele.vance@contoso.com'; StorePath = '.\st' }
        $first = New-MbtCliJob -WorkFolder $script:W -Mode Export -Task export -Parameter $params -WorkingDirectory $script:Cwd -ScriptPath 'x' -Title 't'
        Write-MbtJobStatus -JobFolder $first.Folder -State resumable -JobId $first.JobId
        $again = New-MbtCliJob -WorkFolder $script:W -Mode Export -Task export -Parameter $params -WorkingDirectory $script:Cwd -ScriptPath 'x' -Title 't'
        $again.Reused | Should -BeTrue
        $again.Name | Should -Be $first.Name
        $other = New-MbtCliJob -WorkFolder $script:W -Mode Export -Task export -Parameter @{ Mode = 'Export'; UserPrincipalName = 'alex.wilber@contoso.com'; StorePath = '.\st' } -WorkingDirectory $script:Cwd -ScriptPath 'x' -Title 't'
        $other.Reused | Should -BeFalse
        $elsewhere = New-MbtCliJob -WorkFolder $script:W -Mode Export -Task export -Parameter $params -WorkingDirectory $TestDrive -ScriptPath 'x' -Title 't'
        $elsewhere.Reused | Should -BeFalse
        # how it signs in does not make it another command
        $signin = New-MbtCliJob -WorkFolder $script:W -Mode Export -Task export -Parameter ($params + @{ NoPrompt = [switch]$true }) -WorkingDirectory $script:Cwd -ScriptPath 'x' -Title 't'
        $signin.Name | Should -Be $first.Name
        Write-MbtJobStatus -JobFolder $first.Folder -State complete -JobId $first.JobId
        (New-MbtCliJob -WorkFolder $script:W -Mode Export -Task export -Parameter $params -WorkingDirectory $script:Cwd -ScriptPath 'x' -Title 't').Reused | Should -BeFalse
    }
    It 'gives the console what it needs to continue a job of the command line: the task and the values of its pages' {
        $e = New-MbtCliJob -WorkFolder $script:W -Mode Export -Task export -Parameter @{ Mode = 'Export'; UserPrincipalName = 'adele.vance@contoso.com'; OutputPst = 'D:\out\adele.pst'; PstMaxSizeGB = 10 } -WorkingDirectory $script:Cwd -ScriptPath 'x' -Title 't'
        $r = Read-MbtJobRequest -Folder $e.Folder
        $r.Task | Should -Be 'export'; $r.Dest | Should -Be 'pst'; $r.OutFolder | Should -Be 'D:\out'; $r.OutName | Should -Be 'adele.pst'; $r.SplitGB | Should -Be 10; $r.Mailbox | Should -Be 'adele.vance@contoso.com'
        $c = New-MbtCliJob -WorkFolder $script:W -Mode Copy -Task mailcopy -Parameter @{ Mode = 'Copy'; UserPrincipalName = 'adele.vance@contoso.com'; TargetMailbox = 'Archive'; TargetRoot = 'Copied'; AppOnly = [switch]$true; TargetMailboxId = 'projects@contoso.com' } -WorkingDirectory $script:Cwd -ScriptPath 'x' -Title 't'
        $r = Read-MbtJobRequest -Folder $c.Folder
        $r.Task | Should -Be 'mailcopy'; $r.TgtLoc | Should -Be 'Archive'; $r.TgtName | Should -Be 'Copied'; $r.McTarget | Should -Be 'projects@contoso.com'; $r.AuthMode | Should -Be 'app'
        $x = New-MbtCliJob -WorkFolder $script:W -Mode ExportEml -Task export -Parameter @{ Mode = 'ExportEml'; EmlPath = 'D:\eml' } -WorkingDirectory $script:Cwd -ScriptPath 'x' -Title 't'
        (Read-MbtJobRequest -Folder $x.Folder).Dest | Should -Be 'eml'
        (Read-MbtJobRequest -Folder 'C:\no\such\folder') | Should -BeNullOrEmpty
    }
    It 'the default work folder is the environment variable when it is set' {
        $saved = $env:MAILBOXTRANSFER_WORK_FOLDER
        try {
            $env:MAILBOXTRANSFER_WORK_FOLDER = Join-Path $TestDrive 'from-env'
            Get-MbtDefaultWorkFolder | Should -Be (Join-Path $TestDrive 'from-env')
            $env:MAILBOXTRANSFER_WORK_FOLDER = $null
            Get-MbtDefaultWorkFolder | Should -BeLike '*\MailboxTransfer'
        } finally { $env:MAILBOXTRANSFER_WORK_FOLDER = $saved }
    }
}

Describe 'The entry script: a job of the command line, -Mode Status, -ResumeJob, -RemoveJob' {
    BeforeAll {
        $script:Work = Join-Path $TestDrive 'cli-work'
        $script:RunDir = Join-Path $TestDrive 'cli-run'
        $script:OtherDir = Join-Path $TestDrive 'cli-other'
        foreach ($d in $script:Work, $script:RunDir, $script:OtherDir) { $null = New-Item -ItemType Directory -Force -Path $d }
        function Invoke-CliChild {
            param([string]$In, [string[]]$Arguments)
            Push-Location -LiteralPath $In
            try { $out = & pwsh -NoProfile -NonInteractive -File $script:Entry -DevelopmentMode -WorkFolder $script:Work -LogFolder (Join-Path $script:Work 'logs') @Arguments 2>&1 | Out-String; $code = $LASTEXITCODE } finally { Pop-Location }
            [pscustomobject]@{ Exit = $code; Output = ($out -replace '\x1b\[[0-9;]*m', '') }
        }
        function Get-CliJobFolder { @(Get-ChildItem -LiteralPath (Join-Path $script:Work 'jobs') -Directory | Where-Object { $_.Name -match '-export-' }) }
        $script:Run1 = Invoke-CliChild $script:RunDir @('-Mode', 'Export', '-UserPrincipalName', 'adele.vance@contoso.com', '-NoPrompt', '-StorePath', '.\st')
    }
    It 'a single job writes its job folder in the work folder: status, events, and the request with the command and its folder' {
        $script:Run1.Exit | Should -Be 2
        $folders = @(Get-CliJobFolder)
        $folders.Count | Should -Be 1
        $dir = $folders[0].FullName
        foreach ($f in 'status.json', 'events.jsonl', 'request.json') { Test-Path (Join-Path $dir $f) | Should -BeTrue -Because $f }
        $r = Get-Content (Join-Path $dir 'request.json') -Raw | ConvertFrom-Json
        $r.Source | Should -Be 'cli'
        $r.WorkingDirectory | Should -Be $script:RunDir
        $r.StorePath | Should -Be (Join-Path $script:RunDir 'st')
        $r.CommandLine | Should -BeLike '.\Invoke-MailboxTransfer.ps1 -Mode Export*-StorePath .\st*'
        $s = Get-Content (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json
        $s.state | Should -Be 'resumable'; $s.reason | Should -Be 'signin-required'; $s.step.index | Should -Be 1
    }
    It 'says in its Next steps the short command that continues the job' {
        $name = (Get-CliJobFolder)[0].Name
        $script:Run1.Output | Should -BeLike "*Continue this job: .\Invoke-MailboxTransfer.ps1 -ResumeJob $name*"
    }
    It 'the same command again continues the same job folder' {
        $again = Invoke-CliChild $script:RunDir @('-Mode', 'Export', '-UserPrincipalName', 'adele.vance@contoso.com', '-NoPrompt', '-StorePath', '.\st')
        $again.Exit | Should -Be 2
        @(Get-CliJobFolder).Count | Should -Be 1
        $again.Output | Should -BeLike '*the same command again: this job continues*'
    }
    It '-Mode Status lists the job with why it stopped and the command that continues it' {
        $name = (Get-CliJobFolder)[0].Name
        $status = Invoke-CliChild $script:OtherDir @('-Mode', 'Status')
        $status.Exit | Should -Be 0
        $status.Output | Should -BeLike "*$name*"
        $status.Output | Should -BeLike '*because the sign-in expired*'
        $status.Output | Should -BeLike "*continue: -ResumeJob $name*"
    }
    It '-ResumeJob runs the recorded command in the folder it ran in, from another folder, whatever the case of the name, in the same job folder' {
        $name = (Get-CliJobFolder)[0].Name
        $resume = Invoke-CliChild $script:OtherDir @('-ResumeJob', $name.ToUpperInvariant(), '-NoPrompt')
        $resume.Exit | Should -Be 2
        $resume.Output | Should -BeLike "*Resume the job $name*"
        $resume.Output | Should -BeLike "*runs in: $($script:RunDir)*Store*cli-run*st*"
        @(Get-CliJobFolder).Count | Should -Be 1
        $events = @(Get-Content (Join-Path (Get-CliJobFolder)[0].FullName 'events.jsonl') | ForEach-Object { $_ | ConvertFrom-Json })
        @($events | Where-Object e -EQ 'plan').Count | Should -BeGreaterThan 1
        @($events | Where-Object { $_.e -eq 'plan' })[-1].steps[0].id | Should -Be 'resume'
    }
    It '-ResumeJob of a job that does not exist is an error that says how to list the jobs' {
        $x = Invoke-CliChild $script:OtherDir @('-ResumeJob', 'nothing-like-it')
        $x.Exit | Should -Not -Be 0
        ($x.Output -replace '\s*\|\s*', ' ' -replace '\s+', ' ') | Should -BeLike "*No job is named 'nothing-like-it'*-Mode Status lists the jobs*"
    }
    It '-WhatIf creates no job folder' {
        $before = @(Get-CliJobFolder).Count
        $null = Invoke-CliChild $script:RunDir @('-Mode', 'Export', '-UserPrincipalName', 'alex.wilber@contoso.com', '-NoPrompt', '-StorePath', '.\st2', '-WhatIf')
        @(Get-CliJobFolder).Count | Should -Be $before
    }
    It '-RemoveJob with -WhatIf deletes nothing, with -Confirm:$false deletes the job folder only, whatever the case of the name' {
        $dir = (Get-CliJobFolder)[0]
        $whatIf = Invoke-CliChild $script:OtherDir @('-RemoveJob', $dir.Name.ToUpperInvariant(), '-WhatIf')
        $whatIf.Exit | Should -Be 0
        $whatIf.Output | Should -BeLike '*WhatIf: nothing was deleted*'
        Test-Path $dir.FullName | Should -BeTrue
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $script:RunDir 'st')
        Set-Content -LiteralPath (Join-Path $script:RunDir 'st\catalog.db') -Value 'x'
        $real = Invoke-CliChild $script:OtherDir @('-RemoveJob', $dir.Name.ToUpperInvariant(), '-Confirm:$false')
        $real.Exit | Should -Be 0
        $real.Output | Should -BeLike '*the job folder was deleted*'
        Test-Path $dir.FullName | Should -BeFalse
        Test-Path (Join-Path $script:RunDir 'st\catalog.db') | Should -BeTrue
    }
    It '-RemoveJob of a job that runs is refused with an error block and exit code 1' {
        $dir = Join-Path $script:Work 'jobs\20261125-1200-export-live'
        $null = New-Item -ItemType Directory -Force -Path $dir
        @{ v = 1; jobId = 'live'; task = 'export'; title = 'Export live'; state = 'running'; pid = $PID; started = (Get-TestIso ([datetime]::UtcNow)); runStarted = (Get-TestIso ([datetime]::UtcNow)); updated = (Get-TestIso ([datetime]::UtcNow)) } | ConvertTo-Json | Set-Content (Join-Path $dir 'status.json')
        $x = Invoke-CliChild $script:OtherDir @('-RemoveJob', '20261125-1200-EXPORT-LIVE', '-Confirm:$false')
        $x.Exit | Should -Be 1
        $x.Output | Should -BeLike '*is running now*'
        $x.Output | Should -BeLike '*RemoveJob failed*'
        Test-Path $dir | Should -BeTrue
    }
}

Describe 'The jobs of the demonstration say what the engine says' {
    BeforeAll {
        foreach ($f in 'Core', 'Run', 'Fake') { . (Join-Path $script:Root "module\MailboxTransfer.Gui.$f.ps1") }
        $script:Fake = New-MbtGuiFakeEngine
        $script:W = New-TestWork
        # the instant whose local time is the one the demonstration says
        function Get-TestUtcOf([string]$Local) { Get-TestIso ([datetime]::SpecifyKind([datetime]::Parse($Local, $script:Inv), [DateTimeKind]::Local).ToUniversalTime()) }
        $demo = @(& $script:Fake.Jobs $script:W)
        function Get-Demo([string]$Name) { @($demo | Where-Object Name -EQ $Name)[0] }
        $null = New-TestJob $script:W '20261124-2241-import-3fa2' @{ mode = 'ImportPst'; task = 'import'; state = 'running'; pid = 2147483000; updated = (Get-TestUtcOf '2026-11-24 22:41'); step = @{ id = 'send'; title = 'Send the items'; index = 3; of = 5; done = 5120; total = 9860 } }
        $null = New-TestJob $script:W '20261123-1630-mailcopy-7c1e' @{ mode = 'Copy'; task = 'mailcopy'; state = 'stopped'; reason = 'stopped'; updated = (Get-TestUtcOf '2026-11-23 17:05'); step = @{ id = 'download'; title = 'Download the items'; index = 3; of = 9; done = 12400; total = 31000 } }
        $null = New-TestJob $script:W '20261122-0912-export-b8f8' @{ mode = 'Export'; task = 'export'; state = 'error'; updated = (Get-TestUtcOf '2026-11-22 09:12'); error = @{ code = 'disk-full'; message = 'The disk is full. Free some space and press Try again. What was downloaded is kept.' }; step = @{ id = 'download'; title = 'Download the items'; index = 3; of = 4; done = 2300; total = 4100 } }
        $null = New-TestJob $script:W '20261110-0915-mailcopy-9b1c' @{ mode = 'Copy'; task = 'mailcopy'; state = 'complete'; updated = (Get-TestUtcOf '2026-11-10 11:48'); counts = @{ done = 14200 } }
        $null = New-TestJob $script:W '20261102-1000-export-0d44' @{ mode = 'Export'; task = 'export'; state = 'complete'; updated = (Get-TestUtcOf '2026-11-02 11:00') }
    }
    It 'the line of each job: why it stopped and the step reached' {
        foreach ($name in '20261124-2241-import-3fa2', '20261123-1630-mailcopy-7c1e', '20261122-0912-export-b8f8', '20261110-0915-mailcopy-9b1c', '20261102-1000-export-0d44') {
            $d = Get-MbtJobDescription -Job (Get-TestJobOf $script:W $name)
            (Get-Demo $name).Sub | Should -Be $d.Why -Because $name
            (Get-Demo $name).Sub2 | Should -Be $d.Progress -Because $name
        }
    }
    It 'the Resume view of each job: what is done, what is checked again, where it restarts' {
        $request = @{
            '20261124-2241-import-3fa2' = [pscustomobject]@{ JournalPath = 'D:\MailboxTransfer\jobs\20261124-2241-import-3fa2\work\pstimport.db' }
            '20261123-1630-mailcopy-7c1e' = [pscustomobject]@{}
            '20261122-0912-export-b8f8' = [pscustomobject]@{ Dest = 'pst' }
        }
        foreach ($name in $request.Keys) {
            $job = Get-TestJobOf $script:W $name
            $real = Get-MbtJobResumeText -Job $job -Request $request[$name]
            $fake = & $script:Fake.JobResume $script:W $name
            $fake.Heading | Should -Be $real.Heading -Because $name
            @($fake.Already) | Should -Be @($real.Already) -Because $name
            @($fake.Again) | Should -Be @($real.Again) -Because $name
            $fake.From | Should -Be $real.From -Because $name
            @($fake.Notes) | Should -Be @($real.Notes) -Because $name
            $d = Get-MbtJobDescription -Job $job
            $fake.Why | Should -Be $d.Why -Because $name
            $fake.Progress | Should -Be $d.Progress -Because $name
        }
    }
}
