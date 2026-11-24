#Requires -Version 7.6
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the copy of a list of mailboxes (module\MailboxTransfer.Batch.ps1 and the entry script with -Mode Copy and -MailboxCsv): the checks of the CSV file (every problem at once), the
    arguments of a copy, the scheduler on stand-in child processes (tests\support\batch-child.ps1: lines that write into the same target location, or read the same source, never run together; different
    ones run side by side; resume, sign-in, exit codes), and the command line (-WhatIf, the error card, the mode of a batch). Nothing here reaches the network; one copy of one mailbox into another against the
    fake service is in Copy.Tests.ps1.

    Run:  Invoke-Pester -Path .\tests\BatchCopy.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    $script:Stub = Join-Path $PSScriptRoot 'support\batch-child.ps1'
    $script:SavedLocal = $env:MAILBOXTRANSFER_LOCAL_CONFIG
    $env:MAILBOXTRANSFER_LOCAL_CONFIG = Join-Path $TestDrive 'no-local-config.psd1'
    Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    $script:AppOption = @{ AppOnly = $true; SignInUser = ''; StoreBase = ''; Pass = @{} }

    function New-Csv([string]$Name, [string]$Text, [switch]$Bom) {
        $dir = Join-Path $TestDrive 'csv'
        $null = New-Item -ItemType Directory -Force -Path $dir
        $path = Join-Path $dir $Name
        [IO.File]::WriteAllText($path, $Text, [Text.UTF8Encoding]::new([bool]$Bom))
        return $path
    }
    function New-Work([string]$Name) { $w = Join-Path $TestDrive $Name; $null = New-Item -ItemType Directory -Force -Path $w; return $w }
    # The batch of a list of copies, rows as the CSV reader makes them: lines of "source,target".
    function New-TestCopyBatch {
        param([string]$Work, [string]$Id, [string[]]$Line, [hashtable]$Option = $script:AppOption, [int]$Parallel = 1, [string]$Root = 'Copied')
        $csv = New-Csv "$Id.csv" ("SourceAddress,TargetAddress`r`n" + ($Line -join "`r`n"))
        $read = Read-MbtCopyBatchCsv -Path $csv -DefaultTargetRoot $Root -StoreBase (Join-Path $Work 'stores')
        $read.Problems | Should -BeNullOrEmpty
        return (New-MbtBatch -Id $Id -Csv $csv -Row $read.Rows -Option $Option -Parallel $Parallel -Mode Copy)
    }
    function Invoke-StubBatch {
        param([string]$Work, $Batch, [hashtable]$Option = $script:AppOption, [int]$Parallel = 1)
        $a = @{ WorkFolder = $Work; Batch = $Batch; EntryScript = $script:Stub; Option = $Option; Parallel = $Parallel; PollMilliseconds = 100; StartGapMilliseconds = 50; StopGraceSeconds = 20 }
        return (Invoke-MbtBatch @a 6>&1 | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.InformationRecord] })
    }
    function Get-BatchRow($Batch, [string]$Upn) { @($Batch.rows | Where-Object { $_.upn -ieq $Upn })[0] }
    function Get-JobFile($Work, $Row, [string]$Name) { Join-Path (Join-Path (Join-Path $Work 'jobs') $Row.job) $Name }
    function Get-Given($Work, $Row) { Get-Content -LiteralPath (Get-JobFile $Work $Row 'args.json') -Raw | ConvertFrom-Json }
    function Get-Interval($Work, $Row) {
        $began = [datetime]::Parse((Get-Content -LiteralPath (Get-JobFile $Work $Row 'run-1.txt') -Raw), [cultureinfo]::InvariantCulture, 'RoundtripKind')
        $ended = (Get-Item -LiteralPath (Get-JobFile $Work $Row 'status.json')).LastWriteTimeUtc
        [pscustomobject]@{ Began = $began; Ended = $ended }
    }
    function Get-MaxOverlap($Work, $Rows) {
        $iv = @($Rows | ForEach-Object { Get-Interval $Work $_ })
        $max = 0
        foreach ($i in $iv) { $n = @($iv | Where-Object { $_.Began -le $i.Began.AddMilliseconds(1) -and $_.Ended -ge $i.Began.AddMilliseconds(1) }).Count; if ($n -gt $max) { $max = $n } }
        return $max
    }
}
AfterAll { $env:MAILBOXTRANSFER_LOCAL_CONFIG = $script:SavedLocal }

Describe 'The CSV file of a batch of copies' {
    It 'is read with headers in any case, a semicolon or a comma, with or without a BOM, the defaults of the command and Notes ignored' {
        $csv = New-Csv 'copy-ok.csv' "SOURCEADDRESS;sourcelocation;TargetAddress;TARGETLOCATION;TargetRoot;Notes`r`nalex@contoso.com;archive;adele@contoso.com;primary;From Alex\Archive;first`r`nshared@contoso.com;;adele@contoso.com;Archive;;second`r`n" -Bom
        $r = Read-MbtCopyBatchCsv -Path $csv -DefaultTargetRoot 'Common/' -StoreBase (Join-Path $TestDrive 'stores')
        $r.Problems | Should -BeNullOrEmpty
        $r.Delimiter | Should -Be ';'
        $r.Rows.Count | Should -Be 2
        $r.Rows[0].SourceAddress | Should -Be 'alex@contoso.com'
        $r.Rows[0].SourceLocation | Should -Be 'Archive'
        $r.Rows[0].TargetLocation | Should -Be 'Primary'
        $r.Rows[0].TargetRoot | Should -Be 'From Alex/Archive'
        $r.Rows[1].SourceLocation | Should -Be 'Primary'
        $r.Rows[1].TargetLocation | Should -Be 'Archive'
        $r.Rows[1].TargetRoot | Should -Be 'Common'
        $r.Rows[0].Store | Should -Match ([regex]::Escape((Join-Path $TestDrive 'stores')) + '\\copy-alex@contoso.com-archive-[0-9a-f]{8}$')
        $r.Rows[0].Store | Should -Not -Be $r.Rows[1].Store
    }
    It 'reports every problem at once, with the line of each, and nothing is lost to the first one' {
        $text = "SourceAddress,SourceLocation,TargetAddress,TargetLocation,TargetRoot`r`n" + (@(
                ',Primary,adele@contoso.com,Primary,R',
                'not-an-address,Primary,adele@contoso.com,Primary,R',
                'alex@contoso.com,Both,adele@contoso.com,Primary,R',
                'alex@contoso.com,Primary,,Primary,R',
                'alex@contoso.com,Primary,adele,Primary,R',
                'alex@contoso.com,Primary,adele@contoso.com,Both,R',
                'alex@contoso.com,Primary,adele@contoso.com,Primary,A//B',
                'adele@contoso.com,Primary,ADELE@contoso.com,Primary,R',
                'alex@contoso.com,Primary,adele@contoso.com,Primary,R',
                'ALEX@contoso.com,primary,Adele@Contoso.com,PRIMARY,R2'
            ) -join "`r`n")
        $r = Read-MbtCopyBatchCsv -Path (New-Csv 'copy-bad.csv' $text)
        $r.Rows.Count | Should -Be 1
        $t = $r.Problems -join "`n"
        $t | Should -Match 'line 2: SourceAddress is empty'
        $t | Should -Match "line 3: 'not-an-address' is not a mailbox address"
        $t | Should -Match "line 4: SourceLocation 'Both' is not Primary or Archive \(a copy reads one mailbox location"
        $t | Should -Match 'line 5: TargetAddress is empty'
        $t | Should -Match "line 6: 'adele' is not a mailbox address"
        $t | Should -Match "line 7: TargetLocation 'Both' is not Primary or Archive"
        $t | Should -Match "line 8: TargetRoot 'A//B' has an empty folder name in it"
        $t | Should -Match 'line 9: the source and the target are the same mailbox location \(the Primary mailbox of adele@contoso.com\)'
        $t | Should -Match 'line 11: ALEX@contoso.com \(Primary\) is already copied into Adele@Contoso.com \(Primary\) on line 10'
        $r.Problems.Count | Should -Be 9
    }
    It 'reports a column it does not know, a column twice, a missing required column, an empty file and a missing file' {
        $r = Read-MbtCopyBatchCsv -Path (New-Csv 'copy-col.csv' "SourceAddress,TargetAddress,Target,TargetAddress`r`nalex@contoso.com,adele@contoso.com,R,x`r`n")
        ($r.Problems -join "`n") | Should -Match 'The column Target is not known \(the columns are SourceAddress, SourceLocation, TargetAddress, TargetLocation, TargetRoot, SelectionFile, Notes\)'
        ($r.Problems -join "`n") | Should -Match 'The column TargetAddress is there twice'
        (Read-MbtCopyBatchCsv -Path (New-Csv 'copy-noto.csv' "SourceAddress,TargetLocation`r`nalex@contoso.com,Primary`r`n")).Problems | Should -Match 'SourceAddress and TargetAddress are required \(found: SourceAddress, TargetLocation\)'
        (Read-MbtCopyBatchCsv -Path (New-Csv 'copy-empty.csv' '')).Problems | Should -Match 'The CSV file is empty'
        (Read-MbtCopyBatchCsv -Path (New-Csv 'copy-head.csv' "SourceAddress,TargetAddress`r`n")).Problems | Should -Match 'The CSV file has no copy'
        (Read-MbtCopyBatchCsv -Path (Join-Path $TestDrive 'csv\missing.csv')).Problems | Should -Match 'The CSV file was not found'
    }
    It 'with a delegated sign-in writes into the signed-in account only and reads the archive of that account only; the application sign-in has no such limit' {
        $csv = New-Csv 'copy-dg.csv' "SourceAddress,SourceLocation,TargetAddress`r`nshared@contoso.com,Primary,adele@contoso.com`r`nshared@contoso.com,Archive,adele@contoso.com`r`nalex@contoso.com,Primary,alina@contoso.com`r`nadele@contoso.com,Archive,ADELE@contoso.com`r`n"
        $d = Read-MbtCopyBatchCsv -Path $csv -SignedInUser 'ADELE@contoso.com' -DefaultTargetLocation Primary
        $t = $d.Problems -join "`n"
        $t | Should -Match 'line 3: shared@contoso.com is not the signed-in account \(ADELE@contoso.com\). With a delegated sign-in only its primary mailbox can be read'
        $t | Should -Match 'line 4: alina@contoso.com is not the signed-in account \(ADELE@contoso.com\).*refuses a delegated write into another person.s or a shared mailbox.*-AppOnly'
        $d.Rows.Count | Should -Be 2 -Because 'the primary mailbox of a shared mailbox into the account, and the archive of the account into its primary mailbox'
        $a = Read-MbtCopyBatchCsv -Path $csv -SignedInUser 'adele@contoso.com' -AppOnly
        $a.Problems | Should -BeNullOrEmpty
        $a.Rows.Count | Should -Be 4
    }
    It 'takes one source for two targets and two sources for one target (they will run one after the other), and one store for a source read twice' {
        $csv = New-Csv 'copy-multi.csv' "SourceAddress,TargetAddress,TargetRoot`r`nalex@contoso.com,adele@contoso.com,A`r`nalex@contoso.com,alina@contoso.com,A`r`nshared@contoso.com,adele@contoso.com,B`r`n"
        $r = Read-MbtCopyBatchCsv -Path $csv -StoreBase (Join-Path $TestDrive 'stores')
        $r.Problems | Should -BeNullOrEmpty
        $r.Rows.Count | Should -Be 3
        $r.Rows[0].Store | Should -Be $r.Rows[1].Store -Because 'a source copied into two targets is read once'
        $r.Rows[0].Store | Should -Not -Be $r.Rows[2].Store
    }
    It 'takes a selection of one mailbox location, relative to the folder of the CSV file, refuses one of several, and does not let the column contradict it' {
        $dir = Join-Path $TestDrive 'csv'; $null = New-Item -ItemType Directory -Force -Path $dir
        $one = @'
{ "Version": 1, "Kind": "Archive", "MailboxId": "MBX:11111111-2222-3333-4444-5555555555aa@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d", "Entries": [ { "FolderId": "F1", "FolderPath": "Inbox", "IncludeSubfolders": true } ] }
'@
        [IO.File]::WriteAllText((Join-Path $dir 'sel-one.json'), $one)
        $two = "{ `"Version`": 1, `"Selections`": [ $($one.Trim()), $($one.Trim() -replace 'Archive', 'Primary' -replace '5555555555aa', '555555555555') ] }"
        [IO.File]::WriteAllText((Join-Path $dir 'sel-two.json'), $two)
        $r = Read-MbtCopyBatchCsv -Path (New-Csv 'copy-sel.csv' "SourceAddress,TargetAddress,SelectionFile,SourceLocation`r`nalex@contoso.com,adele@contoso.com,sel-one.json,`r`nalex@contoso.com,alina@contoso.com,sel-two.json,`r`nalex@contoso.com,anna@contoso.com,sel-one.json,Primary`r`nalex@contoso.com,ana@contoso.com,nowhere.json,`r`n")
        $r.Rows.Count | Should -Be 1
        $r.Rows[0].SourceLocation | Should -Be 'Archive' -Because 'the file says which location is read'
        $r.Rows[0].SelectionFile | Should -Be (Join-Path $dir 'sel-one.json')
        $t = $r.Problems -join "`n"
        $t | Should -Match 'line 3: the SelectionFile names 2 mailbox locations: a copy takes one'
        $t | Should -Match 'line 4: SourceLocation Primary does not match the SelectionFile, which is of the Archive mailbox'
        $t | Should -Match 'line 5: the SelectionFile was not found'
    }
}

Describe 'The arguments of a copy' {
    BeforeAll {
        function New-CopyRow([string]$Source = 'alex@contoso.com', [string]$SourceLocation = 'Primary', [string]$Target = 'adele@contoso.com', [string]$TargetLocation = 'Primary', [string]$Root = 'Copied', [string]$Selection = '') {
            [ordered]@{ upn = $Target; location = $TargetLocation; sourceUpn = $Source; sourceLocation = $SourceLocation; targetRoot = $Root; selection = $Selection; store = 'D:\w\stores\s1' }
        }
        $script:Job = [pscustomobject]@{ Id = 'mailcopy-ab12'; Folder = 'D:\w\jobs\x-mailcopy-ab12'; EventsPath = 'D:\w\jobs\x-mailcopy-ab12\events.jsonl'; StopPath = 'D:\w\jobs\x-mailcopy-ab12\stop.request' }
    }
    It 'the application names both mailboxes: the source as the mailbox it works on, the target by -TargetMailboxId; the store, the folder and the options of the batch follow' {
        $a = @(Get-MbtBatchChildArgument -Row (New-CopyRow -SourceLocation Archive -TargetLocation Archive) -Option @{ AppOnly = $true; SignInUser = ''; Pass = @{ Concurrency = 2; Force = $true; ArrivedFrom = '2026-01-01'; NoAutoReconcile = $false } } -Job $script:Job -BatchId 'B1' -Mode Copy)
        ($a -join ' ') | Should -Match '^-Mode Copy -UserPrincipalName alex@contoso.com -TargetMailboxId adele@contoso.com -Mailbox Archive -TargetMailbox Archive -TargetRoot Copied -StorePath D:\\w\\stores\\s1 -TempPath '
        $a | Should -Contain '-Force'
        $a | Should -Not -Contain '-NoAutoReconcile'
        $a[($a.IndexOf('-ArrivedFrom') + 1)] | Should -Be '2026-01-01'
        $a[($a.IndexOf('-BatchId') + 1)] | Should -Be 'B1'
        $a[($a.IndexOf('-JobId') + 1)] | Should -Be 'mailcopy-ab12'
        $a | Should -Contain '-NoPrompt'
    }
    It 'delegated: the signed-in account is the one that signs in and receives; another source is given by its address, a selection names its own mailbox' {
        $d = @(Get-MbtBatchChildArgument -Row (New-CopyRow 'shared@contoso.com' -Target 'ADELE@contoso.com') -Option @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = @{} } -Job $script:Job -BatchId 'B1' -Mode Copy)
        ($d -join ' ') | Should -Match '^-Mode Copy -UserPrincipalName adele@contoso.com -MailboxId shared@contoso.com -Mailbox Primary -TargetMailbox Primary '
        $d | Should -Not -Contain '-TargetMailboxId'
        $own = @(Get-MbtBatchChildArgument -Row (New-CopyRow 'adele@contoso.com' 'Archive' 'adele@contoso.com' 'Primary') -Option @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = @{} } -Job $script:Job -BatchId 'B1' -Mode Copy)
        $own | Should -Not -Contain '-MailboxId'
        $s = @(Get-MbtBatchChildArgument -Row (New-CopyRow 'shared@contoso.com' -Selection 'D:\sel.json') -Option @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = @{} } -Job $script:Job -BatchId 'B1' -Mode Copy)
        $s[($s.IndexOf('-SelectionFile') + 1)] | Should -Be 'D:\sel.json'
        $s | Should -Not -Contain '-MailboxId'
        $s | Should -Not -Contain '-Mailbox' -Because 'the selection says which location'
    }
    It 'a delegated sign-in never gets a line for another target: the arguments are refused before anything is read' {
        { Get-MbtBatchChildArgument -Row (New-CopyRow -Target 'shared@contoso.com') -Option @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = @{} } -Job $script:Job -BatchId 'B1' -Mode Copy } | Should -Throw '*delegated sign-in writes into the mailbox of the signed-in account*'
        { Get-MbtBatchChildArgument -Row (New-CopyRow -Target 'shared@contoso.com') -Option $script:AppOption -Job $script:Job -BatchId 'B1' -Mode Copy } | Should -Not -Throw
    }
    It 'every parameter it uses exists in the entry script' {
        $declared = (Get-Command $script:Entry).Parameters.Keys
        $pass = @{ Concurrency = 2; MaxAttempts = 3; RetryWaitSeconds = 2; IntervalSeconds = 5; ConfigPath = 'x.psd1'; ReconcileDelaySeconds = 60; ReconcileRounds = 2; Folder = 'Inbox'; MappingFile = 'D:\m.csv'; NotReadyTrip = 3; BatchItems = 10; BatchMiB = 100
            NoAutoReconcile = $true; SkipInventory = $true; IncludeRecoverableItems = $true; Force = $true; TenantId = 't'; ClientId = 'c'; CertificateThumbprint = 'ab'; ArrivedFrom = '2026-01-01'; ArrivedTo = '2026-02-01' }
        foreach ($option in @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = $pass }, @{ AppOnly = $true; SignInUser = ''; Pass = $pass }) {
            $a = Get-MbtBatchChildArgument -Row (New-CopyRow) -Option $option -Job $script:Job -BatchId 'B' -Mode Copy
            foreach ($t in $a) { if ($t -match '^-([A-Za-z]+)$') { $declared | Should -Contain $Matches[1] -Because "$t" } }
        }
    }
    It 'the request of the console for the Continue says a copy, the source, the target and the folder' {
        $w = New-Work 'req'
        $module = Get-Module MailboxTransfer
        $job = & $module { param($folder) New-MbtBatchJobFolder -WorkFolder $folder -Kind mailcopy } $w
        $job.Id | Should -Match '^mailcopy-[0-9a-f]{4}$'
        $job.Name | Should -Match '^\d{8}-\d{4}-mailcopy-[0-9a-f]{4}$'
        (& $module { param($folder, $name) New-MbtBatchJobFolder -WorkFolder $folder -Name $name -Kind mailcopy } $w $job.Name).Id | Should -Be $job.Id
        $row = New-CopyRow -Root 'Copied/x'
        $argument = @(Get-MbtBatchChildArgument -Row $row -Option $script:AppOption -Job $job -BatchId 'B' -Mode Copy)
        & $module { param($r, $j, $a, $o) Write-MbtBatchJobRequest -Row $r -Job $j -Argument $a -Option $o -BatchId 'B' -Mode Copy } $row $job $argument @{ AppOnly = $true; SignInUser = ''; Pass = @{ Concurrency = 3 } }
        $q = Get-Content -LiteralPath (Join-Path $job.Folder 'request.json') -Raw | ConvertFrom-Json
        $q.Task | Should -Be 'mailcopy'
        $q.Mailbox | Should -Be 'alex@contoso.com'
        $q.McTarget | Should -Be 'adele@contoso.com'
        $q.TgtLoc | Should -Be 'Primary'
        $q.TgtMode | Should -Be 'new'
        $q.TgtName | Should -Be 'Copied/x'
        $q.Conc | Should -Be 3
        $q.Batch | Should -Be 'B'
        $q.StorePath | Should -Be 'D:\w\stores\s1'
        @($q.Commands[0]) | Should -Contain '-TargetMailboxId'
    }
    It 'a line holds its target location and its source location: the locks say it' {
        $row = New-CopyRow -Source 'Alex@contoso.com' -SourceLocation Archive -Target 'ADELE@contoso.com' -TargetLocation Archive
        $locks = @(Get-MbtBatchRowLock $row)
        $locks | Should -Contain 'target:adele@contoso.com|archive'
        $locks | Should -Contain 'source:alex@contoso.com|archive'
        (@(Get-MbtBatchRowLock (New-CopyRow -Target 'adele@contoso.com' -TargetLocation Primary)) -contains 'target:adele@contoso.com|archive') | Should -BeFalse
    }
}

Describe 'Running a batch of copies on stand-in child processes' {
    It 'runs lines of different targets and sources side by side, never two lines of the same target location or the same source, whatever the number asked at once' {
        $w = New-Work 'serial'
        # two lines into ok-pause-a, two from ok-src-b; line 5 has nothing in common with the others
        $b = New-TestCopyBatch $w 'Serial' @('ok-src-a@contoso.com,ok-pause-a@contoso.com', 'ok-src-b@contoso.com,ok-pause-a@contoso.com', 'ok-src-b@contoso.com,ok-pause-b@contoso.com', 'ok-src-c@contoso.com,ok-pause-c@contoso.com') -Parallel 3
        $b.mode | Should -Be 'Copy'
        $r = Invoke-StubBatch $w $b -Parallel 3
        $r.ExitCode | Should -Be 0
        $r.Counts.Complete | Should -Be 4
        Get-MaxOverlap $w @($r.Batch.rows | Where-Object { $_.upn -ieq 'ok-pause-a@contoso.com' }) | Should -Be 1 -Because 'two lines into one target location ran together'
        Get-MaxOverlap $w @($r.Batch.rows | Where-Object { $_.sourceUpn -ieq 'ok-src-b@contoso.com' }) | Should -Be 1 -Because 'one source was read twice at once'
        Get-MaxOverlap $w @($r.Batch.rows) | Should -BeGreaterThan 1 -Because 'the others run side by side'
        Get-MaxOverlap $w @($r.Batch.rows) | Should -BeLessOrEqual 3
    }
    It 'the Archive and the Primary mailbox of one address are two target locations: they run together' {
        $w = New-Work 'serial-loc'
        $csv = New-Csv 'loc.csv' "SourceAddress,TargetAddress,TargetLocation`r`nok-src-a@contoso.com,ok-pause-a@contoso.com,Primary`r`nok-src-b@contoso.com,ok-pause-a@contoso.com,Archive`r`n"
        $read = Read-MbtCopyBatchCsv -Path $csv -DefaultTargetRoot R -StoreBase (Join-Path $w 'stores')
        $b = New-MbtBatch -Id 'Loc' -Csv $csv -Row $read.Rows -Option $script:AppOption -Parallel 2 -Mode Copy
        $r = Invoke-StubBatch $w $b -Parallel 2
        Get-MaxOverlap $w @($r.Batch.rows) | Should -Be 2
    }
    It 'every line is an ordinary copy job in its own folder with the batch in its status, the arguments of a copy and the counts of both halves' {
        $w = New-Work 'jobs'
        $b = New-TestCopyBatch $w 'Jobs' @('ok-src-1@contoso.com,ok-1@contoso.com', 'ok-src-2@contoso.com,ok-2@contoso.com') -Root 'Restored'
        $r = Invoke-StubBatch $w $b
        $r.ExitCode | Should -Be 0
        foreach ($row in $r.Batch.rows) {
            $row.state | Should -Be 'complete'
            $row.done | Should -Be 100
            $row.exported | Should -Be 100
            $status = Get-Content -LiteralPath (Join-Path (Join-Path (Join-Path $w 'jobs') $row.job) 'status.json') -Raw | ConvertFrom-Json
            $status.batch | Should -Be 'Jobs'
            $status.task | Should -Be 'mailcopy'
            $status.mode | Should -Be 'Copy'
            $given = Get-Given $w $row
            $given[($given.IndexOf('-Mode') + 1)] | Should -Be 'Copy'
            $given[($given.IndexOf('-StorePath') + 1)] | Should -Be $row.store
            $given[($given.IndexOf('-TargetRoot') + 1)] | Should -Be 'Restored'
            $given[($given.IndexOf('-TargetMailboxId') + 1)] | Should -Be $row.upn
            $given[($given.IndexOf('-JobId') + 1)] | Should -Match '^mailcopy-[0-9a-f]{4}$'
            Test-Path -LiteralPath (Join-Path (Join-Path (Join-Path $w 'jobs') $row.job) 'request.json') | Should -BeTrue
        }
        Test-Path -LiteralPath (Join-Path $w 'batches\Jobs\batch.json') | Should -BeTrue
        $html = Get-Content -LiteralPath (Join-Path $w 'batches\Jobs\report.html') -Raw
        $html | Should -Match 'ok-1@contoso.com'
    }
    It 'continues what is not complete and skips what is: the same batch again; a changed list is refused' {
        $w = New-Work 'resume'
        $b = New-TestCopyBatch $w 'Resume' @('ok-src-1@contoso.com,ok-1@contoso.com', 'ok-src-2@contoso.com,partial-2@contoso.com', 'ok-src-3@contoso.com,uncertain-3@contoso.com')
        $first = Invoke-StubBatch $w $b
        $first.ExitCode | Should -Be 2
        (Get-BatchRow $first.Batch 'ok-1@contoso.com').state | Should -Be 'complete'
        (Get-BatchRow $first.Batch 'partial-2@contoso.com').state | Should -Be 'partial'
        (Get-BatchRow $first.Batch 'uncertain-3@contoso.com').state | Should -Be 'partial'
        (Get-BatchRow $first.Batch 'uncertain-3@contoso.com').message | Should -Match 'uncertain items'
        $second = Invoke-StubBatch $w (Read-MbtBatch -Folder (Get-MbtBatchFolder -WorkFolder $w -Name 'Resume'))
        $second.ExitCode | Should -Be 0
        (Get-BatchRow $second.Batch 'ok-1@contoso.com').attempts | Should -Be 1 -Because 'a complete line is not run again'
        (Get-BatchRow $second.Batch 'partial-2@contoso.com').attempts | Should -Be 2
        (Get-BatchRow $second.Batch 'partial-2@contoso.com').job | Should -Be (Get-BatchRow $first.Batch 'partial-2@contoso.com').job -Because 'the same job folder: only what is left'
        $other = Read-MbtCopyBatchCsv -Path (New-Csv 'resume-other.csv' "SourceAddress,TargetAddress`r`nok-src-1@contoso.com,ok-1@contoso.com`r`nok-src-9@contoso.com,ok-9@contoso.com`r`n") -DefaultTargetRoot Copied
        $diff = @(Compare-MbtCopyBatchRow -Batch $second.Batch -Row $other.Rows)
        ($diff -join "`n") | Should -Match 'ok-src-9@contoso.com \(Primary\) into ok-9@contoso.com \(Primary\) folder Copied is in the CSV file and not in the batch'
        ($diff -join "`n") | Should -Match 'is in the batch and not in the CSV file'
        @(Compare-MbtCopyBatchRow -Batch $second.Batch -Row (Read-MbtCopyBatchCsv -Path (Join-Path $TestDrive 'csv\Resume.csv') -DefaultTargetRoot Copied).Rows).Count | Should -Be 0
    }
    It 'a line that failed is reported with its reason and does not stop the others; exit 1' {
        $w = New-Work 'fail'
        $b = New-TestCopyBatch $w 'Fail' @('ok-src-1@contoso.com,fail-1@contoso.com', 'ok-src-2@contoso.com,ok-2@contoso.com')
        $r = Invoke-StubBatch $w $b
        $r.ExitCode | Should -Be 1
        (Get-BatchRow $r.Batch 'fail-1@contoso.com').state | Should -Be 'failed'
        (Get-BatchRow $r.Batch 'fail-1@contoso.com').message | Should -Match 'We cannot find this mailbox'
        (Get-BatchRow $r.Batch 'ok-2@contoso.com').state | Should -Be 'complete'
    }
    It 'a stop request stops the running lines, starts no other, and the batch can be run again' {
        $w = New-Work 'stop'
        $b = New-TestCopyBatch $w 'Stop' @('ok-src-1@contoso.com,slow-1@contoso.com', 'ok-src-2@contoso.com,ok-2@contoso.com')
        $folder = Get-MbtBatchFolder -WorkFolder $w -Name 'Stop'
        $null = New-Item -ItemType Directory -Force -Path $folder
        $job = Start-ThreadJob -ArgumentList $w, $folder -ScriptBlock {
            param($work, $batchFolder)
            $until = [DateTime]::UtcNow.AddSeconds(60)
            while ([DateTime]::UtcNow -lt $until) {
                if (@(Get-ChildItem -LiteralPath (Join-Path $work 'jobs') -Filter 'run-1.txt' -Recurse -ErrorAction SilentlyContinue).Count) { [IO.File]::WriteAllText((Join-Path $batchFolder 'stop.request'), 'stop'); return }
                Start-Sleep -Milliseconds 100
            }
        }
        try { $r = Invoke-StubBatch $w $b } finally { Remove-Job $job -Force }
        $r.Stopped | Should -BeTrue
        $r.ExitCode | Should -Be 2
        (Get-BatchRow $r.Batch 'slow-1@contoso.com').state | Should -Be 'partial'
        (Get-BatchRow $r.Batch 'ok-2@contoso.com').state | Should -Be 'pending' -Because 'no other line starts after the stop'
    }
}

Describe 'The command line of a batch of copies' {
    BeforeAll {
        function Invoke-EntryChild {
            param([string[]]$Argument)
            $out = & (Get-Process -Id $PID).Path -NoProfile -File $script:Entry -DevelopmentMode @Argument 2>&1 | Out-String
            [pscustomobject]@{ Text = $out; Exit = $LASTEXITCODE }
        }
    }
    It '-WhatIf lists every line and writes nothing; the file is checked first and nothing starts while a problem remains' {
        $w = New-Work 'cli'
        $csv = New-Csv 'cli-ok.csv' "SourceAddress,TargetAddress,TargetRoot`r`nalex@contoso.com,adele@contoso.com,From Alex`r`nshared@contoso.com,adele@contoso.com,From shared`r`n"
        $r = Invoke-EntryChild @('-Mode', 'Copy', '-MailboxCsv', $csv, '-BatchName', 'CliPlan', '-WorkFolder', $w, '-AppOnly', '-CertificateThumbprint', ('AB' * 20), '-ClientId', '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d', '-TenantId', '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d', '-WhatIf')
        $r.Exit | Should -Be 0
        $r.Text | Should -Match 'Copy batch planned'
        $r.Text | Should -Match 'alex@contoso.com \(Primary\) into adele@contoso.com \(Primary\)'
        $r.Text | Should -Match 'target mailbox location is written by several lines'
        Test-Path -LiteralPath (Join-Path $w 'jobs') | Should -BeFalse
        $bad = New-Csv 'cli-bad.csv' "SourceAddress,TargetAddress`r`nalex@contoso.com,alex@contoso.com`r`n,adele@contoso.com`r`n"
        $b = Invoke-EntryChild @('-Mode', 'Copy', '-MailboxCsv', $bad, '-BatchName', 'CliBad', '-WorkFolder', $w, '-AppOnly', '-CertificateThumbprint', ('AB' * 20), '-ClientId', 'x', '-TenantId', 'y')
        $b.Exit | Should -Be 1
        $b.Text | Should -Match 'The batch cannot start'
        $b.Text | Should -Match 'line 2: the source and the target are the same mailbox location'
        $b.Text | Should -Match 'line 3: SourceAddress is empty'
        $b.Text | Should -Match '2 problems, listed above: nothing was started'
        Test-Path -LiteralPath (Join-Path $w 'batches\CliBad') | Should -BeFalse
    }
    It 'names a batch for the mode it was made for, and refuses the single-copy parameters' {
        $w = New-Work 'cli-mode'
        $csv = New-Csv 'cli-mode.csv' "SourceAddress,TargetAddress`r`nalex@contoso.com,adele@contoso.com`r`n"
        $null = Invoke-EntryChild @('-Mode', 'Copy', '-MailboxCsv', $csv, '-BatchName', 'Kind', '-WorkFolder', $w, '-AppOnly', '-CertificateThumbprint', ('AB' * 20), '-ClientId', 'x', '-TenantId', 'y', '-WhatIf')
        $r = Invoke-EntryChild @('-Mode', 'Export', '-ResumeBatch', 'Kind', '-WorkFolder', $w)
        $r.Text | Should -Match 'There is no batch named Kind|was made for -Mode Copy'
        foreach ($single in @('-TargetMailboxId', 'adele@contoso.com'), @('-SelectionFile', 'x.json'), @('-MailboxId', 'alex@contoso.com')) {
            $x = Invoke-EntryChild (@('-Mode', 'Copy', '-MailboxCsv', $csv, '-WorkFolder', $w) + $single)
            $x.Text | Should -Match "$([regex]::Escape($single[0])) is for one mailbox"
        }
        (Invoke-EntryChild @('-Mode', 'Copy', '-MailboxCsv', $csv, '-WorkFolder', $w, '-Mailbox', 'All')).Text | Should -Match 'is not for a batch'
    }
    It 'a delegated sign-in is checked against the file: a target that is not the signed-in account stops the batch before anything starts' {
        $w = New-Work 'cli-deleg'
        $csv = New-Csv 'cli-deleg.csv' "SourceAddress,TargetAddress`r`nalex@contoso.com,alina@contoso.com`r`n"
        $r = Invoke-EntryChild @('-Mode', 'Copy', '-MailboxCsv', $csv, '-BatchName', 'Deleg', '-WorkFolder', $w, '-UserPrincipalName', 'adele@contoso.com', '-TenantId', '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d', '-ClientId', '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d', '-WhatIf')
        $r.Exit | Should -Be 1
        $r.Text | Should -Match 'alina@contoso.com is not the signed-in account \(adele@contoso.com\)'
    }
}
