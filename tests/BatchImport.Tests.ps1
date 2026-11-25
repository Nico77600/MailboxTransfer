#Requires -Version 7.6
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the import of a list of mailboxes (module\MailboxTransfer.Batch.ps1 and the entry script with -Mode Import, ImportPst or ImportEml and -MailboxCsv): the checks of the CSV
    file (every problem at once), the arguments of an import, the scheduler on stand-in child processes (tests\support\batch-child.ps1: lines of one target or one source never run together, different
    targets run side by side, resume, sign-in, exit codes), the command line (-WhatIf, the error card, the mode of a batch) and one batch of real EML imports run against the fake service.
    Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\BatchImport.Tests.ps1 -Output Detailed
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
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:Src = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
    $script:AppOption = @{ AppOnly = $true; SignInUser = ''; Pass = @{} }

    function New-Csv([string]$Name, [string]$Text, [switch]$Bom) {
        $dir = Join-Path $TestDrive 'csv'
        $null = New-Item -ItemType Directory -Force -Path $dir
        $path = Join-Path $dir $Name
        [IO.File]::WriteAllText($path, $Text, [Text.UTF8Encoding]::new([bool]$Bom))
        return $path
    }
    function New-Work([string]$Name) { $w = Join-Path $TestDrive $Name; $null = New-Item -ItemType Directory -Force -Path $w; return $w }
    # A source of the kind the mode takes, below the folder of the CSV files (so that a relative Source is tried too); returns its path.
    function New-Source([string]$Mode, [string]$Name) {
        $base = Join-Path $TestDrive 'csv'
        $null = New-Item -ItemType Directory -Force -Path $base
        switch ($Mode) {
            'Import' { $p = Join-Path $base $Name; $null = New-Item -ItemType Directory -Force -Path $p; [IO.File]::WriteAllBytes((Join-Path $p 'catalog.db'), [byte[]](1, 2, 3)) }
            'ImportPst' { $p = Join-Path $base "$Name.pst"; [IO.File]::WriteAllBytes($p, [byte[]](1, 2, 3)) }
            'ImportEml' {
                $p = Join-Path $base $Name; $null = New-Item -ItemType Directory -Force -Path (Join-Path $p 'Sub')
                [IO.File]::WriteAllText((Join-Path $p 'Sub\one.eml'), "From: a@contoso.com`r`nTo: b@contoso.com`r`nSubject: One $Name`r`nDate: Mon, 18 Mar 2024 10:00:00 +0000`r`n`r`nBody`r`n")
            }
        }
        return $p
    }
    # The batch of a list of imports, rows as the CSV reader makes them: pairs of "mailbox" and "source name".
    function New-TestImportBatch {
        param([string]$Work, [string]$Id, [string[]]$Mailbox, [string[]]$Source = @(), [string]$Mode = 'ImportPst', [hashtable]$Option = $script:AppOption, [int]$Parallel = 1, [string]$Root = 'Imported')
        $lines = for ($i = 0; $i -lt $Mailbox.Count; $i++) {
            $name = if ($i -lt $Source.Count) { $Source[$i] } else { "s$i" }
            $null = New-Source $Mode $name
            $path = Join-Path (Join-Path $TestDrive 'csv') $(if ($Mode -eq 'ImportPst') { "$name.pst" } else { $name })
            "$($Mailbox[$i]),$path"
        }
        $csv = New-Csv "$Id.csv" ("Mailbox,Source`r`n" + ($lines -join "`r`n"))
        $read = Read-MbtImportBatchCsv -Path $csv -Mode $Mode -DefaultTargetRoot $Root -WorkFolder $Work
        $read.Problems | Should -BeNullOrEmpty
        return (New-MbtBatch -Id $Id -Csv $csv -Row $read.Rows -Option $Option -Parallel $Parallel -Mode $Mode)
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

Describe 'The CSV file of an import batch' {
    It 'is read with headers in any case, a semicolon or a comma, with or without a BOM, a relative Source, the defaults of the command and Notes ignored' {
        $null = New-Source 'ImportPst' 'rel-a'; $null = New-Source 'ImportPst' 'rel-b'
        $csv = New-Csv 'imp-ok.csv' "MAILBOX;SOURCE;location;TARGETROOT;Notes`r`nadele@contoso.com;rel-a.pst;archive;Imported\A;first`r`nalex@contoso.com;rel-b.pst;;;second`r`n" -Bom
        $r = Read-MbtImportBatchCsv -Path $csv -Mode ImportPst -DefaultTargetRoot 'Common/' -WorkFolder (Join-Path $TestDrive 'w1')
        $r.Problems | Should -BeNullOrEmpty
        $r.Delimiter | Should -Be ';'
        $r.Rows.Count | Should -Be 2
        $r.Rows[0].Mailbox | Should -Be 'adele@contoso.com'
        $r.Rows[0].Location | Should -Be 'Archive'
        $r.Rows[0].TargetRoot | Should -Be 'Imported/A'
        $r.Rows[0].Source | Should -Be (Join-Path (Join-Path $TestDrive 'csv') 'rel-a.pst')
        $r.Rows[1].Location | Should -Be 'Primary'
        $r.Rows[1].TargetRoot | Should -Be 'Common'
        $r.Rows[0].Journal | Should -Be (Join-Path (Join-Path (Join-Path $TestDrive 'w1') 'imports') $r.Rows[0].Key)
        $r.Rows[0].Key | Should -Match '^[0-9a-f]{12}$'
        $r.Rows[0].Key | Should -Not -Be $r.Rows[1].Key
    }
    It 'gives a row the same key whatever the case of the address, the path, the Location and the folder: the journal and the job name do not depend on it' {
        $null = New-Source 'ImportPst' 'case-a'
        $one = New-Csv 'case1.csv' "Mailbox,Source,Location,TargetRoot`r`nAdele@Contoso.com,case-a.pst,Archive,Imported/A`r`n"
        $two = New-Csv 'case2.csv' "mailbox,source,location,targetroot`r`nadele@contoso.com,CASE-A.PST,archive,imported\a`r`n"
        $a = (Read-MbtImportBatchCsv -Path $one -Mode ImportPst).Rows[0]
        $b = (Read-MbtImportBatchCsv -Path $two -Mode ImportPst).Rows[0]
        $a.Key | Should -Be $b.Key
        (Get-MbtImportBatchKey -Mailbox 'a@x.fr' -Location 'Primary' -Source 'D:\S' -TargetRoot 'R' -SourceMailbox '') | Should -Be (Get-MbtImportBatchKey -Mailbox 'A@X.FR' -Location 'primary' -Source 'd:\s' -TargetRoot 'r' -SourceMailbox '')
        (Get-MbtImportBatchKey -Mailbox 'a@x.fr' -Location 'Primary' -Source 'D:\S' -TargetRoot 'R' -SourceMailbox '') | Should -Not -Be (Get-MbtImportBatchKey -Mailbox 'a@x.fr' -Location 'Archive' -Source 'D:\S' -TargetRoot 'R' -SourceMailbox '')
    }
    It 'reports every problem at once, with the line of each, and nothing is lost to the first one' {
        $ok = New-Source 'ImportPst' 'prob-ok'
        $store = New-Source 'Import' 'prob-store'
        $emptyStore = Join-Path $TestDrive 'csv\prob-empty-store'; $null = New-Item -ItemType Directory -Force -Path $emptyStore
        $text = Join-Path $TestDrive 'csv\prob.txt'; [IO.File]::WriteAllText($text, 'x')
        $csv = New-Csv 'imp-bad.csv' ("Mailbox,Source,Location,TargetRoot,SourceMailbox`r`n" + (@(
                    ',prob-ok.pst,Primary,R,',
                    'not-an-address,prob-ok.pst,Primary,R,',
                    "adele@contoso.com,prob-ok.pst,Both,R,",
                    'adele@contoso.com,,Primary,R,',
                    'adele@contoso.com,nowhere.pst,Primary,R,',
                    'adele@contoso.com,prob.txt,Primary,R,',
                    'adele@contoso.com,prob-ok.pst,Primary,,',
                    'adele@contoso.com,prob-ok.pst,Primary,A//B,',
                    'adele@contoso.com,prob-ok.pst,Primary,R,Primary',
                    'alex@contoso.com,prob-ok.pst,Primary,R,',
                    'ALEX@contoso.com,PROB-OK.PST,primary,R2,'
                ) -join "`r`n"))
        $r = Read-MbtImportBatchCsv -Path $csv -Mode ImportPst
        $r.Rows.Count | Should -Be 1
        $text = $r.Problems -join "`n"
        $text | Should -Match 'line 2: Mailbox is empty'
        $text | Should -Match "line 3: 'not-an-address' is not a mailbox address"
        $text | Should -Match "line 4: Location 'Both' is not Primary or Archive"
        $text | Should -Match 'line 5: Source is empty'
        $text | Should -Match 'line 6: the Source was not found: .*nowhere\.pst'
        $text | Should -Match 'line 7: the Source is not a \.pst file: .*prob\.txt'
        $text | Should -Match 'line 8: no TargetRoot: give the column, or -TargetRoot, or -MergeWellKnown'
        $text | Should -Match "line 9: TargetRoot 'A//B' has an empty folder name in it"
        $text | Should -Match 'line 10: SourceMailbox applies to an export store'
        $text | Should -Match 'line 12: ALEX@contoso.com \(Primary\) already imports .* on line 11'
        $r.Problems.Count | Should -Be 10
        # the same with a store and with .eml files
        $s = Read-MbtImportBatchCsv -Path (New-Csv 'imp-bad-store.csv' "Mailbox,Source,SourceMailbox`r`nadele@contoso.com,prob-store,Both`r`nadele@contoso.com,prob-empty-store,`r`nadele@contoso.com,prob.txt,`r`nadele@contoso.com,prob-store,Archive`r`n") -Mode Import
        ($s.Problems -join "`n") | Should -Match "line 2: SourceMailbox 'Both' is not Primary or Archive"
        ($s.Problems -join "`n") | Should -Match 'line 3: the Source is not an export store \(no catalog.db in it\)'
        ($s.Problems -join "`n") | Should -Match 'line 4: the Source is a file: an export store is a folder'
        $s.Rows.Count | Should -Be 1 -Because 'the same store for the Archive of the store is another import'
        $e = Read-MbtImportBatchCsv -Path (New-Csv 'imp-bad-eml.csv' "Mailbox,Source`r`nadele@contoso.com,prob-store`r`nadele@contoso.com,prob-ok.pst`r`n") -Mode ImportEml
        ($e.Problems -join "`n") | Should -Match 'line 2: the Source folder holds no \.eml file'
        ($e.Problems -join "`n") | Should -Match 'line 3: the Source is not an \.eml file or a folder of \.eml files'
        ($e.Problems -join "`n") | Should -Match 'line 2: no TargetRoot: give the column, or -TargetRoot'
    }
    It 'reports a column it does not know, a column twice, a missing required column, an empty file and a missing file' {
        $null = New-Source 'ImportPst' 'col-a'
        $r = Read-MbtImportBatchCsv -Path (New-Csv 'imp-col.csv' "Mailbox,Source,Target,Mailbox`r`nadele@contoso.com,col-a.pst,R,x`r`n") -Mode ImportPst -DefaultTargetRoot R
        ($r.Problems -join "`n") | Should -Match 'The column Target is not known \(the columns are Mailbox, Source, Location, TargetRoot, SourceMailbox, Notes\)'
        ($r.Problems -join "`n") | Should -Match 'The column Mailbox is there twice'
        $none = Read-MbtImportBatchCsv -Path (New-Csv 'imp-nosrc.csv' "Mailbox,Location`r`nadele@contoso.com,Primary`r`n") -Mode ImportPst
        ($none.Problems -join "`n") | Should -Match 'Mailbox and Source are required \(found: Mailbox, Location\)'
        (Read-MbtImportBatchCsv -Path (New-Csv 'imp-empty.csv' '') -Mode ImportPst).Problems | Should -Match 'The CSV file is empty'
        (Read-MbtImportBatchCsv -Path (New-Csv 'imp-head.csv' "Mailbox,Source`r`n") -Mode ImportPst).Problems | Should -Match 'The CSV file has no import'
        (Read-MbtImportBatchCsv -Path (Join-Path $TestDrive 'csv\missing.csv') -Mode ImportPst).Problems | Should -Match 'The CSV file was not found'
    }
    It 'with a delegated sign-in only the mailbox of the signed-in account can be written to; the application sign-in has no such limit' {
        $null = New-Source 'ImportPst' 'dg-a'; $null = New-Source 'ImportPst' 'dg-b'
        $csv = New-Csv 'imp-dg.csv' "Mailbox,Source,TargetRoot`r`nadele@contoso.com,dg-a.pst,R`r`nshared@contoso.com,dg-b.pst,R`r`n"
        $d = Read-MbtImportBatchCsv -Path $csv -Mode ImportPst -SignedInUser 'ADELE@contoso.com'
        $d.Rows.Count | Should -Be 1
        $d.Problems.Count | Should -Be 1
        $d.Problems[0] | Should -Match 'line 3: shared@contoso.com is not the signed-in account \(ADELE@contoso.com\).*refuses a delegated write into another person.s or a shared mailbox.*-AppOnly'
        $a = Read-MbtImportBatchCsv -Path $csv -Mode ImportPst -SignedInUser 'adele@contoso.com' -AppOnly
        $a.Problems | Should -BeNullOrEmpty
        $a.Rows.Count | Should -Be 2
    }
    It 'takes several sources for one target (they will run one after the other) and the Primary and the Archive of one store for one target' {
        $null = New-Source 'ImportPst' 'multi-a'; $null = New-Source 'ImportPst' 'multi-b'; $null = New-Source 'Import' 'multi-store'
        $r = Read-MbtImportBatchCsv -Path (New-Csv 'imp-multi.csv' "Mailbox,Source,TargetRoot`r`nadele@contoso.com,multi-a.pst,A`r`nadele@contoso.com,multi-b.pst,B`r`n") -Mode ImportPst
        $r.Problems | Should -BeNullOrEmpty
        $r.Rows.Count | Should -Be 2
        $s = Read-MbtImportBatchCsv -Path (New-Csv 'imp-multi-store.csv' "Mailbox,Source,SourceMailbox,TargetRoot`r`nadele@contoso.com,multi-store,Primary,P`r`nadele@contoso.com,multi-store,archive,A`r`n") -Mode Import
        $s.Problems | Should -BeNullOrEmpty
        $s.Rows[1].SourceMailbox | Should -Be 'Archive'
        $s.Rows[0].Key | Should -Not -Be $s.Rows[1].Key
        # a store may keep the folder names: no TargetRoot is fine; the global SourceMailbox is the default of a row that has none
        $g = Read-MbtImportBatchCsv -Path (New-Csv 'imp-store-same.csv' "Mailbox,Source`r`nadele@contoso.com,multi-store`r`n") -Mode Import -DefaultSourceMailbox archive
        $g.Problems | Should -BeNullOrEmpty
        $g.Rows[0].SourceMailbox | Should -Be 'Archive'
        $g.Rows[0].TargetRoot | Should -Be ''
        $m = Read-MbtImportBatchCsv -Path (New-Csv 'imp-pst-merge.csv' "Mailbox,Source`r`nadele@contoso.com,multi-a.pst`r`n") -Mode ImportPst -MergeWellKnown
        $m.Problems | Should -BeNullOrEmpty
    }
    It 'accepts a folder of .pst files and one .eml file, relative to the folder of the CSV file' {
        $folder = Join-Path $TestDrive 'csv\pstfolder'; $null = New-Item -ItemType Directory -Force -Path $folder
        [IO.File]::WriteAllBytes((Join-Path $folder 'mail.pst'), [byte[]](1)); [IO.File]::WriteAllBytes((Join-Path $folder 'mail-0002.pst'), [byte[]](1))
        $one = Join-Path $TestDrive 'csv\single.eml'; [IO.File]::WriteAllText($one, "Subject: x`r`n`r`nb")
        (Read-MbtImportBatchCsv -Path (New-Csv 'imp-folder.csv' "Mailbox,Source,TargetRoot`r`nadele@contoso.com,pstfolder,R`r`n") -Mode ImportPst).Problems | Should -BeNullOrEmpty
        (Read-MbtImportBatchCsv -Path (New-Csv 'imp-oneeml.csv' "Mailbox,Source,TargetRoot`r`nadele@contoso.com,single.eml,R`r`n") -Mode ImportEml).Problems | Should -BeNullOrEmpty
    }
}

Describe 'The arguments of an import' {
    BeforeAll {
        function New-ImportRow([string]$Upn = 'adele@contoso.com', [string]$Location = 'Primary', [string]$Source = 'D:\src\a.pst', [string]$Root = 'Imported', [string]$SourceMailbox = '') {
            [ordered]@{ upn = $Upn; location = $Location; source = $Source; targetRoot = $Root; sourceMailbox = $SourceMailbox; key = 'abc123def456'; journal = 'D:\w\imports\abc123def456' }
        }
        $script:Job = [pscustomobject]@{ Id = 'import-ab12'; Folder = 'D:\w\jobs\x-import-ab12'; EventsPath = 'D:\w\jobs\x-import-ab12\events.jsonl'; StopPath = 'D:\w\jobs\x-import-ab12\stop.request' }
    }
    It 'is an ordinary PST import with a silent sign-in, a journal outside the job folder and the batch' {
        $a = @(Get-MbtBatchChildArgument -Row (New-ImportRow) -Option @{ AppOnly = $true; SignInUser = ''; Pass = @{ Concurrency = 2; MergeWellKnown = $true; NoCheck = $false } } -Job $script:Job -BatchId 'B1' -Mode ImportPst)
        ($a -join ' ') | Should -Match '^-Mode ImportPst -UserPrincipalName adele@contoso.com -Mailbox Primary -PstPath D:\\src\\a.pst -TempPath D:\\w\\imports\\abc123def456 -TargetRoot Imported -Concurrency 2 -MergeWellKnown -NoPrompt '
        $a | Should -Not -Contain '-NoCheck'
        $a[($a.IndexOf('-BatchId') + 1)] | Should -Be 'B1'
        $a[($a.IndexOf('-JobId') + 1)] | Should -Be 'import-ab12'
        $a | Should -Not -Contain '-StorePath'
    }
    It 'a store import names its job after the key of the row, and keeps the source mailbox; an EML import keeps its files in the job folder' {
        $s = @(Get-MbtBatchChildArgument -Row (New-ImportRow -Source 'D:\store' -SourceMailbox 'Archive' -Root '') -Option $script:AppOption -Job $script:Job -BatchId 'B1' -Mode Import)
        $s[($s.IndexOf('-StorePath') + 1)] | Should -Be 'D:\store'
        $s[($s.IndexOf('-JobName') + 1)] | Should -Be 'import-abc123def456'
        $s[($s.IndexOf('-SourceMailbox') + 1)] | Should -Be 'Archive'
        $s | Should -Not -Contain '-TargetRoot' -Because 'a store keeps the folder names when the row has no TargetRoot'
        $e = @(Get-MbtBatchChildArgument -Row (New-ImportRow -Source 'D:\eml') -Option @{ AppOnly = $true; SignInUser = ''; Pass = @{ Unread = $true; MimeConversion = 'Own' } } -Job $script:Job -BatchId 'B1' -Mode ImportEml)
        $e[($e.IndexOf('-EmlPath') + 1)] | Should -Be 'D:\eml'
        $e[($e.IndexOf('-TempPath') + 1)] | Should -Be ([IO.Path]::Combine($script:Job.Folder, 'temp'))
        $e | Should -Contain '-Unread'
    }
    It 'delegated: the signed-in account is the one that signs in and receives; the application names the target mailbox itself' {
        $d = @(Get-MbtBatchChildArgument -Row (New-ImportRow 'ADELE@contoso.com' 'Archive') -Option @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = @{} } -Job $script:Job -BatchId 'B1' -Mode ImportPst)
        $d[($d.IndexOf('-UserPrincipalName') + 1)] | Should -Be 'adele@contoso.com'
        $d[($d.IndexOf('-Mailbox') + 1)] | Should -Be 'Archive'
        $d | Should -Not -Contain '-MailboxId'
        $p = @(Get-MbtBatchChildArgument -Row (New-ImportRow 'shared@contoso.com') -Option @{ AppOnly = $true; SignInUser = ''; Pass = @{ CertificateThumbprint = 'ABCD'; TenantId = 'contoso.onmicrosoft.com'; ClientId = 'x' } } -Job $script:Job -BatchId 'B1' -Mode ImportPst)
        $p[($p.IndexOf('-UserPrincipalName') + 1)] | Should -Be 'shared@contoso.com'
        $p | Should -Contain '-CertificateThumbprint'
    }
    It 'a delegated sign-in never receives a line for another mailbox: the arguments are refused before anything is written' {
        { Get-MbtBatchChildArgument -Row (New-ImportRow 'shared@contoso.com') -Option @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = @{} } -Job $script:Job -BatchId 'B1' -Mode ImportPst } | Should -Throw '*delegated sign-in writes into the mailbox of the signed-in account (adele@contoso.com) only*Nothing was written*'
        { Get-MbtBatchChildArgument -Row (New-ImportRow 'ADELE@contoso.com') -Option @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = @{} } -Job $script:Job -BatchId 'B1' -Mode ImportPst } | Should -Not -Throw
        { Get-MbtBatchChildArgument -Row (New-ImportRow 'shared@contoso.com') -Option $script:AppOption -Job $script:Job -BatchId 'B1' -Mode ImportPst } | Should -Not -Throw
    }
    It 'every parameter it uses exists in the entry script' {
        $declared = (Get-Command $script:Entry).Parameters.Keys
        $pass = @{ Concurrency = 2; MaxAttempts = 3; RetryWaitSeconds = 2; IntervalSeconds = 5; ConfigPath = 'x.psd1'; ReconcileDelaySeconds = 60; ReconcileRounds = 2; RepairLevel = 'Rebuild'; Folder = 'Inbox'; MappingFile = 'D:\m.csv'; MimeConversion = 'Own'
            NoAutoReconcile = $true; MergeWellKnown = $true; PerFileSubfolder = $true; NoCheck = $true; Repair = $true; KeepDamagedAttachments = $true; Unread = $true; NoResume = $true; TenantId = 't'; ClientId = 'c'; CertificateThumbprint = 'ab' }
        foreach ($mode in 'Import', 'ImportPst', 'ImportEml') {
            $a = Get-MbtBatchChildArgument -Row (New-ImportRow -SourceMailbox $(if ($mode -eq 'Import') { 'Primary' } else { '' })) -Option @{ AppOnly = $false; SignInUser = 'adele@contoso.com'; Pass = $pass } -Job $script:Job -BatchId 'B' -Mode $mode
            foreach ($t in $a) { if ($t -match '^-([A-Za-z]+)$') { $declared | Should -Contain $Matches[1] -Because "$mode $t" } }
        }
    }
    It 'the request of the console for the Continue says an import, the source, the target and the folder' {
        $w = New-Work 'req'
        $module = Get-Module MailboxTransfer
        $job = & $module { param($folder) New-MbtBatchJobFolder -WorkFolder $folder -Kind import } $w
        $job.Id | Should -Match '^import-[0-9a-f]{4}$'
        $job.Name | Should -Match '^\d{8}-\d{4}-import-[0-9a-f]{4}$'
        (& $module { param($folder, $name) New-MbtBatchJobFolder -WorkFolder $folder -Name $name -Kind import } $w $job.Name).Id | Should -Be $job.Id
        $row = New-ImportRow -Source 'D:\eml' -Root 'Imported/x'
        $argument = @(Get-MbtBatchChildArgument -Row $row -Option $script:AppOption -Job $job -BatchId 'B' -Mode ImportEml)
        & $module { param($r, $j, $a, $o) Write-MbtBatchJobRequest -Row $r -Job $j -Argument $a -Option $o -BatchId 'B' -Mode ImportEml } $row $job $argument @{ AppOnly = $true; SignInUser = ''; Pass = @{ Concurrency = 3 } }
        $q = Get-Content -LiteralPath (Join-Path $job.Folder 'request.json') -Raw | ConvertFrom-Json
        $q.Task | Should -Be 'import'
        $q.ImSrc | Should -Be 'eml'
        $q.ImPath | Should -Be 'D:\eml'
        $q.Mailbox | Should -Be 'adele@contoso.com'
        $q.TgtMode | Should -Be 'new'
        $q.TgtName | Should -Be 'Imported/x'
        $q.Conc | Should -Be 3
        $q.Batch | Should -Be 'B'
        @($q.Commands[0]) | Should -Contain '-EmlPath'
    }
}

Describe 'Running an import batch on stand-in child processes' {
    It 'runs lines of different targets side by side and never two lines of the same target location, whatever the number asked at once' {
        $w = New-Work 'serial'
        $mailboxes = @('ok-pause-a@contoso.com', 'ok-pause-a@contoso.com', 'ok-pause-b@contoso.com', 'ok-pause-b@contoso.com', 'ok-pause-c@contoso.com')
        $b = New-TestImportBatch $w 'Serial' $mailboxes -Source @('sa1', 'sa2', 'sb1', 'sb2', 'sc1') -Parallel 3
        $r = Invoke-StubBatch $w $b -Parallel 3
        $r.ExitCode | Should -Be 0
        $r.Counts.Complete | Should -Be 5
        foreach ($t in 'ok-pause-a@contoso.com', 'ok-pause-b@contoso.com') {
            Get-MaxOverlap $w @($r.Batch.rows | Where-Object { $_.upn -ieq $t }) | Should -Be 1 -Because "two lines of $t ran together"
        }
        Get-MaxOverlap $w @($r.Batch.rows) | Should -BeGreaterThan 1 -Because 'different targets run side by side'
        Get-MaxOverlap $w @($r.Batch.rows) | Should -BeLessOrEqual 3
    }
    It 'the same target in another case is the same target, the Archive of it is another location, and a source read by two lines is never open twice' {
        $w = New-Work 'serial-case'
        $b = New-TestImportBatch $w 'Case' @('ok-pause-a@contoso.com', 'OK-Pause-A@Contoso.com') -Source @('ca1', 'ca2') -Parallel 2
        $r = Invoke-StubBatch $w $b -Parallel 2
        Get-MaxOverlap $w @($r.Batch.rows) | Should -Be 1
        # one source for two targets (a folder of .eml files keeps its state beside the files)
        $w2 = New-Work 'serial-source'
        $b2 = New-TestImportBatch $w2 'Same' @('ok-pause-a@contoso.com', 'ok-pause-b@contoso.com') -Source @('shared', 'shared') -Mode ImportEml -Parallel 2
        $r2 = Invoke-StubBatch $w2 $b2 -Parallel 2
        $r2.ExitCode | Should -Be 0
        Get-MaxOverlap $w2 @($r2.Batch.rows) | Should -Be 1
        # the Primary and the Archive of one mailbox are two locations: they run together
        $w3 = New-Work 'serial-loc'
        $csv = New-Csv 'loc.csv' ("Mailbox,Source,Location`r`nok-pause-a@contoso.com," + (New-Source 'ImportPst' 'la1') + ",Primary`r`nok-pause-a@contoso.com," + (New-Source 'ImportPst' 'la2') + ",Archive`r`n")
        $read = Read-MbtImportBatchCsv -Path $csv -Mode ImportPst -DefaultTargetRoot R
        $b3 = New-MbtBatch -Id 'Loc' -Csv $csv -Row $read.Rows -Option $script:AppOption -Parallel 2 -Mode ImportPst
        $r3 = Invoke-StubBatch $w3 $b3 -Parallel 2
        Get-MaxOverlap $w3 @($r3.Batch.rows) | Should -Be 2
    }
    It 'every line is an ordinary import job in its own folder with the batch in its status, the arguments of an import and its own journal' {
        $w = New-Work 'jobs'
        $b = New-TestImportBatch $w 'Jobs' @('ok-1@contoso.com', 'ok-2@contoso.com') -Source @('j1', 'j2') -Root 'Restored'
        $b.mode | Should -Be 'ImportPst'
        $r = Invoke-StubBatch $w $b
        $r.ExitCode | Should -Be 0
        foreach ($row in $r.Batch.rows) {
            $row.state | Should -Be 'complete'
            $row.done | Should -Be 100
            $status = Get-Content -LiteralPath (Join-Path (Join-Path (Join-Path $w 'jobs') $row.job) 'status.json') -Raw | ConvertFrom-Json
            $status.batch | Should -Be 'Jobs'
            $status.task | Should -Be 'import'
            $status.mode | Should -Be 'ImportPst'
            $given = Get-Given $w $row
            $given[($given.IndexOf('-Mode') + 1)] | Should -Be 'ImportPst'
            $given[($given.IndexOf('-PstPath') + 1)] | Should -Be $row.source
            $given[($given.IndexOf('-TargetRoot') + 1)] | Should -Be 'Restored'
            $given[($given.IndexOf('-TempPath') + 1)] | Should -Be $row.journal
            $given[($given.IndexOf('-JobId') + 1)] | Should -Match '^import-[0-9a-f]{4}$'
            Test-Path -LiteralPath $row.journal -PathType Container | Should -BeTrue
            Test-Path -LiteralPath (Join-Path (Join-Path (Join-Path $w 'jobs') $row.job) 'request.json') | Should -BeTrue
        }
        (Read-MbtBatch -Folder (Join-Path $w 'batches\Jobs')).mode | Should -Be 'ImportPst'
        (Get-MbtBatchList -WorkFolder $w)[0].Mode | Should -Be 'ImportPst'
    }
    It 'continues what is not complete: a line done is skipped, one that stopped goes on in its own job folder and journal, a failed one is tried again; exit 1, then 1, then 0' {
        $w = New-Work 'resume'
        $b = New-TestImportBatch $w 'Res' @('ok-1@contoso.com', 'partial-2@contoso.com', 'fail-3@contoso.com') -Source @('r1', 'r2', 'r3')
        $first = Invoke-StubBatch $w $b
        $first.ExitCode | Should -Be 1
        @($first.Batch.rows.state) | Should -Be @('complete', 'partial', 'failed')
        $job = (Get-BatchRow $first.Batch 'partial-2@contoso.com').job
        $journal = (Get-BatchRow $first.Batch 'partial-2@contoso.com').journal
        $again = Read-MbtBatch -Folder (Join-Path $w 'batches\Res')
        $second = Invoke-StubBatch $w $again
        $second.ExitCode | Should -Be 1
        @($second.Batch.rows.state) | Should -Be @('complete', 'complete', 'failed')
        (Get-BatchRow $second.Batch 'partial-2@contoso.com').job | Should -Be $job
        (Get-BatchRow $second.Batch 'partial-2@contoso.com').journal | Should -Be $journal
        (Get-BatchRow $second.Batch 'ok-1@contoso.com').attempts | Should -Be 1
        (Get-BatchRow $second.Batch 'fail-3@contoso.com').attempts | Should -Be 2
    }
    It 'a line that needs a sign-in is not complete and the others go on; uncertain items are not complete either; items refused for good fail the line' {
        $w = New-Work 'states'
        $b = New-TestImportBatch $w 'States' @('signin-1@contoso.com', 'ok-2@contoso.com', 'uncertain-3@contoso.com', 'refused-4@contoso.com') -Source @('t1', 't2', 't3', 't4')
        $r = Invoke-StubBatch $w $b -Parallel 2
        @($r.Batch.rows.state) | Should -Be @('partial', 'complete', 'partial', 'failed')
        (Get-BatchRow $r.Batch 'signin-1@contoso.com').message | Should -Match '^sign-in required: .*run it again'
        (Get-BatchRow $r.Batch 'uncertain-3@contoso.com').message | Should -Match '^uncertain items: .*they are looked up in the mailbox first'
        (Get-BatchRow $r.Batch 'uncertain-3@contoso.com').unknown | Should -Be 2
        (Get-BatchRow $r.Batch 'refused-4@contoso.com').message | Should -Match 'some items are refused for good'
        (Get-BatchRow $r.Batch 'refused-4@contoso.com').refused | Should -Be 2
        $r.ExitCode | Should -Be 1
        $c = Get-MbtBatchCount $r.Batch
        $c.Complete | Should -Be 1
        $c.Partial | Should -Be 2
        $c.Failed | Should -Be 1
        $c.Done | Should -Be (100 + 3 + 4)
        $c.Unknown | Should -Be 2
        $c.Refused | Should -Be 2
        # without the failed line the batch is only not complete: exit 2, and it ends complete when the lines are run again
        $w2 = New-Work 'states2'
        $b2 = New-TestImportBatch $w2 'States2' @('signin-1@contoso.com', 'uncertain-2@contoso.com') -Source @('u1', 'u2')
        $r2 = Invoke-StubBatch $w2 $b2
        $r2.ExitCode | Should -Be 2
        $again = Read-MbtBatch -Folder (Join-Path $w2 'batches\States2')
        (Get-BatchRow $again 'uncertain-2@contoso.com').job | Should -Not -BeNullOrEmpty
        $r3 = Invoke-StubBatch $w2 $again
        (Get-BatchRow $r3.Batch 'uncertain-2@contoso.com').state | Should -Be 'complete'
        $r3.ExitCode | Should -Be 2 -Because 'the line that needs a sign-in is still not complete'
    }
    It 'writes the report of the batch with the items imported, refused and uncertain in its caption, the jobs of this batch only' {
        $w = New-Work 'report'
        $b = New-TestImportBatch $w 'Rep' @('ok-1@contoso.com', 'uncertain-2@contoso.com', 'refused-3@contoso.com') -Source @('p1', 'p2', 'p3')
        $r = Invoke-StubBatch $w $b
        $html = Get-Content -LiteralPath (Join-Path $w 'batches\Rep\report.html') -Raw
        $html | Should -Match '3 mailboxes in the CSV file \(1 complete, 1 not complete, 1 failed\); 107 items imported, 2 refused, 2 uncertain\.'
        foreach ($row in $r.Batch.rows) { $html | Should -Match ([regex]::Escape($row.job)) }
        Test-Path (Join-Path $w 'jobs\overview.html') | Should -BeFalse -Because 'a job of a batch leaves the overview of all the jobs to the batch'
    }
    It 'the summary card says imported, refused for good, failed and uncertain, and what to do next' {
        $w = New-Work 'card'
        $b = New-TestImportBatch $w 'Card' @('ok-1@contoso.com', 'uncertain-2@contoso.com', 'refused-3@contoso.com') -Source @('k1', 'k2', 'k3')
        $r = Invoke-StubBatch $w $b
        $text = (& { Write-MbtBatchSummary -ExitCode $r.ExitCode -Batch $r.Batch -Folder $r.Folder -Seconds 75 -Report $r.Report -Command '.\Invoke-MailboxTransfer.ps1 -Mode ImportPst -ResumeBatch Card' -Mode ImportPst } *>&1 | Out-String)
        $text | Should -Match 'Import batch finished with failures'
        $text | Should -Match 'exit code 1 - failed'
        $text | Should -Match 'Items\s+107 imported, 2 refused for good, 2 uncertain'
        $text | Should -Match 'Next steps\s+Run the same command again to continue'
        $text | Should -Match '-Mode ImportPst -ResumeBatch Card'
        $complete = (& { Write-MbtBatchSummary -ExitCode 0 -Batch $r.Batch -Mode ImportEml } *>&1 | Out-String)
        $complete | Should -Match 'Import batch finished'
        $complete | Should -Match 'adds nothing: an item that was imported is never created twice'
        $planned = (& { Write-MbtBatchSummary -ExitCode 0 -Batch $r.Batch -WhatIf -Mode Import } *>&1 | Out-String)
        $planned | Should -Match 'Import batch planned \(WhatIf, nothing was started\)'
    }
    It 'a batch written before import existed is an export; the rows of an import batch are compared by target, location, source, store mailbox and folder' {
        (Get-MbtBatchMode @{ rows = @() }) | Should -Be 'Export'
        (Get-MbtBatchMode ([ordered]@{ mode = 'ImportEml' })) | Should -Be 'ImportEml'
        $w = New-Work 'compare'
        $b = New-TestImportBatch $w 'Cmp' @('ok-1@contoso.com', 'ok-2@contoso.com') -Source @('m1', 'm2') -Root 'R'
        $csv = New-Csv 'cmp2.csv' ("Mailbox,Source,TargetRoot`r`nOK-1@contoso.com," + (Join-Path (Join-Path $TestDrive 'csv') 'M1.pst') + ",r`r`nok-3@contoso.com," + (New-Source 'ImportPst' 'm3') + ",R`r`n")
        $now = Read-MbtImportBatchCsv -Path $csv -Mode ImportPst
        $diff = @(Compare-MbtImportBatchRow -Batch $b -Row $now.Rows)
        $diff.Count | Should -Be 2
        ($diff -join "`n") | Should -Match 'ok-3@contoso.com \(Primary\) from .*m3\.pst into R is in the CSV file and not in the batch'
        ($diff -join "`n") | Should -Match 'ok-2@contoso.com \(Primary\) from .*m2\.pst into R is in the batch and not in the CSV file'
    }
}

Describe 'The command line of an import batch' {
    BeforeAll {
        function Invoke-Cli {
            param([hashtable]$Parameters)
            $global:LASTEXITCODE = $null
            $text = & $script:Entry @Parameters *>&1 | Out-String
            $global:CliExit = $LASTEXITCODE
            return $text
        }
        function Get-Parameters($Base, $Change) { $h = $Base.Clone(); foreach ($k in $Change.Keys) { $h[$k] = $Change[$k] }; return $h }
        $script:Work = New-Work 'cli'
        $script:Common = @{ Mode = 'ImportPst'; AppOnly = $true; CertificateThumbprint = 'ABCDEF0123456789ABCDEF0123456789ABCDEF01'; TenantId = $script:Tenant; ClientId = '66666666-7777-8888-9999-000000000000'; WorkFolder = $script:Work; LogFolder = (Join-Path $script:Work 'logs'); TargetRoot = 'Imported' }
    }
    AfterAll { Remove-Variable CliExit -Scope Global -ErrorAction SilentlyContinue }
    It 'prints every problem of the CSV file in the error card, exits 1 and starts nothing' {
        $ok = New-Source 'ImportPst' 'cli-ok'
        $csv = New-Csv 'cli-bad.csv' "Mailbox,Source,Colour`r`n,cli-ok.pst,red`r`nadele@contoso.com,nowhere.pst,`r`nalex@contoso.com,cli-ok.pst,`r`nALEX@contoso.com,CLI-OK.PST,`r`n"
        $text = Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; BatchName = 'Bad' })
        $global:CliExit | Should -Be 1
        $text | Should -Match 'The batch cannot start'
        $text | Should -Match 'The column Colour is not known'
        $text | Should -Match 'line 2: Mailbox is empty'
        $text | Should -Match 'line 3: the Source was not found'
        $text | Should -Match 'line 5: ALEX@contoso.com \(Primary\) already imports .* on line 4'
        $text | Should -Match 'Import batch not started'
        $text | Should -Match '4 problems, listed above: nothing was started'
        Test-Path (Join-Path $script:Work 'batches') | Should -BeFalse
    }
    It '-WhatIf checks the file and the sources, lists every target with its source and folder, and writes and starts nothing' {
        $null = New-Source 'ImportPst' 'cli-p1'; $null = New-Source 'ImportPst' 'cli-p2'; $null = New-Source 'ImportPst' 'cli-p3'
        $csv = New-Csv 'cli-plan.csv' "Mailbox,Source,Location,TargetRoot`r`nadele@contoso.com,cli-p1.pst,Primary,`r`nadele@contoso.com,cli-p2.pst,Archive,Own/A`r`nshared@contoso.com,cli-p3.pst,,`r`n"
        $text = Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; BatchName = 'Plan'; WhatIf = $true; ParallelMailboxes = 2 })
        $global:CliExit | Should -Be 0
        $text | Should -Match '3 imports: 0 complete, 3 to do; 2 at a time'
        $text | Should -Match 'adele@contoso.com - Primary - from .*cli-p1\.pst into the folder Imported'
        $text | Should -Match 'adele@contoso.com - Archive - from .*cli-p2\.pst into the folder Own/A'
        $text | Should -Match 'shared@contoso.com - Primary - from .*cli-p3\.pst into the folder Imported'
        $text | Should -Match 'these 3 mailbox locations will be written to: adele@contoso.com \(Primary\), adele@contoso.com \(Archive\), shared@contoso.com \(Primary\)\. Items are created only'
        $text | Should -Match 'WhatIf: the file is valid and every source was found; nothing was started and nothing was written to any mailbox'
        $text | Should -Match 'Import batch planned'
        Test-Path (Join-Path $script:Work 'batches') | Should -BeFalse
        Test-Path (Join-Path $script:Work 'imports') | Should -BeFalse
    }
    It 'says that lines of one target or one source run one after the other' {
        $null = New-Source 'ImportPst' 'cli-s1'; $null = New-Source 'ImportPst' 'cli-s2'
        $csv = New-Csv 'cli-serial.csv' "Mailbox,Source`r`nadele@contoso.com,cli-s1.pst`r`nadele@contoso.com,cli-s2.pst`r`nalex@contoso.com,cli-s1.pst`r`n"
        $text = Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; BatchName = 'Serial'; WhatIf = $true })
        $global:CliExit | Should -Be 0
        $text | Should -Match '1 target mailbox location is written by several lines: those lines run one after the other'
        $text | Should -Match '1 source is read by several lines: those lines run one after the other'
    }
    It 'with a delegated sign-in refuses a target that is not the signed-in account, in the error card' {
        $null = New-Source 'ImportPst' 'cli-d1'
        $csv = New-Csv 'cli-deleg.csv' "Mailbox,Source`r`nadele@contoso.com,cli-d1.pst`r`nshared@contoso.com,cli-d1.pst`r`n"
        $delegated = @{ Mode = 'ImportPst'; UserPrincipalName = 'adele@contoso.com'; TenantId = $script:Tenant; ClientId = '66666666-7777-8888-9999-000000000000'; WorkFolder = $script:Work; LogFolder = (Join-Path $script:Work 'logs'); TargetRoot = 'Imported' }
        $text = Invoke-Cli ($delegated + @{ MailboxCsv = $csv; BatchName = 'Deleg' })
        $global:CliExit | Should -Be 1
        $text | Should -Match 'line 3: shared@contoso.com is not the signed-in account \(adele@contoso.com\)'
        $text | Should -Match 'sign in as the application \(-AppOnly\)'
    }
    It 'the same command again on a batch that is complete starts nothing and signs in nowhere; the name is found in any case; the mode must be the one of the batch' {
        $w2 = New-Work 'cli-done'
        $p = $script:Common.Clone(); $p.WorkFolder = $w2; $p.LogFolder = (Join-Path $w2 'logs')
        $b = New-TestImportBatch $w2 'Done1' @('ok-1@contoso.com', 'ok-2@contoso.com') -Source @('d1', 'd2') -Root 'Imported'
        foreach ($row in $b.rows) { $row.state = 'complete'; $row.job = ''; $row.done = 5 }
        Save-MbtBatch -Batch $b -Folder (Join-Path $w2 'batches\Done1')
        $text = Invoke-Cli ($p + @{ MailboxCsv = $b.csv; BatchName = 'done1' })
        $global:CliExit | Should -Be 0
        $text | Should -Match 'every mailbox of this batch is already complete: nothing to do'
        $text | Should -Match 'Import batch finished'
        $text | Should -Match 'Mailboxes\s+2 in Done1: 2 complete, 0 not complete, 0 failed, 0 not started'
        $text | Should -Match 'Items\s+10 imported, 0 uncertain'
        $text | Should -Not -Match 'signing in'
        $again = Invoke-Cli ($p + @{ ResumeBatch = 'DONE1' })
        $global:CliExit | Should -Be 0
        $again | Should -Match '2 imports: 2 complete, 0 to do'
        $wrong = Invoke-Cli (Get-Parameters $p @{ ResumeBatch = 'Done1'; Mode = 'ImportEml' })
        $global:CliExit | Should -Be 1
        $wrong | Should -Match 'The batch Done1 was made for -Mode ImportPst, this command says -Mode ImportEml'
        $export = Invoke-Cli (Get-Parameters $p @{ ResumeBatch = 'Done1'; Mode = 'Export' })
        $global:CliExit | Should -Be 1
        $export | Should -Match 'was made for -Mode ImportPst, this command says -Mode Export'
        $other = New-Csv 'done1-other.csv' ("Mailbox,Source`r`nok-1@contoso.com," + (Join-Path (Join-Path $TestDrive 'csv') 'd1.pst') + "`r`nok-3@contoso.com," + (Join-Path (Join-Path $TestDrive 'csv') 'd1.pst') + "`r`n")
        $refused = Invoke-Cli ($p + @{ MailboxCsv = $other; BatchName = 'Done1' })
        $global:CliExit | Should -Be 1
        $refused | Should -Match 'The CSV file is not the list of the batch Done1: ok-3@contoso.com \(Primary\) from .*d1\.pst into Imported is in the CSV file and not in the batch'
        $refused | Should -Match 'ok-2@contoso.com \(Primary\) from .*d2\.pst into Imported is in the batch and not in the CSV file'
        $refused | Should -Match 'use -ResumeBatch to continue it, or another -BatchName'
    }
    It 'a batch made for the application is not continued with a delegated sign-in that is not every one of its mailboxes' {
        $w4 = New-Work 'cli-foreign'
        $b = New-TestImportBatch $w4 'Foreign' @('ok-1@contoso.com', 'ok-2@contoso.com') -Source @('f1', 'f2')
        Save-MbtBatch -Batch $b -Folder (Join-Path $w4 'batches\Foreign')
        $delegated = @{ Mode = 'ImportPst'; UserPrincipalName = 'ok-1@contoso.com'; TenantId = $script:Tenant; ClientId = '66666666-7777-8888-9999-000000000000'; WorkFolder = $w4; LogFolder = (Join-Path $w4 'logs'); ResumeBatch = 'Foreign' }
        $text = Invoke-Cli $delegated
        $global:CliExit | Should -Be 1
        $text | Should -Match 'ok-2@contoso.com is not the signed-in account \(ok-1@contoso.com\)\. A delegated sign-in writes into its own mailbox only'
        $text | Should -Match 'nothing was started'
        $text | Should -Not -Match 'ok-1@contoso.com is not the signed-in account'
        @(Get-ChildItem -LiteralPath $w4 -Directory -Filter jobs -ErrorAction SilentlyContinue).Count | Should -Be 0
    }
    It 'says what is wrong with the options: a source on the command line, a Location that is not for an import, a folder and a source mailbox that do not fit the mode' {
        $null = New-Source 'ImportPst' 'cli-o1'
        $csv = New-Csv 'cli-opt.csv' "Mailbox,Source`r`nadele@contoso.com,cli-o1.pst`r`n"
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; PstPath = 'x.pst' }) } | Should -Throw '*-PstPath is for one import*Source column*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; JobName = 'x' }) } | Should -Throw '*-JobName is for one import*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; Mailbox = 'All' }) } | Should -Throw '*not for a batch of imports*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; Mailbox = 'AuxArchive' }) } | Should -Throw '*not for a batch of imports*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; SourceMailbox = 'Primary' }) } | Should -Throw '*belongs to -Mode Import*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; Folder = 'Inbox' }) } | Should -Throw '*belongs to -Mode Import*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; MappingFile = (Join-Path $TestDrive 'none.csv') }) } | Should -Throw '*-MappingFile not found*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; OutputPst = 'D:\p' }) } | Should -Throw '*-OutputPst belongs to*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; ArrivedFrom = '2026-01-01' }) } | Should -Throw '*an import takes every item*'
        { Invoke-Cli @{ Mode = 'ImportPst'; BatchName = 'X' } } | Should -Throw '*belong to -MailboxCsv*'
    }
    It 'the entry script of every stand-in batch is complete: a batch of a store and of .eml files is read the same way' {
        $null = New-Source 'Import' 'cli-st'; $null = New-Source 'ImportEml' 'cli-em'
        $store = New-Csv 'cli-store.csv' "Mailbox,Source,SourceMailbox`r`nadele@contoso.com,cli-st,Primary`r`n"
        $storeParameters = Get-Parameters $script:Common @{ Mode = 'Import'; MailboxCsv = $store; BatchName = 'StoreBatch'; WhatIf = $true }
        $storeParameters.Remove('TargetRoot')
        $t = Invoke-Cli $storeParameters
        $global:CliExit | Should -Be 0
        $t | Should -Match 'adele@contoso.com - Primary - from .*cli-st \(Primary\) with the folder names of the source'
        $t | Should -Match 'Imports export stores into the mailboxes of a list'
        $em = New-Csv 'cli-eml.csv' "Mailbox,Source,TargetRoot`r`nadele@contoso.com,cli-em,From-eml`r`n"
        $t2 = Invoke-Cli (Get-Parameters $script:Common @{ Mode = 'ImportEml'; MailboxCsv = $em; BatchName = 'EmlBatch'; WhatIf = $true })
        $global:CliExit | Should -Be 0
        $t2 | Should -Match 'adele@contoso.com - Primary - from .*cli-em into the folder From-eml'
    }
}

Describe 'What the retention of the job folders and the list of batches say of an import batch' {
    It 'keeps the jobs and the journal folder of a batch that is not finished, and lists it with its mode and its numbers' {
        $w = New-Work 'keep'
        $b = New-TestImportBatch $w 'Keep' @('ok-1@contoso.com', 'signin-2@contoso.com') -Source @('q1', 'q2')
        $r = Invoke-StubBatch $w $b
        $r.ExitCode | Should -Be 2
        $protect = Get-MbtBatchProtection -WorkFolder $w
        $protect.Batches.Contains('keep') | Should -BeTrue
        foreach ($row in $r.Batch.rows) { $protect.Jobs.Contains($row.job) | Should -BeTrue }
        $list = @(Get-MbtBatchList -WorkFolder $w)
        $list[0].Mode | Should -Be 'ImportPst'
        $list[0].Total | Should -Be 2
        $list[0].Complete | Should -Be 1
        $list[0].Partial | Should -Be 1
        foreach ($row in $r.Batch.rows) { $row.journal | Should -BeLike (Join-Path $w 'imports\*') }
        Test-Path -LiteralPath (Join-Path $w 'jobs') | Should -BeTrue
    }
}

Describe 'A batch of real imports: the entry script run against the fake service for each line' {
    BeforeAll {
        function New-FakeClient {
            param([Parameter(Mandatory)]$Service)
            $o = [MailboxTransfer.GraphOptions]::new()
            $o.DelayScale = 0
            $o.UseJitter = $false
            return [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
        }
        $script:Launcher = {
            param($Argument, $Job, $Batch)
            $named = @{}
            for ($i = 0; $i -lt $Argument.Count; $i++) {
                $t = [string]$Argument[$i]
                if ($t -notmatch '^-([A-Za-z]+)$') { continue }
                $key = $t.Substring(1)
                if (($i + 1) -lt $Argument.Count -and [string]$Argument[$i + 1] -notmatch '^-[A-Za-z]+$') { $named[$key] = [string]$Argument[$i + 1]; $i++ } else { $named[$key] = $true }
            }
            $named.PrimaryMailboxId = $script:Src; $named.ArchiveMailboxId = $script:Src
            $named.DevelopmentMode = $true
            $named.LogFolder = Join-Path $script:Work 'logs'
            $global:MbtEntryClient = New-FakeClient -Service $global:MbtBatchService
            $global:LASTEXITCODE = $null
            $null = & $script:Entry @named *>&1 | Out-String
            [pscustomobject]@{ Process = [pscustomobject]@{ HasExited = $true; ExitCode = $global:LASTEXITCODE }; Pipes = @() }
        }
        $script:Work = New-Work 'real'
    }
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph $global:MbtEntryClient
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtEntryClient, MbtBatchService -Scope Global -ErrorAction SilentlyContinue }
    It 'imports each folder of .eml files as an ordinary job below its TargetRoot; the same batch run again creates nothing twice' {
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $svc.RootOnlyFolders = $true
        $null = $svc.AddMailbox($script:Src)
        $global:MbtBatchService = $svc
        $one = New-Source 'ImportEml' 'real-one'
        $two = New-Source 'ImportEml' 'real-two'
        $csv = New-Csv 'real-eml.csv' "Mailbox,Source,TargetRoot`r`nadele.vance@contoso.com,$one,From-one`r`nadele.vance@contoso.com,$two,From-two`r`n"
        $read = Read-MbtImportBatchCsv -Path $csv -Mode ImportEml -SignedInUser 'adele.vance@contoso.com' -WorkFolder $script:Work
        $read.Problems | Should -BeNullOrEmpty
        $option = @{ AppOnly = $false; SignInUser = 'adele.vance@contoso.com'; Pass = @{ Concurrency = 2; IntervalSeconds = 60 } }
        $batch = New-MbtBatch -Id 'RealEml' -Csv $csv -Row $read.Rows -Option $option -Parallel 2 -Mode ImportEml
        $run = { param($b) Invoke-MbtBatch -WorkFolder $script:Work -Batch $b -EntryScript $script:Entry -Option $option -Parallel 2 -Launcher $script:Launcher -PollMilliseconds 50 -StartGapMilliseconds 0 6>&1 | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.InformationRecord] } }
        $r = & $run $batch
        $r.ExitCode | Should -Be 0 -Because (@($r.Batch.rows.message) -join ' | ')
        @($r.Batch.rows.state) | Should -Be @('complete', 'complete')
        @($svc.ItemsOf($script:Src)).Count | Should -Be 2
        $names = @($svc.FoldersOf($script:Src) | ForEach-Object Name)
        $names | Should -Contain 'From-one'
        $names | Should -Contain 'From-two'
        foreach ($row in $r.Batch.rows) {
            $dir = Join-Path $script:Work "jobs\$($row.job)"
            $status = Get-Content -LiteralPath (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json
            $status.batch | Should -Be 'RealEml'
            $status.state | Should -Be 'complete'
            Test-Path (Join-Path $dir 'report.html') | Should -BeTrue
            $row.done | Should -Be 1
        }
        (Get-MbtBatchCount $r.Batch).Done | Should -Be 2
        $html = Get-Content -LiteralPath (Join-Path $script:Work 'batches\RealEml\report.html') -Raw
        $html | Should -Match '2 mailboxes in the CSV file \(2 complete, 0 not complete, 0 failed\); 2 items imported, 0 refused, 0 uncertain\.'
        # the batch is complete: run again, nothing starts; a new batch with the same lines skips the files that were imported
        $csv2 = New-Csv 'real-eml2.csv' "Mailbox,Source,TargetRoot`r`nADELE.VANCE@contoso.com,$($one.ToUpperInvariant()),from-one`r`n"
        $read2 = Read-MbtImportBatchCsv -Path $csv2 -Mode ImportEml -SignedInUser 'adele.vance@contoso.com' -WorkFolder $script:Work
        $read2.Problems | Should -BeNullOrEmpty
        $batch2 = New-MbtBatch -Id 'RealEml2' -Csv $csv2 -Row $read2.Rows -Option $option -Parallel 1 -Mode ImportEml
        $r2 = & $run $batch2
        $r2.ExitCode | Should -Be 0 -Because (@($r2.Batch.rows.message) -join ' | ')
        @($svc.ItemsOf($script:Src)).Count | Should -Be 2 -Because 'the file was imported by the first batch: no item is created twice'
        $r2.Batch.rows[0].done | Should -Be 0
    }
}
