#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the export of a list of mailboxes (module\MailboxTransfer.Batch.ps1 and the entry script with -MailboxCsv): the checks of the CSV file (every problem at
    once), the output names, the arguments of a mailbox, the scheduler (one at a time and several at a time, resume, stop, sign-in, a crash) on stand-in child processes
    (tests\support\batch-child.ps1), the batch report, what the retention keeps, and one batch whose mailboxes are the real entry script run against the fake service.
    Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\Batch.Tests.ps1 -Output Detailed
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

    function New-Csv([string]$Name, [string]$Text, [switch]$Bom) {
        $dir = Join-Path $TestDrive 'csv'
        $null = New-Item -ItemType Directory -Force -Path $dir
        $path = Join-Path $dir $Name
        [IO.File]::WriteAllText($path, $Text, [Text.UTF8Encoding]::new([bool]$Bom))
        return $path
    }
    function New-Work([string]$Name) { $w = Join-Path $TestDrive $Name; $null = New-Item -ItemType Directory -Force -Path $w; return $w }
    # The batch of a list of mailboxes, rows as the CSV reader makes them.
    function New-TestBatch {
        param([string]$Work, [string]$Id, [string[]]$Mailbox, [hashtable]$Option = @{ AppOnly = $true; SignInUser = ''; Pass = @{} }, [int]$Parallel = 1, [string[]]$Location = @())
        $csv = New-Csv "$Id.csv" ("UserPrincipalName,Location`r`n" + (($Mailbox | ForEach-Object -Begin { $i = 0 } -Process { $l = if ($i -lt $Location.Count) { $Location[$i] } else { '' }; $i++; "$_,$l" }) -join "`r`n"))
        $read = Read-MbtBatchCsv -Path $csv -StoreBase (Join-Path $Work 'stores') -OutputPst (Join-Path $Work 'pst')
        $read.Problems | Should -BeNullOrEmpty
        return (New-MbtBatch -Id $Id -Csv $csv -Row $read.Rows -Option $Option -Parallel $Parallel)
    }
    function Invoke-StubBatch {
        param([string]$Work, $Batch, [hashtable]$Option = @{ AppOnly = $true; SignInUser = ''; Pass = @{} }, [int]$Parallel = 1, [string]$StopFile)
        $a = @{ WorkFolder = $Work; Batch = $Batch; EntryScript = $script:Stub; Option = $Option; Parallel = $Parallel; PollMilliseconds = 100; StartGapMilliseconds = 50; StopGraceSeconds = 20 }
        if ($StopFile) { $a.StopFile = $StopFile }
        return (Invoke-MbtBatch @a 6>&1 | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.InformationRecord] })
    }
    function Get-BatchRow($Batch, [string]$Upn) { @($Batch.rows | Where-Object { $_.upn -ieq $Upn })[0] }
    function Get-JobFile($Work, $Row, [string]$Name) { Join-Path (Join-Path (Join-Path $Work 'jobs') $Row.job) $Name }
    function Get-JobStatus($Work, $Row) { Get-Content -LiteralPath (Get-JobFile $Work $Row 'status.json') -Raw | ConvertFrom-Json }
    # When a stand-in child began and when it wrote its last status: the intervals of the jobs, to see how many ran at the same time.
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
AfterAll { $env:MAILBOXTRANSFER_LOCAL_CONFIG = $script:SavedLocal; Remove-Item Env:\MBT_STUB_STOPFILE -ErrorAction SilentlyContinue }

Describe 'The CSV file of a batch' {
    It 'is read with a comma or a semicolon, with or without a BOM, headers in any case, columns it does not know ignored' {
        foreach ($case in @(
                @{ Text = "UserPrincipalName,Location,Notes`r`nadele@contoso.com,Archive,first`r`nalex@contoso.com,,second`r`n"; Bom = $true }
                @{ Text = "userprincipalname;LOCATION;Notes`nadele@contoso.com;archive;first`nalex@contoso.com;;second`n"; Bom = $false }
                @{ Text = "Notes,USERPRINCIPALNAME,location`r`nfirst,adele@contoso.com,ARCHIVE`r`nsecond,alex@contoso.com,`r`n"; Bom = $true }
            )) {
            $r = Read-MbtBatchCsv -Path (New-Csv 'ok.csv' $case.Text -Bom:$case.Bom)
            $r.Problems | Should -BeNullOrEmpty
            @($r.Rows).Count | Should -Be 2
            @($r.Rows.UserPrincipalName) | Should -Be @('adele@contoso.com', 'alex@contoso.com')
            @($r.Rows.Location) | Should -Be @('Archive', 'Primary')
            @($r.Rows.Line) | Should -Be @(2, 3)
        }
    }
    It 'takes a quoted field that holds the delimiter, and says which delimiter it found' {
        $r = Read-MbtBatchCsv -Path (New-Csv 'quoted.csv' "UserPrincipalName;Notes;Location`r`nadele@contoso.com;`"a; b`";Both`r`n")
        $r.Problems | Should -BeNullOrEmpty
        $r.Delimiter | Should -Be ';'
        $r.Rows[0].Location | Should -Be 'Both'
        (Read-MbtBatchCsv -Path (New-Csv 'comma.csv' "UserPrincipalName,Notes`r`nadele@contoso.com,x`r`n")).Delimiter | Should -Be ','
    }
    It 'the Location of a row that has none is the default of the command, Primary unless -Mailbox says otherwise; All is Both' {
        $text = "UserPrincipalName,Location`r`nadele@contoso.com,`r`nalex@contoso.com,Primary`r`nlee@contoso.com,all`r`n"
        @((Read-MbtBatchCsv -Path (New-Csv 'def.csv' $text)).Rows.Location) | Should -Be @('Primary', 'Primary', 'Both')
        @((Read-MbtBatchCsv -Path (New-Csv 'def.csv' $text) -DefaultLocation Archive).Rows.Location) | Should -Be @('Archive', 'Primary', 'Both')
        @((Read-MbtBatchCsv -Path (New-Csv 'def.csv' $text) -DefaultLocation All).Rows.Location) | Should -Be @('Both', 'Primary', 'Both')
    }
    It 'reports every problem at once, with the line of each, and nothing is lost to the first one' {
        $text = @(
            'UserPrincipalName,Location,SelectionFile,OutputPst'
            'adele@contoso.com,Primary,,'
            ',Archive,,'
            'Adele@Contoso.com,Primary,,'
            'not an address,,,'
            'alex@contoso.com,Everything,,'
            'lee@contoso.com,,missing-selection.json,'
            'kim@contoso.com,,,bad|name.pst'
            'sam@contoso.com,,,aux.pst'
        ) -join "`r`n"
        $r = Read-MbtBatchCsv -Path (New-Csv 'many.csv' $text)
        $r.Problems.Count | Should -Be 7
        ($r.Problems -join "`n") | Should -Match 'line 3: UserPrincipalName is empty'
        ($r.Problems -join "`n") | Should -Match 'line 4: Adele@Contoso.com is already on line 2'
        ($r.Problems -join "`n") | Should -Match "line 5: 'not an address' is not a mailbox address"
        ($r.Problems -join "`n") | Should -Match "line 6: Location 'Everything' is not Primary, Archive or Both"
        ($r.Problems -join "`n") | Should -Match 'line 7: the SelectionFile was not found'
        ($r.Problems -join "`n") | Should -Match 'line 8: .*bad\|name\.pst'
        ($r.Problems -join "`n") | Should -Match 'line 9: the PST file name aux\.pst cannot be used'
        @($r.Rows).Count | Should -Be 1
    }
    It 'takes the same address on several lines for its Primary mailbox, its Archive or other selections, and names their files after the Location or the line; the same Location twice, Both with another, or the same selection twice are refused' {
        $one = New-Csv 'sel-one.json' '{"one":1}'
        $two = New-Csv 'sel-two.json' '{"two":2}'
        $text = "UserPrincipalName,Location,SelectionFile`r`nadele@contoso.com,Primary,`r`nAdele@contoso.com,Archive,`r`nalex@contoso.com,Both,`r`nalex@contoso.com,Primary,`r`nlee@contoso.com,,`r`nlee@contoso.com,Primary,`r`nkim@contoso.com,,sel-one.json`r`nkim@contoso.com,,sel-two.json`r`nkim@contoso.com,,sel-one.json`r`n"
        $r = Read-MbtBatchCsv -Path (New-Csv 'again.csv' $text) -OutputPst 'D:\pst' -StoreBase 'D:\stores'
        $r.Problems.Count | Should -Be 3
        $r.Problems[0] | Should -Match 'line 5: alex@contoso.com is already on line 4'
        $r.Problems[1] | Should -Match 'line 7: lee@contoso.com is already on line 6'
        $r.Problems[2] | Should -Match 'line 10: kim@contoso.com is already on line 8'
        @($r.Rows.Line) | Should -Be @(2, 3, 4, 6, 8, 9)
        @($r.Rows.Output) | Should -Be @('D:\pst\adele@contoso.com-primary.pst', 'D:\pst\Adele@contoso.com-archive.pst', 'D:\pst\alex@contoso.com.pst', 'D:\pst\lee@contoso.com.pst', 'D:\pst\kim@contoso.com-row7.pst', 'D:\pst\kim@contoso.com-row8.pst')
        @($r.Rows.Store) | Should -Be @('D:\stores\adele@contoso.com-primary', 'D:\stores\Adele@contoso.com-archive', 'D:\stores\alex@contoso.com', 'D:\stores\lee@contoso.com', 'D:\stores\kim@contoso.com-row7', 'D:\stores\kim@contoso.com-row8')
    }
    It 'refuses a file without the column, an empty file, a missing file and a header that repeats a column' {
        (Read-MbtBatchCsv -Path (New-Csv 'nocol.csv' "Mail,Location`r`nadele@contoso.com,Primary`r`n")).Problems[0] | Should -Match 'UserPrincipalName is required \(found: Mail, Location\)'
        (Read-MbtBatchCsv -Path (New-Csv 'empty.csv' '')).Problems[0] | Should -Match 'is empty'
        (Read-MbtBatchCsv -Path (New-Csv 'header.csv' "UserPrincipalName`r`n")).Problems[0] | Should -Match 'has no mailbox'
        (Read-MbtBatchCsv -Path (Join-Path $TestDrive 'csv\nothing-here.csv')).Problems[0] | Should -Match 'was not found'
        (Read-MbtBatchCsv -Path (New-Csv 'twice.csv' "UserPrincipalName,userprincipalname`r`nadele@contoso.com,x`r`n")).Problems[0] | Should -Match 'there twice'
    }
    It 'refuses two rows that write the same PST file, or a part of it, and two rows that use the same store folder' {
        $text = "UserPrincipalName,OutputPst`r`nadele@contoso.com,D:\out\same.pst`r`nalex@contoso.com,D:\OUT\Same.PST`r`nlee@contoso.com,D:\out\same-002.pst`r`nkim@contoso.com,D:\out\other.pst`r`n"
        $r = Read-MbtBatchCsv -Path (New-Csv 'collide.csv' $text)
        $r.Problems.Count | Should -Be 2
        $r.Problems[0] | Should -Match 'line 3: .* is also the file of line 2'
        $r.Problems[1] | Should -Match 'line 4: .*same-002\.pst has the name of a part of the PST file of line 2'
        $stores = Read-MbtBatchCsv -Path (New-Csv 'stores.csv' "UserPrincipalName`r`na_b@contoso.com`r`na:b@contoso.com`r`n") -StoreBase 'D:\stores'
        $stores.Problems.Count | Should -Be 1
        $stores.Problems[0] | Should -Match 'line 3: the store folder .* is also the one of line 2'
    }
    It 'names one PST for each mailbox from -OutputPst: a folder gets one file for each address, a .pst name gets the address added, a row says its own' {
        $csv = New-Csv 'out.csv' "UserPrincipalName,OutputPst`r`nadele@contoso.com,`r`nalex@contoso.com,own.pst`r`nlee@contoso.com,D:\elsewhere`r`n"
        $folder = Read-MbtBatchCsv -Path $csv -OutputPst 'D:\pst'
        @($folder.Rows.Output) | Should -Be @('D:\pst\adele@contoso.com.pst', (Join-Path (Split-Path $csv -Parent) 'own.pst'), 'D:\elsewhere\lee@contoso.com.pst')
        $file = Read-MbtBatchCsv -Path (New-Csv 'out2.csv' "UserPrincipalName`r`nadele@contoso.com`r`nalex@contoso.com`r`n") -OutputPst 'D:\pst\mail.pst'
        @($file.Rows.Output) | Should -Be @('D:\pst\mail-adele@contoso.com.pst', 'D:\pst\mail-alex@contoso.com.pst')
        (Read-MbtBatchCsv -Path (New-Csv 'out3.csv' "UserPrincipalName`r`nadele@contoso.com`r`n")).Rows[0].Output | Should -Be ''
        Resolve-MbtBatchOutput -UserPrincipalName 'a b@contoso.com' -GlobalOutput 'D:\pst' | Should -Be 'D:\pst\a_b@contoso.com.pst'
    }
    It 'with a delegated sign-in only the mailbox of the signed-in account has an archive to export; the application sign-in has no such limit' {
        $csv = New-Csv 'deleg.csv' "UserPrincipalName,Location`r`nadele@contoso.com,Both`r`nalex@contoso.com,Archive`r`nlee@contoso.com,Primary`r`n"
        $r = Read-MbtBatchCsv -Path $csv -SignedInUser 'Adele@contoso.com'
        $r.Problems.Count | Should -Be 1
        $r.Problems[0] | Should -Match 'line 3: alex@contoso.com is not the signed-in account \(Adele@contoso.com\)'
        (Read-MbtBatchCsv -Path $csv -SignedInUser 'adele@contoso.com' -AppOnly).Problems | Should -BeNullOrEmpty
    }
    It 'a SelectionFile is found beside the CSV file when it is relative, and its mailbox is the one it names' {
        $sel = New-Csv 'sel.json' '{}'
        $r = Read-MbtBatchCsv -Path (New-Csv 'withsel.csv' "UserPrincipalName,SelectionFile`r`nadele@contoso.com,sel.json`r`n")
        $r.Problems | Should -BeNullOrEmpty
        $r.Rows[0].SelectionFile | Should -Be $sel
    }
}

Describe 'The arguments of a mailbox' {
    BeforeAll {
        $script:Job = [pscustomobject]@{ Id = 'export-ab12'; Folder = 'D:\w\jobs\j'; EventsPath = 'D:\w\jobs\j\events.jsonl'; StopPath = 'D:\w\jobs\j\stop.request' }
        function New-Row($Upn, $Location = 'Primary', $Selection = '', $Output = '') { [ordered]@{ upn = $Upn; location = $Location; selection = $Selection; output = $Output; store = "D:\s\$Upn" } }
    }
    It 'is an ordinary export with a silent sign-in, its own store, job, events, stop file and batch' {
        $a = Get-MbtBatchChildArgument -Row (New-Row 'adele@contoso.com' 'Both' '' 'D:\pst\adele.pst') -Option @{ AppOnly = $true; SignInUser = ''; Pass = @{ Concurrency = 2; SkipCheck = $true; CertificateThumbprint = 'AB12'; RetryWaitSeconds = 5 } } -Job $script:Job -BatchId 'Wave1'
        ($a -join ' ') | Should -Be '-Mode Export -UserPrincipalName adele@contoso.com -Mailbox All -StorePath D:\s\adele@contoso.com -TempPath D:\w\jobs\j\temp -OutputPst D:\pst\adele.pst -CertificateThumbprint AB12 -Concurrency 2 -RetryWaitSeconds 5 -SkipCheck -NoPrompt -EventsFile D:\w\jobs\j\events.jsonl -StopFile D:\w\jobs\j\stop.request -JobId export-ab12 -BatchId Wave1'
    }
    It 'delegated: the own mailbox as it is, another mailbox through its address (the signed-in account needs access to it); a selection names its mailbox' {
        $opt = @{ AppOnly = $false; SignInUser = 'admin@contoso.com'; Pass = @{} }
        $own = Get-MbtBatchChildArgument -Row (New-Row 'ADMIN@contoso.com' 'Archive') -Option $opt -Job $script:Job -BatchId 'W'
        ($own -join ' ') | Should -Match '^-Mode Export -UserPrincipalName admin@contoso.com -Mailbox Archive '
        $other = Get-MbtBatchChildArgument -Row (New-Row 'alex@contoso.com') -Option $opt -Job $script:Job -BatchId 'W'
        ($other -join ' ') | Should -Match '^-Mode Export -UserPrincipalName admin@contoso.com -MailboxId alex@contoso.com -StorePath '
        ($other -join ' ') | Should -Not -Match '-Mailbox '
        $sel = Get-MbtBatchChildArgument -Row (New-Row 'alex@contoso.com' 'Primary' 'D:\sel.json') -Option @{ AppOnly = $true; SignInUser = ''; Pass = @{} } -Job $script:Job -BatchId 'W'
        ($sel -join ' ') | Should -Match '-UserPrincipalName alex@contoso.com -SelectionFile D:\\sel.json -StorePath'
        ($sel -join ' ') | Should -Not -Match '-Mailbox '
    }
    It 'gives the period of the batch to the export of every mailbox, as it was written' {
        $a = Get-MbtBatchChildArgument -Row (New-Row 'adele@contoso.com') -Option @{ AppOnly = $true; SignInUser = ''; Pass = @{ ArrivedFrom = '2026-03-01'; ArrivedTo = '2026-03-31T18:30' } } -Job $script:Job -BatchId 'W'
        ($a -join ' ') | Should -Match ' -ArrivedFrom 2026-03-01 -ArrivedTo 2026-03-31T18:30 '
        $open = Get-MbtBatchChildArgument -Row (New-Row 'adele@contoso.com') -Option @{ AppOnly = $true; SignInUser = ''; Pass = @{ ArrivedFrom = '2026-03-01'; ArrivedTo = '' } } -Job $script:Job -BatchId 'W'
        ($open -join ' ') | Should -Match ' -ArrivedFrom 2026-03-01 '
        ($open -join ' ') | Should -Not -Match 'ArrivedTo'
        $declared = (Get-Command $script:Entry).Parameters.Keys
        $declared | Should -Contain 'ArrivedFrom'
        $declared | Should -Contain 'ArrivedTo'
    }
    It 'every parameter it uses exists in the entry script' {
        $declared = (Get-Command $script:Entry).Parameters.Keys
        $a = Get-MbtBatchChildArgument -Row (New-Row 'adele@contoso.com' 'Both' 'D:\s.json' 'D:\p.pst') -Option @{ AppOnly = $false; SignInUser = 'admin@contoso.com'; Pass = @{ Concurrency = 2; MaxAttempts = 3; RetryWaitSeconds = 2; NotReadyTrip = 3; BatchItems = 20; BatchMiB = 200; IntervalSeconds = 5; Folder = 'Inbox'; SkipInventory = $true; IncludeRecoverableItems = $true; PstMaxSizeGB = 50; PstMaxSizeMiB = 100; SeparateAuxArchives = $true; SkipCheck = $true; Force = $true; ConfigPath = 'x'; TenantId = 't'; ClientId = 'c'; CertificatePath = 'p' } } -Job $script:Job -BatchId 'W'
        foreach ($t in $a) { if ($t -match '^-([A-Za-z]+)$') { $declared | Should -Contain $Matches[1] -Because $t } }
    }
}

Describe 'Running a batch on stand-in child processes' {
    It 'runs one mailbox at a time by default, each an ordinary job in its own folder, with the batch in its status, a request for the Continue of the console and the arguments of an export' {
        $w = New-Work 'seq'
        $b = New-TestBatch $w 'Seq' @('ok-pause-1@contoso.com', 'ok-pause-2@contoso.com', 'ok-pause-3@contoso.com') -Location @('Archive', 'Both', 'Primary')
        $r = Invoke-StubBatch $w $b
        $r.ExitCode | Should -Be 0
        $r.Counts.Complete | Should -Be 3
        $rows = @($r.Batch.rows)
        @($rows.job | Sort-Object -Unique).Count | Should -Be 3
        Get-MaxOverlap $w $rows | Should -Be 1
        foreach ($row in $rows) {
            $row.state | Should -Be 'complete'
            (Get-JobStatus $w $row).batch | Should -Be 'Seq'
            $row.done | Should -Be 100
            $given = Get-Content -LiteralPath (Get-JobFile $w $row 'args.json') -Raw | ConvertFrom-Json
            $given | Should -Contain '-NoPrompt'
            $given | Should -Contain '-BatchId'
            $given[($given.IndexOf('-JobId') + 1)] | Should -Match '^export-[0-9a-f]{4}$'
            $given[($given.IndexOf('-StorePath') + 1)] | Should -Be (Join-Path (Join-Path $w 'stores') $row.upn)
        }
        (Get-Content -LiteralPath (Get-JobFile $w $rows[0] 'args.json') -Raw | ConvertFrom-Json) | Should -Contain 'Archive'
        (Get-Content -LiteralPath (Get-JobFile $w $rows[1] 'args.json') -Raw | ConvertFrom-Json) | Should -Contain 'All'
        $request = Get-Content -LiteralPath (Get-JobFile $w $rows[0] 'request.json') -Raw | ConvertFrom-Json
        $request.Task | Should -Be 'export'
        $request.Mailbox | Should -Be 'ok-pause-1@contoso.com'
        $request.Dest | Should -Be 'pst'
        $request.StorePath | Should -Be $rows[0].store
        @($request.Commands).Count | Should -Be 1
        @($request.Commands[0]) | Should -Contain '-EventsFile'
        @($request.Commands[0]) | Should -Not -Contain '-LogFolder' -Because 'the console adds the log folder itself'
        (Read-MbtBatch -Folder (Join-Path $w 'batches\Seq')).state | Should -Be 'complete'
        Get-Content -LiteralPath (Join-Path $w 'batches\Seq\batch.log') -Raw | Should -Match 'run ended: 3 complete, 0 not complete, 0 failed, 0 not started; exit 0'
    }
    It 'runs up to the number asked at the same time, never more' {
        $w = New-Work 'par'
        $b = New-TestBatch $w 'Par' @(1..6 | ForEach-Object { "ok-pause-$_@contoso.com" })
        $r = Invoke-StubBatch $w $b -Parallel 2
        $r.ExitCode | Should -Be 0
        $overlap = Get-MaxOverlap $w @($r.Batch.rows)
        $overlap | Should -Be 2
        $b3 = New-TestBatch $w 'Par4' @(1..8 | ForEach-Object { "ok-pause-$_@contoso.com" })
        $r4 = Invoke-StubBatch $w $b3 -Parallel 4
        $r4.ExitCode | Should -Be 0
        Get-MaxOverlap $w @($r4.Batch.rows) | Should -BeIn 3, 4
        { Invoke-MbtBatch -WorkFolder $w -Batch $b -EntryScript $script:Stub -Option @{} -Parallel 5 } | Should -Throw
    }
    It 'continues what is not complete: complete mailboxes are skipped, a partial one goes on in its own job folder, a failed one is tried again; exit 1, then 1, then 0' {
        $w = New-Work 'resume'
        $b = New-TestBatch $w 'Res' @('ok-1@contoso.com', 'partial-2@contoso.com', 'fail-3@contoso.com')
        $first = Invoke-StubBatch $w $b
        $first.ExitCode | Should -Be 1
        @($first.Batch.rows.state) | Should -Be @('complete', 'partial', 'failed')
        (Get-BatchRow $first.Batch 'fail-3@contoso.com').message | Should -Match 'We cannot find this mailbox'
        $jobs = @($first.Batch.rows.job)
        $disk = Read-MbtBatch -Folder (Join-Path $w 'batches\Res')
        @($disk.rows.state) | Should -Be @('complete', 'partial', 'failed')
        $second = Invoke-StubBatch $w $disk
        $second.ExitCode | Should -Be 1 -Because 'the failed mailbox fails again'
        @($second.Batch.rows.state) | Should -Be @('complete', 'complete', 'failed')
        @($second.Batch.rows.job) | Should -Be $jobs -Because 'the same job folders'
        (Get-ChildItem -LiteralPath (Join-Path $w ('jobs\' + $jobs[0])) -Filter 'run-*.txt').Count | Should -Be 1 -Because 'the complete mailbox was not run again'
        (Get-ChildItem -LiteralPath (Join-Path $w ('jobs\' + $jobs[1])) -Filter 'run-*.txt').Count | Should -Be 2
        (Get-ChildItem -LiteralPath (Join-Path $w ('jobs\' + $jobs[2])) -Filter 'run-*.txt').Count | Should -Be 2
        (Get-BatchRow $second.Batch 'partial-2@contoso.com').attempts | Should -Be 2
        $partialOnly = New-TestBatch $w 'Res2' @('ok-1@contoso.com', 'partial-2@contoso.com')
        $p1 = Invoke-StubBatch $w $partialOnly
        $p1.ExitCode | Should -Be 2
        (Invoke-StubBatch $w (Read-MbtBatch -Folder (Join-Path $w 'batches\Res2'))).ExitCode | Should -Be 0
        (Invoke-StubBatch $w (Read-MbtBatch -Folder (Join-Path $w 'batches\Res2'))).Finished | Should -Be 0 -Because 'a complete batch starts nothing'
    }
    It 'a mailbox that was running when the command was killed is continued like a partial one' {
        $w = New-Work 'killed'
        $b = New-TestBatch $w 'Kill' @('ok-1@contoso.com', 'ok-2@contoso.com')
        $b.rows[0].state = 'running'
        $b.rows[0].job = '20261114-0900-export-aaaa'
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $w 'jobs\20261114-0900-export-aaaa')
        $r = Invoke-StubBatch $w $b
        $r.ExitCode | Should -Be 0
        $r.Batch.rows[0].job | Should -Be '20261114-0900-export-aaaa'
        Test-Path (Join-Path $w 'jobs\20261114-0900-export-aaaa\run-1.txt') | Should -BeTrue
    }
    It 'stops on a stop request in the batch folder: the running mailboxes stop the way the console stops them, nothing else starts, and the batch continues later' {
        $w = New-Work 'stop'
        $b = New-TestBatch $w 'Stop' @('asker-1@contoso.com', 'slow-quick-2@contoso.com', 'ok-3@contoso.com', 'ok-4@contoso.com')
        $env:MBT_STUB_STOPFILE = Join-Path $w 'batches\Stop\stop.request'
        try { $r = Invoke-StubBatch $w $b -Parallel 2 } finally { Remove-Item Env:\MBT_STUB_STOPFILE -ErrorAction SilentlyContinue }
        $r.Stopped | Should -BeTrue
        $r.ExitCode | Should -Be 2
        @($r.Batch.rows.state) | Should -Be @('complete', 'partial', 'pending', 'pending')
        $slow = Get-BatchRow $r.Batch 'slow-quick-2@contoso.com'
        $slow.message | Should -Match 'stopped'
        (Get-JobStatus $w $slow).state | Should -Be 'stopped'
        Test-Path (Get-JobFile $w $slow 'stop.request') | Should -BeTrue
        (Get-Content -LiteralPath (Join-Path $w 'batches\Stop\report.html') -Raw) | Should -Match '2 have not started yet'
        # the same command again: the stop request is gone, the stopped mailbox goes on in its own job folder, the others run
        $again = Invoke-StubBatch $w (Read-MbtBatch -Folder (Join-Path $w 'batches\Stop')) -Parallel 2
        $again.Stopped | Should -BeFalse
        $again.ExitCode | Should -Be 0
        @($again.Batch.rows.state) | Should -Be @('complete', 'complete', 'complete', 'complete')
        (Get-BatchRow $again.Batch 'slow-quick-2@contoso.com').job | Should -Be $slow.job
        (Get-ChildItem -LiteralPath (Join-Path $w ('jobs\' + $slow.job)) -Filter 'run-*.txt').Count | Should -Be 2
        Test-Path (Get-JobFile $w $slow 'stop.request') | Should -BeFalse -Because 'the stop request of the first run is removed before the job goes on'
    }
    It 'stops on -StopFile too, and run again the batch goes on from where it was' {
        $w = New-Work 'stopfile'
        $b = New-TestBatch $w 'StopFile' @('asker-1@contoso.com', 'slow-quick-2@contoso.com', 'ok-3@contoso.com')
        $stopFile = Join-Path $w 'stop-it'
        $env:MBT_STUB_STOPFILE = $stopFile
        try { $r = Invoke-StubBatch $w $b -Parallel 2 -StopFile $stopFile } finally { Remove-Item Env:\MBT_STUB_STOPFILE -ErrorAction SilentlyContinue }
        $r.Stopped | Should -BeTrue
        (Get-BatchRow $r.Batch 'ok-3@contoso.com').state | Should -Be 'pending'
        (Get-BatchRow $r.Batch 'asker-1@contoso.com').state | Should -Be 'complete'
        Remove-Item -LiteralPath $stopFile
        $again = Invoke-StubBatch $w (Read-MbtBatch -Folder (Join-Path $w 'batches\StopFile')) -Parallel 2
        $again.ExitCode | Should -Be 0
        @($again.Batch.rows.state) | Should -Be @('complete', 'complete', 'complete')
        (Get-ChildItem -LiteralPath (Join-Path $w ('jobs\' + (Get-BatchRow $again.Batch 'asker-1@contoso.com').job)) -Filter 'run-*.txt').Count | Should -Be 1
    }
    It 'a mailbox that needs a sign-in is not complete (it can be continued after signing in) and the batch goes on (application); with a delegated sign-in three in a row end the run and leave the others for the next one' {
        $w = New-Work 'signin'
        $b = New-TestBatch $w 'App' @('signin-1@contoso.com', 'ok-2@contoso.com', 'signin-3@contoso.com')
        $r = Invoke-StubBatch $w $b
        @($r.Batch.rows.state) | Should -Be @('partial', 'complete', 'partial')
        (Get-BatchRow $r.Batch 'signin-1@contoso.com').message | Should -Match '^sign-in required: .*run it again'
        (Get-BatchRow $r.Batch 'signin-1@contoso.com').reason | Should -Be 'signin-required'
        $r.ExitCode | Should -Be 2
        $d = New-TestBatch $w 'Deleg' @('signin-1@contoso.com', 'signin-2@contoso.com', 'signin-3@contoso.com', 'ok-4@contoso.com', 'ok-5@contoso.com') -Option @{ AppOnly = $false; SignInUser = 'admin@contoso.com'; Pass = @{} }
        $rd = Invoke-StubBatch $w $d -Option @{ AppOnly = $false; SignInUser = 'admin@contoso.com'; Pass = @{} }
        @($rd.Batch.rows.state) | Should -Be @('partial', 'partial', 'partial', 'pending', 'pending')
        $rd.Abandoned | Should -BeTrue
        $rd.ExitCode | Should -Be 2
        $given = Get-Content -LiteralPath (Get-JobFile $w $rd.Batch.rows[0] 'args.json') -Raw | ConvertFrom-Json
        $given[($given.IndexOf('-MailboxId') + 1)] | Should -Be 'signin-1@contoso.com'
        $given[($given.IndexOf('-UserPrincipalName') + 1)] | Should -Be 'admin@contoso.com'
    }
    It 'a mailbox whose process ends without a word is failed with the exit code, and its job says so in the overview' {
        $w = New-Work 'crash'
        $b = New-TestBatch $w 'Crash' @('crash-1@contoso.com', 'ok-2@contoso.com')
        $r = Invoke-StubBatch $w $b
        @($r.Batch.rows.state) | Should -Be @('failed', 'complete')
        $r.Batch.rows[0].message | Should -Match 'ended with exit code 1'
        $s = Get-JobStatus $w $r.Batch.rows[0]
        $s.state | Should -Be 'error'
        $s.batch | Should -Be 'Crash'
        $s.reason | Should -Be 'start-failed'
    }
    It 'refuses to run a batch that another command is running' {
        $w = New-Work 'locked'
        $b = New-TestBatch $w 'Lock' @('ok-1@contoso.com')
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $w 'batches\Lock')
        $held = [IO.File]::Open((Join-Path $w 'batches\Lock\batch.lock'), 'OpenOrCreate', 'ReadWrite', 'None')
        try { { Invoke-StubBatch $w $b } | Should -Throw '*already running*' } finally { $held.Dispose() }
        (Invoke-StubBatch $w $b).ExitCode | Should -Be 0
    }
    It 'writes the report of the batch after each mailbox: the jobs of this batch only, links to their reports, the mailboxes not started said above the table' {
        $w = New-Work 'report'
        $foreign = Join-Path $w 'jobs\20261101-1000-export-ffff'
        $null = New-Item -ItemType Directory -Force -Path $foreign
        Set-Content -LiteralPath (Join-Path $foreign 'status.json') -Value '{"v":1,"jobId":"export-ffff","task":"export","title":"Export - other@contoso.com","state":"complete","exitCode":0,"batch":"Other","started":"2026-11-01T10:00:00.000Z","updated":"2026-11-01T10:05:00.000Z"}'
        $b = New-TestBatch $w 'Rep' @('ok-1@contoso.com', 'fail-2@contoso.com')
        $r = Invoke-StubBatch $w $b
        $r.Report | Should -Be (Join-Path $w 'batches\Rep\report.html')
        $html = Get-Content -LiteralPath $r.Report -Raw
        $html | Should -Match '<title>Mailbox Transfer: jobs of batch Rep</title>'
        $html | Should -Match '"base":"\.\./\.\./jobs/"'
        $html | Should -Match '2 mailboxes in the CSV file \(1 complete, 0 not complete, 1 failed\)\.'
        $html | Should -Match 'ok-1@contoso\.com'
        $html | Should -Not -Match 'other@contoso.com'
        foreach ($row in $r.Batch.rows) { $html | Should -Match ([regex]::Escape($row.job)) }
        (Get-Content -LiteralPath (Join-Path $w 'jobs\overview.html') -Raw -ErrorAction SilentlyContinue) | Should -BeNullOrEmpty -Because 'the overview of all the jobs is the command line''s to write, once, at the end'
    }
    It 'counts a mailbox the same way in the caption, the tiles, the chips, the row, the overview of all the jobs, the summary and the exit code: a sign-in needed is not complete, not failed' {
        $w = New-Work 'agree'
        $b = New-TestBatch $w 'Agree' @('ok-1@contoso.com', 'fail-2@contoso.com', 'partial-3@contoso.com', 'signin-4@contoso.com', 'ok-5@contoso.com')
        $r = Invoke-StubBatch $w $b
        $r.ExitCode | Should -Be 1
        @($r.Batch.rows.state) | Should -Be @('complete', 'failed', 'partial', 'partial', 'complete')
        $html = Get-Content -LiteralPath $r.Report -Raw
        $html | Should -Match '5 mailboxes in the CSV file \(2 complete, 2 not complete, 1 failed\)\.'
        $html | Should -Match 'Complete</div><div class="tile-value">2</div>'
        $html | Should -Match 'Failed</div><div class="tile-value">1</div>'
        $html | Should -Match 'Need a look</div><div class="tile-value">3</div>'
        $data = [Text.Json.JsonDocument]::Parse([regex]::Match($html, '<script type="application/json" id="mbt-overview-data">(.*?)</script>', 'Singleline').Groups[1].Value)
        $codes = @($data.RootElement.GetProperty('rows').EnumerateArray() | ForEach-Object { $_[7].GetInt32() })
        @($codes | Where-Object { $_ -eq 0 }).Count | Should -Be 2
        @($codes | Where-Object { $_ -eq 1 }).Count | Should -Be 1
        @($codes | Where-Object { $_ -eq 2 }).Count | Should -Be 2 -Because 'the mailbox that needs a sign-in is Not complete in the page, as in the caption'
        $all = Get-Content -LiteralPath (Write-MbtJobOverview -WorkFolder $w) -Raw
        $all | Should -Match 'Failed</div><div class="tile-value">1</div>'
        $all | Should -Match 'Need a look</div><div class="tile-value">3</div>'
        $text = (& { Write-MbtBatchSummary -ExitCode $r.ExitCode -Batch $r.Batch -Folder $r.Folder -Report $r.Report } *>&1 | Out-String)
        $text | Should -Match 'Mailboxes\s+5 in Agree: 2 complete, 2 not complete, 1 failed, 0 not started'
    }
    It 'a mailbox that could not be started is failed in its job as well, not left running' {
        $w = New-Work 'nostart'
        $b = New-TestBatch $w 'NoStart' @('ok-1@contoso.com')
        $a = @{ WorkFolder = $w; Batch = $b; EntryScript = $script:Stub; Option = @{ AppOnly = $true; SignInUser = ''; Pass = @{} }; PollMilliseconds = 50; StartGapMilliseconds = 0; Launcher = { throw 'no process' } }
        $r = Invoke-MbtBatch @a 6>&1 | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.InformationRecord] }
        $r.Batch.rows[0].state | Should -Be 'failed'
        $s = Get-JobStatus $w $r.Batch.rows[0]
        $s.state | Should -Be 'error'
        $s.batch | Should -Be 'NoStart'
        $r.ExitCode | Should -Be 1
    }
    It 'every link of the report of a batch and of the overview of all the jobs leads to a file: the report of the job, else its CSV, else its status' {
        $w = New-Work 'links'
        $b = New-TestBatch $w 'Links' @('ok-1@contoso.com', 'fail-2@contoso.com', 'ok-3@contoso.com')
        $r = Invoke-StubBatch $w $b
        $jobs = Join-Path $w 'jobs'
        Set-Content -LiteralPath (Join-Path $jobs ($r.Batch.rows[0].job + '\report.html')) -Value '<html></html>'
        Set-Content -LiteralPath (Join-Path $jobs ($r.Batch.rows[1].job + '\report.csv')) -Value 'a,b'
        $page = Write-MbtBatchReport -WorkFolder $w -Batch $r.Batch -Folder $r.Folder
        $all = Write-MbtJobOverview -WorkFolder $w
        $want = @{ ($r.Batch.rows[0].job) = 'report.html'; ($r.Batch.rows[1].job) = 'report.csv'; ($r.Batch.rows[2].job) = 'status.json' }
        foreach ($file in $page, $all) {
            $html = Get-Content -LiteralPath $file -Raw
            $d = [Text.Json.JsonDocument]::Parse([regex]::Match($html, '<script type="application/json" id="mbt-overview-data">(.*?)</script>', 'Singleline').Groups[1].Value).RootElement
            $base = $d.GetProperty('base').GetString()
            $columns = @($d.GetProperty('columns').EnumerateArray() | ForEach-Object { $_.GetString() })
            $rows = @($d.GetProperty('rows').EnumerateArray())
            $rows.Count | Should -Be 3
            foreach ($row in $rows) {
                $id = $row[$columns.IndexOf('id')].GetString()
                $kind = $row[$columns.IndexOf('link')].GetInt32()
                $name = switch ($kind) { 1 { 'report.html' } 2 { 'report.csv' } 3 { 'status.json' } default { '' } }
                $name | Should -Be $want[$id]
                [IO.Path]::GetFullPath([IO.Path]::Combine((Split-Path $file -Parent), $base.Replace('/', '\'), $id, $name)) | Should -Exist
            }
        }
    }
    It 'one failed mailbox does not stop the others, and the command line summary is the house card with the numbers, the report and the next steps' {
        $w = New-Work 'summary'
        $b = New-TestBatch $w 'Sum' @('ok-1@contoso.com', 'fail-2@contoso.com', 'partial-3@contoso.com')
        $r = Invoke-StubBatch $w $b
        $text = (& { Write-MbtBatchSummary -ExitCode $r.ExitCode -Batch $r.Batch -Folder $r.Folder -Seconds 75 -Report $r.Report -Command '.\Invoke-MailboxTransfer.ps1 -Mode Export -ResumeBatch Sum' } *>&1 | Out-String)
        $text | Should -Match 'Export batch finished with failures'
        $text | Should -Match 'exit code 1 - failed'
        $text | Should -Match 'Mailboxes\s+3 in Sum: 1 complete, 1 not complete, 1 failed, 0 not started'
        $text | Should -Match 'Items\s+103 exported, 1 failed'
        $text | Should -Match 'Report\s+.*report\.html'
        $text | Should -Match 'Next steps\s+Run the same command again to continue'
        $text | Should -Match '-ResumeBatch Sum'
        $r.Batch.rows[1].state | Should -Be 'failed'
    }
}

Describe 'The command line of a batch' {
    BeforeAll {
        function Invoke-Cli {
            param([hashtable]$Parameters)
            $global:LASTEXITCODE = $null
            $text = & $script:Entry @Parameters *>&1 | Out-String
            $global:CliExit = $LASTEXITCODE
            return $text
        }
        $script:Work = New-Work 'cli'
        $script:Common = @{ Mode = 'Export'; AppOnly = $true; CertificateThumbprint = 'ABCDEF0123456789ABCDEF0123456789ABCDEF01'; TenantId = $script:Tenant; ClientId = '66666666-7777-8888-9999-000000000000'; WorkFolder = $script:Work; LogFolder = (Join-Path $script:Work 'logs') }
    }
    AfterAll { Remove-Variable CliExit -Scope Global -ErrorAction SilentlyContinue }
    It 'prints every problem of the CSV file in the error card, exits 1 and starts nothing' {
        $csv = New-Csv 'cli-bad.csv' "UserPrincipalName,Location`r`n,Primary`r`nadele@contoso.com,Primary`r`nalex@contoso.com,Nowhere`r`nADELE@contoso.com,Primary`r`n"
        $text = Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; BatchName = 'Bad' })
        $global:CliExit | Should -Be 1
        $text | Should -Match 'The batch cannot start'
        $text | Should -Match 'line 2: UserPrincipalName is empty'
        $text | Should -Match "line 4: Location 'Nowhere' is not Primary, Archive or Both"
        $text | Should -Match 'line 5: ADELE@contoso.com is already on line 3'
        $text | Should -Match 'Export batch not started'
        $text | Should -Match '3 problems, listed above: nothing was started'
        Test-Path (Join-Path $script:Work 'batches') | Should -BeFalse
    }
    It '-WhatIf checks the file and lists the mailboxes with their files, and writes and starts nothing' {
        $csv = New-Csv 'cli-ok.csv' "UserPrincipalName,Location`r`nadele@contoso.com,Both`r`nalex@contoso.com,`r`n"
        $text = Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; BatchName = 'Plan'; OutputPst = (Join-Path $script:Work 'pst'); WhatIf = $true; ParallelMailboxes = 2 })
        $global:CliExit | Should -Be 0
        $text | Should -Match '2 mailboxes: 0 complete, 2 to do; 2 at a time'
        $text | Should -Match 'adele@contoso.com - Both - PST .*adele@contoso.com\.pst'
        $text | Should -Match 'alex@contoso.com - Primary'
        $text | Should -Match 'WhatIf: the file is valid, nothing was started'
        $text | Should -Match 'Export batch planned'
        Test-Path (Join-Path $script:Work 'batches') | Should -BeFalse
        Test-Path (Join-Path $script:Work 'pst') | Should -BeFalse
    }
    It 'the same command again on a batch that is complete starts nothing and signs in nowhere; the name is found in any case; a CSV that is not the one of the batch is refused with what differs' {
        $w2 = New-Work 'cli-done'
        $p = $script:Common.Clone(); $p.WorkFolder = $w2; $p.LogFolder = (Join-Path $w2 'logs')
        $b = New-TestBatch $w2 'Done1' @('ok-1@contoso.com', 'ok-2@contoso.com') -Location @('Both', 'Archive')
        foreach ($row in $b.rows) { $row.state = 'complete'; $row.job = ''; $row.done = 5 }
        Save-MbtBatch -Batch $b -Folder (Join-Path $w2 'batches\Done1')
        $text = Invoke-Cli ($p + @{ MailboxCsv = $b.csv; BatchName = 'done1' })
        $global:CliExit | Should -Be 0
        $text | Should -Match 'every mailbox of this batch is already complete: nothing to do'
        $text | Should -Match 'Export batch finished'
        $text | Should -Match 'Mailboxes\s+2 in Done1: 2 complete, 0 not complete, 0 failed, 0 not started'
        Test-Path (Join-Path $w2 'batches\Done1\report.html') | Should -BeTrue
        $text | Should -Not -Match 'signing in'
        $again = Invoke-Cli ($p + @{ ResumeBatch = 'DONE1' })
        $global:CliExit | Should -Be 0
        $again | Should -Match '2 mailboxes: 2 complete, 0 to do'
        $other = New-Csv 'done1-other.csv' "UserPrincipalName,Location`r`nok-1@contoso.com,Both`r`nok-3@contoso.com,Primary`r`n"
        $refused = Invoke-Cli ($p + @{ MailboxCsv = $other; BatchName = 'Done1' })
        $global:CliExit | Should -Be 1
        $refused | Should -Match 'The CSV file is not the list of the batch Done1: ok-3@contoso.com \(Primary\) is in the CSV file and not in the batch'
        $refused | Should -Match 'ok-2@contoso.com \(Archive\) is in the batch and not in the CSV file'
        $refused | Should -Match 'use -ResumeBatch to continue it, or another -BatchName'
    }
    It 'a batch keeps the period it was started with: the same period is accepted, another one is refused before anything starts, a date that cannot be read stops the command' {
        $w3 = New-Work 'cli-period'
        $p = $script:Common.Clone(); $p.WorkFolder = $w3; $p.LogFolder = (Join-Path $w3 'logs')
        $option = @{ AppOnly = $true; SignInUser = ''; Pass = @{ ArrivedFrom = '2026-03-01'; ArrivedTo = '2026-03-31' } }
        $b = New-TestBatch $w3 'Per1' @('ok-1@contoso.com', 'ok-2@contoso.com') -Option $option
        foreach ($row in $b.rows) { $row.state = 'complete'; $row.job = ''; $row.done = 5 }
        Save-MbtBatch -Batch $b -Folder (Join-Path $w3 'batches\Per1')
        $same = Invoke-Cli ($p + @{ ResumeBatch = 'Per1'; ArrivedFrom = '2026-03-01T00:00:00'; ArrivedTo = '2026-03-31T23:59:59.9999999' })
        $global:CliExit | Should -Be 0
        $same | Should -Match '2 mailboxes: 2 complete, 0 to do'
        $refused = Invoke-Cli ($p + @{ ResumeBatch = 'Per1'; ArrivedTo = '2026-06-30' })
        $global:CliExit | Should -Be 1
        $refused | Should -Match 'was started with the period'
        $refused | Should -Match 'Continuing it would mix two periods in the stores of the batch'
        $refused | Should -Match 'nothing was started'
        { Invoke-Cli ($p + @{ ResumeBatch = 'Per1'; ArrivedFrom = '2026-04-01'; ArrivedTo = '2026-03-01' }) } | Should -Throw '*starts after it ends*'
        { Invoke-Cli ($p + @{ ResumeBatch = 'Per1'; ArrivedFrom = '1/2/2026' }) } | Should -Throw '*not a date this tool reads*'
        (Get-Content -LiteralPath (Join-Path $w3 'batches\Per1\batch.json') -Raw) | Should -Match '"ArrivedFrom": "2026-03-01'
    }
    It 'says what is wrong with the batch name, a batch that does not exist, a CSV that is not the one of the batch, and options that belong to one mailbox' {
        $csv = New-Csv 'cli-name.csv' "UserPrincipalName`r`nadele@contoso.com`r`n"
        (Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; BatchName = '..\x' })) | Should -Match 'The batch name can hold letters'
        $global:CliExit | Should -Be 1
        (Invoke-Cli ($script:Common + @{ ResumeBatch = 'Nope' })) | Should -Match 'There is no batch named Nope'
        $global:CliExit | Should -Be 1
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; ParallelMailboxes = 9 }) } | Should -Throw
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; SelectionFile = 'x.json' }) } | Should -Throw '*-SelectionFile is for one mailbox*'
        { Invoke-Cli (@{ MailboxCsv = $csv; Mode = 'Import' }) } | Should -Throw '*belong to -Mode Export*'
        { Invoke-Cli @{ Mode = 'Export'; BatchName = 'X' } } | Should -Throw '*belong to -MailboxCsv*'
        { Invoke-Cli ($script:Common + @{ MailboxCsv = $csv; Mailbox = 'AuxArchive' }) } | Should -Throw '*AuxArchive*'
    }
}

Describe 'What the retention of the job folders keeps of a batch' {
    BeforeAll {
        function New-BatchJob([string]$Work, [string]$Name, [double]$DaysAgo, [string]$Batch) {
            $dir = Join-Path $Work "jobs\$Name"
            $null = New-Item -ItemType Directory -Force -Path $dir
            $at = [datetime]::UtcNow.AddDays(-$DaysAgo)
            $stamp = { param($d) $d.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z') }
            $s = [ordered]@{ v = 1; jobId = $Name; task = 'export'; title = 'Export - x'; mode = 'Export'; batch = $Batch; state = 'complete'; exitCode = 0; counts = @{ done = 1 }; started = (& $stamp $at.AddMinutes(-5)); updated = (& $stamp $at); pid = 999999 }
            [IO.File]::WriteAllText((Join-Path $dir 'status.json'), ($s | ConvertTo-Json -Depth 4))
            (Get-Item -LiteralPath $dir).LastWriteTimeUtc = $at
        }
        function New-BatchFolder([string]$Work, [string]$Id, [string[]]$State, [string[]]$Job) {
            $rows = @(for ($i = 0; $i -lt $State.Count; $i++) { [ordered]@{ n = $i + 1; upn = "u$i@contoso.com"; location = 'Primary'; selection = ''; output = ''; store = ''; state = $State[$i]; job = $Job[$i]; attempts = 1 } })
            $b = [ordered]@{ v = 1; id = $Id; name = $Id; created = '2026-01-01T00:00:00.000Z'; updated = '2026-01-01T00:00:00.000Z'; state = ''; mode = 'Export'; csv = ''; parallel = 1; options = @{}; counts = $null; rows = $rows }
            Save-MbtBatch -Batch $b -Folder (Join-Path $Work "batches\$Id")
        }
    }
    It 'never removes a job of a batch that is not finished, however old and complete; the jobs of a complete batch go like any other' {
        $w = New-Work 'keepbatch'
        New-BatchJob $w 'old-in-open' 200 'Open'
        New-BatchJob $w 'old-in-open-nostatusfield' 200 ''
        New-BatchJob $w 'old-in-done' 200 'Done'
        New-BatchJob $w 'old-alone' 200 ''
        New-BatchFolder $w 'Open' @('complete', 'failed') @('old-in-open', 'later-job')
        New-BatchFolder $w 'Done' @('complete', 'complete') @('old-in-done', 'other-job')
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 90
        @($r.Removed | Sort-Object) | Should -Be @('old-alone', 'old-in-done', 'old-in-open-nostatusfield')
        (Get-ChildItem -LiteralPath (Join-Path $w 'jobs') -Directory).Name | Should -Be @('old-in-open')
        @($r.Protected) | Should -Contain 'old-in-open'
    }
    It 'keeps the jobs of a batch whose batch.json cannot be read, and tells a batch that is finished from one that is not' {
        $w = New-Work 'keepbatch2'
        New-BatchJob $w 'job-a' 200 'Broken'
        New-BatchJob $w 'job-b' 200 'Waiting'
        New-BatchJob $w 'job-c' 200 'Finished'
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $w 'batches\Broken')
        Set-Content -LiteralPath (Join-Path $w 'batches\Broken\batch.json') -Value 'not json'
        New-BatchFolder $w 'Waiting' @('complete', 'pending') @('job-b', '')
        New-BatchFolder $w 'Finished' @('complete') @('job-c')
        $p = Get-MbtBatchProtection -WorkFolder $w
        $p.Batches.Contains('broken') | Should -BeTrue
        $p.Batches.Contains('WAITING') | Should -BeTrue
        $p.Batches.Contains('Finished') | Should -BeFalse
        $p.Jobs.Contains('JOB-B') | Should -BeTrue
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 90
        @($r.Removed) | Should -Be @('job-c')
    }
}

Describe 'The batches of a work folder' {
    It 'lists them newest first with their numbers and report, and the names of their jobs' {
        $w = New-Work 'list'
        $a = New-TestBatch $w 'First' @('ok-1@contoso.com', 'ok-2@contoso.com')
        $null = Invoke-StubBatch $w $a
        Start-Sleep -Milliseconds 50
        $b = New-TestBatch $w 'Second' @('fail-1@contoso.com', 'ok-2@contoso.com', 'ok-3@contoso.com')
        $rb = Invoke-StubBatch $w $b
        $list = @(Get-MbtBatchList -WorkFolder $w)
        @($list.Id) | Should -Be @('Second', 'First')
        $list[0].State | Should -Be 'failed'
        $list[0].Total | Should -Be 3; $list[0].Complete | Should -Be 2; $list[0].Failed | Should -Be 1
        $list[0].Report | Should -Be (Join-Path $w 'batches\Second\report.html')
        @($list[0].Jobs).Count | Should -Be 3
        $list[1].State | Should -Be 'complete'
        (Get-MbtBatchList -WorkFolder (Join-Path $TestDrive 'no-batches')) | Should -BeNullOrEmpty
        # a job of the console's Continue list: the batch is on it, and the jobs that can be left out are left out
        $jobs = @(Get-MbtJob -WorkFolder $w -Limit 50 -ExcludeFolder @($list[1].Jobs))
        $jobs.Count | Should -Be 3
        @($jobs.Batch | Sort-Object -Unique) | Should -Be @('Second')
        $rb.Counts.Total | Should -Be 3
    }
}

Describe 'A batch of real exports: the entry script run against the fake service for each mailbox' {
    BeforeAll {
        function New-Message([string]$Subject) {
            $b = [MailboxTransfer.Tests.FxBuilder]::new()
            $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $Subject)
            return $b.ToArray()
        }
        function New-SourceService {
            $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
            $null = $svc.AddMailbox($script:Src)
            $null = $svc.AddFolder($script:Src, 'F-INBOX', 'Inbox')
            foreach ($i in 1..4) { $null = $svc.AddItem($script:Src, 'F-INBOX', "I$i", (New-Message "Inbox $i"), "ck-I$i") }
            return $svc
        }
        $script:Work = New-Work 'real'
        # the launcher runs the entry script itself, in this process, with the arguments the batch gives; the mailbox address is resolved to the id of the fake mailbox
        $script:Launcher = {
            param($Argument, $Job, $Batch)
            $named = @{}
            for ($i = 0; $i -lt $Argument.Count; $i++) {
                $t = [string]$Argument[$i]
                if ($t -notmatch '^-([A-Za-z]+)$') { continue }
                $key = $t.Substring(1)
                if (($i + 1) -lt $Argument.Count -and [string]$Argument[$i + 1] -notmatch '^-[A-Za-z]+$') { $named[$key] = [string]$Argument[$i + 1]; $i++ } else { $named[$key] = $true }
            }
            $named.Remove('MailboxId'); $named.PrimaryMailboxId = $script:Src; $named.ArchiveMailboxId = $script:Src
            $named.DevelopmentMode = $true
            $named.LogFolder = Join-Path $script:Work 'logs'
            $global:MbtEntryClient = New-FakeClient -Service $global:MbtBatchService
            $global:LASTEXITCODE = $null
            $null = & $script:Entry @named *>&1 | Out-String
            [pscustomobject]@{ Process = [pscustomobject]@{ HasExited = $true; ExitCode = $global:LASTEXITCODE }; Pipes = @() }
        }
        function New-FakeClient {
            param([Parameter(Mandatory)]$Service)
            $o = [MailboxTransfer.GraphOptions]::new()
            $o.DelayScale = 0
            $o.UseJitter = $false
            return [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
        }
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
    It 'applies the period to every mailbox of the batch and says how many items each one left out' {
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Src)
        $null = $svc.AddFolder($script:Src, 'F-INBOX', 'Inbox')
        foreach ($i in 1..2) { $null = $svc.AddItem($script:Src, 'F-INBOX', "O$i", (New-Message "Old $i"), "ck-O$i", [datetime]::new(2025, 5, $i, 12, 0, 0, [DateTimeKind]::Utc)) }
        foreach ($i in 1..3) { $null = $svc.AddItem($script:Src, 'F-INBOX', "M$i", (New-Message "March $i"), "ck-M$i", [datetime]::new(2026, 3, 10 + $i, 12, 0, 0, [DateTimeKind]::Utc)) }
        $global:MbtBatchService = $svc
        $work = New-Work 'real-period'
        $csv = New-Csv 'real-period.csv' "UserPrincipalName,Location`r`nadele.vance@contoso.com,Primary`r`nalex.wilber@contoso.com,Primary`r`n"
        $read = Read-MbtBatchCsv -Path $csv -StoreBase (Join-Path $work 'stores') -SignedInUser 'adele.vance@contoso.com'
        $option = @{ AppOnly = $false; SignInUser = 'adele.vance@contoso.com'; Pass = @{ Concurrency = 2; IntervalSeconds = 60; ArrivedFrom = '2026-03-01'; ArrivedTo = '2026-03-31' } }
        $batch = New-MbtBatch -Id 'Period' -Csv $csv -Row $read.Rows -Option $option -Parallel 1
        $r = Invoke-MbtBatch -WorkFolder $work -Batch $batch -EntryScript $script:Entry -Option $option -Launcher $script:Launcher -PollMilliseconds 50 -StartGapMilliseconds 0 6>&1 | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.InformationRecord] }
        $r.ExitCode | Should -Be 0 -Because (@($r.Batch.rows.message) -join ' | ')
        foreach ($row in $r.Batch.rows) {
            $row.state | Should -Be 'complete'
            $row.done | Should -Be 3
            $row.outsidePeriod | Should -Be 2
            $dir = Join-Path $work "jobs\$($row.job)"
            $status = Get-Content -LiteralPath (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json
            $status.period.text | Should -Be 'arrived between 2026-03-01 and 2026-03-31'
            $status.counts.outsidePeriod | Should -Be 2
            $status.counts.kept | Should -Be 3
        }
        (Get-MbtBatchCount $r.Batch).OutsidePeriod | Should -Be 4
    }
    It 'exports each mailbox as an ordinary job: events, status with the batch, report.html, a store, a PST named after the mailbox; one mailbox after the other; the batch report lists them' {
        $global:MbtBatchService = New-SourceService
        $csv = New-Csv 'real.csv' "UserPrincipalName,Location`r`nadele.vance@contoso.com,Primary`r`nalex.wilber@contoso.com,Primary`r`n"
        $read = Read-MbtBatchCsv -Path $csv -StoreBase (Join-Path $script:Work 'stores') -OutputPst (Join-Path $script:Work 'pst') -SignedInUser 'adele.vance@contoso.com'
        $read.Problems | Should -BeNullOrEmpty
        $option = @{ AppOnly = $false; SignInUser = 'adele.vance@contoso.com'; Pass = @{ Concurrency = 2; IntervalSeconds = 60 } }
        $batch = New-MbtBatch -Id 'Real' -Csv $csv -Row $read.Rows -Option $option -Parallel 1
        $r = Invoke-MbtBatch -WorkFolder $script:Work -Batch $batch -EntryScript $script:Entry -Option $option -Launcher $script:Launcher -PollMilliseconds 50 -StartGapMilliseconds 0 6>&1 | Where-Object { $_ -isnot [string] -and $_ -isnot [System.Management.Automation.InformationRecord] }
        $r.ExitCode | Should -Be 0 -Because (@($r.Batch.rows.message) -join ' | ')
        @($r.Batch.rows.state) | Should -Be @('complete', 'complete')
        foreach ($row in $r.Batch.rows) {
            $dir = Join-Path $script:Work "jobs\$($row.job)"
            (Get-Content -LiteralPath (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json).batch | Should -Be 'Real'
            (Get-Content -LiteralPath (Join-Path $dir 'status.json') -Raw | ConvertFrom-Json).state | Should -Be 'complete'
            Test-Path (Join-Path $dir 'events.jsonl') | Should -BeTrue
            Test-Path (Join-Path $dir 'report.html') | Should -BeTrue
            Test-Path (Join-Path $dir 'request.json') | Should -BeTrue
            Test-Path (Join-Path $row.store 'catalog.db') | Should -BeTrue
            Test-Path $row.output | Should -BeTrue
            $row.done | Should -Be 4
        }
        @($r.Batch.rows.output | ForEach-Object { Split-Path $_ -Leaf }) | Should -Be @('adele.vance@contoso.com.pst', 'alex.wilber@contoso.com.pst')
        $html = Get-Content -LiteralPath (Join-Path $script:Work 'batches\Real\report.html') -Raw
        $html | Should -Match '2 mailboxes in the CSV file \(2 complete, 0 not complete, 0 failed\)\.'
        Test-Path (Join-Path $script:Work 'jobs\overview.html') | Should -BeFalse -Because 'a job of a batch leaves the overview of all the jobs to the batch'
    }
}
