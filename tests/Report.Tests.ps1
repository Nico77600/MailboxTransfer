#Requires -Version 7.6
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the HTML report of a job (module\MailboxTransfer.Report.ps1, templates\Report.template.html) and of the way the
    console picks the report to open (Get-MbtGuiReportPath, Write-MbtGuiRunReport). The jobs are synthetic job folders (events.jsonl and
    status.json written here); nothing reaches the network.

    Run:  Invoke-Pester -Path .\tests\Report.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:Root = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    foreach ($f in 'Core', 'Run', 'Jobs') { . (Join-Path $script:Root "module\MailboxTransfer.Gui.$f.ps1") }
    $script:T0 = [datetime]::new(2026, 11, 9, 14, 0, 0, [DateTimeKind]::Utc)

    function New-JobFolder([string]$Name) {
        $dir = Join-Path $TestDrive $Name
        $null = New-Item -ItemType Directory -Force -Path $dir
        return $dir
    }

    # Writes events.jsonl: each event is @{ at = seconds after the start; e = kind; ... }
    function Write-Events([string]$Dir, [object[]]$Events, [string[]]$Raw = @()) {
        $lines = foreach ($ev in $Events) {
            $o = [ordered]@{ v = 1; t = $script:T0.AddSeconds([double]$ev.at).ToString('yyyy-MM-dd\THH:mm:ss.fff\Z'); e = $ev.e }
            foreach ($k in $ev.Keys) { if ($k -notin 'at', 'e') { $o[$k] = $ev[$k] } }
            $o | ConvertTo-Json -Compress -Depth 6
        }
        Set-Content -LiteralPath (Join-Path $Dir 'events.jsonl') -Value (@($lines) + $Raw) -Encoding utf8
    }

    function Write-Status([string]$Dir, [hashtable]$Fields) {
        $s = [ordered]@{ v = 1; jobId = 'imp-7c1e'; task = 'import'; title = 'Import - adele@contoso.com'; mode = 'Import'; store = 'D:\Stores\adele'; state = 'resumable' }
        foreach ($k in $Fields.Keys) { $s[$k] = $Fields[$k] }
        $s | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $Dir 'status.json') -Encoding utf8
    }

    function Get-ReportData([string]$Html) {
        $m = [regex]::Match($Html, '<script type="application/json" id="mbt-report-data">(.*?)</script>', 'Singleline')
        $m.Success | Should -BeTrue
        return ($m.Groups[1].Value | ConvertFrom-Json)
    }

    # An import that was stopped in the middle: steps, phases, waits, a sleep, warnings, folders.
    function New-StoppedImport([string]$Name) {
        $dir = New-JobFolder $Name
        Write-Status $dir @{}
        Write-Events $dir @(
            @{ at = 0; e = 'plan'; job = 'imp-7c1e'; task = 'import'; mode = 'Import'; steps = @(@{ id = 'check'; title = 'Verify the export store' }, @{ id = 'folders'; title = 'Create the folders' }, @{ id = 'send'; title = 'Send the items' }) }
            @{ at = 1; e = 'step'; id = 'check'; state = 'running'; msg = 'Signing in to Microsoft' }
            @{ at = 4; e = 'step'; id = 'check'; state = 'done'; msg = 'The store is complete.' }
            @{ at = 4; e = 'step'; id = 'folders'; state = 'running' }
            @{ at = 6; e = 'step'; id = 'folders'; state = 'done'; done = 3; total = 3 }
            @{ at = 10; e = 'step'; id = 'send'; state = 'running'; done = 0; total = 100; bytes = 0; bytesTotal = 5000000 }
            @{ at = 11; e = 'object'; step = 'send'; key = 'Inbox'; state = 'running'; done = 10; total = 60 }
            @{ at = 20; e = 'wait'; reason = 'throttle'; seconds = 38; count = 1; waited = 38 }
            @{ at = 60; e = 'wait'; reason = 'throttle'; seconds = 12; count = 2; waited = 50 }
            @{ at = 75; e = 'wait'; reason = 'throttle'; seconds = 10; count = 3; waited = 60 }
            @{ at = 100; e = 'warn'; code = 'engine'; step = 'send'; key = 'Quarterly report'; msg = 'The item was refused: too large.' }
            @{ at = 120; e = 'slept'; from = $script:T0.AddSeconds(110).ToString('o'); to = $script:T0.AddSeconds(190).ToString('o'); seconds = 80 }
            @{ at = 200; e = 'object'; step = 'send'; key = 'Inbox'; state = 'done'; done = 60; total = 60; bytes = 3000000 }
            @{ at = 210; e = 'object'; step = 'send'; key = 'Inbox/Projects'; state = 'partial'; done = 30; total = 40; bytes = 1200000; msg = '10 items failed' }
            @{ at = 312; e = 'step'; id = 'send'; state = 'running'; done = 90; total = 100; bytes = 4200000; bytesTotal = 5000000 }
            @{ at = 312; e = 'result'; exit = 2; complete = $false; counts = @{ done = 90; failed = 10; pending = 0; refused = 0; total = 100; foldersCreated = 3 }; next = 'continue'; reason = 'stopped' }
        )
        return $dir
    }
}

Describe 'The HTML report of a job' {
    It 'a copy of the Primary and the Archive shows the steps of both parts, with the part in their titles, in the chronology and in the data' {
        $dir = New-JobFolder 'both-parts'
        Write-Status $dir @{ jobId = 'mailcopy-1a2b'; task = 'mailcopy'; title = 'Copy - alex@contoso.com to adele@contoso.com'; mode = 'Copy'; state = 'resumable' }
        $steps = @(@{ id = 'check'; title = 'Verify the export store (Primary part)' }, @{ id = 'send'; title = 'Send the items (Primary part)' }, @{ id = 'check-2'; title = 'Verify the export store (Archive part)' }, @{ id = 'send-2'; title = 'Send the items (Archive part)' })
        Write-Events $dir @(
            @{ at = 0; e = 'plan'; job = 'mailcopy-1a2b'; task = 'mailcopy'; mode = 'Copy'; steps = $steps }
            @{ at = 1; e = 'step'; id = 'check'; state = 'done' }
            @{ at = 2; e = 'step'; id = 'send'; state = 'done'; done = 10; total = 10 }
            @{ at = 3; e = 'step'; id = 'check-2'; state = 'done' }
            @{ at = 4; e = 'step'; id = 'send-2'; state = 'running'; done = 3; total = 7 }
            @{ at = 5; e = 'object'; step = 'send-2'; key = 'Archive/Old'; state = 'running'; done = 3; total = 7 }
            @{ at = 6; e = 'result'; exit = 2; complete = $false; counts = @{ done = 13; total = 17; failed = 0 }; next = 'continue'; reason = 'stopped' })
        $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw
        $html | Should -Match 'Send the items \(Primary part\) done'
        $html | Should -Match 'Verify the export store \(Archive part\) done'
        $html | Should -Match 'Send the items \(Archive part\) started'
        $data = Get-ReportData $html
        @($data.steps).Count | Should -Be 4
        (@($data.steps | Where-Object id -EQ 'send-2')[0]).state | Should -Be 'running'
        (@($data.steps | Where-Object id -EQ 'check-2')[0]).title | Should -Be 'Verify the export store (Archive part)'
    }

    It 'is written in the job folder and says the task, the result and the numbers' {
        $dir = New-StoppedImport 'full'
        $path = Write-MbtJobHtmlReport -Folder $dir
        $path | Should -Be (Join-Path $dir 'report.html')
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Match '<title>Import - adele@contoso.com</title>'
        $html | Should -Match 'Stopped'
        $html | Should -Match 'Items done</div><div class="tile-value">90</div>' -Because 'the tile of the items done'
        $html | Should -Match '<div class="tile fail"><div class="tile-label">Failed</div><div class="tile-value">10</div>'
        $html | Should -Match 'Throttling waits</div><div class="tile-value">3</div><div class="tile-note">1 min 00 s waited'
        $html | Should -Match 'Folders created'
        $html | Should -Match '5 min 12 s' -Because 'the duration of the job'
        $html | Should -Match 'Continue the job'
        $html | Should -Match 'D:\\Stores\\adele'
        $html | Should -Not -Match '\{\{[A-Z_]+\}\}' -Because 'every marker of the template is replaced'
    }

    It 'says the period of an export, and gives the items kept, the items outside the period and the items without a date as numbers, in the page, the data and the CSV' {
        $dir = New-JobFolder 'period'
        Write-Status $dir @{ jobId = 'exp-3a2b'; task = 'export'; title = 'Export - adele@contoso.com'; mode = 'Export'; store = 'D:\Stores\adele'; state = 'complete'; exitCode = 0; period = @{ from = '2026-02-28T23:00:00.000Z'; to = '2026-03-31T21:59:59.999Z'; text = 'received between 2026-03-01 and 2026-03-31' } }
        Write-Events $dir @(
            @{ at = 0; e = 'plan'; job = 'exp-3a2b'; task = 'export'; mode = 'Export'; period = @{ from = '2026-02-28T23:00:00.000Z'; to = '2026-03-31T21:59:59.999Z'; text = 'received between 2026-03-01 and 2026-03-31' }; steps = @(@{ id = 'tree'; title = 'Read the folder tree' }, @{ id = 'list'; title = 'List the items you chose' }, @{ id = 'download'; title = 'Download the items' }) }
            @{ at = 2; e = 'step'; id = 'tree'; state = 'done' }
            @{ at = 5; e = 'step'; id = 'list'; state = 'done'; msg = '4,612 items outside the period left out; 3 without a date kept' }
            @{ at = 9; e = 'step'; id = 'download'; state = 'done'; done = 1208; total = 1208 }
            @{ at = 10; e = 'result'; exit = 0; complete = $true; counts = @{ done = 1208; failed = 0; bytes = 6200000000; kept = 1208; outsidePeriod = 4612; withoutDate = 3; periodDate = 'mixed'; periodNote = 'received date of the messages; 1 folder(s) filtered on the arrival date: ''Inbox/Projects'' (the Mail API answered HTTP 403 ErrorAccessDenied)' }; next = 'done' }
        )
        $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw
        $html | Should -Match '<span>Period</span>Items received between 2026-03-01 and 2026-03-31'
        $html | Should -Match '<span>Date used</span>received date of the messages; 1 folder\(s\) filtered on the arrival date: &#39;Inbox/Projects&#39; \(the Mail API answered HTTP 403 ErrorAccessDenied\)|<span>Date used</span>received date of the messages; 1 folder\(s\) filtered on the arrival date: .Inbox/Projects. \(the Mail API answered HTTP 403 ErrorAccessDenied\)'
        $html | Should -Match 'Items kept</div><div class="tile-value">1,208</div>'
        $html | Should -Match 'Outside the period</div><div class="tile-value">4,612</div>'
        $html | Should -Match 'Kept without a date</div><div class="tile-value">3</div>'
        $html | Should -Match '4,612 items outside the period left out; 3 without a date kept' -Because 'the step that listed the items says it, in the chronology'
        $data = Get-ReportData $html
        $data.period | Should -Be 'received between 2026-03-01 and 2026-03-31'
        $data.counts.outsidePeriod | Should -Be 4612
        $plain = New-StoppedImport 'noperiod'
        (Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $plain) -Raw) | Should -Not -Match '<span>Period</span>|Outside the period'
        $run = New-MbtGuiRun -Task export -Title 'Export - adele@contoso.com'
        $null = Add-MbtGuiEvent -Run $run -Event ([pscustomobject]@{ v = 1; t = $script:T0.ToString('o'); e = 'result'; exit = 0; complete = $true; counts = [pscustomobject]@{ done = 1208; kept = 1208; outsidePeriod = 4612; withoutDate = 3 }; next = 'done' })
        $files = @(Write-MbtGuiRunReport -Run $run -Folder $dir -Task export -NoHtml)
        $csv = @(Import-Csv -LiteralPath (Join-Path $dir 'report.csv'))
        $row = $csv | Where-Object Kind -EQ 'period'
        $row.Done | Should -Be '4612'
        $row.State | Should -Be 'skipped'
        $row.Message | Should -Match '1208 items kept; 4612 items outside the period were skipped \(not an error\); 3 items without a date were kept'
    }

    It 'has the chronology: every step with its start, end and duration, and the events in time order' {
        $dir = New-StoppedImport 'chrono'
        $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw
        $html = $html.Substring($html.IndexOf('id="chrono"', [StringComparison]::Ordinal))
        $order = 'Verify the export store started', 'Signing in to Microsoft', 'Verify the export store done: The store is complete.', 'Create the folders started', 'Create the folders done', 'Send the items started',
            'Microsoft asked to wait because of too many requests: 3 waits', 'The item was refused: too large.', 'This PC was asleep for 1 min 20 s'
        $at = -1
        foreach ($text in $order) {
            $i = $html.IndexOf($text, [StringComparison]::Ordinal)
            $i | Should -BeGreaterThan $at -Because "'$text' comes after the one before it"
            $at = $i
        }
        $html = Get-Content -LiteralPath (Join-Path $dir 'report.html') -Raw
        $data = Get-ReportData $html
        @($data.steps).Count | Should -Be 3
        $send = @($data.steps | Where-Object id -EQ 'send')[0]
        $send.state | Should -Be 'running'
        ([datetime]$send.started).ToUniversalTime() | Should -Be $script:T0.AddSeconds(10)
        $send.finished | Should -BeNullOrEmpty -Because 'the step did not end'
        $check = @($data.steps | Where-Object id -EQ 'check')[0]
        (([datetime]$check.finished) - ([datetime]$check.started)).TotalSeconds | Should -Be 3
        $html | Should -Match 'not finished' -Because 'a step that did not end says so'
        $html | Should -Match '<svg class="timeline"' -Because 'the picture of the chronology'
        @([regex]::Matches($html, 'data-kind="wait"')).Count | Should -Be 1 -Because 'the waits that follow each other are one row'
        @([regex]::Matches($html, 'data-kind="sleep"')).Count | Should -Be 1
        $data.waits.count | Should -Be 3
        $data.waits.seconds | Should -Be 60
    }

    It 'lists the folders with their items, size and state, the warnings, and links the files that exist' {
        $dir = New-StoppedImport 'folders'
        Set-Content -LiteralPath (Join-Path $dir 'report.csv') -Value 'Kind' -Encoding utf8
        $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw
        $html | Should -Match '<td>Inbox/Projects</td><td>Send the items</td><td><span class="badge warn">partial</span></td><td class="num">30 of 40</td><td class="num">1.1 MiB</td><td>10 items failed</td>'
        $html | Should -Match '<td>Inbox</td><td>Send the items</td><td><span class="badge ok">done</span></td><td class="num">60</td>' -Because 'a folder that is done says what was done, not "60 of 60"'
        $html | Should -Match '<td class="mono">engine</td><td>Quarterly report</td><td>The item was refused: too large.</td>'
        $html | Should -Match '<a href="report.csv">Report \(CSV\)</a>'
        $html | Should -Match '<a href="events.jsonl">'
        $html | Should -Not -Match 'href="report.json"' -Because 'this file does not exist'
    }

    It 'a job that failed: the error, its technical message, and what to do' {
        $dir = New-JobFolder 'failed'
        Write-Status $dir @{ task = 'export'; mode = 'Export'; title = 'Export - adele@contoso.com'; state = 'error' }
        Write-Events $dir @(
            @{ at = 0; e = 'plan'; job = 'exp-1'; task = 'export'; mode = 'Export'; steps = @(@{ id = 'tree'; title = 'Read the folder tree' }) }
            @{ at = 1; e = 'step'; id = 'tree'; state = 'running'; msg = 'Signing in to Microsoft' }
            @{ at = 5; e = 'step'; id = 'tree'; state = 'failed'; msg = 'The mailbox was not found.' }
            @{ at = 5; e = 'error'; code = 'mailbox-not-found'; msg = 'The mailbox was not found. Nothing was changed.'; tech = 'GET /mailboxes/x -> 404 ErrorInvalidUser' }
            @{ at = 5; e = 'result'; exit = 1; complete = $false; counts = @{ done = 0; failed = 0 }; next = 'retry' }
        )
        $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw
        $html | Should -Match '<div class="tile-value">Failed</div>' -Because 'the result tile'
        $html | Should -Match '<section class="card problem"><h2>What went wrong</h2><p>The mailbox was not found. Nothing was changed.'
        $html | Should -Match 'mailbox-not-found'
        $html | Should -Match '<details><summary>Technical message</summary><pre>GET /mailboxes/x -&gt; 404 ErrorInvalidUser</pre></details>'
        $html | Should -Match 'fix its cause'
        $html | Should -Match 'class="badge fail">failed</span>'
        $html | Should -Match 'data-kind="error"'
        (Get-ReportData $html).errors[0].code | Should -Be 'mailbox-not-found'
    }

    It 'a job that finished clean, one that refused items, one that needs a decision, one that was never ended' {
        $cases = @(
            @{ Name = 'ok'; Result = @{ exit = 0; complete = $true; counts = @{ done = 5; failed = 0 }; next = 'done' }; Text = 'Complete'; Next = 'Nothing is left to do.' }
            @{ Name = 'refused'; Result = @{ exit = 1; complete = $true; counts = @{ done = 5; refused = 2; failed = 0; report = 'D:\work\not-imported.csv' }; next = 'done'; reason = 'refused-items' }; Text = 'Finished, items refused'; Next = 'Nothing can be run again' }
            @{ Name = 'decide'; Result = @{ exit = 3; complete = $false; counts = @{ done = 5; unknown = 1 }; next = 'decide' }; Text = 'Needs your decision'; Next = 'Settle the uncertain items' }
            @{ Name = 'signin'; Result = @{ exit = 2; complete = $false; counts = @{ done = 5 }; next = 'signin'; reason = 'signin-required' }; Text = 'Sign-in needed'; Next = 'Sign in again' }
            @{ Name = 'noend'; Result = $null; Text = 'Interrupted'; Next = 'Run the same job again' }
        )
        foreach ($c in $cases) {
            $dir = New-JobFolder "verdict-$($c.Name)"
            Write-Status $dir @{}
            $events = @(@{ at = 0; e = 'plan'; steps = @(@{ id = 'send'; title = 'Send the items' }) }, @{ at = 2; e = 'step'; id = 'send'; state = 'running'; done = 1; total = 5 })
            if ($c.Result) { $r = @{ at = 9; e = 'result' }; foreach ($k in $c.Result.Keys) { $r[$k] = $c.Result[$k] }; $events += $r }
            Write-Events $dir $events
            $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw
            $html | Should -Match ([regex]::Escape("<div class=""tile-value"">$($c.Text)</div>")) -Because $c.Name
            $html | Should -Match ([regex]::Escape($c.Next)) -Because $c.Name
        }
        (Get-Content -LiteralPath (Join-Path $TestDrive 'verdict-refused\report.html') -Raw) | Should -Match 'not-imported.csv'
        (Get-Content -LiteralPath (Join-Path $TestDrive 'verdict-noend\report.html') -Raw) | Should -Match 'not written \(last event'
    }

    It 'an empty job folder, a folder without status, and a line that is not JSON do not stop the report' {
        $empty = New-JobFolder 'empty'
        $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $empty) -Raw
        $html | Should -Match 'Interrupted'
        $html | Should -Match 'The job wrote no step.'
        $html | Should -Match 'The job wrote no event.'
        $html | Should -Not -Match '\{\{[A-Z_]+\}\}'
        $data = Get-ReportData $html
        $data.interrupted | Should -BeTrue

        $odd = New-JobFolder 'oddline'
        Write-Events $odd @(@{ at = 0; e = 'plan'; steps = @(@{ id = 'run'; title = 'Run' }) }, @{ at = 3; e = 'step'; id = 'run'; state = 'done' }, @{ at = 3; e = 'result'; exit = 0; complete = $true; counts = @{ done = 1 }; next = 'done' }) -Raw @('this is not json', '{"v":1,"t":"2026-11-09T14:00:09.000Z","e":"info","msg":"after the odd line"}')
        $html2 = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $odd) -Raw
        $html2 | Should -Match 'Complete'
        $html2 | Should -Match 'after the odd line'
        { Write-MbtJobHtmlReport -Folder (Join-Path $TestDrive 'nowhere') } | Should -Throw '*does not exist*'
    }

    It 'a folder, a message or a title with angle brackets, an ampersand, quotes and a marker of the template is encoded, never taken for markup or for a marker' {
        $dir = New-JobFolder 'encoding'
        $evil = '<img src=x onerror=alert(1)>&"'' {{TITLE}}'
        $closer = '</script><script>alert(2)</script>'
        Write-Status $dir @{ title = "Import - $evil"; store = $evil }
        Write-Events $dir @(
            @{ at = 0; e = 'plan'; steps = @(@{ id = 'send'; title = "Send $evil" }) }
            @{ at = 1; e = 'step'; id = 'send'; state = 'running'; msg = $closer }
            @{ at = 2; e = 'object'; step = 'send'; key = "Inbox/$evil"; state = 'done'; done = 1; total = 1; msg = $closer }
            @{ at = 3; e = 'item'; step = 'send'; key = $evil; state = 'refused'; msg = $closer; detail = $evil }
            @{ at = 4; e = 'warn'; code = '<b>'; step = 'send'; key = $evil; msg = $closer }
            @{ at = 5; e = 'error'; code = 'error'; msg = $closer; tech = $evil }
            @{ at = 6; e = 'result'; exit = 1; complete = $false; counts = @{ done = 1 }; next = 'retry' }
        )
        $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw
        $html | Should -Not -Match ([regex]::Escape('<img src=x'))
        $html | Should -Not -Match ([regex]::Escape('<script>alert(2)'))
        $html | Should -Not -Match ([regex]::Escape('onerror=alert(1)>'))
        $html | Should -Match ([regex]::Escape('&lt;img src=x onerror=alert(1)&gt;&amp;&quot;&#39; {{TITLE}}'))
        @([regex]::Matches($html, '<script')).Count | Should -Be 3 -Because 'the theme script, the data, the table filter: nothing came from the job'
        @([regex]::Matches($html, 'id="mbt-report-data"')).Count | Should -Be 1
        $data = Get-ReportData $html
        $data.errors[0].msg | Should -Be $closer -Because 'the data holds the text as it was, and nothing in it can close the script element'
        $data.errors[0].tech | Should -Be $evil
        $data.folders[0].key | Should -Be "Inbox/$evil"
        $data.title | Should -Be "Import - $evil"
        $escape = [string][char]92 + 'u003c'
        $html | Should -Match ([regex]::Escape($escape))
    }

    It 'many rows are capped and say how many were left out' {
        $dir = New-JobFolder 'cap'
        $events = @(@{ at = 0; e = 'plan'; steps = @(@{ id = 'send'; title = 'Send' }) })
        foreach ($i in 1..30) { $events += @{ at = $i; e = 'object'; step = 'send'; key = "Folder $i"; state = 'done'; done = 1; total = 1 } }
        foreach ($i in 1..12) { $events += @{ at = 40 + $i; e = 'warn'; code = 'w'; step = 'send'; key = "k$i"; msg = "warning $i" } }
        $events += @{ at = 99; e = 'result'; exit = 0; complete = $true; counts = @{ done = 30 }; next = 'done' }
        Write-Events $dir $events
        $m = Get-MbtJobReportModel -Folder $dir -MaxFolders 10 -MaxWarnings 5 -MaxChronology 8
        @($m.Folders).Count | Should -Be 10
        $m.LeftFolders | Should -Be 20
        @($m.Warnings).Count | Should -Be 5
        $m.LeftWarnings | Should -Be 7
        $m.LeftChronology | Should -BeGreaterThan 0
        $m.Counts['done'] | Should -Be 30
        $full = Get-MbtJobReportModel -Folder $dir
        $full.LeftFolders | Should -Be 0
        @($full.Folders).Count | Should -Be 30
    }

    It 'says what the job reads from and writes to: status.json, else the title' {
        $dir = New-JobFolder 'route'
        Write-Status $dir @{ source = 'D:\Stores\adele'; destination = 'adele@contoso.com, folder Restore' }
        Write-Events $dir @(@{ at = 0; e = 'plan'; steps = @(@{ id = 'send'; title = 'Send' }) }, @{ at = 1; e = 'result'; exit = 0; complete = $true; counts = @{ done = 0 }; next = 'done' })
        $html = Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw
        $html | Should -Match '<span>Source</span>D:\\Stores\\adele'
        $html | Should -Match '<span>Destination</span>adele@contoso.com, folder Restore'
        $derived = New-JobFolder 'route-derived'
        Write-Status $derived @{}
        Write-Events $derived @(@{ at = 0; e = 'result'; exit = 0; complete = $true; counts = @{ done = 0 }; next = 'done' })
        $m = Get-MbtJobReportModel -Folder $derived
        $m.Source | Should -Be 'D:\Stores\adele'
        $m.Destination | Should -Be 'adele@contoso.com'
    }

    It 'reads a file that is still open for writing, and replaces an older report' {
        $dir = New-StoppedImport 'locked'
        $stream = [IO.FileStream]::new((Join-Path $dir 'events.jsonl'), 'Append', 'Write', 'ReadWrite')
        try {
            $path = Write-MbtJobHtmlReport -Folder $dir
            Set-Content -LiteralPath $path -Value 'old'
            $path = Write-MbtJobHtmlReport -Folder $dir
            (Get-Content -LiteralPath $path -Raw) | Should -Match '<!doctype html>'
            @(Get-ChildItem -LiteralPath $dir -Filter '*.tmp').Count | Should -Be 0
        } finally { $stream.Dispose() }
    }
}

Describe 'The wording of the report' {
    BeforeAll {
        function New-ReportOf([string]$Name, [object[]]$Events, [hashtable]$Status = @{}) {
            $dir = New-JobFolder $Name
            Write-Status $dir $Status
            Write-Events $dir $Events
            return (Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw)
        }
        function Get-Cell([string]$Html, [string]$Step, [int]$Column) {
            $row = [regex]::Match($Html, '<tr><td>' + [regex]::Escape($Step) + '</td>(.*?)</tr>').Groups[1].Value
            return ([regex]::Matches($row, '<td[^>]*>(.*?)</td>') | ForEach-Object { $_.Groups[1].Value })[$Column]
        }
    }

    It 'the card of what to do next says one sentence when there is nothing more to do, and a second one only when it is another sentence' {
        $cases = @(
            @{ Name = 'ok'; Result = @{ exit = 0; complete = $true; counts = @{ done = 5 }; next = 'done' }; Paragraphs = 1 }
            @{ Name = 'warnings-done'; Result = @{ exit = 2; complete = $true; counts = @{ done = 5 }; next = 'done'; reason = 'warnings' }; Paragraphs = 1 }
            @{ Name = 'refused'; Result = @{ exit = 1; complete = $true; counts = @{ done = 5; refused = 2 }; next = 'done'; reason = 'refused-items' }; Paragraphs = 2 }
            @{ Name = 'failed'; Result = @{ exit = 1; complete = $false; counts = @{ done = 1 }; next = 'retry' }; Paragraphs = 2 }
            @{ Name = 'errors'; Result = @{ exit = 1; complete = $false; counts = @{ done = 1 }; next = 'repair'; reason = 'errors' }; Paragraphs = 2 }
            @{ Name = 'stopped'; Result = @{ exit = 2; complete = $false; counts = @{ done = 1 }; next = 'continue'; reason = 'stopped' }; Paragraphs = 2 }
            @{ Name = 'signin'; Result = @{ exit = 2; complete = $false; counts = @{ done = 1 }; next = 'signin'; reason = 'signin-required' }; Paragraphs = 2 }
            @{ Name = 'partial'; Result = @{ exit = 2; complete = $false; counts = @{ done = 1 }; next = 'continue' }; Paragraphs = 2 }
            @{ Name = 'decide'; Result = @{ exit = 3; complete = $false; counts = @{ done = 1; unknown = 1 }; next = 'decide' }; Paragraphs = 2 }
            @{ Name = 'other'; Result = @{ exit = 9; complete = $false; counts = @{ done = 1 }; next = 'retry' }; Paragraphs = 2 }
            @{ Name = 'noend'; Result = $null; Paragraphs = 2 }
        )
        foreach ($c in $cases) {
            $events = @(@{ at = 0; e = 'plan'; steps = @(@{ id = 'send'; title = 'Send the items' }) }, @{ at = 2; e = 'step'; id = 'send'; state = 'running'; done = 1; total = 5 })
            if ($c.Result) { $r = @{ at = 9; e = 'result' }; foreach ($k in $c.Result.Keys) { $r[$k] = $c.Result[$k] }; $events += $r }
            $html = New-ReportOf "next-$($c.Name)" $events
            $card = [regex]::Match($html, '<section class="card next[^"]*"><h2>What to do next</h2>(.*?)</section>').Groups[1].Value
            $paragraphs = @([regex]::Matches($card, '<p>(?:<strong>)?(.*?)(?:</strong>)?</p>') | ForEach-Object { $_.Groups[1].Value })
            $paragraphs.Count | Should -Be $c.Paragraphs -Because $c.Name
            @($paragraphs | Select-Object -Unique).Count | Should -Be $paragraphs.Count -Because "no sentence of '$($c.Name)' is said twice"
            $card | Should -Not -Match 'Nothing is left to do\.[\s\S]*nothing is left to do' -Because $c.Name
        }
    }

    It 'the header says the task, never the title again' {
        $html = New-ReportOf 'task-export' @(@{ at = 0; e = 'result'; exit = 0; complete = $true; counts = @{ done = 0 }; next = 'done' }) @{ task = 'export'; mode = 'Export'; title = 'Export - adele@contoso.com' }
        $html | Should -Match '<span>Task</span>Export</div>'
        $html | Should -Not -Match '<span>Task</span>Export -'
        $html | Should -Match '<title>Export - adele@contoso.com</title>' -Because 'the title stays the title'
        Get-MbtReportTaskName -Task 'export' -Mode 'Export' | Should -Be 'Export'
        Get-MbtReportTaskName -Task 'export' -Mode 'ExportPst' | Should -Be 'Export (PST)'
        Get-MbtReportTaskName -Task 'import' -Mode 'ImportEml' | Should -Be 'Import (EML files)'
        Get-MbtReportTaskName -Task 'copy' -Mode 'ExportPst' | Should -Be 'Copy (PST)'
        Get-MbtReportTaskName -Task 'repair' -Mode 'RepairPst' | Should -Be 'Repair'
        Get-MbtReportTaskName -Task 'IMPORT' -Mode 'Import' | Should -Be 'Import'
        Get-MbtReportTaskName -Task 'export-plan' -Mode 'Export' | Should -Be 'Export (plan check)'
        Get-MbtReportTaskName -Task '' -Mode 'ImportPst' | Should -Be 'Import (PST)'
        Get-MbtReportTaskName -Task '' -Mode '' | Should -BeNullOrEmpty
    }

    It 'a step that is done says only what it really did, never more than the total that was announced' {
        $events = @(
            @{ at = 0; e = 'plan'; steps = @(@{ id = 'download'; title = 'Download the items' }) }
            @{ at = 1; e = 'step'; id = 'download'; state = 'running'; done = 0; total = 53; bytes = 0; bytesTotal = 8320848 }
            @{ at = 3; e = 'step'; id = 'download'; state = 'running'; done = 40; total = 53; bytes = 5242880; bytesTotal = 8320848 }
            @{ at = 5; e = 'step'; id = 'download'; state = 'done'; done = 53; total = 53; bytes = 8620875; bytesTotal = 8320848 }
            @{ at = 5; e = 'result'; exit = 0; complete = $true; counts = @{ done = 53; bytes = 8620875 }; next = 'done' }
        )
        $html = New-ReportOf 'bytes-done' $events
        Get-Cell $html 'Download the items' 4 | Should -Be '53'
        Get-Cell $html 'Download the items' 5 | Should -Be '8.2 MiB'
        $html | Should -Not -Match '8\.2 MiB of'
        $html | Should -Not -Match '53 of 53'

        $open = New-ReportOf 'bytes-open' ($events[0..2])
        Get-Cell $open 'Download the items' 4 | Should -Be '40 of 53'
        Get-Cell $open 'Download the items' 5 | Should -Be '5.0 MiB of about 7.9 MiB'

        $over = New-ReportOf 'bytes-over' @($events[0], @{ at = 3; e = 'step'; id = 'download'; state = 'running'; done = 60; total = 53; bytes = 9000000; bytesTotal = 8320848 })
        Get-Cell $over 'Download the items' 4 | Should -Be '60' -Because 'more than announced: only what was done'
        Get-Cell $over 'Download the items' 5 | Should -Be '8.6 MiB'
    }

    It 'a folder that was announced with more or fewer items than it holds is never shown with more done than the total' {
        $events = @(
            @{ at = 0; e = 'plan'; steps = @(@{ id = 'download'; title = 'Download the items' }) }
            @{ at = 1; e = 'object'; step = 'download'; key = 'Inbox'; state = 'running'; done = 12; total = 10 }
            @{ at = 1; e = 'object'; step = 'download'; key = 'Sent'; state = 'running'; done = 3; total = 10 }
            @{ at = 1; e = 'object'; step = 'download'; key = 'Drafts'; state = 'done'; done = 7; total = 5 }
        )
        $html = New-ReportOf 'folder-over' $events
        $html | Should -Match '<td>Inbox</td>.*?<td class="num">12</td>'
        $html | Should -Match '<td>Sent</td>.*?<td class="num">3 of 10</td>'
        $html | Should -Match '<td>Drafts</td>.*?<td class="num">7</td>'
        $html | Should -Not -Match '12 of 10|7 of 5'
        Format-MbtReportProgress 5 0 | Should -Be '5'
        Format-MbtReportProgress 0 0 | Should -Be ''
        Format-MbtReportProgress 0 9 -Finished | Should -Be '9'
        Format-MbtReportProgress 1048576 2097152 -Bytes -About | Should -Be '1.0 MiB of about 2.0 MiB'
    }

    It 'a step that is done keeps no progress text: its own final message only, and a running step shows its progress' {
        $events = @(
            @{ at = 0; e = 'plan'; steps = @(@{ id = 'tree'; title = 'Read the folder tree' }, @{ id = 'list'; title = 'List the items' }, @{ id = 'send'; title = 'Send the items' }) }
            @{ at = 1; e = 'step'; id = 'tree'; state = 'running'; msg = 'Measuring 1 folder the service gives no size for: 1 done' }
            @{ at = 2; e = 'step'; id = 'tree'; state = 'done' }
            @{ at = 3; e = 'step'; id = 'list'; state = 'running'; msg = 'Counting' }
            @{ at = 4; e = 'step'; id = 'list'; state = 'done'; msg = 'Each item is checked as it arrives.' }
            @{ at = 5; e = 'step'; id = 'send'; state = 'running'; msg = 'Sending 4 of 9' }
        )
        $html = New-ReportOf 'msg-done' $events
        Get-Cell $html 'Read the folder tree' 6 | Should -Be ''
        Get-Cell $html 'List the items' 6 | Should -Be 'Each item is checked as it arrives.'
        Get-Cell $html 'Send the items' 6 | Should -Be 'Sending 4 of 9'
        $chrono = $html.Substring($html.IndexOf('id="chrono"', [StringComparison]::Ordinal))
        $chrono | Should -Match 'Measuring 1 folder the service gives no size for: 1 done' -Because 'the progress stays in the chronology'
    }

    It 'a number of times or of rows agrees with its noun' {
        Format-MbtReportQuantity 1 'time' | Should -Be '1 time'
        Format-MbtReportQuantity 3 'time' | Should -Be '3 times'
        Format-MbtReportQuantity 1200 'folder' | Should -Be '1,200 folders'
        $html = New-ReportOf 'network-once' @(
            @{ at = 0; e = 'plan'; steps = @(@{ id = 'send'; title = 'Send' }) }
            @{ at = 5; e = 'wait'; reason = 'network'; seconds = 30; count = 1; waited = 30 }
            @{ at = 40; e = 'result'; exit = 0; complete = $true; counts = @{ done = 1 }; next = 'done' }
        )
        $html | Should -Match 'tried again \(1 time, 30 s in all\)'
        $html | Should -Not -Match '\(s\)|1 times'
        $events = @(@{ at = 0; e = 'plan'; steps = @(@{ id = 'send'; title = 'Send' }) })
        foreach ($i in 1..2) { $events += @{ at = $i; e = 'object'; step = 'send'; key = "Folder $i"; state = 'done'; done = 1; total = 1 } }
        $events += @{ at = 9; e = 'result'; exit = 0; complete = $true; counts = @{ done = 2 }; next = 'done' }
        $dir = New-JobFolder 'caption-one'
        Write-Status $dir @{}
        Write-Events $dir $events
        $model = Get-MbtJobReportModel -Folder $dir -MaxFolders 1
        $page = ConvertTo-MbtJobReportHtml -Model $model
        $page | Should -Match '1 more folder is in events.jsonl'
    }
}

Describe 'The console and the report' {
    It 'opens report.html when the job has one, else report.csv, and the CSV of a job made by an older version' {
        $dir = New-JobFolder 'pick'
        Get-MbtGuiReportPath -Folder $dir | Should -Be (Join-Path $dir 'report.csv') -Because 'nothing was written: the CSV is the answer an older job gets'
        Set-Content -LiteralPath (Join-Path $dir 'report.csv') -Value 'Kind'
        Get-MbtGuiReportPath -Reports @((Join-Path $dir 'report.json'), (Join-Path $dir 'report.csv')) -Folder $dir | Should -Be (Join-Path $dir 'report.csv')
        Set-Content -LiteralPath (Join-Path $dir 'report.html') -Value '<html></html>'
        Get-MbtGuiReportPath -Reports @((Join-Path $dir 'report.csv')) -Folder $dir | Should -Be (Join-Path $dir 'report.html') -Because 'the HTML is next to the CSV'
        Get-MbtGuiReportPath -Reports @((Join-Path $dir 'report.json'), (Join-Path $dir 'report.csv'), (Join-Path $dir 'report.html')) -Folder $dir | Should -Be (Join-Path $dir 'report.html')
        Remove-Item -LiteralPath (Join-Path $dir 'report.html')
        Get-MbtGuiReportPath -Reports @((Join-Path $dir 'report.csv'), (Join-Path $dir 'report.html')) -Folder $dir | Should -Be (Join-Path $dir 'report.csv') -Because 'a report that is listed but missing is not offered'
        Get-MbtGuiReportPath | Should -BeNullOrEmpty
    }

    It 'the console writes the HTML report beside the CSV and the JSON, from the events of the job, and it links them' {
        $dir = New-StoppedImport 'console'
        $run = New-MbtGuiRun -Task import -Title 'Import - adele@contoso.com'
        $files = @(Write-MbtGuiRunReport -Run $run -Folder $dir -Task import)
        $files.Count | Should -Be 3
        ($files | Where-Object { $_ -like '*.html' }) | Should -Be (Join-Path $dir 'report.html')
        $html = Get-Content -LiteralPath (Join-Path $dir 'report.html') -Raw
        $html | Should -Match '<a href="report.csv">'
        $html | Should -Match '<a href="report.json">'
        Get-MbtGuiReportPath -Reports $files -Folder $dir | Should -Be (Join-Path $dir 'report.html')
    }

    It 'a run without events.jsonl (the tests, the demonstration) writes the CSV and the JSON only' {
        $dir = New-JobFolder 'noevents'
        $run = New-MbtGuiRun -Task export
        @(Write-MbtGuiRunReport -Run $run -Folder $dir -Task export).Count | Should -Be 2
        (Join-Path $dir 'report.html') | Should -Not -Exist
    }
}

Describe 'The shipped template' {
    It 'is in the package, needs no network and has every marker once' {
        $file = Join-Path $script:Root 'templates\Report.template.html'
        $file | Should -Exist
        $t = Get-Content -LiteralPath $file -Raw
        $t | Should -Not -Match '(src|href)\s*=\s*"https?:'
        $t | Should -Not -Match 'url\(\s*https?:'
        foreach ($marker in 'TITLE', 'VERSION', 'SUBTITLE', 'META', 'TILES', 'NEXT', 'ERRORS', 'TIMELINE', 'STEPS', 'CHRONOLOGY', 'FOLDERS', 'ITEMS', 'WARNINGS', 'DECISIONS', 'FILES', 'DATA', 'GENERATED') {
            [regex]::Matches($t, "\{\{$marker\}\}").Count | Should -BeGreaterThan 0 -Because $marker
        }
        $t | Should -Match '--cp-accent'
        $t | Should -Match 'data-theme="dark"'
    }

    It 'falls back to a plain page when the template is missing' {
        $dir = New-StoppedImport 'notemplate'
        $path = Write-MbtJobHtmlReport -Folder $dir -TemplatePath (Join-Path $TestDrive 'missing.html')
        (Get-Content -LiteralPath $path -Raw) | Should -Match '<h1>Import - adele@contoso.com</h1>'
    }
}

Describe 'The report of a very large job' {
    BeforeAll {
        # events.jsonl of a large export, written in one go: a plan, a step, a pair of folder events for each folder (the last three fail), warnings, waits and notes, a result
        function New-LargeJob([string]$Name, [int]$Folders, [int]$Warnings, [int]$Waits, [switch]$WithCsv) {
            $dir = New-JobFolder $Name
            $writer = [IO.StreamWriter]::new((Join-Path $dir 'events.jsonl'), $false, [Text.UTF8Encoding]::new($false))
            $at = 0.0
            $stamp = { param([double]$S) $script:T0.AddSeconds($S).ToString('yyyy-MM-dd\THH:mm:ss.fff\Z') }
            try {
                $writer.WriteLine('{"v":1,"t":"' + (& $stamp 0) + '","e":"plan","job":"big-0001","task":"export","mode":"Export","steps":[{"id":"download","title":"Download the items"}]}')
                $writer.WriteLine('{"v":1,"t":"' + (& $stamp 1) + '","e":"step","id":"download","state":"running","done":0,"total":1000000}')
                for ($i = 0; $i -lt $Folders; $i++) {
                    $at += 0.5
                    $failed = $i -ge $Folders - 3
                    $writer.WriteLine('{"v":1,"t":"' + (& $stamp $at) + '","e":"object","step":"download","key":"Folder ' + $i + '","state":"running","done":0,"total":10}')
                    $writer.WriteLine('{"v":1,"t":"' + (& $stamp ($at + 0.2)) + '","e":"object","step":"download","key":"Folder ' + $i + '","state":"' + $(if ($failed) { 'failed' } else { 'done' }) + '","done":10,"total":10,"bytes":' + ($i * 100) + '}')
                }
                for ($i = 0; $i -lt $Warnings; $i++) { $at += 0.7; $writer.WriteLine('{"v":1,"t":"' + (& $stamp $at) + '","e":"warn","code":"engine","step":"download","key":"Item ' + $i + '","msg":"The item ' + $i + ' was refused: too large."}') }
                for ($i = 0; $i -lt $Waits; $i++) {
                    $at += 3; $writer.WriteLine('{"v":1,"t":"' + (& $stamp $at) + '","e":"wait","reason":"throttle","seconds":2,"count":' + ($i + 1) + ',"waited":' + (2 * ($i + 1)) + '}')
                    $at += 1; $writer.WriteLine('{"v":1,"t":"' + (& $stamp $at) + '","e":"info","msg":"Went on ' + $i + '"}')
                }
                $writer.WriteLine('{"v":1,"t":"' + (& $stamp ($at + 5)) + '","e":"result","exit":0,"complete":true,"counts":{"done":123456,"failed":3},"next":"done"}')
            } finally { $writer.Dispose() }
            Write-Status $dir @{ jobId = 'big-0001'; task = 'export'; title = 'Export - adele@contoso.com'; mode = 'Export'; state = 'complete'; exitCode = 0 }
            if ($WithCsv) { Set-Content -LiteralPath (Join-Path $dir 'report.csv') -Value 'Kind' }
            return $dir
        }
        $script:Large = New-LargeJob 'large' 5000 1200 1500 -WithCsv
        $script:LargeTime = Measure-Command { $script:LargePath = Write-MbtJobHtmlReport -Folder $script:Large }
        $script:LargeHtml = Get-Content -LiteralPath $script:LargePath -Raw
        $script:LargeModel = Get-MbtJobReportModel -Folder $script:Large
    }

    It 'keeps the page small: the folders and the chronology are capped, whatever the size of the job' {
        (Get-Item -LiteralPath $script:LargePath).Length | Should -BeLessThan 3MB
        $script:LargeTime.TotalSeconds | Should -BeLessThan 90 -Because 'a job of 5,000 folders is reported in seconds, not minutes'
        $folders = $script:LargeHtml.Substring($script:LargeHtml.IndexOf('id="folders"', [StringComparison]::Ordinal))
        $folders = $folders.Substring(0, $folders.IndexOf('</table>', [StringComparison]::Ordinal))
        ([regex]::Matches($folders, '<tr><td>')).Count | Should -Be 2000
        $chrono = $script:LargeHtml.Substring($script:LargeHtml.IndexOf('id="chrono"', [StringComparison]::Ordinal))
        $chrono = $chrono.Substring(0, $chrono.IndexOf('</table>', [StringComparison]::Ordinal))
        ([regex]::Matches($chrono, '<tr class=')).Count | Should -Be 2000
        $data = Get-ReportData $script:LargeHtml
        @($data.folders).Count | Should -Be 2000
        @($data.chronology).Count | Should -Be 2000
        @($data.warnings).Count | Should -Be 500
    }

    It 'says how many rows were left out, and where they are' {
        $script:LargeHtml | Should -Match '3,000 more folders are in report\.csv and events\.jsonl\. The folders that are not done come first\.'
        $script:LargeHtml | Should -Match 'more events are in events\.jsonl\. This table shows the first 2,000\.'
        $script:LargeHtml | Should -Match '700 more warnings are in events\.jsonl\.'
        $script:LargeModel.LeftFolders | Should -Be 3000
        $script:LargeModel.LeftWarnings | Should -Be 700
        # 1 step started, 1200 warnings, 1500 waits, 1500 notes
        $script:LargeModel.LeftChronology + @($script:LargeModel.Chronology).Count | Should -Be (1 + 1200 + 1500 + 1500)
        $dir = New-JobFolder 'large-nocsv'
        Copy-Item -LiteralPath (Join-Path $script:Large 'events.jsonl'), (Join-Path $script:Large 'status.json') -Destination $dir
        (Get-Content -LiteralPath (Write-MbtJobHtmlReport -Folder $dir) -Raw) | Should -Match '3,000 more folders are in events\.jsonl\. The folders'
    }

    It 'lists the folders that are not done first, and keeps the numbers of the whole job' {
        $names = @($script:LargeModel.Folders | ForEach-Object { $_.Key })
        $names | Should -Contain 'Folder 4999'
        $names | Should -Contain 'Folder 4997'
        @($script:LargeModel.Folders | Where-Object State -EQ 'failed').Count | Should -Be 3
        $script:LargeModel.Folders[0].Key | Should -Be 'Folder 0' -Because 'the folders keep the order of the job'
        $script:LargeModel.Counts['done'] | Should -Be 123456
        $script:LargeModel.Counts['failed'] | Should -Be 3
        $script:LargeModel.Waits.Count | Should -Be 1500
        $script:LargeModel.Waits.Seconds | Should -Be 3000
        $script:LargeModel.EventCount | Should -Be (1 + 1 + 10000 + 1200 + 3000 + 1)
        $script:LargeHtml | Should -Match 'Throttling waits</div><div class="tile-value">1,500</div>'
    }

    It 'draws the waits of the picture as a few bars, not one for each wait' {
        ([regex]::Matches($script:LargeHtml, 'class="bar wait"')).Count | Should -BeLessThan 800
        ([regex]::Matches($script:LargeHtml, 'class="bar wait"')).Count | Should -BeGreaterThan 0
    }

    It 'a message of a job that is far too long is cut' {
        $dir = New-JobFolder 'longtext'
        Write-Events $dir @(
            @{ at = 0; e = 'plan'; steps = @(@{ id = 'send'; title = 'Send' }) }
            @{ at = 1; e = 'warn'; code = 'w'; step = 'send'; key = 'k'; msg = ('x' * 50000) }
            @{ at = 2; e = 'result'; exit = 0; complete = $true; counts = @{ done = 1 }; next = 'done' }
        )
        $m = Get-MbtJobReportModel -Folder $dir
        $m.Warnings[0].Msg.Length | Should -BeLessThan 2100
        (Get-Item -LiteralPath (Write-MbtJobHtmlReport -Folder $dir)).Length | Should -BeLessThan 100KB
    }

    It 'the caps can be changed, and the folders that are not done stay in' {
        $m = Get-MbtJobReportModel -Folder $script:Large -MaxFolders 50 -MaxChronology 40
        @($m.Folders).Count | Should -Be 50
        @($m.Chronology).Count | Should -Be 40
        $m.LeftFolders | Should -Be 4950
        @($m.Folders | Where-Object State -EQ 'failed').Count | Should -Be 3
    }

    It 'a job of 50,000 folders and 200,000 events is reported in minutes and the page stays at a few MB (opt in: MBT_BIG_REPORT=1)' -Skip:(-not $env:MBT_BIG_REPORT) {
        $dir = New-LargeJob 'huge' 50000 20000 20000
        $elapsed = Measure-Command { $path = Write-MbtJobHtmlReport -Folder $dir }
        $elapsed.TotalMinutes | Should -BeLessThan 5
        (Get-Item -LiteralPath $path).Length | Should -BeLessThan 4MB
    }
}
