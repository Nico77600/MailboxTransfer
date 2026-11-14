#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the overview of all the jobs (module\MailboxTransfer.Overview.ps1, templates\Overview.template.html): the rows, their order and
    filters as data, the encoding, the plan checks left out, a damaged status.json, the cache, the atomic write, the batch field, a folder of 5,000 jobs.
    The jobs are synthetic job folders; nothing reaches the network.

    Run:  Invoke-Pester -Path .\tests\Overview.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:Root = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    $script:SavedLocal = $env:MAILBOXTRANSFER_LOCAL_CONFIG
    $env:MAILBOXTRANSFER_LOCAL_CONFIG = Join-Path $TestDrive 'no-local-config.psd1'
    Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
    $script:Now = [datetime]::UtcNow

    # A job folder of <work>\jobs: status.json (unless -NoStatus), the files that are listed, the age of the job in hours.
    function New-TestJob {
        param([string]$Work, [string]$Name, [hashtable]$Status = @{}, [double]$HoursAgo = 1, [string[]]$Files = @(), [switch]$NoStatus, [string]$StatusText)
        $dir = Join-Path (Join-Path $Work 'jobs') $Name
        $null = New-Item -ItemType Directory -Force -Path $dir
        $started = $script:Now.AddHours(-$HoursAgo)
        $s = [ordered]@{ v = 1; jobId = "export-$Name"; task = 'export'; title = 'Export - adele@contoso.com'; mode = 'Export'; store = 'D:\Stores\adele'; source = 'adele@contoso.com'; destination = 'D:\Stores\adele'
            state = 'complete'; exitCode = 0; counts = [ordered]@{ done = 100; failed = 2; bytes = 5000 }; next = 'done'; reason = $null
            started = $started.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z'); updated = $started.AddMinutes(5).ToString('yyyy-MM-dd\THH:mm:ss.fff\Z'); pid = 4 }
        foreach ($k in $Status.Keys) { $s[$k] = $Status[$k] }
        if ($PSBoundParameters.ContainsKey('StatusText')) { [IO.File]::WriteAllText((Join-Path $dir 'status.json'), $StatusText) }
        elseif (-not $NoStatus) { [IO.File]::WriteAllText((Join-Path $dir 'status.json'), ($s | ConvertTo-Json -Depth 5), [Text.UTF8Encoding]::new($false)) }
        foreach ($f in $Files) { [IO.File]::WriteAllText((Join-Path $dir $f), 'x') }
        return $dir
    }

    # The data of the page: the columns named, the shared strings put back.
    function Get-OverviewData([string]$Html) {
        $m = [regex]::Match($Html, '<script type="application/json" id="mbt-overview-data">(.*?)</script>', 'Singleline')
        $m.Success | Should -BeTrue
        $d = $m.Groups[1].Value | ConvertFrom-Json
        $rows = @(foreach ($r in $d.rows) {
                $o = [ordered]@{}
                for ($i = 0; $i -lt $d.columns.Count; $i++) { $o[$d.columns[$i]] = $r[$i] }
                foreach ($c in 'task', 'source', 'destination', 'result') { $o[$c] = $d.strings[[int]$o[$c]] }
                $o['batch'] = if ([int]$o['batch'] -ge 0) { $d.strings[[int]$o['batch']] } else { '' }
                [pscustomobject]$o
            })
        [pscustomobject]@{ Data = $d; Rows = $rows }
    }
    function Get-OverviewPage([string]$Work) { Get-Content -LiteralPath (Write-MbtJobOverview -WorkFolder $Work) -Raw }
}
AfterAll { $env:MAILBOXTRANSFER_LOCAL_CONFIG = $script:SavedLocal }

Describe 'The overview of all the jobs' {
    It 'is written in the jobs folder, one row for each job, newest first, with the numbers of each' {
        $w = Join-Path $TestDrive 'basic'
        $null = New-TestJob $w '20261110-1000-export-aaaa' @{} 30 -Files 'report.html', 'report.csv'
        $null = New-TestJob $w '20261112-1000-import-bbbb' @{ jobId = 'import-bbbb'; task = 'import'; mode = 'Import'; source = 'D:\Stores\adele'; destination = 'adele@contoso.com'; state = 'error'; exitCode = 1; counts = @{ done = 7; failed = 0 } } 5 -Files 'report.csv'
        $null = New-TestJob $w '20261111-1000-exportpst-cccc' @{ jobId = 'exportpst-cccc'; task = 'exportpst'; mode = 'ExportPst'; counts = @{ done = 3; fileBytes = 777 } } 12
        $path = Write-MbtJobOverview -WorkFolder $w
        $path | Should -Be (Join-Path $w 'jobs\overview.html')
        $html = Get-Content -LiteralPath $path -Raw
        $html | Should -Not -Match '\{\{[A-Z_]+\}\}'
        $html | Should -Match '<title>Mailbox Transfer: all jobs</title>'
        $o = Get-OverviewData $html
        $o.Rows.Count | Should -Be 3
        @($o.Rows.id) | Should -Be @('20261112-1000-import-bbbb', '20261111-1000-exportpst-cccc', '20261110-1000-export-aaaa') -Because 'newest first'
        $a = $o.Rows | Where-Object id -EQ '20261110-1000-export-aaaa'
        $a.task | Should -Be 'Export'
        $a.source | Should -Be 'adele@contoso.com'
        $a.destination | Should -Be 'D:\Stores\adele'
        $a.result | Should -Be 'Complete'
        $a.done | Should -Be 100; $a.failed | Should -Be 2; $a.bytes | Should -Be 5000
        $a.seconds | Should -Be 300
        $a.link | Should -Be 1
        ([datetime]::new(1970, 1, 1, 0, 0, 0, [DateTimeKind]::Utc).AddSeconds($a.started) - $script:Now.AddHours(-30)).TotalSeconds | Should -BeLessThan 2
        $b = $o.Rows | Where-Object id -EQ '20261112-1000-import-bbbb'
        $b.task | Should -Be 'Import'; $b.result | Should -Be 'Failed'; $b.code | Should -Be 1; $b.link | Should -Be 2 -Because 'no HTML: the CSV'
        $c = $o.Rows | Where-Object id -EQ '20261111-1000-exportpst-cccc'
        $c.task | Should -Be 'Export (PST)'; $c.bytes | Should -Be 777; $c.link | Should -Be 0
        $html | Should -Match 'Jobs</div><div class="tile-value">3</div>'
        $html | Should -Match 'Failed</div><div class="tile-value">1</div>'
        $html | Should -Match 'Items done</div><div class="tile-value">110</div>'
    }

    It 'puts the jobs that need a look before the finished ones, and lets the page filter by result, task and period' {
        $w = Join-Path $TestDrive 'results'
        $cases = @(
            @{ n = 'complete'; s = @{}; code = 0 }
            @{ n = 'failed'; s = @{ state = 'error'; exitCode = 1 }; code = 1 }
            @{ n = 'stopped'; s = @{ state = 'stopped'; exitCode = 2; reason = 'stopped'; next = 'continue' }; code = 2 }
            @{ n = 'decision'; s = @{ state = 'needs-decision'; exitCode = 3; next = 'decide' }; code = 3 }
            @{ n = 'refused'; s = @{ exitCode = 1; reason = 'refused-items' }; code = 6 }
            @{ n = 'interrupted'; s = @{ state = 'running'; exitCode = $null; pid = 999999 }; code = 4 }
            @{ n = 'running'; s = @{ state = 'running'; exitCode = $null; pid = $PID; started = ((Get-Process -Id $PID).StartTime.ToUniversalTime().ToString('yyyy-MM-dd\THH:mm:ss.fff\Z')) }; code = 5 }
        )
        $i = 0
        foreach ($c in $cases) { $i++; $null = New-TestJob $w ('2026111{0}-1000-export-{1:x4}' -f $i, $i) $c.s ($i * 2) }
        $o = Get-OverviewData (Get-OverviewPage $w)
        @($o.Rows | Sort-Object { $_.id } | ForEach-Object code) | Should -Be @(0, 1, 2, 3, 6, 4, 5)
        $running = $o.Rows | Where-Object code -EQ 5
        $running.result | Should -Be 'Running'
        $o.Rows | Where-Object code -EQ 4 | ForEach-Object result | Should -Be 'Interrupted'
        $rank = @{}; foreach ($r in $o.Data.results) { $rank[[int]$r.code] = [int]$r.rank }
        $rank[1] | Should -BeLessThan $rank[3] -Because 'failed first, then the decision'
        $rank[3] | Should -BeLessThan $rank[4]
        $rank[4] | Should -BeLessThan $rank[2]
        $rank[2] | Should -BeLessThan $rank[6]
        $rank[6] | Should -BeLessThan $rank[5]
        $rank[5] | Should -BeLessThan $rank[0] -Because 'the finished jobs come last'
        @($o.Data.results | Where-Object { $_.name -eq 'Needs a decision' }).Count | Should -Be 1
        # what the filters of the page work on: every row says its task, its result code and its start
        @($o.Rows | Where-Object { -not $_.task -or $null -eq $_.started }).Count | Should -Be 0
    }

    It 'leaves out the checks of a plan, by their task, the name of their folder or their id' {
        $w = Join-Path $TestDrive 'plans'
        $null = New-TestJob $w '20261110-1000-export-aaaa'
        $null = New-TestJob $w '20261110-1100-export-plan-bbbb' @{ task = 'export-plan' }
        $null = New-TestJob $w '20261110-1200-import-plan-cccc' -NoStatus
        $null = New-TestJob $w '20261110-1300-import-plan-dddd' -StatusText '{ not json'
        $null = New-TestJob $w '20261110-1400-export-plan-eeee' @{ jobId = 'export-plan-eeee' } -Files 'report.csv'
        $o = Get-OverviewData (Get-OverviewPage $w)
        @($o.Rows.id) | Should -Be @('20261110-1000-export-aaaa')
    }

    It 'a missing, empty or damaged status.json, and a folder that is not a job, never stop it' {
        $w = Join-Path $TestDrive 'damaged'
        $null = New-TestJob $w '20261110-1000-export-aaaa'
        $null = New-TestJob $w '20261110-1100-export-bbbb' -NoStatus
        $null = New-TestJob $w '20261110-1200-export-cccc' -StatusText '{ "state": "complete", '
        $null = New-TestJob $w '20261110-1300-export-dddd' -StatusText ''
        $null = New-TestJob $w '20261110-1400-export-eeee' -StatusText '[1,2,3]'
        $null = New-TestJob $w 'some folder' -NoStatus
        { Write-MbtJobOverview -WorkFolder $w } | Should -Not -Throw
        $o = Get-OverviewData (Get-OverviewPage $w)
        $o.Rows.Count | Should -Be 6
        ($o.Rows | Where-Object id -EQ '20261110-1000-export-aaaa').result | Should -Be 'Complete' -Because 'one damaged file does not cost the others'
        foreach ($n in '20261110-1100-export-bbbb', '20261110-1200-export-cccc', '20261110-1300-export-dddd', '20261110-1400-export-eeee') {
            $r = $o.Rows | Where-Object id -EQ $n
            $r.result | Should -Be 'Unknown' -Because $n
            $r.task | Should -Be 'Export' -Because 'the task is in the name of the folder'
        }
        ($o.Rows | Where-Object id -EQ 'some folder').task | Should -BeNullOrEmpty
    }

    It 'a work folder without jobs gives an empty page, and the counts come from report.json when status.json has none' {
        $w = Join-Path $TestDrive 'empty'
        $o = Get-OverviewData (Get-OverviewPage $w)
        $o.Rows.Count | Should -Be 0
        $w2 = Join-Path $TestDrive 'reportcounts'
        $dir = New-TestJob $w2 '20261110-1000-export-aaaa' @{ counts = @{} }
        [IO.File]::WriteAllText((Join-Path $dir 'report.json'), '{ "counts": { "done": 42, "failed": 1, "bytes": 99 }, "rows": [] }')
        $r = (Get-OverviewData (Get-OverviewPage $w2)).Rows[0]
        $r.done | Should -Be 42; $r.failed | Should -Be 1; $r.bytes | Should -Be 99
    }

    It 'encodes every text of a job: markup, a closing script tag, quotes and a marker of the template are only text' {
        $w = Join-Path $TestDrive 'a&b'
        $evil = '</script><img src=x onerror=alert(1)> & "quoted" {{TITLE}} <!-- x'
        $null = New-TestJob $w '20261110-1000-export-aaaa' @{ source = $evil; destination = "D:\$evil"; batch = $evil }
        $html = Get-OverviewPage $w
        $html | Should -Not -Match ([regex]::Escape('<img src=x'))
        $html | Should -Match '<title>Mailbox Transfer: all jobs</title>' -Because 'a marker that comes from a job is not taken for one of the template'
        ([regex]::Matches($html, '</script>')).Count | Should -Be 3 -Because 'the theme, the data and the page: a job cannot add one'
        ([regex]::Matches($html, '<script')).Count | Should -Be 3
        $html | Should -Match ([regex]::Escape('a&amp;b')) -Because 'the folder in the subtitle is HTML-encoded'
        $html | Should -Match ([regex]::Escape([string][char]92 + 'u003C/script'))
        $o = Get-OverviewData $html
        $o.Rows[0].source | Should -Be $evil
        $o.Rows[0].destination | Should -Be "D:\$evil"
        $o.Rows[0].batch | Should -Be $evil
        $html | Should -Not -Match '[^\x00-\x7F]' -Because 'a text above ASCII is escaped too'
    }

    It 'links the report of each job relative to the jobs folder (or to the base the page says), built in the page from the name of the folder' {
        $t = Get-Content -LiteralPath (Join-Path $script:Root 'templates\Overview.template.html') -Raw
        $t | Should -Match ([regex]::Escape('(data.base || "./") + encodeURIComponent(r[col.id])'))
        (Get-OverviewData (Get-OverviewPage (Join-Path $TestDrive 'basic'))).Data.base | Should -Be './'
        $t | Should -Not -Match 'innerHTML|outerHTML|document\.write|insertAdjacentHTML' -Because 'no text of a job is ever put in the page as markup'
    }

    It 'lists the jobs of a batch only, whatever the case of its name, and a job without a batch is not affected' {
        $w = Join-Path $TestDrive 'batches'
        $null = New-TestJob $w '20261110-1000-export-aaaa' @{ batch = 'Wave-1' }
        $null = New-TestJob $w '20261110-1100-export-bbbb' @{ batch = 'wave-2' }
        $null = New-TestJob $w '20261110-1200-export-cccc'
        $all = Get-OverviewData (Get-OverviewPage $w)
        $all.Rows.Count | Should -Be 3
        @($all.Rows | Where-Object id -EQ '20261110-1000-export-aaaa' | ForEach-Object batch) | Should -Be @('Wave-1')
        $one = Get-Content -LiteralPath (Write-MbtJobOverview -WorkFolder $w -Batch 'WAVE-1' -FileName 'overview-wave-1.html') -Raw
        @((Get-OverviewData $one).Rows.id) | Should -Be @('20261110-1000-export-aaaa')
        $one | Should -Match '<h1>Jobs of batch WAVE-1</h1>'
        Join-Path $w 'jobs\overview-wave-1.html' | Should -Exist
        (Get-ChildItem -LiteralPath (Join-Path $w 'jobs') -Filter 'overview*.html').Count | Should -Be 2
        { Write-MbtJobOverview -WorkFolder $w -FileName '..\x.html' } | Should -Throw
        $s = Join-Path $TestDrive 'status-batch'
        Write-MbtJobStatus -JobFolder $s -State running -JobId 'j' -Batch 'Wave-9'
        (Get-Content -LiteralPath (Join-Path $s 'status.json') -Raw | ConvertFrom-Json).batch | Should -Be 'Wave-9'
        Write-MbtJobStatus -JobFolder (Join-Path $TestDrive 'status-nobatch') -State running -JobId 'j'
        (Get-Content -LiteralPath (Join-Path $TestDrive 'status-nobatch\status.json') -Raw | ConvertFrom-Json).batch | Should -BeNullOrEmpty
    }

    It 'lists the newest jobs only beyond the cap, and says so' {
        $w = Join-Path $TestDrive 'cap'
        1..5 | ForEach-Object { $null = New-TestJob $w ('2026111{0}-1000-export-{1:x4}' -f $_, $_) @{} (100 - $_) }
        $html = Get-Content -LiteralPath (Write-MbtJobOverview -WorkFolder $w -MaxJobs 3) -Raw
        $o = Get-OverviewData $html
        $o.Rows.Count | Should -Be 3
        $o.Data.left | Should -Be 2
        $html | Should -Match 'lists the 3 newest jobs\. 2 older jobs are'
    }

    It 'asked for by the folder of a job, finds the work folder; a job that is not in a folder named jobs gives nothing' {
        $w = Join-Path $TestDrive 'byjob'
        $dir = New-TestJob $w '20261110-1000-export-aaaa'
        (Write-MbtJobOverview -JobFolder $dir) | Should -Be (Join-Path $w 'jobs\overview.html')
        $elsewhere = Join-Path $TestDrive 'elsewhere\job1'
        $null = New-Item -ItemType Directory -Force -Path $elsewhere
        Write-MbtJobOverview -JobFolder $elsewhere | Should -BeNullOrEmpty
        Join-Path $TestDrive 'elsewhere\overview.html' | Should -Not -Exist
    }

    It 'keeps what it learned of the folders that did not change, and sees one that did, one that is new and one that is gone' {
        $w = Join-Path $TestDrive 'cache'
        $a = New-TestJob $w '20261110-1000-export-aaaa'
        $null = New-TestJob $w '20261110-1100-export-bbbb'
        $null = Write-MbtJobOverview -WorkFolder $w
        Join-Path $w 'jobs\overview.cache.json' | Should -Exist
        # a job ends (status.json is replaced, as Write-MbtJobStatus does): its folder changed
        Write-MbtJobStatus -JobFolder $a -State error -ExitCode 1 -Counts @{ done = 5 }
        $null = New-TestJob $w '20261110-1200-export-cccc'
        Remove-Item -LiteralPath (Join-Path $w 'jobs\20261110-1100-export-bbbb') -Recurse -Force
        $o = Get-OverviewData (Get-OverviewPage $w)
        @($o.Rows.id | Sort-Object) | Should -Be @('20261110-1000-export-aaaa', '20261110-1200-export-cccc')
        ($o.Rows | Where-Object id -EQ '20261110-1000-export-aaaa').result | Should -Be 'Failed'
        # the same rows, read from the cache or from the folders
        $cached = (Get-Content -LiteralPath (Write-MbtJobOverview -WorkFolder $w) -Raw) -replace 'Written [^.]*\.', ''
        $fresh = (Get-Content -LiteralPath (Write-MbtJobOverview -WorkFolder $w -NoCache) -Raw) -replace 'Written [^.]*\.', ''
        $cached | Should -Be $fresh
        Set-Content -LiteralPath (Join-Path $w 'jobs\overview.cache.json') -Value 'not json'
        { Write-MbtJobOverview -WorkFolder $w } | Should -Not -Throw
        (Get-OverviewData (Get-Content -LiteralPath (Join-Path $w 'jobs\overview.html') -Raw)).Rows.Count | Should -Be 2
    }

    It 'writes the page whole: no temporary file is left, a reader that holds the file for a moment is waited for' {
        $w = Join-Path $TestDrive 'atomic'
        $null = New-TestJob $w '20261110-1000-export-aaaa'
        $path = Write-MbtJobOverview -WorkFolder $w
        Set-Content -LiteralPath $path -Value 'old'
        $path = Write-MbtJobOverview -WorkFolder $w
        (Get-Content -LiteralPath $path -Raw) | Should -Match '</html>\s*$'
        @(Get-ChildItem -LiteralPath (Join-Path $w 'jobs') -Filter '*.tmp').Count | Should -Be 0
        # another process holds the page open without allowing a replace, for about a second
        $flag = Join-Path $TestDrive 'held.flag'
        $hold = "`$s = [IO.File]::Open('$path', 'Open', 'Read', 'Read'); Set-Content '$flag' 'held'; Start-Sleep -Milliseconds 1200; `$s.Dispose()"
        $child = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList @('-NoProfile', '-Command', $hold) -PassThru -WindowStyle Hidden
        $waited = [Diagnostics.Stopwatch]::StartNew()
        while (-not (Test-Path -LiteralPath $flag) -and $waited.Elapsed.TotalSeconds -lt 30) { Start-Sleep -Milliseconds 50 }
        { Write-MbtJobOverview -WorkFolder $w } | Should -Not -Throw
        $child.WaitForExit(10000) | Out-Null
        (Get-Content -LiteralPath $path -Raw) | Should -Match '</html>\s*$'
        @(Get-ChildItem -LiteralPath (Join-Path $w 'jobs') -Filter '*.tmp').Count | Should -Be 0
    }

    It 'several jobs that end together leave one whole page' {
        $w = Join-Path $TestDrive 'parallel'
        1..4 | ForEach-Object { $null = New-TestJob $w ('2026111{0}-1000-export-{1:x4}' -f $_, $_) }
        $module = Join-Path $script:Root 'MailboxTransfer.psd1'
        $null = 1..6 | ForEach-Object -ThrottleLimit 6 -Parallel {
            Import-Module $using:module -Force
            $env:MAILBOXTRANSFER_LOCAL_CONFIG = Join-Path $using:TestDrive 'none.psd1'
            Write-MbtJobOverview -WorkFolder $using:w
        }
        $html = Get-Content -LiteralPath (Join-Path $w 'jobs\overview.html') -Raw
        $html | Should -Match '</html>\s*$'
        (Get-OverviewData $html).Rows.Count | Should -Be 4
        @(Get-ChildItem -LiteralPath (Join-Path $w 'jobs') -Filter '*.tmp').Count | Should -Be 0
        { Get-Content -LiteralPath (Join-Path $w 'jobs\overview.cache.json') -Raw | ConvertFrom-Json } | Should -Not -Throw
    }
}

Describe 'The shipped overview template' {
    It 'is in the package, needs no network and has every marker once' {
        $file = Join-Path $script:Root 'templates\Overview.template.html'
        $file | Should -Exist
        $t = Get-Content -LiteralPath $file -Raw
        $t | Should -Not -Match '(src|href)\s*=\s*"https?:'
        $t | Should -Not -Match 'url\(\s*https?:'
        foreach ($marker in 'TITLE', 'HEADING', 'VERSION', 'SUBTITLE', 'TILES', 'CAP', 'GENERATED', 'DATA') {
            [regex]::Matches($t, "\{\{$marker\}\}").Count | Should -BeGreaterThan 0 -Because $marker
        }
        $t | Should -Match '--cp-accent'
        $t | Should -Match 'data-theme="dark"'
        $r = Get-Content -LiteralPath (Join-Path $script:Root 'templates\Report.template.html') -Raw
        foreach ($token in 'cp-bg', 'cp-surface', 'cp-border', 'cp-text', 'cp-accent', 'cp-success', 'cp-danger', 'cp-warning', 'cp-link') {
            ([regex]::Match($t, "--$token`: ([^;]+);").Groups[1].Value) | Should -Be ([regex]::Match($r, "--$token`: ([^;]+);").Groups[1].Value) -Because "the colour $token is the one of the report"
        }
    }
    It 'falls back to a plain page when the template is missing' {
        $w = Join-Path $TestDrive 'notemplate'
        $null = New-TestJob $w '20261110-1000-export-aaaa'
        $path = Write-MbtJobOverview -WorkFolder $w -TemplatePath (Join-Path $TestDrive 'missing.html')
        (Get-Content -LiteralPath $path -Raw) | Should -Match '<h1>Mailbox Transfer: all jobs</h1>'
    }
    It 'is a script that a browser runs without an error, when node is there to try it' -Skip:(-not (Get-Command node -ErrorAction SilentlyContinue)) {
        $w = Join-Path $TestDrive 'js'
        1..30 | ForEach-Object { $null = New-TestJob $w ('2026111{0:x}-1000-export-{1:x4}' -f ($_ % 9), $_) @{ source = "user$_@contoso.com" } $_ }
        $html = Get-OverviewPage $w
        $script = [regex]::Matches($html, '<script>(.*?)</script>', 'Singleline')[1].Groups[1].Value
        $data = [regex]::Match($html, '<script type="application/json" id="mbt-overview-data">(.*?)</script>', 'Singleline').Groups[1].Value
        $harness = @"
const data = $($data | ConvertTo-Json -Compress);
const els = {};
function el(tag) { const e = { tag, children: [], dataset: {}, style: {}, attrs: {}, listeners: {}, className: '', textContent: '', hidden: false, value: '',
  appendChild(c) { this.children.push(c); return c; }, setAttribute(k, v) { this.attrs[k] = v; }, addEventListener(t, f) { this.listeners[t] = f; }, querySelector() { return el('button'); } };
  Object.defineProperty(e, 'textContent', { get() { return this._t || ''; }, set(v) { this._t = v; if (v === '') this.children = []; } }); return e; }
['mbt-overview-data','chips','find','task','period','batch','batchField','reset','shown','rows','more'].forEach(id => { els[id] = el('x'); });
els['mbt-overview-data'].textContent = data;
global.window = { location: { search: '' }, matchMedia: () => ({ matches: false }) };
global.document = { documentElement: el('html'), getElementById: id => els[id], createElement: el, createTextNode: t => ({ t }), querySelectorAll: () => [] };
global.URLSearchParams = URLSearchParams; global.Option = function (t, v) { return { text: t, value: v }; }; global.Intl = Intl;
$script
console.log('rows', els.rows.children.length, 'chips', els.chips.children.length);
"@
        $file = Join-Path $TestDrive 'harness.js'
        Set-Content -LiteralPath $file -Value $harness -Encoding utf8
        $out = & node $file 2>&1 | Out-String
        $LASTEXITCODE | Should -Be 0 -Because $out
        $out | Should -Match 'rows 30 chips \d+'
    }
}

Describe 'The time budget of the overview' {
    BeforeAll {
        # every block of job folders takes 600 ms to read, as a virus scanner can make it: the build of the page cannot wait for ever
        Mock -ModuleName MailboxTransfer Read-MbtOverviewStatusBlock { Start-Sleep -Milliseconds 600; return @{} }
        function New-ManyJobs([string]$Work, [int]$Count) {
            foreach ($i in 1..$Count) { $null = New-TestJob $Work ('20261110-1000-export-{0:x4}' -f $i) @{ jobId = "export-$i" } 20 }
        }
    }
    It 'stops after the budget, keeps the page that is there, says how many job folders are left, and keeps what it read for the next build' {
        $w = Join-Path $TestDrive 'budget'
        New-ManyJobs $w 300
        $page = Join-Path $w 'jobs\overview.html'
        Set-Content -LiteralPath $page -Value 'the page of the last build'
        $info = @{}
        $path = Write-MbtJobOverview -WorkFolder $w -TimeBudgetSeconds 1 -Info $info
        $path | Should -BeNullOrEmpty
        $info.Complete | Should -BeFalse
        $info.Folders | Should -Be 300
        $info.Unread | Should -BeGreaterThan 100
        $info.Unread | Should -BeLessThan 300
        (Get-Content -LiteralPath $page -Raw).Trim() | Should -Be 'the page of the last build'
        $cache = Get-Content -LiteralPath (Join-Path $w 'jobs\overview.cache.json') -Raw | ConvertFrom-Json
        @($cache.rows).Count | Should -Be (300 - $info.Unread)
    }
    It 'is finished by several short builds: each one goes on where the last one stopped, and the page is then the one a build without a budget gives' {
        $w = Join-Path $TestDrive 'budget-converge'
        New-ManyJobs $w 200
        $builds = 0
        $info = @{}
        do { $builds++; $info = @{}; $path = Write-MbtJobOverview -WorkFolder $w -TimeBudgetSeconds 1 -Info $info } while (-not $info.Complete -and $builds -lt 12)
        $info.Complete | Should -BeTrue
        $builds | Should -BeGreaterThan 1
        $path | Should -Be (Join-Path $w 'jobs\overview.html')
        $page = (Get-Content -LiteralPath $path -Raw) -replace 'Written [^.]*\.', ''
        $fresh = (Get-Content -LiteralPath (Write-MbtJobOverview -WorkFolder $w -NoCache) -Raw) -replace 'Written [^.]*\.', ''
        $page | Should -Be $fresh
    }
    It 'a build without a budget reads everything and writes the page' {
        $w = Join-Path $TestDrive 'budget-none'
        New-ManyJobs $w 60
        $info = @{}
        $path = Write-MbtJobOverview -WorkFolder $w -Info $info
        $path | Should -Be (Join-Path $w 'jobs\overview.html')
        $info.Complete | Should -BeTrue
        $info.Unread | Should -Be 0
    }
}

Describe 'A folder of 5,000 jobs' {
    BeforeAll {
        $script:Big = Join-Path $TestDrive 'big'
        $jobs = Join-Path $script:Big 'jobs'
        $null = [IO.Directory]::CreateDirectory($jobs)
        $states = @(@('complete', 0, 'null', 'done'), @('complete', 0, 'null', 'done'), @('complete', 0, 'null', 'done'), @('error', 1, 'null', 'retry'), @('resumable', 2, '"stopped"', 'continue'), @('needs-decision', 3, 'null', 'decide'))
        $rnd = [Random]::new(7)
        for ($i = 0; $i -lt 5000; $i++) {
            $at = $script:Now.AddMinutes(-$i * 7)
            $st = $states[$rnd.Next($states.Count)]
            $plan = ($i % 8 -eq 0)
            $task = if ($plan) { 'export-plan' } else { 'export' }
            $dir = '{0}\{1:yyyyMMdd-HHmm}-{2}-{3:x4}' -f $jobs, $at.ToLocalTime(), $task, $i
            [void][IO.Directory]::CreateDirectory($dir)
            $json = '{"v":1,"jobId":"export-%i","task":"%task","title":"Export - user%m@contoso.com","mode":"Export","store":"D:\\Stores\\u","source":"user%m@contoso.com","destination":"D:\\Stores\\u","state":"%s","exitCode":%e,"counts":{"done":%d,"failed":0,"bytes":%b},"next":"%n","reason":%r,"started":"%t","updated":"%u","pid":4}'
            $json = $json.Replace('%i', "$i").Replace('%task', $task).Replace('%m', "$($i % 300)").Replace('%s', $st[0]).Replace('%e', "$($st[1])").Replace('%d', "$($i * 3)").Replace('%b', "$($i * 4000)").Replace('%n', $st[3]).Replace('%r', $st[2]).Replace('%t', $at.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z')).Replace('%u', $at.AddMinutes(4).ToString('yyyy-MM-dd\THH:mm:ss.fff\Z'))
            [IO.File]::WriteAllText("$dir\status.json", $json)
            if ($i % 2 -eq 0) { [IO.File]::WriteAllText("$dir\report.html", 'x') }
        }
        $script:Cold = Measure-Command { $script:BigPath = Write-MbtJobOverview -WorkFolder $script:Big }
        $script:Warm = Measure-Command { $null = Write-MbtJobOverview -WorkFolder $script:Big }
    }
    It 'is built in seconds, and again in less when nothing changed, in a page of well under a megabyte' {
        $script:Cold.TotalSeconds | Should -BeLessThan 180
        $script:Warm.TotalSeconds | Should -BeLessThan 60
        (Get-Item -LiteralPath $script:BigPath).Length | Should -BeLessThan 1MB
    }
    It 'lists every job that is not a plan check' {
        $o = Get-OverviewData (Get-Content -LiteralPath $script:BigPath -Raw)
        $o.Rows.Count | Should -Be (5000 - 625)
        $o.Data.strings.Count | Should -BeLessThan 2000 -Because 'the strings the rows share are written once'
        @($o.Rows.started | Sort-Object -Descending) | Should -Be @($o.Rows.started)
    }
}
