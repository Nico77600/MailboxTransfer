#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the retention of the job folders (Invoke-MbtJobRetention, the configuration key Jobs.RetentionDays, the command line and the
    console that apply it). The jobs are synthetic job folders; nothing reaches the network.

    Run:  Invoke-Pester -Path .\tests\JobRetention.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:Root = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    $script:SavedLocal = $env:MAILBOXTRANSFER_LOCAL_CONFIG
    $env:MAILBOXTRANSFER_LOCAL_CONFIG = Join-Path $TestDrive 'no-local-config.psd1'
    Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
    $script:Now = [datetime]::UtcNow

    # A job folder in <work>\jobs, last touched DaysAgo days ago (its status.json says so too).
    function New-OldJob {
        param([string]$Work, [string]$Name, [double]$DaysAgo, [string]$State = 'complete', [string]$Request, [string[]]$Files = @(), [switch]$NoStatus, [string]$StatusText)
        $dir = Join-Path (Join-Path $Work 'jobs') $Name
        $null = New-Item -ItemType Directory -Force -Path $dir
        $at = $script:Now.AddDays(-$DaysAgo)
        $stamp = { param($d) $d.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z') }
        $exit = switch ($State) { 'complete' { 0 } 'error' { 1 } 'needs-decision' { 3 } 'resumable' { 2 } 'stopped' { 2 } default { $null } }
        $s = [ordered]@{ v = 1; jobId = $Name; task = 'export'; title = 'Export - adele@contoso.com'; mode = 'Export'; state = $State; exitCode = $exit; counts = @{ done = 1 }; started = (& $stamp $at.AddMinutes(-5)); updated = (& $stamp $at); pid = 999999 }
        if ($PSBoundParameters.ContainsKey('StatusText')) { [IO.File]::WriteAllText((Join-Path $dir 'status.json'), $StatusText) }
        elseif (-not $NoStatus) { [IO.File]::WriteAllText((Join-Path $dir 'status.json'), ($s | ConvertTo-Json -Depth 4)) }
        if ($Request) { [IO.File]::WriteAllText((Join-Path $dir 'request.json'), $Request) }
        foreach ($f in $Files) { $p = Join-Path $dir $f; $null = New-Item -ItemType Directory -Force -Path (Split-Path $p -Parent); [IO.File]::WriteAllText($p, 'x') }
        (Get-Item -LiteralPath $dir).LastWriteTimeUtc = $at
        return $dir
    }
    function Get-JobNames([string]$Work) { @(Get-ChildItem -LiteralPath (Join-Path $Work 'jobs') -Directory | ForEach-Object Name | Sort-Object) }
    function ConvertTo-JsonPath([string]$Path) { $Path.Replace('\', '\\') }
}
AfterAll { $env:MAILBOXTRANSFER_LOCAL_CONFIG = $script:SavedLocal }

Describe 'The configuration key Jobs.RetentionDays' {
    BeforeAll {
        function Read-Config([string]$Text) {
            $file = Join-Path $TestDrive ('cfg-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.psd1')
            Set-Content -LiteralPath $file -Value $Text -Encoding UTF8
            Import-MbtConfiguration -Path $file
        }
    }
    It 'is 90 days by default, in the code and in the shipped file' {
        (Import-MbtConfiguration).Jobs.RetentionDays | Should -Be 90
        $shipped = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Root 'config\MailboxTransfer.config.psd1')
        $shipped.Jobs.RetentionDays | Should -Be 90
    }
    It 'reads a number of days, and 0 for keeping everything' {
        (Read-Config '@{ Jobs = @{ RetentionDays = 30 } }').Jobs.RetentionDays | Should -Be 30
        (Read-Config '@{ Jobs = @{ RetentionDays = 0 } }').Jobs.RetentionDays | Should -Be 0
        (Read-Config '@{ Jobs = @{ RetentionDays = 3650 } }').Jobs.RetentionDays | Should -Be 3650
        (Read-Config "@{ Jobs = @{ RetentionDays = '45' } }").Jobs.RetentionDays | Should -Be 45
    }
    It 'refuses what is not a number of days between 0 and 3650, and a key it does not know' {
        foreach ($bad in '-1', '3651', "'many'", '1.5', '$null', "''") {
            { Read-Config "@{ Jobs = @{ RetentionDays = $bad } }" } | Should -Throw '*Jobs.RetentionDays*' -Because $bad
        }
        { Read-Config '@{ Jobs = @{ Keep = 1 } }' } | Should -Throw '*unknown key*'
    }
}

Describe 'The checks of a plan' {
    It 'a plan check is removed after a day, whatever its state, interrupted or never finished included; a younger one and every other job stay' {
        $w = Join-Path $TestDrive 'plans'
        $null = New-OldJob $w '20261101-1000-export-plan-aa01' 2
        $null = New-OldJob $w '20261101-1001-export-plan-aa02' 2 -State error
        $null = New-OldJob $w '20261101-1002-import-plan-aa03' 3 -State running
        $null = New-OldJob $w '20261101-1003-export-plan-aa04' 5 -State resumable
        $null = New-OldJob $w '20261101-1004-export-plan-aa05' 40 -State stopped
        $null = New-OldJob $w '20261101-1005-export-plan-aa06' 2 -NoStatus
        $null = New-OldJob $w '20261101-1006-export-plan-aa07' 2 -StatusText 'not json'
        $null = New-OldJob $w '20261114-0900-export-plan-bb01' 0.5
        $null = New-OldJob $w '20261101-1010-export-aa08' 2
        $null = New-OldJob $w '20261101-1011-export-aa09' 20 -State stopped
        $null = New-OldJob $w '20261101-1012-export-planet-aa0a' 5
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 90
        @($r.Removed | Sort-Object) | Should -Be @('20261101-1000-export-plan-aa01', '20261101-1001-export-plan-aa02', '20261101-1002-import-plan-aa03', '20261101-1003-export-plan-aa04', '20261101-1004-export-plan-aa05', '20261101-1005-export-plan-aa06', '20261101-1006-export-plan-aa07')
        Get-JobNames $w | Should -Be @('20261101-1010-export-aa08', '20261101-1011-export-aa09', '20261101-1012-export-planet-aa0a', '20261114-0900-export-plan-bb01')
    }
    It 'the retention of 0 days keeps everything, the checks of a plan too; a plan that a job that stays names in its request stays' {
        $w = Join-Path $TestDrive 'plans-keep'
        $null = New-OldJob $w '20261101-1000-export-plan-cc01' 9
        (Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 0).Removed.Count | Should -Be 0
        Get-JobNames $w | Should -Be @('20261101-1000-export-plan-cc01')
        $kept = New-OldJob $w '20261101-1001-export-cc02' 1 -State stopped -Request ('{"StorePath":"' + (ConvertTo-JsonPath (Join-Path $w 'jobs\20261101-1000-export-plan-cc01\temp')) + '"}')
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 90
        @($r.Removed).Count | Should -Be 0
        @($r.Protected) | Should -Contain '20261101-1000-export-plan-cc01'
    }
}

Describe 'Removing the folders of old jobs' {
    It 'removes a finished job that is older than the retention, and keeps a young one, whatever its result' {
        $w = Join-Path $TestDrive 'basic'
        $null = New-OldJob $w 'old-complete' 120
        $null = New-OldJob $w 'old-error' 100 -State error
        $null = New-OldJob $w 'young-complete' 10
        $null = New-OldJob $w 'young-error' 89 -State error
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 90
        @($r.Removed | Sort-Object) | Should -Be @('old-complete', 'old-error')
        Get-JobNames $w | Should -Be @('young-complete', 'young-error')
        $r.RetentionDays | Should -Be 90
        @($r.Failed).Count | Should -Be 0
    }

    It 'uses the retention of the configuration when none is given, and 0 keeps every job' {
        $w = Join-Path $TestDrive 'config'
        $null = New-OldJob $w 'old' 200
        $null = New-OldJob $w 'older-than-config' 45
        $null = New-OldJob $w 'young' 10
        (Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 0).Removed.Count | Should -Be 0
        Get-JobNames $w | Should -Be @('old', 'older-than-config', 'young')
        $file = Join-Path $TestDrive 'retention-30.psd1'
        Set-Content -LiteralPath $file -Value '@{ Jobs = @{ RetentionDays = 30 } }'
        $env:MAILBOXTRANSFER_LOCAL_CONFIG = $file
        try { $r = Invoke-MbtJobRetention -WorkFolder $w } finally { $env:MAILBOXTRANSFER_LOCAL_CONFIG = Join-Path $TestDrive 'no-local-config.psd1' }
        $r.RetentionDays | Should -Be 30
        Get-JobNames $w | Should -Be @('young')
        $w2 = Join-Path $TestDrive 'config-default'
        $null = New-OldJob $w2 'old' 200
        $null = New-OldJob $w2 'between' 95
        $null = New-OldJob $w2 'young' 80
        (Invoke-MbtJobRetention -WorkFolder $w2).Removed.Count | Should -Be 2 -Because 'the default is 90 days'
        Get-JobNames $w2 | Should -Be @('young')
    }

    It 'never removes a job that can still be continued, however old: stopped, not complete, waiting for a decision, running, interrupted, without a result' {
        $w = Join-Path $TestDrive 'continuable'
        $null = New-OldJob $w 'stopped' 400 -State stopped
        $null = New-OldJob $w 'resumable' 400 -State resumable
        $null = New-OldJob $w 'decision' 400 -State needs-decision
        $null = New-OldJob $w 'running-dead' 400 -State running
        $null = New-OldJob $w 'unknown-state' 400 -State weird
        $null = New-OldJob $w 'old-complete' 400
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30
        @($r.Removed) | Should -Be @('old-complete')
        Get-JobNames $w | Should -Be @('decision', 'resumable', 'running-dead', 'stopped', 'unknown-state')
    }

    It 'never removes a folder whose status.json is missing, empty or damaged, or has no date' {
        $w = Join-Path $TestDrive 'unreadable'
        $null = New-OldJob $w 'no-status' 400 -NoStatus
        $null = New-OldJob $w 'empty' 400 -StatusText ''
        $null = New-OldJob $w 'damaged' 400 -StatusText '{ "state": "complete", '
        $null = New-OldJob $w 'no-date' 400 -StatusText '{ "state": "complete", "exitCode": 0 }'
        $null = New-OldJob $w 'array' 400 -StatusText '["complete"]'
        (Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30).Removed.Count | Should -Be 0
        (Get-JobNames $w).Count | Should -Be 5
    }

    It 'keeps the job that starts, a file of the jobs folder, and a link to somewhere else with what it points to' {
        $w = Join-Path $TestDrive 'special'
        $current = New-OldJob $w 'current' 400
        $null = New-OldJob $w 'old' 400
        $jobs = Join-Path $w 'jobs'
        [IO.File]::WriteAllText((Join-Path $jobs 'overview.html'), 'page')
        $outside = New-OldJob (Join-Path $TestDrive 'special-outside') 'precious' 400 -Files 'keep.txt'
        $link = Join-Path $jobs 'linked'
        New-Item -ItemType Junction -Path $link -Target $outside | Out-Null
        $r = Invoke-MbtJobRetention -JobFolder $current -RetentionDays 30
        @($r.Removed) | Should -Be @('old')
        Get-JobNames $w | Should -Be @('current', 'linked')
        (Join-Path $outside 'keep.txt') | Should -Exist
        (Join-Path $jobs 'overview.html') | Should -Exist -Because 'the overview is a file of the jobs folder, never a job'
    }

    It 'never leaves the jobs folder: the stores, the selections, the outputs and the other folders of the work folder stay' {
        $w = Join-Path $TestDrive 'outside'
        $outDir = Join-Path $TestDrive 'outside-documents'
        $paths = @('stores\default\catalog.db', 'selections\export-1234abcd.json', 'logs\MailboxTransfer_20250101.log', 'token.cache') | ForEach-Object { Join-Path $w $_ }
        foreach ($p in $paths) { $null = New-Item -ItemType Directory -Force -Path (Split-Path $p -Parent); [IO.File]::WriteAllText($p, 'x'); (Get-Item -LiteralPath $p).LastWriteTimeUtc = $script:Now.AddDays(-900) }
        $null = New-Item -ItemType Directory -Force -Path $outDir
        [IO.File]::WriteAllText((Join-Path $outDir 'export.pst'), 'pst'); [IO.File]::WriteAllText((Join-Path $outDir 'item.eml'), 'eml')
        $request = '{ "StorePath": "' + (ConvertTo-JsonPath (Join-Path $w 'stores\default')) + '", "SelectionFiles": [ "' + (ConvertTo-JsonPath (Join-Path $w 'selections\export-1234abcd.json')) + '" ], "OutFolder": "' + (ConvertTo-JsonPath $outDir) + '", "OutName": "export.pst" }'
        $null = New-OldJob $w 'old-export' 400 -Request $request -Files 'temp\part-1.bin', 'report.csv'
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30
        @($r.Removed) | Should -Be @('old-export')
        foreach ($p in $paths) { $p | Should -Exist }
        (Join-Path $outDir 'export.pst') | Should -Exist
        (Join-Path $outDir 'item.eml') | Should -Exist
        Join-Path $w 'jobs' | Should -Exist -Because 'the jobs folder itself stays, empty'
        (Get-JobNames $w).Count | Should -Be 0
    }

    It 'keeps a folder that a job that stays names in its request, and a job that wrote something into its own folder' {
        $w = Join-Path $TestDrive 'named'
        $jobs = Join-Path $w 'jobs'
        $null = New-OldJob $w 'old-used' 400 -Files 'store\catalog.db'
        $null = New-OldJob $w 'old-own-output' 400 -Request ('{ "OutFolder": "' + (ConvertTo-JsonPath (Join-Path $jobs 'old-own-output\pst')) + '" }') -Files 'pst\export.pst'
        $null = New-OldJob $w 'old-standard' 400 -Request ('{ "TempPath": "' + (ConvertTo-JsonPath (Join-Path $jobs 'old-standard\temp')) + '", "Events": "' + (ConvertTo-JsonPath (Join-Path $jobs 'old-standard\events.jsonl')) + '" }') -Files 'temp\x.bin'
        $null = New-OldJob $w 'old-free' 400
        $null = New-OldJob $w 'continuable' 5 -State stopped -Request ('{ "StorePath": "' + (ConvertTo-JsonPath (Join-Path $jobs 'old-used\store')) + '" }')
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30
        @($r.Removed | Sort-Object) | Should -Be @('old-free', 'old-standard')
        @($r.Protected | Sort-Object) | Should -Be @('old-own-output', 'old-used')
        Get-JobNames $w | Should -Be @('continuable', 'old-own-output', 'old-used')
        (Join-Path $jobs 'old-used\store\catalog.db') | Should -Exist
        (Join-Path $jobs 'old-own-output\pst\export.pst') | Should -Exist
    }

    It 'does not follow a link inside a job folder: the folder it points to stays' {
        $w = Join-Path $TestDrive 'inner-link'
        $target = Join-Path $TestDrive 'inner-link-target'
        $null = New-Item -ItemType Directory -Force -Path $target
        [IO.File]::WriteAllText((Join-Path $target 'data.txt'), 'data')
        $dir = New-OldJob $w 'old' 400 -Files 'temp\x.bin'
        New-Item -ItemType Junction -Path (Join-Path $dir 'temp\link') -Target $target | Out-Null
        (Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30).Removed | Should -Be @('old')
        (Join-Path $target 'data.txt') | Should -Exist
        Join-Path $dir 'x' | Should -Not -Exist
    }

    It 'follows a link inside a job folder only when its reparse tag is a cloud tag: a tag that redirects or is not known is a link, a cloud folder is an ordinary one' {
        $w = Join-Path $TestDrive 'tags'
        $cloudTarget = Join-Path $TestDrive 'tags-cloud-target'
        $otherTarget = Join-Path $TestDrive 'tags-other-target'
        foreach ($target in $cloudTarget, $otherTarget) { $null = New-Item -ItemType Directory -Force -Path $target; [IO.File]::WriteAllText((Join-Path $target 'data.txt'), 'data') }
        $cloud = New-OldJob $w 'old-cloud' 400 -Files 'temp\x.bin'
        $other = New-OldJob $w 'old-other' 400 -Files 'temp\y.bin'
        New-Item -ItemType Junction -Path (Join-Path $cloud 'temp\link') -Target $cloudTarget | Out-Null
        New-Item -ItemType Junction -Path (Join-Path $other 'temp\link') -Target $otherTarget | Out-Null
        Mock -ModuleName MailboxTransfer Get-PathReparseTag { [uint32]0x9000001Au } -ParameterFilter { $Path -like '*\old-cloud\temp\link' }
        Mock -ModuleName MailboxTransfer Get-PathReparseTag { [uint32]0x8000001Bu } -ParameterFilter { $Path -like '*\old-other\temp\link' }
        @((Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30).Removed | Sort-Object) | Should -Be @('old-cloud', 'old-other')
        Join-Path $cloudTarget 'data.txt' | Should -Not -Exist
        Join-Path $otherTarget 'data.txt' | Should -Exist
    }

    It 'judges a folder entry by its reparse tag: a normal folder and a cloud folder are not links; a tag that redirects, an unknown tag and a tag that cannot be read are' {
        $target = Join-Path $TestDrive 'entry-target'
        $null = New-Item -ItemType Directory -Force -Path $target
        $plain = Join-Path $TestDrive 'entry-plain'
        $null = New-Item -ItemType Directory -Force -Path $plain
        $link = Join-Path $TestDrive 'entry-link'
        New-Item -ItemType Junction -Path $link -Target $target | Out-Null
        InModuleScope MailboxTransfer -Parameters @{ Plain = $plain; Link = $link } {
            param($Plain, $Link)
            Test-MbtFolderLink ([IO.DirectoryInfo]::new($Plain)) | Should -BeFalse
            Test-MbtFolderLink ([IO.DirectoryInfo]::new($Link)) | Should -BeTrue -Because 'a junction has a tag that redirects'
        }
        foreach ($case in @(@{ Tag = 0x9000001Au; Link = $false }, @{ Tag = 0x9000F01Au; Link = $false }, @{ Tag = 0x80000013u; Link = $false }, @{ Tag = 0xA000000Cu; Link = $true }, @{ Tag = 0x8000001Bu; Link = $true })) {
            Mock -ModuleName MailboxTransfer Get-PathReparseTag ([scriptblock]::Create('[uint32]0x' + $case.Tag.ToString('X8') + 'u'))
            InModuleScope MailboxTransfer -Parameters @{ Link = $link; Expected = $case.Link } {
                param($Link, $Expected)
                Test-MbtFolderLink ([IO.DirectoryInfo]::new($Link)) | Should -Be $Expected
            }
        }
        Mock -ModuleName MailboxTransfer Get-PathReparseTag { throw 'cannot be read' }
        InModuleScope MailboxTransfer -Parameters @{ Link = $link } {
            param($Link)
            Test-MbtFolderLink ([IO.DirectoryInfo]::new($Link)) | Should -BeTrue
        }
    }
    It 'says what it would remove with -WhatIf and removes nothing' {
        $w = Join-Path $TestDrive 'whatif'
        $null = New-OldJob $w 'old' 400
        $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30 -WhatIf
        @($r.Removed) | Should -Be @('old')
        Get-JobNames $w | Should -Be @('old')
    }

    It 'goes on with the other folders when one cannot be removed, and says which' {
        $w = Join-Path $TestDrive 'locked'
        $locked = New-OldJob $w 'locked' 400 -Files 'temp\busy.bin'
        $null = New-OldJob $w 'free' 400
        $held = [IO.File]::Open((Join-Path $locked 'temp\busy.bin'), 'Open', 'Read', 'None')
        try { $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30 } finally { $held.Dispose() }
        @($r.Removed) | Should -Be @('free')
        @($r.Failed) | Should -Be @('locked')
        (Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30).Removed | Should -Be @('locked') -Because 'the next job tries again'
    }

    It 'logs what it removed, and writes the overview again without it' {
        $w = Join-Path $TestDrive 'logged'
        $null = New-OldJob $w 'old-one' 400
        $null = New-OldJob $w 'old-two' 400 -State error
        $keep = New-OldJob $w 'young' 1
        $null = Write-MbtJobOverview -WorkFolder $w
        (Get-Content -LiteralPath (Join-Path $w 'jobs\overview.html') -Raw) | Should -Match 'old-one'
        $log = Start-MbtLog -Directory (Join-Path $w 'logs')
        try { $r = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30 } finally { Stop-MbtLog }
        $text = Get-Content -LiteralPath $log -Raw
        $text | Should -Match 'Job folder removed \(finished, older than 30 days\): .*old-one'
        $text | Should -Match 'Job folder removed \(finished, older than 30 days\): .*old-two'
        $text | Should -Not -Match 'young'
        $r.Overview | Should -Be (Join-Path $w 'jobs\overview.html')
        $page = Get-Content -LiteralPath $r.Overview -Raw
        $page | Should -Not -Match 'old-one'
        $page | Should -Match 'young'
        $nothing = Invoke-MbtJobRetention -WorkFolder $w -RetentionDays 30
        $nothing.Overview | Should -BeNullOrEmpty -Because 'nothing was removed: the overview is left as it is'
    }

    It 'a work folder without jobs, and a folder that is not in a folder named jobs, are left alone' {
        (Invoke-MbtJobRetention -WorkFolder (Join-Path $TestDrive 'nothing-here') -RetentionDays 30).Removed.Count | Should -Be 0
        Join-Path $TestDrive 'nothing-here\jobs' | Should -Not -Exist
        $elsewhere = Join-Path $TestDrive 'elsewhere\job1'
        $null = New-Item -ItemType Directory -Force -Path $elsewhere
        Invoke-MbtJobRetention -JobFolder $elsewhere -RetentionDays 30 | Should -BeNullOrEmpty
        $elsewhere | Should -Exist
    }
}

Describe 'The command line and the console apply the retention' {
    It 'the command line removes old finished jobs when a job starts, says so, and ends with the All jobs line' {
        $w = Join-Path $TestDrive 'cli'
        $null = New-OldJob $w 'old-complete' 400
        $null = New-OldJob $w 'old-stopped' 400 -State stopped
        $null = New-OldJob $w 'young' 3
        $cfg = Join-Path $TestDrive 'cli-config.psd1'
        Set-Content -LiteralPath $cfg -Value '@{ Jobs = @{ RetentionDays = 30 } }'
        $store = Join-Path $TestDrive 'cli-nostore'
        $ev = Join-Path $w 'jobs\20261114-1000-status-aaaa\events.jsonl'
        $out = & pwsh -NoProfile -NonInteractive -File $script:Entry -DevelopmentMode -ConfigPath $cfg -Mode CheckPst -Path (Join-Path $TestDrive 'missing.pst') -EventsFile $ev 2>&1 | Out-String
        $out | Should -Match 'removed 1 old job folder \(finished more than 30 days ago\)'
        Get-JobNames $w | Should -Be @('20261114-1000-status-aaaa', 'old-stopped', 'young')
        $overview = Join-Path $w 'jobs\overview.html'
        $overview | Should -Exist
        $out | Should -Match ('(?m)^\s+\S*\s*All jobs\s+' + [regex]::Escape($overview))
        $page = Get-Content -LiteralPath $overview -Raw
        $page | Should -Match '20261114-1000-status-aaaa'
        $page | Should -Not -Match 'old-complete'
    }

    It 'a job whose folder is not in a jobs folder writes no overview and removes nothing' {
        $w = Join-Path $TestDrive 'cli-plain'
        $null = New-OldJob $w 'old-complete' 400
        $ev = Join-Path $w 'a-job\events.jsonl'
        $out = & pwsh -NoProfile -NonInteractive -File $script:Entry -DevelopmentMode -Mode CheckPst -Path (Join-Path $TestDrive 'missing.pst') -EventsFile $ev 2>&1 | Out-String
        $out | Should -Not -Match 'All jobs'
        Get-JobNames $w | Should -Be @('old-complete')
        Join-Path $w 'overview.html' | Should -Not -Exist
    }

    It 'the console starts the removal in the background, and the demonstration removes nothing' {
        $start = Get-Content -LiteralPath (Join-Path $script:Root 'Start-MailboxTransferConsole.ps1') -Raw
        $start | Should -Match 'New-MbtGuiRealEngine[^\r\n]*\r?\n[^\r\n]*\r?\n\s*try \{ & \$engine\.Retention \$WorkFolder \}'
        $engineSource = Get-Content -LiteralPath (Join-Path $script:Root 'module\MailboxTransfer.Gui.Engine.ps1') -Raw
        $engineSource | Should -Match 'Invoke-MbtJobRetention -WorkFolder \$work'
        $fake = Get-Content -LiteralPath (Join-Path $script:Root 'module\MailboxTransfer.Gui.Fake.ps1') -Raw
        $fake | Should -Match '\$engine\.Retention = \{ param\(\$Work\) \}'
    }
}
