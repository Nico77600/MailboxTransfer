#Requires -Version 7.6
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - every number of the tool is written in en-US (221.6 KiB, 13.8 s, 3,826,688), whatever the language of the session. The tests put the
    thread in fr-FR (decimal comma, narrow no-break space as thousands separator) and look at the helpers, the summaries of the engine, the tiles of the console,
    and at the end card and the log of a run of the entry script. Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\Culture.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\PstDamage.ps1')
    foreach ($f in 'Core', 'Run') { . (Join-Path $script:Root "module\MailboxTransfer.Gui.$f.ps1") }
    $script:Good = New-TestPst $TestDrive 'good'
    $script:SavedCulture = [Threading.Thread]::CurrentThread.CurrentCulture
    $script:Fr = [Globalization.CultureInfo]::GetCultureInfo('fr-FR')
    # a decimal comma in a number that is not a thousands group: 222,6 / 13,8 / 3,9 (a thousands group has three digits: 3,826,688)
    $script:DecimalComma = '\d,\d{1,2}(?![\d,])'
    function Enter-French { [Threading.Thread]::CurrentThread.CurrentCulture = $script:Fr }
    function Exit-French { [Threading.Thread]::CurrentThread.CurrentCulture = $script:SavedCulture }
}
AfterAll { [Threading.Thread]::CurrentThread.CurrentCulture = $script:SavedCulture }

Describe 'The session is really in French for these tests' {
    It 'writes a decimal comma' {
        Enter-French
        try { (1.5).ToString() | Should -Be '1,5' } finally { Exit-French }
    }
}

Describe 'Helpers of the console output and of the log' {
    BeforeEach { Enter-French }
    AfterEach { Exit-French }
    It 'Format-MbtBytes, Format-MbtDuration, Format-MbtRate and Format-MbtNumber use a decimal point and commas for thousands' {
        Format-MbtBytes 227942.4 | Should -Be '222.6 KiB'
        Format-MbtBytes (3.9MB) | Should -Be '3.9 MiB'
        Format-MbtBytes 512 | Should -Be '512 B'
        Format-MbtDuration 13.84 | Should -Be '13.8 s'
        Format-MbtDuration 0.25 | Should -Be '250 ms'
        Format-MbtDuration 3700 | Should -Be '01:01:40'
        Format-MbtRate 4000000 1.2 | Should -Be '3.18 MiB/s'
        Format-MbtNumber 3826688 | Should -Be '3,826,688'
    }
    It 'the numbers of the plan and of the tiles of the console (bytes, counts, spans, time left) are en-US too' {
        Format-MbtGuiBytes 2.5GB | Should -Be '2.5 GiB'
        Format-MbtGuiBytes 47.3GB | Should -Be '47.3 GiB'
        Format-MbtGuiBytes (3.9MB) | Should -Be '4 MiB'
        Format-MbtGuiCount 3826688 | Should -Be '3,826,688'
        Format-MbtGuiSpan 100 | Should -Be '1.7 h'
        (Get-MbtGuiPlanEta -Task import -Bytes 47.3GB).Text | Should -Not -Match $script:DecimalComma
        $r = New-MbtGuiRun
        $t0 = [datetime]::UtcNow.AddHours(-2)
        $null = Add-MbtGuiEvent -Run $r -Event ([pscustomobject]@{ v = 1; t = $t0.ToString('o'); e = 'step'; id = 's'; state = 'running'; done = 1; total = 100; bytes = 1GB; bytesTotal = 50GB })
        foreach ($m in 5, 10, 15, 20, 25, 30) {
            $null = Add-MbtGuiEvent -Run $r -Event ([pscustomobject]@{ v = 1; t = $t0.AddMinutes($m - 1).ToString('o'); e = 'step'; id = 's'; state = 'running'; done = $m; total = 100; bytes = 1GB; bytesTotal = 50GB })
            $null = Add-MbtGuiEvent -Run $r -Event ([pscustomobject]@{ v = 1; t = $t0.AddMinutes($m).ToString('o'); e = 'rate'; gibPerHour = 3.5 + $m / 30 })
        }
        $eta = Get-MbtGuiEta -Run $r -Now $t0.AddMinutes(31)
        $eta.State | Should -Be 'ready'
        $eta.Basis | Should -Match '\d+\.\d GiB per hour'
        $eta.Basis | Should -Not -Match $script:DecimalComma
    }
}

Describe 'Summaries written by the engine' {
    BeforeEach { Enter-French }
    AfterEach { Exit-French }
    It 'the check of a PST writes its sizes and its time in en-US' {
        $text = (Test-MbtPst -Path $script:Good).Report.Summary()
        $text | Should -Match 'result: 0 error\(s\), 0 warning\(s\), 0 information\(s\) in \d+\.\d s'
        $text | Should -Match '\(\d{1,3}(,\d{3})+ bytes\)'
        $text | Should -Not -Match $script:DecimalComma
        $text | Should -Not -Match '[  ]'
    }
    It 'the repair of a PST writes its counts and its time in en-US' {
        $res = Repair-MbtPst -Path $script:Good -Destination (Join-Path $TestDrive 'repaired.pst') -Level Recompute
        $text = $res.Report.Summary()
        $text | Should -Match 'action\(s\) in \d+\.\d s'
        $text | Should -Not -Match $script:DecimalComma
        $text | Should -Not -Match '[  ]'
    }
    It 'the N0 and N1 formats of the engine are invariant' {
        $files = Get-ChildItem -LiteralPath (Join-Path $script:Root 'src') -Filter 'Engine.*.cs'
        foreach ($f in $files) {
            $lines = Get-Content -LiteralPath $f.FullName | Where-Object { $_ -match 'ToString\("N[0-9]"\)' }
            @($lines).Count | Should -Be 0 -Because "$($f.Name) formats a number with the culture of the thread"
        }
    }
}

Describe 'A run of the entry script started from a French session' {
    It 'writes the end card and the log with decimal points' {
        $logs = Join-Path $TestDrive 'logs-fr'
        $french = "`$fr = [Globalization.CultureInfo]::GetCultureInfo('fr-FR'); [Globalization.CultureInfo]::DefaultThreadCurrentCulture = `$fr; [Globalization.CultureInfo]::DefaultThreadCurrentUICulture = `$fr; [Threading.Thread]::CurrentThread.CurrentCulture = `$fr; [Threading.Thread]::CurrentThread.CurrentUICulture = `$fr; & '$($script:Entry)' -DevelopmentMode -Mode CheckPst -Path '$($script:Good)' -LogFolder '$logs'"
        # the warning about the unsigned source tree is written by PowerShell before the script runs, in the language of the session: not a number of the tool
        $out = (& pwsh -NoProfile -NonInteractive -Command $french 2>&1 | Out-String) -split "`r?`n" | Where-Object { $_ -notmatch 'DevelopmentMode: no package manifest' } | Join-String -Separator "`n"
        $LASTEXITCODE | Should -Be 0
        $out | Should -Match 'CheckPst finished'
        $out | Should -Match 'Duration\s+\d+\.\d s'
        $out | Should -Not -Match $script:DecimalComma
        $out | Should -Not -Match '[  ]'
        $log = (Get-ChildItem -LiteralPath $logs -Filter '*.log' | ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw }) -join "`n"
        $log | Should -Match 'Summary - Duration: \d+\.\d s'
        $log | Should -Not -Match $script:DecimalComma
        $log | Should -Not -Match '[  ]'
    }
    It 'the entry scripts set the culture of the process and of the threads that start later' {
        foreach ($name in 'Invoke-MailboxTransfer.ps1', 'Start-MailboxTransferConsole.ps1') {
            $text = Get-Content -LiteralPath (Join-Path $script:Root $name) -Raw
            $text | Should -Match 'DefaultThreadCurrentCulture = \$mbtCulture' -Because $name
            $text | Should -Match 'DefaultThreadCurrentUICulture = \$mbtCulture' -Because $name
            $text | Should -Match 'CurrentThread\.CurrentCulture = \$mbtCulture' -Because $name
        }
    }
}
