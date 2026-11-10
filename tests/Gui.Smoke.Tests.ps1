#Requires -Version 7.4
# Console: headless smoke test. The real WPF window is built (not shown) in a single-threaded pwsh of its own, on the fake engine, and walked
# through the state machine with a fake events file. No sign-in, no cloud.
BeforeAll {
    $package = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    $script = Join-Path $PSScriptRoot 'support\gui-smoke.ps1'
    $work = Join-Path ([IO.Path]::GetTempPath()) ('mbt-gui-smoke-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $out = & (Get-Process -Id $PID).Path -NoProfile -STA -File $script -Package $package -Work $work 2>$null
    $script:Raw = ($out | Out-String).Trim()
    $script:Result = $null
    try { $script:Result = $script:Raw | ConvertFrom-Json } catch { }
    Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Console smoke test (headless WPF)' {
    It 'ran and reported' {
        $script:Result | Should -Not -BeNullOrEmpty -Because $script:Raw
        @($script:Result.Checks).Count | Should -BeGreaterThan 100
    }
    It 'raised no exception' {
        @($script:Result.Errors) | Should -BeNullOrEmpty
    }
    It 'passes every check (pages of every scenario, light and dark, the state machine, the run states)' {
        $failed = @($script:Result.Checks | Where-Object { -not $_.Ok } | ForEach-Object { "$($_.Name): $($_.Detail)" })
        $failed | Should -BeNullOrEmpty
    }
    It 'walked every state of a run: running, slept, decision, stopped, complete, error' {
        $names = @($script:Result.Checks | ForEach-Object Name)
        foreach ($n in 'run: PC slept then resumed', 'run: needs a decision', 'run: stopped result stays on the run page with Continue', 'run: a complete result opens the done page', 'run: an error shows the plain message and offers Try again') {
            $names | Should -Contain $n
        }
    }
}
