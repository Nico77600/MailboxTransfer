#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
# Console: the REAL adapters of Gui.Engine.ps1 (sign-in, locations, folders, messages, preview, selection file, jobs, PST, store and .eml sources, the check of a PST)
# on a worker that runs inline, over stubs of the engine calls that reach the cloud, with 0, 1 and several results shaped like the real ones. The fake engine of the
# other console tests never calls these adapters and always answers with several items; the worker of the console hands every result as an array. Each check is one
# test: what the adapter hands to its page (one object or an array, the properties the page reads under StrictMode Latest). Nothing here reaches the network.
BeforeDiscovery {
    $package = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    $script = Join-Path $PSScriptRoot 'support\gui-engine.ps1'
    $work = Join-Path ([IO.Path]::GetTempPath()) ('mbt-gui-engine-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    $stdout = Join-Path ([IO.Path]::GetTempPath()) ('mbt-gui-engine-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.json')
    $stderr = $stdout + '.err'
    $process = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList @('-NoProfile', '-STA', '-File', $script, '-Package', $package, '-Work', (Join-Path $work 'w')) -RedirectStandardOutput $stdout -RedirectStandardError $stderr -WindowStyle Hidden -PassThru
    $null = $process.WaitForExit(900000)
    $raw = if (Test-Path -LiteralPath $stdout) { ([string](Get-Content -LiteralPath $stdout -Raw)).Trim() } else { '' }
    $err = if (Test-Path -LiteralPath $stderr) { [string](Get-Content -LiteralPath $stderr -Raw) } else { '' }
    $result = $null
    try { $result = $raw | ConvertFrom-Json } catch { }
    $checks = @()
    if ($result) { $checks = @($result.Checks | ForEach-Object { @{ Name = [string]$_.Name; Ok = [bool]$_.Ok; Detail = [string]$_.Detail } }); $checks += @($result.Errors | ForEach-Object { @{ Name = 'raised an exception'; Ok = $false; Detail = [string]$_ } }) }
    $checks += @{ Name = 'the adapters ran and reported (more than 60 checks)'; Ok = ($checks.Count -gt 60); Detail = "$($checks.Count) checks. $raw $err" }
    $script:EngineChecks = $checks
    Remove-Item -LiteralPath $stdout, $stderr, $work -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Console adapters of the real engine (inline worker, stubs of the cloud calls)' {
    It '<Name>' -ForEach $script:EngineChecks {
        $Ok | Should -BeTrue -Because $Detail
    }
}
