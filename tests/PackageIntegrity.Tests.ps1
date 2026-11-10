#Requires -Version 7.4
<#
    Package integrity: the contract of the release package (signed PowerShell files, hash manifest, byte-exact Git attributes) and the
    logic of the verifier on a controlled copy of the package.

    The runtime fails closed on any missing, changed or untrusted file once a PackageIntegrity.psd1 is there. The fixture below is a COPY of
    package\ in which exactly one thing is substituted: every call of Microsoft.PowerShell.Security\Get-AuthenticodeSignature returns a
    Valid, timestamped, owner-thumbprint result. Everything else is the real code and the real bytes: the verifier, the manifest
    generator (tools\New-PackageIntegrity.ps1), the SHA256 comparison, the path and reparse-point guards and the scan of the tree.
    This proves the integrity LOGIC. It does not prove Authenticode trust: the real signatures are checked by the last Describe, on the
    unmodified package, when the owner has signed it.
#>
# Needed at discovery time (-Skip conditions).
$script:Package = Join-Path (Split-Path $PSScriptRoot -Parent) package
BeforeAll {
    $script:Repo = Split-Path $PSScriptRoot -Parent
    $script:Package = Join-Path $script:Repo 'package'
    $script:Signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
    $script:Token = 'Microsoft.PowerShell.Security\Get-AuthenticodeSignature'
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('mbt-integrity-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp
    $script:Substitute = '& { [CmdletBinding()] param([string]$LiteralPath) [pscustomobject]@{ Status = ''Valid''; SignerCertificate = [pscustomobject]@{ Thumbprint = ''' + $script:Signer + ''' }; TimeStamperCertificate = [pscustomobject]@{ Subject = ''TEST-FIXTURE-SUBSTITUTE'' } } }'

    function Get-SignedPaths {
        @(Get-ChildItem -LiteralPath $script:Package -Recurse -File -Include *.ps1, *.psm1, *.psd1 |
            Where-Object { $_.FullName -notlike (Join-Path $script:Package 'config\*') -and $_.Name -ne 'PackageIntegrity.psd1' } |
            ForEach-Object { [IO.Path]::GetRelativePath($script:Package, $_.FullName) } | Sort-Object)
    }
    function New-IntegrityFixture {
        param([Parameter(Mandatory)][string]$Name, [switch]$WithoutManifest)
        $root = Join-Path $script:Tmp $Name
        foreach ($file in Get-ChildItem -LiteralPath $script:Package -Recurse -File -Force) {
            $relative = [IO.Path]::GetRelativePath($script:Package, $file.FullName)
            if ($relative -like 'docs\images\*' -or $relative -like 'config\*.local.psd1' -or $relative -eq 'PackageIntegrity.psd1') { continue }
            $target = Join-Path $root $relative
            $null = New-Item -ItemType Directory -Force -Path (Split-Path $target -Parent)
            Copy-Item -LiteralPath $file.FullName -Destination $target
        }
        foreach ($name2 in 'CHANGELOG.md', 'SECURITY.md') { Copy-Item -LiteralPath (Join-Path $script:Repo $name2) -Destination (Join-Path $root $name2) }
        foreach ($relative in 'Invoke-MailboxTransfer.ps1', 'Start-MailboxTransferConsole.ps1', 'MailboxTransfer.psm1', 'src\Test-PackageIntegrity.ps1') {
            $path = Join-Path $root $relative
            [IO.File]::WriteAllText($path, [IO.File]::ReadAllText($path).Replace($script:Token, $script:Substitute), [Text.UTF8Encoding]::new($false))
        }
        if (-not $WithoutManifest) { $null = & (Join-Path $script:Repo 'tools\New-PackageIntegrity.ps1') -Root $root 6>&1 }
        return $root
    }
    function Invoke-Verifier {
        param([Parameter(Mandatory)][string]$Root, [switch]$Distribution, [switch]$DevelopmentMode)
        $arguments = @('-NoProfile', '-NonInteractive', '-File', (Join-Path $Root 'src\Test-PackageIntegrity.ps1'), '-Root', $Root)
        if ($Distribution) { $arguments += '-Distribution' }
        if ($DevelopmentMode) { $arguments += '-DevelopmentMode' }
        $out = & pwsh @arguments 2>&1 | Out-String
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Text = $out }
    }
    function Set-FixtureText([string]$Path, [string]$Text) { [IO.File]::WriteAllText($Path, $Text, [Text.UTF8Encoding]::new($false)) }
}
AfterAll {
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Static package contract' {
    It 'lists in the verifier exactly the PowerShell files of the package, and the signing tool finds the same ones' {
        $text = [IO.File]::ReadAllText((Join-Path $script:Package 'src\Test-PackageIntegrity.ps1'))
        $block = [regex]::Match($text, '(?s)\$signed = @\((.*?)\r?\n\)').Groups[1].Value
        $listed = @([regex]::Matches($block, "'([^']+)'") | ForEach-Object { $_.Groups[1].Value } | Sort-Object)
        $actual = Get-SignedPaths
        $listed | Should -Be $actual
        $actual.Count | Should -Be 20
    }
    It 'keeps the signed files ASCII (the signature must read the same in Windows PowerShell 5.1 and PowerShell 7)' {
        foreach ($relative in Get-SignedPaths) {
            $bytes = [IO.File]::ReadAllBytes((Join-Path $script:Package $relative))
            ($bytes | Where-Object { $_ -gt 127 }).Count | Should -Be 0 -Because $relative
        }
    }
    It 'pins the fixed owner thumbprint in the verifier, the entry scripts, the module and the package builder' {
        foreach ($relative in 'src\Test-PackageIntegrity.ps1', 'Invoke-MailboxTransfer.ps1', 'Start-MailboxTransferConsole.ps1', 'MailboxTransfer.psm1') {
            [IO.File]::ReadAllText((Join-Path $script:Package $relative)) | Should -Match $script:Signer -Because $relative
        }
        [IO.File]::ReadAllText((Join-Path $script:Repo 'tools\New-MailboxTransferPackage.ps1')) | Should -Match $script:Signer
        [IO.File]::ReadAllText((Join-Path $script:Repo 'tools\New-PackageIntegrity.ps1')) | Should -Match $script:Signer
    }
    It 'offers no unsigned mode of the verifier: a missing manifest or signature is an error, and -DevelopmentMode does not skip the checks' {
        $text = [IO.File]::ReadAllText((Join-Path $script:Package 'src\Test-PackageIntegrity.ps1'))
        $text | Should -Match 'PackageIntegrity.psd1 is missing'
        $text | Should -Not -Match '(?i)\$env:MBT_|\$env:MAILBOXTRANSFER_(SKIP|UNSIGNED|BYPASS)'
        ([regex]::Matches($text, 'Assert-OwnerSignature \$')).Count | Should -BeGreaterOrEqual 3
    }
    It 'keeps every signed and hashed path byte-exact through .gitattributes (-text, cr-at-eol on the signed files)' {
        $paths = @(Get-SignedPaths | ForEach-Object { 'package/' + ($_ -replace '\\', '/') }) + 'package/PackageIntegrity.psd1', 'package/README.md', 'package/docs/MailboxTransfer-UserGuide.html', 'package/lib/sqlite/Microsoft.Data.Sqlite.dll', 'CHANGELOG.md', 'SECURITY.md'
        $attr = @(git -C $script:Repo check-attr text whitespace -- @paths)
        foreach ($path in $paths) {
            $attr | Should -Contain "${path}: text: unset"
            $attr | Should -Contain "${path}: whitespace: blank-at-eol,space-before-tab,blank-at-eof,cr-at-eol"
        }
    }
    It 'has one version: the module manifest, the module, the entry scripts, product.json and the top of the changelog' {
        $v = '1.0.0'
        (Import-PowerShellDataFile -LiteralPath (Join-Path $script:Package 'MailboxTransfer.psd1')).ModuleVersion | Should -Be $v
        [IO.File]::ReadAllText((Join-Path $script:Package 'MailboxTransfer.psm1')) | Should -Match "(?m)^\`$script:ToolVersion = '$v'"
        [IO.File]::ReadAllText((Join-Path $script:Package 'Start-MailboxTransferConsole.ps1')) | Should -Match "(?m)^\`$script:ConsoleVersion = '$v'"
        (Get-Content -LiteralPath (Join-Path $script:Package 'product.json') -Raw | ConvertFrom-Json).version | Should -Be $v
        [IO.File]::ReadAllText((Join-Path $script:Repo 'CHANGELOG.md')) | Should -Match "(?m)^## $([regex]::Escape($v)) "
    }
    It 'distributes no data, no private configuration and no test or tool (builder dry run)' {
        $destination = Join-Path $script:Tmp 'dry-run'
        $null = & (Join-Path $script:Repo 'tools\New-MailboxTransferPackage.ps1') -Destination $destination -DevelopmentMode 6>&1
        $names = @(Get-ChildItem -LiteralPath $destination -Recurse -File | ForEach-Object { [IO.Path]::GetRelativePath($destination, $_.FullName) })
        $names | Should -Contain 'CHANGELOG.md'
        $names | Should -Contain 'SECURITY.md'
        $names | Should -Contain 'docs\MailboxTransfer-UserGuide.html'
        $names | Should -Not -Contain 'config\MailboxTransfer.config.local.psd1'
        @($names | Where-Object { $_ -match '^(tests|tools|out)\\' -or $_ -match '\.(fxs|sqlite|pst|eml)$' -or $_ -like 'docs\images\*' }).Count | Should -Be 0
    }
    It 'refuses a source tree without manifest unless -DevelopmentMode (entry script, in a fresh process)' -Skip:(Test-Path -LiteralPath (Join-Path $script:Package 'PackageIntegrity.psd1')) {
        $out = & pwsh -NoProfile -NonInteractive -File (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1') -Mode Status -StorePath (Join-Path $script:Tmp 'none') 2>&1 | Out-String
        $out | Should -Match 'PackageIntegrity.psd1 is missing'
        $out | Should -Match 'DevelopmentMode'
    }
}

Describe 'Controlled fixture - real verifier, hash and path logic (signature call substituted)' {
    BeforeAll { $script:Fx = New-IntegrityFixture -Name 'ok' }
    It 'accepts the fixture package, in a distribution and in development mode' {
        $r = Invoke-Verifier -Root $script:Fx -Distribution
        $r.ExitCode | Should -Be 0 -Because $r.Text
        $r.Text | Should -Match 'MailboxTransfer'
        (Invoke-Verifier -Root $script:Fx -DevelopmentMode).ExitCode | Should -Be 0
    }
    It 'accepts the package without the files that only a distribution carries (CHANGELOG.md, SECURITY.md)' {
        $fx = New-IntegrityFixture -Name 'source'
        Remove-Item -LiteralPath (Join-Path $fx 'CHANGELOG.md'), (Join-Path $fx 'SECURITY.md')
        (Invoke-Verifier -Root $fx).ExitCode | Should -Be 0
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
    }
    It 'refuses -Distribution together with -DevelopmentMode' {
        $r = Invoke-Verifier -Root $script:Fx -Distribution -DevelopmentMode
        $r.ExitCode | Should -Not -Be 0
        $r.Text | Should -Match 'cannot be combined'
    }
    It 'rejects a changed stable file: <Path>' -ForEach @(
        @{ Path = 'src\Engine.Store.cs' }, @{ Path = 'module\MailboxTransfer.Gui.xaml' }, @{ Path = 'lib\sqlite\Microsoft.Data.Sqlite.dll' },
        @{ Path = 'MailboxTransfer.psm1' }, @{ Path = 'README.md' }, @{ Path = 'docs\MailboxTransfer-UserGuide.html' }, @{ Path = 'CHANGELOG.md' }) {
        $fx = New-IntegrityFixture -Name ('changed-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        $file = Join-Path $fx $Path
        [IO.File]::AppendAllText($file, ' ')
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        $r.Text | Should -Match 'SHA256 mismatch'
    }
    It 'rejects a missing stable file: <Path>' -ForEach @(@{ Path = 'src\Engine.Auth.cs' }, @{ Path = 'module\MailboxTransfer.Events.ps1' }, @{ Path = 'LICENSE' }, @{ Path = 'Start-MailboxTransferConsole.cmd' }) {
        $fx = New-IntegrityFixture -Name ('missing-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        Remove-Item -LiteralPath (Join-Path $fx $Path)
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
    }
    It 'rejects an unexpected file: <Path>' -ForEach @(@{ Path = 'module\Extra.ps1' }, @{ Path = 'src\Engine.Extra.cs' }, @{ Path = 'lib\sqlite\extra.dll' }, @{ Path = 'extra.txt' }) {
        $fx = New-IntegrityFixture -Name ('extra-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        Set-FixtureText (Join-Path $fx $Path) 'x'
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        $r.Text | Should -Match 'unexpected stable file'
    }
    It 'accepts a distribution without the pictures and the Markdown guides, which the manifest lists but the builder does not ship' {
        $fx = New-IntegrityFixture -Name ('docs-' + [guid]::NewGuid().ToString('N').Substring(0, 6)) -WithoutManifest
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $fx 'docs\images')
        [IO.File]::WriteAllBytes((Join-Path $fx 'docs\images\picture.png'), [byte[]](137, 80, 78, 71, 13, 10, 26, 10, 1))
        $null = & (Join-Path $script:Repo 'tools\New-PackageIntegrity.ps1') -Root $fx 6>&1
        (Get-Content -LiteralPath (Join-Path $fx 'PackageIntegrity.psd1') -Raw) | Should -Match 'docs\\images\\picture\.png'
        (Invoke-Verifier -Root $fx -Distribution).ExitCode | Should -Be 0
        Remove-Item -LiteralPath (Join-Path $fx 'docs\images') -Recurse -Force
        foreach ($guide in 'docs\MailboxTransfer-Guide.md', 'docs\MailboxTransfer-UserGuide.md') { Remove-Item -LiteralPath (Join-Path $fx $guide) }
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Be 0 -Because $r.Text
        (Invoke-Verifier -Root $fx).ExitCode | Should -Be 0
    }
    It 'still refuses a picture or a Markdown guide that is present with another content, or that is not listed' -ForEach @(
        @{ Path = 'docs\images\picture.png'; Listed = $true }, @{ Path = 'docs\MailboxTransfer-UserGuide.md'; Listed = $true },
        @{ Path = 'docs\images\unlisted.png'; Listed = $false }, @{ Path = 'docs\Unlisted.md'; Listed = $false }) {
        $fx = New-IntegrityFixture -Name ('docs-' + [guid]::NewGuid().ToString('N').Substring(0, 6)) -WithoutManifest
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $fx 'docs\images')
        [IO.File]::WriteAllBytes((Join-Path $fx 'docs\images\picture.png'), [byte[]](137, 80, 78, 71, 13, 10, 26, 10, 1))
        $null = & (Join-Path $script:Repo 'tools\New-PackageIntegrity.ps1') -Root $fx 6>&1
        if ($Listed) { [IO.File]::AppendAllText((Join-Path $fx $Path), ' ') } else { Set-FixtureText (Join-Path $fx $Path) 'x' }
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        $r.Text | Should -Match $(if ($Listed) { 'SHA256 mismatch' } else { 'unexpected stable file' })
    }
    It 'still requires every executable and code path of the manifest, the pictures excepted: <Path>' -ForEach @(
        @{ Path = 'product.json' }, @{ Path = 'module\MailboxTransfer.Gui.xaml' }, @{ Path = 'src\Gui.ViewModels.cs' }, @{ Path = 'Start-MailboxTransferConsole.cmd' },
        @{ Path = 'lib\sqlite\SQLitePCLRaw.core.dll' }, @{ Path = 'module\MailboxTransfer.Gui.State.ps1' }, @{ Path = 'docs\MailboxTransfer-UserGuide.html' }, @{ Path = 'README.md' }) {
        $fx = New-IntegrityFixture -Name ('required-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        Remove-Item -LiteralPath (Join-Path $fx $Path)
        (Invoke-Verifier -Root $fx -Distribution).ExitCode | Should -Not -Be 0
        (Invoke-Verifier -Root $fx).ExitCode | Should -Not -Be 0
    }
    It 'permits the editable configuration, its local copy and the generated folders (data, outside the integrity)' {
        $fx = New-IntegrityFixture -Name 'editable'
        Set-FixtureText (Join-Path $fx 'config\MailboxTransfer.config.psd1') '@{ Output = @{ Root = ''D:\x'' } }'
        Set-FixtureText (Join-Path $fx 'config\MailboxTransfer.config.local.psd1') '@{ }'
        foreach ($folder in 'logs', 'out', 'data', 'bin') { $null = New-Item -ItemType Directory -Path (Join-Path $fx $folder); Set-FixtureText (Join-Path $fx "$folder\x.txt") 'x' }
        (Invoke-Verifier -Root $fx -Distribution).ExitCode | Should -Be 0
    }
    It 'rejects a manifest that is <Name>' -ForEach @(
        @{ Name = 'missing'; Edit = { param($t) $null } },
        @{ Name = 'with a duplicate path'; Edit = { param($t) $m = [regex]::Match($t, "(?m)^\s+@\{ Path = 'LICENSE'.*$"); $t.Insert($m.Index, $m.Value + "`r`n") } },
        @{ Name = 'with a bad hash'; Edit = { param($t) [regex]::Replace($t, "(Path = 'LICENSE'; Sha256 = ')[0-9A-F]", '${1}Z') } },
        @{ Name = 'without a signed script'; Edit = { param($t) [regex]::Replace($t, "(?m)^\s+@\{ Path = 'module\\MailboxTransfer.Eml.ps1'.*\r?\n", '') } },
        @{ Name = 'for another signer'; Edit = { param($t) $t.Replace('0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6', 'FFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFFF') } },
        @{ Name = 'with a path that leaves the package'; Edit = { param($t) $t.Replace("Path = 'LICENSE'", "Path = '..\LICENSE'") } }) {
        $fx = New-IntegrityFixture -Name ('manifest-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        $manifest = Join-Path $fx 'PackageIntegrity.psd1'
        $new = & $Edit ([IO.File]::ReadAllText($manifest))
        if ($null -eq $new) { Remove-Item -LiteralPath $manifest } else { Set-FixtureText $manifest $new }
        (Invoke-Verifier -Root $fx -Distribution).ExitCode | Should -Not -Be 0
    }
    It 'rejects a reparse point in the stable tree: <Path>' -ForEach @(@{ Path = 'module' }, @{ Path = 'docs' }, @{ Path = 'lib' }) {
        $fx = New-IntegrityFixture -Name ('junction-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        $real = Join-Path $script:Tmp ('real-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        Move-Item -LiteralPath (Join-Path $fx $Path) -Destination $real
        $null = New-Item -ItemType Junction -Path (Join-Path $fx $Path) -Target $real
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        $r.Text | Should -Match 'reparse|unsafe path'
    }
    It 'fails the entry script and the module before any engine code when a stable file changed' {
        $fx = New-IntegrityFixture -Name 'entry-changed'
        [IO.File]::AppendAllText((Join-Path $fx 'src\Engine.Store.cs'), ' ')
        $out = & pwsh -NoProfile -NonInteractive -File (Join-Path $fx 'Invoke-MailboxTransfer.ps1') -Mode Status -StorePath (Join-Path $script:Tmp 'none') 2>&1 | Out-String
        $out | Should -Match 'SHA256 mismatch'
        $out = & pwsh -NoProfile -NonInteractive -Command "Import-Module '$(Join-Path $fx 'MailboxTransfer.psd1')'" 2>&1 | Out-String
        $out | Should -Match 'SHA256 mismatch'
    }
    It 'starts from the intact fixture package and reports the store state' {
        $out = & pwsh -NoProfile -NonInteractive -File (Join-Path $script:Fx 'Invoke-MailboxTransfer.ps1') -Mode Status -StorePath (Join-Path $script:Tmp 'none') 2>&1 | Out-String
        $out | Should -Not -Match 'Package integrity'
    }
    It 'refuses a tree without manifest in the module only as a warning, and in the entry script as an error' {
        $fx = New-IntegrityFixture -Name 'no-manifest' -WithoutManifest
        $out = & pwsh -NoProfile -NonInteractive -File (Join-Path $fx 'Invoke-MailboxTransfer.ps1') -Mode Status -StorePath (Join-Path $script:Tmp 'none') 2>&1 | Out-String
        $out | Should -Match 'PackageIntegrity.psd1 is missing'
        $out = & pwsh -NoProfile -NonInteractive -Command "Import-Module '$(Join-Path $fx 'MailboxTransfer.psd1')' -WarningVariable w; `$w -join ' '" 2>&1 | Out-String
        $out | Should -Match 'unverified source tree'
    }
}

Describe 'Package builder - controlled fixture' {
    It 'refuses a source that is not authenticated, before writing anything' -Skip:(Test-Path -LiteralPath (Join-Path $script:Package 'PackageIntegrity.psd1')) {
        $destination = Join-Path $script:Tmp 'refused'
        { & (Join-Path $script:Repo 'tools\New-MailboxTransferPackage.ps1') -Destination $destination 6>&1 | Out-Null } | Should -Throw
        Test-Path -LiteralPath $destination | Should -BeFalse
    }
}

Describe 'Native state of the real package - no substitution (honest, state-dependent)' {
    It 'is a source tree (no manifest yet) or a signed package that the real verifier accepts' {
        $manifest = Join-Path $script:Package 'PackageIntegrity.psd1'
        if (-not (Test-Path -LiteralPath $manifest)) {
            Set-ItResult -Skipped -Because 'PackageIntegrity.psd1 does not exist yet: the files are not signed (release preparation)'
            return
        }
        $r = Invoke-Verifier -Root $script:Package -DevelopmentMode
        $r.ExitCode | Should -Be 0 -Because $r.Text
    }
}
