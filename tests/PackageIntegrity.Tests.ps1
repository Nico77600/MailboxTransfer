#Requires -Version 7.6
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
        $actual.Count | Should -Be 24
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
        $v = '1.3.0'
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

Describe 'Reparse points are judged by their tag' {
    # A symbolic link needs a privilege that an ordinary account may not hold; known at discovery, for -Skip.
    $canSymlink = $false
    $probeTarget = Join-Path ([IO.Path]::GetTempPath()) ('mbt-symlink-probe-' + [guid]::NewGuid().ToString('N'))
    try {
        [void](New-Item -ItemType SymbolicLink -Path $probeTarget -Target $env:SystemRoot -ErrorAction Stop)
        $canSymlink = $true
        (Get-Item -LiteralPath $probeTarget -Force).Delete()
    } catch { $canSymlink = $false }
    BeforeAll {
        $script:HelperNames = 'Get-PathReparseTag', 'Test-ReparseTagAllowed', 'Get-RefusedReparseTag', 'Test-CloudReadErrorCode', 'Get-ReadFailureCode', 'Assert-PlainPath'
        $verifierAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Package 'src\Test-PackageIntegrity.ps1'), [ref]$null, [ref]$null)
        foreach ($name in $script:HelperNames) {
            $function = $verifierAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name }, $true)
            if ($null -eq $function) { throw "The verifier function $name is missing." }
            . ([scriptblock]::Create($function.Extent.Text))
        }
        $script:TagFx = New-IntegrityFixture -Name 'tag-ok'
        # A fixture whose verifier text is changed at an anchor; the manifest is written again so that the changed verifier is the signed one.
        function Set-FixtureVerifier([string]$Fixture, [string]$Search, [string]$Replace) {
            $verifier = Join-Path $Fixture 'src\Test-PackageIntegrity.ps1'
            $text = [IO.File]::ReadAllText($verifier)
            if (-not $text.Contains($Search)) { throw "The anchor was not found in the verifier: $Search" }
            [IO.File]::WriteAllText($verifier, $text.Replace($Search, $Replace), [Text.UTF8Encoding]::new($false))
            $null = & (Join-Path $script:Repo 'tools\New-PackageIntegrity.ps1') -Root $Fixture 6>&1
        }
        function Set-FixtureTag([string]$Fixture, [uint32]$Tag) {
            $literal = '0x' + $Tag.ToString('X8') + 'u'
            Set-FixtureVerifier $Fixture "Assert-PlainPath `$rootPath `$true`r`n" ('function Get-PathReparseTag { param($Path) [uint32]' + $literal + " }`r`nAssert-PlainPath `$rootPath `$true`r`n")
        }
        # docs\ becomes a junction to a copy of itself: a directory that is a reparse point and whose files keep their bytes.
        function New-DocsJunction([string]$Fixture) {
            $link = Join-Path $Fixture 'docs'
            $copy = Join-Path $script:Tmp ('docs-copy-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            Move-Item -LiteralPath $link -Destination $copy
            [void](New-Item -ItemType Junction -Path $link -Target $copy)
            [pscustomobject]@{ Link = $link; Target = $copy }
        }
        function ConvertTo-FlatText([string]$Text) { ($Text -replace '\s*\|\s*', ' ' -replace '\s+', ' ') }
    }

    It 'judges the tag <Name> (0x<Tag>): accepted is <Accepted>' -ForEach @(
        @{ Name = 'symbolic link'; Tag = 'A000000C'; Accepted = $false }
        @{ Name = 'junction and mount point'; Tag = 'A0000003'; Accepted = $false }
        @{ Name = 'OneDrive cloud files'; Tag = '9000001A'; Accepted = $true }
        @{ Name = 'OneDrive cloud files, variant 1'; Tag = '9000101A'; Accepted = $true }
        @{ Name = 'OneDrive cloud files, variant 2'; Tag = '9000201A'; Accepted = $true }
        @{ Name = 'OneDrive cloud files, variant 7'; Tag = '9000701A'; Accepted = $true }
        @{ Name = 'OneDrive cloud files, variant F'; Tag = '9000F01A'; Accepted = $true }
        @{ Name = 'data deduplication'; Tag = '80000013'; Accepted = $true }
        @{ Name = 'WOF (CompactOS)'; Tag = '80000017'; Accepted = $true }
        @{ Name = 'a cloud tag that carries the name-surrogate bit'; Tag = 'B000001A'; Accepted = $false }
        @{ Name = 'an unknown tag with the name-surrogate bit (WSL symbolic link)'; Tag = 'A000001D'; Accepted = $false }
        @{ Name = 'the name-surrogate bit alone'; Tag = '20000000'; Accepted = $false }
        @{ Name = 'all bits set'; Tag = 'FFFFFFFF'; Accepted = $false }
        @{ Name = 'an unknown tag without the bit (app execution alias)'; Tag = '8000001B'; Accepted = $false }
        @{ Name = 'an unknown tag without the bit (hierarchical storage)'; Tag = 'C0000004'; Accepted = $false }
        @{ Name = 'an unknown Microsoft tag'; Tag = '80000023'; Accepted = $false }
        @{ Name = 'a cloud tag with another low byte'; Tag = '9000001B'; Accepted = $false }
        @{ Name = 'a cloud tag with a variant in the wrong nibble'; Tag = '9000011A'; Accepted = $false }
        @{ Name = 'a cloud tag with another high word'; Tag = '9001001A'; Accepted = $false }
        @{ Name = 'a deduplication tag with a variant'; Tag = '80001013'; Accepted = $false }
        @{ Name = 'a tag that is not set'; Tag = '00000000'; Accepted = $false }
        @{ Name = 'a non-Microsoft tag'; Tag = '00000011'; Accepted = $false }
    ) {
        Test-ReparseTagAllowed ([Convert]::ToUInt32($Tag, 16)) | Should -Be $Accepted
    }
    It 'reads no tag from a normal file or folder, and refuses to guess for a path that does not exist' {
        $file = Join-Path $script:TagFx 'README.md'
        Get-PathReparseTag $file | Should -BeNullOrEmpty
        Get-PathReparseTag $script:TagFx | Should -BeNullOrEmpty
        Get-RefusedReparseTag (Get-Item -LiteralPath $file -Force) | Should -BeNullOrEmpty
        { Get-PathReparseTag (Join-Path $script:TagFx 'missing.txt') } | Should -Throw -ExpectedMessage '*Windows error 2*'
    }
    It 'reads the tag of a junction without opening its target, from a path of more than 260 characters too' {
        $target = Join-Path $script:Tmp 'tag-target'
        [void][IO.Directory]::CreateDirectory($target)
        $link = Join-Path $script:Tmp 'tag-junction'
        [void](New-Item -ItemType Junction -Path $link -Target $target)
        try {
            '{0:X8}' -f (Get-PathReparseTag $link) | Should -Be 'A0000003'
            Get-RefusedReparseTag (Get-Item -LiteralPath $link -Force) | Should -Be 'A0000003'
            $deep = Join-Path $script:Tmp ((1..4 | ForEach-Object { 'd' * 70 }) -join '\')
            [void][IO.Directory]::CreateDirectory('\\?\' + $deep)
            $deepLink = Join-Path $deep 'deep-junction'
            [void](New-Item -ItemType Junction -Path ('\\?\' + $deepLink) -Target $target)
            $deepLink.Length | Should -BeGreaterThan 260
            '{0:X8}' -f (Get-PathReparseTag $deepLink) | Should -Be 'A0000003'
            [IO.Directory]::Delete('\\?\' + $deepLink)
        }
        finally { (Get-Item -LiteralPath $link -Force).Delete() }
    }
    It 'reads the tag of a symbolic link to a file and to a folder' -Skip:(-not $canSymlink) {
        $target = Join-Path $script:Tmp 'symlink-target'
        [void][IO.Directory]::CreateDirectory($target)
        $targetFile = Join-Path $target 'file.txt'
        Set-Content -LiteralPath $targetFile -Value 'x'
        $fileLink = Join-Path $script:Tmp 'file-link'
        $folderLink = Join-Path $script:Tmp 'folder-link'
        [void](New-Item -ItemType SymbolicLink -Path $fileLink -Target $targetFile)
        [void](New-Item -ItemType SymbolicLink -Path $folderLink -Target $target)
        try {
            '{0:X8}' -f (Get-PathReparseTag $fileLink) | Should -Be 'A000000C'
            '{0:X8}' -f (Get-PathReparseTag $folderLink) | Should -Be 'A000000C'
        }
        finally { (Get-Item -LiteralPath $fileLink -Force).Delete(); (Get-Item -LiteralPath $folderLink -Force).Delete() }
    }
    It 'keeps one text of the three tag functions in the verifier, the module, both entry scripts and both tools' {
        $files = 'package\src\Test-PackageIntegrity.ps1', 'package\MailboxTransfer.psm1', 'package\Invoke-MailboxTransfer.ps1', 'package\Start-MailboxTransferConsole.ps1', 'tools\New-MailboxTransferPackage.ps1', 'tools\New-PackageIntegrity.ps1'
        foreach ($name in 'Get-PathReparseTag', 'Test-ReparseTagAllowed', 'Get-RefusedReparseTag') {
            $texts = foreach ($relative in $files) {
                $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $script:Repo $relative), [ref]$null, [ref]$null)
                $found = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name }, $true))
                $found.Count | Should -Be 1 -Because "$name in $relative"
                $found[0].Extent.Text.Replace("`r`n", "`n")
            }
            @($texts | Select-Object -Unique).Count | Should -Be 1 -Because $name
        }
    }
    It 'judges a path by its tag and not by the bare ReparsePoint attribute, in every place that judges one' {
        foreach ($relative in 'package\src\Test-PackageIntegrity.ps1', 'package\MailboxTransfer.psm1', 'package\Invoke-MailboxTransfer.ps1', 'package\Start-MailboxTransferConsole.ps1', 'package\module\MailboxTransfer.Events.ps1', 'tools\New-MailboxTransferPackage.ps1', 'tools\New-PackageIntegrity.ps1') {
            $text = [IO.File]::ReadAllText((Join-Path $script:Repo $relative))
            @($text -split "`r?`n" | Where-Object { $_ -match '\$\w+\.Attributes -band \[IO\.FileAttributes\]::ReparsePoint' -and $_ -notmatch '\$(Item|Entry)\.Attributes' }).Count | Should -Be 0 -Because $relative
        }
    }

    It 'refuses a file, a folder and an ancestor whose tag is <Name>, and names the tag' -ForEach @(
        @{ Name = 'a symbolic link'; Tag = 'A000000C' }
        @{ Name = 'a junction'; Tag = 'A0000003' }
        @{ Name = 'unknown with the name-surrogate bit'; Tag = 'A000001D' }
        @{ Name = 'unknown without it'; Tag = '8000001B' }
        @{ Name = 'a cloud tag with the surrogate bit'; Tag = 'B000001A' }
    ) {
        $target = Join-Path $script:Tmp ('guard-refused-' + $Tag)
        [void][IO.Directory]::CreateDirectory($target)
        Set-Content -LiteralPath (Join-Path $target 'file.txt') -Value 'x'
        $link = Join-Path $script:Tmp ('guard-link-' + $Tag)
        [void](New-Item -ItemType Junction -Path $link -Target $target)
        Mock Get-PathReparseTag ([scriptblock]::Create("[uint32]0x${Tag}u"))
        try {
            { Assert-PlainPath $link $true } | Should -Throw -ExpectedMessage "*reparse point that is not accepted (tag 0x$Tag)*"
            { Assert-PlainPath (Join-Path $link 'file.txt') } | Should -Throw -ExpectedMessage "*reparse-point ancestor that is not accepted (tag 0x$Tag)*"
        }
        finally { (Get-Item -LiteralPath $link -Force).Delete() }
    }
    It 'accepts a file and a folder whose tag is <Name>' -ForEach @(
        @{ Name = 'a OneDrive cloud placeholder'; Tag = '9000001A' }
        @{ Name = 'a OneDrive cloud placeholder, variant F'; Tag = '9000F01A' }
        @{ Name = 'a deduplicated item'; Tag = '80000013' }
        @{ Name = 'a WOF item'; Tag = '80000017' }
    ) {
        $target = Join-Path $script:Tmp ('guard-accepted-' + $Tag)
        [void][IO.Directory]::CreateDirectory($target)
        Set-Content -LiteralPath (Join-Path $target 'file.txt') -Value 'x'
        $link = Join-Path $script:Tmp ('guard-accepted-link-' + $Tag)
        [void](New-Item -ItemType Junction -Path $link -Target $target)
        Mock Get-PathReparseTag ([scriptblock]::Create("[uint32]0x${Tag}u"))
        try {
            { Assert-PlainPath $link $true } | Should -Not -Throw
            { Assert-PlainPath (Join-Path $link 'file.txt') } | Should -Not -Throw
        }
        finally { (Get-Item -LiteralPath $link -Force).Delete() }
    }
    It 'still refuses a path of the wrong type, whatever its tag' {
        { Assert-PlainPath (Join-Path $script:TagFx 'README.md') $true } | Should -Throw -ExpectedMessage '*unsafe path or unexpected type*'
        { Assert-PlainPath (Join-Path $script:TagFx 'docs') } | Should -Throw -ExpectedMessage '*unsafe path or unexpected type*'
        { Assert-PlainPath (Join-Path $script:TagFx 'README.md') } | Should -Not -Throw
    }
    It 'refuses a symbolic link to a file and to a folder of the package' -Skip:(-not $canSymlink) {
        $fx = New-IntegrityFixture -Name ('symlink-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        $readme = Join-Path $fx 'README.md'
        $copy = Join-Path $script:Tmp ('README-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.copy')
        Move-Item -LiteralPath $readme -Destination $copy
        [void](New-Item -ItemType SymbolicLink -Path $readme -Target $copy)
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        (ConvertTo-FlatText $r.Text) | Should -Match 'reparse point that is not accepted \(tag 0xA000000C\)'
        (Get-Item -LiteralPath $readme -Force).Delete()
        Move-Item -LiteralPath $copy -Destination $readme
        $docs = Join-Path $fx 'docs'
        $docsCopy = Join-Path $script:Tmp ('docs-' + [guid]::NewGuid().ToString('N').Substring(0, 6) + '.copy')
        Move-Item -LiteralPath $docs -Destination $docsCopy
        [void](New-Item -ItemType SymbolicLink -Path $docs -Target $docsCopy)
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        (ConvertTo-FlatText $r.Text) | Should -Match 'reparse point that is not accepted \(tag 0xA000000C\)'
        (Get-Item -LiteralPath $docs -Force).Delete()
    }
    It 'refuses a junction in the package tree, whatever it points to, and names the tag' {
        $fx = New-IntegrityFixture -Name ('junction-tag-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        $link = New-DocsJunction $fx
        try {
            $r = Invoke-Verifier -Root $fx -Distribution
            $r.ExitCode | Should -Not -Be 0
            (ConvertTo-FlatText $r.Text) | Should -Match 'that is not accepted \(tag 0xA0000003\)'
        }
        finally { (Get-Item -LiteralPath $link.Link -Force).Delete() }
    }
    It 'accepts a package in a normal folder' {
        $r = Invoke-Verifier -Root $script:TagFx -Distribution
        $r.ExitCode | Should -Be 0 -Because $r.Text
    }
    It 'accepts a folder whose tag is <Name> and verifies the files in it' -ForEach @(
        @{ Name = 'a OneDrive cloud placeholder'; Tag = 0x9000001Au }
        @{ Name = 'a OneDrive cloud placeholder, variant F'; Tag = 0x9000F01Au }
        @{ Name = 'a deduplicated item'; Tag = 0x80000013u }
        @{ Name = 'a WOF item'; Tag = 0x80000017u }
    ) {
        $fx = New-IntegrityFixture -Name ('accepted-' + $Tag.ToString('X8'))
        Set-FixtureTag $fx $Tag
        $link = New-DocsJunction $fx
        try {
            $r = Invoke-Verifier -Root $fx -Distribution
            $r.ExitCode | Should -Be 0 -Because $r.Text
        }
        finally { (Get-Item -LiteralPath $link.Link -Force).Delete() }
    }
    It 'refuses a folder whose tag is <Name>, even when the tag looks harmless' -ForEach @(
        @{ Name = 'unknown with the name-surrogate bit'; Tag = 0xA000001Du; Text = 'A000001D' }
        @{ Name = 'unknown without it'; Tag = 0x8000001Bu; Text = '8000001B' }
    ) {
        $fx = New-IntegrityFixture -Name ('refused-' + $Text)
        Set-FixtureTag $fx $Tag
        $link = New-DocsJunction $fx
        try {
            $r = Invoke-Verifier -Root $fx -Distribution
            $r.ExitCode | Should -Not -Be 0
            (ConvertTo-FlatText $r.Text) | Should -Match "that is not accepted \(tag 0x$Text\)"
        }
        finally { (Get-Item -LiteralPath $link.Link -Force).Delete() }
    }
    It 'still refuses a changed file inside an accepted reparse point: the SHA256 decides' {
        $fx = New-IntegrityFixture -Name 'accepted-changed'
        Set-FixtureTag $fx 0x9000001Au
        $link = New-DocsJunction $fx
        try {
            (Invoke-Verifier -Root $fx -Distribution).ExitCode | Should -Be 0
            [IO.File]::AppendAllText((Join-Path $link.Target 'MailboxTransfer-UserGuide.html'), "`r`n<!-- synchronised change -->")
            $r = Invoke-Verifier -Root $fx -Distribution
            $r.ExitCode | Should -Not -Be 0
            $r.Text | Should -Match 'SHA256 mismatch: docs\\MailboxTransfer-UserGuide.html'
        }
        finally { (Get-Item -LiteralPath $link.Link -Force).Delete() }
    }
    It 'refuses a package under a junction, in the verifier, the entry script, the console and the module, and names the tag' {
        $via = Join-Path $script:Tmp ('via-junction-' + [guid]::NewGuid().ToString('N').Substring(0, 6))
        [void](New-Item -ItemType Junction -Path $via -Target $script:Tmp)
        try {
            $linked = Join-Path $via (Split-Path $script:TagFx -Leaf)
            $r = Invoke-Verifier -Root $linked -Distribution
            $r.ExitCode | Should -Not -Be 0
            (ConvertTo-FlatText $r.Text) | Should -Match 'reparse-point ancestor that is not accepted \(tag 0xA0000003\)'
            $out = & pwsh -NoProfile -NonInteractive -File (Join-Path $linked 'Invoke-MailboxTransfer.ps1') -Mode Status -StorePath (Join-Path $script:Tmp 'none') 2>&1 | Out-String
            (ConvertTo-FlatText $out) | Should -Match 'that is not accepted \(tag 0xA0000003\)'
            $out = & pwsh -NoProfile -NonInteractive -File (Join-Path $linked 'Start-MailboxTransferConsole.ps1') 2>&1 | Out-String
            (ConvertTo-FlatText $out) | Should -Match 'that is not accepted \(tag 0xA0000003\)'
            $out = & pwsh -NoProfile -NonInteractive -Command "Import-Module '$(Join-Path $linked 'MailboxTransfer.psd1')'" 2>&1 | Out-String
            (ConvertTo-FlatText $out) | Should -Match 'that is not accepted \(tag 0xA0000003\)'
        }
        finally { (Get-Item -LiteralPath $via -Force).Delete() }
    }
    It 'accepts a package under a folder whose tag is a cloud tag, and refuses one whose tag redirects' {
        foreach ($case in @(@{ Tag = 0x9000001Au; Accepted = $true }, @{ Tag = 0xA000000Cu; Accepted = $false })) {
            $fx = New-IntegrityFixture -Name ('via-' + $case.Tag.ToString('X8'))
            Set-FixtureTag $fx $case.Tag
            $via = Join-Path $script:Tmp ('via-cloud-' + $case.Tag.ToString('X8'))
            [void](New-Item -ItemType Junction -Path $via -Target $script:Tmp)
            try {
                $r = Invoke-Verifier -Root (Join-Path $via (Split-Path $fx -Leaf)) -Distribution
                if ($case.Accepted) { $r.ExitCode | Should -Be 0 -Because $r.Text }
                else {
                    $r.ExitCode | Should -Not -Be 0
                    (ConvertTo-FlatText $r.Text) | Should -Match 'reparse-point ancestor that is not accepted \(tag 0xA000000C\)'
                }
            }
            finally { (Get-Item -LiteralPath $via -Force).Delete() }
        }
    }
    It 'refuses a link in the package when the manifest is written, and names the tag' {
        $fx = New-IntegrityFixture -Name ('generator-' + [guid]::NewGuid().ToString('N').Substring(0, 6)) -WithoutManifest
        $link = New-DocsJunction $fx
        try {
            { & (Join-Path $script:Repo 'tools\New-PackageIntegrity.ps1') -Root $fx 6>&1 | Out-Null } | Should -Throw -ExpectedMessage '*not accepted (tag 0xA0000003)*docs*'
        }
        finally { (Get-Item -LiteralPath $link.Link -Force).Delete() }
    }
    It 'accepts a file that Windows compressed with WOF (CompactOS) and keeps judging its bytes' {
        $fx = New-IntegrityFixture -Name 'wof'
        $readme = Join-Path $fx 'README.md'
        $null = & compact.exe /C /EXE:XPRESS4K /F $readme
        $compacted = $LASTEXITCODE -eq 0
        (Invoke-Verifier -Root $fx -Distribution).ExitCode | Should -Be 0
        $tag = if ($compacted) { Get-PathReparseTag $readme }
        if ($null -eq $tag) { Set-ItResult -Skipped -Because 'this Windows build does not expose a reparse tag on a compacted file'; return }
        '{0:X8}' -f $tag | Should -Be '80000017'
        [IO.File]::WriteAllText($readme, ('# another content' * 2000))
        $null = & compact.exe /C /EXE:XPRESS4K /F $readme
        '{0:X8}' -f (Get-PathReparseTag $readme) | Should -Be '80000017'
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        $r.Text | Should -Match 'SHA256 mismatch: README.md'
    }

    It 'maps the Windows error <Code> (<Name>) of a read to the cloud message: <Mapped>' -ForEach @(
        @{ Code = 362; Name = 'ERROR_CLOUD_FILE_PROVIDER_NOT_RUNNING'; Mapped = $true }
        @{ Code = 369; Name = 'ERROR_FILE_SYSTEM_VIRTUALIZATION_UNAVAILABLE'; Mapped = $true }
        @{ Code = 386; Name = 'ERROR_CLOUD_FILE_AUTHENTICATION_FAILED'; Mapped = $true }
        @{ Code = 388; Name = 'ERROR_CLOUD_FILE_NETWORK_UNAVAILABLE'; Mapped = $true }
        @{ Code = 389; Name = 'ERROR_CLOUD_FILE_UNSUCCESSFUL'; Mapped = $true }
        @{ Code = 393; Name = 'ERROR_CLOUD_FILE_REQUEST_ABORTED'; Mapped = $true }
        @{ Code = 398; Name = 'ERROR_CLOUD_FILE_REQUEST_CANCELED'; Mapped = $true }
        @{ Code = 404; Name = 'ERROR_CLOUD_FILE_PROVIDER_TERMINATED'; Mapped = $true }
        @{ Code = 426; Name = 'ERROR_CLOUD_FILE_REQUEST_TIMEOUT'; Mapped = $true }
        @{ Code = 395; Name = 'ERROR_CLOUD_FILE_ACCESS_DENIED'; Mapped = $false }
        @{ Code = 5; Name = 'ERROR_ACCESS_DENIED'; Mapped = $false }
        @{ Code = 32; Name = 'ERROR_SHARING_VIOLATION'; Mapped = $false }
        @{ Code = 2; Name = 'ERROR_FILE_NOT_FOUND'; Mapped = $false }
        @{ Code = 0; Name = 'no error'; Mapped = $false }
    ) {
        Test-CloudReadErrorCode $Code | Should -Be $Mapped
    }
    It 'reads the Windows error from a Win32 HRESULT, a Win32Exception or an inner exception, and ignores the message text' {
        Get-ReadFailureCode ([IO.IOException]::new('any text, in any language', [int]0x80070184)) | Should -Be 388
        Get-ReadFailureCode ([IO.IOException]::new('Le fournisseur de synchronisation du cloud n''a pas pu', [int]0x8007016A)) | Should -Be 362
        Get-ReadFailureCode ([ComponentModel.Win32Exception]::new(426)) | Should -Be 426
        Get-ReadFailureCode ([InvalidOperationException]::new('outer', [IO.IOException]::new('inner', [int]0x80070185))) | Should -Be 389
        Get-ReadFailureCode ([InvalidOperationException]::new('ERROR_CLOUD_FILE_NETWORK_UNAVAILABLE 388')) | Should -Be 0
        Get-ReadFailureCode ([IO.IOException]::new('not a Win32 HRESULT', [int]0x80004005)) | Should -Be 0
        Get-ReadFailureCode $null | Should -Be 0
    }
    It 'names a cloud file that cannot be downloaded, with the way out, when its read fails with <Code>' -ForEach @(
        @{ Code = 388; HResult = [int]0x80070184 }
        @{ Code = 362; HResult = [int]0x8007016A }
        @{ Code = 426; HResult = [int]0x800701AA }
    ) {
        $fx = New-IntegrityFixture -Name ('cloud-read-' + $Code)
        Set-FixtureVerifier $fx 'try { (Microsoft.PowerShell.Utility\Get-FileHash' "try { if (`$Path -like '*\README.md') { throw [IO.IOException]::new('texte localise', $HResult) }; (Microsoft.PowerShell.Utility\Get-FileHash"
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        (ConvertTo-FlatText $r.Text) | Should -Match "this cloud file \(OneDrive\) is not downloaded.*Windows error $Code.*README\.md.*Reconnect this computer to the network.*outside OneDrive"
    }
    It 'leaves the other read failures as they are' {
        $fx = New-IntegrityFixture -Name 'cloud-read-other'
        Set-FixtureVerifier $fx 'try { (Microsoft.PowerShell.Utility\Get-FileHash' "try { if (`$Path -like '*\README.md') { throw [IO.IOException]::new('texte localise', [int]0x80070005) }; (Microsoft.PowerShell.Utility\Get-FileHash"
        $r = Invoke-Verifier -Root $fx -Distribution
        $r.ExitCode | Should -Not -Be 0
        $r.Text | Should -Match 'texte localise'
        $r.Text | Should -Not -Match 'OneDrive'
    }
}
