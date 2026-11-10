#Requires -Version 7.4
<#
.SYNOPSIS
    Verifies the authenticated Mailbox Transfer package before any of its code runs.

.DESCRIPTION
    Invoked by Invoke-MailboxTransfer.ps1, by Start-MailboxTransferConsole.ps1, by the module (import) and by the package
    builder. It requires a valid, timestamped owner signature on this script, on PackageIntegrity.psd1 and on every
    PowerShell file of the package, then compares the SHA256 of every file listed in the signed manifest (C# sources, SQLite
    libraries, window definitions, HTML guides, README, notices, launcher) with the manifest. Missing, changed, unlisted,
    duplicated or reparse-point files fail closed.

    Not protected, by design: config\MailboxTransfer.config.psd1 and config\MailboxTransfer.config.local.psd1 (edited by the
    administrator, data that is never executed) and the generated folders bin\, data\, logs\ and out\ (no compiled cache
    is ever loaded).

.PARAMETER Root
    Package folder to verify, not its src subfolder. This script must belong to that package's src folder.

.PARAMETER Distribution
    Verifies a release package: the files that the package builder copies from the repository root (CHANGELOG.md,
    SECURITY.md) are required.

.PARAMETER DevelopmentMode
    For a source tree or a developer package: the manifest is still enforced when it is there. Refused together with
    -Distribution.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.1
#>
[CmdletBinding()]
param([Parameter(Mandatory)][string]$Root, [switch]$Distribution, [switch]$DevelopmentMode)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ($Distribution -and $DevelopmentMode) { throw 'Package integrity: -Distribution and -DevelopmentMode cannot be combined.' }
$signer = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'
# What to do, at the end of the refusals an unsigned or re-saved package meets first.
$releaseHint = 'Use the complete release ZIP (https://github.com/Nico77600/MailboxTransfer/releases/latest): extract it again and unblock it. A modified, re-saved or unsigned file cannot be started. A source checkout is started with -DevelopmentMode.'
$rootPath = [IO.Path]::GetFullPath($Root).TrimEnd('\')
# The PowerShell files of the package: every one carries the owner signature, and the manifest must list every one.
$signed = @(
    'Invoke-MailboxTransfer.ps1'
    'Start-MailboxTransferConsole.ps1'
    'MailboxTransfer.psd1'
    'MailboxTransfer.psm1'
    'module\MailboxTransfer.Eml.ps1'
    'module\MailboxTransfer.Events.ps1'
    'module\MailboxTransfer.Gui.Browse.ps1'
    'module\MailboxTransfer.Gui.Core.ps1'
    'module\MailboxTransfer.Gui.DemoJob.ps1'
    'module\MailboxTransfer.Gui.Engine.ps1'
    'module\MailboxTransfer.Gui.Fake.ps1'
    'module\MailboxTransfer.Gui.Jobs.ps1'
    'module\MailboxTransfer.Gui.Pages.Export.ps1'
    'module\MailboxTransfer.Gui.Pages.Run.ps1'
    'module\MailboxTransfer.Gui.Pages.Source.ps1'
    'module\MailboxTransfer.Gui.Pages.ps1'
    'module\MailboxTransfer.Gui.Run.ps1'
    'module\MailboxTransfer.Gui.Scenarios.ps1'
    'module\MailboxTransfer.Gui.State.ps1'
    'src\Test-PackageIntegrity.ps1'
)
# Other files that must be in every manifest (the rest of the inventory is whatever the owner-signed manifest lists).
$required = @($signed) + @(
    'module\MailboxTransfer.Gui.xaml'
    'module\MailboxTransfer.Gui.Browse.xaml'
    'src\Gui.ViewModels.cs'
    'src\Engine.Store.cs'
    'src\Engine.Auth.cs'
    'lib\sqlite\Microsoft.Data.Sqlite.dll'
    'lib\sqlite\SQLitePCLRaw.core.dll'
    'lib\sqlite\SQLitePCLRaw.provider.e_sqlite3.dll'
    'lib\sqlite\SQLitePCLRaw.batteries_v2.dll'
    'lib\sqlite\runtimes\win-x64\e_sqlite3.dll'
    'lib\sqlite\runtimes\win-arm64\e_sqlite3.dll'
    'Start-MailboxTransferConsole.cmd'
    'product.json'
    'README.md'
    'LICENSE'
    'THIRD-PARTY-NOTICES.md'
)
# Copied from the repository root by the package builder: hashed, required only in a distribution.
$distributionOnly = @('CHANGELOG.md', 'SECURITY.md')
# In the manifest (the whole package\ folder is hashed) but not copied into the release ZIP by the package builder: the pictures and the Markdown guides (the HTML guides are self-contained). Verified when present, not required.
$notDistributed = @('docs\MailboxTransfer-Guide.md', 'docs\MailboxTransfer-UserGuide.md')
function Test-NotDistributed([string]$Path) { $Path -cin $notDistributed -or $Path.StartsWith('docs\images\', [StringComparison]::Ordinal) }
# Written or edited by the administrator; data, never executed.
$editableFiles = @('config\MailboxTransfer.config.psd1', 'config\MailboxTransfer.config.local.psd1', 'config\MailboxTransfer.config.cache')
# Created at run time, never loaded as code: exact names, at the package root only.
$generatedFolders = @('bin', 'data', 'logs', 'out')

function Assert-PlainPath([string]$Path, [bool]$Directory = $false) {
    $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $Path -Force -ErrorAction Stop
    if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint -or $item.PSIsContainer -ne $Directory) {
        throw "Package integrity: unsafe path or unexpected type: $Path"
    }
    $parent = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($Path))
    while ($parent) {
        $ancestor = Microsoft.PowerShell.Management\Get-Item -LiteralPath $parent -Force -ErrorAction Stop
        if ($ancestor.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            throw "Package integrity: reparse-point ancestor: $parent"
        }
        $parent = [IO.Path]::GetDirectoryName($parent)
    }
}
function Assert-OwnerSignature([string]$Path) {
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne $signer -or $null -eq $signature.TimeStamperCertificate) {
        throw "Package integrity: expected trusted, timestamped owner signature: $Path. $releaseHint"
    }
}
function Test-RelativePath([string]$Path) {
    if ([string]::IsNullOrEmpty($Path) -or $Path.StartsWith('\') -or $Path.Contains('/') -or $Path.Contains(':')) { return $false }
    foreach ($part in $Path.Split('\')) { if ($part -eq '' -or $part -eq '.' -or $part -eq '..') { return $false } }
    return $true
}

Assert-PlainPath $rootPath $true
if (-not [string]::Equals([IO.Path]::GetFullPath($PSCommandPath),
    [IO.Path]::Combine($rootPath, 'src', 'Test-PackageIntegrity.ps1'), [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Package integrity: verifier must belong to this package root.'
}
Assert-PlainPath $PSCommandPath
Assert-OwnerSignature $PSCommandPath
$manifestPath = [IO.Path]::Combine($rootPath, 'PackageIntegrity.psd1')
if (-not [IO.File]::Exists($manifestPath)) { throw "Package integrity: PackageIntegrity.psd1 is missing. $releaseHint" }
Assert-PlainPath $manifestPath
Assert-OwnerSignature $manifestPath
# Authenticate the raw PSD1 before parsing its restricted data language.
$manifest = Microsoft.PowerShell.Utility\Import-PowerShellDataFile -LiteralPath $manifestPath -ErrorAction Stop
if ($manifest.Count -ne 5 -or -not $manifest.ContainsKey('SchemaVersion') -or
    -not $manifest.ContainsKey('Algorithm') -or -not $manifest.ContainsKey('SignerThumbprint') -or
    -not $manifest.ContainsKey('Version') -or -not $manifest.ContainsKey('Files') -or
    $manifest.SchemaVersion -isnot [int] -or $manifest.SchemaVersion -ne 1 -or $manifest.Algorithm -cne 'SHA256' -or
    $manifest.SignerThumbprint -cne $signer -or $manifest.Version -isnot [string] -or $manifest.Files -isnot [array]) {
    throw 'Package integrity: unsupported or incomplete manifest schema.'
}
$listed = [Collections.Generic.Dictionary[string, string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($entry in $manifest.Files) {
    if ($entry -isnot [hashtable] -or $entry.Count -ne 2 -or -not $entry.ContainsKey('Path') -or -not $entry.ContainsKey('Sha256') -or
        $entry.Path -isnot [string] -or $entry.Sha256 -isnot [string] -or $entry.Sha256 -cnotmatch '^[0-9A-F]{64}$' -or
        -not (Test-RelativePath $entry.Path) -or $entry.Path -ceq 'PackageIntegrity.psd1' -or $entry.Path -cin $editableFiles -or
        -not $listed.TryAdd($entry.Path, $entry.Sha256)) {
        throw 'Package integrity: invalid, duplicate or unapproved inventory path/hash.'
    }
}
foreach ($path in $required + $distributionOnly) {
    if (-not $listed.ContainsKey($path)) { throw "Package integrity: missing inventory path: $path" }
}
foreach ($pair in $listed.GetEnumerator()) {
    $path = [IO.Path]::Combine($rootPath, $pair.Key)
    if (-not [IO.File]::Exists($path) -and (($pair.Key -cin $distributionOnly -and -not $Distribution) -or (Test-NotDistributed $pair.Key))) { continue }
    Assert-PlainPath $path
    $hash = (Microsoft.PowerShell.Utility\Get-FileHash -LiteralPath $path -Algorithm SHA256 -ErrorAction Stop).Hash
    if ($hash -cne $pair.Value) { throw "Package integrity: SHA256 mismatch: $($pair.Key)" }
    if ($pair.Key -cmatch '\.ps(m1|d1|1)$') { Assert-OwnerSignature $path }
}
foreach ($path in $signed) {
    if (-not $listed.ContainsKey($path)) { throw "Package integrity: unsigned inventory path: $path" }
}
# Every file of the tree is in the manifest (or is editable data, or a generated folder): nothing else is allowed to sit in the package.
$queue = [Collections.Generic.Queue[string]]::new()
$queue.Enqueue($rootPath)
while ($queue.Count) {
    $directory = $queue.Dequeue()
    foreach ($item in Microsoft.PowerShell.Management\Get-ChildItem -LiteralPath $directory -Force -ErrorAction Stop) {
        $relative = [IO.Path]::GetRelativePath($rootPath, $item.FullName)
        if ($directory -ceq $rootPath -and $item.PSIsContainer -and $item.Name -cin $generatedFolders) { continue }
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: protected reparse point: $relative" }
        if ($item.PSIsContainer) { $queue.Enqueue($item.FullName); continue }
        # The owner-signed manifest itself (authenticated above, before it was parsed).
        if ($relative -ceq 'PackageIntegrity.psd1') { continue }
        if ($relative -cin $editableFiles) { continue }
        if ($listed.ContainsKey($relative)) { continue }
        throw "Package integrity: unexpected stable file: $relative"
    }
}
[pscustomobject]@{
    Name    = 'MailboxTransfer'
    Version = $manifest.Version
    Files   = $manifest.Files.Count
    Status  = $(if ($DevelopmentMode) { 'Valid (DevelopmentMode)' } else { 'Valid' })
    Signer  = $signer
}
