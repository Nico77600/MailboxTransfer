#Requires -Version 7.4
<#
.SYNOPSIS
    Writes package\PackageIntegrity.psd1: the SHA256 of every file of the package, to be signed by the owner.

.DESCRIPTION
    Run it when every PowerShell file of the package is signed and timestamped (tools\Get-SigningInventory.ps1 lists them): the
    hashes are those of the final bytes, signature blocks included. The list holds every file of package\ except
    PackageIntegrity.psd1 itself, the editable configuration (config\MailboxTransfer.config.psd1 and its local copy) and the
    generated folders, plus CHANGELOG.md and SECURITY.md of the repository root (the package builder copies them into the
    release ZIP). The manifest is then signed and timestamped like the scripts; src\Test-PackageIntegrity.ps1 refuses a
    package whose manifest is not signed.

    The script does not sign and does not verify signatures.

.PARAMETER Root
    Package folder. Default: package\ of this repository.

.PARAMETER OutputPath
    Manifest to write. Default: PackageIntegrity.psd1 in the package folder.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>
[CmdletBinding()]
param([string]$Root, [string]$OutputPath)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repository = Split-Path $PSScriptRoot -Parent
if (-not $Root) { $Root = Join-Path $repository 'package' }
$Root = [IO.Path]::GetFullPath($Root).TrimEnd('\')
if (-not $OutputPath) { $OutputPath = Join-Path $Root 'PackageIntegrity.psd1' }
$version = [regex]::Match([IO.File]::ReadAllText((Join-Path $Root 'MailboxTransfer.psm1')), "(?m)^\`$script:ToolVersion\s*=\s*'([0-9]+\.[0-9]+\.[0-9]+)'").Groups[1].Value
if (-not $version) { throw 'Version not found in MailboxTransfer.psm1 ($script:ToolVersion).' }
$editable = @('config\MailboxTransfer.config.psd1', 'config\MailboxTransfer.config.local.psd1', 'config\MailboxTransfer.config.cache')
$generated = @('bin', 'data', 'logs', 'out')
$entries = [Collections.Generic.List[object]]::new()
$queue = [Collections.Generic.Queue[string]]::new()
$queue.Enqueue($Root)
while ($queue.Count) {
    $directory = $queue.Dequeue()
    foreach ($item in Get-ChildItem -LiteralPath $directory -Force) {
        $relative = [IO.Path]::GetRelativePath($Root, $item.FullName)
        if ($directory -ceq $Root -and $item.PSIsContainer -and $item.Name -cin $generated) { continue }
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Reparse point in the package: $relative" }
        if ($item.PSIsContainer) { $queue.Enqueue($item.FullName); continue }
        if ($relative -ceq 'PackageIntegrity.psd1' -or $relative -cin $editable -or $relative -cin 'CHANGELOG.md', 'SECURITY.md') { continue }
        $entries.Add([pscustomobject]@{ Path = $relative; Source = $item.FullName })
    }
}
foreach ($name in 'CHANGELOG.md', 'SECURITY.md') {
    $source = Join-Path $repository $name
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "$name not found at the repository root." }
    $entries.Add([pscustomobject]@{ Path = $name; Source = $source })
}
$lines = foreach ($e in ($entries | Sort-Object { $_.Path.ToLowerInvariant() })) {
    "        @{{ Path = '{0}'; Sha256 = '{1}' }}" -f $e.Path, (Get-FileHash -LiteralPath $e.Source -Algorithm SHA256).Hash
}
$text = @(
    '@{'
    '    SchemaVersion = 1'
    "    Algorithm = 'SHA256'"
    "    SignerThumbprint = '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6'"
    "    Version = '$version'"
    '    Files = @('
    $lines
    '    )'
    '}'
) | ForEach-Object { $_ }
[IO.File]::WriteAllBytes($OutputPath, [Text.UTF8Encoding]::new($false).GetBytes((($text -join "`r`n") + "`r`n")))
Write-Host ("{0}: {1} files, version {2}. Sign and timestamp it, then run the package builder." -f $OutputPath, $entries.Count, $version)
