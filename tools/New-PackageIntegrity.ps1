#Requires -Version 7.6
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
function Get-PathReparseTag {
    <#
        The tag of the reparse point at Path, or $null when Path is not one. WIN32_FIND_DATAW.dwReserved0 holds the tag when FILE_ATTRIBUTE_REPARSE_POINT
        is set; FindFirstFileW reads it from the directory entry without opening the item, so a cloud placeholder is not downloaded. The two calls are
        declared with Reflection.Emit: no C# is compiled, because this runs before the engine exists. Layout of the 592-byte structure: dwFileAttributes
        at offset 0, dwReserved0 at offset 36. The same text is in the verifier, the module, both entry scripts and the two tools of the repository.
    #>
    param([Parameter(Mandatory)][string]$Path)
    $api = Get-Variable -Name FindFileApi -Scope Script -ValueOnly -ErrorAction Ignore
    if ($null -eq $api) {
        $module = [Reflection.Emit.AssemblyBuilder]::DefineDynamicAssembly([Reflection.AssemblyName]::new('ReparseTag'), [Reflection.Emit.AssemblyBuilderAccess]::Run).DefineDynamicModule('ReparseTag')
        $native = $module.DefineType('Native', 'Public, Abstract, Sealed')
        foreach ($call in @(@('FindFirstFileW', [IntPtr], @([string], [IntPtr])), @('FindClose', [bool], @([IntPtr])))) {
            $method = $native.DefinePInvokeMethod($call[0], 'kernel32.dll', $call[0], 'Public, Static, PinvokeImpl', 'Standard', $call[1], [type[]]$call[2], 'Winapi', 'Unicode')
            $method.SetImplementationFlags('PreserveSig')
        }
        $api = $native.CreateType()
        Set-Variable -Name FindFileApi -Scope Script -Value $api
    }
    $full = [IO.Path]::GetFullPath($Path)
    $full = if ($full.StartsWith('\\?\', [StringComparison]::Ordinal)) { $full } elseif ($full.StartsWith('\\', [StringComparison]::Ordinal)) { '\\?\UNC\' + $full.Substring(2) } else { '\\?\' + $full }
    $data = [Runtime.InteropServices.Marshal]::AllocHGlobal(592)
    try {
        $handle = $api::FindFirstFileW($full, $data)
        if ($handle -eq [IntPtr]::new(-1)) { throw "Package integrity: the reparse tag of this path cannot be read (Windows error $([Runtime.InteropServices.Marshal]::GetLastSystemError())): $Path" }
        [void]$api::FindClose($handle)
        if (([Runtime.InteropServices.Marshal]::ReadInt32($data, 0) -band 0x400) -eq 0) { return $null }
        return [uint32]([long][Runtime.InteropServices.Marshal]::ReadInt32($data, 36) -band 0xFFFFFFFFL)
    }
    finally { [Runtime.InteropServices.Marshal]::FreeHGlobal($data) }
}
function Test-ReparseTagAllowed {
    <#
        The rule for the tag of a reparse point, not for the bare attribute. Refused: every tag that redirects to another path, that is every tag with the
        name-surrogate bit 0x20000000 (symbolic link 0xA000000C, junction and mount point 0xA0000003, and the others), and every tag that is not in the closed
        list below (fail closed). Accepted, the Microsoft types that do not redirect: the cloud files of OneDrive (0x9000001A and its variants 0x9000101A to
        0x9000F01A, that is the tag with the bits 0x0000F000 cleared), data deduplication 0x80000013 and WOF (Windows Overlay Filter, CompactOS) 0x80000017.
        The SHA256 of the signed manifest still detects any change of content, including one made by a synchronisation.
    #>
    param([Parameter(Mandatory)][uint32]$Tag)
    if ($Tag -band 0x20000000u) { return $false }
    return (($Tag -band 0xFFFF0FFFu) -eq 0x9000001Au) -or $Tag -eq 0x80000013u -or $Tag -eq 0x80000017u
}
function Get-RefusedReparseTag {
    <# $null when the item (from Get-Item -Force) is not a reparse point or its tag is accepted by Test-ReparseTagAllowed; otherwise the refused tag as eight hexadecimal digits. #>
    param([Parameter(Mandatory)]$Item)
    if (-not ($Item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { return $null }
    $tag = [uint32](Get-PathReparseTag $Item.FullName)
    if (Test-ReparseTagAllowed $tag) { return $null }
    return $tag.ToString('X8')
}
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
        $refused = Get-RefusedReparseTag $item
        if ($refused) { throw "Reparse point that is not accepted (tag 0x$refused) in the package: $relative" }
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
