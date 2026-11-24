#Requires -Version 7.4
<#
.SYNOPSIS
    Copies the files needed to run Mailbox Transfer into a separate folder (and, with -Zip, into the release ZIP).

.DESCRIPTION
    The runtime lives in the package\ folder of the repository. The distribution keeps the same paths:
        Invoke-MailboxTransfer.ps1, Start-MailboxTransferConsole.ps1 and .cmd, MailboxTransfer.psd1 and .psm1, PackageIntegrity.psd1,
        product.json, README.md, LICENSE, THIRD-PARTY-NOTICES.md, CHANGELOG.md and SECURITY.md (copied from the repository root),
        module\ (console, events, EML), src\ (C# engine sources and the verifier), lib\sqlite\, config\MailboxTransfer.config.psd1,
        docs\ (the two HTML guides, which carry their pictures, and the events contract).
    It never copies tests\, tools\, out\, the local configuration (config\*.local.psd1) or the Markdown guides and the pictures of
    docs\images (the HTML guides are self-contained).

    The source and the finished package must pass the owner-signed integrity check (src\Test-PackageIntegrity.ps1 and
    PackageIntegrity.psd1). A source that is unsigned or changed is refused before anything is written, unless -DevelopmentMode is
    given (a dry run on a tree that is not signed yet: nothing is verified, the package is not a release).

.PARAMETER Destination
    Package folder. Default: MailboxTransfer-<version> next to the repository folder.

.PARAMETER Zip
    Also writes <Destination>.zip and prints its size and SHA256 (the values the release notes and the asset inventory need).

.PARAMETER Force
    Replace the destination folder (and ZIP) if it already exists.

.PARAMETER DevelopmentMode
    Build from a tree that is not signed or not hashed yet (a dry run). The result is not verified.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>
[CmdletBinding()]
param([string]$Destination, [switch]$Zip, [switch]$Force, [switch]$DevelopmentMode)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$packageRoot = Join-Path $root 'package'
$version = [regex]::Match([IO.File]::ReadAllText((Join-Path $packageRoot 'MailboxTransfer.psm1')), "(?m)^\`$script:ToolVersion\s*=\s*'([0-9]+\.[0-9]+\.[0-9]+)'").Groups[1].Value
if (-not $version) { throw 'Version not found in package\MailboxTransfer.psm1 ($script:ToolVersion).' }

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

function Assert-AuthenticatedPackage([string]$Path, [switch]$Distribution) {
    $verifier = Join-Path $Path 'src\Test-PackageIntegrity.ps1'
    $cursor = [IO.Path]::GetFullPath($verifier)
    while ($cursor) {
        $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        $refused = Get-RefusedReparseTag $item
        if ($refused) { throw "Package integrity: reparse point that is not accepted (tag 0x$refused): $cursor" }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $verifier -ErrorAction Stop
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or $null -eq $signature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature.'
    }
    if ($Distribution) { $null = & $verifier -Root $Path -Distribution } else { $null = & $verifier -Root $Path }
}
if (-not $DevelopmentMode) { Assert-AuthenticatedPackage $packageRoot }
if (-not $Destination) { $Destination = Join-Path (Split-Path $root -Parent) "MailboxTransfer-$version" }
if (-not [IO.Path]::IsPathRooted($Destination)) { $Destination = Join-Path (Get-Location).Path $Destination }
$Destination = [IO.Path]::GetFullPath($Destination).TrimEnd('\')
$rootPrefix = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
if (($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be outside the repository: $Destination"
}
foreach ($path in @($Destination, "$Destination.zip")) {
    if (Test-Path -LiteralPath $path) {
        if (-not $Force) { throw "The destination already exists: $path. Use -Force to replace it." }
        if ((Test-Path -LiteralPath $path -PathType Container) -and -not (Test-Path -LiteralPath (Join-Path $path 'Invoke-MailboxTransfer.ps1'))) { throw "The destination is not a Mailbox Transfer package, it is not replaced: $path" }
        Remove-Item -LiteralPath $path -Recurse -Force
    }
}

# ---- Files needed at run time (paths relative to package\, the same as in the extracted ZIP) -----------------
$skipDirectories = @('docs\images')
$skipFiles = @('docs\MailboxTransfer-Guide.md', 'docs\MailboxTransfer-UserGuide.md', 'config\MailboxTransfer.config.local.psd1', 'config\MailboxTransfer.config.cache')
$files = [Collections.Generic.List[string]]::new()
foreach ($file in Get-ChildItem -LiteralPath $packageRoot -Recurse -File -Force) {
    $relative = [IO.Path]::GetRelativePath($packageRoot, $file.FullName)
    if ($relative -cin $skipFiles) { continue }
    if ($skipDirectories | Where-Object { $relative.StartsWith($_ + '\', [StringComparison]::OrdinalIgnoreCase) }) { continue }
    if ($relative -match '^(bin|data|logs|out)\\') { continue }
    $files.Add($relative)
}
foreach ($f in $files) {
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath (Join-Path $packageRoot $f) -Destination $target
}
# The two documents copied from the repository root (hashed by the manifest).
foreach ($name in 'CHANGELOG.md', 'SECURITY.md') {
    if (-not (Test-Path -LiteralPath (Join-Path $root $name) -PathType Leaf)) { throw "Missing file at the repository root: $name" }
    Copy-Item -LiteralPath (Join-Path $root $name) -Destination (Join-Path $Destination $name)
    $files.Add($name)
}

# ---- Checks ---------------------------------------------------------------------------------------------------
$problems = [Collections.Generic.List[string]]::new()
foreach ($name in 'tests', 'tools', 'out', 'logs', 'data', 'bin') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
foreach ($required in 'Invoke-MailboxTransfer.ps1', 'Start-MailboxTransferConsole.ps1', 'Start-MailboxTransferConsole.cmd', 'MailboxTransfer.psd1', 'MailboxTransfer.psm1', 'product.json',
    'README.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md', 'CHANGELOG.md', 'SECURITY.md', 'config\MailboxTransfer.config.psd1', 'docs\MailboxTransfer-Guide.html', 'docs\MailboxTransfer-UserGuide.html', 'src\Test-PackageIntegrity.ps1') {
    if (-not (Test-Path -LiteralPath (Join-Path $Destination $required) -PathType Leaf)) { $problems.Add("Expected file missing from the package: $required") }
}
if (-not $DevelopmentMode -and -not (Test-Path -LiteralPath (Join-Path $Destination 'PackageIntegrity.psd1'))) { $problems.Add('PackageIntegrity.psd1 is missing.') }
# No private value: the shipped configuration holds defaults only (the local file is never copied).
$config = Import-PowerShellDataFile -LiteralPath (Join-Path $Destination 'config\MailboxTransfer.config.psd1')
foreach ($key in 'TenantId', 'ClientId', 'UserPrincipalName', 'UserObjectId', 'PrimaryMailboxId', 'ArchiveMailboxId') {
    foreach ($section in 'Tenant', 'Lab') {
        if ($config.ContainsKey($section) -and $config[$section].ContainsKey($key) -and $config[$section][$key]) { $problems.Add("The shipped configuration holds a value for $section.$key.") }
    }
}
foreach ($file in Get-ChildItem -LiteralPath $Destination -Recurse -File | Where-Object { $_.Extension -in '.fxs', '.sqlite', '.pst', '.eml' }) {
    $problems.Add("Data file in the package: $($file.FullName.Substring($Destination.Length + 1))")
}
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }
if (-not $DevelopmentMode) {
    try { Assert-AuthenticatedPackage $Destination -Distribution }
    catch { Remove-Item -LiteralPath $Destination -Recurse -Force -ErrorAction SilentlyContinue; throw }
}
$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  Mailbox Transfer $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
Write-Host ('  Content  : {0} files, {1:N1} MB' -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host $(if ($DevelopmentMode) { '  Integrity: NOT verified (-DevelopmentMode): this is a dry run, not a release' } else { '  Integrity: owner signatures and SHA256 inventory verified' })
if ($Zip) {
    Compress-Archive -Path (Join-Path $Destination '*') -DestinationPath "$Destination.zip" -CompressionLevel Optimal
    $z = Get-Item -LiteralPath "$Destination.zip"
    Write-Host ('  ZIP      : {0} ({1} bytes, SHA256 {2})' -f $z.FullName, $z.Length, (Get-FileHash -LiteralPath $z.FullName -Algorithm SHA256).Hash)
}
