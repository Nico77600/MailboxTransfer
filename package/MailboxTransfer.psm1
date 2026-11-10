#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer - PowerShell module.

.DESCRIPTION
    Helper functions used by Invoke-MailboxTransfer.ps1 and by the lab scripts in tools\, organised in regions in
    the order of an execution:

        1. Console                 Write-Mbt* and Format-Mbt* functions (what the administrator sees)
        2. Configuration           Import-MbtConfiguration
        3. Engine                  Initialize-MbtEngine (SQLite from lib\sqlite + C# compiled in memory), Initialize-MbtAuthEngine
        4. Sign-in                 Connect-MbtGraph (browser, -DeviceCode, or app-only with -CertificateThumbprint / -CertificatePath), Get-MbtAccessToken, Get-MbtTokenInfo, Get-MbtCertificate, Disconnect-MbtGraph
        5. Mailbox                 Test-MbtMailboxId, ConvertTo-MbtMailboxId, Get-MbtPrimaryMailboxId, Get-MbtMailboxLocation (primary, archive, auxiliary
                                   archives of an auto-expanding archive), Read-MbtMailboxLocationFile, Get-MbtMailboxTree, Resolve-MbtFolder,
                                   Get-MbtFolderItems, New-MbtFolderPath (creates missing folders, create only)
        5b. Browse                 Get-MbtBrowseLocation, Get-MbtBrowseFolder, Get-MbtBrowseItem, Get-MbtBrowsePreview, selection -> export plan
                                   (live view of a mailbox before the export: Mail API for Primary and Archive, import-export API for an auxiliary archive)
        6. Export                  Export-MbtItems (streams every item to a file, writes manifest.json)
        7. Store                   Open-MbtExportStore, Save-MbtFolderTree, Save-MbtItemList, Export-MbtItemsToStore, Get-MbtStoreItems,
                                   Import-MbtStoreItem (catalog.db + packs: folders, items, streams, previews, import jobs)
        8. Decode                  Read-MbtFxStream (file or store item), Show-MbtFxReport (MS-OXCFXICS reader)
        9. Import                  Import-MbtItem (from a file), Import-MbtStoreItem (from the store): creation only
       10. Transfer                Whole mailbox export and import, status
       11. PST                     Test-MbtPst (integrity checker), Repair-MbtPst (repair into a new file); the PST writer and reader are in the engine (PstWriter, PstReader)

    The Microsoft Graph calls, the streaming of the exported data and the FastTransfer reader are in the C# engine
    (src\Engine.*.cs), compiled in memory the first time they are used.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

    As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.
#>
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$script:ToolVersion = '1.0.0'
$script:ToolRoot = $PSScriptRoot
# Package integrity. A release package carries the owner-signed PackageIntegrity.psd1 manifest; when it is present, the signed
# verifier (src\Test-PackageIntegrity.ps1) must accept the tree before any engine code is compiled. A source tree without a manifest
# is an unverified development tree: the entry scripts refuse it unless -DevelopmentMode is given, and the module says so once.
function Assert-MbtPackageIntegrity {
    [CmdletBinding()]
    param(
        [string]$Root = $script:ToolRoot,
        [switch]$Distribution,
        [switch]$DevelopmentMode,
        [switch]$RequireManifest
    )
    $packageRoot = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($Root))
    $manifest = Join-Path $packageRoot 'PackageIntegrity.psd1'
    if ((Test-Path -LiteralPath $manifest) -and -not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        throw 'Package integrity: PackageIntegrity.psd1 must be a file.'
    }
    if (-not (Test-Path -LiteralPath $manifest -PathType Leaf)) {
        if ($RequireManifest -and -not $DevelopmentMode) {
            throw 'Package integrity: PackageIntegrity.psd1 is missing. This command requires the verified package; use -DevelopmentMode explicitly for an unverified source checkout.'
        }
        # Without a manifest nothing is verified: say so once in every process that imports the module (the flag lives in the process).
        if (-not [AppDomain]::CurrentDomain.GetData('MailboxTransfer.UnverifiedTreeWarned')) {
            [AppDomain]::CurrentDomain.SetData('MailboxTransfer.UnverifiedTreeWarned', $true)
            Write-Warning 'Package integrity: no PackageIntegrity.psd1; this is an unverified source tree (package byte integrity and publisher authentication are not established).'
        }
        return
    }
    $verifier = Join-Path $packageRoot 'src\Test-PackageIntegrity.ps1'
    $cursor = [IO.Path]::GetFullPath($verifier)
    while ($cursor) {
        if (-not (Test-Path -LiteralPath $cursor)) { throw 'Package integrity: the verifier is missing; no engine or SQLite library was loaded.' }
        $item = Microsoft.PowerShell.Management\Get-Item -LiteralPath $cursor -Force -ErrorAction Stop
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $cursor" }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    $signature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $verifier -ErrorAction Stop
    if ([string]$signature.Status -cne 'Valid' -or $null -eq $signature.SignerCertificate -or
        $signature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
        $null -eq $signature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature.'
    }
    $arguments = @{ Root = $packageRoot }
    if ($Distribution) { $arguments.Distribution = $true }
    if ($DevelopmentMode) { $arguments.DevelopmentMode = $true }
    $result = @(& $verifier @arguments)
    if ($result.Count -ne 1 -or -not $result[0].PSObject.Properties['Status'] -or $result[0].Status -notlike 'Valid*') {
        throw 'Package integrity: the verifier did not return an explicit Valid result; engine initialization is blocked.'
    }
}
Assert-MbtPackageIntegrity
$script:Session = $null
# Engine sources compiled together (Engine.Store.cs needs the SQLite assemblies of lib\sqlite, loaded first).
# Engine.Auth.cs is compiled apart (it needs MSAL.NET).
$script:EngineFiles = @('Engine.Graph.cs', 'Engine.JsonStream.cs', 'Engine.Mailbox.cs', 'Engine.FastTransfer.cs', 'Engine.FxModel.cs', 'Engine.FxStream.cs', 'Engine.PstNdb.cs', 'Engine.PstLtp.cs', 'Engine.PstFile.cs', 'Engine.PstWriter.cs', 'Engine.PstReader.cs', 'Engine.PstCompare.cs', 'Engine.FxDiff.cs', 'Engine.FxImport.cs', 'Engine.PstImportPlan.cs', 'Engine.PstImportJob.cs', 'Engine.PstChecker.cs', 'Engine.PstScan.cs', 'Engine.PstNameMap.cs', 'Engine.PstSalvagedBlob.cs', 'Engine.PstRepair.cs', 'Engine.Message.cs', 'Engine.Store.cs', 'Engine.Transfer.cs', 'Engine.Export.cs', 'Engine.Import.cs', 'Engine.Browse.cs', 'Engine.PstExport.cs', 'Engine.Mime.cs', 'Engine.EmlRtf.cs', 'Engine.EmlHtml.cs', 'Engine.EmlMapi.cs', 'Engine.Eml.cs', 'Engine.EmlImport.cs', 'Engine.FolderFilter.cs', 'Engine.PstPreview.cs', 'Engine.PstCopy.cs', 'Engine.Events.cs')

#region 1. Console ---------------------------------------------------------------------------

# Console theme (the same rules as PRA Remote Mailbox and Exchange Log Report):
#   - icons: emoji in Windows Terminal and VS Code, characters of the classic console fonts elsewhere; MBT_ICONS = Emoji | Symbols | Ascii forces a style.
#   - colours: none when the output is redirected (a scheduled task, the console of the window) or when NO_COLOR is set.
#   - every line that matters is also written to the log file (Start-MbtLog), without colours or icons.
function Get-MbtIconSet {
    <# Icons of one console style. Symbols: only characters of the classic console fonts. #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)
    $u = { param([int]$Code) [char]::ConvertFromUtf32($Code) }
    switch ($Style) {
        'Emoji' { return @{
                Logo = & $u 0x1F4EC; Ok = & $u 0x2705; Warn = (& $u 0x26A0) + [char]0xFE0F; Fail = & $u 0x274C; Info = & $u 0x1F539; Skip = & $u 0x23E9
                Plan = & $u 0x1F50E; Mode = & $u 0x1F3AF; Store = & $u 0x1F4BE; Folder = & $u 0x1F4C1; Log = & $u 0x1F4DD; Done = & $u 0x1F389; Clock = & $u 0x23F3
            } }
        'Symbols' { return @{
                Logo = & $u 0x2666; Ok = & $u 0x221A; Warn = & $u 0x25B2; Fail = & $u 0x00D7; Info = & $u 0x2022; Skip = & $u 0x00BB
                Plan = & $u 0x25BA; Mode = & $u 0x25D9; Store = & $u 0x25A0; Folder = & $u 0x2302; Log = & $u 0x00B6; Done = & $u 0x221A; Clock = & $u 0x25CB
            } }
        default { return @{
                Logo = '*'; Ok = '+'; Warn = '!'; Fail = 'x'; Info = '-'; Skip = '>'; Plan = '?'; Mode = 'o'; Store = '#'; Folder = '>'; Log = '='; Done = '*'; Clock = '~'
            } }
    }
}

function Get-MbtFrameSet {
    <# Frame characters: rounded corners, ASCII for the Ascii style. #>
    param([Parameter(Mandatory)][ValidateSet('Emoji', 'Symbols', 'Ascii')][string]$Style)
    if ($Style -eq 'Ascii') { return @{ TopLeft = [char]'+'; TopRight = [char]'+'; BottomLeft = [char]'+'; BottomRight = [char]'+'; Horizontal = [char]'-'; Vertical = [char]'|' } }
    return @{ TopLeft = [char]0x256D; TopRight = [char]0x256E; BottomLeft = [char]0x2570; BottomRight = [char]0x256F; Horizontal = [char]0x2500; Vertical = [char]0x2502 }
}

$script:IconStyle = if ($env:MBT_ICONS -in 'Emoji', 'Symbols', 'Ascii') { $env:MBT_ICONS }
    elseif ([Console]::IsOutputRedirected) { 'Symbols' }
    elseif ($env:WT_SESSION -or $env:TERM_PROGRAM -eq 'vscode') { 'Emoji' }
    else { 'Symbols' }
$script:Icons = Get-MbtIconSet $script:IconStyle
$script:Frame = Get-MbtFrameSet $script:IconStyle
# Emoji are two columns wide, symbols one: pad symbols so that the text stays aligned.
$script:IconPad = if ($script:IconStyle -eq 'Emoji') { ' ' } else { '  ' }
# Console colours of the theme. Accent = frames and step headers.
$script:Theme = @{ Accent = 'Magenta'; Dim = 'DarkGray'; Ok = 'Green'; Warn = 'Yellow'; Fail = 'Red'; Info = 'Cyan' }
$script:LogPath = $null
$script:LogFailed = $false

function Get-MbtIcon {
    param([Parameter(Mandatory)][string]$Name)
    $icon = $script:Icons[$Name]
    if (-not $icon) { $icon = $script:Icons['Info'] }
    return $icon + $script:IconPad
}

function Write-MbtHost {
    <# One console line in one colour (empty: the console default). Console output only. #>
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost', '', Justification = 'Single console writer of the tool.')]
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text, [string]$Color)
    if ($Color -and -not [Console]::IsOutputRedirected -and -not $env:NO_COLOR) { Write-Host $Text -ForegroundColor $Color } else { Write-Host $Text }
}

function Start-MbtLog {
    <#
    .SYNOPSIS
        Opens (or continues) today's log file in a folder and deletes the log files older than the retention. Returns the path.
    .DESCRIPTION
        MailboxTransfer_yyyyMMdd.log, one line per message: "2026-11-09 14:03:12.481 [WARN  ] text". The runs of a day (and the jobs that the
        console starts) append to the same file. The log never contains colours, icons or a secret.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Directory, [ValidateRange(1, 3650)][int]$RetentionDays = 14)
    [void][IO.Directory]::CreateDirectory($Directory)
    $script:LogPath = Join-Path $Directory ('MailboxTransfer_{0:yyyyMMdd}.log' -f (Get-Date))
    $script:LogFailed = $false
    $limit = (Get-Date).AddDays(-$RetentionDays)
    Get-ChildItem -LiteralPath $Directory -Filter 'MailboxTransfer_*.log' -File -ErrorAction SilentlyContinue |
        Where-Object LastWriteTime -LT $limit | Remove-Item -Force -ErrorAction SilentlyContinue
    return $script:LogPath
}

function Stop-MbtLog { $script:LogPath = $null }

function Get-MbtLogPath { return $script:LogPath }

function Write-MbtLog {
    <# One line in the log file only (nothing when no log is open). A log that cannot be written is reported once and never stops a long job. #>
    param([ValidateSet('INFO', 'OK', 'WARN', 'FAIL', 'SKIP', 'STEP', 'DETAIL')][string]$Level = 'INFO', [Parameter(Mandatory)][AllowEmptyString()][string]$Message)
    if (-not $script:LogPath) { return }
    $line = '{0} [{1,-6}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff'), $Level, ($Message -replace '[\r\n]+', ' | ')
    try {
        $stream = [IO.FileStream]::new($script:LogPath, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
        try { $bytes = [Text.UTF8Encoding]::new($false).GetBytes($line + "`r`n"); $stream.Write($bytes, 0, $bytes.Length) } finally { $stream.Dispose() }
    } catch {
        if (-not $script:LogFailed) { $script:LogFailed = $true; Write-Warning "The log file cannot be written ($($script:LogPath)): $($_.Exception.Message)" }
    }
}

function Write-MbtBanner {
    <#
    .SYNOPSIS
        Title card at the start of an execution: a rounded frame with the title, the version and the author, an optional subtitle, then one row per detail.
    .PARAMETER Details
        Ordered list of rows: key = label, value = @(IconName, Text) or plain text.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Title, [string]$Subtitle, [System.Collections.Specialized.OrderedDictionary]$Details)
    $F = $script:Frame; $width = 76
    $right = "v$($script:ToolVersion) $([char]0x00B7) Nicolas Fabert"
    $left = "  $($script:Icons.Logo)  $Title"
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { $script:Icons.Logo.Length }
    $visible = $left.Length - $script:Icons.Logo.Length + $iconWidth
    $gap = [Math]::Max(1, $width - $visible - $right.Length - 2)
    Write-Host ''
    Write-MbtHost ('  ' + $F.TopLeft + ([string]$F.Horizontal * $width) + $F.TopRight) $script:Theme.Accent
    Write-MbtHost ('  ' + $F.Vertical + $left + (' ' * $gap) + $right + '  ' + $F.Vertical) $script:Theme.Accent
    if ($Subtitle) {
        $sub = '     ' + $Subtitle
        if ($sub.Length -gt $width) { $sub = $sub.Substring(0, $width) }
        Write-MbtHost ('  ' + $F.Vertical + $sub.PadRight($width) + $F.Vertical) $script:Theme.Accent
    }
    Write-MbtHost ('  ' + $F.BottomLeft + ([string]$F.Horizontal * $width) + $F.BottomRight) $script:Theme.Accent
    Write-MbtLog STEP "=== $Title v$($script:ToolVersion) ==="
    if ($Details) {
        foreach ($key in $Details.Keys) {
            $value = $Details[$key]
            $icon = '   '; $text = $value
            if ($value -is [array]) { $icon = Get-MbtIcon $value[0]; $text = $value[1] }
            Write-Host ('     {0}{1,-11} {2}' -f $icon, $key, $text)
            Write-MbtLog INFO ('{0}: {1}' -f $key, $text)
        }
    }
}

function Write-MbtStep {
    param([Parameter(Mandatory)][string]$Text, [int]$Number, [int]$Total, [string]$Icon = 'Plan')
    $label = if ($Number -gt 0 -and $Total -gt 0) { '{0,5}  ' -f ('{0}/{1}' -f $Number, $Total) } else { '       ' }
    Write-Host ''
    Write-MbtHost ('  ' + $label + (Get-MbtIcon $Icon) + $Text) $script:Theme.Accent
    Write-MbtLog STEP $(if ($Number -gt 0 -and $Total -gt 0) { "[$Number/$Total] $Text" } else { $Text })
}

function Write-MbtItem {
    <# One indented result line with a status icon (OK, WARN, FAIL, INFO, SKIP), also written to the log. Warnings, failures and skipped lines are coloured. #>
    param([Parameter(Mandatory)][ValidateSet('OK', 'WARN', 'FAIL', 'INFO', 'SKIP')][string]$Status, [Parameter(Mandatory)][AllowEmptyString()][string]$Text, [string]$Icon)
    $key = @{ OK = 'Ok'; WARN = 'Warn'; FAIL = 'Fail'; INFO = 'Info'; SKIP = 'Skip' }[$Status]
    $color = switch ($Status) { 'WARN' { $script:Theme.Warn } 'FAIL' { $script:Theme.Fail } 'SKIP' { $script:Theme.Dim } default { '' } }
    Write-MbtHost ('      ' + (Get-MbtIcon $(if ($Icon) { $Icon } else { $key })) + $Text) $color
    Write-MbtLog $Status $Text
}

function Write-MbtDetail {
    <# A secondary line, indented and dimmed, also written to the log. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    Write-MbtHost "         $Text" $script:Theme.Dim
    Write-MbtLog DETAIL $Text
}

function Write-MbtSummary {
    <#
    .SYNOPSIS
        Final summary card of a run: a frame in the colour of the outcome (green, yellow, red) with one row per value.
    .PARAMETER Values
        Ordered list: key = label, value = @(IconName, Text) or plain text. A text can be an array (several lines under the same label).
    #>
    param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][System.Collections.Specialized.OrderedDictionary]$Values, [ValidateSet('Ok', 'Warn', 'Fail')][string]$Status = 'Ok')
    $F = $script:Frame; $width = 76
    $color = $script:Theme[$Status]
    $icon = $script:Icons[@{ Ok = 'Done'; Warn = 'Warn'; Fail = 'Fail' }[$Status]]
    $iconWidth = if ($script:IconStyle -eq 'Emoji') { 2 } else { $icon.Length }
    $head = " $icon  $Title "
    $rest = [Math]::Max(2, $width - 1 - ($head.Length - $icon.Length + $iconWidth))
    Write-Host ''
    Write-MbtHost ('  ' + $F.TopLeft + $F.Horizontal + $head + ([string]$F.Horizontal * $rest) + $F.TopRight) $color
    foreach ($key in $Values.Keys) {
        $value = $Values[$key]
        $rowIcon = '   '; $text = $value
        if ($value -is [array] -and $value.Count -eq 2 -and $script:Icons.ContainsKey([string]$value[0])) { $rowIcon = Get-MbtIcon $value[0]; $text = $value[1] }
        $first = $true
        foreach ($line in @($text)) {
            $label = if ($first) { '{0,-11}' -f $key } else { ' ' * 11 }
            Write-Host ('     {0}{1} {2}' -f $(if ($first) { $rowIcon } else { '   ' }), $label, $line)
            Write-MbtLog INFO ('Summary - {0}: {1}' -f $key, $line)
            $first = $false
        }
    }
    Write-MbtHost ('  ' + $F.BottomLeft + ([string]$F.Horizontal * $width) + $F.BottomRight) $color
}

function Format-MbtNumber {
    <# A whole number with thousands separators, always the same way (en-US). #>
    param([Parameter(Mandatory)][AllowNull()]$Value)
    if ($null -eq $Value) { return '-' }
    return ([long]$Value).ToString('N0', [Globalization.CultureInfo]::GetCultureInfo('en-US'))
}

function Format-MbtBytes {
    param([Parameter(Mandatory)][double]$Bytes)
    if ($Bytes -lt 0) { return '-' }
    $units = 'B', 'KiB', 'MiB', 'GiB', 'TiB'
    $i = 0
    while ($Bytes -ge 1024 -and $i -lt $units.Count - 1) { $Bytes /= 1024; $i++ }
    if ($i -eq 0) { return ('{0:0} B' -f $Bytes) }
    return ('{0:0.0} {1}' -f $Bytes, $units[$i])
}

function Format-MbtDuration {
    param([Parameter(Mandatory)][double]$Seconds)
    if ($Seconds -lt 1) { return ('{0:0} ms' -f ($Seconds * 1000)) }
    if ($Seconds -lt 120) { return ('{0:0.0} s' -f $Seconds) }
    return ([TimeSpan]::FromSeconds($Seconds)).ToString('hh\:mm\:ss')
}

function Format-MbtRate {
    <# Throughput in MiB/s. #>
    param([Parameter(Mandatory)][double]$Bytes, [Parameter(Mandatory)][double]$Seconds)
    if ($Seconds -le 0) { return '-' }
    return ('{0:0.00} MiB/s' -f ($Bytes / 1MB / $Seconds))
}

#endregion

#region 2. Configuration ---------------------------------------------------------------------

function Import-MbtConfiguration {
    <#
    .SYNOPSIS
        Reads config\MailboxTransfer.config.psd1 (or another file) over the built-in defaults, then the private local file when there is one.
    .DESCRIPTION
        Config\MailboxTransfer.config.local.psd1 is read after the shipped file (same sections and keys). It is
        git-ignored and never part of a package: it holds the private defaults of one machine, typically the Lab section (tenant,
        application, user, mailbox ids) of the lab scripts in tools\. The environment variable MAILBOXTRANSFER_LOCAL_CONFIG names another
        local file (a missing file means none); the tests use it so that no private value can reach them. It is read after any
        file given with -Path too, so that the jobs the console starts keep the same defaults.
    #>
    [CmdletBinding()]
    param([string]$Path = (Join-Path $script:ToolRoot 'config\MailboxTransfer.config.psd1'))
    $cfg = @{
        Graph = @{ BaseUrl = 'https://graph.microsoft.com/v1.0'; MaxAttempts = 6; BaseBackoffSeconds = 2; MaxBackoffSeconds = 60; RequestTimeoutMinutes = 15; ExportBatchSize = 20; ItemPageSize = 0 }
        Auth  = @{ Scopes = @('https://graph.microsoft.com/.default'); RedirectUri = 'http://localhost'; CacheFile = ''; AppTenantId = ''; AppClientId = ''; CertificateThumbprint = ''; CertificatePath = '' }
        Output = @{ Root = '.\out' }
        Lab = @{
            TenantId = ''; ClientId = ''; UserPrincipalName = ''; UserObjectId = ''; PrimaryMailboxId = ''; ArchiveMailboxId = ''; SourceFolder = ''; TargetFolder = ''
        }
    }
    $files = [Collections.Generic.List[string]]::new()
    if (Test-Path -LiteralPath $Path) { $files.Add($Path) }
    $local = if ($null -ne $env:MAILBOXTRANSFER_LOCAL_CONFIG) { $env:MAILBOXTRANSFER_LOCAL_CONFIG } else { Join-Path $script:ToolRoot 'config\MailboxTransfer.config.local.psd1' }
    if ($local -and (Test-Path -LiteralPath $local)) { $files.Add($local) }
    foreach ($configFile in $files) {
        $file = Import-PowerShellDataFile -LiteralPath $configFile
        foreach ($section in @($file.Keys)) {
            if (-not $cfg.ContainsKey($section)) { throw "Configuration $configFile : unknown section '$section'." }
            foreach ($key in @($file[$section].Keys)) {
                if (-not $cfg[$section].ContainsKey($key)) { throw "Configuration $configFile : unknown key '$section.$key'." }
                $cfg[$section][$key] = $file[$section][$key]
            }
        }
    }
    if ($cfg.Graph.ExportBatchSize -lt 1 -or $cfg.Graph.ExportBatchSize -gt 20) { throw 'Graph.ExportBatchSize must be between 1 and 20 (limit of exportItems).' }
    if ([string]::IsNullOrWhiteSpace($cfg.Auth.CacheFile)) { $cfg.Auth.CacheFile = Join-Path $env:LOCALAPPDATA 'MailboxTransfer\token.cache' }
    return [pscustomobject]$cfg
}

#endregion

#region 3. Engine (compiled C#) --------------------------------------------------------------

function Get-MbtSourceHash {
    <# SHA-256 (hex) of C# source files, file names included: identifies the build of an engine. #>
    param([Parameter(Mandatory)][string[]]$Path)
    $sha = [Security.Cryptography.IncrementalHash]::CreateHash([Security.Cryptography.HashAlgorithmName]::SHA256)
    try {
        foreach ($file in $Path) {
            $sha.AppendData([Text.Encoding]::UTF8.GetBytes((Split-Path -Leaf $file) + "`n"))
            $sha.AppendData([IO.File]::ReadAllBytes($file))
        }
        return [BitConverter]::ToString($sha.GetHashAndReset()).Replace('-', '')
    } finally { $sha.Dispose() }
}

function Assert-MbtEngineCurrent {
    <#
    .SYNOPSIS
        Fails at once when the engine type already loaded in this process was compiled from other sources than the ones on disk.
    .DESCRIPTION
        Add-Type assemblies cannot be unloaded nor recompiled under the same type names, so after an update of src\*.cs the
        reloaded module would run against the old engine (missing members found at the end of a long run). The hash of the sources
        is kept in the AppDomain (not in the module, which a reload resets) when the engine is compiled and compared here.
        Nothing happens while the engine is not loaded yet.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Engine', 'Auth')][string]$Name, [Parameter(Mandatory)][string]$Root)
    $typeName, $files = if ($Name -eq 'Auth') { 'MailboxTransfer.GraphAuth', @('Engine.Auth.cs') } else { 'MailboxTransfer.GraphClient', $script:EngineFiles }
    if (-not ($typeName -as [type])) { return $false }
    $paths = @($files | ForEach-Object { Join-Path $Root "src\$_" })
    $missing = @($paths | Where-Object { -not (Test-Path -LiteralPath $_) })
    if ($missing) { throw "Engine source file not found: $($missing -join ', ')" }
    $expected = Get-MbtSourceHash -Path $paths
    $loaded = [AppDomain]::CurrentDomain.GetData("MailboxTransfer.$Name.SourceHash")
    if ($loaded -ne $expected) {
        $was = if ($loaded) { "sources $($loaded.Substring(0, 12))" } else { 'unknown sources' }
        throw ("The MailboxTransfer $($Name.ToLower()) engine loaded in this PowerShell session is an older build ($was, current sources $($expected.Substring(0, 12))). " +
            'A compiled engine cannot be replaced in a running process: open a new PowerShell session (pwsh) to use the updated code.')
    }
    return $true
}

function Initialize-MbtEngine {
    <#
    .SYNOPSIS
        Loads SQLite (lib\sqlite: the native library of the architecture, then the managed assemblies) and compiles the C#
        sources of the engine (the files of $script:EngineFiles) in memory, once per process.
    .DESCRIPTION
        -AdditionalSource: more C# files compiled in the same unit (the test helpers of tests\support). An in-memory assembly
        cannot be referenced by a later Add-Type, so they must be given to the call that compiles the engine; if the engine
        is already compiled in the process they are not added.
    #>
    [CmdletBinding()]
    param([string]$Root = $script:ToolRoot, [string[]]$AdditionalSource = @())
    if (Assert-MbtEngineCurrent -Name Engine -Root $Root) { return }
    $lib = Join-Path $Root 'lib\sqlite'
    $arch = if ([Runtime.InteropServices.RuntimeInformation]::ProcessArchitecture -eq 'Arm64') { 'win-arm64' } else { 'win-x64' }
    $native = Join-Path $lib "runtimes\$arch\e_sqlite3.dll"
    if (-not (Test-Path -LiteralPath $native)) { throw "SQLite native library not found: $native" }
    [void][Runtime.InteropServices.NativeLibrary]::Load($native)
    foreach ($name in 'SQLitePCLRaw.core', 'SQLitePCLRaw.provider.e_sqlite3', 'SQLitePCLRaw.batteries_v2', 'Microsoft.Data.Sqlite') {
        Add-Type -LiteralPath (Join-Path $lib "$name.dll")
    }
    [SQLitePCL.Batteries_V2]::Init()

    $sources = @(foreach ($name in $script:EngineFiles) {
        $file = Join-Path $Root "src\$name"
        if (-not (Test-Path -LiteralPath $file)) { throw "Engine source file not found: $file" }
        $file
    }) + @($AdditionalSource | Where-Object { $_ })
    $references = @((Join-Path $lib 'Microsoft.Data.Sqlite.dll'), (Join-Path $lib 'SQLitePCLRaw.core.dll'),
        'System.Data.Common', 'System.Buffers', 'System.IO', 'System.Runtime.InteropServices', 'System.Net.Http', 'System.Text.Json', 'System.Collections', 'System.Linq', 'System.Runtime', 'System.Memory',
        'System.Security.Cryptography', 'System.Text.RegularExpressions', 'System.Threading', 'System.Threading.Tasks', 'System.Threading.Thread',
        'netstandard', 'System.Net.Primitives', 'System.IO.FileSystem', 'System.Console', 'System.Text.Encoding.Extensions',
        'System.Text.Encoding.CodePages', 'System.Runtime.Extensions', 'System.Collections.Concurrent', 'System.Diagnostics.Process',
        'System.Private.Uri', 'System.Net.Security', 'System.Linq.Expressions', 'System.ComponentModel.Primitives')
    Add-Type -LiteralPath $sources -ReferencedAssemblies $references -IgnoreWarnings -WarningAction SilentlyContinue
    [AppDomain]::CurrentDomain.SetData('MailboxTransfer.Engine.SourceHash', (Get-MbtSourceHash -Path ($sources | Select-Object -First $script:EngineFiles.Count)))
}

function Find-MbtMsal {
    <#
    .SYNOPSIS
        Where MSAL.NET is: lib\msal (if one day shipped), else the Microsoft.Graph.Authentication module (newest version).
    #>
    [CmdletBinding()]
    param([string]$Root = $script:ToolRoot)
    $local = Join-Path $Root 'lib\msal\Microsoft.Identity.Client.dll'
    if (Test-Path -LiteralPath $local) {
        return [pscustomobject]@{ Msal = $local; Abstractions = Join-Path $Root 'lib\msal\Microsoft.IdentityModel.Abstractions.dll' }
    }
    $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module) { throw 'MSAL.NET not found: install the module Microsoft.Graph.Authentication (Install-Module Microsoft.Graph.Authentication -Scope CurrentUser).' }
    $deps = Join-Path $module.ModuleBase 'Dependencies'
    return [pscustomobject]@{ Msal = Join-Path $deps 'Core\Microsoft.Identity.Client.dll'; Abstractions = Join-Path $deps 'Microsoft.IdentityModel.Abstractions.dll' }
}

function Initialize-MbtAuthEngine {
    <#
    .SYNOPSIS
        Loads MSAL.NET and compiles src\Engine.Auth.cs in memory, once per process.
    #>
    [CmdletBinding()]
    param([string]$Root = $script:ToolRoot)
    if (Assert-MbtEngineCurrent -Name Auth -Root $Root) { return }
    $loaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' } | Select-Object -First 1
    if ($loaded) {
        $msalPath = $loaded.Location
    } else {
        $found = Find-MbtMsal -Root $Root
        $abstractionsLoaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Microsoft.IdentityModel.Abstractions' }
        if (-not $abstractionsLoaded -and (Test-Path -LiteralPath $found.Abstractions)) { Add-Type -LiteralPath $found.Abstractions }
        Add-Type -LiteralPath $found.Msal
        $msalPath = $found.Msal
    }
    Add-Type -AssemblyName System.Security.Cryptography.ProtectedData
    $references = @($msalPath, 'System.Security.Cryptography.ProtectedData', 'System.Text.Json', 'System.Linq', 'System.Runtime', 'System.Threading',
        'System.Threading.Tasks', 'System.Collections', 'System.Memory', 'netstandard', 'System.Security.Cryptography',
        'System.Text.Encoding.Extensions', 'System.Runtime.Extensions', 'System.IO.FileSystem', 'System.Console', 'System.Private.Uri',
        'System.ComponentModel.Primitives', 'System.Security.Cryptography.X509Certificates')
    $authSource = Join-Path $Root 'src\Engine.Auth.cs'
    Add-Type -LiteralPath $authSource -ReferencedAssemblies $references -IgnoreWarnings -WarningAction SilentlyContinue
    [AppDomain]::CurrentDomain.SetData('MailboxTransfer.Auth.SourceHash', (Get-MbtSourceHash -Path $authSource))
}

#endregion

#region 4. Sign-in ---------------------------------------------------------------------------

function New-MbtGraphOptions {
    param([Parameter(Mandatory)]$Configuration)
    $options = [MailboxTransfer.GraphOptions]::new()
    $options.BaseUrl = $Configuration.Graph.BaseUrl
    $options.MaxAttempts = [int]$Configuration.Graph.MaxAttempts
    $options.BaseBackoffSeconds = [double]$Configuration.Graph.BaseBackoffSeconds
    $options.MaxBackoffSeconds = [double]$Configuration.Graph.MaxBackoffSeconds
    $options.RequestTimeout = [TimeSpan]::FromMinutes([double]$Configuration.Graph.RequestTimeoutMinutes)
    return $options
}

function New-MbtAuth {
    <# The MSAL.NET sign-in object (not exported): browser, or device code with the line written to the console and to a file. #>
    param(
        [Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$ClientId, [Parameter(Mandatory)][string[]]$Scopes,
        [Parameter(Mandatory)][string]$CacheFile, [Parameter(Mandatory)]$Configuration, [switch]$DeviceCode, [string]$DeviceCodeFile
    )
    Initialize-MbtAuthEngine
    if ($DeviceCode -and -not $DeviceCodeFile) { $DeviceCodeFile = Join-Path (Join-Path (Get-Location).Path $Configuration.Output.Root) 'device-code.txt' }
    return [MailboxTransfer.GraphAuth]::new($ClientId, $TenantId, $Scopes, $CacheFile, $Configuration.Auth.RedirectUri, [bool]$DeviceCode, $(if ($DeviceCode) { $DeviceCodeFile }))
}

function Get-MbtAccessToken {
    <#
    .SYNOPSIS
        An access token of a public client application (delegated), from the token cache, else after a sign-in in the browser
        or with -DeviceCode. Used by tools\Grant-LabConsent.ps1 for the administrator (Connect-MgGraph -DeviceCode stops
        waiting after 120 seconds, this one waits until the code expires). Returns AccessToken, Account and ExpiresOn.
    .DESCRIPTION
        Same engine and same DPAPI cache protection as Connect-MbtGraph; give -CacheFile to keep the cache of another account
        apart (the administrator's one is not the lab user's).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][string[]]$Scopes,
        [string]$CacheFile,
        [string]$UserPrincipalName,
        [switch]$DeviceCode,
        [string]$DeviceCodeFile,
        $Configuration = (Import-MbtConfiguration)
    )
    if (-not $CacheFile) { $CacheFile = Join-Path (Split-Path $Configuration.Auth.CacheFile) 'token-admin.cache' }
    $auth = New-MbtAuth -TenantId $TenantId -ClientId $ClientId -Scopes $Scopes -CacheFile $CacheFile -Configuration $Configuration -DeviceCode:$DeviceCode -DeviceCodeFile $DeviceCodeFile
    $token = $auth.GetToken($(if ($UserPrincipalName) { $UserPrincipalName }), $true, [Threading.CancellationToken]::None)
    return [pscustomobject]@{ AccessToken = $token; Account = $auth.Account; ExpiresOn = $auth.ExpiresOn }
}

function Find-MbtStoreCertificate {
    <# One certificate of the "My" store of a location by thumbprint ($null when it is not there). Separate so that the tests can replace the stores. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Thumbprint, [Parameter(Mandatory)][ValidateSet('CurrentUser', 'LocalMachine')][string]$StoreLocation)
    $store = [Security.Cryptography.X509Certificates.X509Store]::new('My', $StoreLocation)
    try {
        $store.Open([Security.Cryptography.X509Certificates.OpenFlags]::ReadOnly -bor [Security.Cryptography.X509Certificates.OpenFlags]::OpenExistingOnly)
        $found = $store.Certificates.Find([Security.Cryptography.X509Certificates.X509FindType]::FindByThumbprint, $Thumbprint, $false)
        if ($found.Count -gt 0) { return $found[0] }
        return $null
    } finally { $store.Dispose() }
}

function Get-MbtCertificate {
    <#
    .SYNOPSIS
        The certificate of the application (app-only sign-in): by thumbprint from CurrentUser\My then LocalMachine\My, or from a .pfx file.
    .DESCRIPTION
        The certificate must have its private key (the sign-in request is signed with it) and be inside its validity period. A thumbprint is
        accepted as copied from the certificate console (spaces and invisible characters are dropped). A .pfx is loaded in memory only
        (EphemeralKeySet: its key is not written to the profile); -Password is a SecureString.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Thumbprint')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')][string]$Thumbprint,
        [Parameter(Mandatory, ParameterSetName = 'Path')][string]$Path,
        [Parameter(ParameterSetName = 'Path')][securestring]$Password
    )
    $x509 = [Security.Cryptography.X509Certificates.X509Certificate2]
    if ($PSCmdlet.ParameterSetName -eq 'Path') {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Certificate file not found: $Path" }
        $flags = [Security.Cryptography.X509Certificates.X509KeyStorageFlags]::EphemeralKeySet
        try {
            $cert = if ($Password) { $x509::new((Resolve-Path -LiteralPath $Path).ProviderPath, $Password, $flags) } else { $x509::new((Resolve-Path -LiteralPath $Path).ProviderPath, [string]'', $flags) }
        } catch { throw "Cannot open the certificate file $Path (wrong password, or not a .pfx with a private key): $($_.Exception.Message)" }
        $source = "file $Path"
    } else {
        $clean = $Thumbprint -replace '[^0-9A-Fa-f]', ''
        if ($clean.Length -ne 40) { throw "'$Thumbprint' is not a certificate thumbprint (40 hexadecimal characters, the SHA-1 thumbprint shown by the certificate console or Get-ChildItem Cert:\CurrentUser\My)." }
        $cert = $null
        foreach ($location in 'CurrentUser', 'LocalMachine') {
            $cert = Find-MbtStoreCertificate -Thumbprint $clean -StoreLocation $location
            if ($cert) { $source = "$location\My"; break }
        }
        if (-not $cert) { throw "No certificate with the thumbprint $clean in CurrentUser\My or LocalMachine\My." }
    }
    $now = [DateTime]::Now
    $problem = if (-not $cert.HasPrivateKey) { "The certificate ($source) has no private key: the application cannot sign its sign-in request with it." }
    elseif ($cert.NotAfter -lt $now) { "The certificate ($source) expired on $($cert.NotAfter.ToString('yyyy-MM-dd'))." }
    elseif ($cert.NotBefore -gt $now) { "The certificate ($source) is not valid before $($cert.NotBefore.ToString('yyyy-MM-dd HH:mm'))." }
    if ($problem) { $cert.Dispose(); throw $problem }
    return $cert
}

function New-MbtAppAuth {
    <# The app-only sign-in object (not exported): client credentials with a certificate. The tests replace it with one that has no network. #>
    param([Parameter(Mandatory)][string]$TenantId, [Parameter(Mandatory)][string]$ClientId, [Parameter(Mandatory)][string[]]$Scopes, [Parameter(Mandatory)]$Certificate)
    Initialize-MbtAuthEngine
    return [MailboxTransfer.GraphAppAuth]::new($ClientId, $TenantId, $Certificate, $Scopes)
}

function Test-MbtAppOnly {
    <# $true when the session (the current one, or the one given) signed in as the application (certificate) and not as a user. #>
    param($Session = $script:Session)
    return [bool]($Session -and $Session.PSObject.Properties['AppOnly'] -and $Session.AppOnly)
}

function Get-MbtSessionUser {
    <# The user whose mailbox the session works on: the signed-in account (delegated), or the target mailbox owner given at sign-in (app-only: no account). #>
    param($Session = $script:Session)
    if (-not $Session) { return $null }
    if (Test-MbtAppOnly $Session) { return $(if ($Session.PSObject.Properties['UserPrincipalName']) { $Session.UserPrincipalName }) }
    return $(if ($Session.Token) { $Session.Token.Account })
}

function Get-MbtExchangeSettingsPath {
    <#
    .SYNOPSIS
        The paths that answer "which mailboxes has this user" (settings/exchange): always /users/{id}; /me only for a delegated session of that very user.
    .DESCRIPTION
        An app-only token has no user, so /me does not exist for it: every call names the user (or the mailbox id).
    #>
    param([Parameter(Mandatory)][string]$UserPrincipalName, $Session = $script:Session)
    $paths = @("/users/$([uri]::EscapeDataString($UserPrincipalName))/settings/exchange")
    if (-not (Test-MbtAppOnly $Session) -and $Session.Token -and $Session.Token.Account -ieq $UserPrincipalName) { $paths += '/me/settings/exchange' }
    return $paths
}

function Get-MbtAppPermission {
    <#
    .SYNOPSIS
        The APPLICATION permissions (app-only sign-in) a task needs, as the code uses them: one row per permission, with the roles that satisfy it.
    .DESCRIPTION
        Browse (Primary, Archive)  Mail.Read (live view through the Mail API) + MailboxFolder.Read.All and MailboxItem.Read.All (the folders and items the Mail API
                                   does not list: Calendar, Contacts, Tasks ... are merged from the import-export API)
        Browse (AuxArchive)        MailboxFolder.Read.All + MailboxItem.Read.All (the Mail API cannot open an auxiliary archive)
        Export                     MailboxFolder.Read.All + MailboxItem.Read.All (folder tree, inventory) + MailboxItem.Export.All (exportItems; ImportExport.All also does)
        Import, ImportPst          MailboxFolder.ReadWrite.All (creates the folders) + MailboxItem.Read.All (looks up an item whose answer was lost) + MailboxItem.ImportExport.All (importItem)
        ImportEml                  as Import; with -MimeConversion Exchange also Mail.ReadWrite (a draft is created with the Mail API, then exported)
        ExportEml                  as Browse (Mail.Read: the message is read with the Mail API)
        Status, ExportPst, CheckPst, RepairPst  nothing (local), except -Locations / -Unknown -Check (as Browse of the folders)
        User.Read.All              optional in every mode: only to find a mailbox id from an address (GET /users/{id}/settings/exchange); without it give -PrimaryMailboxId / -MailboxId.
        The names are those of the application permissions of Microsoft Graph (the Exchange RBAC for Applications roles carry the same effect, see docs\MailboxTransfer-Guide.md, chapter 4.6).
    #>
    [CmdletBinding()]
    param(
        [ValidateSet('Browse', 'Export', 'Import', 'ImportPst', 'ExportEml', 'ImportEml', 'ExportPst', 'Status', 'CheckPst', 'RepairPst')][string]$Mode = 'Browse',
        [ValidateSet('Primary', 'Archive', 'AuxArchive', 'All')][string]$Mailbox = 'Primary',
        [switch]$ExchangeConversion
    )
    $row = { param($p, $any, $why, $required = $true) [pscustomobject]@{ Permission = $p; AnyOf = @($any); Required = $required; Why = $why } }
    $folderRead = & $row 'MailboxFolder.Read.All' @('MailboxFolder.Read.All', 'MailboxFolder.ReadWrite.All') 'folder tree and folder counts (import-export API)'
    $itemRead = & $row 'MailboxItem.Read.All' @('MailboxItem.Read.All', 'MailboxItem.ReadWrite.All') 'item listing (import-export API)'
    $mailRead = & $row 'Mail.Read' @('Mail.Read', 'Mail.ReadWrite') 'live view with subject, sender and body (Mail API); without it the listing has no subjects'
    $rows = switch ($Mode) {
        'Browse' { if ($Mailbox -eq 'AuxArchive') { @($folderRead, $itemRead) } else { @($mailRead, $folderRead, $itemRead) } }
        'ExportEml' { if ($Mailbox -eq 'AuxArchive') { @($folderRead, $itemRead) } else { @($mailRead, $folderRead, $itemRead) } }
        'Export' { @($folderRead, $itemRead, (& $row 'MailboxItem.Export.All' @('MailboxItem.Export.All', 'MailboxItem.ImportExport.All') 'exportItems')) }
        { $_ -in 'Import', 'ImportPst', 'ImportEml' } {
            @((& $row 'MailboxFolder.ReadWrite.All' @('MailboxFolder.ReadWrite.All') 'create the folders (and read them)'),
                $itemRead,
                (& $row 'MailboxItem.ImportExport.All' @('MailboxItem.ImportExport.All') 'importItem'))
            if ($Mode -eq 'ImportEml' -and $ExchangeConversion) { & $row 'Mail.ReadWrite' @('Mail.ReadWrite') 'draft created with the Mail API (-MimeConversion Exchange)' }
        }
        default { @() }
    }
    $rows = @($rows)
    if ($rows.Count) { $rows += & $row 'User.Read.All' @('User.Read.All', 'User.ReadWrite.All') 'find a mailbox id from an address (settings/exchange); without it give the mailbox id' $false }
    return $rows
}

function Get-MbtMissingAppPermission {
    <# The required application permissions (Get-MbtAppPermission) that the roles of an app-only token do not hold. Token: from Get-MbtTokenInfo (Roles). #>
    [CmdletBinding()]
    param($Token, [string]$Mode = 'Browse', [string]$Mailbox = 'Primary', [switch]$ExchangeConversion)
    $roles = if ($Token -and $Token.PSObject.Properties['Roles']) { @($Token.Roles) } else { @() }
    $args2 = @{ Mode = $Mode; Mailbox = $Mailbox; ExchangeConversion = $ExchangeConversion }
    return @(Get-MbtAppPermission @args2 | Where-Object { $_.Required -and -not @($_.AnyOf | Where-Object { $_ -in $roles }).Count } | ForEach-Object Permission)
}

function Connect-MbtGraph {
    <#
    .SYNOPSIS
        Signs in (delegated, interactive the first time, then silent with the refresh token; or app-only with a certificate) and prepares the Graph
        client of the module. Returns the session.
    .DESCRIPTION
        Delegated: MSAL.NET, public client, redirect URI http://localhost. The first time the sign-in opens the system browser, or with
        -DeviceCode shows a code: the line "DEVICE CODE: <code> URL: <url>" is written to the console at once and to
        -DeviceCodeFile (default: out\device-code.txt, overwritten), then the call waits until the code expires (about 15
        minutes). The token cache is protected by DPAPI (current user) and serves both modes, so a later run of a script does
        not ask again until the refresh token expires.
        App-only (-CertificateThumbprint, or -CertificatePath with -CertificatePassword): client credentials, no user and no prompt: the
        application signs its request with the certificate and gets a token for https://graph.microsoft.com/.default that carries its APPLICATION
        permissions (claim roles). The token is kept in memory only. -UserPrincipalName is then only the OWNER OF THE TARGET MAILBOX (the session
        has no /me: every call names the user or the mailbox id). -TenantId must be a tenant id or domain (not common).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$ClientId,
        [string]$UserPrincipalName,
        [string]$CertificateThumbprint,
        [string]$CertificatePath,
        [securestring]$CertificatePassword,
        [switch]$DeviceCode,
        [string]$DeviceCodeFile,
        [switch]$ForceRefresh,
        # Never prompts (no browser, no device code): fails at once when no valid token is cached.
        [switch]$NoPrompt,
        $Configuration = (Import-MbtConfiguration)
    )
    Initialize-MbtEngine
    if ($CertificateThumbprint -and $CertificatePath) { throw 'Give -CertificateThumbprint or -CertificatePath, not both.' }
    if ($CertificateThumbprint -or $CertificatePath) {
        if ($DeviceCode) { throw '-DeviceCode is a delegated sign-in: it cannot be combined with a certificate (app-only).' }
        if ($TenantId -in 'common', 'organizations', 'consumers') { throw "App-only sign-in needs the tenant id or domain of the tenant, not '$TenantId'." }
        $certificate = if ($CertificateThumbprint) { Get-MbtCertificate -Thumbprint $CertificateThumbprint } else { Get-MbtCertificate -Path $CertificatePath -Password $CertificatePassword }
        $auth = New-MbtAppAuth -TenantId $TenantId -ClientId $ClientId -Scopes ([string[]]$Configuration.Auth.Scopes) -Certificate $certificate
        $first = $auth.GetToken([Threading.CancellationToken]::None, [bool]$ForceRefresh)
        $graph = [MailboxTransfer.GraphClient]::new($auth.Provider(), (New-MbtGraphOptions -Configuration $Configuration))
        $script:Session = [pscustomobject]@{
            Auth = $auth; Graph = $graph; Api = [MailboxTransfer.MailboxApi]::new($graph); Configuration = $Configuration
            TenantId = $TenantId; ClientId = $ClientId; Token = (Get-MbtTokenInfo -AccessToken $first)
            AppOnly = $true; UserPrincipalName = $UserPrincipalName; Certificate = $certificate
        }
        return $script:Session
    }
    $auth = New-MbtAuth -TenantId $TenantId -ClientId $ClientId -Scopes ([string[]]$Configuration.Auth.Scopes) -CacheFile $Configuration.Auth.CacheFile -Configuration $Configuration -DeviceCode:$DeviceCode -DeviceCodeFile $DeviceCodeFile
    $first = $auth.GetToken($(if ($UserPrincipalName) { $UserPrincipalName }), (-not $NoPrompt),[Threading.CancellationToken]::None, [bool]$ForceRefresh)
    $graph = [MailboxTransfer.GraphClient]::new($auth.Provider($(if ($UserPrincipalName) { $UserPrincipalName })), (New-MbtGraphOptions -Configuration $Configuration))
    $script:Session = [pscustomobject]@{
        Auth = $auth; Graph = $graph; Api = [MailboxTransfer.MailboxApi]::new($graph); Configuration = $Configuration
        TenantId = $TenantId; ClientId = $ClientId; Token = (Get-MbtTokenInfo -AccessToken $first)
        AppOnly = $false; UserPrincipalName = $UserPrincipalName; Certificate = $null
    }
    return $script:Session
}

function Set-MbtSession {
    <# Test hook: uses an existing GraphClient (for instance with a fake HTTP handler) instead of signing in. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][MailboxTransfer.GraphClient]$Graph, $Configuration = (Import-MbtConfiguration))
    $script:Session = [pscustomobject]@{
        Auth = $null; Graph = $Graph; Api = [MailboxTransfer.MailboxApi]::new($Graph); Configuration = $Configuration
        TenantId = $null; ClientId = $null; Token = $null; AppOnly = $false; UserPrincipalName = $null; Certificate = $null
    }
    return $script:Session
}

function Get-MbtSession {
    if (-not $script:Session) { throw 'Not connected: run Connect-MbtGraph first.' }
    return $script:Session
}

function Get-MbtTokenInfo {
    <#
    .SYNOPSIS
        Claims of an access token (payload only): account, tenant, application, permissions and expiry.
    .DESCRIPTION
        Scopes is what the token may do, under the names of the DELEGATED permissions: the scp claim of a user token, or, for an app-only token, its
        roles with the suffix .All removed (MailboxItem.ImportExport.All -> MailboxItem.ImportExport; Mail.Read stays Mail.Read) so that one
        set of checks serves both. Roles keeps the application roles as the token carries them; AppOnly says which kind of token it is.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$AccessToken)
    $doc = [MailboxTransfer.Json]::JwtClaims($AccessToken)
    try {
        $root = $doc.RootElement
        $get = { param($n) $v = [Text.Json.JsonElement]::new(); if ($root.TryGetProperty($n, [ref]$v)) { $v.ToString() } }
        $scp = & $get 'scp'
        $exp = & $get 'exp'
        $roles = @()
        $rv = [Text.Json.JsonElement]::new()
        if ($root.TryGetProperty('roles', [ref]$rv) -and $rv.ValueKind -eq 'Array') { $roles = @($rv.EnumerateArray() | ForEach-Object { $_.GetString() }) }
        $appOnly = ((& $get 'idtyp') -eq 'app') -or (-not $scp -and $roles.Count -gt 0)
        $scopes = if ($appOnly) { @($roles | ForEach-Object { $_ -replace '\.All$', '' } | Select-Object -Unique) } else { @(if ($scp) { $scp -split ' ' }) }
        [pscustomobject]@{
            Account   = & $get 'upn'
            Name      = & $get 'name'
            ObjectId  = & $get 'oid'
            TenantId  = & $get 'tid'
            AppId     = & $get 'appid'
            Audience  = & $get 'aud'
            Scopes    = $scopes
            Roles     = $(if ($appOnly) { $roles } else { @() })
            AppOnly   = [bool]$appOnly
            ExpiresOn = if ($exp) { [DateTimeOffset]::FromUnixTimeSeconds([long]$exp) } else { $null }
        }
    } finally { $doc.Dispose() }
}

function Disconnect-MbtGraph {
    [CmdletBinding()]
    param([switch]$ForgetAccount)
    if ($script:Session -and $script:Session.Auth -and $ForgetAccount) { $script:Session.Auth.SignOut() }
    if ($script:Session -and $script:Session.Graph) { $script:Session.Graph.Dispose() }
    if ($script:Session -and $script:Session.PSObject.Properties['Certificate'] -and $script:Session.Certificate) { $script:Session.Certificate.Dispose() }
    $script:Session = $null
}

#endregion

#region 5. Mailbox ---------------------------------------------------------------------------

function Test-MbtMailboxId {
    <#
    .SYNOPSIS
        Does not trust a physical mailbox id: checks its syntax and tenant, whether the API answers for it, and (if the
        signed-in account may read it) whether Graph itself gives it as the primary mailbox of the user.
    .DESCRIPTION
        GET /users/{upn}/settings/exchange returns primaryMailboxId but needs User.Read.All and supports only primary and
        shared mailboxes: an archive mailbox id cannot be resolved this way. The archive id is then only proven usable
        (the folders answer) and recognisable by its content.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$MailboxId,
        [Parameter(Mandatory)][string]$TenantId,
        [string]$UserPrincipalName
    )
    $s = Get-MbtSession
    $r = [ordered]@{ MailboxId = $MailboxId; Syntax = $false; TenantMatches = $false; PrimaryMailboxIdFromGraph = $null; IsPrimaryPerGraph = $null; GraphLookup = 'not tried'; Reachable = $false; Error = $null }
    if ($MailboxId -match '^MBX:(?<m>[0-9a-fA-F-]{36})@(?<t>[0-9a-fA-F-]{36})$') {
        $r.Syntax = $true
        $r.TenantMatches = $Matches.t -eq $TenantId
    }
    if ($UserPrincipalName) {
        $paths = Get-MbtExchangeSettingsPath -UserPrincipalName $UserPrincipalName -Session $s
        foreach ($path in $paths) {
            try {
                $doc = $s.Graph.GetJson($s.Graph.Url($path))
                try {
                    $v = [Text.Json.JsonElement]::new()
                    if ($doc.RootElement.TryGetProperty('primaryMailboxId', [ref]$v)) {
                        $r.PrimaryMailboxIdFromGraph = $v.GetString()
                        $r.IsPrimaryPerGraph = $v.GetString() -eq $MailboxId
                        $r.GraphLookup = 'ok'
                    }
                } finally { $doc.Dispose() }
                break
            } catch {
                $ge = Get-MbtGraphException $_
                if (-not $ge) { throw }
                $r.GraphLookup = "HTTP $($ge.StatusCode) $($ge.ErrorCode)"
            }
        }
    }
    try {
        $doc = $s.Graph.GetJson($s.Graph.Url([MailboxTransfer.MailboxApi]::MailboxPath($MailboxId) + '/folders?$top=1'))
        $doc.Dispose()
        $r.Reachable = $true
    } catch {
        $ge = Get-MbtGraphException $_
        if (-not $ge) { throw }
        $r.Error = $ge.Message
    }
    return [pscustomobject]$r
}

function ConvertTo-MbtMailboxId {
    <#
    .SYNOPSIS
        Physical mailbox id "MBX:<mailbox guid>@<tenant id>" from that form or from the bare mailbox guid (the tenant id is added).
    .DESCRIPTION
        The guids of a user's mailboxes come from Exchange Online PowerShell:
            Get-MailboxLocation -User <upn> | Format-Table MailboxLocationType, MailboxGuid
        (MailboxLocationType Primary, MainArchive, AuxArchive).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Value, [Parameter(Mandatory)][string]$TenantId)
    $v = $Value.Trim()
    $guid = [guid]::Empty
    if ($v -match '^MBX:(?<m>[^@]+)@(?<t>.+)$' -and [guid]::TryParse($Matches.m, [ref]$guid) -and [guid]::TryParse($Matches.t, [ref]$guid)) { return $v }
    if ([guid]::TryParse($v, [ref]$guid)) { return "MBX:$($guid.ToString())@$TenantId" }
    throw "'$Value' is not a mailbox id: expected MBX:<mailbox guid>@<tenant id> (or the bare mailbox guid; an SMTP address / UPN is resolved by Resolve-MbtMailboxAddress). Get it with Exchange Online PowerShell: Get-MailboxLocation -User <upn>."
}

function Get-MbtPrimaryMailboxId {
    <#
    .SYNOPSIS
        Primary mailbox id of a user through GET /users/{upn}/settings/exchange (primaryMailboxId).
    .DESCRIPTION
        Needs the delegated permission User.Read.All (not part of the lab consent). Without it the error says so and asks for
        the id to be given (-PrimaryMailboxId). The archive mailbox id has no such API: it comes from Get-MailboxLocation.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserPrincipalName, [Parameter(Mandatory)][string]$TenantId)
    $s = Get-MbtSession
    $paths = Get-MbtExchangeSettingsPath -UserPrincipalName $UserPrincipalName -Session $s
    $why = 'no answer'
    foreach ($path in $paths) {
        try {
            $doc = $s.Graph.GetJson($s.Graph.Url($path))
            try {
                $v = [Text.Json.JsonElement]::new()
                if ($doc.RootElement.TryGetProperty('primaryMailboxId', [ref]$v) -and $v.GetString()) { return (ConvertTo-MbtMailboxId -Value $v.GetString() -TenantId $TenantId) }
                $why = 'the answer has no primaryMailboxId'
            } finally { $doc.Dispose() }
        } catch {
            $ge = Get-MbtGraphException $_
            if (-not $ge) { throw }
            $why = "HTTP $($ge.StatusCode) $($ge.ErrorCode)"
        }
    }
    throw ("Cannot read the primary mailbox id of $UserPrincipalName through Graph ($why; GET /users/{id}/settings/exchange needs the delegated permission User.Read.All). " +
        "Give it with -PrimaryMailboxId MBX:<mailbox guid>@$TenantId. The guids come from Exchange Online PowerShell: Get-MailboxLocation -User $UserPrincipalName " +
        "(MailboxLocationType Primary, MainArchive, AuxArchive).")
}

function Test-MbtMailboxAddress {
    <# $true for an SMTP address / UPN (user@domain), $false for a physical id (MBX:...) or a bare guid. #>
    [CmdletBinding()]
    param([AllowNull()][AllowEmptyString()][string]$Value)
    if (-not $Value) { return $false }
    $v = $Value.Trim()
    return ($v -notmatch '^MBX:') -and ($v -match '^[^@\s]+@[^@\s]+$')
}

function Get-MbtAppOnlyAccessHint {
    <# The 403 of an app-only session: the application permission missing in the token, else the Exchange RBAC for Applications scope. #>
    param([ValidateSet('Browse', 'Export', 'Import', 'Resolve')][string]$Mode = 'Browse', [string[]]$Scopes = @())
    # each need: the delegated-style names that satisfy it (the token's roles without .All), and the application permission to grant
    $needs = switch ($Mode) {
        'Import' { @(@{ Any = @('MailboxFolder.ReadWrite'); Grant = 'MailboxFolder.ReadWrite.All' }, @{ Any = @('MailboxItem.ImportExport'); Grant = 'MailboxItem.ImportExport.All' }) }
        'Export' { @(@{ Any = @('MailboxFolder.Read', 'MailboxFolder.ReadWrite'); Grant = 'MailboxFolder.Read.All' }, @{ Any = @('MailboxItem.Read', 'MailboxItem.ReadWrite', 'MailboxItem.Export', 'MailboxItem.ImportExport'); Grant = 'MailboxItem.Read.All' }) }
        default { @(@{ Any = @('MailboxFolder.Read', 'MailboxFolder.ReadWrite'); Grant = 'MailboxFolder.Read.All' }, @{ Any = @('MailboxItem.Read', 'MailboxItem.ReadWrite', 'MailboxItem.Export', 'MailboxItem.ImportExport'); Grant = 'MailboxItem.Read.All' }) }
    }
    $missing = @($needs | Where-Object { $n = $_; -not @($n.Any | Where-Object { $_ -in $Scopes }).Count } | ForEach-Object { $_.Grant })
    $text = ' Access denied to the application on this mailbox.'
    if ($missing -and $Scopes) { $text += ' Application permission missing in the token: ' + ($missing -join ', ') + ' (an administrator grants it and consents to it for the application, or assigns the matching Exchange application role with RBAC for Applications; then sign in again).' }
    elseif ($Scopes) { $text += ' The application permissions needed are in the token, so the refusal comes from Exchange: the mailbox may be outside the management scope of the application (RBAC for Applications).' }
    $text += ' Check with Exchange Online PowerShell: Test-ServicePrincipalAuthorization -Identity <application id> -Resource <mailbox> (InScope must be True). A change of scope takes 30 minutes to 2 hours to apply.'
    return $text
}

function Get-MbtAccessHint {
    <#
    .SYNOPSIS
        What to check when the service answers 403 (ErrorAccessDenied) for a mailbox that is not one of the signed-in user: the text to append
        to the error, or '' when the message is not an access denial.
    .DESCRIPTION
        Delegated access to another user's mailbox (a shared mailbox, a colleague's) is the sum of two things: the Exchange right of the
        signed-in user on that mailbox (FullAccess; it can take up to 60 minutes to apply) and the delegated permissions in the token.
        Mail API: Mail.Read only opens the signed-in user's own mailbox, another one needs Mail.Read.Shared (browse then falls back to the
        import-export listing). Import-export API: MailboxFolder.Read and MailboxItem.Read to browse and export, MailboxFolder.ReadWrite
        to create folders, MailboxItem.ImportExport to import. Observed in the lab (FullAccess granted, the five permissions in the token):
        reads answer, creating a folder and importItem are refused with 403.
    #>
    [CmdletBinding()]
    param([AllowNull()][string]$Message, [string[]]$Scopes = @(), [ValidateSet('Browse', 'Export', 'Import', 'Resolve')][string]$Mode = 'Browse', [switch]$OtherMailbox, [switch]$AppOnly)
    if (-not $Message -or $Message -notmatch '403|ErrorAccessDenied|Access is denied') { return '' }
    if ($AppOnly) { return (Get-MbtAppOnlyAccessHint -Mode $Mode -Scopes $Scopes) }
    if (-not $OtherMailbox) { return '' }
    # each need: the permission names that satisfy it
    $needs = switch ($Mode) {
        'Import' { @(, @('MailboxFolder.ReadWrite')) + @(, @('MailboxItem.ImportExport')) }
        'Export' { @(, @('MailboxFolder.Read', 'MailboxFolder.ReadWrite')) + @(, @('MailboxItem.Read', 'MailboxItem.ImportExport')) }
        default { @(, @('MailboxFolder.Read', 'MailboxFolder.ReadWrite')) + @(, @('MailboxItem.Read', 'MailboxItem.ImportExport')) }
    }
    $missing = @($needs | Where-Object { $group = $_; -not @($group | Where-Object { $_ -in $Scopes }).Count } | ForEach-Object { $_[0] })
    $text = ' Access denied on a mailbox that is not the signed-in user''s.'
    if ($Message -match "requires the following permissions: '(?<p>[^']+)'") {
        $text += " The service says this delegated call needs: $($Matches.p). The .Shared variants do not exist as delegated permissions on Microsoft Graph (checked in the lab tenant), so delegated writes to another user's mailbox through the import-export API are not possible:" +
            ' use application permissions (.All) with a certificate for a shared mailbox (-CertificateThumbprint, docs\MailboxTransfer-Guide.md, chapter 4.6).'
    }
    elseif ($missing -and $Scopes) { $text += ' Missing in the token: ' + ($missing -join ', ') + ' (an administrator grants it to the application, tools\Grant-LabConsent.ps1; then sign in again).' }
    elseif ($Scopes) { $text += ' The permissions needed are in the token, so the refusal comes from Exchange.' }
    $text += ' The signed-in user needs the Exchange right FullAccess on that mailbox (Add-MailboxPermission -Identity <mailbox> -User <user> -AccessRights FullAccess; it can take up to 60 minutes to apply).'
    if ($Mode -eq 'Browse') { $text += ' The Mail API opens another mailbox only with Mail.Read.Shared (browse falls back to the import-export listing).' }
    return $text
}

function Resolve-MbtMailboxAddress {
    <#
    .SYNOPSIS
        The physical id MBX:<guid>@<tenant> of the mailbox of an SMTP address / UPN (a shared mailbox, another user's mailbox).
    .DESCRIPTION
        1. GET /users/{address}/settings/exchange (primaryMailboxId; needs User.Read.All, which the lab consent does not carry).
        2. Else the import-export API, which accepts the address in place of the id: the first folder of GET /admin/exchange/mailboxes/{address}/folders
           carries parentMailboxUrl, which holds the physical id. Needs the Exchange right of the signed-in user on that mailbox (FullAccess).
        The error says which call failed and how, and what is missing.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Address, [Parameter(Mandatory)][string]$TenantId)
    $s = Get-MbtSession
    $why = @()
    try { return (Get-MbtPrimaryMailboxId -UserPrincipalName $Address -TenantId $TenantId) }
    catch { $why += "Graph settings/exchange: $($_.Exception.Message -replace '\s+', ' ' -replace '^(.{0,160}).*$', '$1')" }
    try {
        $doc = $s.Graph.GetJson($s.Graph.Url([MailboxTransfer.MailboxApi]::MailboxPath($Address) + '/folders?$top=1'))
        try {
            $value = [Text.Json.JsonElement]::new()
            if ($doc.RootElement.TryGetProperty('value', [ref]$value) -and $value.ValueKind -eq 'Array') {
                foreach ($f in $value.EnumerateArray()) {
                    $u = [Text.Json.JsonElement]::new()
                    if ($f.TryGetProperty('parentMailboxUrl', [ref]$u) -and $u.ValueKind -eq 'String') {
                        $id = [MailboxTransfer.MailboxApi]::MailboxIdOfUrl($u.GetString())
                        if ($id) { return $id }
                    }
                }
            }
            $why += 'import-export API: the folder listing carries no parentMailboxUrl'
        } finally { $doc.Dispose() }
    } catch {
        $ge = Get-MbtGraphException $_
        if (-not $ge) { throw }
        $why += "import-export API: HTTP $($ge.StatusCode) $($ge.ErrorCode)"
    }
    throw ("Cannot find the mailbox id of $Address (" + ($why -join '; ') + '). Give it with -MailboxId MBX:<mailbox guid>@' + $TenantId +
        ' (Exchange Online PowerShell: Get-MailboxLocation -User ' + $Address + '), or ' + $(if (Test-MbtAppOnly $s) { 'grant the application User.Read.All (application permission) and put the mailbox in its Exchange RBAC for Applications scope.' } else { 'grant the signed-in user FullAccess on that mailbox (up to 60 minutes to apply), or User.Read.All.' }) +
        (Get-MbtAccessHint -Message ($why -join ' ') -Mode Resolve -OtherMailbox -AppOnly:(Test-MbtAppOnly $s) -Scopes $(if ($s.Token) { $s.Token.Scopes })))
}

function ConvertTo-MbtLocationKind {
    <# 'Primary', 'Archive' (MainArchive) or 'AuxArchive' from a MailboxLocationType of Exchange Online (Primary, MainArchive, AuxArchive) or from our own kind names; $null for anything else. #>
    param([AllowNull()][string]$Type)
    if (-not $Type) { return $null }
    switch -Regex ($Type.Trim()) {
        '^Primary$' { return 'Primary' }
        '^(MainArchive|Archive)$' { return 'Archive' }
        '^(AuxArchive|AuxiliaryArchive|Aux)$' { return 'AuxArchive' }
    }
    return $null
}

function Read-MbtMailboxLocationFile {
    <#
    .SYNOPSIS
        Reads what the operator got from Exchange Online PowerShell, without requiring that module here:
            Get-MailboxLocation -User <upn> | Select-Object MailboxGuid, MailboxLocationType | Export-Csv locations.csv   (or ConvertTo-Json, or Format-Table > txt)
        Returns one object per mailbox (Kind Primary / Archive / AuxArchive, MailboxId MBX:<guid>@<tenant>).
    .DESCRIPTION
        Tolerant on purpose: a JSON file (an array of objects with MailboxGuid and MailboxLocationType) or any text (CSV, Format-Table output):
        every line that holds a guid and one of the words Primary, MainArchive, AuxArchive is a location. Other types
        (Aggregated, ComponentShared...) are ignored.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$TenantId)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Locations file not found: $Path" }
    $text = (Get-Content -LiteralPath $Path -Raw).TrimStart([char]0xFEFF)
    $pairs = [System.Collections.Generic.List[object]]::new()
    $guidRx = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'
    $trim = $text.TrimStart()
    if ($trim.StartsWith('[') -or $trim.StartsWith('{')) {
        foreach ($o in @($trim | ConvertFrom-Json)) {
            $g = $o.PSObject.Properties | Where-Object { $_.Name -in 'MailboxGuid', 'ExchangeGuid' } | Select-Object -First 1
            $t = $o.PSObject.Properties | Where-Object { $_.Name -in 'MailboxLocationType', 'Type', 'Kind' } | Select-Object -First 1
            if ($g -and $t) { $pairs.Add([pscustomobject]@{ Guid = [string]$g.Value; Type = [string]$t.Value }) }
        }
    } else {
        foreach ($line in $text -split "`r?`n") {
            if ($line -match "(?<g>$guidRx)") {
                $guid = $Matches.g
                if ($line -match '\b(?<t>MainArchive|AuxArchive|Primary)\b') { $pairs.Add([pscustomobject]@{ Guid = $guid; Type = $Matches.t }) }
            }
        }
    }
    foreach ($pair in $pairs) {
        $kind = ConvertTo-MbtLocationKind $pair.Type
        if (-not $kind) { continue }
        [pscustomobject]@{ Kind = $kind; MailboxId = (ConvertTo-MbtMailboxId -Value $pair.Guid -TenantId $TenantId); Source = 'OperatorFile'; Detail = "MailboxLocationType $($pair.Type) in $(Split-Path $Path -Leaf)" }
    }
}

function Get-MbtGraphMailboxLocation {
    <#
    .SYNOPSIS
        What Microsoft Graph says about the mailboxes of a user: GET /beta/users/{upn}/settings/exchange (primaryMailboxId and
        inPlaceArchiveMailboxId), else GET /v1.0/... (primaryMailboxId only). Needs a session and the delegated permission User.Read.All.
        No Graph API lists the AUXILIARY archives. Returns the locations found and, in Lookup, what each call answered.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$UserPrincipalName, [Parameter(Mandatory)][string]$TenantId)
    $s = Get-MbtSession
    $base = $s.Graph.Options.BaseUrl
    $beta = $base -replace '/v1\.0$', '/beta'
    $paths = Get-MbtExchangeSettingsPath -UserPrincipalName $UserPrincipalName -Session $s
    $found = @(); $lookup = @()
    foreach ($root in @($beta, $base) | Select-Object -Unique) {
        foreach ($path in $paths) {
            try {
                $doc = $s.Graph.GetJson($root + $path)
                try {
                    $note = "$root$path ok"
                    foreach ($pair in @(@('primaryMailboxId', 'Primary'), @('inPlaceArchiveMailboxId', 'Archive'))) {
                        $v = [Text.Json.JsonElement]::new()
                        if ($doc.RootElement.TryGetProperty($pair[0], [ref]$v) -and $v.ValueKind -eq 'String' -and $v.GetString()) {
                            $id = try { ConvertTo-MbtMailboxId -Value $v.GetString() -TenantId $TenantId } catch { $v.GetString() }
                            $found += [pscustomobject]@{ Kind = $pair[1]; MailboxId = $id; Source = 'Graph'; Detail = "$($pair[0]) of GET $($root -replace '^.*/', '')$path" }
                        }
                    }
                    $lookup += $note
                } finally { $doc.Dispose() }
            } catch {
                $ge = Get-MbtGraphException $_
                if (-not $ge) { throw }
                $lookup += "$root$path HTTP $($ge.StatusCode) $($ge.ErrorCode)"
                continue
            }
            if ($found) { break }
        }
        if ($found) { break }
    }
    [pscustomobject]@{ Locations = $found; Lookup = $lookup }
}

function Get-MbtMailboxLocation {
    <#
    .SYNOPSIS
        Lists the locations of the mailboxes of a user: Primary, MainArchive (Kind Archive) and every AuxArchive (auto-expanding archive), with
        their MBX:<guid>@<tenant> ids, and how each one was learned. Several sources are merged, in this order of preference:
          1. Microsoft Graph, when a session is open and the token has User.Read.All: GET /beta/users/{id}/settings/exchange gives primaryMailboxId and
             inPlaceArchiveMailboxId (v1.0 gives the primary only). No Graph API lists the auxiliary archives.
          2. An operator file or objects (-LocationsFile, -InputObject): the output of Exchange Online
             Get-MailboxLocation -User <upn> | Select MailboxGuid, MailboxLocationType. Exchange Online PowerShell is not required by this tool.
          3. Parameters (-PrimaryMailboxId, -ArchiveMailboxId, -AuxArchiveMailboxId).
          4. The store (-Store): the locations recorded by earlier runs, and above all the auxiliary archives LEARNED from the redirects of
             the service (308 Location, ErrorArchiveFolderMovedPermanently, 409 "target session is expected in mailbox"), which always works.
        What is given (operator, parameters, Graph) beats what was guessed from a redirect. With a writable -Store the new locations are recorded.
    #>
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName,
        [string]$TenantId,
        $Store,
        [string]$LocationsFile,
        [Parameter(ValueFromPipeline)]$InputObject,
        [string]$PrimaryMailboxId,
        [string]$ArchiveMailboxId,
        [string[]]$AuxArchiveMailboxId,
        [switch]$SkipGraph,
        [switch]$NoSave
    )
    begin { $piped = [System.Collections.Generic.List[object]]::new() }
    process { if ($null -ne $InputObject) { $piped.Add($InputObject) } }
    end {
        $found = [System.Collections.Generic.List[object]]::new()
        if (-not $TenantId -and $script:Session -and $script:Session.TenantId) { $TenantId = $script:Session.TenantId }
        if ($Store) { Initialize-MbtEngine }
        # 1. Graph
        if (-not $SkipGraph -and $script:Session -and $UserPrincipalName) {
            $g = Get-MbtGraphMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $(if ($TenantId) { $TenantId } else { $script:Session.TenantId })
            foreach ($l in $g.Locations) { $found.Add($l) }
            foreach ($n in $g.Lookup) { Write-Verbose "Graph: $n" }
            if (-not $g.Locations) { Write-Verbose 'Graph gave no mailbox id (User.Read.All is needed); the other sources are used. No Graph API lists the auxiliary archives.' }
        }
        # 2. operator file and objects
        if ($LocationsFile) {
            if (-not $TenantId) { throw '-TenantId is needed to complete the mailbox guids of -LocationsFile.' }
            foreach ($l in Read-MbtMailboxLocationFile -Path $LocationsFile -TenantId $TenantId) { $found.Add($l) }
        }
        foreach ($o in $piped) {
            $guid = if ($o.PSObject.Properties['MailboxGuid']) { $o.MailboxGuid } else { $null }
            $type = if ($o.PSObject.Properties['MailboxLocationType']) { [string]$o.MailboxLocationType } else { '' }
            if (-not $guid) { continue }
            $kind = ConvertTo-MbtLocationKind $type
            if (-not $kind) { continue }
            if (-not $TenantId) { throw '-TenantId is needed to complete the mailbox guids of the objects given.' }
            $found.Add([pscustomobject]@{ Kind = $kind; MailboxId = (ConvertTo-MbtMailboxId -Value ([string]$guid) -TenantId $TenantId); Source = 'OperatorInput'; Detail = "MailboxLocationType $type" })
        }
        # 3. parameters
        foreach ($pair in @(@('Primary', $PrimaryMailboxId), @('Archive', $ArchiveMailboxId))) {
            if ($pair[1]) { $found.Add([pscustomobject]@{ Kind = $pair[0]; MailboxId = (ConvertTo-MbtMailboxId -Value $pair[1] -TenantId $TenantId); Source = 'Parameter'; Detail = "-$($pair[0])MailboxId or configuration" }) }
        }
        foreach ($a in @($AuxArchiveMailboxId | Where-Object { $_ })) { $found.Add([pscustomobject]@{ Kind = 'AuxArchive'; MailboxId = (ConvertTo-MbtMailboxId -Value $a -TenantId $TenantId); Source = 'Parameter'; Detail = '-AuxArchiveMailboxId' }) }
        # given sources are recorded in the store (before the store's own rows are read back)
        $now = [DateTime]::UtcNow
        if ($Store -and -not $Store.IsReadOnly -and -not $NoSave) {
            foreach ($l in $found) {
                [void]$Store.RecordLocation([MailboxTransfer.StoreMailboxLocation]@{ MailboxId = $l.MailboxId; Kind = $l.Kind; Upn = $UserPrincipalName; Source = $l.Source; Detail = $l.Detail; Rank = [MailboxTransfer.StoreMailboxLocation]::RankGiven; DiscoveredUtc = $now })
            }
        }
        $result = [System.Collections.Generic.List[object]]::new()
        $byId = @{}
        $order = @{ Primary = 0; Archive = 1; AuxArchive = 2 }
        foreach ($l in $found) {
            $key = $l.MailboxId.ToLowerInvariant()
            if ($byId.ContainsKey($key)) { if ($byId[$key].Kind -ne $l.Kind) { Write-Warning "Mailbox $($l.MailboxId) is $($byId[$key].Kind) per $($byId[$key].Source) but $($l.Kind) per $($l.Source): the first is kept." }; continue }
            $obj = [pscustomobject]@{ Kind = $l.Kind; MailboxId = $l.MailboxId; Source = $l.Source; Detail = $l.Detail; DiscoveredUtc = $now; Learned = $false }
            $byId[$key] = $obj; $result.Add($obj)
        }
        # 4. the store
        if ($Store) {
            foreach ($sl in $Store.GetLocations()) {
                $key = $sl.MailboxId.ToLowerInvariant()
                if ($byId.ContainsKey($key)) { continue }
                $obj = [pscustomobject]@{ Kind = $sl.Kind.ToString(); MailboxId = $sl.MailboxId; Source = $sl.Source; Detail = $sl.Detail; DiscoveredUtc = $sl.DiscoveredUtc; Learned = ($sl.Rank -lt [MailboxTransfer.StoreMailboxLocation]::RankGiven) }
                $byId[$key] = $obj; $result.Add($obj)
            }
        }
        $result | Sort-Object { $order[$_.Kind] }, DiscoveredUtc, MailboxId
    }
}

function Get-MbtGraphException {
    <# The GraphException inside an error record (PowerShell wraps the exceptions of .NET methods), or $null. #>
    param([Parameter(Mandatory)]$ErrorRecord)
    $e = $ErrorRecord.Exception
    while ($e -and $e -isnot [MailboxTransfer.GraphException]) { $e = $e.InnerException }
    return $e
}

function Stop-MbtImportFailure {
    <#
    .SYNOPSIS
        Rethrows the error of an import. The answer 409 "Invalid import session. The target session is expected in mailbox
        MBX:..." (the folder lives in an auxiliary archive of an auto-expanding archive) gets an explanation; any other error
        is rethrown unchanged.
    #>
    param([Parameter(Mandatory)]$ErrorRecord)
    $ge = Get-MbtGraphException $ErrorRecord
    if ($ge -and $ge.StatusCode -eq 409 -and $ge.Message -match 'expected in mailbox\s+(?<id>MBX:[0-9A-Fa-f-]+@[0-9A-Fa-f-]+)') {
        throw ("$($ge.Message)`nThe target folder lives in an auxiliary archive (auto-expanding archive): the import session must be created for $($Matches.id) " +
            "(MailboxLocationType AuxArchive in Get-MailboxLocation), with the id of the folder as that mailbox lists it. The single-item import does not follow it " +
            "(the whole-mailbox import, Import-MbtMailbox / -Mode Import, does): nothing was imported.")
    }
    throw $ErrorRecord
}

function Get-MbtMailboxTree {
    <#
    .SYNOPSIS
        Complete folder tree of a mailbox (primary or archive) with item counts. -Sizes adds the size of each folder by
        adding up the sizes of its items (one listing per folder: slow on large mailboxes).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$MailboxId, [switch]$Sizes, [int]$PageSize = -1)
    $s = Get-MbtSession
    if ($PageSize -lt 0) { $PageSize = [int]$s.Configuration.Graph.ItemPageSize }
    $tree = @($s.Api.GetFolderTree($MailboxId))
    if ($Sizes) {
        $todo = @($tree | Where-Object { $_.TotalItemCount -gt 0 })
        $n = 0
        foreach ($folder in $todo) {
            $n++
            Write-Progress -Activity 'Measuring folders' -Status $folder.Path -PercentComplete (100 * $n / $todo.Count)
            [long]$sum = 0
            foreach ($item in $s.Api.EnumerateItems($MailboxId, $folder.Id, $PageSize)) { $sum += $item.Size }
            $folder.TotalSize = $sum
        }
        foreach ($folder in ($tree | Where-Object { $_.TotalItemCount -eq 0 })) { $folder.TotalSize = 0 }
        Write-Progress -Activity 'Measuring folders' -Completed
    }
    return $tree
}

function Resolve-MbtFolder {
    <#
    .SYNOPSIS
        Finds a folder of a tree by id, by full path, or by the end of its path ("Attachments", "Lab/Attachments").
        Several matches is an error that lists them: the script never guesses. No match is an error too, or $null
        with -NullIfMissing.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Tree, [Parameter(Mandatory)][string]$Name, [switch]$NullIfMissing)
    $byId = @($Tree | Where-Object { $_.Id -ceq $Name })
    if ($byId.Count -eq 1) { return $byId[0] }
    $wanted = $Name.Trim('/')
    $candidates = @($Tree | Where-Object { $_.Path -ieq $wanted -or $_.Path.EndsWith('/' + $wanted, [StringComparison]::OrdinalIgnoreCase) })
    # The shortest path wins only if it is the only exact one: "A/B" must not silently stand for "X/A/B".
    $exact = @($candidates | Where-Object { $_.Path -ieq $wanted })
    if ($exact.Count -eq 1) { return $exact[0] }
    if ($candidates.Count -eq 1) { return $candidates[0] }
    if ($candidates.Count -eq 0) {
        if ($NullIfMissing) { return $null }
        throw "Folder '$Name' not found in this mailbox."
    }
    throw "Folder '$Name' is ambiguous, $($candidates.Count) matches:`n  " + (($candidates | ForEach-Object { "$($_.Path)  [$($_.Id)]" }) -join "`n  ")
}

function Get-MbtFolderItems {
    <# Items of a folder (id, type, size, dates, change key). The listing carries no subject. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$MailboxId, [Parameter(Mandatory)][string]$FolderId, [int]$PageSize = -1, [int]$First = 0)
    $s = Get-MbtSession
    if ($PageSize -lt 0) { $PageSize = [int]$s.Configuration.Graph.ItemPageSize }
    $items = [System.Collections.Generic.List[object]]::new()
    foreach ($item in $s.Api.EnumerateItems($MailboxId, $FolderId, $PageSize)) {
        $items.Add($item)
        if ($First -gt 0 -and $items.Count -ge $First) { break }
    }
    return $items.ToArray()
}

function New-MbtFolderPath {
    <#
    .SYNOPSIS
        Finds the folder of the path in the tree, creating the missing ones (create only: nothing is renamed, moved or deleted).
        Returns the last folder (FolderInfo); with -WhatIf, $null as soon as a folder would have to be created.
    .DESCRIPTION
        POST /admin/exchange/mailboxes/{id}/folders (top level) or .../folders/{parentId}/childFolders, delegated permission
        MailboxFolder.ReadWrite, body displayName and type (IPF.Note by default).
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$MailboxId,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Tree,
        [Parameter(Mandatory)][string]$Path,
        [string]$Type = 'IPF.Note'
    )
    $s = Get-MbtSession
    $segments = @($Path.Trim('/') -split '/' | Where-Object { $_ })
    if (-not $segments) { throw 'Folder path is empty.' }
    $parent = $null
    $prefix = ''
    foreach ($name in $segments) {
        $prefix = if ($prefix) { "$prefix/$name" } else { $name }
        $existing = @($Tree | Where-Object { $_.Path -ieq $prefix })
        if ($existing.Count -gt 1) { throw "Folder '$prefix' is ambiguous, $($existing.Count) folders have this path." }
        if ($existing.Count -eq 1) { $parent = $existing[0]; continue }
        if (-not $PSCmdlet.ShouldProcess($MailboxId, "Create folder '$prefix'")) { return $null }
        $parent = $s.Api.CreateFolder($MailboxId, $name, $(if ($parent) { $parent.Id }), $Type, $(if ($parent) { $parent.Path }))
    }
    return $parent
}

#endregion

#region 5b. Browse (live view of a mailbox before the export) --------------------------------

function New-MbtBrowseLocation {
    <#
    .SYNOPSIS
        A mailbox to browse: Primary, Archive (main archive) or AuxArchive, with its MBX:<guid>@<tenant> id. Primary and Archive are read
        with the Mail API (full headers), an auxiliary archive with the import-export API (the Mail API does not reach it).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][ValidateSet('Primary', 'Archive', 'AuxArchive')][string]$Kind, [Parameter(Mandatory)][string]$MailboxId, [string]$Source)
    Initialize-MbtEngine
    return [MailboxTransfer.MailboxBrowser]::Location($Kind, $MailboxId, $Source)
}

function Get-MbtBrowseLocation {
    <#
    .SYNOPSIS
        The locations of the user as the browse view takes them: Get-MbtMailboxLocation (Graph, operator file, parameters, the store) converted to
        browse locations. Same parameters as Get-MbtMailboxLocation; nothing is recorded in the store.
    #>
    [CmdletBinding()]
    param(
        [string]$UserPrincipalName, [string]$TenantId, $Store, [string]$LocationsFile, [string]$PrimaryMailboxId, [string]$ArchiveMailboxId,
        [string[]]$AuxArchiveMailboxId, [switch]$SkipGraph
    )
    $p = @{ NoSave = $true }
    foreach ($k in 'UserPrincipalName', 'TenantId', 'Store', 'LocationsFile', 'PrimaryMailboxId', 'ArchiveMailboxId', 'AuxArchiveMailboxId') { if ($PSBoundParameters.ContainsKey($k)) { $p[$k] = $PSBoundParameters[$k] } }
    if ($SkipGraph) { $p['SkipGraph'] = $true }
    foreach ($l in @(Get-MbtMailboxLocation @p)) { New-MbtBrowseLocation -Kind $l.Kind -MailboxId $l.MailboxId -Source $l.Source }
}

function Get-MbtBrowser {
    <# The browser object of the session (C#: MailboxTransfer.MailboxBrowser). -RecoverableItems adds Recoverable Items to the tree (-RecoverableItemsFolders: only these folders of it, else all six); -NoNonMailFolders leaves out Calendar, Contacts... #>
    [CmdletBinding()]
    param([switch]$RecoverableItems, [string[]]$RecoverableItemsFolders, [switch]$NoNonMailFolders)
    $s = Get-MbtSession
    Initialize-MbtEngine
    $b = [MailboxTransfer.MailboxBrowser]::new($s.Api)
    $b.IncludeRecoverableItems = [bool]$RecoverableItems
    if ($RecoverableItemsFolders) { $b.RecoverableItemsFolders = [MailboxTransfer.RecoverableItems]::Names($RecoverableItemsFolders) }
    $b.IncludeNonMailFolders = -not $NoNonMailFolders
    return $b
}

function Get-MbtBrowseFolder {
    <#
    .SYNOPSIS
        The folders of a mailbox with item counts and sizes (Mail API for Primary and Archive, import-export API for an auxiliary archive or when the
        Mail API refuses). Without -Lazy the whole tree (one request per 200 folders); with -Lazy the children of -Parent (the first level without it).
        -Path / -Folder: only the folder with this id or path. -AsTree returns the BrowseTree (folders and notes) instead of the folders.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Location,
        [Alias('Folder')][string]$Path,
        [switch]$Lazy,
        $Parent,
        [switch]$RecoverableItems,
        [string[]]$RecoverableItemsFolders,
        [switch]$NoNonMailFolders,
        [switch]$AsTree,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )
    $b = Get-MbtBrowser -RecoverableItems:$RecoverableItems -RecoverableItemsFolders $RecoverableItemsFolders -NoNonMailFolders:$NoNonMailFolders
    if ($Lazy) { return $b.ListFolders($Location, $Parent, $CancellationToken) }
    $tree = $b.GetFolderTree($Location, $CancellationToken, $null)
    if ($AsTree) { return $tree }
    foreach ($n in $tree.Notes) { Write-Warning $n }
    if ($Path) {
        $f = $tree.Find($Path)
        if (-not $f) { throw "Folder '$Path' not found in this mailbox." }
        return $f
    }
    return $tree.Folders.ToArray()
}

function Find-MbtBrowseFolder {
    <#
    .SYNOPSIS
        One folder by its path (Inbox/Reports), walking down the tree one level at a time (one request a level: seconds on an archive of thousands of
        folders where the whole tree takes minutes). Not found that way (the end of a path, an id): the whole tree is read and searched. Two folders of
        the same name at one level is an error that lists their ids.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Location, [Parameter(Mandatory)][string]$Path, [switch]$RecoverableItems, [string[]]$RecoverableItemsFolders, [switch]$NoNonMailFolders, [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None)
    $b = Get-MbtBrowser -RecoverableItems:$RecoverableItems -RecoverableItemsFolders $RecoverableItemsFolders -NoNonMailFolders:$NoNonMailFolders
    $current = $null
    $found = $true
    foreach ($segment in @($Path.Trim('/') -split '/' | Where-Object { $_ })) {
        $hit = @($b.ListFolders($Location, $current, $CancellationToken) | Where-Object { $_.DisplayName -ieq $segment })
        if ($hit.Count -gt 1) { throw "Folder '$Path' is ambiguous at '$segment', $($hit.Count) folders have this name:`n  " + (($hit | ForEach-Object { "$($_.Path)  [$($_.Id)]" }) -join "`n  ") + "`nGive the id with -Folder." }
        if ($hit.Count -eq 0) { $found = $false; break }
        $current = $hit[0]
    }
    if ($found -and $current) { return $current }
    $tree = $b.GetFolderTree($Location, $CancellationToken, $null)
    $folder = $tree.Find($Path)
    if (-not $folder) { throw "Folder '$Path' not found in this mailbox." }
    return $folder
}

function Get-MbtBrowseItem {
    <#
    .SYNOPSIS
        The rows of a folder: subject, sender, To/Cc counts, received and sent dates, size, attachments, importance, sensitivity, item class,
        categories, flag (Mail API), or id, class, size and dates only (import-export API). Paged and lazy: -Top stops after n rows, -Page returns
        one BrowsePage (with NextToken) to continue with -Token.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Location,
        [Parameter(Mandatory)]$Folder,
        [int]$Top = 0,
        [ValidateRange(1, 1000)][int]$PageSize = 50,
        [ValidateSet('ReceivedDescending', 'ReceivedAscending', 'SentDescending', 'SentAscending', 'SubjectAscending', 'SubjectDescending')][string]$Sort = 'ReceivedDescending',
        [switch]$NoRecipients,
        [string]$Token,
        [switch]$Page,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )
    $b = Get-MbtBrowser
    if ($Folder -is [string]) { $Folder = Find-MbtBrowseFolder -Location $Location -Path $Folder -CancellationToken $CancellationToken }
    $q = [MailboxTransfer.BrowseItemsQuery]::new()
    $q.PageSize = if ($Top -gt 0 -and $Top -lt $PageSize) { $Top } else { $PageSize }
    $q.Sort = [MailboxTransfer.BrowseSort]$Sort
    $q.Recipients = -not $NoRecipients
    if ($Token) { $q.Token = $Token }
    if ($Page) { return $b.ListItems($Location, $Folder, $q, $CancellationToken, $null) }
    $n = 0
    $first = $true
    while ($true) {
        $pg = $b.ListItems($Location, $Folder, $q, $CancellationToken, $null)
        if ($first) { foreach ($note in $pg.Notes) { Write-Warning $note }; $first = $false }
        foreach ($i in $pg.Items) {
            $i
            $n++
            if ($Top -gt 0 -and $n -ge $Top) { return }
        }
        if (-not $pg.HasMore) { return }
        $q.Token = $pg.NextToken
    }
}

function Get-MbtBrowsePreview {
    <#
    .SYNOPSIS
        Header fields, an excerpt of the text body, the attachments (with the attached items read one level down) of one item. -Item is a row of
        Get-MbtBrowseItem; -ItemId alone reads a message of the Mail API by id (Primary / Archive only).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Location,
        [Parameter(Mandatory, ParameterSetName = 'Row', ValueFromPipeline)]$Item,
        [Parameter(Mandatory, ParameterSetName = 'Id')][string]$ItemId,
        [int]$BodyChars = 2000,
        [switch]$NoAttachments,
        [ValidateRange(0, 3)][int]$NestedDepth = 1,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )
    process {
        $b = Get-MbtBrowser
        if ($PSCmdlet.ParameterSetName -eq 'Id') {
            $Item = [MailboxTransfer.BrowseItem]@{ Id = $ItemId; Source = [MailboxTransfer.MailboxBrowser]::MailApi }
        }
        $o = [MailboxTransfer.BrowsePreviewOptions]::new()
        $o.BodyChars = $BodyChars; $o.Attachments = -not $NoAttachments; $o.NestedDepth = $NestedDepth
        $b.GetItemPreview($Location, $Item, $o, $CancellationToken)
    }
}

function New-MbtBrowseSelection {
    <# A selection of one mailbox (what the user ticks in the browse view): add folders and items with Add-MbtBrowseSelection, save it with Save-MbtBrowseSelection. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Location)
    Initialize-MbtEngine
    return [MailboxTransfer.BrowseSelection]::For($Location)
}

function Add-MbtBrowseSelection {
    <# Ticks a folder (-IncludeSubfolders: and the folders below it), or the given items of a folder. Returns the selection. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Selection,
        [Parameter(Mandatory)]$Folder,
        [switch]$IncludeSubfolders,
        $Item
    )
    if ($Item) { $null = $Selection.AddItems($Folder, [MailboxTransfer.BrowseItem[]]@($Item)) }
    else { $null = $Selection.AddFolder($Folder, [bool]$IncludeSubfolders) }
    return $Selection
}

function Save-MbtBrowseSelection {
    <# Writes the selection as JSON (the file -Mode Export -SelectionFile reads). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Selection, [Parameter(Mandatory)][string]$Path)
    Initialize-MbtEngine
    $full = if ([IO.Path]::IsPathRooted($Path)) { $Path } else { Join-Path (Get-Location).Path $Path }
    [MailboxTransfer.FileNaming]::Require($full, $true, 'Selection file')
    $dir = Split-Path $full -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Force -Path $dir }
    $Selection.Save($full)
    return $full
}

function Read-MbtBrowseSelection {
    <# Reads a selection file written by Save-MbtBrowseSelection (or by the console). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    Initialize-MbtEngine
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Selection file not found: $Path" }
    $full = (Resolve-Path -LiteralPath $Path).Path
    return [MailboxTransfer.BrowseSelection]::Load($full)
}

function Resolve-MbtBrowsePlan {
    <#
    .SYNOPSIS
        Turns a selection into an export plan: the folder ids and item ids of the import-export API (read only: one GET per folder, one listing per
        folder with ticked items). Each entry says how it was matched (SameId, ByPath, Unresolved); Problems lists what cannot be exported.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Selection, [switch]$NoVerify,
        # the other Archive / AuxArchive locations of the user (New-MbtBrowseLocation): a folder, or items, that moved there since the selection was made are found by their folder entry key
        $AlsoSearch,
        # the folders of Recoverable Items that count when the plan takes the root with its subfolders (default set: Deletions, Purges, Versions, DiscoveryHolds)
        [string[]]$RecoverableItemsFolders,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )
    Initialize-MbtEngine
    $b = Get-MbtBrowser -RecoverableItemsFolders ([MailboxTransfer.RecoverableItems]::Chosen([string[]]@($RecoverableItemsFolders | Where-Object { $_ })))
    $others = [System.Collections.Generic.List[MailboxTransfer.BrowseLocation]]::new()
    foreach ($l in @($AlsoSearch | Where-Object { $_ })) { $others.Add($l) }
    return $b.ResolvePlan($Selection, -not $NoVerify, $CancellationToken, $null, $others)
}

function Add-MbtBrowseMoveInfo {
    <#
    .SYNOPSIS
        Finds the folders that exist in the main archive and in an auxiliary archive (same folder entry key in the id; the path only for an id that is not in the
        Exchange format) and writes on each of them MoveState (None, BeingMoved, Moved, BothSides) and MoveText ("being moved to aux archive 1 (1a2b3c06)").
        Takes the trees of Get-MbtBrowseFolder -AsTree (Archive and AuxArchive ones count); returns the groups. Read only.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][object[]]$Tree)
    Initialize-MbtEngine
    return [MailboxTransfer.FolderCorrelation]::Annotate([MailboxTransfer.BrowseTree[]]$Tree)
}

function Format-MbtBrowseDate {
    param($Value)
    if ($null -eq $Value) { return '-' }
    return ([DateTimeOffset]$Value).UtcDateTime.ToString('yyyy-MM-dd HH:mm')
}

function Format-MbtBrowseText {
    param([AllowNull()][string]$Text, [int]$Width)
    if ($null -eq $Text) { $Text = '' }
    $Text = ($Text -replace '[\r\n\t]+', ' ')
    if ($Text.Length -gt $Width) { return $Text.Substring(0, [Math]::Max(1, $Width - 1)) + '~' }
    return $Text.PadRight($Width)
}

function Show-MbtBrowseTree {
    <#
    .SYNOPSIS
        Prints the folder tree: indented names, item count, size ("unknown" when the service does not give it, never a number), unread count, and the move
        state of a folder that exists in the main archive and in an auxiliary archive (Add-MbtBrowseMoveInfo). Folders from the import-export API are marked
        when the tree comes from the Mail API. The last line is the total: unknown sizes are counted apart, not added.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Tree, [int]$MaxDepth = 99)
    foreach ($n in $Tree.Notes) { Write-MbtItem WARN $n }
    Write-MbtItem INFO ("{0} folder(s), source: {1}" -f $Tree.Folders.Count, $(if ($Tree.FromMailApi) { 'Microsoft Graph Mail API (counts and sizes of each folder)' } else { 'mailbox import and export API (item counts; the size of a folder is not given, the sizes of the items are in the listing of a folder)' }))
    foreach ($f in $Tree.Folders) {
        if ($f.Depth -gt $MaxDepth) { continue }
        $name = (' ' * (2 * $f.Depth)) + $f.DisplayName
        $size = if ($f.SizeBytes -ge 0) { Format-MbtBytes $f.SizeBytes } else { 'unknown' }
        $unread = if ($f.UnreadItemCount -gt 0) { "  unread $($f.UnreadItemCount)" } else { '' }
        $tag = if ($f.Recoverable) { '  [Recoverable Items]' } elseif ($Tree.FromMailApi -and $f.Source -eq 'ImportExportApi') { "  [$($f.FolderClass), import-export API]" } else { '' }
        $move = if ($f.MoveText) { "  [$($f.MoveText)]" } else { '' }
        $state = if ($f.State -eq 'ContentNotAvailable') { "  [content not available yet: $($f.StateMessage)]" } else { '' }
        Write-Host ('  {0} {1,7} item(s) {2,10}{3}{4}{5}{6}' -f (Format-MbtBrowseText $name 46), $f.TotalItemCount, $size, $unread, $tag, $move, $state)
    }
    Write-Host ('  total: ' + (Format-MbtBrowseTotals $Tree.Totals()))
}

function Format-MbtBrowseTotals {
    <# "40 folder(s), 1234 item(s), 12.3 GiB (3 folders with unknown size, not counted)"; the size is "unknown" when no folder gives one. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Totals)
    $size = if ($Totals.SizeUnknown) { 'size unknown' } else { Format-MbtBytes $Totals.KnownSizeBytes }
    $unknown = if ($Totals.FoldersWithUnknownSize -gt 0 -and -not $Totals.SizeUnknown) { " ($($Totals.FoldersWithUnknownSize) folder$(if ($Totals.FoldersWithUnknownSize -ne 1) { 's' }) with unknown size, not counted)" } else { '' }
    return ('{0} folder(s), {1} item(s), {2}{3}' -f $Totals.Folders, $Totals.Items, $size, $unknown)
}

function Show-MbtBrowseItems {
    <# Prints the rows of a folder, one line each (numbered), from Get-MbtBrowseItem. #>
    [CmdletBinding()]
    param([Parameter(Mandatory, ValueFromPipeline)]$Item)
    begin { $n = 0 }
    process {
        $n++
        if ($Item.Source -eq 'MailApi') {
            $flags = (@($(if ($Item.HasAttachments) { 'A' }), $(if ($Item.Importance -eq 'high') { '!' } elseif ($Item.Importance -eq 'low') { 'v' }), $(if ($Item.IsRead -eq $false) { 'U' }), $(if ($Item.FlagStatus -and $Item.FlagStatus -ne 'notFlagged') { 'F' }), $(if ($Item.Sensitivity -and $Item.Sensitivity -ne 'normal') { 'S' })) | Where-Object { $_ }) -join ''
            $from = if ($Item.FromName) { $Item.FromName } else { $Item.FromAddress }
            $recipients = if ($Item.ToCount -ge 0) { "to {0}/cc {1}" -f $Item.ToCount, $Item.CcCount } else { '' }
            $cats = if ($Item.Categories.Count) { '  {' + ($Item.Categories -join ',') + '}' } else { '' }
            Write-Host ('  {0,3}  {1}  {2}  {3}  {4,9}  {5,-5} {6,-10} {7}{8}' -f $n, (Format-MbtBrowseDate $(if ($Item.Received) { $Item.Received } else { $Item.Sent })), (Format-MbtBrowseText $from 24), (Format-MbtBrowseText $Item.Subject 46),
                $(if ($Item.Size -ge 0) { Format-MbtBytes $Item.Size } else { '-' }), $flags, $recipients, $Item.ItemClass, $cats)
        } else {
            Write-Host ('  {0,3}  created {1}  modified {2}  {3,9}  {4,-18} {5}' -f $n, (Format-MbtBrowseDate $Item.Created), (Format-MbtBrowseDate $Item.Modified), $(if ($Item.Size -ge 0) { Format-MbtBytes $Item.Size } else { '-' }), $Item.ItemClass, $Item.Id.Substring([Math]::Max(0, $Item.Id.Length - 16)))
        }
    }
}

function Show-MbtBrowsePreview {
    <# Prints a preview: header, recipients, body excerpt, attachments (attached items indented). #>
    [CmdletBinding()]
    param([Parameter(Mandatory, ValueFromPipeline)]$Preview, [int]$Indent = 2)
    process {
        $pad = ' ' * $Indent
        $h = $Preview.Header
        foreach ($n in $Preview.Notes) { Write-MbtItem WARN $n }
        if ($h.Source -eq 'MailApi') {
            Write-Host "${pad}Subject    : $($h.Subject)"
            Write-Host "${pad}From       : $($h.FromName) <$($h.FromAddress)>"
            if ($Preview.To.Count) { Write-Host "${pad}To         : $($Preview.To -join '; ')" }
            if ($Preview.Cc.Count) { Write-Host "${pad}Cc         : $($Preview.Cc -join '; ')" }
            if ($Preview.Bcc.Count) { Write-Host "${pad}Bcc        : $($Preview.Bcc -join '; ')" }
            Write-Host "${pad}Received   : $(Format-MbtBrowseDate $h.Received)   Sent: $(Format-MbtBrowseDate $h.Sent)   Size: $(if ($h.Size -ge 0) { Format-MbtBytes $h.Size } else { '-' })"
            Write-Host "${pad}Class      : $($h.ItemClass)   Importance: $($h.Importance)   Sensitivity: $($h.Sensitivity)   Flag: $($h.FlagStatus)   Categories: $($h.Categories -join ', ')"
            if ($Preview.BodyText) { Write-Host "${pad}Body       : $(($Preview.BodyText -replace '\s+', ' ').Trim())$(if ($Preview.BodyTrimmed) { ' [...]' })" }
            elseif ($Preview.BodyPreview) { Write-Host "${pad}Body       : $(($Preview.BodyPreview -replace '\s+', ' ').Trim())" }
        } else {
            Write-Host "${pad}Item       : $($h.Id)"
            Write-Host "${pad}Class      : $($h.ItemClass)   Size: $(if ($h.Size -ge 0) { Format-MbtBytes $h.Size } else { '-' })   Created: $(Format-MbtBrowseDate $h.Created)   Modified: $(Format-MbtBrowseDate $h.Modified)"
        }
        foreach ($a in $Preview.Attachments) {
            Write-Host ("${pad}Attachment : {0} ({1}, {2}{3})" -f $a.Name, $a.Kind, (Format-MbtBytes $a.Size), $(if ($a.IsInline) { ', inline' }))
            if ($a.Item) { Show-MbtBrowsePreview -Preview $a.Item -Indent ($Indent + 4) }
        }
    }
}

function Show-MbtExportPlan {
    <# Prints what an export plan will take, and what cannot be exported. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Plan)
    foreach ($f in $Plan.Folders) {
        $what = if ($f.ItemIds) { "$($f.ItemIds.Count) item(s)" } elseif ($f.IncludeSubfolders) { 'folder and its subfolders' } else { 'folder' }
        $where = if ($f.MailboxId) { " in $($f.MailboxId)" } else { '' }
        # what the entry takes, as the service announces it now: the items, and the size when every folder gives one (the import-export API of an auxiliary archive gives none)
        $count = ''
        if ($f.Resolved -and $f.Items -ge 0) {
            $size = if ($f.SizeComplete) { Format-MbtBytes $f.Bytes } elseif ($f.Bytes -ge 0) { "{0} known, {1} folder(s) of unknown size" -f (Format-MbtBytes $f.Bytes), $f.FoldersUnknownSize } else { 'size unknown' }
            $count = "{0} item(s) in {1} folder(s), {2}" -f $f.Items, $f.FolderCount, $size
        }
        Write-MbtJobObject -Step plan -Key $f.Path -State $(if ($f.Resolved) { 'planned' } else { 'refused' }) -Total $(if ($f.Resolved) { $f.Items } else { -1 }) -Bytes $(if ($f.Resolved -and $f.SizeComplete) { $f.Bytes } else { -1 }) -Msg ("{0} [{1}]{2}" -f $what, $f.Match, $(if ($f.Note) { " - $($f.Note)" }))
        Write-MbtItem $(if ($f.Resolved) { 'OK' } else { 'FAIL' }) ("{0}: {1}{2} [{3}]{4}{5}" -f $f.Path, $what, $where, $f.Match, $(if ($count) { " - $count" }), $(if ($f.Resolved -and $f.Note) { " - $($f.Note)" }))
    }
    foreach ($n in $Plan.Notes) { Write-MbtItem INFO $n }
    foreach ($p in $Plan.Problems) { Write-MbtItem WARN $p }
}

#endregion

#region 6. Export ----------------------------------------------------------------------------

function Export-MbtItems {
    <#
    .SYNOPSIS
        Exports items with exportItems, 20 at most per request, and writes each one to a file as the answer arrives (nothing
        is held in memory). Adds the items to manifest.json in the output folder and returns what was done.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$MailboxId,
        [Parameter(Mandatory)][string[]]$ItemId,
        [Parameter(Mandatory)][string]$OutputPath,
        [string]$FolderId,
        [string]$FolderPath,
        [object[]]$ItemInfo = @()
    )
    $s = Get-MbtSession
    $null = New-Item -ItemType Directory -Force -Path $OutputPath
    $listed = @{}
    foreach ($i in $ItemInfo) { $listed[$i.Id] = $i }
    $batchSize = [int]$s.Configuration.Graph.ExportBatchSize
    $entries = [System.Collections.Generic.List[object]]::new()
    $batches = [System.Collections.Generic.List[object]]::new()
    $n = 0
    foreach ($group in [MailboxTransfer.MailboxApi]::Chunk([string[]]$ItemId, $batchSize)) {
        $n++
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $result = $s.Api.ExportItems($MailboxId, $group, $OutputPath, $entries.Count)
        $watch.Stop()
        $bytes = [long]0
        foreach ($e in $result) { $bytes += $e.Bytes }
        $seconds = $watch.Elapsed.TotalSeconds
        $batches.Add([pscustomobject]@{ Batch = $n; Items = $group.Count; Bytes = $bytes; Seconds = $seconds; MiBPerSecond = if ($seconds -gt 0) { [math]::Round($bytes / 1MB / $seconds, 3) } else { 0 } })
        foreach ($e in $result) {
            $info = if ($e.ItemId) { $listed[$e.ItemId] } else { $null }
            $entries.Add([pscustomobject][ordered]@{
                SourceMailboxId = $MailboxId
                SourceFolderId = $FolderId
                SourceFolderPath = $FolderPath
                ItemId = $e.ItemId
                ChangeKey = $e.ChangeKey
                ItemType = if ($info) { $info.Type } else { $null }
                ListedSize = if ($info) { $info.Size } else { $null }
                File = if ($e.FilePath) { Split-Path $e.FilePath -Leaf } else { $null }
                Bytes = $e.Bytes
                Sha256 = $e.Sha256
                ExportedUtc = [DateTime]::UtcNow.ToString('o')
                BatchSeconds = [math]::Round($seconds, 3)
                BatchBytes = $bytes
                ErrorCode = $e.ErrorCode
                ErrorMessage = $e.ErrorMessage
            })
        }
    }
    $manifestPath = Join-Path $OutputPath 'manifest.json'
    $all = [System.Collections.Generic.List[object]]::new()
    if (Test-Path -LiteralPath $manifestPath) {
        $existing = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        foreach ($old in @($existing.Items)) { $all.Add($old) }
    }
    foreach ($e in $entries) { $all.Add($e) }
    [ordered]@{ Tool = 'MailboxTransfer'; Version = $script:ToolVersion; UpdatedUtc = [DateTime]::UtcNow.ToString('o'); Items = $all } |
        ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $manifestPath -Encoding utf8
    return [pscustomobject]@{ Entries = $entries.ToArray(); Batches = $batches.ToArray(); ManifestPath = $manifestPath }
}

#endregion

#region 7. Store (export store: folders, items, packs, previews, import jobs) ------------------

function Open-MbtExportStore {
    <#
    .SYNOPSIS
        Opens the export store of a folder (catalog.db + packs, see src\Engine.Store.cs), creating it when there is none.
        One process at a time may hold it open for writing: dispose it ($store.Dispose()) when done. -ReadOnly browses it
        while another process exports.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [string]$Name = 'phase0', [string]$Source = 'MailboxTransfer', [switch]$ReadOnly)
    Initialize-MbtEngine
    $full = if ([IO.Path]::IsPathRooted($Path)) { $Path } else { Join-Path (Get-Location).Path $Path }
    [MailboxTransfer.FileNaming]::Require($full, $false, 'Store folder')
    if (Test-Path -LiteralPath (Join-Path $full 'catalog.db')) {
        if ($ReadOnly) { return [MailboxTransfer.ExportStore]::OpenReadOnly($full) }
        return [MailboxTransfer.ExportStore]::Open($full)
    }
    if ($ReadOnly) { throw "No export store in $full." }
    return [MailboxTransfer.ExportStore]::Create($full, $Name, $Source)
}

function Save-MbtFolderTree {
    <# Records a mailbox and its folder tree (Get-MbtMailboxTree) in the store: the structure of the source. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store, [Parameter(Mandatory)][string]$MailboxId, [Parameter(Mandatory)][ValidateSet('Primary', 'Archive')][string]$Kind,
        [string]$UserPrincipalName, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Tree
    )
    [void]$Store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $MailboxId; Kind = $Kind; Upn = $UserPrincipalName; DisplayName = "$Kind mailbox" })
    $folders = [System.Collections.Generic.List[MailboxTransfer.StoreFolder]]::new()
    foreach ($f in $Tree) {
        $folders.Add([MailboxTransfer.StoreFolder]@{
            MailboxId = $MailboxId; FolderId = $f.Id; ParentFolderId = $f.ParentFolderId; DisplayName = $f.DisplayName; FolderType = $f.Type
            SourceItemCount = [long]$f.TotalItemCount; SourceSize = if ($f.TotalSize -ge 0) { [long]$f.TotalSize } else { $null }
        })
    }
    if ($folders.Count) { [void]$Store.UpsertFolders($folders) }
}

function Save-MbtItemList {
    <# Records the items listed in a folder (Get-MbtFolderItems) in the store, as Pending (nothing is downloaded yet). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Store, [Parameter(Mandatory)][string]$MailboxId, [Parameter(Mandatory)][string]$FolderId, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Items)
    $entries = [System.Collections.Generic.List[MailboxTransfer.StoreItemEntry]]::new()
    foreach ($i in $Items) {
        $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = $MailboxId; ItemId = $i.Id; FolderId = $FolderId; ChangeKey = $i.ChangeKey; ItemType = $i.Type; Size = [long]$i.Size })
    }
    if ($entries.Count) { [void]$Store.UpsertItems($entries) }
}

function ConvertTo-MbtStorePreview {
    <# What the browser of the store shows of an item, from the summary of the FastTransfer reader. #>
    param([Parameter(Mandatory)]$Summary, [Parameter(Mandatory)][long]$ItemKey)
    $m = $Summary
    $names = @($m.Attachments | ForEach-Object { if ($_.FileName) { $_.FileName } elseif ($_.DisplayName) { $_.DisplayName } })
    $to = @($m.Recipients | Where-Object { $_.Type -eq 1 } | ForEach-Object { if ($_.DisplayName) { $_.DisplayName } else { $_.Smtp } })
    $sent = if ($m.ClientSubmitTime) { $m.ClientSubmitTime.UtcDateTime } else { $null }
    $received = if ($m.DeliveryTime) { $m.DeliveryTime.UtcDateTime } else { $sent }
    $p = [MailboxTransfer.StorePreview]@{
        ItemKey = $ItemKey; Subject = $m.Subject; From = (($m.SenderName, $m.SenderSmtp | Where-Object { $_ }) -join ' <') + $(if ($m.SenderName -and $m.SenderSmtp) { '>' })
        To = ($to -join '; '); MessageClass = $m.MessageClass; HasAttachments = ($m.Attachments.Count -gt 0); AttachmentCount = $m.Attachments.Count
        AttachmentNames = ($names -join '; '); BodyPreview = $(if ($m.BodyTextPreview) { $m.BodyTextPreview } else { $m.BodyHtmlPreview })
    }
    $p.SentUtc = $sent
    $p.ReceivedUtc = $received
    return $p
}

function Export-MbtItemsToStore {
    <#
    .SYNOPSIS
        Exports items with exportItems (20 at most per request) into the store: each answer is streamed to a temporary file
        (nothing is held in memory), copied into a pack of the store (SHA-256 compared with the one computed on arrival),
        decoded for its preview (subject, sender, recipients, dates, attachments, body excerpt), then the temporary file is
        deleted. An item that fails is marked Failed in the store with the error. Returns Entries and Batches.
    .DESCRIPTION
        The items must be known by the store (Save-MbtItemList); with -FolderId an unknown item is registered on the fly.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store,
        [Parameter(Mandatory)][string]$MailboxId,
        [Parameter(Mandatory)][string[]]$ItemId,
        [Parameter(Mandatory)][string]$TempPath,
        [string]$FolderId,
        [object[]]$ItemInfo = @()
    )
    $s = Get-MbtSession
    $tmp = Join-Path $TempPath 'export-tmp'
    $null = New-Item -ItemType Directory -Force -Path $tmp
    $listed = @{}
    foreach ($i in $ItemInfo) { $listed[$i.Id] = $i }
    $unknown = @($ItemId | Where-Object { $null -eq $Store.FindItemKey($MailboxId, $_) })
    if ($unknown.Count) {
        if (-not $FolderId) { throw "$($unknown.Count) item(s) are not in the store: list their folder first (Save-MbtItemList) or give -FolderId." }
        Save-MbtItemList -Store $Store -MailboxId $MailboxId -FolderId $FolderId -Items @($unknown | ForEach-Object { [pscustomobject]@{ Id = $_; ChangeKey = $null; Type = $null; Size = 0 } })
    }
    $batchSize = [int]$s.Configuration.Graph.ExportBatchSize
    $entries = [System.Collections.Generic.List[object]]::new()
    $batches = [System.Collections.Generic.List[object]]::new()
    $n = 0
    try {
        foreach ($group in [MailboxTransfer.MailboxApi]::Chunk([string[]]$ItemId, $batchSize)) {
            $n++
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $result = $s.Api.ExportItems($MailboxId, $group, $tmp, $entries.Count)
            $watch.Stop()
            $bytes = [long]0
            foreach ($e in $result) { $bytes += $e.Bytes }
            $seconds = $watch.Elapsed.TotalSeconds
            $batches.Add([pscustomobject]@{ Batch = $n; Items = $group.Count; Bytes = $bytes; Seconds = $seconds; MiBPerSecond = if ($seconds -gt 0) { [math]::Round($bytes / 1MB / $seconds, 3) } else { 0 } })
            foreach ($e in $result) {
                $info = if ($e.ItemId) { $listed[$e.ItemId] } else { $null }
                $key = if ($e.ItemId) { $Store.FindItemKey($MailboxId, $e.ItemId) } else { $null }
                $entry = [ordered]@{
                    ItemKey = $key; ItemId = $e.ItemId; ChangeKey = $e.ChangeKey; ListedSize = if ($info) { $info.Size } else { $null }
                    Bytes = $e.Bytes; Sha256 = $e.Sha256; PackId = $null; Offset = $null; Sha256Matches = $null; Summary = $null; ErrorCode = $e.ErrorCode; ErrorMessage = $e.ErrorMessage
                    BatchSeconds = [math]::Round($seconds, 3)
                }
                if ($e.ErrorCode -or -not $e.FilePath) {
                    if ($null -ne $key) { $Store.MarkItemFailed($key, ("{0} {1}" -f $e.ErrorCode, $e.ErrorMessage).Trim()) }
                } elseif ($null -eq $key) {
                    $entry.ErrorCode = 'NotInStore'; $entry.ErrorMessage = 'the answer carries an item the store does not know'
                } else {
                    $in = [IO.File]::OpenRead($e.FilePath)
                    try { $w = $Store.WriteItem($key, $in, [Threading.CancellationToken]::None, $true) } finally { $in.Dispose() }
                    $entry.PackId = $w.PackId; $entry.Offset = $w.Offset; $entry.Sha256Matches = ($w.Sha256 -ieq $e.Sha256)
                    $read = [MailboxTransfer.FxReader]::ReadFile($e.FilePath)
                    $entry.Summary = [MailboxTransfer.FxMessage]::Summarize($read)
                    $Store.SetPreview($key, (ConvertTo-MbtStorePreview -Summary $entry.Summary -ItemKey $key))
                    Remove-Item -LiteralPath $e.FilePath -Force
                }
                $entries.Add([pscustomobject]$entry)
            }
        }
    } finally {
        if ((Test-Path -LiteralPath $tmp) -and -not (Get-ChildItem -LiteralPath $tmp -Force)) { Remove-Item -LiteralPath $tmp -Force }
    }
    return [pscustomobject]@{ Entries = $entries.ToArray(); Batches = $batches.ToArray() }
}

function Get-MbtStoreItems {
    <# The items of the store in a state (Exported by default), folder by folder, with their folder path. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Store, [string]$MailboxId, [MailboxTransfer.StoreItemState]$State = [MailboxTransfer.StoreItemState]::Exported, [int]$PageSize = 500)
    $ids = if ($MailboxId) { @($MailboxId) } else { @($Store.GetMailboxes() | ForEach-Object Id) }
    foreach ($id in $ids) {
        $stack = [System.Collections.Generic.Stack[MailboxTransfer.StoreFolderNode]]::new()
        foreach ($root in $Store.GetFolderTree($id)) { $stack.Push($root) }
        while ($stack.Count) {
            $node = $stack.Pop()
            foreach ($child in $node.Children) { $stack.Push($child) }
            $page = 0
            do {
                $result = $Store.GetItems($node.FolderKey, $page, $PageSize, [MailboxTransfer.StoreItemSort]::ReceivedDate, $false, $State)
                foreach ($item in $result.Items) { $item | Add-Member -NotePropertyName FolderPath -NotePropertyValue $node.Path -PassThru }
                $page++
            } while ($page -lt $result.PageCount)
        }
    }
}

function Import-MbtStoreItem {
    <#
    .SYNOPSIS
        Creates ONE item in a folder of a mailbox from an item of the store (import job of the store: Pending -> InFlight ->
        Done / Failed, createImportSession, importItem in "create" mode streamed from the pack, no Authorization header).
        Nothing is replaced, nothing is deleted. A failure after which the item may exist (timeout, network) leaves the job
        InFlight: the store turns it into Unknown at its next opening and never replays it by itself.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]$Store,
        [Parameter(Mandatory)][long]$ItemKey,
        [Parameter(Mandatory)][string]$MailboxId,
        [Parameter(Mandatory)][string]$FolderId,
        [string]$JobName = 'import'
    )
    $s = Get-MbtSession
    $item = $Store.GetItem($ItemKey)
    if (-not $item -or $item.State -ne [MailboxTransfer.StoreItemState]::Exported) { throw "Store item $ItemKey is not exported." }
    if (-not $PSCmdlet.ShouldProcess("$MailboxId folder $FolderId", "Create an item from store item $ItemKey ($(Format-MbtBytes $item.Length))")) { return }
    $job = $Store.CreateImportJob($JobName, $MailboxId)
    [void]$Store.AddImportItems($job, @([MailboxTransfer.StoreImportPlan]@{ ItemKey = $ItemKey; TargetMailbox = $MailboxId; TargetFolder = $FolderId }))
    $todo = @($Store.NextImportItems($job, 1))[0]
    try {
        $session = $s.Api.CreateImportSession($MailboxId)
        $result = $s.Api.ImportStoreItem($session, $FolderId, $Store, $ItemKey)
    } catch {
        $ge = Get-MbtGraphException $_
        if ($ge -and $ge.StatusCode -ge 400 -and $ge.StatusCode -lt 500) { $Store.MarkImportFailed($todo.ImportKey, $ge.Message) }
        Stop-MbtImportFailure $_
    }
    $Store.MarkImportDone($todo.ImportKey, $result.ItemId)
    return $result
}

#endregion

#region 8. Decode (FastTransfer reader) ------------------------------------------------------

function Read-MbtFxStream {
    <#
    .SYNOPSIS
        Decodes an exported file, or an item of the store (-Store, -ItemKey: read from its pack), with the MS-OXCFXICS reader
        (never throws, never loads an attachment in memory).
        Returns Read (envelope, every property: tag, type, size) and Summary (subject, sender, recipients, dates, body, attachments).
    #>
    [CmdletBinding(DefaultParameterSetName = 'File')]
    param(
        [Parameter(Mandatory, ParameterSetName = 'File')][string]$Path,
        [Parameter(Mandatory, ParameterSetName = 'Store')]$Store,
        [Parameter(Mandatory, ParameterSetName = 'Store')][long]$ItemKey
    )
    Initialize-MbtEngine
    if ($PSCmdlet.ParameterSetName -eq 'Store') {
        $in = $Store.OpenItem($ItemKey)
        try { $read = [MailboxTransfer.FxReader]::Read($in) } finally { $in.Dispose() }
        $Path = "store item $ItemKey"
    } else {
        $read = [MailboxTransfer.FxReader]::ReadFile($Path)
    }
    [pscustomobject]@{ Path = $Path; Read = $read; Summary = [MailboxTransfer.FxMessage]::Summarize($read) }
}

function Show-MbtFxReport {
    <# Human readable report of Read-MbtFxStream. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Decoded, [switch]$AllProperties, [int]$MaxProperties = 400)
    $r = $Decoded.Read
    $m = $Decoded.Summary
    Write-MbtItem INFO ("{0}: {1} read, envelope: {2}" -f (Split-Path $Decoded.Path -Leaf), (Format-MbtBytes $r.BytesRead), $r.Envelope)
    Write-MbtDetail "first bytes: $($r.HeadHex)"
    foreach ($rec in $r.EnvelopeRecords) {
        $kind = if ($rec.IsData) { 'FastTransfer data' } else { 'header' }
        Write-MbtDetail ("envelope record type {0} length {1} at offset {2} ({3}) {4}" -f $rec.Type, $rec.Length, $rec.Offset, $kind, $rec.HeadHex)
    }
    Write-MbtDetail ("FastTransfer data starts at offset {0}; {1} properties, {2} markers" -f $r.DataOffset, $r.PropertyCount, $r.MarkerCount)
    if ($r.Error) { Write-MbtItem WARN "reader stopped at offset $($r.ErrorOffset): $($r.Error)" }
    if ($r.Truncated) { Write-MbtItem WARN 'the stream is truncated (or shorter than announced)' }

    $rows = @($r.Elements | ForEach-Object {
        [pscustomobject]@{
            Offset = $_.Offset
            Tag = if ($_.IsMarker) { $_.Name } else { $_.TagText }
            Name = if ($_.IsMarker) { 'marker' } else { $_.Name }
            Type = $_.TypeName
            Bytes = if ($_.IsMarker) { '' } else { $_.ValueBytes }
            Where = $_.Path
            Value = $_.Display
        }
    })
    $shown = if ($AllProperties) { $rows } else { $rows | Select-Object -First $MaxProperties }
    $shown | Format-Table -AutoSize -Wrap | Out-String -Width 220 | ForEach-Object { $_.TrimEnd() } | Write-Host
    if ($rows.Count -gt $shown.Count) { Write-MbtDetail "... $($rows.Count - $shown.Count) more elements (use -AllProperties)" }
    $named = @($r.Elements | Where-Object { $_.IsNamed })
    if ($named.Count) {
        Write-MbtDetail 'named properties (GUID, lid or name):'
        foreach ($e in ($named | Select-Object -First 20)) { Write-MbtDetail ("  {0} {1}" -f $e.TagText, $e.NamedText) }
    }

    Write-Host ''
    Write-Host '  Extracted' -ForegroundColor Cyan
    $fmt = { param($label, $value) if ($null -ne $value -and "$value" -ne '') { Write-Host ('    {0,-16} {1}' -f $label, $value) } }
    & $fmt 'Subject' $m.Subject
    & $fmt 'Class' $m.MessageClass
    & $fmt 'From' (($m.SenderName, $m.SenderSmtp, $m.SenderAddress | Where-Object { $_ }) -join ' | ')
    & $fmt 'Sent as' (($m.SentRepresentingName, $m.SentRepresentingSmtp, $m.SentRepresentingAddress | Where-Object { $_ }) -join ' | ')
    & $fmt 'Submit time' $m.ClientSubmitTime
    & $fmt 'Delivery time' $m.DeliveryTime
    & $fmt 'Created' $m.CreationTime
    & $fmt 'Modified' $m.LastModificationTime
    & $fmt 'Message id' $m.InternetMessageId
    foreach ($rc in $m.Recipients) { & $fmt "Recipient ($($rc.TypeName))" (($rc.DisplayName, $rc.Smtp, $rc.Address | Where-Object { $_ }) -join ' | ') }
    if ($m.BodyTextBytes -ge 0) { & $fmt 'Body text' ("{0} bytes: {1}" -f $m.BodyTextBytes, $m.BodyTextPreview) }
    if ($m.BodyHtmlBytes -ge 0) { & $fmt 'Body HTML' ("{0} bytes: {1}" -f $m.BodyHtmlBytes, $m.BodyHtmlPreview) }
    if ($m.HasCompressedRtf) { & $fmt 'Body RTF' ("COMPRESSED RTF present ({0} bytes), not decompressed" -f $m.RtfBytes) }
    foreach ($a in $m.Attachments) {
        $label = if ($a.FileName) { $a.FileName } elseif ($a.DisplayName) { $a.DisplayName } else { '(no name)' }
        $kind = if ($a.EmbeddedMessage) { " embedded message '$($a.EmbeddedSubject)'" } else { '' }
        & $fmt 'Attachment' ("{0}  size {1} (stream data {2}){3}{4}" -f $label, (Format-MbtBytes $a.Size), (Format-MbtBytes $a.DataBytes), $kind, $(if ($a.DataTruncated) { ' TRUNCATED' } else { '' }))
    }
    if ($m.HasFolderInfo) { Write-MbtItem OK 'Folder information IS in the stream (PidTagFolderId / ParentFolderId / ParentSourceKey):' }
    else { Write-MbtItem INFO 'No PidTagFolderId / PidTagParentFolderId / PidTagParentSourceKey in the stream: the source folder is not stored in it.' }
    foreach ($p in $m.FolderInfoProperties) { Write-MbtDetail $p }
}

#endregion

#region 9. Import ----------------------------------------------------------------------------

function Import-MbtItem {
    <#
    .SYNOPSIS
        Creates ONE item in a folder of a mailbox from an exported file (createImportSession, then importItem in "create"
        mode on the pre-authenticated URL, without Authorization header). Nothing is replaced, nothing is deleted.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)][string]$MailboxId,
        [Parameter(Mandatory)][string]$FolderId,
        [Parameter(Mandatory)][string]$Path
    )
    $s = Get-MbtSession
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "File not found: $Path" }
    if ($PSCmdlet.ShouldProcess("$MailboxId folder $FolderId", "Create an item from $(Split-Path $Path -Leaf) ($(Format-MbtBytes (Get-Item -LiteralPath $Path).Length))")) {
        try {
            $session = $s.Api.CreateImportSession($MailboxId)
            return $s.Api.ImportItem($session, $FolderId, (Resolve-Path -LiteralPath $Path).Path)
        } catch { Stop-MbtImportFailure $_ }
    }
}

#endregion

#region 10. Transfer (whole mailbox export and import, status) -------------------------------

# Context: the owner's own tooling to migrate or back up Exchange Online mailboxes the operator is authorised to access (delegated
# consent of the mailbox user). Throttling is obeyed, not worked around: Retry-After exactly, concurrency halved on 429 / 503
# (see src\Engine.Transfer.cs), never more than the 4 concurrent requests per mailbox the service documents, one application only.

function Resolve-MbtMailboxId {
    <#
    .SYNOPSIS
        The physical id of the primary or archive mailbox: the id given, else (primary only) read from Graph, else an error that says
        how to get it (no API lists the archive of a user: Exchange Online PowerShell, Get-MailboxLocation).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][ValidateSet('Primary', 'Archive', 'AuxArchive')][string]$Mailbox,
        [Parameter(Mandatory)][string]$TenantId,
        [Parameter(Mandatory)][string]$UserPrincipalName,
        [string]$PrimaryMailboxId,
        [string]$ArchiveMailboxId,
        # one explicit location (any kind): wins over everything else
        [string]$MailboxId,
        # where else to look: the store (learned and recorded locations) and the operator's Get-MailboxLocation output
        $Store,
        [string]$LocationsFile
    )
    if ($MailboxId) {
        if (Test-MbtMailboxAddress $MailboxId) { return Resolve-MbtMailboxAddress -Address $MailboxId.Trim() -TenantId $TenantId }
        return ConvertTo-MbtMailboxId -Value $MailboxId -TenantId $TenantId
    }
    $given = if ($Mailbox -eq 'Primary') { $PrimaryMailboxId } elseif ($Mailbox -eq 'Archive') { $ArchiveMailboxId } else { $null }
    if ($given) { return ConvertTo-MbtMailboxId -Value $given -TenantId $TenantId }
    $known = @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $Store -LocationsFile $LocationsFile -NoSave | Where-Object Kind -EQ $Mailbox)
    if ($known.Count -eq 1) { return $known[0].MailboxId }
    if ($known.Count -gt 1) {
        throw ("$($known.Count) $Mailbox mailboxes are known for ${UserPrincipalName}: " + (($known | ForEach-Object { $_.MailboxId }) -join ', ') + '. Choose one with -MailboxId.')
    }
    if ($Mailbox -eq 'Primary') { return Get-MbtPrimaryMailboxId -UserPrincipalName $UserPrincipalName -TenantId $TenantId }
    if ($Mailbox -eq 'AuxArchive') {
        throw ("No auxiliary archive of $UserPrincipalName is known yet. Pass -MailboxId MBX:<mailbox guid>@$TenantId (Exchange Online PowerShell: Get-MailboxLocation -User $UserPrincipalName, " +
            "MailboxLocationType AuxArchive), or -LocationsFile with that output; they are also learned from the redirects of an Archive export or import.")
    }
    throw ("The archive mailbox id of $UserPrincipalName is not known (Graph beta gives inPlaceArchiveMailboxId with User.Read.All, nothing else does): pass -ArchiveMailboxId MBX:<mailbox guid>@$TenantId " +
        "(Exchange Online PowerShell: Get-MailboxLocation -User $UserPrincipalName, MailboxLocationType MainArchive) or -LocationsFile with that output.")
}

function Get-MbtConfiguredMailboxId {
    <#
    .SYNOPSIS
        The Primary and Archive mailbox ids of the configuration (Lab section), for the user of the configuration only.
    .DESCRIPTION
        An id written in the configuration is the id of one mailbox. It is returned only when -UserPrincipalName is the user the configuration
        names (Lab.UserPrincipalName); for any other user both are $null and the ids are discovered (Graph for the primary mailbox, the locations
        of the store or -LocationsFile, an explicit -PrimaryMailboxId / -ArchiveMailboxId), so that a request for one user never lands in the
        mailbox of another.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Configuration, [string]$UserPrincipalName)
    $r = [pscustomobject]@{ Primary = $null; Archive = $null }
    $lab = if ($Configuration -and $Configuration.PSObject.Properties['Lab']) { $Configuration.Lab } else { $null }
    if ($lab -and $UserPrincipalName -and $lab.UserPrincipalName -and $lab.UserPrincipalName -ieq $UserPrincipalName) {
        if ($lab.PrimaryMailboxId) { $r.Primary = [string]$lab.PrimaryMailboxId }
        if ($lab.ArchiveMailboxId) { $r.Archive = [string]$lab.ArchiveMailboxId }
    }
    return $r
}

function Assert-MbtSignInSettings {
    <#
    .SYNOPSIS
        Stops with a message that says what to give when the tenant or the application (client) id of the sign-in is missing or malformed.
    #>
    [CmdletBinding()]
    param([string]$TenantId, [string]$ClientId)
    $guid = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    $problems = @()
    if (-not $TenantId) { $problems += 'the tenant (directory id, or the domain name of the tenant)' }
    elseif ($TenantId -notmatch $guid -and $TenantId -notmatch '^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') { $problems += "a valid tenant (a GUID or a domain name, not '$TenantId')" }
    if (-not $ClientId) { $problems += 'the application (client) id of the app registration' }
    elseif ($ClientId -notmatch $guid) { $problems += "a valid application (client) id (a GUID, not '$ClientId')" }
    if ($problems) {
        throw ('Sign-in needs ' + ($problems -join ' and ') + '. Give -TenantId and -ClientId, or write TenantId and ClientId in the Lab section of the configuration file (-ConfigPath); ' +
            'keep that file out of version control (config\*.local.psd1). The application is the one registered for this tool in the tenant (public client, delegated permissions).')
    }
}

function Write-MbtJobMessage {
    <# One message of a job (warnings and refusals in yellow, the rest in grey). #>
    param([Parameter(Mandatory)][string]$Text)
    $status = if ($Text -match 'failed|UNKNOWN|not listed|not created|not imported|throttled|lowered|refused|Unknown') { 'WARN' } else { 'INFO' }
    Write-MbtItem $status $Text
}

function Wait-MbtTransferJob {
    <#
    .SYNOPSIS
        Follows a job of the engine (ExportJob, ImportJob, ReplicateJob) from this thread: writes its messages, a progress line every
        -IntervalSeconds (items, bytes, MiB/s overall and recent, GiB/h, 429, concurrency) and the windows of 5 minutes, until it ends.
        Ctrl+C stops the job after the requests in flight (the store stays consistent: run again to resume).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Job, [int]$IntervalSeconds = 5)
    $next = [DateTime]::UtcNow.AddSeconds($IntervalSeconds)
    $lastPhase = ''
    $stopSaid = $false
    try {
        # a stop request made before the job began (the file was there at the start) is honoured at once
        if (Test-MbtStopRequested) {
            $stopSaid = $true
            Write-MbtItem WARN 'stop requested: the requests in flight finish first, run the command again to resume'
            $Job.Cancel()
        }
        while (-not $Job.Wait(500)) {
            $messages = @($Job.TakeMessages())
            foreach ($m in $messages) { Write-MbtJobMessage $m }
            # a stop-request file (the console's "Stop and keep progress") has the effect of Ctrl+C: the requests in flight finish, run again to resume
            if (-not $stopSaid -and (Test-MbtStopRequested)) {
                $stopSaid = $true
                Write-MbtItem WARN 'stop requested: the requests in flight finish first, run the command again to resume'
                $Job.Cancel()
            }
            if ([DateTime]::UtcNow -ge $next) {
                $next = [DateTime]::UtcNow.AddSeconds($IntervalSeconds)
                Write-MbtDetail ($Job.Snapshot().ToString())
                Update-MbtJobProgress -Job $Job -Messages $messages
            } elseif ($messages.Count) { Update-MbtJobProgress -Job $Job -Messages $messages }
        }
    } finally {
        if (-not $Job.IsCompleted) { Write-MbtItem WARN 'stopping: the requests in flight finish first, run the command again to resume'; $Job.Cancel(); [void]$Job.Wait(120000) }
    }
    $last = @($Job.TakeMessages())
    foreach ($m in $last) { Write-MbtJobMessage $m }
    Update-MbtJobProgress -Job $Job -Messages $last -Final
    return $Job
}

function Write-MbtThrottleReport {
    <# Every 429 / 503 of a job with the time waited, and the windows of the sustained rate. #>
    param([Parameter(Mandatory)]$Job)
    $events = @($Job.GetThrottleEvents())
    if ($events.Count) {
        Write-MbtItem INFO "$($events.Count) throttling answer(s) (429 / 503), each one waited for as the service asked:"
        foreach ($e in $events | Select-Object -First 50) { Write-MbtDetail $e.ToString() }
        if ($events.Count -gt 50) { Write-MbtDetail "... $($events.Count - 50) more" }
    } else { Write-MbtItem OK 'no throttling answer (429 / 503)' }
    $windows = @($Job.GetWindows())
    if ($windows.Count) {
        Write-MbtItem INFO 'sustained rate by window:'
        foreach ($w in $windows) { Write-MbtDetail $w.ToString() }
    }
}

function Export-MbtMailbox {
    <#
    .SYNOPSIS
        Exports a whole mailbox (primary, or the archive with -Archive) into the store: folder tree, inventory of the items, then the
        items in batches of 20 items / 200 MiB by up to 4 workers, streamed into the packs. A new run does only what is left (Pending
        items, Failed ones tried fewer than -MaxAttempts times, items whose change key moved). Needs Connect-MbtGraph first.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Store,
        [Parameter(Mandatory)][string]$MailboxId,
        [switch]$Archive,
        # Primary (default), Archive (main archive: expanded folders are followed into the auxiliary archives) or AuxArchive (one auxiliary archive by itself)
        [ValidateSet('Primary', 'Archive', 'AuxArchive')][string]$Kind,
        [string]$UserPrincipalName,
        [ValidateRange(1, 4)][int]$Concurrency = 4,
        [ValidateRange(1, 20)][int]$BatchItems = 20,
        [long]$BatchBytes = 200MB,
        [int]$MaxAttempts = 3,
        # after this many folders in a row whose content is not available yet, the next ones are asked once instead of 3 times (0: always 3 times)
        [ValidateRange(0, 100000)][int]$NotReadyTrip = 3,
        [switch]$SkipInventory,
        [string]$Folder,
        # What a browse selection asked for (Resolve-MbtBrowsePlan): these folders / items only. Exclusive with -Folder.
        $Plan,
        # Also export Recoverable Items (deleted items, purges, versions, holds), under the folder "Recoverable Items": -RecoverableItemsFolders names the folders (default Deletions, Purges, Versions, DiscoveryHolds)
        [switch]$IncludeRecoverableItems,
        [string[]]$RecoverableItemsFolders,
        [Parameter(Mandatory)][string]$TempPath,
        [int]$IntervalSeconds = 5
    )
    $s = Get-MbtSession
    $o = [MailboxTransfer.ExportOptions]::new()
    if (-not $Kind) { $Kind = if ($Archive) { 'Archive' } else { 'Primary' } }
    $o.MailboxId = $MailboxId; $o.Kind = $Kind; $o.Upn = $UserPrincipalName
    $o.Concurrency = $Concurrency; $o.BatchItems = $BatchItems; $o.BatchBytes = $BatchBytes; $o.MaxAttempts = $MaxAttempts; $o.NotReadyTrip = $NotReadyTrip
    $o.SkipInventory = [bool]$SkipInventory; $o.FolderPath = if ($Folder) { $Folder }; $o.TempPath = $TempPath
    if ($Plan) { $o.Plan = $Plan }
    $o.IncludeRecoverableItems = [bool]$IncludeRecoverableItems
    if ($RecoverableItemsFolders) { $o.RecoverableItemsFolders = [string[]]$RecoverableItemsFolders }
    $job = [MailboxTransfer.ExportJob]::Start($s.Api, $Store, $o)
    $null = Wait-MbtTransferJob -Job $job -IntervalSeconds $IntervalSeconds
    $r = $job.Result
    if ($job.Error) { Write-MbtItem FAIL "export stopped: $($job.Error.Message)" }
    Write-MbtItem $(if ($r.Failed -eq 0 -and $r.Completed) { 'OK' } else { 'WARN' }) ("export {0}: {1} items exported ({2}), {3} failed, {4} folders, {5} items listed, {6} previews with a warning" -f
        $(if ($r.Completed) { 'finished' } else { 'NOT finished' }), $r.Exported, (Format-MbtBytes $r.Bytes), $r.Failed, $r.Folders, $r.Inventoried, $r.PreviewWarnings)
    if ($r.Final) { Write-MbtDetail $r.Final.ToString() }
    if ($r.RecoverableFolders -gt 0) { Write-MbtItem INFO ("Recoverable Items: {0} folder(s) recorded under '{1}'. An import puts them in a normal folder named '{2}', never back into Recoverable Items." -f $r.RecoverableFolders, [MailboxTransfer.RecoverableItems]::TopName, [MailboxTransfer.RecoverableItems]::RestoredName) }
    if ($r.FoldersMoved -gt 0 -or $r.AuxMailboxes.Count -gt 0) {
        Write-MbtItem INFO ("auto-expanding archive: {0} expanded folder(s) (recorded once in the main archive tree), items recorded in {1} auxiliary archive(s): {2}{3}" -f $r.FoldersMoved, $r.AuxMailboxes.Count, ($r.AuxMailboxes -join ', '),
            $(if ($r.ItemsMoved) { "; $($r.ItemsMoved) item(s) moved to their auxiliary archive after export" } else { '' }))
    }
    foreach ($m in $r.MovesInProgress) { Write-MbtItem INFO $m }
    if ($r.ItemsDeduplicated -gt 0 -or $r.ItemsRehomed -gt 0) {
        Write-MbtItem INFO ("move in progress: {0} item(s) listed on both sides of a move were recorded once, {1} record(s) of items not exported yet followed the item to the mailbox that lists it now" -f $r.ItemsDeduplicated, $r.ItemsRehomed)
    }
    if ($r.FoldersDeferred.Count -gt 0) {
        Write-MbtItem WARN ("{0} folder(s) DEFERRED, not exported (the service cannot read their content yet: a folder being moved to an auxiliary archive{1}); run the same command again later:" -f $r.FoldersDeferred.Count, $(if ($r.FoldersDeferredFast -gt 0) { "; $($r.FoldersDeferredFast) of them asked once because the location kept refusing, -NotReadyTrip" } else { '' }))
        foreach ($d in $r.FoldersDeferred | Select-Object -First 50) { Write-MbtDetail $d.ToString() }
        if ($r.FoldersDeferred.Count -gt 50) { Write-MbtDetail "... $($r.FoldersDeferred.Count - 50) more" }
    }
    Write-MbtThrottleReport -Job $job
    return [pscustomobject]@{ Result = $r; Job = $job; Error = $job.Error }
}

function Import-MbtMapping {
    <#
    .SYNOPSIS
        Reads a folder mapping file (.psd1 or .json): source path prefix -> target path prefix (relative to -TargetRoot), as a
        dictionary the engine takes. Example (.psd1): @{ 'Inbox/Projects' = 'Archive/Work'; 'Sent Items' = 'Archive/Sent' }
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Mapping file not found: $Path" }
    $raw = if ($Path -like '*.json') { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json -AsHashtable } else { Import-PowerShellDataFile -LiteralPath $Path }
    $map = [System.Collections.Generic.Dictionary[string, string]]::new()
    foreach ($key in $raw.Keys) {
        if ($raw[$key] -isnot [string]) { throw "Mapping file ${Path}: the target of '$key' must be a path (text)." }
        $map[[string]$key] = [string]$raw[$key]
    }
    return $map
}

function Write-MbtRestoredDeletedNote {
    <# An export that holds Recoverable Items: say plainly where the import puts them (a normal folder, never Recoverable Items). #>
    [CmdletBinding()]
    param([string[]]$SourcePath)
    Initialize-MbtEngine
    if (@($SourcePath | Where-Object { $_ -and [MailboxTransfer.RecoverableItems]::IsPath($_) }).Count) {
        Write-MbtItem INFO ("deleted items (Recoverable Items) in this source are imported into a normal folder named '{0}' in the target (below -TargetRoot if given), never back into Recoverable Items: Exchange does not accept an import there. They come back as ordinary items of that folder." -f [MailboxTransfer.RecoverableItems]::RestoredName)
    }
}

function Import-MbtMailbox {
    <#
    .SYNOPSIS
        Imports the Exported items of the store into a mailbox (create only): recreates the folder tree below -TargetRoot (or at the
        same paths, or following -MappingFile), then imports the items by up to 4 workers with one import session reused until it
        expires. States of the store: Pending -> InFlight -> Done / Failed / Unknown. Running again with the same -JobName resumes.
        With -WhatIf only the plan is shown: nothing is created, nothing is sent.
        Automatic reconciliation (on by default, -NoAutoReconcile turns it off): an item whose answer is missing (timeout, 504, connection
        lost after the whole body was sent) does not stop the run. After -ReconcileDelaySeconds (default 300: the service may still be
        creating it) the target folder is listed and the item looked for (same size, not known as imported, created after the attempt began):
        found once = Done; not found = imported again in the same run (one more attempt, -MaxAttempts respected, at most -ReconcileRounds
        lookups per item); several candidates = left Unknown and reported. Unknown rows left by an earlier run are looked up the same way
        at the start of the run, once their delay is over.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory)]$Store,
        [Parameter(Mandatory)][string]$SourceMailboxId,
        [Parameter(Mandatory)][string]$TargetMailboxId,
        [string]$TargetRoot,
        [string]$MappingFile,
        [string]$SourceFolder,
        [string]$JobName = 'import',
        [ValidateRange(1, 4)][int]$Concurrency = 4,
        [int]$MaxAttempts = 3,
        [switch]$NoAutoReconcile,
        [ValidateRange(0, 86400)][double]$ReconcileDelaySeconds = 300,
        [ValidateRange(1, 20)][int]$ReconcileRounds = 3,
        [int]$IntervalSeconds = 5
    )
    $s = Get-MbtSession
    $o = [MailboxTransfer.ImportOptions]::new()
    $o.AutoReconcile = -not $NoAutoReconcile; $o.ReconcileDelaySeconds = $ReconcileDelaySeconds; $o.ReconcileRounds = $ReconcileRounds
    $o.SourceMailboxId = $SourceMailboxId; $o.TargetMailboxId = $TargetMailboxId; $o.TargetRoot = if ($TargetRoot) { $TargetRoot }
    $o.Concurrency = $Concurrency; $o.JobName = $JobName; $o.MaxAttempts = $MaxAttempts; $o.SourceFolderPath = if ($SourceFolder) { $SourceFolder }
    if ($MappingFile) { $o.Mapping = Import-MbtMapping -Path $MappingFile }
    $plan = [MailboxTransfer.ImportPlanner]::Build($Store, $o, $s.Api.GetFolderTree($TargetMailboxId))
    Write-MbtItem INFO ("plan: {0} items ({1}) from {2} folders; {3} folders already exist in the target, {4} to create" -f $plan.Items, (Format-MbtBytes $plan.Bytes), $plan.Folders.Count, $plan.FoldersExisting, $plan.PathsToCreate.Count)
    foreach ($p in $plan.PathsToCreate | Select-Object -First 40) { Write-MbtDetail "would create folder: $p" }
    if ($plan.PathsToCreate.Count -gt 40) { Write-MbtDetail "... $($plan.PathsToCreate.Count - 40) more folders" }
    Write-MbtRestoredDeletedNote -SourcePath @($plan.Folders | ForEach-Object SourcePath)
    foreach ($f in $plan.Folders | Where-Object { $_.Items -gt 0 } | Select-Object -First 40) { Write-MbtDetail ("{0} -> {1}: {2} items, {3}" -f $f.SourcePath, $f.TargetPath, $f.Items, (Format-MbtBytes $f.Bytes)) }
    foreach ($f in $plan.Folders | Where-Object { $_.Items -gt 0 } | Select-Object -First 500) { Write-MbtJobObject -Step plan -Key $f.TargetPath -State planned -Total $f.Items -Bytes $f.Bytes -Msg ("from " + $f.SourcePath) }
    if ($plan.PathsToCreate.Count) { Write-MbtJobObject -Step plan -Key '(folders)' -State planned -Total $plan.PathsToCreate.Count -Msg 'folders to create' }
    if ($NoAutoReconcile) { Write-MbtItem INFO 'automatic reconciliation is off (-NoAutoReconcile): an item that ends Unknown stays Unknown until a person settles it' }
    else {
        $earlier = 0
        $existing = $Store.FindImportJob($JobName)
        if ($existing) { $earlier = [long]($Store.GetImportStats($existing.JobId) | Where-Object State -EQ 'Unknown').Count }
        Write-MbtItem INFO ("automatic reconciliation is on: an item that ends Unknown is looked up in its target folder {0} s after the request ended (at most {1} lookups per item); {2} Unknown row(s) of the job '{3}' left by earlier runs would be looked up first" -f $ReconcileDelaySeconds, $ReconcileRounds, $earlier, $JobName)
    }
    if (-not $PSCmdlet.ShouldProcess($TargetMailboxId, "Create $($plan.PathsToCreate.Count) folders and import $($plan.Items) items (create only)")) { return [pscustomobject]@{ Plan = $plan; Result = $null; Job = $null; Error = $null } }
    $job = [MailboxTransfer.ImportJob]::Start($s.Api, $Store, $o)
    $null = Wait-MbtTransferJob -Job $job -IntervalSeconds $IntervalSeconds
    $r = $job.Result
    if ($job.Error) { Write-MbtItem FAIL "import stopped: $($job.Error.Message)" }
    # The counts of the line are THIS run's; the job state after the semicolon is the whole job (earlier runs included).
    $jobState = if (-not $r.JobKnown) { "job '$($r.JobName)' state not read" }
        else { "job '{0}' {1}: {2} of {3} done, {4} failed, {5} unknown, {6} pending" -f $r.JobName, $(if ($r.JobComplete) { 'COMPLETE' } else { 'INCOMPLETE' }), $r.JobDone, $r.JobTotal, $r.JobFailed, $r.JobUnknown, $r.JobPending }
    Write-MbtItem $(if ($r.Failed -eq 0 -and $r.Unknown -eq 0 -and $r.Completed -and $r.JobComplete) { 'OK' } else { 'WARN' }) ("import {0}: {1} items imported ({2}), {3} failed (nothing created), {4} unknown (may exist), {5} folders created in this run; {6}" -f
        $(if ($r.Completed) { 'finished' } else { 'NOT finished' }), $r.Done, (Format-MbtBytes $r.Bytes), $r.Failed, $r.Unknown, $r.FoldersCreated, $jobState)
    Write-MbtItem $(if ($r.JobComplete) { 'OK' } else { 'WARN' }) $r.JobSummary()
    if ($r.ReconcileSummary()) { Write-MbtItem $(if ($r.ReconcileAmbiguous -or $r.ReconcileExhausted) { 'WARN' } else { 'OK' }) $r.ReconcileSummary() }
    if ($r.FoldersInAux -gt 0) { Write-MbtItem INFO ("auto-expanding archive: {0} folder(s) of the target live in {1} auxiliary archive(s) ({2}); their items were sent there, nothing was created in an auxiliary archive by this routing" -f $r.FoldersInAux, $r.AuxArchives.Count, ($r.AuxArchives -join ', ')) }
    if ($r.SessionLifetimeSeconds.Count) { Write-MbtDetail ("{0} import session(s) created, announced lifetime {1}" -f $r.SessionsCreated, (($r.SessionLifetimeSeconds | ForEach-Object { Format-MbtDuration $_ }) -join ', ')) }
    if ($r.Final) { Write-MbtDetail $r.Final.ToString() }
    if ($r.JobUnknown -gt 0) {
        $why = if ($NoAutoReconcile) { 'automatic reconciliation is off' } else { 'the lookup in the target folder could not settle them (several items could be them, or the rounds are used up, or the run was stopped)' }
        Write-MbtItem WARN ("{0} Unknown row(s) in the job, {1} of them from an earlier run ({2}): list them with Get-MbtImportUnknown -Check (or Invoke-MailboxTransfer.ps1 -Mode Status -Unknown -Check), settle them with Resolve-MbtImportUnknown (one row: -ImportKey <key> -Imported|-NotImported; every row with no match in the target: -Absent), then run the import again" -f $r.JobUnknown, $r.JobUnknownEarlier, $why)
    }
    elseif ($r.JobKnown -and -not $r.JobComplete) { Write-MbtItem WARN 'run the same import again (same -JobName) to continue: Failed rows are tried again while attempts remain' }
    Write-MbtThrottleReport -Job $job
    return [pscustomobject]@{ Plan = $plan; Result = $r; Job = $job; Error = $job.Error }
}

function Show-MbtPstImportPlan {
    <# The plan of an import from PST (Import-MbtPst -WhatIf and the start of a real run): files, folders, items, size, estimated duration. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Plan, [int]$MaxFolders = 60)
    foreach ($s in $Plan.Sources) {
        $level = if ($s.Damaged -and -not $s.Repaired) { 'WARN' } else { 'INFO' }
        Write-MbtItem $level ("{0} ({1}): {2}" -f $s.SourcePath, (Format-MbtBytes $s.Length), $(if ($s.Checked) { $s.Note } else { 'not checked' }))
    }
    foreach ($n in $Plan.Notes) { Write-MbtItem WARN $n }
    $est = if ($Plan.EstimatedSeconds -gt 0) { "estimated $(Format-MbtDuration $Plan.EstimatedSeconds) at the observed rate of about 4 GiB per hour per mailbox" } else { 'duration not estimated' }
    Write-MbtItem INFO ("plan: {0} items ({1}) in {2} folders; {3} folders already exist in the target, {4} to create; {5}" -f $Plan.Items, (Format-MbtBytes $Plan.Bytes), $Plan.Folders.Count, $Plan.FoldersExisting, $Plan.PathsToCreate.Count, $est)
    if ($Plan.Unreadable) { Write-MbtItem WARN "$($Plan.Unreadable) item(s) cannot be read from the PST or expressed as a FastTransfer stream: they will be reported as refused, the others are imported" }
    if ($Plan.Associated) { Write-MbtItem INFO "$($Plan.Associated) folder-associated item(s) (rules, views, settings) are not imported" }
    foreach ($p in $Plan.PathsToCreate | Select-Object -First $MaxFolders) { Write-MbtDetail "would create folder: $p" }
    if ($Plan.PathsToCreate.Count -gt $MaxFolders) { Write-MbtDetail "... $($Plan.PathsToCreate.Count - $MaxFolders) more folders" }
    Write-MbtRestoredDeletedNote -SourcePath @($Plan.Folders | ForEach-Object PstPath)
    foreach ($f in $Plan.Folders | Where-Object { $_.Items -gt 0 } | Select-Object -First $MaxFolders) {
        $wk = if ($f.WellKnown) { " [$($f.WellKnown)]" } else { '' }
        Write-MbtDetail ("{0}{1} -> {2}: {3} items, {4}" -f $(if ($f.PstPath) { $f.PstPath } else { '(top)' }), $wk, $f.TargetPath, $f.Items, (Format-MbtBytes $f.Bytes))
    }
    # the plan for the console: one object per folder of the target (the same keys the folders of the run have)
    foreach ($f in $Plan.Folders | Where-Object { $_.Items -gt 0 } | Select-Object -First 500) { Write-MbtJobObject -Step plan -Key $f.TargetPath -State planned -Total $f.Items -Bytes $f.Bytes -Msg ("from " + $(if ($f.PstPath) { $f.PstPath } else { '(top)' })) }
    if ($Plan.PathsToCreate.Count) { Write-MbtJobObject -Step plan -Key '(folders)' -State planned -Total $Plan.PathsToCreate.Count -Msg 'folders to create' }
}

function Import-MbtPst {
    <#
    .SYNOPSIS
        Imports the items of one or more PST files into a mailbox (create only): the folder tree of the PST is recreated below -TargetFolder
        (container classes kept), the items are read with our own PST reader, converted to FastTransfer streams with our encoder and sent with
        importItem (up to 4 workers, one import session per mailbox, throttling obeyed). Nothing is replaced, moved or deleted in the mailbox.
    .DESCRIPTION
        -Path             a PST file, several files, or a folder of PST files (the parts of a split PST, name.pst, name-0002.pst, ..., are imported in order
                          as one tree).
        -TargetFolder     the folder below which the tree of the PST is created (for instance Imported/2019). The Inbox, Sent Items, Deleted Items ... of
                          the PST become subfolders of it: they are never merged into the real Inbox of the mailbox unless -MergeWellKnown is given
                          (then the well-known folders of the PST go into the matching folders of the mailbox; the others below -TargetFolder).
        -FolderMap        a hashtable (or a .psd1 / .json file): PST path prefix (below the top folder of the PST) -> target path prefix (below -TargetFolder).
        -PerFileSubfolder each PST file gets its own subfolder (named after its store).
        -IncludeFolder / -ExcludeFolder   only some folders of the PST: paths below the top folder ("Inbox", "Inbox/Projects"), wildcards * and ?. A path chooses
                          the folder and the folders below it; "Inbox/*" is what is below the Inbox without the Inbox itself; an exclusion wins. A pattern that matches no
                          folder is reported, and an import in which no folder is chosen stops before anything is created.
        The journal (-JournalPath, a SQLite file; default <Output.Root>\transfer\pstimport.db) has one row per message keyed by PST file, node id and a fingerprint
        of its content: running the same command again resumes the job (-JobName) and never creates an item twice. An item whose answer is missing is
        looked up in the target folder after -ReconcileDelaySeconds (the candidates are exported and compared with the message); -NoAutoReconcile leaves it Unknown.
        Each file is checked first with our integrity checker. A file with errors is refused, unless -Repair: then it is repaired into a temporary copy
        (-WorkPath; the file itself is never modified) and the copy is imported. -NoCheck skips the check.
        An item that cannot be read from the PST, cannot be expressed or is refused by the service is reported with the reason and the run goes on.
        -WhatIf shows the plan (files and their check, folders, items, size, estimated duration) and creates nothing.
        Returns Plan, Result (PstImportRunResult: counts, NotImported with the reason of each item), Job and Error. Needs Connect-MbtGraph first.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param(
        [Parameter(Mandatory, Position = 0)][string[]]$Path,
        [ValidateSet('Primary', 'Archive', 'AuxArchive')][string]$Mailbox = 'Primary',
        [string]$MailboxId,
        [string]$PrimaryMailboxId,
        [string]$ArchiveMailboxId,
        [string]$LocationsFile,
        [string]$TargetFolder,
        $FolderMap,
        [switch]$MergeWellKnown,
        [switch]$PerFileSubfolder,
        [string[]]$IncludeFolder,
        [string[]]$ExcludeFolder,
        [switch]$NoCheck,
        [switch]$Repair,
        [ValidateSet('Recompute', 'Rebuild', 'Salvage')][string]$RepairLevel = 'Rebuild',
        [switch]$KeepDamagedAttachments,
        [string]$JournalPath,
        [string]$WorkPath,
        [string]$JobName = 'pstimport',
        [ValidateRange(1, 4)][int]$Concurrency = 4,
        [int]$MaxAttempts = 3,
        [switch]$NoAutoReconcile,
        [ValidateRange(0, 86400)][double]$ReconcileDelaySeconds = 300,
        [ValidateRange(1, 20)][int]$ReconcileRounds = 3,
        [double]$AssumedGiBPerHour = 4,
        [int]$IntervalSeconds = 5
    )
    Initialize-MbtEngine
    $s = Get-MbtSession
    $upn = Get-MbtSessionUser $s
    # the ids of the configuration only for the user of the configuration (never in app-only for another mailbox); for anyone else they are discovered
    $configured = if ($s.Configuration) { Get-MbtConfiguredMailboxId -Configuration $s.Configuration -UserPrincipalName $upn } else { $null }
    if (-not $PrimaryMailboxId -and $configured -and $configured.Primary) { $PrimaryMailboxId = $configured.Primary }
    if (-not $ArchiveMailboxId -and $configured -and $configured.Archive) { $ArchiveMailboxId = $configured.Archive }
    $tenant = $s.TenantId
    if (-not $tenant -and $MailboxId -match '@(?<t>[0-9a-fA-F-]{36})$') { $tenant = $Matches.t }
    if (-not $tenant) { $tenant = 'unknown' }
    if (-not $upn) { $upn = 'unknown' }
    $targetId = Resolve-MbtMailboxId -Mailbox $Mailbox -TenantId $tenant -UserPrincipalName $upn -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -MailboxId $MailboxId -LocationsFile $LocationsFile
    if (-not $TargetFolder -and -not $MergeWellKnown) { throw '-TargetFolder is required: the folder below which the folders of the PST are created (for instance Imported/2019), or -MergeWellKnown.' }
    $o = [MailboxTransfer.PstImportOptions]::new()
    foreach ($p in $Path) { $o.PstPaths.Add($p) }
    $o.TargetMailboxId = $targetId; $o.TargetFolder = $TargetFolder
    $o.MergeWellKnown = [bool]$MergeWellKnown; $o.PerFileSubfolder = [bool]$PerFileSubfolder
    foreach ($f in @($IncludeFolder | Where-Object { $_ })) { $o.IncludeFolders.Add($f) }
    foreach ($f in @($ExcludeFolder | Where-Object { $_ })) { $o.ExcludeFolders.Add($f) }
    $o.Check = -not $NoCheck; $o.Repair = [bool]$Repair; $o.RepairLevel = $RepairLevel; $o.KeepDamagedAttachments = [bool]$KeepDamagedAttachments
    $o.WorkPath = $WorkPath; $o.JobName = $JobName; $o.Concurrency = $Concurrency; $o.MaxAttempts = $MaxAttempts
    $o.AutoReconcile = -not $NoAutoReconcile; $o.ReconcileDelaySeconds = $ReconcileDelaySeconds; $o.ReconcileRounds = $ReconcileRounds; $o.AssumedGiBPerHour = $AssumedGiBPerHour
    if ($FolderMap) {
        $map = if ($FolderMap -is [string]) { Import-MbtMapping -Path $FolderMap } else {
            $m = [System.Collections.Generic.Dictionary[string, string]]::new()
            foreach ($k in $FolderMap.Keys) { $m[[string]$k] = [string]$FolderMap[$k] }
            $m
        }
        $o.FolderMap = $map
    }
    $root = if ($s.Configuration -and $s.Configuration.Output.Root) { $s.Configuration.Output.Root } else { '.\out' }
    if (-not $JournalPath) { $JournalPath = Join-Path (Join-Path $root 'transfer') 'pstimport.db' }
    if (-not [IO.Path]::IsPathRooted($JournalPath)) { $JournalPath = Join-Path (Get-Location).Path $JournalPath }
    if (-not $WorkPath) { $o.WorkPath = Join-Path (Split-Path $JournalPath -Parent) 'pstimport-work' }

    Write-MbtItem INFO ("reading the PST file(s) and the folder tree of the {0} mailbox {1}" -f $Mailbox, $targetId)
    $sink = [MailboxTransfer.PstProgressSink]::new([Action[MailboxTransfer.PstProgress]]{
            param($e)
            if ($e.BytesTotal -gt 0) { Write-Progress -Activity 'Import-MbtPst' -Status ("{0}: {1:N0} / {2:N0} bytes" -f $e.Phase, $e.BytesDone, $e.BytesTotal) -PercentComplete ([math]::Min(100, [int](100 * $e.BytesDone / $e.BytesTotal))) -Id 7 }
            elseif ($e.Items -gt 0) { Write-Progress -Activity 'Import-MbtPst' -Status ("{0}: {1} items {2}" -f $e.Phase, $e.Items, $e.CurrentFolder) -Id 7 }
        })
    $plan = [MailboxTransfer.PstImportPlanner]::Build($o, $s.Api.GetFolderTree($targetId), $sink, [Threading.CancellationToken]::None)
    Write-Progress -Activity 'Import-MbtPst' -Completed -Id 7
    Show-MbtPstImportPlan -Plan $plan
    if ($NoAutoReconcile) { Write-MbtItem INFO 'automatic reconciliation is off (-NoAutoReconcile): an item that ends Unknown stays Unknown' }
    $blocked = $plan.Blocked -and -not $Repair
    if ($blocked) { Write-MbtItem FAIL 'a PST file is damaged: nothing will be imported. Repair it (Repair-MbtPst) or run again with -Repair to import from a repaired copy.' }
    if (-not $PSCmdlet.ShouldProcess($targetId, "Create $($plan.PathsToCreate.Count) folders and import $($plan.Items) items from $($plan.Sources.Count) PST file(s) (create only)")) {
        return [pscustomobject]@{ Plan = $plan; Result = $null; Job = $null; Error = $null }
    }
    if ($blocked) { throw 'A PST file is damaged (see above): nothing was imported. Use -Repair to import from a repaired copy.' }
    $journal = [MailboxTransfer.PstImportJournal]::new($JournalPath)
    try {
        $job = [MailboxTransfer.PstImportJob]::Start($s.Api, $journal, $o, $o.WorkPath, $null)
        $null = Wait-MbtTransferJob -Job $job -IntervalSeconds $IntervalSeconds
        $r = $job.Result
        if ($job.Error) { Write-MbtItem FAIL "import stopped: $($job.Error.Message)" }
        $jobState = if (-not $r.JobKnown) { "job '$($r.JobName)' state not read" }
            else { "job '{0}' {1}: {2} of {3} done, {4} failed, {5} unknown, {6} refused, {7} pending" -f $r.JobName, $(if ($r.JobComplete) { 'COMPLETE' } else { 'INCOMPLETE' }), $r.JobDone, $r.JobTotal, $r.JobFailed, $r.JobUnknown, $r.JobRefused, $r.JobPending }
        Write-MbtItem $(if ($r.Failed -eq 0 -and $r.Unknown -eq 0 -and $r.Refused -eq 0 -and $r.Completed -and $r.JobClean) { 'OK' } else { 'WARN' }) ("import {0}: {1} items imported ({2}), {3} failed (nothing created), {4} unknown (may exist), {5} refused, {6} folders created in this run; {7}" -f
            $(if ($r.Completed) { 'finished' } else { 'NOT finished' }), $r.Done, (Format-MbtBytes $r.Bytes), $r.Failed, $r.Unknown, $r.Refused, $r.FoldersCreated, $jobState)
        Write-MbtItem $(if ($r.JobClean) { 'OK' } else { 'WARN' }) $r.JobSummary()
        $reportPath = $null
        if ($r.NotImported.Count -gt 0) {
            $reportPath = Join-Path (Split-Path $JournalPath -Parent) ("{0}-not-imported.csv" -f ($JobName -replace '[^\w\.\-]', '_'))
            try { $null = Export-MbtPstImportReport -Result $r -Path $reportPath; Write-MbtItem INFO "the items that were not imported are listed, one per line, with the reason, in $reportPath" }
            catch { $reportPath = $null; Write-MbtItem WARN "the list of the items that were not imported could not be written: $($_.Exception.Message)" }
        }
        if ($r.ReconcileSummary()) { Write-MbtItem $(if ($r.ReconcileAmbiguous -or $r.ReconcileExhausted) { 'WARN' } else { 'OK' }) $r.ReconcileSummary() }
        foreach ($i in $r.NotImported | Select-Object -First 50) {
            Write-MbtItem WARN ("{0}: {1} / {2} [{3}] {4}" -f $i.State, $i.Folder, $(if ($i.Subject) { $i.Subject } else { '(no subject)' }), $i.MessageClass, $i.Reason)
        }
        if ($r.NotImported.Count -gt 50) { Write-MbtDetail "... $($r.NotImported.Count - 50) more item(s) not imported (Result.NotImported)" }
        if ($r.SessionLifetimeSeconds.Count) { Write-MbtDetail ("{0} import session(s) created, announced lifetime {1}" -f $r.SessionsCreated, (($r.SessionLifetimeSeconds | ForEach-Object { Format-MbtDuration $_ }) -join ', ')) }
        if ($r.Final) { Write-MbtDetail $r.Final.ToString() }
        if ($r.JobUnknown -gt 0) { Write-MbtItem WARN ("{0} Unknown row(s) in the job ({1} from an earlier run): the item may exist in the target folder; run the same command again (the lookup is done once the delay is over) or check the folder" -f $r.JobUnknown, $r.JobUnknownEarlier) }
        elseif ($r.JobKnown -and -not $r.JobComplete -and ($r.JobFailed + $r.JobPending) -gt 0) { Write-MbtItem WARN 'run the same command again (same -JobName) to continue: Failed rows are tried again while attempts remain' }
        Write-MbtThrottleReport -Job $job
        return [pscustomobject]@{ Plan = $plan; Result = $r; Job = $job; Error = $job.Error; ReportPath = $reportPath }
    } finally { $journal.Dispose() }
}

function Export-MbtPstImportReport {
    <#
    .SYNOPSIS
        Writes the items of an import from PST that were not imported (Refused for good, Failed, Unknown) as a CSV, one line per item with the reason.
    .DESCRIPTION
        -Result is the Result of Import-MbtPst (PstImportRunResult.NotImported). Columns: State, Folder (in the PST), Subject, Class, SizeBytes, Attempts, Reason, PstFile, Node.
        State Refused is final (the service refuses the content, the item cannot be read or expressed, or the attempts allowed are used); Failed is tried again by the next
        run; Unknown may exist in the mailbox. UTF-8 with a byte order mark so that Excel reads the accents. Returns the path.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Result, [Parameter(Mandatory)][string]$Path)
    $dir = Split-Path ([IO.Path]::GetFullPath($Path)) -Parent
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Force -Path $dir }
    $rows = foreach ($i in @($Result.NotImported)) {
        [pscustomobject][ordered]@{ State = [string]$i.State; Folder = [string]$i.Folder; Subject = [string]$i.Subject; Class = [string]$i.MessageClass; SizeBytes = [long]$i.Size; Attempts = [int]$i.Attempts; Reason = [string]$i.Reason; PstFile = [string]$i.Pst; Node = [long]$i.Nid }
    }
    @($rows) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding utf8BOM
    return $Path
}

function Get-MbtImportUnknown {
    <#
    .SYNOPSIS
        Lists the import rows in the state Unknown (the service may or may not have created the item: timeout, network error, 5xx,
        interruption, crash). With -Check, lists the target folder and shows the items that could be the one (same size as the source
        item, not known as imported by any job of the store, created or modified after the attempt began; the listing of the items returns
        no subject nor message id): a person decides, nothing is replayed by itself. The import settles these rows by itself unless
        -NoAutoReconcile was given.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Store, [string]$JobName, [switch]$Check)
    $WhatIfPreference = $false   # read only: a -WhatIf of the caller must not make ForEach-Object skip the property reads below
    $jobs = @($Store.GetImportJobs() | Where-Object { -not $JobName -or $_.Name -eq $JobName })
    foreach ($job in $jobs) {
        $unknown = @($Store.GetImportItems($job.JobId, [MailboxTransfer.StoreImportState]::Unknown, 100000))
        if (-not $unknown) { continue }
        $known = $null
        $listings = @{}
        if ($Check) {
            $s = Get-MbtSession
            $known = $Store.GetDoneTargetItemIds($null)
        }
        foreach ($u in $unknown) {
            $source = $Store.GetItem($u.ItemKey)
            $row = [ordered]@{
                JobName = $job.Name; JobId = $job.JobId; ImportKey = $u.ImportKey; ItemKey = $u.ItemKey; Subject = $source.Subject; Size = $source.Size
                TargetMailbox = $u.TargetMailbox; TargetFolder = $u.TargetFolder; Error = $u.LastError; Candidates = @()
            }
            if ($Check) {
                $k = "$($u.TargetMailbox)|$($u.TargetFolder)"
                if (-not $listings.ContainsKey($k)) { $listings[$k] = @($s.Api.EnumerateItems($u.TargetMailbox, $u.TargetFolder, 0)) }
                $notBefore = [MailboxTransfer.ImportReconciler]::NotBefore($null, $u.UpdatedUtc)
                $row.Candidates = @([MailboxTransfer.ImportReconciler]::Candidates($source.Size, [MailboxTransfer.ItemInfo[]]@($listings[$k]), $known, $notBefore))
            }
            [pscustomobject]$row
        }
    }
}

function Resolve-MbtImportUnknown {
    <#
    .SYNOPSIS
        Settles an Unknown import row once you know: -Imported (the item exists in the target; -TargetItemId records its id) or
        -NotImported (the row goes back to Pending and a later import creates the item). Nothing else ever replays an Unknown row.
        -Absent settles many rows at once: it looks for the item of every Unknown row in its target folder (the check of
        Get-MbtImportUnknown -Check, a Graph session is needed) and puts back to Pending only the rows for which no matching item was
        found; a row with candidates is left Unknown and named. -JobName limits it to one job, -WhatIf shows what would be settled.
        Run it when nothing else writes to the target folders: an item that was just created may not be listed yet.
    #>
    [CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Imported')]
    param(
        [Parameter(Mandatory)]$Store,
        [Parameter(Mandatory, ParameterSetName = 'Imported')][Parameter(Mandatory, ParameterSetName = 'NotImported')][long]$ImportKey,
        [Parameter(Mandatory, ParameterSetName = 'Imported')][switch]$Imported,
        [Parameter(Mandatory, ParameterSetName = 'NotImported')][switch]$NotImported,
        [Parameter(Mandatory, ParameterSetName = 'Absent')][switch]$Absent,
        [Parameter(ParameterSetName = 'Absent')][string]$JobName,
        [string]$TargetItemId
    )
    if ($Absent) {
        $rows = @(Get-MbtImportUnknown -Store $Store -JobName $JobName -Check)   # all the listings are read before anything is settled
        $settled = 0; $wouldSettle = 0
        foreach ($u in $rows) {
            $what = "import {0} (job '{1}'): item {2} '{3}' -> folder {4}" -f $u.ImportKey, $u.JobName, $u.ItemKey, $u.Subject, $u.TargetFolder
            if ($u.Candidates.Count) {
                Write-MbtItem WARN "$what left Unknown: $($u.Candidates.Count) item(s) in the target folder could be it ($($u.Candidates -join ', '))"
                $done = $false
            }
            elseif ($PSCmdlet.ShouldProcess("import row $($u.ImportKey)", 'no matching item in the target folder: put back to Pending')) {
                $Store.ResolveImportUnknown($u.ImportKey, $false, $null)
                Write-MbtItem OK "$what settled: no matching item found in the target folder, back to Pending"
                $settled++
                $done = $true
            }
            else { $done = $false; if ($WhatIfPreference) { $wouldSettle++ } }
            [pscustomobject]@{ ImportKey = $u.ImportKey; JobName = $u.JobName; ItemKey = $u.ItemKey; Subject = $u.Subject; Candidates = $u.Candidates; Settled = $done }
        }
        if ($WhatIfPreference) { Write-MbtItem INFO ("{0} Unknown row(s) checked: {1} would be settled as not imported (WhatIf), {2} left Unknown" -f $rows.Count, $wouldSettle, ($rows.Count - $wouldSettle)) }
        else { Write-MbtItem INFO ("{0} Unknown row(s) checked: {1} settled as not imported, {2} left Unknown" -f $rows.Count, $settled, ($rows.Count - $settled)) }
        return
    }
    if ($PSCmdlet.ShouldProcess("import row $ImportKey", $(if ($Imported) { 'mark as imported (Done)' } else { 'put back to Pending' }))) {
        $Store.ResolveImportUnknown($ImportKey, [bool]$Imported, $TargetItemId)
    }
}

function Get-MbtStoreStatus {
    <# The state of a store: mailboxes, folder trees with counts and sizes by state, packs, import jobs. Read only. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Store)
    $stats = $Store.GetStats()
    $mailboxes = foreach ($m in $Store.GetMailboxes()) {
        [pscustomobject]@{ Mailbox = $m; Items = $Store.GetItemStats($m.Id); Tree = $Store.GetFolderTree($m.Id) }
    }
    $jobs = foreach ($j in $Store.GetImportJobs()) { [pscustomobject]@{ Job = $j; States = $Store.GetImportStats($j.JobId) } }
    [pscustomobject]@{ Name = $Store.Name; Source = $Store.Source; CreatedUtc = $Store.CreatedUtc; Stats = $stats; Mailboxes = @($mailboxes); ImportJobs = @($jobs); Locations = @($Store.GetLocations()) }
}

function Show-MbtStoreStatus {
    <# Writes Get-MbtStoreStatus: tree of each mailbox (items by state and sizes), packs, import jobs. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Status, [int]$MaxFolders = 200)
    $st = $Status.Stats
    Write-MbtItem INFO ("store '{0}' (source {1}), created {2:u}" -f $Status.Name, $Status.Source, $Status.CreatedUtc)
    Write-MbtItem INFO ("{0} mailbox(es), {1} folders, {2} items, {3} previews; {4} stored in {5} pack(s) ({6} on disk)" -f $st.Mailboxes, $st.Folders, $st.Items, $st.Previews, (Format-MbtBytes $st.StoredBytes), $st.Packs, (Format-MbtBytes $st.PackBytesOnDisk))
    if ($Status.Locations.Count) {
        Write-MbtStep 'Mailbox locations (primary, main archive, auxiliary archives of an auto-expanding archive)'
        foreach ($l in $Status.Locations) {
            Write-MbtItem INFO ("{0,-10} {1}  [{2}, {3:u}{4}]" -f $l.Kind, $l.MailboxId, $l.Source, $l.DiscoveredUtc, $(if ($l.Detail) { ": $($l.Detail)" } else { '' }))
        }
    }
    foreach ($m in $Status.Mailboxes) {
        Write-MbtStep ("{0} mailbox {1}" -f $m.Mailbox.Kind, $m.Mailbox.Id)
        if ($m.Mailbox.Kind -eq 'AuxArchive') { Write-MbtDetail 'auxiliary archive: its items belong to folders recorded once in the main archive tree (shown there, with these items counted)' }
        Write-MbtDetail (($m.Items | Where-Object { $_.Count } | ForEach-Object { "{0} {1} ({2})" -f $_.State, $_.Count, (Format-MbtBytes $_.AnnouncedBytes) }) -join ', ')
        $shown = 0
        $walk = $null
        $walk = {
            param($node)
            if ($script:mbtShown -ge $MaxFolders) { return }
            $script:mbtShown++
            $t = $node.Total
            Write-Host ('  {0,-48} {1,7} items  exported {2,6}  pending {3,6}  failed {4,5}  {5,10}' -f (('  ' * $node.Depth) + $node.DisplayName), $t.Count, $t.Exported, $t.Pending, $t.Failed, (Format-MbtBytes $t.Size))
            foreach ($c in $node.Children) { & $walk $c }
        }
        $script:mbtShown = 0
        foreach ($root in $m.Tree) { & $walk $root }
        if ($script:mbtShown -ge $MaxFolders) { Write-MbtDetail "... only the first $MaxFolders folders are shown" }
        Remove-Variable -Name mbtShown -Scope Script -ErrorAction SilentlyContinue
    }
    if ($Status.ImportJobs.Count) {
        Write-MbtStep 'Import jobs'
        foreach ($j in $Status.ImportJobs) {
            Write-MbtItem INFO ("job {0} '{1}' -> {2}, created {3:u}: {4}" -f $j.Job.JobId, $j.Job.Name, $j.Job.TargetMailbox, $j.Job.CreatedUtc,
                (($j.States | Where-Object { $_.Count } | ForEach-Object { "{0} {1} ({2})" -f $_.State, $_.Count, (Format-MbtBytes $_.StoredBytes) }) -join ', '))
        }
    }
}

#endregion

#region 11. PST (own MS-PST writer, reader and integrity checker) -------------------------------

function Test-MbtPst {
    <#
    .SYNOPSIS
        Checks the integrity of a Unicode PST file (report only, nothing is repaired): header and CRCs,
        allocation maps, block and node B-trees, every block, data and subnode trees, heaps, property and table contexts, message
        store, name-to-id map, folder tree, contents tables against the messages, folder counts, recipient and attachment tables,
        embedded messages. Works on a PST written by this tool or by Outlook; a corrupt file gives findings, never an exception.
    .DESCRIPTION
        Returns an object: Path, Ok, ExitCode (0 clean, 2 warnings only, 1 errors), Errors, Warnings, Infos, Stats, Elapsed, Findings.
        Each finding has Severity (Info / Warning / Error), Layer (NDB / LTP / Messaging), Code (stable identifier), Message, Offset,
        Bid, Nid and Where (folder path and subject). -Detailed also prints the findings (at most -MaxPerCode per code).
        The whole file is read once (-NoBlockCheck skips the CRC of every block, -NoDeep skips the messages); the memory used is about
        1/256 of the file size plus 30 bytes per block and per node.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)][Alias('FullName')][string]$Path,
        [switch]$Detailed,
        [switch]$NoBlockCheck,
        [switch]$NoDeep,
        [int]$MaxPerCode = 50,
        # Called with a [MailboxTransfer.PstProgress] (Phase, BytesDone / BytesTotal of the blocks verified, Items / ItemsTotal of the messages checked) from the thread that runs the check
        [scriptblock]$OnProgress,
        # Stops the check between two blocks / two messages: the result then has Cancelled = $true and the findings so far
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )
    process {
        Initialize-MbtEngine
        $o = [MailboxTransfer.PstCheckOptions]::new()
        $o.VerifyBlocks = -not $NoBlockCheck
        $o.Deep = -not $NoDeep
        $o.MaxPerCode = $MaxPerCode
        $o.Cancellation = $CancellationToken
        $hook = if ($OnProgress) { $OnProgress } else { Get-MbtPstProgressHook -Kind Check }
        if ($hook) { $o.ProgressReport = [MailboxTransfer.PstProgressSink]::new([Action[MailboxTransfer.PstProgress]]({ param($e) & $hook $e }.GetNewClosure())) }
        $report = [MailboxTransfer.PstChecker]::Check((Resolve-Path -LiteralPath $Path).Path, $o)
        if ($Detailed) {
            Write-Host $report.Summary()
            foreach ($f in $report.Findings) {
                $color = switch ($f.Severity) { 'Error' { 'Red' } 'Warning' { 'Yellow' } default { 'Gray' } }
                Write-Host ('  ' + $f.ToString()) -ForegroundColor $color
            }
        }
        [pscustomobject]@{
            Path = $report.Path; Ok = $report.Ok; ExitCode = $report.ExitCode; Errors = $report.Errors; Warnings = $report.Warnings; Infos = $report.Infos
            Stats = $report.Stats; Elapsed = $report.Elapsed; Findings = $report.Findings; Cancelled = $report.Cancelled; Report = $report
        }
    }
}

function Repair-MbtPst {
    <#
    .SYNOPSIS
        Repairs a Unicode PST into a NEW file (the source is never modified): recomputes counts, sizes, header counters and
        allocation maps, rebuilds folder and contents tables from the nodes, recovers orphans under "Recovered items", or salvages every
        readable message of a heavily damaged file. With -WhatIf only the planned actions (from the integrity check) are shown.
    .DESCRIPTION
        -Level Recompute  (default) folders and messages are read through the hierarchy and contents tables; counts, sizes, header
                          counters, AMap / PMap / FMap / FPMap, B-trees are written again clean; references to nodes that do not exist are
                          dropped; a damaged message is kept as far as it decodes.
        -Level Rebuild    folders and messages come from the nodes of the node B-tree (their parent), the tables are not trusted: messages
                          not listed are kept in their folder, folders and messages whose parent is gone go under "Recovered items".
        -Level Salvage    the file is scanned for valid blocks and pages (signature and CRC); the B-trees are rebuilt from the scan; then as Rebuild.
        A level that cannot walk the file goes on to the next one unless -NoEscalate is given.
        At the levels Rebuild and Salvage a lost or damaged name-to-id map (named properties: calendar times, recurrence, contact fields) is searched in the blocks of the file; the result says where the map came from (NameMapSource: Original, RecoveredByScan, RebuiltFromUsage).
        -KeepDamagedAttachments  a file attachment with unreadable blocks is kept with the readable blocks and zeros in their place, its name suffixed ('name (damaged).ext') and the zero-filled byte ranges listed in the item (default: the data of such an attachment is dropped and reported).
        Returns an object: Success, Level, LevelUsed, Items counts (Recovered / Partial / Lost), Before and After (checker reports),
        Actions (code, status, message, location, the checker findings they deal with), Items (per item status and what was lost), Report.
        -Level Automatic checks the file once and repairs at the level the check recommends. Entry script mode: Invoke-MailboxTransfer.ps1 -Mode RepairPst.
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory, ValueFromPipelineByPropertyName)][Alias('FullName')][string]$Path,
        [Parameter(Mandatory)][string]$Destination,
        # Automatic: the file is checked first and repaired at the level the check recommends (the lowest that deals with all the findings)
        [ValidateSet('Automatic', 'Recompute', 'Rebuild', 'Salvage')][string]$Level = 'Recompute',
        [switch]$NoEscalate,
        [switch]$Force,
        [switch]$NoCheck,
        [switch]$KeepDamagedAttachments,
        [switch]$Detailed,
        [int]$MaxItems = 100000,
        # Called with a [MailboxTransfer.PstProgress] (the checks before and after, and the phase "repair" with Items / ItemsTotal)
        [scriptblock]$OnProgress,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )
    process {
        Initialize-MbtEngine
        $src = (Resolve-Path -LiteralPath $Path).Path
        $dst = if ([IO.Path]::IsPathRooted($Destination)) { $Destination } else { Join-Path (Get-Location).Path $Destination }
        [MailboxTransfer.FileNaming]::Require($dst, $true, 'Repair destination')
        $o = [MailboxTransfer.PstRepairOptions]::new()
        $o.Level = if ($Level -eq 'Automatic') { 'Recompute' } else { $Level }; $o.Escalate = -not $NoEscalate; $o.Overwrite = [bool]$Force; $o.MaxItemRecords = $MaxItems; $o.KeepDamagedAttachments = [bool]$KeepDamagedAttachments
        if ($NoCheck) { $o.CheckBefore = $false; $o.CheckAfter = $false }
        $toResult = {
            param($r)
            [pscustomobject]@{
                Path = $r.SourcePath; Destination = $r.TargetPath; Success = $r.Success; PlanOnly = $r.PlanOnly; Level = $r.Requested; LevelUsed = $r.Used; Recommended = $r.Recommended
                Escalated = $r.Escalated; Cancelled = $r.Cancelled; Error = $r.Error
                ItemsRecovered = $r.ItemsRecovered; ItemsPartial = $r.ItemsPartial; ItemsLost = $r.ItemsLost; OrphanItems = $r.OrphanItems; DanglingDropped = $r.DanglingDropped
                RecoveryRate = $r.RecoveryRate; Elapsed = $r.Elapsed
                NameMapSource = $r.NameMapSource; NameMapDetail = $r.NameMapDetail; AttachmentsKeptDamaged = $r.AttachmentsKeptDamaged
                Before = $r.Before; After = $r.After; Actions = $r.Actions; Items = $r.Items; Report = $r
            }
        }
        $show = {
            param($r)
            Write-Host $r.Summary()
            foreach ($a in $r.Actions) { Write-Host ('  ' + $a.ToString()) -ForegroundColor $(if ($a.Status -eq 'NotApplied') { 'Yellow' } else { 'Gray' }) }
            if ($r.ActionsOmitted) { Write-Host "  ... $($r.ActionsOmitted) more action(s) not listed" -ForegroundColor Gray }
            foreach ($i in ($r.Items | Where-Object Status -NE 'Recovered' | Select-Object -First 50)) { Write-Host ('  ' + $i.ToString()) -ForegroundColor $(if ($i.Status -eq 'Lost') { 'Red' } else { 'Yellow' }) }
        }
        $hook = if ($OnProgress) { $OnProgress } else { Get-MbtPstProgressHook -Kind Repair }
        $hookSink = if ($hook) { [MailboxTransfer.PstProgressSink]::new([Action[MailboxTransfer.PstProgress]]({ param($e) & $hook $e }.GetNewClosure())) }
        if ($WhatIfPreference) {
            $plan = [MailboxTransfer.PstRepair]::Plan($src, $o, $hookSink, $CancellationToken)
            Write-Host "What if: repair '$src' into '$dst' at level $Level$(if ($Level -eq 'Automatic') { " (recommended: $($plan.Recommended))" })."
            & $show $plan
            return (& $toResult $plan)
        }
        if (-not $PSCmdlet.ShouldProcess($dst, "Repair '$src' (level $Level) into a new file")) { return }
        if ($Level -eq 'Automatic') {
            # the check of the source is done once: its report is the "before" of the repair, and its recommendation the level
            $plan = [MailboxTransfer.PstRepair]::Plan($src, $o, $hookSink, $CancellationToken)
            if ($plan.Cancelled -or $plan.Error) { return (& $toResult $plan) }
            $o.Level = $plan.Recommended
            $o.BeforeReport = $plan.Before
            Write-Host "Automatic level: the check recommends $($plan.Recommended)."
        }
        $sink = [MailboxTransfer.PstProgressSink]::new([Action[MailboxTransfer.PstProgress]]{
                param($e)
                if ($hookSink) { $hookSink.Report($e) }
                if ($e.ItemsTotal -gt 0) { Write-Progress -Activity 'Repair-MbtPst' -Status ("{0}: {1} / {2} {3}" -f $e.Phase, $e.Items, $e.ItemsTotal, $e.CurrentFolder) -PercentComplete ([math]::Min(100, [int](100 * $e.Items / $e.ItemsTotal))) }
                elseif ($e.BytesTotal -gt 0) { Write-Progress -Activity 'Repair-MbtPst' -Status ("{0}: {1:N0} / {2:N0} bytes" -f $e.Phase, $e.BytesDone, $e.BytesTotal) -PercentComplete ([math]::Min(100, [int](100 * $e.BytesDone / $e.BytesTotal))) }
            })
        $r = [MailboxTransfer.PstRepair]::Repair($src, $dst, $o, $sink, $CancellationToken)
        Write-Progress -Activity 'Repair-MbtPst' -Completed
        if ($Detailed) { & $show $r }
        & $toResult $r
    }
}

function Resolve-MbtPstDestination {
    <#
    The first PST of an export: -Destination as a .pst file, else a folder (the file is named after the user of the store, and
    "<user>-AuxArchive1" for an export of an auxiliary archive alone: see [MailboxTransfer.PstExporter]::DefaultStem; -Options = the
    PstExportOptions that say which mailboxes are exported).
    A name given by the operator that cannot be created (Windows device name CON PRN AUX NUL COM0-9 LPT0-9 with any extension, character
    not allowed, trailing dot or space) is an error with the name to use instead, before any work: Outlook crashes on a PST named aux.pst.
    A name derived from the user is made safe ("aux" -> "aux_") and reported.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Destination, $Store, $Options)
    Initialize-MbtEngine
    $full = if ([IO.Path]::IsPathRooted($Destination)) { $Destination } else { Join-Path (Get-Location).Path $Destination }
    if ($full -match '\.pst$') {
        [MailboxTransfer.FileNaming]::Require($full, $true, 'PST destination')
        return [IO.Path]::GetFullPath($full)
    }
    [MailboxTransfer.FileNaming]::Require($full, $false, 'PST destination folder')
    $user = if ($Store) { @($Store.GetMailboxes() | ForEach-Object Upn | Where-Object { $_ })[0] } else { $null }
    $stem = if ($Store) { [MailboxTransfer.PstExporter]::DefaultStem($Store, $(if ($Options) { $Options } else { [MailboxTransfer.PstExportOptions]::new() })) } else { 'mailbox' }
    $plain = if ($user) { ($user -split '@')[0] -replace '[^\w\.\-]', '_' } else { 'mailbox' }
    if ([MailboxTransfer.FileNaming]::IsReservedDeviceName($plain)) {
        Write-MbtItem INFO ("the name derived from the user, '{0}', is a reserved Windows device name (Outlook crashes on a PST named like this): the file is named '{1}.pst'" -f $plain, $stem)
    }
    return [IO.Path]::GetFullPath((Join-Path $full "$stem.pst"))
}

function Show-MbtPstExportPlan {
    <#
    .SYNOPSIS
        Prints the plan of an export to PST ([MailboxTransfer.PstExporter]::Plan): the locations, the folders that hold items with their item counts
        and announced sizes, and the estimated number of PST files.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Plan, [string]$Destination, [int]$MaxFolders = 60)
    $names = ''
    if ($Destination) {
        $shown = @(1..[math]::Max(1, $Plan.EstimatedParts) | Select-Object -First 3 | ForEach-Object { Split-Path ([MailboxTransfer.PstExporter]::PartPath($Destination, $_)) -Leaf })
        $names = ' (' + ($shown -join ', ') + $(if ($Plan.EstimatedParts -gt 3) { ', ...' }) + ')'
    }
    Write-MbtItem INFO ("plan: {0} item(s), {1} announced by the source, {2} stored; estimated {3} PST file(s) of at most {4}{5}" -f $Plan.Items, (Format-MbtBytes $Plan.AnnouncedBytes), (Format-MbtBytes $Plan.StoredBytes), $Plan.EstimatedParts, (Format-MbtBytes $Plan.MaxPartBytes), $names)
    foreach ($l in $Plan.Locations) {
        Write-MbtItem INFO ("{0} [{1}; {2}]: {3} folder(s), {4} item(s), {5}" -f $l.Name, $l.Kind, ($l.MailboxIds -join ', '), $l.Folders, $l.Items, (Format-MbtBytes $l.AnnouncedBytes))
        $shown = 0
        foreach ($f in $Plan.Folders | Where-Object { $_.Location -eq $l.Index -and $_.Included -and $_.Items -gt 0 }) {
            if ($shown -ge $MaxFolders) { Write-MbtDetail '...'; break }
            Write-MbtDetail ("{0,-60} {1,6} item(s) {2,10}  {3}" -f $f.Path, $f.Items, (Format-MbtBytes $f.AnnouncedBytes), $f.ContainerClass)
            Write-MbtJobObject -Step plan -Key ("$($l.Name)/" + $f.Path) -State planned -Total $f.Items -Bytes $f.AnnouncedBytes
            $shown++
        }
        $empty = @($Plan.Folders | Where-Object { $_.Location -eq $l.Index -and $_.Included -and -not $_.IsTop -and $_.Items -eq 0 }).Count
        if ($empty) { Write-MbtDetail "$empty folder(s) without items (created in the first PST)" }
    }
    foreach ($n in $Plan.Notes) { Write-MbtItem WARN $n }
}


function Show-MbtLiveExportPlan {
    <#
    .SYNOPSIS
        What an export (-WhatIf) of one location would take, from the live folder tree (Get-MbtBrowseFolder -AsTree): the folders that hold items with
        their item counts and announced sizes (-1 when the service does not give a folder size), the totals, and how many PST files the announced
        size would make at -PstMaxBytes. Read only. Returns the totals of the folders shown.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Tree, [Parameter(Mandatory)][string]$Kind, [string]$Folder, [long]$PstMaxBytes = 51200MB, [int]$MaxFolders = 60,
        # the tree holds Recoverable Items (Get-MbtBrowseFolder -RecoverableItems -RecoverableItemsFolders ...): they count in the plan; otherwise they are left out
        [switch]$IncludeRecoverableItems)
    foreach ($n in $Tree.Notes) { Write-MbtItem WARN $n }
    $folders = @($Tree.Folders | Where-Object { $IncludeRecoverableItems -or -not $_.Recoverable })
    if ($Folder) {
        $root = $Tree.Find($Folder)
        if (-not $root) { Write-MbtItem WARN "folder '$Folder' is not in the $Kind mailbox: nothing would be exported from it"; $folders = @() }
        else { $folders = @($folders | Where-Object { $_.Path -ieq $root.Path -or $_.Path.StartsWith($root.Path + '/', [StringComparison]::OrdinalIgnoreCase) }) }
    }
    $totals = [MailboxTransfer.BrowseTotals]::Of([MailboxTransfer.BrowseFolder[]]$folders)
    $source = if ($Tree.FromMailApi) { 'Mail API' } else { 'import-export API' }
    Write-MbtItem INFO ("{0} {1}: {2} ({3})" -f $Kind, $Tree.Location.MailboxId, (Format-MbtBrowseTotals $totals), $source)
    $shown = 0
    $withItems = @($folders | Where-Object { $_.TotalItemCount -gt 0 })
    foreach ($f in $withItems | Where-Object { -not $_.Recoverable }) {
        if ($shown -ge $MaxFolders) { Write-MbtDetail '...'; break }
        Write-MbtDetail ("{0,-60} {1,7} item(s) {2,10}" -f $f.Path, $f.TotalItemCount, $(if ($f.SizeBytes -ge 0) { Format-MbtBytes $f.SizeBytes } else { 'unknown' }))
        Write-MbtJobObject -Step plan -Key ("$Kind/" + $f.Path) -State planned -Total $f.TotalItemCount -Bytes $f.SizeBytes
        $shown++
    }
    # Recoverable Items are few: all of them are listed, whatever the number of normal folders
    $deleted = @($withItems | Where-Object Recoverable)
    if ($deleted.Count) { Write-MbtDetail 'deleted items (Recoverable Items), saved apart under the folder "Recoverable Items":' }
    foreach ($f in $deleted) {
        Write-MbtDetail ("{0,-60} {1,7} item(s) {2,10}" -f $f.Path, $f.TotalItemCount, $(if ($f.SizeBytes -ge 0) { Format-MbtBytes $f.SizeBytes } else { 'unknown' }))
        Write-MbtJobObject -Step plan -Key ("$Kind/" + $f.Path) -State planned -Total $f.TotalItemCount -Bytes $f.SizeBytes
    }
    $empty = @($folders | Where-Object { $_.TotalItemCount -eq 0 }).Count
    if ($empty) { Write-MbtDetail "$empty folder(s) without items" }
    if ($totals.Items -gt 0) {
        if ($totals.SizeUnknown) { Write-MbtDetail ("PST files: cannot be estimated, the service gives no folder size for this location (an estimate of the number of files at {0} per file needs the sizes)" -f (Format-MbtBytes $PstMaxBytes)) }
        else { Write-MbtDetail ("PST files: about {0} at most {1} per file (announced size {2}{3})" -f [math]::Max(1, [math]::Ceiling($totals.KnownSizeBytes / $PstMaxBytes)), (Format-MbtBytes $PstMaxBytes), (Format-MbtBytes $totals.KnownSizeBytes), $(if (-not $totals.SizeComplete) { ', unknown sizes not counted' })) }
    }
    return $totals
}

function Export-MbtPst {
    <#
    .SYNOPSIS
        Converts an export store into PST files with our own MS-PST writer (no Outlook, no third-party component): the real folder tree of each
        location, split by size, resumable, with a JSON manifest and an integrity check of each file.
    .DESCRIPTION
        Layout: one top folder per location (Primary, Archive, "Archive auxiliaire n (guid8)") under the PST root, then the folders of the
        mailbox with their names, hierarchy and container class (IPF.Note, IPF.Appointment, IPF.Contact...). The standard folders (Inbox, Drafts,
        Sent Items, Deleted Items, Outbox, Junk Email, Calendar, Contacts, Tasks, Notes, Journal, also by their French or German names) come first.
        The auxiliary archives of an auto-expanding archive are merged into the Archive tree by default (the same folder entry key = the same logical
        folder); -SeparateAuxArchives gives each one its own top folder.
        Split: at most -MaxSizeGB (default 50 = 51,200 MiB, Outlook's default; or -MaxSizeMiB for a smaller limit) per file, on item boundaries;
        every file recreates the folders its items need. <name>.pst, <name>-002.pst, ... and <name>.manifest.json (parts, folders, items with the
        SHA-256 of the source stream, sizes, items larger than the limit, items that could not be written). An item larger than the limit is written
        alone in its own file and reported.
        Resumable: run the same command again after an interruption: the parts the manifest lists are verified (size and PST header) and kept, the
        unfinished part is deleted and rebuilt. -Force starts again from zero. -SkipCheck leaves out the integrity check (Test-MbtPst) of each part.
        -WhatIf prints the plan (locations, folders, item counts, announced sizes, estimated number of files) and writes nothing.
        The store is only read. Returns Plan, Result (the engine's PstExportResult) and ExitCode (0 complete, 1 error or a part with check errors,
        2 cancelled or items not written).
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory)]$Store,
        [Parameter(Mandatory)][string]$Destination,
        [ValidateRange(1, 1024)][double]$MaxSizeGB = 50,
        [ValidateRange(1, 1048576)][long]$MaxSizeMiB,
        [ValidateSet('Primary', 'Archive', 'AuxArchive', 'All')][string]$Mailbox = 'All',
        [string[]]$MailboxId,
        [string]$SelectionFile,
        [switch]$SeparateAuxArchives,
        [switch]$SkipCheck,
        [switch]$Force,
        [string]$DisplayName,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )
    Initialize-MbtEngine
    $own = $Store -is [string]
    $st = if ($own) { Open-MbtExportStore -Path $Store -ReadOnly } else { $Store }
    try {
        $o = [MailboxTransfer.PstExportOptions]::new()
        $o.MaxPartBytes = if ($PSBoundParameters.ContainsKey('MaxSizeMiB')) { $MaxSizeMiB * 1MB } else { [long]($MaxSizeGB * 1024) * 1MB }
        $o.MergeAuxIntoArchive = -not $SeparateAuxArchives
        $o.Check = -not $SkipCheck
        $o.Force = [bool]$Force
        if ($DisplayName) { $o.DisplayName = $DisplayName }
        if ($Mailbox -ne 'All') { $o.Kinds = [MailboxTransfer.StoreMailboxKind[]]@([MailboxTransfer.StoreMailboxKind]$Mailbox) }
        if ($MailboxId) { $o.MailboxIds = [string[]]$MailboxId }
        $first = Resolve-MbtPstDestination -Destination $Destination -Store $st -Options $o
        $o.Destination = $first
        if ($SelectionFile) {
            $sel = Read-MbtBrowseSelection -Path $SelectionFile
            $r = [MailboxTransfer.PstExportSelection]::Resolve($st, $sel)
            foreach ($p in $r.Problems) { Write-MbtItem WARN $p }
            if ($r.EntriesMatched -eq 0) { throw "None of the $($sel.Entries.Count) entr(ies) of the selection $SelectionFile is in the store." }
            $o.FolderKeys = $r.FolderKeys
            if ($r.ItemKeys.Count) { $o.ItemKeys = $r.ItemKeys }
        }
        try { $plan = [MailboxTransfer.PstExporter]::Plan($st, $o, $CancellationToken) }
        catch {
            if (-not ($_.Exception -is [OperationCanceledException] -or $_.Exception.InnerException -is [OperationCanceledException])) { throw }
            Write-MbtItem WARN 'export to PST stopped before it began: nothing was written, run the same command again'
            return [pscustomobject]@{ Plan = $null; Result = $null; ExitCode = 2 }
        }
        Show-MbtPstExportPlan -Plan $plan -Destination $first
        $manifestPath = [MailboxTransfer.PstExportManifest]::PathFor($first)
        if (-not $PSCmdlet.ShouldProcess($first, "Write about $($plan.EstimatedParts) PST file(s) from $($plan.Items) item(s) ($(Format-MbtBytes $plan.StoredBytes)); manifest $manifestPath")) {
            return [pscustomobject]@{ Plan = $plan; Result = $null; ExitCode = 0 }
        }
        $last = [Diagnostics.Stopwatch]::StartNew()
        $state = @{ Phase = '' }
        # with an event log open (a job of the console) the folders of the PST are followed: the Run page lists them
        $follow = Start-MbtFolderEvents -Step pst
        if ($follow) { $o.Folders = $follow.Tracker }
        # the sink is a closure: it does not see the private functions of the module, so it gets this one by reference
        $pstEvent = ${function:Write-MbtPstExportEvent}
        $sink = [MailboxTransfer.PstExportProgressSink]::new([Action[MailboxTransfer.PstExportProgress]]({
                param($e)
                if ($e.Phase -ne $state.Phase -or $last.ElapsedMilliseconds -ge 500) {
                    & $pstEvent -Progress $e
                    $state.Phase = $e.Phase; $last.Restart()
                    $pct = if ($e.BytesTotal -gt 0) { [math]::Min(100, [int](100 * $e.BytesDone / $e.BytesTotal)) } else { 0 }
                    Write-Progress -Activity 'Export-MbtPst' -Status ("{0}: part {1}, {2} / {3} items, {4}" -f $e.Phase, $e.PartIndex, $e.ItemsDone, $e.ItemsTotal, $e.Folder) -PercentComplete $pct
                }
            }.GetNewClosure()))
        try { $res = [MailboxTransfer.PstExporter]::Run($st, $o, $sink, $CancellationToken) } finally { Stop-MbtFolderEvents $follow }
        Write-Progress -Activity 'Export-MbtPst' -Completed
        foreach ($n in $res.Notes) { Write-MbtItem INFO $n }
        if ($res.PartsReused -gt 0) { Write-MbtItem INFO ("{0} part(s) of an earlier run verified and kept ({1} item(s)); {2} part(s) written again" -f $res.PartsReused, $res.ItemsReused, $res.PartsRebuilt) }
        foreach ($p in $res.Parts | Where-Object { -not $_.Reused }) {
            $check = if (-not $p.Checked) { 'not checked' } elseif ($p.CheckErrors) { "CHECK: $($p.CheckErrors) error(s), $($p.CheckWarnings) warning(s)" } else { "check clean ($($p.CheckWarnings) warning(s), $(Format-MbtDuration $p.CheckSeconds))" }
            Write-MbtItem $(if ($p.CheckErrors) { 'FAIL' } else { 'OK' }) ("{0}: {1}, {2} item(s), {3} folder(s), written in {4}, {5}" -f $p.File, (Format-MbtBytes $p.Bytes), $p.ItemCount, $p.FolderCount, (Format-MbtDuration $p.Seconds), $check)
        }
        foreach ($ov in $res.Oversize) { Write-MbtItem WARN ("item {0} '{1}' (about {2}) is larger than the PST limit of {3}: written alone in part {4}" -f $ov.ItemKey, $ov.Subject, (Format-MbtBytes $ov.EstimatedBytes), (Format-MbtBytes $o.MaxPartBytes), $ov.Part) }
        foreach ($f in $res.Failed | Select-Object -First 50) { Write-MbtItem WARN ("item {0} '{1}' in {2} is not in the PST: {3}" -f $f.ItemKey, $f.Subject, $f.Folder, $f.Reason) }
        if ($res.Failed.Count -gt 50) { Write-MbtDetail "... $($res.Failed.Count - 50) more" }
        if ($res.Error) { Write-MbtItem FAIL "export to PST stopped: $($res.Error)" }
        if ($res.Cancelled) { Write-MbtItem WARN 'export to PST cancelled: the finished parts are kept, run the same command again to continue' }
        $code = if ($res.Error -or $res.PartsWithErrors) { 1 } elseif ($res.Cancelled -or $res.Failed.Count -or $plan.ItemsNotExported -gt 0) { 2 } else { 0 }
        Write-MbtItem $(if ($code -eq 0) { 'OK' } else { 'WARN' }) ("export to PST {0}: {1} part(s) in total, {2} item(s) written now, {3} reused, {4} in {5}; manifest {6}" -f $(if ($res.Complete -and -not $res.Cancelled -and -not $res.Error) { 'finished' } else { 'NOT finished' }),
            $res.Parts.Count, $res.ItemsWritten, $res.ItemsReused, (Format-MbtBytes $res.BytesWritten), (Format-MbtDuration $res.Elapsed.TotalSeconds), $res.ManifestPath)
        return [pscustomobject]@{ Plan = $plan; Result = $res; ExitCode = $code }
    } finally { if ($own -and $st) { $st.Dispose() } }
}

function Copy-MbtPst {
    <#
    .SYNOPSIS
        Copies chosen folders and messages of a PST (or of the parts of a split PST) into new PST files with our own writer: PST to PST, no Outlook.
    .DESCRIPTION
        -Path          the PST file, several files, or a folder of PST files (split parts in order), only read.
        -IncludeFolder / -ExcludeFolder   the folders to copy: paths below the top folder of the PST ("Inbox", "Inbox/Projects"), wildcards * and ?; a path takes the folder
                       and what is below it, an exclusion wins ("Inbox" with "Inbox/*" excluded is the Inbox alone). Neither: everything.
        -ItemNid       messages chosen one by one (node ids of the PST, one source file), in addition to the folders; with this alone only these messages are copied, in
                       the folders that hold them.
        Layout and rules as Export-MbtPst: the folders keep their names, hierarchy and container class; split at -MaxSizeGB (default 50 = 51,200 MiB) or -MaxSizeMiB on item
        boundaries (name.pst, name-002.pst, ...), each part recreates the folders its items need and is checked with the integrity checker unless -SkipCheck. Folder-associated
        items (rules, views) and search folders are not copied. A message that cannot be read or written is reported and the copy goes on.
        No resume: a copy that is stopped or fails leaves no file. The destination is never overwritten (-Force deletes the files of an earlier copy of the same name).
        -WhatIf prints the plan (folders, items, estimated size and number of files) and writes nothing.
        Returns Plan, Result (PstCopyResult) and ExitCode: 0 complete, 1 error, a part with check errors, or items that could not be copied (the copy is finished), 2 stopped.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'Medium')]
    param(
        [Parameter(Mandatory, Position = 0)][string[]]$Path,
        [Parameter(Mandatory)][string]$Destination,
        [string[]]$IncludeFolder,
        [string[]]$ExcludeFolder,
        [uint32[]]$ItemNid,
        [ValidateRange(1, 1024)][double]$MaxSizeGB = 50,
        [ValidateRange(1, 1048576)][long]$MaxSizeMiB,
        [switch]$SkipCheck,
        [switch]$Force,
        [string]$DisplayName,
        [Threading.CancellationToken]$CancellationToken = [Threading.CancellationToken]::None
    )
    Initialize-MbtEngine
    $o = [MailboxTransfer.PstCopyOptions]::new()
    foreach ($f in $Path) { $o.SourcePaths.Add($f) }
    $o.Destination = $(if ([IO.Path]::IsPathRooted($Destination)) { $Destination } else { Join-Path (Get-Location).Path $Destination })
    # A Windows device name (aux.pst): refused here, with the clear message, before anything resolves the path (on Windows 10 and Server 2022 the name is the device itself).
    [MailboxTransfer.FileNaming]::Require($o.Destination, $true, 'PST destination')
    $o.MaxPartBytes = if ($PSBoundParameters.ContainsKey('MaxSizeMiB')) { $MaxSizeMiB * 1MB } else { [long]($MaxSizeGB * 1024) * 1MB }
    foreach ($f in @($IncludeFolder | Where-Object { $_ })) { $o.IncludeFolders.Add($f) }
    foreach ($f in @($ExcludeFolder | Where-Object { $_ })) { $o.ExcludeFolders.Add($f) }
    if ($ItemNid) { $o.ItemNids = [Collections.Generic.HashSet[uint32]]::new([uint32[]]$ItemNid) }
    $o.Check = -not $SkipCheck; $o.Force = [bool]$Force
    if ($DisplayName) { $o.DisplayName = $DisplayName }
    try { $plan = [MailboxTransfer.PstCopier]::Plan($o, $CancellationToken) }
    catch {
        if (-not ($_.Exception -is [OperationCanceledException] -or $_.Exception.InnerException -is [OperationCanceledException])) { throw }
        Write-MbtItem WARN 'copy to PST stopped before it began: nothing was written, run the same command again'
        return [pscustomobject]@{ Plan = $null; Result = $null; ExitCode = 2 }
    }
    $names = ' (' + ((1..[math]::Max(1, $plan.EstimatedParts) | Select-Object -First 3 | ForEach-Object { Split-Path ([MailboxTransfer.PstExporter]::PartPath($o.Destination, $_)) -Leaf }) -join ', ') + $(if ($plan.EstimatedParts -gt 3) { ', ...' }) + ')'
    Write-MbtItem INFO ("plan: {0} item(s) in {1} folder(s) of {2} (the new files would be about {3}, {4} PST file(s) of at most {5}{6}; the size is estimated from the share of the items)" -f $plan.Items, $plan.Folders.Count, (Format-MbtBytes $plan.SourceBytes), (Format-MbtBytes $plan.EstimatedBytes), $plan.EstimatedParts, (Format-MbtBytes $o.MaxPartBytes), $names)
    foreach ($f in $plan.Folders | Where-Object { $_.Items -gt 0 } | Select-Object -First 60) {
        Write-MbtDetail ("{0,-60} {1,6} item(s)  {2}" -f $(if ($f.Path) { $f.Path } else { '(top)' }), $f.Items, $f.ContainerClass)
    }
    foreach ($f in $plan.Folders | Where-Object { $_.Items -gt 0 } | Select-Object -First 500) { Write-MbtJobObject -Step plan -Key $(if ($f.Path) { $f.Path } else { '(top)' }) -State planned -Total $f.Items }
    foreach ($n in $plan.Notes) { Write-MbtItem WARN $n }
    if (-not $PSCmdlet.ShouldProcess($o.Destination, "Write about $($plan.EstimatedParts) PST file(s) from $($plan.Items) item(s)")) {
        return [pscustomobject]@{ Plan = $plan; Result = $null; ExitCode = 0 }
    }
    $last = [Diagnostics.Stopwatch]::StartNew()
    $state = @{ Phase = '' }
    $follow = Start-MbtFolderEvents -Step pst
    if ($follow) { $o.Folders = $follow.Tracker }
    $pstEvent = ${function:Write-MbtPstExportEvent}
    $sink = [MailboxTransfer.PstExportProgressSink]::new([Action[MailboxTransfer.PstExportProgress]]({
            param($e)
            if ($e.Phase -ne $state.Phase -or $last.ElapsedMilliseconds -ge 500) {
                & $pstEvent -Progress $e
                $state.Phase = $e.Phase; $last.Restart()
                $pct = if ($e.ItemsTotal -gt 0) { [math]::Min(100, [int](100 * $e.ItemsDone / $e.ItemsTotal)) } else { 0 }
                Write-Progress -Activity 'Copy-MbtPst' -Status ("{0}: part {1}, {2} / {3} items, {4}" -f $e.Phase, $e.PartIndex, $e.ItemsDone, $e.ItemsTotal, $e.Folder) -PercentComplete $pct
            }
        }.GetNewClosure()))
    try { $res = [MailboxTransfer.PstCopier]::Run($o, $sink, $CancellationToken) } finally { Stop-MbtFolderEvents $follow }
    Write-Progress -Activity 'Copy-MbtPst' -Completed
    foreach ($n in $res.Notes) { Write-MbtItem INFO $n }
    foreach ($p in $res.Parts) {
        $check = if (-not $p.Checked) { 'not checked' } elseif ($p.CheckErrors) { "CHECK: $($p.CheckErrors) error(s), $($p.CheckWarnings) warning(s)" } else { "check clean ($($p.CheckWarnings) warning(s), $(Format-MbtDuration $p.CheckSeconds))" }
        Write-MbtItem $(if ($p.CheckErrors) { 'FAIL' } else { 'OK' }) ("{0}: {1}, {2} item(s), {3} folder(s), written in {4}, {5}" -f $p.File, (Format-MbtBytes $p.Bytes), $p.ItemCount, $p.FolderCount, (Format-MbtDuration $p.Seconds), $check)
    }
    foreach ($ov in $res.Oversize) { Write-MbtItem WARN ("message {0} '{1}' (about {2}) is larger than the PST limit of {3}: written alone in part {4}" -f $ov.ItemKey, $ov.Subject, (Format-MbtBytes $ov.EstimatedBytes), (Format-MbtBytes $o.MaxPartBytes), $ov.Part) }
    foreach ($f in $res.Failed | Select-Object -First 50) { Write-MbtItem WARN ("message {0} '{1}' in {2} is not in the new PST: {3}" -f $f.ItemKey, $f.Subject, $f.Folder, $f.Reason) }
    if ($res.Failed.Count -gt 50) { Write-MbtDetail "... $($res.Failed.Count - 50) more" }
    if ($res.Error) { Write-MbtItem FAIL "copy to PST stopped: $($res.Error)" }
    if ($res.Cancelled) { Write-MbtItem WARN 'copy to PST stopped on request: no file was kept, run the same command again to start it again' }
    $code = if ($res.Error -or $res.PartsWithErrors -or ($res.Failed.Count -and -not $res.Cancelled)) { 1 } elseif ($res.Cancelled) { 2 } else { 0 }
    Write-MbtItem $(if ($code -eq 0) { 'OK' } else { 'WARN' }) ("copy to PST {0}: {1} file(s), {2} message(s) written, {3} not written, {4} in {5}" -f $(if ($res.Ok) { 'finished' } else { 'NOT clean' }), $res.Parts.Count, $res.ItemsWritten, $res.Failed.Count, (Format-MbtBytes $res.BytesWritten), (Format-MbtDuration $res.Elapsed.TotalSeconds))
    return [pscustomobject]@{ Plan = $plan; Result = $res; ExitCode = $code }
}

function Get-MbtPstPreview {
    <#
    .SYNOPSIS
        The preview of one message of a PST, as the reading pane shows it: header, text body (and HTML body when there is one), attachments (name, type, size). Read only.
    .DESCRIPTION
        -Nid is the node id of the message (PstReader.MessageNids lists them). The body is cut at -BodyChars characters. Returns a PstMessagePreview.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][uint32]$Nid, [ValidateRange(0, 1000000)][int]$BodyChars = 20000)
    Initialize-MbtEngine
    $r = [MailboxTransfer.PstReader]::Open((Resolve-Path -LiteralPath $Path).Path)
    try { return $r.GetPreview($Nid, $BodyChars) } finally { $r.Dispose() }
}

#endregion

# EML export / import / compare (functions of the module, in a file of their own)
. (Join-Path $PSScriptRoot 'module\MailboxTransfer.Eml.ps1')
# Events of a job (JSON lines), stop request, keep awake, job status, capabilities.
. (Join-Path $PSScriptRoot 'module\MailboxTransfer.Events.ps1')

# An engine left in this process by an older build of the module is refused here, before any network or store operation.
[void](Assert-MbtEngineCurrent -Name Engine -Root $script:ToolRoot)
[void](Assert-MbtEngineCurrent -Name Auth -Root $script:ToolRoot)

Export-ModuleMember -Function @(
    'Write-MbtStep', 'Write-MbtItem', 'Write-MbtDetail', 'Write-MbtBanner', 'Write-MbtSummary', 'Write-MbtLog', 'Start-MbtLog', 'Stop-MbtLog', 'Get-MbtLogPath', 'Format-MbtBytes', 'Format-MbtNumber', 'Format-MbtDuration', 'Format-MbtRate',
    'Import-MbtConfiguration', 'Initialize-MbtEngine', 'Initialize-MbtAuthEngine', 'Find-MbtMsal',
    'Connect-MbtGraph', 'Get-MbtAccessToken', 'Disconnect-MbtGraph', 'Get-MbtSession', 'Set-MbtSession', 'Get-MbtTokenInfo', 'Get-MbtCertificate', 'Get-MbtAppPermission', 'Get-MbtMissingAppPermission', 'Test-MbtAppOnly', 'Get-MbtSessionUser', 'Get-MbtExchangeSettingsPath',
    'Test-MbtMailboxId', 'ConvertTo-MbtMailboxId', 'Get-MbtPrimaryMailboxId', 'Test-MbtMailboxAddress', 'Resolve-MbtMailboxAddress', 'Get-MbtAccessHint', 'Get-MbtMailboxLocation', 'Read-MbtMailboxLocationFile', 'Get-MbtGraphException', 'Get-MbtMailboxTree', 'Resolve-MbtFolder', 'Get-MbtFolderItems', 'New-MbtFolderPath',
    'Export-MbtItems', 'Open-MbtExportStore', 'Save-MbtFolderTree', 'Save-MbtItemList', 'Export-MbtItemsToStore', 'Get-MbtStoreItems', 'Import-MbtStoreItem',
    'Read-MbtFxStream', 'Show-MbtFxReport', 'Import-MbtItem'
    'Resolve-MbtMailboxId', 'Wait-MbtTransferJob', 'Export-MbtMailbox', 'Import-MbtMapping', 'Import-MbtMailbox', 'Import-MbtPst', 'Export-MbtPstImportReport', 'Show-MbtPstImportPlan', 'Get-MbtImportUnknown', 'Resolve-MbtImportUnknown', 'Get-MbtStoreStatus', 'Show-MbtStoreStatus'
    'New-MbtBrowseLocation', 'Get-MbtBrowseLocation', 'Get-MbtBrowser', 'Get-MbtBrowseFolder', 'Find-MbtBrowseFolder', 'Get-MbtBrowseItem', 'Get-MbtBrowsePreview'
    'New-MbtBrowseSelection', 'Add-MbtBrowseSelection', 'Save-MbtBrowseSelection', 'Read-MbtBrowseSelection', 'Resolve-MbtBrowsePlan', 'Add-MbtBrowseMoveInfo', 'Format-MbtBrowseTotals'
    'Show-MbtBrowseTree', 'Show-MbtBrowseItems', 'Show-MbtBrowsePreview', 'Show-MbtExportPlan'
    'Test-MbtPst', 'Repair-MbtPst', 'Export-MbtPst', 'Copy-MbtPst', 'Get-MbtPstPreview', 'Show-MbtPstExportPlan', 'Resolve-MbtPstDestination', 'Show-MbtLiveExportPlan',
    'Export-MbtEml', 'Import-MbtEml', 'Resolve-MbtEmlUnknown', 'Compare-MbtEml',
    'Open-MbtEventLog', 'Close-MbtEventLog', 'Write-MbtEvent', 'Write-MbtJobPlan', 'Write-MbtJobStep', 'Write-MbtJobObject', 'Write-MbtJobResult', 'Write-MbtJobError', 'ConvertTo-MbtPlainError', 'Write-MbtJobStatus', 'Get-MbtJob'
    'Start-MbtStopWatch', 'Stop-MbtStopWatch', 'Test-MbtStopRequested', 'Get-MbtStopToken', 'Start-MbtKeepAwake', 'Stop-MbtKeepAwake', 'Get-MbtCapability'
    'Get-MbtConfiguredMailboxId', 'Assert-MbtSignInSettings'
)

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCcY/EBLuzeYFqX
# y0AQQ059VBTqYJcu2O6aELD2AbB7jKCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYEDCCGAwCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCzRrvz9Zpm
# V3yF1R/nXTjzI9AhFodczgOpdyKEQEOSBzANBgkqhkiG9w0BAQEFAASCAgCnXA03
# kELHnz65licz6akT1WwODHpS4GwyE8+eqisFD0ELqTPKNgiFpiTIArBfMq8u4Ok2
# RKHL4cUZWzxzh9I+qw1d9VR0bI4LYZ9amO9VN/vp9c+7thqoJhL2/W8XsuWBbCpt
# 2uwdJ7MopOLtD4RiB6nWR1pafgHCBoqTbFtyatRHwrTR30cY6pVto9ZU2oxhtvpW
# fwR+d8l7ntt5Ovnk1xB2sL5ddu8nvasQndfqYKMNqm4T3lc3jiQu4ziGlWOhGzLO
# SiAp5QfsQfKtDeXy49czMdug+r5Wc11ipHIRpCwP/qaft+4nyj1nxyUBPxr3Rh5j
# ObH6iQ9R6iFILDMwZGcyaLKIaL7y84ZJkEffbd8cDoZuD3+PxQ3qJYQCjtnTdeHQ
# Oab6Eif/jivjnkYWtDIT1m1hmBCBcvHV2QhbaNOvaNAKLofT+iYaQGkzmvUHdVor
# kWJqgpL3xHL9QQuMwVqww5rWdKH85GQkCbF5+jjwPJgg7ploIWbc3XHzciXbl3CB
# d6VMbXQ0AgHjd989Exum0sEJMs2i54XtI2BPOs1UvBfen0FUngYehXi8cUXSTirZ
# uOEB7f6ujkWjcEwie13YjzBKXwvszk6/HmjU3hzj/Xi84eerAe8DG1zxG3Cf+Jzm
# EGeUJlaVcy/Liksd7EZSh2se+IuaqqUfESm3oqGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIKW6D2KV0weKFuI77lsCVu4zsP+B7VtaiiSDrVU79pflAggfsxEK9BPf
# TBgPMjAyNjExMTAwOTE3MjJaMAMCAQECEAN9FfMWi1G6lpPCD8m/qPigghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlUwggJRAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA5MTcyMlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIDY/
# jOWgfMffRqESLb48z/zFNVPMhDhoJwJ3zNAdx/B+MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiB8CXUZaZrKCcd46EuByqnAAG62cZmL1dTeAFzOmTIB5QIgQN09
# vOfjF0bpmM1ZN0KW7/7x6tYvbvmx0BQDCQbWPyo=
# SIG # End signature block
