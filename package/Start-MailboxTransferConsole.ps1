#Requires -Version 7.4

<#
.SYNOPSIS
    Mailbox Transfer console - the window that guides an operator through browse and export, import, check and repair of a PST, and open.

.DESCRIPTION
    WPF window (Fluent theme of .NET 9 and later; light or dark like Windows). Task, From, To, Plan, Run, Done: one question per page, safe
    defaults, the plan read before anything starts, the long jobs run by Invoke-MailboxTransfer.ps1 in a child process and followed through
    their events file. Sign-in, browse and preview run in this process. See README, section "Console".

    -Demo runs the whole window on synthetic data with a fake engine and a job that only writes events: no network, no mailbox, no file outside
    the work folder. -DocScenario <name> opens one page in one state of that demo (for the documentation images). -Theme Light or Dark is for the
    images too; the window follows Windows.

.PARAMETER WorkFolder
    Where jobs and stores live (default %LOCALAPPDATA%\MailboxTransfer).

.PARAMETER Theme
    System (default, like Windows), Light or Dark.

.PARAMETER Demo
    Synthetic data, fake engine, a demonstration job.

.PARAMETER DocScenario
    One state of one page, on synthetic data (implies -Demo): see Get-MbtGuiScenarioName.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.1.0 (console)
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.
#>
[CmdletBinding()]
param(
    [string]$WorkFolder,
    [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System',
    [switch]$Demo,
    [string]$DocScenario,
    [switch]$DevelopmentMode
)

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    # WPF needs a single-threaded apartment: start again in one.
    $arguments = @('-NoProfile', '-STA', '-File', $PSCommandPath)
    foreach ($k in $PSBoundParameters.Keys) {
        $v = $PSBoundParameters[$k]
        if ($v -is [switch]) { if ($v) { $arguments += "-$k" } } else { $arguments += "-$k"; $arguments += [string]$v }
    }
    & (Get-Process -Id $PID).Path @arguments
    exit $LASTEXITCODE
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# The package is verified before any of its code runs: the verifier must carry the owner signature, then it checks the signed manifest.
# A tree without a manifest (a source checkout) is accepted only with -DevelopmentMode.
$integrityManifest = Join-Path $PSScriptRoot 'PackageIntegrity.psd1'
if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $integrityManifest) {
    $integrityVerifier = Join-Path $PSScriptRoot 'src\Test-PackageIntegrity.ps1'
    $integrityCursor = [IO.Path]::GetFullPath($integrityVerifier)
    while ($integrityCursor) {
        $integrityItem = Microsoft.PowerShell.Management\Get-Item -LiteralPath $integrityCursor -Force -ErrorAction Stop
        if ($integrityItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $integrityCursor" }
        $integrityCursor = [IO.Path]::GetDirectoryName($integrityCursor)
    }
    $integritySignature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $integrityVerifier -ErrorAction Stop
    if ([string]$integritySignature.Status -cne 'Valid' -or $null -eq $integritySignature.SignerCertificate -or
        $integritySignature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
        $null -eq $integritySignature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature. Use the complete release ZIP (https://github.com/Nico77600/MailboxTransfer/releases/latest): extract it again and unblock it. A source checkout is started with -DevelopmentMode.'
    }
    if ($DevelopmentMode) { & $integrityVerifier -Root $PSScriptRoot -DevelopmentMode | Out-Null } else { & $integrityVerifier -Root $PSScriptRoot -Distribution | Out-Null }
} elseif (-not $DevelopmentMode) {
    throw 'Package integrity: PackageIntegrity.psd1 is missing. Use the complete release ZIP (https://github.com/Nico77600/MailboxTransfer/releases/latest): extract it again and unblock it. A source checkout is started with -DevelopmentMode.'
} else {
    Write-Warning 'DevelopmentMode: no package manifest; package byte integrity and publisher authentication are not established.'
    # The module does not repeat this warning in the same process.
    [AppDomain]::CurrentDomain.SetData('MailboxTransfer.UnverifiedTreeWarned', $true)
}
$script:ConsoleVersion = '1.1.0'
$packageRoot = $PSScriptRoot
foreach ($f in 'Core', 'Run', 'Jobs', 'Browse', 'State', 'Pages', 'Pages.Export', 'Pages.Run', 'Pages.Source', 'Fake', 'Engine', 'Scenarios') {
    . (Join-Path $packageRoot "module\MailboxTransfer.Gui.$f.ps1")
}
$script:GuiRoot = $packageRoot
$script:GuiDevelopmentMode = [bool]$DevelopmentMode

if ($DocScenario) { $Demo = $true }
$entry = Join-Path $packageRoot 'Invoke-MailboxTransfer.ps1'
if (-not $WorkFolder) { $WorkFolder = Join-Path $env:LOCALAPPDATA 'MailboxTransfer' }
Initialize-MbtGuiTypes
if ($Demo) {
    $engine = New-MbtGuiFakeEngine -Capability @{ PstImport = $true; EmlExport = $true; EmlImport = $true; SharedMailbox = $true; FolderChoice = $true; PstCopy = $true; PstPreview = $true }
    $demoJob = Join-Path $packageRoot 'module\MailboxTransfer.Gui.DemoJob.ps1'
    $WorkFolder = Join-Path ([IO.Path]::GetTempPath()) 'MailboxTransferConsoleDemo'
    $window = New-MbtGuiWindow -Engine $engine -Theme $Theme -WorkFolder $WorkFolder -EntryScript $demoJob -Version $script:ConsoleVersion
    if ($DocScenario) { Set-MbtGuiScenario -Name $DocScenario }
} else {
    $engine = New-MbtGuiRealEngine -PackageRoot $packageRoot -EntryScript $entry -WorkFolder $WorkFolder
    $window = New-MbtGuiWindow -Engine $engine -Theme $Theme -WorkFolder $WorkFolder -EntryScript $entry -Version $script:ConsoleVersion
}
[void]$window.ShowDialog()
