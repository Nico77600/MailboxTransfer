#Requires -Version 7.4
<#
.SYNOPSIS
    Renders the pages of the console from the real WPF window (synthetic data, fake engine) into PNG files, light and dark.

.DESCRIPTION
    Builds the window with the fake engine, puts it in each scenario (package\module\MailboxTransfer.Gui.Scenarios.ps1), draws it into a bitmap at
    1480 x 940 and saves out\console-screens\<theme>\<NN>-<scenario>.png. Nothing is read or written outside the output folder. Runs in its own
    single-threaded apartment (it starts itself again when needed). The window is shown without being visible (opacity 0) so that WPF builds
    every template, then closed.

.PARAMETER Scenario
    Scenario names (default: all of them).

.PARAMETER Theme
    Light, Dark (default: both).

.PARAMETER Version
    The version shown in the corner of the window (default: ModuleVersion of the manifest).

.PARAMETER OutDir
    Default: <repository>\out\console-screens.
#>
[CmdletBinding()]
param([string[]]$Scenario, [string[]]$Theme = @('Light', 'Dark'), [string]$OutDir, [int]$Width = 1480, [int]$Height = 940, [string]$Version)

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $a = @('-NoProfile', '-STA', '-File', $PSCommandPath)
    foreach ($k in $PSBoundParameters.Keys) { $v = $PSBoundParameters[$k]; $a += "-$k"; $a += (@($v) -join ',') }
    & (Get-Process -Id $PID).Path @a
    exit $LASTEXITCODE
}
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$package = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
if (-not $Version) { $Version = (Import-PowerShellDataFile -LiteralPath (Join-Path $package 'MailboxTransfer.psd1')).ModuleVersion }
$Scenario = @($Scenario | ForEach-Object { $_ -split ',' } | Where-Object { $_ }); $Theme = @($Theme | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
if (-not $OutDir) { $OutDir = Join-Path (Split-Path $PSScriptRoot -Parent) 'out\console-screens' }
foreach ($f in 'Core', 'Run', 'Jobs', 'Browse', 'State', 'Pages', 'Pages.Export', 'Pages.Run', 'Pages.Source', 'Fake', 'Engine', 'Scenarios') { . (Join-Path $package "module\MailboxTransfer.Gui.$f.ps1") }
$script:GuiRoot = $package
if (-not $Scenario) { $Scenario = Get-MbtGuiScenarioName }
$stateOnly = $null

function Wait-Dispatcher {
    for ($i = 0; $i -lt 4; $i++) {
        $frame = [Windows.Threading.DispatcherFrame]::new()
        [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::ContextIdle, [Action] { $frame.Continue = $false })
        [Windows.Threading.Dispatcher]::PushFrame($frame)
    }
}

foreach ($t in $Theme) {
    $dir = Join-Path $OutDir $t.ToLowerInvariant()
    $null = New-Item -ItemType Directory -Force -Path $dir
    $engine = New-MbtGuiFakeEngine -Capability @{ PstImport = $true; EmlExport = $true; EmlImport = $true; SharedMailbox = $true; FolderChoice = $true; PstCopy = $true; PstPreview = $true }
    $window = New-MbtGuiWindow -Engine $engine -Theme $t -WorkFolder 'D:\MailboxTransfer' -EntryScript (Join-Path $package 'module\MailboxTransfer.Gui.DemoJob.ps1') -Version $Version
    $window.WindowStartupLocation = 'Manual'; $window.Left = 0; $window.Top = 0
    $window.Width = $Width + 16; $window.Height = $Height + 39; $window.MinWidth = 0; $window.MinHeight = 0
    $window.Opacity = 0; $window.ShowInTaskbar = $false; $window.ShowActivated = $false
    $window.Show()
    $n = 0
    foreach ($name in $Scenario) {
        $n++
        $g = $script:Gui
        # a clean state for each scenario (the window and its browsers live on)
        $g.S = New-MbtGuiState -WorkFolder 'D:\MailboxTransfer'; $g.Run = $null; $g.Job = $null
        foreach ($b in $g.Browsers.Values) { Set-MbtGuiBrowserRoots -Browser $b -Roots @() }
        Set-MbtGuiScenario -Name $name
        Wait-Dispatcher
        $root = $window.Content
        $root.UpdateLayout()
        $w = [int]$root.ActualWidth; $h = [int]$root.ActualHeight
        $bmp = [Windows.Media.Imaging.RenderTargetBitmap]::new($w, $h, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
        $bmp.Render($root)
        $enc = [Windows.Media.Imaging.PngBitmapEncoder]::new()
        $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bmp))
        $file = Join-Path $dir ('{0:00}-{1}.png' -f $n, $name)
        $fs = [IO.File]::Create($file); try { $enc.Save($fs) } finally { $fs.Dispose() }
        Write-Host "$t $name -> $file ($w x $h)"
    }
    $g.ForceClose = $true
    $window.Close()
}
