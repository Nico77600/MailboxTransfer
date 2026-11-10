#Requires -Version 7.4
<#
.SYNOPSIS
    Regenerates the pictures of the documentation (package\docs\images\*.png), light and dark, from synthetic data only.

.DESCRIPTION
    No mailbox, no network and no Outlook are needed. Two kinds of pictures are made:

      1. Window pictures, console-*-<theme>.png: the real WPF window of the console on the fake engine (tools\New-ConsoleScreens.ps1,
         package\module\MailboxTransfer.Gui.Scenarios.ps1), contoso.com names only, with numbered red callouts drawn on the key ones.
      2. Drawings, readme-*-<theme>.png and guide-*-<theme>.png: small diagrams drawn with the same WPF classes and the same colours as the
         window (the path, the principles, the permission matrix, the glossary, ...). They are rendered at twice their size.
      3. The README banner, readme-banner-<theme>.svg: an SVG with the template of the ExchangeLogReport hero (same canvas, frame, accent bar, tint,
         icon, eyebrow, title, pills, footer line and stat cards).

    Naming follows the house standard: <area>-<page or subject>-<variant>-<theme>.png. The README shows the pictures with <picture> so that the
    light or the dark one is chosen by the reader's theme.

    Privacy: every text of a drawing is checked before it is saved. A text that holds a GUID, a tenant name other than contoso, or a term of the
    optional private deny-list (tests\privacy-denylist.local.txt, git-ignored) stops the script. The window pictures come from the fake engine,
    which holds contoso values only.

.PARAMETER Images
    All (default), Console (the window pictures only) or Diagrams (the drawings only).

.PARAMETER OutDir
    Default: package\docs\images.

.PARAMETER Version
    The version drawn in the corner of the window pictures (default: ModuleVersion of the manifest).

.PARAMETER Name
    Only the pictures whose file name starts with one of these words (for instance readme-path, console-run).

.EXAMPLE
    .\tools\New-DocumentationImages.ps1
    .\tools\New-DocumentationImages.ps1 -Images Diagrams -Name readme-permissions

.NOTES
    Author  : Nicolas Fabert
    Part of : Mailbox Transfer (repository tool, not in the package)
    Needs   : Windows with WPF, PowerShell 7.4 or later; Segoe UI Variable and Segoe Fluent Icons (Windows 11) for the exact look.
#>
[CmdletBinding()]
param([ValidateSet('All', 'Console', 'Diagrams')][string]$Images = 'All', [string]$OutDir, [string]$Version, [string[]]$Name)

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
    $a = @('-NoProfile', '-STA', '-File', $PSCommandPath)
    foreach ($k in $PSBoundParameters.Keys) { $a += "-$k"; $a += (@($PSBoundParameters[$k]) -join ',') }
    & (Get-Process -Id $PID).Path @a
    exit $LASTEXITCODE
}
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
$package = Join-Path $repo 'package'
if (-not $OutDir) { $OutDir = Join-Path $package 'docs\images' }
$null = New-Item -ItemType Directory -Force -Path $OutDir
$Name = @($Name | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
if (-not $Version) { $Version = (Import-PowerShellDataFile -LiteralPath (Join-Path $package 'MailboxTransfer.psd1')).ModuleVersion }
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase

function Test-Wanted([string]$File) { (-not $Name) -or [bool]($Name | Where-Object { $File -like "$_*" }) }

# ---- privacy check ------------------------------------------------------------------------------------------------------------------------
$script:Deny = @()
$denyFile = Join-Path $repo 'tests\privacy-denylist.local.txt'
if (Test-Path -LiteralPath $denyFile) { $script:Deny = @(Get-Content -LiteralPath $denyFile | Where-Object { $_.Trim() -and -not $_.StartsWith('#') } | ForEach-Object { $_.Trim() }) }
function Assert-PublicText([string]$Text) {
    if ($Text -match '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}') { throw 'A picture text holds a GUID.' }
    if ($Text -match '(?i)\b(?!contoso\b)[a-z0-9-]+\.onmicrosoft\.com') { throw 'A picture text names a tenant other than contoso.' }
    foreach ($t in $script:Deny) { if ($Text.IndexOf($t, [StringComparison]::OrdinalIgnoreCase) -ge 0) { throw 'A picture text holds a term of the private deny-list.' } }
}

# ---- 1. window pictures -----------------------------------------------------------------------------------------------------------------------
$screens = [ordered]@{
    Start = 'console-task'; StartImport = 'console-task-import'; ExportFromSignedOut = 'console-signin'; ExportFromCode = 'console-signin-code'
    ExportFrom = 'console-from-browse'; ExportTo = 'console-to-pst'; ExportPlan = 'console-plan-export'
    RunExport = 'console-run-export'; RunImport = 'console-run-import'; RunImportSleep = 'console-run-import-sleep'
    RunImportDecision = 'console-run-import-decision'; RunImportStopped = 'console-run-import-stopped'; RunRepair = 'console-run-repair'
    DoneExportComplete = 'console-done-export-complete'; DoneExportPartial = 'console-done-export-partial'; DoneExportError = 'console-done-export-error'
    ImportFrom = 'console-import-from'; ImportTo = 'console-import-to'; ImportPlan = 'console-import-plan'
    DoneImportComplete = 'console-done-import-complete'; DoneImportPartial = 'console-done-import-partial'
    DoneImportDecision = 'console-done-import-decision'; DoneImportError = 'console-done-import-error'
    RepairFromDamaged = 'console-repair-from-damaged'; RepairFromClean = 'console-repair-from-clean'; RepairTo = 'console-repair-to'
    RepairPlan = 'console-repair-plan'; DoneRepair = 'console-repair-done'; Open = 'console-open'
}
# Numbered callouts (centre x, y on the 1480 x 940 picture). The numbers are explained in the caption of the picture in the guide.
$callouts = @{
    Start = @(@(1, 50, 116), @(2, 16, 195), @(3, 16, 352), @(4, 868, 352))
    ExportFromSignedOut = @(@(1, 16, 176), @(2, 16, 244))
    ExportFromCode = @(@(1, 16, 176), @(2, 16, 330))
    ExportFrom = @(@(1, 16, 176), @(2, 16, 400), @(3, 452, 204), @(4, 1104, 204), @(5, 16, 586))
    ExportTo = @(@(1, 16, 176), @(2, 440, 176), @(3, 1252, 913))
    ExportPlan = @(@(1, 16, 176), @(2, 440, 176), @(3, 1252, 913))
    RunExport = @(@(1, 16, 178), @(2, 16, 334), @(3, 290, 334), @(4, 1146, 334), @(5, 16, 913))
    RunImportDecision = @(@(1, 16, 330), @(2, 30, 381), @(3, 16, 913))
    RepairFromDamaged = @(@(1, 16, 215), @(2, 440, 215), @(3, 1252, 913))
    ImportFrom = @(@(1, 16, 215), @(2, 440, 215))
    ImportPlan = @(@(1, 16, 215), @(2, 440, 215), @(3, 1252, 913))
    DoneRepair = @(,@(1, 16, 178))
}

function Add-Callouts([string]$File, $Points) {
    $img = [Windows.Media.Imaging.BitmapImage]::new(); $img.BeginInit(); $img.UriSource = [Uri]::new($File); $img.CacheOption = 'OnLoad'; $img.EndInit()
    $dv = [Windows.Media.DrawingVisual]::new()
    $dc = $dv.RenderOpen()
    $dc.DrawImage($img, [Windows.Rect]::new(0, 0, $img.PixelWidth, $img.PixelHeight))
    $red = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(0xE5, 0x32, 0x2D)); $white = [Windows.Media.Brushes]::White
    $pen = [Windows.Media.Pen]::new($white, 2.5)
    foreach ($p in $Points) {
        $dc.DrawEllipse($red, $pen, [Windows.Point]::new($p[1], $p[2]), 14, 14)
        $ft = [Windows.Media.FormattedText]::new([string]$p[0], [Globalization.CultureInfo]::InvariantCulture, 'LeftToRight', [Windows.Media.Typeface]::new('Segoe UI Variable Text, Segoe UI'), 16, $white, 1.0)
        $ft.SetFontWeight([Windows.FontWeights]::Bold)
        $dc.DrawText($ft, [Windows.Point]::new($p[1] - $ft.Width / 2, $p[2] - $ft.Height / 2))
    }
    $dc.Close()
    $bmp = [Windows.Media.Imaging.RenderTargetBitmap]::new($img.PixelWidth, $img.PixelHeight, 96, 96, [Windows.Media.PixelFormats]::Pbgra32)
    $bmp.Render($dv)
    $enc = [Windows.Media.Imaging.PngBitmapEncoder]::new(); $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bmp))
    $fs = [IO.File]::Create($File); try { $enc.Save($fs) } finally { $fs.Dispose() }
}

if ($Images -in 'All', 'Console') {
    $work = Join-Path ([IO.Path]::GetTempPath()) ('mbt-doc-images-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    try {
        $want = @($screens.Keys | Where-Object { (Test-Wanted ($screens[$_] + '-light')) })
        if ($want) {
            & (Get-Process -Id $PID).Path -NoProfile -STA -File (Join-Path $PSScriptRoot 'New-ConsoleScreens.ps1') -OutDir $work -Version $Version -Scenario ($want -join ',') | Out-Null
            if ($LASTEXITCODE -ne 0) { throw 'New-ConsoleScreens.ps1 failed.' }
            foreach ($theme in 'light', 'dark') {
                $files = Get-ChildItem -LiteralPath (Join-Path $work $theme) -Filter '*.png'
                foreach ($f in $files) {
                    $scenarioName = $f.BaseName -replace '^\d+-', ''
                    $target = Join-Path $OutDir ('{0}-{1}.png' -f $screens[$scenarioName], $theme)
                    Copy-Item -LiteralPath $f.FullName -Destination $target -Force
                    if ($callouts.ContainsKey($scenarioName)) { Add-Callouts $target $callouts[$scenarioName] }
                    Write-Host "window  $([IO.Path]::GetFileName($target))"
                }
            }
        }
    } finally { if (Test-Path -LiteralPath $work) { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue } }
}
if ($Images -eq 'Console') { return }

# ---- 2. drawings: a small toolkit over WPF -----------------------------------------------------------------------------------------------------
$palettes = @{
    light = @{ bg = '#FFFFFF'; card = '#FFFFFF'; stroke = '#DEDEDE'; text = '#242424'; text2 = '#5C5C5C'; text3 = '#8A8A8A'; accent = '#B11F4B'; accentText = '#9A1A41'; soft = '#F7E4EA'
        onAccent = '#FFFFFF'; ok = '#15803D'; okBg = '#EAF7EE'; warn = '#92400E'; warnBg = '#FFF7E8'; bad = '#B91C1C'; badBg = '#FDECEC'; info = '#F3F3F3'; white = '#FFFFFF'; bar = '#B11F4B'; barSoft = '#E7B9C7' }
    dark  = @{ bg = '#0D1117'; card = '#2B2B2B'; stroke = '#3D3D3D'; text = '#FFFFFF'; text2 = '#C5C5C5'; text3 = '#9A9A9A'; accent = '#FD8EA1'; accentText = '#FD8EA1'; soft = '#40303A'
        onAccent = '#0D1117'; ok = '#4ADE80'; okBg = '#24362B'; warn = '#FBBF24'; warnBg = '#3A3322'; bad = '#F87171'; badBg = '#3A2626'; info = '#2E2E2E'; white = '#FFFFFF'; bar = '#FD8EA1'; barSoft = '#6B4450' }
}
$script:Canvas = $null; $script:Pal = $null; $script:Brush = @{}
function Get-B([string]$Key) {
    if (-not $script:Brush.ContainsKey($Key)) {
        $hex = if ($script:Pal.ContainsKey($Key)) { $script:Pal[$Key] } else { $Key }
        $b = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color][Windows.Media.ColorConverter]::ConvertFromString($hex)); $b.Freeze(); $script:Brush[$Key] = $b
    }
    $script:Brush[$Key]
}
function Set-Pos($El, [double]$X, [double]$Y) { [Windows.Controls.Canvas]::SetLeft($El, $X); [Windows.Controls.Canvas]::SetTop($El, $Y); [void]$script:Canvas.Children.Add($El) }
function Add-Rect([double]$X, [double]$Y, [double]$W, [double]$H, [string]$Fill = 'card', [string]$Stroke = 'stroke', [double]$Radius = 12, [double]$Thickness = 1, [switch]$Dash) {
    $r = [Windows.Shapes.Rectangle]::new(); $r.Width = $W; $r.Height = $H; $r.RadiusX = $Radius; $r.RadiusY = $Radius
    if ($Fill) { $r.Fill = Get-B $Fill }
    if ($Stroke) { $r.Stroke = Get-B $Stroke; $r.StrokeThickness = $Thickness }
    if ($Dash) { $r.StrokeDashArray = [Windows.Media.DoubleCollection]::new([double[]]@(4, 3)) }
    Set-Pos $r $X $Y
}
function Add-Circle([double]$X, [double]$Y, [double]$D, [string]$Fill = 'accent', [string]$Stroke = '') {
    $e = [Windows.Shapes.Ellipse]::new(); $e.Width = $D; $e.Height = $D
    if ($Fill) { $e.Fill = Get-B $Fill }
    if ($Stroke) { $e.Stroke = Get-B $Stroke; $e.StrokeThickness = 1.5 }
    Set-Pos $e $X $Y
}
function Add-Text([string]$Text, [double]$X, [double]$Y, [double]$Size = 16, [string]$Color = 'text', [string]$Weight = 'Normal', [double]$W = 0, [string]$Align = 'Left') {
    Assert-PublicText $Text
    $t = [Windows.Controls.TextBlock]::new(); $t.Text = $Text; $t.FontSize = $Size; $t.Foreground = Get-B $Color
    $t.FontFamily = [Windows.Media.FontFamily]::new('Segoe UI Variable Text, Segoe UI')
    $t.FontWeight = [Windows.FontWeights]::$Weight
    if ($W -gt 0) { $t.Width = $W; $t.TextWrapping = 'Wrap'; $t.TextAlignment = $Align }
    Set-Pos $t $X $Y
    $t.Measure([Windows.Size]::new(10000, 10000))
    return $t.DesiredSize
}
function Add-Icon([int]$Code, [double]$X, [double]$Y, [double]$Size = 24, [string]$Color = 'accent') {
    $t = [Windows.Controls.TextBlock]::new(); $t.Text = [string][char]$Code; $t.FontSize = $Size; $t.Foreground = Get-B $Color
    $t.FontFamily = [Windows.Media.FontFamily]::new('Segoe Fluent Icons, Segoe MDL2 Assets')
    Set-Pos $t $X $Y
}
function Add-Tile([int]$Code, [double]$X, [double]$Y, [double]$D = 56, [string]$Fill = 'soft', [string]$Color = 'accent') {
    Add-Rect $X $Y $D $D $Fill '' ($D / 4)
    Add-Icon $Code ($X + $D * 0.24) ($Y + $D * 0.2) ($D * 0.5) $Color
}
function Add-Pill([string]$Text, [double]$X, [double]$Y, [string]$Fill = 'okBg', [string]$Ink = 'ok', [double]$Size = 14) {
    $m = [Windows.Controls.TextBlock]::new(); $m.Text = $Text; $m.FontSize = $Size; $m.FontWeight = [Windows.FontWeights]::SemiBold
    $m.FontFamily = [Windows.Media.FontFamily]::new('Segoe UI Variable Text, Segoe UI'); $m.Measure([Windows.Size]::new(10000, 10000))
    $w = $m.DesiredSize.Width + 22; $h = $m.DesiredSize.Height + 8
    Add-Rect $X $Y $w $h $Fill '' ($h / 2)
    [void](Add-Text $Text ($X + 11) ($Y + 4) $Size $Ink 'SemiBold')
    return $w
}
function Add-Badge([string]$Text, [double]$X, [double]$Y, [switch]$Filled, [double]$Size = 15) {
    $m = [Windows.Controls.TextBlock]::new(); $m.Text = $Text; $m.FontSize = $Size
    $m.FontFamily = [Windows.Media.FontFamily]::new('Segoe UI Variable Text, Segoe UI'); $m.Measure([Windows.Size]::new(10000, 10000))
    $w = $m.DesiredSize.Width + 28; $h = $m.DesiredSize.Height + 14
    if ($Filled) { Add-Rect $X $Y $w $h 'accent' '' ($h / 2); [void](Add-Text $Text ($X + 14) ($Y + 7) $Size 'onAccent' 'SemiBold') }
    else { Add-Rect $X $Y $w $h 'card' 'stroke' ($h / 2) 1.2; [void](Add-Text $Text ($X + 14) ($Y + 7) $Size 'text2') }
    return $w
}
function Add-Arrow([double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [string]$Color = 'text3', [double]$Thickness = 2.5, [switch]$Dash) {
    $l = [Windows.Shapes.Line]::new(); $l.X1 = $X1; $l.Y1 = $Y1; $l.X2 = $X2; $l.Y2 = $Y2; $l.Stroke = Get-B $Color; $l.StrokeThickness = $Thickness
    if ($Dash) { $l.StrokeDashArray = [Windows.Media.DoubleCollection]::new([double[]]@(3, 3)) }
    [void]$script:Canvas.Children.Add($l)
    $ang = [Math]::Atan2($Y2 - $Y1, $X2 - $X1); $s = 11
    $p = [Windows.Shapes.Polygon]::new(); $p.Fill = Get-B $Color
    foreach ($d in 0.45, -0.45) { $p.Points.Add([Windows.Point]::new($X2 - $s * [Math]::Cos($ang + $d), $Y2 - $s * [Math]::Sin($ang + $d))) }
    $p.Points.Add([Windows.Point]::new($X2, $Y2)); [void]$script:Canvas.Children.Add($p)
}
function Add-Line([double]$X1, [double]$Y1, [double]$X2, [double]$Y2, [string]$Color = 'stroke', [double]$Thickness = 1) {
    $l = [Windows.Shapes.Line]::new(); $l.X1 = $X1; $l.Y1 = $Y1; $l.X2 = $X2; $l.Y2 = $Y2; $l.Stroke = Get-B $Color; $l.StrokeThickness = $Thickness
    [void]$script:Canvas.Children.Add($l)
}
function Add-Card([double]$X, [double]$Y, [double]$W, [double]$H, [string]$Fill = 'card', [string]$Stroke = 'stroke') { Add-Rect $X $Y $W $H $Fill $Stroke 14 1 }
function Add-Step([int]$N, [double]$X, [double]$Y, [double]$D = 44, [string]$Fill = 'accent') {
    Add-Circle $X $Y $D $Fill
    [void](Add-Text ([string]$N) $X ($Y + $D * 0.2) ($D * 0.42) 'white' 'Bold' $D 'Center')
}

function Save-Drawing([string]$Base, [string]$Theme, [int]$W, [int]$H, [int]$PixelWidth = 2160) {
    # The house size of a README picture: 2160 px wide (PraRemoteMailbox, ExchangeLogReport).
    $s = $PixelWidth / [double]$W
    $pxH = [int][Math]::Ceiling($H * $s)
    $c = $script:Canvas; $c.Width = $W; $c.Height = $pxH / $s
    $c.Measure([Windows.Size]::new($W, $c.Height)); $c.Arrange([Windows.Rect]::new(0, 0, $W, $c.Height)); $c.UpdateLayout()
    $bmp = [Windows.Media.Imaging.RenderTargetBitmap]::new($PixelWidth, $pxH, 96 * $s, 96 * $s, [Windows.Media.PixelFormats]::Pbgra32)
    $bmp.Render($c)
    $enc = [Windows.Media.Imaging.PngBitmapEncoder]::new(); $enc.Frames.Add([Windows.Media.Imaging.BitmapFrame]::Create($bmp))
    $file = Join-Path $OutDir ('{0}-{1}.png' -f $Base, $Theme)
    $fs = [IO.File]::Create($file); try { $enc.Save($fs) } finally { $fs.Dispose() }
    Write-Host "drawing $([IO.Path]::GetFileName($file))"
}

# ---- the drawings -----------------------------------------------------------------------------------------------------------------------------
$drawings = [ordered]@{}

$drawings['readme-why'] = @{ W = 1280; H = 400; Draw = {
        Add-Rect 0 0 1280 400 'bg' '' 0
        $items = @(@(0xE715, 'A mailbox', ('Primary, archive,' + [char]10 + 'auxiliary archives')), @(0xE74E, 'A PST file or a store', ('Checked, in parts,' + [char]10 + 'you can look inside')), @(0xE715, 'Another mailbox', ('Items are only added.' + [char]10 + 'Nothing is deleted.')))
        $x = 80
        for ($i = 0; $i -lt 3; $i++) {
            Add-Card $x 70 300 220
            Add-Tile $items[$i][0] ($x + 122) 96 56
            [void](Add-Text $items[$i][1] ($x + 10) 170 24 'text' 'SemiBold' 280 'Center')
            [void](Add-Text $items[$i][2] ($x + 10) 210 18 'text2' 'Normal' 280 'Center')
            if ($i -lt 2) { Add-Arrow ($x + 312) 180 ($x + 396) 180 'accent' 4 }
            $x += 400
        }
        [void](Add-Text 'Made for the big, boring, long job: 50 GB, a night, a PC that may sleep.' 80 330 22 'text2' 'Normal' 1120 'Center')
    } }

$drawings['readme-tasks'] = @{ W = 1280; H = 380; Draw = {
        Add-Rect 0 0 1280 380 'bg' '' 0
        $tasks = @(@(0xE896, 'Browse and export', 'Look inside a mailbox, tick what to keep, save it as PST files, EML files or an export store.', 'Reads only', 'ok'),
            @(0xE898, 'Import', 'Put a PST file, an export store or EML files into a mailbox. Items are added, never replaced.', 'Adds to a mailbox', 'warn'),
            @(0xE90F, 'Check and repair a PST', 'Test a PST. If it is damaged, write a repaired copy. The original is never changed.', 'Local files only', 'ok'),
            @(0xE838, 'Open an export or a PST', 'Read a PST or an export store like a mailbox: folders, messages, attachments.', 'Reads only', 'ok'))
        $x = 40
        foreach ($t in $tasks) {
            Add-Card $x 40 285 300
            Add-Tile $t[0] ($x + 20) 62 52
            [void](Add-Text $t[1] ($x + 20) 130 21 'text' 'SemiBold' 250)
            [void](Add-Text $t[2] ($x + 20) 166 16 'text2' 'Normal' 250)
            if ($t[4] -eq 'ok') { [void](Add-Pill $t[3] ($x + 20) 290 'okBg' 'ok') } else { [void](Add-Pill $t[3] ($x + 20) 290 'warnBg' 'warn') }
            $x += 305
        }
    } }

$drawings['readme-start'] = @{ W = 1280; H = 540; Draw = {
        Add-Rect 0 0 1280 540 'bg' '' 0
        Add-Rect 10 10 1260 520 '' 'stroke' 18 1.5 -Dash
        Add-Icon 0xE713 36 30 20 'accentText'
        [void](Add-Text 'GET READY, ONCE' 66 30 18 'accentText' 'Bold')
        $once = @(@('Get your rights', 'ask your administrator'), @('Put the tool on the PC', 'unzip, then unblock'), @('Write your two numbers', 'tenant id, application id'), @('Open the console', 'double-click the .cmd'))
        for ($i = 0; $i -lt 4; $i++) {
            $x = 36 + $i * 305
            Add-Card $x 66 285 80
            Add-Step ($i + 1) ($x + 14) 86 40
            [void](Add-Text $once[$i][0] ($x + 64) 80 18 'text' 'SemiBold' 215)
            [void](Add-Text $once[$i][1] ($x + 64) 108 16 'text2' 'Normal' 215)
            if ($i -lt 3) { Add-Arrow ($x + 289) 106 ($x + 300) 106 'text3' 2 }
        }
        $branches = @(
            @('A', 0xE896, 'Browse and export', @(@('Sign in, pick the mailbox', 'or open a shared mailbox'), @('Tick, choose where to save', 'read the plan, then Start'), @('Read the result', 'the Done page, then the log'))),
            @('B', 0xE898, 'Import', @(@('Pick the file and the mailbox', 'a test mailbox first'), @('Read the plan, type IMPORT', 'then Start'), @('Read the result', 'what was added, what was refused'))),
            @('C', 0xE90F, 'Check and repair a PST', @(@('Pick the file', 'nothing is written yet'), @('Check, then repair', 'into a new file: the original stays'), @('Read before and after', 'the repair report')))
        )
        for ($b = 0; $b -lt 3; $b++) {
            $x = 36 + $b * 405
            Add-Arrow ($x + 190) 154 ($x + 190) 192 'accent' 2
            Add-Rect $x 196 380 70 'accent' '' 16
            Add-Rect $x 201 380 311 'card' 'stroke' 16 1.2
            Add-Rect ($x + 18) 218 44 44 'text' '' 10
            [void](Add-Text $branches[$b][0] ($x + 18) 224 24 'bg' 'Bold' 44 'Center')
            Add-Tile $branches[$b][1] ($x + 74) 218 44 'soft' 'accent'
            [void](Add-Text $branches[$b][2] ($x + 132) 226 21 'text' 'SemiBold' 240)
            for ($i = 0; $i -lt 3; $i++) {
                $y = 290 + $i * 74
                Add-Step ($i + 5) ($x + 20) $y 38
                [void](Add-Text $branches[$b][3][$i][0] ($x + 74) ($y - 2) 17 'text' 'SemiBold' 290)
                [void](Add-Text $branches[$b][3][$i][1] ($x + 74) ($y + 24) 15 'text2' 'Normal' 290)
            }
        }
    } }

$drawings['readme-path'] = @{ W = 1280; H = 240; Draw = {
        Add-Rect 0 0 1280 240 'bg' '' 0
        $steps = @(@('Task', 'Pick what to do'), @('From', 'Choose the source'), @('To', 'Choose where it goes'), @('Plan', 'Read before you start'), @('Run', 'Wait, stop, continue'), @('Done', 'Read the result'))
        for ($i = 0; $i -lt 6; $i++) {
            $x = 40 + $i * 200
            Add-Step ($i + 1) ($x + 60) 40 56 $(if ($i -eq 3) { 'accent' } else { 'accent' })
            [void](Add-Text $steps[$i][0] $x 108 24 'text' 'SemiBold' 176 'Center')
            [void](Add-Text $steps[$i][1] $x 142 16 'text2' 'Normal' 176 'Center')
            if ($i -lt 5) { Add-Arrow ($x + 128) 68 ($x + 190) 68 'text3' 2.5 }
        }
        [void](Add-Text 'The circles are at the top of the window: the highlighted one is the page you are on.' 40 196 18 'text3' 'Normal' 1200 'Center')
    } }

$drawings['readme-principles'] = @{ W = 1280; H = 400; Draw = {
        Add-Rect 0 0 1280 400 'bg' '' 0
        $p = @(@(0xE710, 'Import only creates', 'Never deletes, overwrites or renames a folder.'), @(0xE8D7, 'Export only reads', 'The mailbox is not changed.'),
            @(0xE73E, 'Everything is checked', 'Each item against its fingerprint, each PST part when written.'), @(0xE90F, 'A repair keeps the original', 'The new file appears only when complete.'),
            @(0xE916, 'Throttling is obeyed', 'The tool waits as long as Microsoft asks.'))
        $x = 24
        foreach ($t in $p) {
            Add-Card $x 50 232 290
            Add-Tile $t[0] ($x + 88) 76 56
            [void](Add-Text $t[1] ($x + 12) 156 20 'text' 'SemiBold' 208 'Center')
            [void](Add-Text $t[2] ($x + 14) 214 16 'text2' 'Normal' 204 'Center')
            $x += 248
        }
    } }

$drawings['readme-speed'] = @{ W = 1280; H = 440; Draw = {
        Add-Rect 0 0 1280 440 'bg' '' 0
        Add-Card 40 30 1200 380
        $x0 = 400; $scale = 800 / 160.0      # 0 to 160 GiB per hour
        for ($g = 0; $g -le 160; $g += 40) {
            $x = $x0 + $g * $scale
            Add-Line $x 100 $x 330 'stroke' 1
            [void](Add-Text ('{0}' -f $g) ($x - 30) 340 15 'text3' 'Normal' 60 'Center')
        }
        [void](Add-Text 'GiB per hour' ($x0 + 340) 368 15 'text3' 'Normal' 130 'Center')
        [void](Add-Text 'How fast, honestly' 70 50 24 'text' 'SemiBold')
        [void](Add-Text 'Import into one mailbox' 70 120 19 'text' 'SemiBold' 310)
        [void](Add-Text 'Microsoft throttles imports per mailbox: more workers do not help.' 70 148 15 'text2' 'Normal' 310)
        Add-Rect $x0 130 (4 * $scale) 28 'bar' '' 6
        [void](Add-Text 'about 4' ($x0 + 4 * $scale + 12) 132 17 'accentText' 'SemiBold')
        [void](Add-Text 'Export from a mailbox' 70 205 19 'text' 'SemiBold' 310)
        [void](Add-Text 'Varies with Microsoft''s load. Observed up to about 150.' 70 233 15 'text2' 'Normal' 310)
        Add-Rect ($x0 + 30 * $scale) 215 (120 * $scale) 28 'barSoft' '' 6
        Add-Rect ($x0 + 30 * $scale) 215 (120 * $scale) 28 '' 'bar' 6 1.5
        [void](Add-Text '30 to 150' ($x0 + 150 * $scale + 12) 217 17 'accentText' 'SemiBold')
        [void](Add-Text 'Write a PST, check, repair' 70 290 19 'text' 'SemiBold' 310)
        [void](Add-Text 'Local disk speed. The window measures it once it runs.' 70 318 15 'text2' 'Normal' 310)
        Add-Rect $x0 300 (600 * 1.0) 28 'info' 'stroke' 6 1 -Dash
        [void](Add-Text 'as fast as your disk' ($x0 + 14) 303 16 'text2')
    } }

$drawings['readme-formats'] = @{ W = 1280; H = 460; Draw = {
        Add-Rect 0 0 1280 460 'bg' '' 0
        Add-Card 40 30 570 400
        Add-Tile 0xE8B7 62 52 56
        [void](Add-Text 'Export store' 134 56 26 'text' 'SemiBold')
        [void](Add-Text 'Our own format, in one folder' 134 92 16 'text2')
        [void](Add-Text 'D:\MailboxTransfer\store' 66 150 17 'accentText' 'SemiBold')
        [void](Add-Text 'catalog.db       the list of every item' 80 184 16 'text2')
        [void](Add-Text 'packs\pack-000001.bin   the items, about 1 GiB each' 80 214 16 'text2')
        [void](Add-Text 'Checked item by item, resumable, quick to import into a mailbox. Only this tool reads it.' 66 290 18 'text' 'Normal' 500)
        Add-Card 670 30 570 400
        Add-Tile 0xE74E 692 52 56
        [void](Add-Text 'PST files' 764 56 26 'text' 'SemiBold')
        [void](Add-Text 'The format Outlook opens' 764 92 16 'text2')
        [void](Add-Text 'D:\Exports\adele' 696 150 17 'accentText' 'SemiBold')
        [void](Add-Text 'adele.pst            the first 50 GB' 710 184 16 'text2')
        [void](Add-Text 'adele-002.pst       the next part' 710 214 16 'text2')
        [void](Add-Text 'adele.manifest.json   what is in each part' 710 244 16 'text2')
        [void](Add-Text 'The real folder tree. Open the file in Outlook. Keep the parts together.' 696 310 18 'text' 'Normal' 500)
    } }

$drawings['readme-permissions'] = @{ W = 1280; H = 700; Draw = {
        Add-Rect 0 0 1280 700 'bg' '' 0
        Add-Card 20 20 1240 660
        $cols = @('Folder.Read', 'Folder.ReadWrite', 'Item.Read', 'Item.Export', 'Item.ImportExport', 'Mail.Read', 'Mail.Read.Shared', 'Mail.ReadWrite')
        $x0 = 440; $cw = 98
        [void](Add-Text 'Delegated permissions (Microsoft Graph), the minimum for each task' 44 36 20 'text' 'SemiBold')
        for ($i = 0; $i -lt $cols.Count; $i++) {
            $label = $cols[$i]; $head = if ($label -like 'Mail.*') { 'Mail' + [char]10 + '.' + $label.Substring(5) } else { 'Mailbox' + $label.Split('.')[0] + [char]10 + '.' + $label.Split('.', 2)[1] }
            [void](Add-Text $head ($x0 + $i * $cw) 70 13 'text2' 'SemiBold' $cw 'Center')
        }
        # marks: 1 required, 2 accepted instead of the required one
        $rows = @(
            @('Open, check or repair a PST or a store', @()),
            @('Browse: folders, sizes, dates (no subject)', @(1, 0, 1, 0, 0, 0, 0, 0)),
            @('Browse: subject, sender, preview', @(1, 0, 1, 0, 0, 1, 0, 0)),
            @('Export to a store or to PST files', @(1, 0, 1, 1, 2, 0, 0, 0)),
            @('Export EML files from the live mailbox', @(1, 0, 1, 0, 0, 1, 0, 0)),
            @('Import a PST, a store or EML files', @(2, 1, 1, 0, 1, 0, 0, 0)),
            @('Import EML with the Exchange conversion', @(2, 1, 1, 0, 1, 0, 0, 1)),
            @('Browse a shared mailbox (FullAccess)', @(1, 0, 1, 0, 0, 0, 1, 0)),
            @('Write to a shared mailbox', @('app'))
        )
        $y = 118
        foreach ($r in $rows) {
            Add-Rect 36 ($y - 6) 1208 52 $(if ($rows.IndexOf($r) % 2 -eq 0) { 'info' } else { 'card' }) '' 8
            [void](Add-Text $r[0] 52 ($y + 4) 17 'text' 'Normal' 376)
            $m = $r[1]
            if ($m.Count -eq 0) { [void](Add-Pill 'no permission, no sign-in' $x0 ($y + 4) 'okBg' 'ok') }
            elseif ($m[0] -eq 'app') { [void](Add-Pill 'application permissions (.All) and an app-only sign-in with a certificate: see the reference guide, 4.6' $x0 ($y + 4) 'warnBg' 'warn') }
            else {
                for ($i = 0; $i -lt 8; $i++) {
                    $cx = $x0 + $i * $cw + $cw / 2 - 10
                    if ($m[$i] -eq 1) { Add-Circle $cx ($y + 6) 20 'accent' }
                    elseif ($m[$i] -eq 2) { Add-Circle $cx ($y + 6) 20 '' 'accent' }
                }
            }
            $y += 56
        }
        Add-Circle 52 646 16 'accent'; [void](Add-Text 'required' 74 643 15 'text2')
        Add-Circle 170 646 16 '' 'accent'; [void](Add-Text 'accepted instead of the filled one (the ReadWrite or ImportExport variant covers a Read or Export)' 192 643 15 'text2')
    } }

$drawings['guide-dictionary'] = @{ W = 1280; H = 470; Draw = {
        Add-Rect 0 0 1280 470 'bg' '' 0
        $w = @(@(0xE715, 'Mailbox', 'Your mail on the Microsoft 365 servers.'), @(0xE7B8, 'Archive', 'A second mailbox for old mail. It may be split into auxiliary archives.'),
            @(0xE8B7, 'Export store', 'A folder where this tool keeps what it downloaded, checked.'), @(0xE74E, 'PST file', 'A file Outlook opens. Big ones come in parts.'),
            @(0xE9D5, 'Plan', 'The list of what will be done. Read it, then press Start.'), @(0xE916, 'Throttling', 'Microsoft slows us down. Normal. The tool waits.'),
            @(0xE897, 'Uncertain item', 'Sent, but no answer came. The tool looks for it, it never guesses.'), @(0xE72D, 'Job', 'One run. You can stop it and continue it later.'))
        for ($i = 0; $i -lt 8; $i++) {
            $x = 30 + ($i % 4) * 305; $y = 30 + [Math]::Floor($i / 4) * 220
            Add-Card $x $y 285 200
            Add-Tile $w[$i][0] ($x + 20) ($y + 22) 52
            [void](Add-Text $w[$i][1] ($x + 20) ($y + 92) 22 'text' 'SemiBold' 250)
            [void](Add-Text $w[$i][2] ($x + 20) ($y + 130) 16 'text2' 'Normal' 250)
        }
    } }

$drawings['guide-aux-archive'] = @{ W = 1280; H = 520; Draw = {
        Add-Rect 0 0 1280 520 'bg' '' 0
        Add-Card 40 40 520 440
        [void](Add-Text 'Archive (main)' 66 58 22 'text' 'SemiBold')
        $f = @('Finance / Budgets', 'Old Mail / 2018', 'Projects')
        for ($i = 0; $i -lt 3; $i++) { Add-Icon 0xE8B7 70 (118 + $i * 70) 26; [void](Add-Text $f[$i] 110 (118 + $i * 70) 19 'text'); }
        [void](Add-Pill 'a placeholder' 330 (118 + 70) 'warnBg' 'warn')
        [void](Add-Text 'The archive is full, so Microsoft adds auxiliary archives and moves folders there over several hours.' 66 350 17 'text2' 'Normal' 470)
        Add-Card 720 40 520 440
        [void](Add-Text 'Auxiliary archive 1' 746 58 22 'text' 'SemiBold')
        Add-Icon 0xE8B7 750 188 26; [void](Add-Text 'Old Mail / 2018' 790 188 19 'text')
        [void](Add-Pill 'being filled' 1030 188 'warnBg' 'warn')
        [void](Add-Text 'While it is being filled the folder cannot be read: the window says "not ready", the tool tries again later and never exports it as empty. Press Continue later.' 746 350 17 'text2' 'Normal' 470)
        Add-Arrow 570 210 710 210 'warn' 3.5
        [void](Add-Text 'moves' 596 176 16 'warn' 'SemiBold')
    } }

$drawings['guide-uncertain'] = @{ W = 1280; H = 440; Draw = {
        Add-Rect 0 0 1280 440 'bg' '' 0
        $s = @(@(0xE898, 'The item is sent', ''), @(0xE9CE, 'No answer comes', 'timeout, or the line drops'), @(0xE916, 'Wait 5 minutes', 'Microsoft may still be creating it'), @(0xE721, 'Look in the folder', 'items created after the attempt'))
        $x = 40
        for ($i = 0; $i -lt 4; $i++) {
            Add-Card $x 40 260 150
            Add-Tile $s[$i][0] ($x + 18) 58 48
            [void](Add-Text $s[$i][1] ($x + 18) 118 18 'text' 'SemiBold' 230)
            if ($s[$i][2]) { [void](Add-Text $s[$i][2] ($x + 18) 148 14.5 'text2' 'Normal' 230) }
            if ($i -lt 3) { Add-Arrow ($x + 264) 115 ($x + 296) 115 'text3' 2.5 }
            $x += 300
        }
        Add-Arrow 1170 195 1170 240 'text3' 2.5
        $o = @(@('Found once', 'It arrived: marked done.', 'okBg', 'ok'), @('Not found', 'It did not arrive: sent again.', 'okBg', 'ok'), @('Several could be it', 'The window asks you. We do not guess.', 'warnBg', 'warn'))
        $x = 340
        for ($i = 0; $i -lt 3; $i++) {
            Add-Card $x 250 290 130 $o[$i][2] ''
            [void](Add-Text $o[$i][0] ($x + 18) 270 20 $o[$i][3] 'SemiBold' 260)
            [void](Add-Text $o[$i][1] ($x + 18) 308 16.5 'text' 'Normal' 258)
            $x += 305
        }
        [void](Add-Text 'Three ways out' 40 300 20 'text2' 'SemiBold' 280)
    } }

$drawings['guide-pst-split'] = @{ W = 1280; H = 480; Draw = {
        Add-Rect 0 0 1280 480 'bg' '' 0
        Add-Card 40 60 360 340
        [void](Add-Text 'One mailbox, 62 GB' 64 78 20 'text' 'SemiBold')
        $f = @('Inbox', 'Projects', 'Sent Items', 'Archive / Old Mail', 'Archive / Budgets')
        for ($i = 0; $i -lt $f.Count; $i++) { Add-Icon 0xE8B7 70 (130 + $i * 46) 22; [void](Add-Text $f[$i] 106 (130 + $i * 46) 18 'text2') }
        Add-Arrow 410 230 520 230 'accent' 4
        Add-Card 530 40 340 160
        Add-Tile 0xE74E 550 58 48; [void](Add-Text 'adele.pst' 614 62 22 'text' 'SemiBold'); [void](Add-Text 'the first 50 GB, with the whole folder tree' 550 124 16 'text2' 'Normal' 300)
        Add-Card 530 220 340 160
        Add-Tile 0xE74E 550 238 48; [void](Add-Text 'adele-002.pst' 614 242 22 'text' 'SemiBold'); [void](Add-Text 'the rest, in the same folders' 550 304 16 'text2' 'Normal' 300)
        Add-Card 900 130 340 160 'okBg' ''
        [void](Add-Text 'What is guaranteed' 920 146 19 'ok' 'SemiBold' 300)
        [void](Add-Text 'An item is never cut: the split is between items.' 920 182 16 'text' 'Normal' 300)
        [void](Add-Text 'Each part is checked as soon as it is written.' 920 230 16 'text' 'Normal' 300)
        [void](Add-Text 'Choose 50, 20, 10 or 5 GB in the window (any size from 1 GB on the command line).' 530 410 17 'text2' 'Normal' 710)
    } }

$drawings['guide-repair-levels'] = @{ W = 1280; H = 480; Draw = {
        Add-Rect 0 0 1280 480 'bg' '' 0
        $l = @(@('Recompute', 'Trusts the folder and message lists.', 'Fixes counters, sizes, maps. Keeps what it can read.'), @('Rebuild', 'Trusts the messages themselves, not the lists.', 'Also finds messages and folders no list shows. Recommended when the check says so.'), @('Salvage', 'Scans the whole file for readable blocks.', 'Also works when the header or the indexes are broken, or the file is cut off.'))
        $x = 40
        for ($i = 0; $i -lt 3; $i++) {
            Add-Card $x 40 370 330
            Add-Step ($i + 1) ($x + 20) 62 44
            [void](Add-Text $l[$i][0] ($x + 80) 70 26 'text' 'SemiBold')
            [void](Add-Text $l[$i][1] ($x + 20) 130 18 'accentText' 'SemiBold' 330)
            [void](Add-Text $l[$i][2] ($x + 20) 200 17 'text2' 'Normal' 330)
            if ($i -lt 2) { Add-Arrow ($x + 376) 205 ($x + 420) 205 'text3' 2.5 }
            $x += 420
        }
        [void](Add-Text 'Automatic (the default) starts at the level the check recommends and goes up by itself when a level cannot read the file. The original file is never changed.' 40 400 18 'text2' 'Normal' 1200)
    } }

$drawings['guide-continue'] = @{ W = 1280; H = 380; Draw = {
        Add-Rect 0 0 1280 380 'bg' '' 0
        $s = @(@(0xE768, 'Run', 'the job works'), @(0xE71A, 'Stop, sleep, close', 'or power cut'), @(0xE895, 'Continue', 'the same job, the same command'), @(0xE73E, 'Only what is left', 'items done are skipped'))
        $x = 40
        for ($i = 0; $i -lt 4; $i++) {
            Add-Card $x 60 260 200 $(if ($i -eq 1) { 'warnBg' } else { 'card' }) $(if ($i -eq 1) { '' } else { 'stroke' })
            Add-Tile $s[$i][0] ($x + 18) 82 52 $(if ($i -eq 1) { 'card' } else { 'soft' }) $(if ($i -eq 1) { 'warn' } else { 'accent' })
            [void](Add-Text $s[$i][1] ($x + 18) 150 20 'text' 'SemiBold' 230)
            [void](Add-Text $s[$i][2] ($x + 18) 184 16 'text2' 'Normal' 230)
            if ($i -lt 3) { Add-Arrow ($x + 264) 160 ($x + 296) 160 'text3' 2.5 }
            $x += 300
        }
        [void](Add-Text 'Nothing is lost. An item whose answer was lost is looked for, not guessed.' 40 300 20 'text2' 'Normal' 1200 'Center')
    } }

# ---- the README banner (hero): an SVG, the template of the ExchangeLogReport hero (1600 x 490, same coordinates, same sizes) -----------------------
# Same file kind as ExchangeLogReport (readme-banner-light.svg / readme-banner-dark.svg, <picture> in the README). Only the words change; the texts are
# checked for privacy like every picture text, and measured so that a line cannot run into the stat cards (a 12 % margin for a reader without Segoe UI).
$hero = [ordered]@{
    Title    = 'Mailbox Transfer'
    Eyebrow  = 'EXCHANGE ONLINE / MAILBOXES AND PST FILES'
    Sub      = 'Export and import Exchange Online mailboxes.'
    Lines    = @('Verified and resumable, with a guided Windows console.', 'Checks and repairs PST files, without Outlook.')
    Pills    = @("$Version release", 'PowerShell 7.5+', 'No Outlook needed')
    Footer   = 'Nicolas Fabert / MIT license'
    Desc     = 'Exchange Online mailbox export and import to a verified store or to PST files, with a guided Windows console and PST check and repair.'
    Stats    = @(
        @{ N = '4'; Label = 'Tasks in one console'; Sub = 'Export, import, check PST, open' }
        @{ N = '50'; Label = 'GB per PST part'; Sub = 'Each part checked when written' }
        @{ N = '0'; Label = 'Items deleted'; Sub = 'Import adds, export only reads' }
    )
}
$heroColors = @{
    light = @{ S0 = '#ffffff'; S1 = '#f9eff2'; Line = '#dddddd'; Accent = '#b71c50'; Ring = '.08'; Text = '#242424'; Text2 = '#535353'; Card = '#ffffff'; OnAccent = '#ffffff' }
    dark  = @{ S0 = '#262626'; S1 = '#46353b'; Line = '#535353'; Accent = '#ff89a5'; Ring = '.1'; Text = '#f5f5f5'; Text2 = '#d1d1d1'; Card = '#262626'; OnAccent = '#242424' }
}
function Get-HeroTextWidth([string]$Text, [double]$Size, [string]$Weight = 'Normal', [double]$Spacing = 0) {
    $m = [Windows.Controls.TextBlock]::new(); $m.Text = $Text; $m.FontSize = $Size; $m.FontWeight = [Windows.FontWeights]::$Weight
    $m.FontFamily = [Windows.Media.FontFamily]::new('Segoe UI'); $m.Measure([Windows.Size]::new(10000, 10000))
    return $m.DesiredSize.Width + $Spacing * $Text.Length
}
function Assert-HeroFits([string]$Text, [double]$Size, [double]$Available, [string]$Weight = 'Normal', [double]$Spacing = 0) {
    $w = (Get-HeroTextWidth $Text $Size $Weight $Spacing) * 1.12
    if ($w -gt $Available) { throw "The banner text '$Text' is too wide ($([int]$w) px for $Available)." }
}
function New-HeroSvg([string]$Theme) {
    $c = $heroColors[$Theme]; $h = $hero
    $all = @($h.Title, $h.Eyebrow, $h.Sub, $h.Footer, $h.Desc) + $h.Lines + $h.Pills + @($h.Stats | ForEach-Object { $_.N; $_.Label; $_.Sub })
    foreach ($s in $all) { Assert-PublicText $s }
    Assert-HeroFits $h.Eyebrow 16 840 'Bold' 2; Assert-HeroFits $h.Title 43 880 'Bold'; Assert-HeroFits $h.Sub 27 990
    foreach ($l in $h.Lines) { Assert-HeroFits $l 23 990 }
    foreach ($s in $h.Stats) { Assert-HeroFits $s.Label 21 320 'SemiBold'; Assert-HeroFits $s.Sub 19 320; Assert-HeroFits $s.N 32 80 'Bold' }
    # the pills: 13 px apart from x = 55, 48 px high; the first one (filled) is 212 px wide as in ExchangeLogReport, the others fit their text
    $x = 55; $pills = [Text.StringBuilder]::new()
    for ($k = 0; $k -lt $h.Pills.Count; $k++) {
        $txt = $h.Pills[$k]
        $w = if ($k -eq 0) { [Math]::Max(212, [int][Math]::Ceiling((Get-HeroTextWidth $txt 19 'SemiBold') + 60)) } else { [int][Math]::Ceiling((Get-HeroTextWidth $txt 19) + 60) }
        $cx = [int]($x + $w / 2)
        if ($k -eq 0) {
            [void]$pills.AppendLine("    <rect x=""$x"" y=""335"" width=""$w"" height=""48"" rx=""24"" fill=""$($c.Accent)""/>")
            [void]$pills.AppendLine("    <text x=""$cx"" y=""365"" text-anchor=""middle"" fill=""$($c.OnAccent)"" font-size=""19"" font-weight=""600"">$txt</text>")
        } else {
            [void]$pills.AppendLine("    <rect x=""$x"" y=""335"" width=""$w"" height=""48"" rx=""24"" fill=""$($c.Card)"" stroke=""$($c.Line)""/>")
            [void]$pills.AppendLine("    <text x=""$cx"" y=""365"" text-anchor=""middle"" fill=""$($c.Text2)"" font-size=""19"">$txt</text>")
        }
        $x += $w + 13
    }
    if ($x - 13 -gt 1040) { throw 'The banner pills run into the stat cards.' }
    $rows = 0..2 | ForEach-Object { 75 + 120 * $_ }
    $cards = ($rows | ForEach-Object { "      <rect x=""1080"" y=""$_"" width=""455"" height=""100"" rx=""22""/>" }) -join "`n"
    $nums = (0..2 | ForEach-Object { "<text x=""1110"" y=""$($rows[$_] + 61)"">$($h.Stats[$_].N)</text>" }) -join ''
    $labels = (0..2 | ForEach-Object { "<text x=""1200"" y=""$($rows[$_] + 44)"">$($h.Stats[$_].Label)</text>" }) -join ''
    $subs = (0..2 | ForEach-Object { "<text x=""1200"" y=""$($rows[$_] + 75)"">$($h.Stats[$_].Sub)</text>" }) -join ''
    @"
<svg xmlns="http://www.w3.org/2000/svg" width="1600" height="490" viewBox="0 0 1600 490" role="img" aria-labelledby="title desc">
  <title id="title">$($h.Title)</title>
  <desc id="desc">$($h.Desc)</desc>
  <defs>
    <linearGradient id="surface" x1="0" y1="0" x2="1" y2="1">
      <stop stop-color="$($c.S0)"/><stop offset="1" stop-color="$($c.S1)"/>
    </linearGradient>
    <clipPath id="card"><rect x="10" y="10" width="1580" height="470" rx="26"/></clipPath>
  </defs>
  <rect x="10" y="10" width="1580" height="470" rx="26" fill="url(#surface)" stroke="$($c.Line)" stroke-width="2"/>
  <g clip-path="url(#card)"><circle cx="1430" cy="40" r="235" fill="none" stroke="$($c.Accent)" stroke-opacity="$($c.Ring)" stroke-width="70"/></g>
  <path d="M37 13H1563" stroke="$($c.Accent)" stroke-width="5" stroke-linecap="round"/>
  <g font-family="Segoe UI, Arial, sans-serif">
    <rect x="55" y="55" width="88" height="88" rx="24" fill="$($c.Accent)"/>
    <g fill="none" stroke="$($c.OnAccent)" stroke-width="3.5" stroke-linejoin="round"><rect x="76" y="82" width="46" height="34" rx="5"/><path d="M79 88l20 15 20-15"/></g>
    <text x="170" y="77" fill="$($c.Accent)" font-size="16" font-weight="700" letter-spacing="2">$($h.Eyebrow)</text>
    <text x="170" y="126" fill="$($c.Text)" font-size="43" font-weight="700">$($h.Title)</text>
    <text x="55" y="205" fill="$($c.Text)" font-size="27">$($h.Sub)</text>
    <text x="55" y="252" fill="$($c.Text2)" font-size="23">$($h.Lines[0])</text>
    <text x="55" y="286" fill="$($c.Text2)" font-size="23">$($h.Lines[1])</text>
$($pills.ToString().TrimEnd())
    <text x="55" y="435" fill="$($c.Text2)" font-size="17">$($h.Footer)</text>
    <g fill="$($c.Card)" stroke="$($c.Line)" stroke-width="2">
$cards
    </g>
    <g fill="$($c.Accent)" font-weight="700" font-size="32">$nums</g>
    <g fill="$($c.Text)" font-size="21" font-weight="600">$labels</g>
    <g fill="$($c.Text2)" font-size="19">$subs</g>
  </g>
</svg>
"@
}
foreach ($theme in 'light', 'dark') {
    if (-not (Test-Wanted "readme-banner-$theme")) { continue }
    $file = Join-Path $OutDir "readme-banner-$theme.svg"
    [IO.File]::WriteAllText($file, ((New-HeroSvg $theme).Replace("`r`n", "`n").TrimEnd() + "`n"), [Text.UTF8Encoding]::new($false))
    Write-Host "banner  $([IO.Path]::GetFileName($file))"
}
Get-ChildItem -LiteralPath $OutDir -Filter 'readme-banner-*.png' -ErrorAction SilentlyContinue | Remove-Item -Force   # the banner was a PNG until 1.0.1
foreach ($key in $drawings.Keys) {
    if (-not (Test-Wanted "$key-light")) { continue }
    $d = $drawings[$key]
    foreach ($theme in 'light', 'dark') {
        $script:Pal = $palettes[$theme]; $script:Brush = @{}
        $script:Canvas = [Windows.Controls.Canvas]::new()
        $script:Canvas.Background = Get-B 'bg'
        & $d.Draw
        $pw = if ($d.ContainsKey('PixelWidth')) { $d.PixelWidth } else { 2160 }
        Save-Drawing $key $theme $d.W $d.H $pw
    }
}
