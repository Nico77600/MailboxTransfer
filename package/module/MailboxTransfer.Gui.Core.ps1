#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - WPF foundations (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    Types (view models), theme, small element factories, text and number formats, the background worker. The look is the one of Mailbox Message
    Report 2.1: Fluent theme of WPF (.NET 9 and later) light or dark like Windows, the accent of the house (#B11F4B / #FD8EA1) instead of the
    Windows accent, cards on a flat page, 24 px margins, 12 px between cards.

    Nothing here knows the engine. Pages (MailboxTransfer.Gui.Pages.ps1) and the state machine (MailboxTransfer.Gui.State.ps1) are built from it.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

$script:GuiRoot = Split-Path $PSScriptRoot -Parent
$script:Gui = $null
# A console started from a source checkout (-DevelopmentMode) starts its jobs the same way.
$script:GuiDevelopmentMode = $false

#region Types and theme ----------------------------------------------------------------------

function Initialize-MbtGuiTypes {
    <# Loads WPF and compiles the view models (src\Gui.ViewModels.cs) once per process. #>
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    if (-not ('MailboxTransferGui.FolderNode' -as [type])) {
        $src = Join-Path $script:GuiRoot 'src\Gui.ViewModels.cs'
        Add-Type -TypeDefinition (Get-Content -LiteralPath $src -Raw) -Language CSharp
    }
}

function Test-MbtGuiDarkMode {
    <# Windows shows the applications in dark mode (AppsUseLightTheme = 0); light when the setting is missing. #>
    $personalize = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue
    if (-not $personalize) { return $false }
    return [string]$personalize.AppsUseLightTheme -eq '0'
}

function Initialize-MbtGuiTheme {
    <# Applies the Fluent theme to the application (one application per process). Returns Fluent and Dark. #>
    param([ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')
    Initialize-MbtGuiTypes
    $dark = $Theme -eq 'Dark'
    if ($Theme -eq 'System') { $dark = Test-MbtGuiDarkMode }
    $app = [Windows.Application]::Current
    if (-not $app) {
        $app = [Windows.Application]::new()
        $app.ShutdownMode = [Windows.ShutdownMode]::OnExplicitShutdown
    }
    $fluent = $null -ne [Windows.Application].GetProperty('ThemeMode')
    if ($fluent) { $app.ThemeMode = [Windows.ThemeMode]::new($(if ($dark) { 'Dark' } else { 'Light' })) }
    [pscustomobject]@{ Fluent = $fluent; Dark = $dark; Application = $app }
}

function New-MbtGuiBrush {
    param([Parameter(Mandatory)][string]$Color)
    $brush = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush
}

function Set-MbtGuiTheme {
    <# The accent of the house on the Fluent accent resources, the status colours, and (classic theme) the Fluent resources the window uses. #>
    param([Parameter(Mandatory)][Windows.Window]$Window, [Parameter(Mandatory)]$Theme)
    $dark = $Theme.Dark
    $r = $Window.Resources
    $set = { param([string[]]$Keys, [string]$LightColor, [string]$DarkColor) $b = New-MbtGuiBrush $(if ($dark) { $DarkColor } else { $LightColor }); foreach ($k in $Keys) { $r[$k] = $b } }
    if (-not $Theme.Fluent) {
        & $set 'ApplicationBackgroundBrush' '#F7F4EF' '#202020'
        & $set 'CardBackgroundFillColorDefaultBrush' '#FFFFFF' '#2B2B2B'
        & $set 'CardStrokeColorDefaultBrush', 'ControlStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'ControlFillColorDefaultBrush' '#FFFFFF' '#2D2D2D'
        & $set 'ControlFillColorSecondaryBrush' '#F5F5F5' '#323232'
        & $set 'DividerStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'LayerFillColorDefaultBrush' '#FCFBF8' '#262626'
        & $set 'TextFillColorPrimaryBrush' '#242424' '#FFFFFF'
        & $set 'TextFillColorSecondaryBrush' '#5C5C5C' '#C5C5C5'
        & $set 'TextFillColorTertiaryBrush' '#8A8A8A' '#9A9A9A'
    }
    & $set 'AccentFillColorDefaultBrush', 'AccentButtonBackground', 'AccentButtonBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'AccentFillColorSecondaryBrush', 'AccentButtonBackgroundPointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'AccentFillColorTertiaryBrush', 'AccentButtonBackgroundPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'AccentTextFillColorPrimaryBrush' '#9A1A41' '#FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillChecked', 'CheckBoxCheckBackgroundStrokeChecked', 'RadioButtonOuterEllipseCheckedFill', 'RadioButtonOuterEllipseCheckedStroke',
        'ProgressBarForeground', 'CalendarViewSelectedBackground', 'CalendarViewSelectedBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillCheckedPointerOver', 'CheckBoxCheckBackgroundStrokeCheckedPointerOver', 'RadioButtonOuterEllipseCheckedStrokePointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillIndeterminate', 'CheckBoxCheckBackgroundStrokeIndeterminate', 'TreeViewItemSelectionIndicatorForeground' '#B11F4B' '#FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillIndeterminatePointerOver', 'CheckBoxCheckBackgroundStrokeIndeterminatePointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillIndeterminatePressed', 'CheckBoxCheckBackgroundStrokeIndeterminatePressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'CheckBoxCheckBackgroundFillCheckedPressed', 'CheckBoxCheckBackgroundStrokeCheckedPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'DataGridRowSelectedBackgroundThemeBrush' '#1FB11F4B' '#40FD8EA1'
    & $set 'ListViewItemPillFillBrush' '#B11F4B' '#FD8EA1'
    & $set 'DataGridRowSelectedForegroundThemeBrush' '#242424' '#FFFFFF'
    $r[[Windows.SystemColors]::HighlightBrushKey] = $r['DataGridRowSelectedBackgroundThemeBrush']
    $r[[Windows.SystemColors]::InactiveSelectionHighlightBrushKey] = $r['DataGridRowSelectedBackgroundThemeBrush']
    $r[[Windows.SystemColors]::HighlightTextBrushKey] = $r['DataGridRowSelectedForegroundThemeBrush']
    $r[[Windows.SystemColors]::InactiveSelectionHighlightTextBrushKey] = $r['DataGridRowSelectedForegroundThemeBrush']
    & $set 'MbtBrand' '#B11F4B' '#B11F4B'
    & $set 'MbtBrandText' '#B11F4B' '#FD8EA1'
    & $set 'MbtAccentSoft' '#14B11F4B' '#33FD8EA1'
    & $set 'MbtSuccess' '#16A34A' '#4ADE80'
    & $set 'MbtSuccessInk' '#15803D' '#4ADE80'
    & $set 'MbtCaution' '#D97706' '#FBBF24'
    & $set 'MbtCautionInk' '#92400E' '#FBBF24'
    & $set 'MbtCritical' '#DC2626' '#F87171'
    & $set 'MbtCriticalInk' '#B91C1C' '#F87171'
    & $set 'MbtInfoBackground' '#F3F3F3' '#2E2E2E'
    & $set 'MbtCautionBackground' '#FFF7E8' '#33FBBF24'
    & $set 'MbtCriticalBackground' '#FDECEC' '#33F87171'
    & $set 'MbtSuccessBackground' '#EAF7EE' '#334ADE80'
    & $set 'MbtScrim' '#66000000' '#A3000000'
}

#endregion

#region Text and numbers --------------------------------------------------------------------

# Glyphs of Segoe Fluent Icons (Segoe MDL2 Assets as the fallback), by the names the pages use.
$script:Glyphs = @{
    mail = 0xE715; folder = 0xE8B7; archive = 0xE7B8; cal = 0xE787; people = 0xE716; trash = 0xE74D; send = 0xE724; pencil = 0xE70F; person = 0xE77B
    cloud = 0xE753; down = 0xE896; up = 0xE898; pst = 0xE8A5; wrench = 0xE90F; search = 0xE721; check = 0xE73E; warn = 0xE7BA; x = 0xE711; chev = 0xE76C
    clock = 0xE121; pause = 0xE769; play = 0xE768; stop = 0xE71A; info = 0xE946; key = 0xE8D7; open = 0xE838; refresh = 0xE72C; moon = 0xE708
    disk = 0xEDA2; doc = 0xE7C3; shield = 0xEA18; clip = 0xE723; dot = 0xEA3A; lock = 0xE72E
}

function ConvertTo-MbtGuiThickness {
    <# "0,10,0,0" or "8" to a Thickness. #>
    param([string]$Text)
    $p = @($Text -split '[,\s]+' | Where-Object { $_ } | ForEach-Object { [double]$_ })
    switch ($p.Count) { 1 { [Windows.Thickness]::new($p[0]) } 2 { [Windows.Thickness]::new($p[0], $p[1], $p[0], $p[1]) } default { [Windows.Thickness]::new($p[0], $p[1], $p[2], $p[3]) } }
}

function Get-MbtGlyph {
    param([string]$Name)
    if (-not $Name -or -not $script:Glyphs.ContainsKey($Name)) { return '' }
    return [string][char]$script:Glyphs[$Name]
}

# Figures of the Plan page and the ETA, in ONE place (spec 14.7: they are shown as ranges; to be re-measured before the guide prints them).
$script:Throughput = @{
    ImportGiBPerHour = @(3, 5)        # Microsoft lets about 4 GiB per hour into one mailbox
    ExportGiBPerHour = @(30, 150)     # about 150 GiB/h observed, to be measured again on a second tenant
    ExportMinimumMinutes = @(2, 6)    # never show less than this range for a small export
    EtaWarmupMinutes = 10             # no ETA for the first minutes of a run
    EtaWindows = 6                    # windows of 5 minutes kept
    SleepGapSeconds = 120             # a gap longer than this between two events is a pause, not a rate
    ForceStopMinutes = 10             # "Force stop" is offered after this
}

function Format-MbtGuiBytes {
    param([double]$Bytes)
    if ($Bytes -lt 0) { return 'unknown' }
    if ($Bytes -eq 0) { return '0 B' }
    $gb = 1GB; $mb = 1MB; $kb = 1KB
    if ($Bytes -ge $gb) { $v = $Bytes / $gb; return ('{0} GiB' -f [math]::Round($v, $(if ($v -ge 10) { 1 } else { 2 })).ToString([cultureinfo]::InvariantCulture)) }
    if ($Bytes -ge $mb) { return ('{0} MiB' -f [math]::Round($Bytes / $mb)) }
    return ('{0} KiB' -f [math]::Max(1, [math]::Round($Bytes / $kb)))
}

function Format-MbtGuiCount { param([double]$N) return ([long]$N).ToString('N0', [cultureinfo]::InvariantCulture) }

function Format-MbtGuiSelectionSize {
    <# The size of a selection in words: a folder whose size Microsoft does not give (an auxiliary archive, a PST folder) is never counted as 0 B. #>
    param([Parameter(Mandatory)]$Selection)
    $known = [double]$Selection.Bytes; $unknown = [int]$Selection.UnknownSize
    if ($unknown -le 0) { return (Format-MbtGuiBytes $known) }
    if ($known -le 0) { return 'size unknown' }
    return ('{0} + {1} folder{2} of unknown size' -f (Format-MbtGuiBytes $known), $unknown, $(if ($unknown -gt 1) { 's' } else { '' }))
}

function Format-MbtGuiSelectionItems {
    <# The number of items of a selection: "at least" while a ticked folder has subfolders that are not read yet (they are counted when the plan is checked). #>
    param([Parameter(Mandatory)]$Selection)
    $n = Format-MbtGuiCount $Selection.Items
    if ($Selection.SubtreeUnread) { return "at least $n" }
    return $n
}

function Format-MbtGuiDuration {
    <# "5 h 50 min", "20 min", "3 min": for elapsed times and ETA. #>
    param([double]$Seconds, [switch]$WithSeconds)
    $s = [math]::Max(0, [long]$Seconds)
    $h = [math]::Floor($s / 3600); $m = [math]::Floor(($s % 3600) / 60)
    if ($h -gt 0) { return $(if ($m -gt 0) { "$h h $m min" } else { "$h h" }) }
    if ($m -gt 0) { return $(if ($WithSeconds -and ($s % 60) -gt 0) { "$m min $($s % 60) s" } else { "$m min" }) }
    return "$s s"
}

function ConvertTo-MbtGuiRuns {
    <# Adds the text of a TextBlock with **bold** parts as runs. #>
    param([Parameter(Mandatory)][Windows.Controls.TextBlock]$Block, [AllowEmptyString()][string]$Text)
    $Block.Inlines.Clear()
    $bold = $false
    foreach ($part in ($Text -split '\*\*')) {
        if ($part.Length -gt 0) {
            $run = [Windows.Documents.Run]::new($part)
            if ($bold) { $run.FontWeight = [Windows.FontWeights]::SemiBold }
            $Block.Inlines.Add($run)
        }
        $bold = -not $bold
    }
}

function Set-MbtGuiResource {
    param([Parameter(Mandatory)][Windows.FrameworkElement]$Element, [Parameter(Mandatory)]$Property, [Parameter(Mandatory)][string]$Key)
    $Element.SetResourceReference($Property, $Key)
}

#endregion

#region Element factories -------------------------------------------------------------------

function New-GuiText {
    <# A TextBlock. Style: Body, Hint (12, secondary), Label (12, secondary, space above), Muted, Small, Title. **bold** is supported. #>
    param([AllowEmptyString()][string]$Text = '', [ValidateSet('Body', 'Hint', 'Label', 'Muted', 'Small', 'Title', 'H2', 'Big')][string]$Style = 'Body', [switch]$NoWrap, [string]$Margin)
    $t = [Windows.Controls.TextBlock]::new()
    ConvertTo-MbtGuiRuns $t $Text
    $t.TextWrapping = $(if ($NoWrap) { 'NoWrap' } else { 'Wrap' })
    $primary = 'TextFillColorPrimaryBrush'; $secondary = 'TextFillColorSecondaryBrush'
    switch ($Style) {
        'Body' { Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $primary }
        'Hint' { $t.FontSize = 12; $t.Margin = [Windows.Thickness]::new(0, 6, 0, 0); Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $secondary }
        'Label' { $t.FontSize = 12; $t.Margin = [Windows.Thickness]::new(0, 10, 0, 4); Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $secondary }
        'Muted' { Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $secondary }
        'Small' { $t.FontSize = 12; Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $secondary }
        'Title' { $t.FontSize = 14; $t.FontWeight = [Windows.FontWeights]::SemiBold; Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $primary }
        'H2' { $t.FontSize = 20; $t.FontWeight = [Windows.FontWeights]::SemiBold; $t.Margin = [Windows.Thickness]::new(0, 2, 0, 10); Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $primary }
        'Big' { $t.FontSize = 18; $t.FontWeight = [Windows.FontWeights]::SemiBold; Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $primary }
    }
    if ($Margin) { $t.Margin = ConvertTo-MbtGuiThickness ($Margin) }
    return $t
}

function New-GuiIcon {
    param([Parameter(Mandatory)][string]$Name, [double]$Size = 14, [string]$Brush = 'MbtBrandText')
    $t = [Windows.Controls.TextBlock]::new()
    $t.Text = Get-MbtGlyph $Name
    $t.FontFamily = [Windows.Media.FontFamily]::new('Segoe Fluent Icons, Segoe MDL2 Assets')
    $t.FontSize = $Size
    $t.VerticalAlignment = 'Center'
    Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $Brush
    return $t
}

function New-GuiStack {
    param([object[]]$Children = @(), [ValidateSet('Vertical', 'Horizontal')][string]$Orientation = 'Vertical', [string]$Margin)
    $s = [Windows.Controls.StackPanel]::new()
    $s.Orientation = $Orientation
    foreach ($c in @($Children)) { if ($null -ne $c) { [void]$s.Children.Add($c) } }
    if ($Margin) { $s.Margin = ConvertTo-MbtGuiThickness ($Margin) }
    return $s
}

function New-GuiPill {
    <# Kind: '' (grey), ok, warn, bad, run. #>
    param([Parameter(Mandatory)][string]$Text, [ValidateSet('', 'ok', 'warn', 'bad', 'run')][string]$Kind = '', [switch]$Light)
    $bg = switch ($Kind) { 'ok' { 'MbtSuccessBackground' } 'warn' { 'MbtCautionBackground' } 'bad' { 'MbtCriticalBackground' } 'run' { 'MbtAccentSoft' } default { 'MbtInfoBackground' } }
    $fg = switch ($Kind) { 'ok' { 'MbtSuccessInk' } 'warn' { 'MbtCautionInk' } 'bad' { 'MbtCriticalInk' } 'run' { 'MbtBrandText' } default { 'TextFillColorSecondaryBrush' } }
    $b = [Windows.Controls.Border]::new()
    $b.CornerRadius = [Windows.CornerRadius]::new(10)
    $b.Padding = [Windows.Thickness]::new(10, 2, 10, 2)
    $b.VerticalAlignment = 'Center'
    $b.HorizontalAlignment = 'Left'
    Set-MbtGuiResource $b ([Windows.Controls.Border]::BackgroundProperty) $bg
    $t = [Windows.Controls.TextBlock]::new()
    $t.Text = $Text; $t.FontSize = 12; $t.TextWrapping = 'NoWrap'
    if (-not $Light) { $t.FontWeight = [Windows.FontWeights]::SemiBold }
    Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $fg
    $b.Child = $t
    $b.Tag = $Text
    return $b
}

function New-GuiMessage {
    <# A coloured message: Kind info, ok, warn, bad; Icon is a glyph name. #>
    param([ValidateSet('info', 'ok', 'warn', 'bad')][string]$Kind = 'info', [string]$Icon = 'info', [Parameter(Mandatory)][string]$Text, [object[]]$After = @())
    $bg = switch ($Kind) { 'ok' { 'MbtSuccessBackground' } 'warn' { 'MbtCautionBackground' } 'bad' { 'MbtCriticalBackground' } default { 'MbtInfoBackground' } }
    $fg = switch ($Kind) { 'ok' { 'MbtSuccessInk' } 'warn' { 'MbtCautionInk' } 'bad' { 'MbtCriticalInk' } default { 'TextFillColorSecondaryBrush' } }
    $ic = switch ($Kind) { 'ok' { 'MbtSuccess' } 'warn' { 'MbtCaution' } 'bad' { 'MbtCritical' } default { 'TextFillColorSecondaryBrush' } }
    $b = [Windows.Controls.Border]::new()
    $b.CornerRadius = [Windows.CornerRadius]::new(4)
    $b.Padding = [Windows.Thickness]::new(10, 8, 10, 8)
    $b.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    Set-MbtGuiResource $b ([Windows.Controls.Border]::BackgroundProperty) $bg
    $g = [Windows.Controls.Grid]::new()
    $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::new(24)
    $c2 = [Windows.Controls.ColumnDefinition]::new(); $c2.Width = [Windows.GridLength]::new(1, 'Star')
    $g.ColumnDefinitions.Add($c1); $g.ColumnDefinitions.Add($c2)
    $glyph = New-GuiIcon $Icon 13 $ic
    $glyph.VerticalAlignment = 'Top'; $glyph.Margin = [Windows.Thickness]::new(0, 2, 0, 0)
    $g.Children.Add($glyph) | Out-Null
    $inner = [Windows.Controls.StackPanel]::new()
    [Windows.Controls.Grid]::SetColumn($inner, 1)
    $t = [Windows.Controls.TextBlock]::new()
    ConvertTo-MbtGuiRuns $t $Text
    $t.FontSize = 12; $t.TextWrapping = 'Wrap'
    Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $fg
    $inner.Children.Add($t) | Out-Null
    foreach ($a in $After) { if ($a) { [void]$inner.Children.Add($a) } }
    $g.Children.Add($inner) | Out-Null
    $b.Child = $g
    return $b
}

function ConvertTo-MbtGuiHandler {
    <#
        The handler of a control, written as a script block in the function that builds the page. The control raises its event after that function
        has returned, so the block gets a copy of the variables of that function (taken when the control is made, like a closure) and keeps its own
        scope, where the functions and the script variables of the console are. A block that is already a closure is left as it is.
        -CallerScope: how many scopes up from this function the function that wrote the block is (2 when a New-Gui... function of this file calls it, 1 when the builder calls it).
    #>
    param([Parameter(Mandatory)][scriptblock]$Block, [int]$CallerScope = 2)
    if ($Block.Module) { return $Block }
    $skip = @('args', 'input', 'this', 'PSCmdlet', 'MyInvocation', 'PSBoundParameters', 'PSCommandPath', 'PSScriptRoot', 'foreach', 'switch', 'Error', '_', 'ExecutionContext', 'Host', 'StackTrace', 'null', 'true', 'false', '?', '^', '$')
    $locked = [Management.Automation.ScopedItemOptions]::Constant -bor [Management.Automation.ScopedItemOptions]::ReadOnly -bor [Management.Automation.ScopedItemOptions]::AllScope
    $vars = [Collections.Generic.List[psvariable]]::new()
    foreach ($v in @(Get-Variable -Scope $CallerScope -ErrorAction SilentlyContinue)) {
        if ($v.Name -notin $skip -and -not ($v.Options -band $locked)) { $vars.Add([psvariable]::new($v.Name, $v.Value)) }
    }
    $inner = $Block
    return { param($sender, $e) $null = $inner.InvokeWithContext($null, $vars, @($sender, $e)) }.GetNewClosure()
}

function Add-MbtGuiHandler {
    <# Subscribes a script block to an event of a control (Click, Checked, SelectionChanged...). Every handler of the console goes through here or through a New-Gui... function: see ConvertTo-MbtGuiHandler. #>
    param([Parameter(Mandatory)]$Control, [Parameter(Mandatory)][string]$Event, [Parameter(Mandatory)][scriptblock]$Block)
    $handler = ConvertTo-MbtGuiHandler -Block $Block -CallerScope 2
    $null = $Control.('Add_' + $Event).Invoke($handler)
}

function New-GuiButton {
    param(
        [Parameter(Mandatory)][string]$Text, [scriptblock]$OnClick, [switch]$Accent, [switch]$Small, [switch]$Danger, [string]$Icon, [switch]$Disabled, [string]$Name, [double]$MinWidth = 0
    )
    $b = [Windows.Controls.Button]::new()
    if ($Name) { $b.Name = $Name }
    if ($Icon) {
        $sp = [Windows.Controls.StackPanel]::new(); $sp.Orientation = 'Horizontal'
        $ic = New-GuiIcon $Icon 13 $(if ($Accent) { 'TextOnAccentFillColorPrimaryBrush' } else { 'MbtBrandText' })
        $ic.Margin = [Windows.Thickness]::new(0, 1, 7, 0)
        $sp.Children.Add($ic) | Out-Null
        $tx = [Windows.Controls.TextBlock]::new(); $tx.Text = $Text
        $sp.Children.Add($tx) | Out-Null
        $b.Content = $sp
    } else { $b.Content = $Text }
    if ($Small) { $b.Padding = [Windows.Thickness]::new(10, 2, 10, 3); $b.FontSize = 13 }
    if ($Accent) { $b.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'AccentButtonStyle') }
    if ($Danger) { Set-MbtGuiResource $b ([Windows.Controls.Control]::BorderBrushProperty) 'MbtCritical'; Set-MbtGuiResource $b ([Windows.Controls.Control]::ForegroundProperty) 'MbtCriticalInk' }
    if ($MinWidth) { $b.MinWidth = $MinWidth }
    $b.IsEnabled = -not $Disabled
    $b.Margin = [Windows.Thickness]::new(0, 0, 8, 0)
    if ($OnClick) { $b.Add_Click((ConvertTo-MbtGuiHandler $OnClick)) }
    return $b
}

function New-GuiLink {
    param([Parameter(Mandatory)][string]$Text, [scriptblock]$OnClick)
    $b = [Windows.Controls.Button]::new()
    $b.Content = $Text
    $b.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtLink')
    if ($OnClick) { $b.Add_Click((ConvertTo-MbtGuiHandler $OnClick)) }
    return $b
}

function New-GuiTextBox {
    param([string]$Text = '', [string]$Placeholder, [switch]$ReadOnly, [scriptblock]$OnChange, [string]$Name, [string]$Automation)
    $t = [Windows.Controls.TextBox]::new()
    $t.Text = $Text
    if ($ReadOnly) { $t.IsReadOnly = $true }
    if ($Name) { $t.Name = $Name }
    if ($Automation) { [Windows.Automation.AutomationProperties]::SetName($t, $Automation) }
    if ($Placeholder) { $t.Tag = $Placeholder; try { $t.SetValue([Windows.Controls.TextBox]::TextProperty, $Text) } catch { } }
    if ($OnChange) { $t.Add_TextChanged((ConvertTo-MbtGuiHandler $OnChange)) }
    return $t
}

function New-GuiRow {
    <# A text box and a button side by side (folder and Browse...). #>
    param([Parameter(Mandatory)][Windows.Controls.TextBox]$Box, [Parameter(Mandatory)][Windows.Controls.Button]$Button)
    $g = [Windows.Controls.Grid]::new()
    $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::new(1, 'Star')
    $c2 = [Windows.Controls.ColumnDefinition]::new(); $c2.Width = [Windows.GridLength]::Auto
    $g.ColumnDefinitions.Add($c1); $g.ColumnDefinitions.Add($c2)
    $Button.Margin = [Windows.Thickness]::new(8, 0, 0, 0)
    [Windows.Controls.Grid]::SetColumn($Button, 1)
    [void]$g.Children.Add($Box); [void]$g.Children.Add($Button)
    return $g
}

function New-GuiOption {
    <# A radio button or a check box with a title and a hint under it. #>
    param([ValidateSet('Radio', 'Check')][string]$Type = 'Radio', [string]$Group, [Parameter(Mandatory)][string]$Title, [string]$Hint, [bool]$Checked = $false, [scriptblock]$OnChange, [switch]$Disabled, [string]$Margin = '0,10,0,0')
    $c = if ($Type -eq 'Radio') { $r = [Windows.Controls.RadioButton]::new(); if ($Group) { $r.GroupName = $Group }; $r } else { [Windows.Controls.CheckBox]::new() }
    $sp = [Windows.Controls.StackPanel]::new()
    $sp.Children.Add((New-GuiText $Title 'Body')) | Out-Null
    if ($Hint) { $h = New-GuiText $Hint 'Hint'; $h.Margin = [Windows.Thickness]::new(0, 1, 0, 0); $sp.Children.Add($h) | Out-Null }
    $c.Content = $sp
    $c.IsChecked = $Checked
    $c.IsEnabled = -not $Disabled
    $c.Margin = ConvertTo-MbtGuiThickness ($Margin)
    $c.VerticalContentAlignment = 'Top'
    $c.Padding = [Windows.Thickness]::new(8, 0, 0, 0)
    if ($OnChange) { $handler = ConvertTo-MbtGuiHandler $OnChange; $c.Add_Checked($handler); if ($Type -eq 'Check') { $c.Add_Unchecked($handler) } }
    return $c
}

function New-GuiMore {
    <# A collapsed "More options" section: a thin rule above, an accent header. #>
    param([Parameter(Mandatory)][string]$Summary, [Parameter(Mandatory)][object[]]$Content, [switch]$Open)
    $e = [Windows.Controls.Expander]::new()
    $e.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtMore')
    $h = [Windows.Controls.TextBlock]::new(); $h.Text = $Summary; $h.FontSize = 13
    Set-MbtGuiResource $h ([Windows.Controls.TextBlock]::ForegroundProperty) 'AccentTextFillColorPrimaryBrush'
    $e.Header = $h
    $e.Content = (New-GuiStack $Content -Margin '0,6,0,0')
    $e.IsExpanded = [bool]$Open
    return $e
}

function New-GuiCard {
    <# A card: title, an information text beside it, a control at the right (a pill), then the content. #>
    param([string]$Title, [string]$Info, $Right, [object[]]$Content = @(), [string]$Margin, [string]$BorderKey)
    $card = [Windows.Controls.Border]::new()
    $card.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtCard')
    if ($Margin) { $card.Margin = ConvertTo-MbtGuiThickness ($Margin) }
    if ($BorderKey) { Set-MbtGuiResource $card ([Windows.Controls.Border]::BorderBrushProperty) $BorderKey }
    $stack = [Windows.Controls.StackPanel]::new()
    if ($Title -or $Info -or $Right) {
        $head = [Windows.Controls.DockPanel]::new()
        $head.Margin = [Windows.Thickness]::new(0, 0, 0, 8)
        if ($Right) { [Windows.Controls.DockPanel]::SetDock($Right, 'Right'); $Right.Margin = [Windows.Thickness]::new(12, 0, 0, 0); [void]$head.Children.Add($Right) }
        $t = New-GuiText $Title 'Title' -NoWrap
        [Windows.Controls.DockPanel]::SetDock($t, 'Left')
        [void]$head.Children.Add($t)
        if ($Info) {
            $i = New-GuiText $Info 'Small'
            $i.Margin = [Windows.Thickness]::new(12, 1, 0, 0); $i.TextTrimming = 'CharacterEllipsis'; $i.TextWrapping = 'NoWrap'
            [void]$head.Children.Add($i)
        }
        [void]$stack.Children.Add($head)
    }
    foreach ($c in @($Content)) { if ($null -ne $c) { [void]$stack.Children.Add($c) } }
    $card.Child = $stack
    return $card
}

function New-GuiFacts {
    <# Label / value pairs: small grey label, normal value. #>
    param([Parameter(Mandatory)][object[]]$Pairs)
    $g = [Windows.Controls.Grid]::new()
    $g.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::Auto
    $c2 = [Windows.Controls.ColumnDefinition]::new(); $c2.Width = [Windows.GridLength]::new(1, 'Star')
    $g.ColumnDefinitions.Add($c1); $g.ColumnDefinitions.Add($c2)
    $row = 0
    foreach ($p in $Pairs) {
        $g.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
        $k = New-GuiText ([string]$p[0]) 'Small'; $k.Margin = [Windows.Thickness]::new(0, 3, 14, 0); $k.VerticalAlignment = 'Top'
        [Windows.Controls.Grid]::SetRow($k, $row)
        $v = if ($p[1] -is [Windows.UIElement]) { $p[1] } else { New-GuiText ([string]$p[1]) 'Body' }
        $v.FontSize = 13
        $v.Margin = [Windows.Thickness]::new(0, 2, 0, 0)
        [Windows.Controls.Grid]::SetRow($v, $row); [Windows.Controls.Grid]::SetColumn($v, 1)
        [void]$g.Children.Add($k); [void]$g.Children.Add($v)
        $row++
    }
    return $g
}

function New-GuiTable {
    <# A table with a thin border: Headers, Rows (cells are strings or elements), Widths ('*' or pixels). Right = column indexes aligned right. #>
    param([Parameter(Mandatory)][string[]]$Headers, [object[]]$Rows = @(), [string[]]$Widths, [int[]]$Right = @())
    $outer = [Windows.Controls.Border]::new()
    $outer.BorderThickness = [Windows.Thickness]::new(1); $outer.CornerRadius = [Windows.CornerRadius]::new(4)
    $outer.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    Set-MbtGuiResource $outer ([Windows.Controls.Border]::BorderBrushProperty) 'DividerStrokeColorDefaultBrush'
    $g = [Windows.Controls.Grid]::new()
    for ($c = 0; $c -lt $Headers.Count; $c++) {
        $cd = [Windows.Controls.ColumnDefinition]::new()
        $w = if ($Widths -and $c -lt $Widths.Count) { $Widths[$c] } else { '*' }
        $cd.Width = if ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } elseif ($w -eq 'Auto') { [Windows.GridLength]::Auto } else { [Windows.GridLength]::new([double]$w) }
        $g.ColumnDefinitions.Add($cd)
    }
    $g.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
    for ($c = 0; $c -lt $Headers.Count; $c++) {
        $h = New-GuiText $Headers[$c] 'Small'; $h.Margin = [Windows.Thickness]::new(10, 6, 10, 6)
        if ($c -in $Right) { $h.TextAlignment = 'Right' }
        [Windows.Controls.Grid]::SetColumn($h, $c); [void]$g.Children.Add($h)
    }
    $r = 1
    foreach ($row in $Rows) {
        $g.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
        $rule = [Windows.Controls.Border]::new(); $rule.BorderThickness = [Windows.Thickness]::new(0, 1, 0, 0)
        Set-MbtGuiResource $rule ([Windows.Controls.Border]::BorderBrushProperty) 'DividerStrokeColorDefaultBrush'
        [Windows.Controls.Grid]::SetRow($rule, $r); [Windows.Controls.Grid]::SetColumnSpan($rule, $Headers.Count); [void]$g.Children.Add($rule)
        for ($c = 0; $c -lt $Headers.Count; $c++) {
            $cell = if ($c -lt $row.Count) { $row[$c] } else { '' }
            $el = if ($cell -is [Windows.UIElement]) { $cell } else { $x = New-GuiText ([string]$cell) 'Body'; $x.FontSize = 13; $x }
            if ($el -is [Windows.Controls.TextBlock] -and $c -in $Right) { $el.TextAlignment = 'Right' }
            $el.Margin = [Windows.Thickness]::new(10, 6, 10, 6)
            [Windows.Controls.Grid]::SetRow($el, $r); [Windows.Controls.Grid]::SetColumn($el, $c); [void]$g.Children.Add($el)
        }
        $r++
    }
    $outer.Child = $g
    return $outer
}

function New-GuiTile {
    <# A result tile: big number, small label, tone ok / warn / bad / ''. #>
    param([string]$Value, [string]$Label, [ValidateSet('', 'ok', 'warn', 'bad')][string]$Tone = '')
    $b = [Windows.Controls.Border]::new()
    $b.BorderThickness = [Windows.Thickness]::new(1); $b.CornerRadius = [Windows.CornerRadius]::new(8); $b.Padding = [Windows.Thickness]::new(12, 10, 12, 10)
    $b.Margin = [Windows.Thickness]::new(0, 0, 10, 0); $b.MinWidth = 150
    Set-MbtGuiResource $b ([Windows.Controls.Border]::BorderBrushProperty) 'DividerStrokeColorDefaultBrush'
    $v = New-GuiText $Value 'Body' -NoWrap; $v.FontSize = 22; $v.FontWeight = [Windows.FontWeights]::SemiBold
    $key = switch ($Tone) { 'ok' { 'MbtSuccessInk' } 'warn' { 'MbtCautionInk' } 'bad' { 'MbtCriticalInk' } default { $null } }
    if ($key) { Set-MbtGuiResource $v ([Windows.Controls.TextBlock]::ForegroundProperty) $key }
    $k = New-GuiText $Label 'Small'
    $b.Child = (New-GuiStack @($v, $k))
    return $b
}

function New-GuiTwoColumns {
    <# The house window: left column (410 px at most), right area. #>
    param([Parameter(Mandatory)]$Left, [Parameter(Mandatory)]$Right, [double]$LeftWidth = 410, [string]$RightWidth = '*')
    $g = [Windows.Controls.Grid]::new()
    $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::new($LeftWidth)
    $c2 = [Windows.Controls.ColumnDefinition]::new(); $c2.Width = [Windows.GridLength]::new(16)
    $c3 = [Windows.Controls.ColumnDefinition]::new(); $c3.Width = [Windows.GridLength]::new(1, 'Star')
    $g.ColumnDefinitions.Add($c1); $g.ColumnDefinitions.Add($c2); $g.ColumnDefinitions.Add($c3)
    [Windows.Controls.Grid]::SetColumn($Right, 2)
    [void]$g.Children.Add($Left); [void]$g.Children.Add($Right)
    return $g
}

function New-GuiGrid {
    <# Equal columns (or the given widths) of elements, 12 px apart. #>
    param([Parameter(Mandatory)][object[]]$Children, [string[]]$Widths)
    $g = [Windows.Controls.Grid]::new()
    $n = $Children.Count
    for ($i = 0; $i -lt $n; $i++) {
        if ($i -gt 0) { $sp = [Windows.Controls.ColumnDefinition]::new(); $sp.Width = [Windows.GridLength]::new(12); $g.ColumnDefinitions.Add($sp) }
        $cd = [Windows.Controls.ColumnDefinition]::new()
        $w = if ($Widths -and $i -lt $Widths.Count) { $Widths[$i] } else { '*' }
        $cd.Width = if ($w -match '^\d+(\.\d+)?\*$') { [Windows.GridLength]::new([double]$w.TrimEnd('*'), 'Star') } elseif ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } else { [Windows.GridLength]::new([double]$w) }
        $g.ColumnDefinitions.Add($cd)
        $el = $Children[$i]
        [Windows.Controls.Grid]::SetColumn($el, $i * 2)
        [void]$g.Children.Add($el)
    }
    return $g
}

function New-GuiProgressBar {
    <# A thin bar (4 px) or the big one (8 px). Tone: run (accent), ok, warn. Fraction 0..1. #>
    param([double]$Fraction = 0, [ValidateSet('run', 'ok', 'warn')][string]$Tone = 'run', [double]$Height = 4)
    $track = [Windows.Controls.Border]::new()
    $track.Height = $Height; $track.CornerRadius = [Windows.CornerRadius]::new($Height / 2)
    Set-MbtGuiResource $track ([Windows.Controls.Border]::BackgroundProperty) 'DividerStrokeColorDefaultBrush'
    $g = [Windows.Controls.Grid]::new()
    $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::new([math]::Max(0.0001, [math]::Min(1.0, $Fraction)), 'Star')
    $c2 = [Windows.Controls.ColumnDefinition]::new(); $c2.Width = [Windows.GridLength]::new([math]::Max(0.0001, 1.0 - [math]::Min(1.0, $Fraction)), 'Star')
    $g.ColumnDefinitions.Add($c1); $g.ColumnDefinitions.Add($c2)
    $fill = [Windows.Controls.Border]::new(); $fill.CornerRadius = [Windows.CornerRadius]::new($Height / 2)
    $key = switch ($Tone) { 'ok' { 'MbtSuccess' } 'warn' { 'MbtCaution' } default { 'AccentFillColorDefaultBrush' } }
    Set-MbtGuiResource $fill ([Windows.Controls.Border]::BackgroundProperty) $key
    if ($Fraction -le 0) { $fill.Visibility = 'Hidden' }
    [void]$g.Children.Add($fill)
    $track.Child = $g
    return $track
}

#endregion

#region Pickers -----------------------------------------------------------------------------

function Set-MbtGuiClipboard {
    <# Puts a text on the clipboard. Another program may hold the clipboard for a moment: that is not an error of the console, the button then does not say "Copied". #>
    param([string]$Text)
    try { [Windows.Clipboard]::SetText($Text); return $true } catch { Write-Verbose "clipboard: $($_.Exception.Message)"; return $false }
}

function Select-MbtGuiFile {
    param([string]$Title, [string]$Filter = 'All files (*.*)|*.*', [string]$InitialPath, [switch]$Save)
    $d = if ($Save) { [Microsoft.Win32.SaveFileDialog]::new() } else { [Microsoft.Win32.OpenFileDialog]::new() }
    $d.Title = $Title; $d.Filter = $Filter
    if ($InitialPath -and (Test-Path -LiteralPath (Split-Path $InitialPath -Parent -ErrorAction SilentlyContinue) -ErrorAction SilentlyContinue)) { $d.InitialDirectory = Split-Path $InitialPath -Parent }
    if ($d.ShowDialog($script:Gui.Window)) { return $d.FileName }
    return $null
}

function Select-MbtGuiFolder {
    param([string]$Title, [string]$InitialPath)
    $d = [Microsoft.Win32.OpenFolderDialog]::new()
    $d.Title = $Title
    if ($InitialPath -and (Test-Path -LiteralPath $InitialPath)) { $d.InitialDirectory = $InitialPath }
    if ($d.ShowDialog($script:Gui.Window)) { return $d.FolderName }
    return $null
}

#endregion

#region Background worker --------------------------------------------------------------------

function New-MbtGuiWorker {
    <# One runspace that runs the engine calls one after the other (the sign-in session lives in it). #>
    param([string]$ModulePath)
    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'MTA'
    $rs.Open()
    $w = [pscustomobject]@{ Runspace = $rs; Queue = [Collections.Generic.Queue[object]]::new(); Current = $null; ModulePath = $ModulePath; Ready = $false; Inline = $false }
    return $w
}

function Start-MbtGuiWork {
    <#
        Queues a script block on the worker. OnDone gets the result array, OnError the error text; both run on the UI thread (Step-MbtGuiWork, a
        100 ms timer). With a worker that is Inline (the fake engine of the tests and of the documentation scenarios) the script runs at once.
    #>
    param([Parameter(Mandatory)]$Worker, [Parameter(Mandatory)][scriptblock]$Script, [object[]]$Arguments = @(), [scriptblock]$OnDone, [scriptblock]$OnError)
    $job = [pscustomobject]@{ Script = $Script; Arguments = $Arguments; OnDone = $OnDone; OnError = $OnError; Shell = $null; Handle = $null }
    if ($Worker.Inline) {
        try { $result = @(& $Script @Arguments) } catch { if ($OnError) { & $OnError $_.Exception.Message }; return }
        if ($OnDone) { & $OnDone $result }
        return
    }
    $Worker.Queue.Enqueue($job)
    Step-MbtGuiWork $Worker
}

function Write-MbtGuiFault {
    <# A step of the console that failed after an engine call: written to the log (with where it failed), so that a page that stops changing can be explained. #>
    param([string]$Where, $Record)
    Write-Verbose "${Where}: $Record"
    if (Get-Command Write-MbtLog -ErrorAction SilentlyContinue) {
        try { Write-MbtLog -Level FAIL -Message ('console: {0} failed: {1} {2}' -f $Where, $Record, ($Record.ScriptStackTrace -split "`r?`n" | Select-Object -First 1)) } catch { Write-Verbose "log: $_" }
    }
}

function Step-MbtGuiWork {
    param([Parameter(Mandatory)]$Worker)
    if ($Worker.Inline) { return }
    $cur = $Worker.Current
    if ($cur -and $cur.Handle.IsCompleted) {
        $Worker.Current = $null
        $err = $null; $result = @()
        try {
            $result = @($cur.Shell.EndInvoke($cur.Handle))
            if ($cur.Shell.HadErrors) { $err = ($cur.Shell.Streams.Error | Select-Object -First 1).ToString() }
        } catch { $err = $_.Exception.Message }
        $cur.Shell.Dispose()
        if ($err) { if ($cur.OnError) { try { & $cur.OnError $err } catch { Write-MbtGuiFault 'OnError' $_ } } }
        elseif ($cur.OnDone) { try { & $cur.OnDone $result } catch { Write-MbtGuiFault 'OnDone' $_ } }
    }
    if (-not $Worker.Current -and $Worker.Queue.Count -gt 0) {
        $job = $Worker.Queue.Dequeue()
        $ps = [powershell]::Create()
        $ps.Runspace = $Worker.Runspace
        $prefix = ''
        if (-not $Worker.Ready -and $Worker.ModulePath) {
            $prefix = "Set-StrictMode -Version Latest; `$ErrorActionPreference = 'Stop'; Import-Module -Name '$($Worker.ModulePath.Replace("'", "''"))' -Force; Initialize-MbtEngine`n"
            $Worker.Ready = $true
        }
        if ($prefix) { [void]$ps.AddScript($prefix, $true).AddStatement() }
        [void]$ps.AddScript($job.Script.ToString(), $true)
        foreach ($a in $job.Arguments) { [void]$ps.AddArgument($a) }
        $job.Shell = $ps
        $job.Handle = $ps.BeginInvoke()
        $Worker.Current = $job
    }
}

function Stop-MbtGuiWorker {
    param($Worker)
    if (-not $Worker -or $Worker.Inline) { return }
    try { if ($Worker.Current) { $Worker.Current.Shell.Stop() } } catch { Write-Verbose "stop: $_" }
    try { $Worker.Runspace.Dispose() } catch { Write-Verbose "dispose: $_" }
}

#endregion

