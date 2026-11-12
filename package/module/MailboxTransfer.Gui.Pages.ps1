#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - the pages Setup, Task (What, From, To), Verify (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    One function per page: each returns the element of the page and sets the action bar. A page is built again when the state changes (the browse
    cards live on, see Gui.Browse.ps1).
    Features that depend on the engine (PST import, EML, shared mailboxes) read $script:Gui.Engine.Capability: a missing one is shown disabled with
    "Coming in this version", or left out, so the page lights up by itself after the merge of the engine work.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

$script:PageBuilders = @{
    setup = { Build-MbtGuiSetupPage }
    start = { Build-MbtGuiStartPage }
    exFrom = { Build-MbtGuiExportFromPage }
    exTo = { Build-MbtGuiExportToPage }
    exPlan = { Build-MbtGuiExportPlanPage }
    imFrom = { Build-MbtGuiImportFromPage }
    imTo = { Build-MbtGuiImportToPage }
    imPlan = { Build-MbtGuiImportPlanPage }
    rpFrom = { Build-MbtGuiRepairFromPage }
    rpTo = { Build-MbtGuiRepairToPage }
    rpPlan = { Build-MbtGuiRepairPlanPage }
    opFrom = { Build-MbtGuiOpenPage }
    cpTo = { Build-MbtGuiCopyToPage }
    cpPlan = { Build-MbtGuiCopyPlanPage }
    run = { Build-MbtGuiRunPage }
    done = { Build-MbtGuiDonePage }
}

#region Shared pieces ------------------------------------------------------------------------------

function Get-MbtGuiCapability {
    param([string]$Name)
    $cap = $script:Gui.Engine.Capability
    return [bool]($cap.ContainsKey($Name) -and $cap[$Name])
}

function New-MbtGuiBrowseFolderRow {
    <# A text box with Browse...: a folder or a file; ReadOnly: no. #>
    param([string]$Value, [ValidateSet('Folder', 'File', 'SaveFile')][string]$Kind, [string]$Title, [string]$Filter, [scriptblock]$OnChange, [string]$Automation)
    $box = New-GuiTextBox -Text $Value -Automation $Automation
    $state = @{ Box = $box; OnChange = $OnChange }
    if ($OnChange) {
        $state.OnChange = ConvertTo-MbtGuiHandler -Block $OnChange -CallerScope 2
        Add-MbtGuiHandler $box 'LostFocus' { & $state.OnChange $state.Box.Text }.GetNewClosure()
        Add-MbtGuiHandler $box 'KeyDown' { param($sender, $e) if ($e.Key -eq 'Return') { & $state.OnChange $state.Box.Text; $e.Handled = $true } }.GetNewClosure()
    }
    $btn = New-GuiButton -Text 'Browse...' -OnClick {
        $picked = switch ($Kind) {
            'Folder' { Select-MbtGuiFolder -Title $Title -InitialPath $state.Box.Text }
            'SaveFile' { Select-MbtGuiFile -Title $Title -Filter $Filter -InitialPath $state.Box.Text -Save }
            default { Select-MbtGuiFile -Title $Title -Filter $Filter -InitialPath $state.Box.Text }
        }
        if ($picked) { $state.Box.Text = $picked; if ($state.OnChange) { & $state.OnChange $picked } }
    }.GetNewClosure()
    $btn.Margin = [Windows.Thickness]::new(0)
    return (New-GuiRow -Box $box -Button $btn)
}

function Get-MbtGuiBrowser {
    <# The browse card of a use (export, import, open), created once. #>
    param([ValidateSet('export', 'import', 'open')][string]$Use)
    $g = $script:Gui
    if ($g.Browsers.ContainsKey($Use)) { return $g.Browsers[$Use] }
    switch ($Use) {
        'export' { return (New-MbtGuiBrowser -Name 'export' -Title 'Browse' -Info 'Live view of the mailbox. Read only.' -TrayEmpty 'Nothing ticked yet.') }
        'import' { return (New-MbtGuiBrowser -Name 'import' -Title 'What is inside' -Info 'Read only. The default is to import everything. Tick only if you want part of it.' -Badge 'Read only' -TrayEmpty 'Nothing ticked: **everything will be imported**.' -Kind source) }
        'open' { return (New-MbtGuiBrowser -Name 'open' -Title 'Browse' -Info 'Folders, messages and attachments, like a mailbox.' -Badge 'Read only' -TrayEmpty 'Nothing ticked yet.' -Kind source) }
    }
}

function Place-MbtGuiBrowser {
    <# Puts the card of a browser in a page (it may still belong to the page that was replaced). #>
    param($Browser, [Windows.Controls.Panel]$Into)
    $root = $Browser.Root
    $parent = $root.Parent
    if ($parent -is [Windows.Controls.Panel]) { $parent.Children.Remove($root) }
    elseif ($parent -is [Windows.Controls.Decorator]) { $parent.Child = $null }
    elseif ($parent -is [Windows.Controls.ContentControl]) { $parent.Content = $null }
    [void]$Into.Children.Add($root)
}

function Format-MbtGuiSelectionFooter {
    param($Selection, [bool]$Signed)
    if ($Selection.Items -gt 0) { return ('**{0} items** ticked in {1} folder{2}, {3}' -f (Format-MbtGuiSelectionItems $Selection), $Selection.Folders, $(if ($Selection.Folders -gt 1) { 's' } else { '' }), (Format-MbtGuiSelectionSize $Selection)) }
    if ($Signed) { return 'Tick what to keep, then continue.' }
    return 'Sign in first.'
}

function New-MbtGuiSelectionCard {
    param($Selection)
    if ($Selection.Items -gt 0) {
        $body = New-GuiFacts @(@('Items', (Format-MbtGuiSelectionItems $Selection)), @('Size', (Format-MbtGuiSelectionSize $Selection)), @('Folders', [string]$Selection.Folders))
    } else { $body = New-GuiText 'Tick a folder or some messages in the browser.' 'Hint' }
    return (New-GuiCard -Title 'Your selection' -Content @($body))
}

#endregion

#region Start ----------------------------------------------------------------------------------------

function New-MbtGuiTaskCard {
    param([string]$Key, [bool]$Selected)
    $t = $script:Tasks[$Key]
    $btn = [Windows.Controls.Primitives.ToggleButton]::new()
    $btn.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtTaskCard')
    $btn.IsChecked = $Selected
    $btn.Tag = $Key
    [Windows.Automation.AutomationProperties]::SetName($btn, $t.Title)
    $g = [Windows.Controls.Grid]::new()
    $head = [Windows.Controls.StackPanel]::new(); $head.Orientation = 'Horizontal'
    $tile = [Windows.Controls.Border]::new(); $tile.Width = 40; $tile.Height = 40; $tile.CornerRadius = [Windows.CornerRadius]::new(8)
    Set-MbtGuiResource $tile ([Windows.Controls.Border]::BackgroundProperty) 'MbtAccentSoft'
    $ic = New-GuiIcon $t.Icon 22; $ic.HorizontalAlignment = 'Center'; $tile.Child = $ic
    $titles = [Windows.Controls.StackPanel]::new(); $titles.Margin = [Windows.Thickness]::new(12, 0, 0, 0); $titles.VerticalAlignment = 'Center'
    $eb = New-GuiText $t.Eyebrow 'Small' -NoWrap; $eb.FontWeight = [Windows.FontWeights]::SemiBold; Set-MbtGuiResource $eb ([Windows.Controls.TextBlock]::ForegroundProperty) 'MbtBrandText'
    $ti = New-GuiText $t.Title 'Body' -NoWrap; $ti.FontSize = 15; $ti.FontWeight = [Windows.FontWeights]::SemiBold
    $titles.Children.Add($eb) | Out-Null; $titles.Children.Add($ti) | Out-Null
    $head.Children.Add($tile) | Out-Null; $head.Children.Add($titles) | Out-Null
    $desc = New-GuiText $t.Text 'Muted'; $desc.FontSize = 13; $desc.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
    $col = New-GuiStack @($head, $desc)
    $g.Children.Add($col) | Out-Null
    $tick = [Windows.Controls.Border]::new(); $tick.Width = 22; $tick.Height = 22; $tick.CornerRadius = [Windows.CornerRadius]::new(11); $tick.HorizontalAlignment = 'Right'; $tick.VerticalAlignment = 'Top'
    Set-MbtGuiResource $tick ([Windows.Controls.Border]::BackgroundProperty) 'AccentFillColorDefaultBrush'
    $ti2 = New-GuiIcon 'check' 12 'TextOnAccentFillColorPrimaryBrush'; $ti2.HorizontalAlignment = 'Center'; $tick.Child = $ti2
    $tick.Visibility = $(if ($Selected) { 'Visible' } else { 'Collapsed' })
    $g.Children.Add($tick) | Out-Null
    $btn.Content = $g
    return $btn
}

function New-MbtGuiJobList {
    $g = $script:Gui
    $jobs = @()
    try { $jobs = @(& $g.Engine.Jobs $g.S.WorkFolder) } catch { Write-Verbose "jobs: $_" }
    if (-not $jobs.Count) { return (New-GuiText 'No job yet. A job you stop, or that is not finished, is listed here so you can continue it.' 'Hint') }
    $stack = [Windows.Controls.StackPanel]::new()
    $first = $true
    foreach ($j in $jobs) {
        $row = [Windows.Controls.Grid]::new()
        $row.Margin = [Windows.Thickness]::new(0, 0, 0, 0)
        foreach ($w in 'Auto', '*', 'Auto') { $cd = [Windows.Controls.ColumnDefinition]::new(); $cd.Width = if ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } else { [Windows.GridLength]::Auto }; $row.ColumnDefinitions.Add($cd) }
        $glyphName = switch ($j.State) { 'ok' { 'check' } 'warn' { 'warn' } default { 'pause' } }
        $brush = switch ($j.State) { 'ok' { 'MbtSuccess' } 'warn' { 'MbtCaution' } default { 'MbtBrandText' } }
        $ic = New-GuiIcon $glyphName 14 $brush; $ic.Margin = [Windows.Thickness]::new(0, 0, 12, 0); $ic.VerticalAlignment = 'Center'
        $txt = [Windows.Controls.StackPanel]::new(); [Windows.Controls.Grid]::SetColumn($txt, 1)
        $t1 = New-GuiText $j.Title 'Body'; $t1.FontSize = 13; $t1.FontWeight = [Windows.FontWeights]::SemiBold
        $t2 = New-GuiText $j.Sub 'Small'
        $txt.Children.Add($t1) | Out-Null; $txt.Children.Add($t2) | Out-Null
        $btn = New-GuiButton -Text $j.Button -Small -OnClick { Resume-MbtGuiJob $j }.GetNewClosure()
        $btn.Margin = [Windows.Thickness]::new(12, 0, 0, 0); $btn.VerticalAlignment = 'Center'
        [Windows.Controls.Grid]::SetColumn($btn, 2)
        $row.Children.Add($ic) | Out-Null; $row.Children.Add($txt) | Out-Null; $row.Children.Add($btn) | Out-Null
        $wrap = [Windows.Controls.Border]::new(); $wrap.Child = $row; $wrap.Padding = [Windows.Thickness]::new(0, 9, 0, 9)
        if (-not $first) { $wrap.BorderThickness = [Windows.Thickness]::new(0, 1, 0, 0); Set-MbtGuiResource $wrap ([Windows.Controls.Border]::BorderBrushProperty) 'DividerStrokeColorDefaultBrush' }
        $stack.Children.Add($wrap) | Out-Null
        $first = $false
    }
    return $stack
}

function New-MbtGuiCheckRow {
    <# One line of a list of lights or checks: a green dot (ok), a red one (bad) or a grey clock (wait), the check in bold and what was found after it. #>
    param([ValidateSet('ok', 'bad', 'wait')][string]$State, [Parameter(Mandatory)][string]$Text, [string]$Detail = '')
    $row = [Windows.Controls.Grid]::new(); $row.Margin = [Windows.Thickness]::new(0, 4, 0, 4)
    $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::Auto
    $c2 = [Windows.Controls.ColumnDefinition]::new(); $c2.Width = [Windows.GridLength]::new(1, 'Star')
    $row.ColumnDefinitions.Add($c1); $row.ColumnDefinitions.Add($c2)
    $dot = [Windows.Controls.Border]::new(); $dot.Width = 18; $dot.Height = 18; $dot.CornerRadius = [Windows.CornerRadius]::new(9); $dot.VerticalAlignment = 'Top'; $dot.Margin = [Windows.Thickness]::new(0, 1, 0, 0)
    Set-MbtGuiResource $dot ([Windows.Controls.Border]::BackgroundProperty) $(switch ($State) { 'ok' { 'MbtSuccessBackground' } 'bad' { 'MbtCriticalBackground' } default { 'MbtInfoBackground' } })
    $ic = New-GuiIcon $(switch ($State) { 'ok' { 'check' } 'bad' { 'x' } default { 'clock' } }) 10 $(switch ($State) { 'ok' { 'MbtSuccess' } 'bad' { 'MbtCritical' } default { 'TextFillColorSecondaryBrush' } })
    $ic.HorizontalAlignment = 'Center'; $dot.Child = $ic
    $tx = New-GuiText ("**$Text**" + $(if ($Detail) { '  ' + [char]0x00B7 + " $Detail" } else { '' })) 'Body'; $tx.FontSize = 13; $tx.Margin = [Windows.Thickness]::new(8, 0, 0, 0)
    [Windows.Controls.Grid]::SetColumn($tx, 1)
    $row.Children.Add($dot) | Out-Null; $row.Children.Add($tx) | Out-Null
    return $row
}

function New-MbtGuiReadinessCard {
    <# The lights of this computer, all shown: PowerShell, sign-in component, work folder and its room, network, unblocked files. #>
    $g = $script:Gui
    $lights = @()
    try { $lights = @((& $g.Engine.Readiness).Lights) } catch { $lights = @(@{ Name = 'The engine could not be checked'; Detail = $_.Exception.Message; Ok = $false }) }
    $red = @($lights | Where-Object { -not $_.Ok })
    $list = [Windows.Controls.StackPanel]::new()
    foreach ($l in $lights) { $list.Children.Add((New-MbtGuiCheckRow -State $(if ($l.Ok) { 'ok' } else { 'bad' }) -Text $l.Name -Detail $l.Detail)) | Out-Null }
    $again = New-GuiButton -Text 'Check again' -Small -Icon refresh -OnClick { Show-MbtGuiPage }
    $again.Margin = [Windows.Thickness]::new(0, 10, 0, 0); $again.HorizontalAlignment = 'Left'
    $note = New-GuiText 'A red light says what to do. Check and repair, and Open, need no sign-in and no network.' 'Hint'
    $g.SetupRed = $red.Count
    return (New-GuiCard -Title 'This computer' -Right (New-GuiPill $(if ($red.Count) { 'Check' } else { 'Ready' }) $(if ($red.Count) { 'bad' } else { 'ok' })) -Content @($list, $note, $again))
}

function Open-MbtGuiSetup {
    <# From a page of the task: to Setup, and back to the same page. #>
    $s = $script:Gui.S
    $s.SetupReturn = $s.Step; $s.Step = 0
    Show-MbtGuiPage
}

function New-MbtGuiValueLine {
    <# A value read from the configuration file: its label, the value, and a pill when it was changed for this run. #>
    param([string]$Label, [string]$Value, [bool]$Changed)
    $line = [Windows.Controls.StackPanel]::new(); $line.Orientation = 'Horizontal'
    $t = New-GuiText $(if ($Value) { $Value } else { 'not set in the configuration file' }) 'Body'; $t.FontSize = 13
    if (-not $Value) { Set-MbtGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) 'MbtCautionInk' }
    $line.Children.Add($t) | Out-Null
    if ($Changed) { $p = New-GuiPill ('modified ' + [char]0x00B7 + ' this run only') 'warn' -Light; $p.Margin = [Windows.Thickness]::new(10, 0, 0, 0); $line.Children.Add($p) | Out-Null }
    return (New-GuiStack @((New-GuiText $Label 'Label'), $line))
}

function New-MbtGuiSigninSetupCard {
    <# Sign in as Me or the application; the tenant, the application and the certificate come from the configuration file and can be changed for this run only. #>
    $g = $script:Gui; $s = $g.S
    $appOn = Get-MbtGuiCapability 'AppOnly'
    $app = ($s.AuthMode -eq 'app') -and $appOn
    $switchTo = { param($mode) if ($s.AuthMode -ne $mode) { $s.AuthMode = $mode; $s.AuthEdit = $false; if ($s.Signed) { & $g.Engine.SignOut; $s.Signed = $false; $s.Locations = @(); $s.Code = $null; if ($g.Browsers.ContainsKey('export')) { Set-MbtGuiBrowserRoots -Browser $g.Browsers['export'] -Roots @() } }; $s.SignInError = ''; Show-MbtGuiPage } }.GetNewClosure()
    $content = [Collections.Generic.List[object]]::new()
    $content.Add((New-GuiText 'Sign in as' 'Label'))
    $content.Add((New-GuiOption -Type Radio -Group authmode -Title 'Me (delegated, recommended)' -Hint 'A window from Microsoft opens. You see the mailboxes you are allowed to open.' -Checked (-not $app) -Margin '0,4,0,0' -OnChange { & $switchTo 'user' }.GetNewClosure()))
    if ($appOn) { $content.Add((New-GuiOption -Type Radio -Group authmode -Title 'The application (certificate)' -Hint 'No window and no user: for scheduled or bulk jobs, and to write into a shared mailbox. Your administrator sets the application up first (see the app-only guide).' -Checked $app -Margin '0,4,0,0' -OnChange { & $switchTo 'app' }.GetNewClosure())) }
    $eff = Get-MbtGuiEffectiveSignIn -S $s
    $edit = $s.AuthEdit -or $eff.Missing.Count
    $box = [Windows.Controls.Border]::new(); $box.BorderThickness = [Windows.Thickness]::new(1); $box.CornerRadius = [Windows.CornerRadius]::new(8); $box.Padding = [Windows.Thickness]::new(12, 4, 12, 12); $box.Margin = [Windows.Thickness]::new(0, 14, 0, 0)
    Set-MbtGuiResource $box ([Windows.Controls.Border]::BorderBrushProperty) 'DividerStrokeColorDefaultBrush'; Set-MbtGuiResource $box ([Windows.Controls.Border]::BackgroundProperty) 'LayerFillColorDefaultBrush'
    $inner = [Collections.Generic.List[object]]::new()
    $head = New-GuiText $(if ($app) { 'Values for the application' } else { 'Values for the sign-in' }) 'Label'; $head.FontWeight = [Windows.FontWeights]::SemiBold
    $inner.Add($head)
    $fields = @(@('Tenant', 'Tenant', $(if ($app) { 'AppTenant' } else { 'TenantId' }), $eff.Tenant, $eff.TenantChanged), @('Application (client) id', 'Application id', $(if ($app) { 'AppClient' } else { 'ClientId' }), $eff.Client, $eff.ClientChanged))
    if ($app) { $fields += , @('Certificate thumbprint', 'Certificate thumbprint', 'AppThumb', $eff.Thumb, $eff.ThumbChanged) }
    foreach ($f in $fields) {
        $key = $f[2]
        if ($edit) {
            $inner.Add((New-GuiText $f[0] 'Label'))
            $inner.Add((New-GuiTextBox -Text $(if ($s[$key]) { $s[$key] } else { $f[3] }) -Automation $f[1] -OnChange { param($x, $e) $s[$key] = $x.Text }.GetNewClosure()))
        } else { $inner.Add((New-MbtGuiValueLine -Label $f[0] -Value $f[3] -Changed ([bool]$f[4]))) }
    }
    if ($edit) {
        $done = New-GuiButton -Text 'Done' -Small -OnClick { $s.AuthEdit = $false; Show-MbtGuiPage }
        $reset = New-GuiLink 'Use the values of the configuration file' { foreach ($k in 'TenantId', 'ClientId', 'AppTenant', 'AppClient', 'AppThumb') { $s[$k] = '' }; $s.AuthEdit = $false; Show-MbtGuiPage }; $reset.FontSize = 12; $reset.VerticalAlignment = 'Center'
        $bar = [Windows.Controls.StackPanel]::new(); $bar.Orientation = 'Horizontal'; $bar.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
        $bar.Children.Add($done) | Out-Null; $bar.Children.Add($reset) | Out-Null
        $inner.Add($bar)
    } else {
        $chg = New-GuiButton -Text 'Change for this run' -Small -Icon pencil -OnClick { $s.AuthEdit = $true; Show-MbtGuiPage }
        $chg.Margin = [Windows.Thickness]::new(0, 12, 0, 0); $chg.HorizontalAlignment = 'Left'
        $inner.Add($chg)
    }
    $inner.Add((New-GuiText $(if ($app) { 'Read from the configuration file. The certificate is looked up on this PC (your certificates, then the computer ones). A value you change applies to this run only; nothing is saved.' } else { 'Read from the configuration file. A value you change applies to this run only; nothing is saved.' }) 'Hint'))
    $box.Child = (New-GuiStack $inner.ToArray())
    $content.Add($box)
    return (New-GuiCard -Title 'Sign-in' -Right (New-GuiPill $(if ($app) { 'Application' } else { 'You' }) '') -Content $content.ToArray())
}

function New-MbtGuiPlacesCard {
    <# The configuration file (Reload, Open) and the work folder (Choose a folder...). #>
    $g = $script:Gui; $s = $g.S
    $hasPath = if ($g.Engine -is [Collections.IDictionary]) { $g.Engine.Contains('ConfigPath') } else { [bool]$g.Engine.PSObject.Properties['ConfigPath'] }
    $cfgPath = if ($s.ConfigPath) { $s.ConfigPath } elseif ($hasPath) { [string](& $g.Engine.ConfigPath) } else { '' }
    $cfgBox = New-GuiTextBox -Text $cfgPath -ReadOnly -Automation 'Configuration file'
    $reload = New-GuiButton -Text 'Reload' -Small -Icon refresh -OnClick { Show-MbtGuiPage }
    $openCfg = New-GuiButton -Text 'Open' -Small -Icon open -OnClick { if ($cfgPath -and (Test-Path -LiteralPath $cfgPath)) { Start-Process -FilePath $cfgPath } }.GetNewClosure()
    $bar = [Windows.Controls.StackPanel]::new(); $bar.Orientation = 'Horizontal'; $bar.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
    $bar.Children.Add($reload) | Out-Null; $bar.Children.Add($openCfg) | Out-Null
    $free = $null; try { $free = [double](& $g.Engine.FreeSpace $s.WorkFolder) } catch { $free = $null }
    $wfLabel = [Windows.Controls.StackPanel]::new(); $wfLabel.Orientation = 'Horizontal'
    $wfLabel.Children.Add((New-GuiText 'Work folder' 'Label')) | Out-Null
    if ($s.Changed.ContainsKey('WorkFolder')) { $p = New-GuiPill ('modified ' + [char]0x00B7 + ' this run only') 'warn' -Light; $p.Margin = [Windows.Thickness]::new(8, 6, 0, 0); $wfLabel.Children.Add($p) | Out-Null }
    $wf = New-GuiTextBox -Text $s.WorkFolder -ReadOnly -Automation 'Work folder'
    $pick = New-GuiButton -Text 'Choose a folder...' -OnClick { $v = Select-MbtGuiFolder -Title 'Work folder' -InitialPath $s.WorkFolder; if ($v -and $v -ne $s.WorkFolder) { $s.WorkFolder = $v; $s.Changed['WorkFolder'] = $true; Show-MbtGuiPage } }.GetNewClosure()
    $hint = "Jobs, export stores and logs live here. $(if ($null -ne $free) { "$(Format-MbtGuiBytes $free) free. " })A folder you choose here applies to this run only."
    return (New-GuiCard -Title 'Files and folders' -Content @((New-GuiText 'Configuration file' 'Label'), $cfgBox, $bar, (New-GuiText 'Tenant, application and defaults are read from this file. Reload reads it again; Open shows it in your editor.' 'Hint'), $wfLabel, (New-GuiRow -Box $wf -Button $pick), (New-GuiText $hint 'Hint')))
}

function Build-MbtGuiSetupPage {
    $g = $script:Gui; $s = $g.S
    $left = New-GuiStack @((New-MbtGuiReadinessCard), (New-MbtGuiPlacesCard))
    $awake = New-GuiCard -Title 'While a job runs' -Content @((New-GuiOption -Type Check -Title 'Keep this PC awake while it runs' -Hint 'Stops Windows from going to sleep by itself. Closing the lid or choosing Sleep still sleeps.' -Checked $s.Awake -Margin '0' -OnChange { param($x, $e) $s.Awake = [bool]$x.IsChecked }))
    $right = New-GuiStack @((New-MbtGuiSigninSetupCard), $awake)
    $red = [int]$g.SetupRed
    $footer = if ($red) { 'A light is red. Cloud tasks need every light green; Check and repair, and Open, do not.' } else { 'Everything needed is in place. Nothing here is saved: it applies to this run.' }
    $back = $s.SetupReturn -ge 2
    Set-MbtGuiActions -Back $false -Footer $footer -Label $(if ($back) { 'Back to the task' } else { 'Choose a task' }) -OnClick {
        $s = $script:Gui.S
        $to = if ($s.SetupReturn -ge 1) { $s.SetupReturn } else { 1 }
        $s.SetupReturn = -1
        Set-MbtGuiStep $to
    }
    return (New-GuiStack @((New-GuiText 'Is this computer ready?' 'H2'), (New-GuiTwoColumns -Left $left -Right $right -LeftWidth 560)))
}

function Build-MbtGuiStartPage {
    $g = $script:Gui; $s = $g.S; $engine = $g.Engine
    $sel = $s.Sel
    $grid = [Windows.Controls.Primitives.UniformGrid]::new(); $grid.Columns = 4
    foreach ($k in @($script:Tasks.Keys | Where-Object { -not (Get-MbtGuiProp $script:Tasks[$_] 'Hidden') })) {
        $card = New-MbtGuiTaskCard -Key $k -Selected ($sel -eq $k)
        $card.Margin = [Windows.Thickness]::new(0, 0, $(if ($k -eq 'open') { 0 } else { 12 }), 0)
        Add-MbtGuiHandler $card 'Click' { param($sender, $e) $script:Gui.S.Sel = [string]$sender.Tag; Show-MbtGuiPage }
        $grid.Children.Add($card) | Out-Null
    }
    $left = New-GuiCard -Title 'Continue a job' -Info ('Found in your work folder ' + $s.WorkFolder) -Content @((New-MbtGuiJobList))
    $right = if ($sel) {
        $t = $script:Tasks[$sel]
        $ol = [Windows.Controls.StackPanel]::new()
        $i = 1
        foreach ($w in $t.Will) { $row = New-GuiText ("$i.  $w") 'Body'; $row.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $ol.Children.Add($row) | Out-Null; $i++ }
        New-GuiCard -Title 'What will happen' -Content @($ol, (New-GuiText $t.Safe 'Hint'))
    } else { New-GuiCard -Title 'What will happen' -Content @((New-GuiText 'Pick a task above to see the steps.' 'Muted')) }
    $cols = New-GuiGrid -Children @($left, $right) -Widths @('1.4*', '1*')
    $stack = New-GuiStack @((New-GuiText 'What do you want to do?' 'H2'), $grid, (New-GuiStack @($cols) -Margin '0,12,0,0'))
    $footer = if ($sel) { "You chose **$($script:Tasks[$sel].Title)**." } else { 'Pick what you want to do. The console guides you from there.' }
    Set-MbtGuiActions -Footer $footer -Label 'Continue' -Enabled ([bool]$sel) -OnClick {
        $s = $script:Gui.S
        $s.Task = $s.Sel
        $s.Step = $script:Flows[$s.Task].IndexOf($(switch ($s.Task) { 'export' { 'exFrom' } 'import' { 'imFrom' } 'repair' { 'rpFrom' } default { 'opFrom' } }))
        Initialize-MbtGuiTask
        Show-MbtGuiPage
    }
    return $stack
}

function Initialize-MbtGuiTask {
    <# A task was chosen: what its pages start with. #>
    $g = $script:Gui; $s = $g.S
    $s.Outcome = $null; $s.PlanState = 'idle'; $s.PlanMsg = ''; $s.RStep = $null; $s.RObj = $null; $s.Confirmed = $false; $s.ConfirmKey = ''; $s.JobName = ''
    $g.Run = $null; $g.Job = $null
    switch ($s.Task) {
        'export' { if (-not $s.OutFolder) { $s.OutFolder = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'MailboxTransfer' } }
        'import' { $s.Confirmed = $false; $s.ConfirmKey = '' }
        'repair' { $s.RpCheck = $s.RpCheck }
    }
}

#endregion

