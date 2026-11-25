#Requires -Version 7.6
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
    mcTo = { Build-MbtGuiMailCopyToPage }
    mcPlan = { Build-MbtGuiMailCopyPlanPage }
    imFrom = { Build-MbtGuiImportFromPage }
    imTo = { Build-MbtGuiImportToPage }
    imPlan = { Build-MbtGuiImportPlanPage }
    rpFrom = { Build-MbtGuiRepairFromPage }
    rpTo = { Build-MbtGuiRepairToPage }
    rpPlan = { Build-MbtGuiRepairPlanPage }
    opFrom = { Build-MbtGuiOpenPage }
    cpTo = { Build-MbtGuiCopyToPage }
    cpPlan = { Build-MbtGuiCopyPlanPage }
    resume = { Build-MbtGuiResumePage }
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
    if ($Selection.Items -gt 0) { return ('**{0}** ticked in {1}, {2}' -f (Format-MbtGuiSelectionItems $Selection -WithNoun), (Format-MbtGuiQuantity $Selection.Folders 'folder'), (Format-MbtGuiSelectionSize $Selection)) }
    if ($Signed) { return 'Tick what to keep, then continue.' }
    return 'Sign in first.'
}

function New-MbtGuiSelectionCard {
    param($Selection, $Period)
    if ($Selection.Items -gt 0) {
        $atMost = [bool]($Period -and $Period.Active)
        $body = New-GuiFacts @(@('Items', (Format-MbtGuiSelectionItems $Selection -AtMost:$atMost)), @('Size', ($(if ($atMost) { 'at most ' } else { '' }) + (Format-MbtGuiSelectionSize $Selection))), @('Folders', [string]$Selection.Folders))
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

function New-MbtGuiJobRow {
    <# One job of the list: the icon of its state, its title, why it stopped and when, the step it reached, and its buttons (Continue for a job that is not finished, Open the report, Delete). A batch has its report only. #>
    param([Parameter(Mandatory)]$Job, [bool]$First)
    $state = [string](Get-MbtGuiProp $Job 'State' 'run')
    $isBatch = ((Get-MbtGuiProp $Job 'Task' '') -eq 'batch')
    $glyphName = switch ($state) { 'ok' { 'check' } 'warn' { 'warn' } default { 'pause' } }
    $brush = switch ($state) { 'ok' { 'MbtSuccess' } 'warn' { 'MbtCaution' } default { 'MbtBrandText' } }
    $row = [Windows.Controls.Grid]::new()
    foreach ($w in 'Auto', '*') { $cd = [Windows.Controls.ColumnDefinition]::new(); $cd.Width = if ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } else { [Windows.GridLength]::Auto }; $row.ColumnDefinitions.Add($cd) }
    $ic = New-GuiIcon $glyphName 14 $brush; $ic.Margin = [Windows.Thickness]::new(0, 2, 12, 0); $ic.VerticalAlignment = 'Top'
    $txt = [Windows.Controls.StackPanel]::new(); [Windows.Controls.Grid]::SetColumn($txt, 1)
    $t1 = New-GuiText ([string](Get-MbtGuiProp $Job 'Title' '')) 'Body'; $t1.FontSize = 13; $t1.FontWeight = [Windows.FontWeights]::SemiBold
    $txt.Children.Add($t1) | Out-Null
    $why = [string](Get-MbtGuiProp $Job 'Sub' '')
    if ($why) { $txt.Children.Add((New-GuiText $why 'Small')) | Out-Null }
    $progress = [string](Get-MbtGuiProp $Job 'Sub2' '')
    if ($progress) { $txt.Children.Add((New-GuiText $progress 'Small')) | Out-Null }
    $bar = [Windows.Controls.WrapPanel]::new(); $bar.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
    if ($isBatch) {
        $bar.Children.Add((New-GuiButton -Text ([string](Get-MbtGuiProp $Job 'Button' 'Open the report')) -Small -OnClick { Resume-MbtGuiJob $Job }.GetNewClosure())) | Out-Null
    } else {
        if ([bool](Get-MbtGuiProp $Job 'CanContinue' $false)) { $bar.Children.Add((New-GuiButton -Text 'Continue' -Small -Icon play -OnClick { Resume-MbtGuiJob $Job }.GetNewClosure())) | Out-Null }
        $bar.Children.Add((New-GuiButton -Text 'Open the report' -Small -Icon doc -OnClick { Open-MbtGuiJobReport $Job }.GetNewClosure())) | Out-Null
        $bar.Children.Add((New-GuiButton -Text 'Delete' -Small -Icon trash -OnClick { Show-MbtGuiJobDelete $Job }.GetNewClosure())) | Out-Null
    }
    $txt.Children.Add($bar) | Out-Null
    $row.Children.Add($ic) | Out-Null; $row.Children.Add($txt) | Out-Null
    $wrap = [Windows.Controls.Border]::new(); $wrap.Child = $row; $wrap.Padding = [Windows.Thickness]::new(0, 9, 0, 9)
    if (-not $First) { $wrap.BorderThickness = [Windows.Thickness]::new(0, 1, 0, 0); Set-MbtGuiResource $wrap ([Windows.Controls.Border]::BorderBrushProperty) 'DividerStrokeColorDefaultBrush' }
    return $wrap
}

function New-MbtGuiJobList {
    <# The jobs of the work folder: the ones that are not finished first, every one of them, in a list that scrolls inside the card; the finished ones in a fold below. #>
    $g = $script:Gui
    $jobs = @()
    try { $jobs = @(& $g.Engine.Jobs $g.S.WorkFolder) } catch { Write-Verbose "jobs: $_" }
    if (-not $jobs.Count) { return (New-GuiText 'No job yet. A job you stop, or that is not finished, is listed here so you can continue it.' 'Hint') }
    $open = @($jobs | Where-Object { -not [bool](Get-MbtGuiProp $_ 'Finished' $false) })
    $done = @($jobs | Where-Object { [bool](Get-MbtGuiProp $_ 'Finished' $false) })
    $stack = [Windows.Controls.StackPanel]::new()
    $first = $true
    foreach ($j in $open) { $stack.Children.Add((New-MbtGuiJobRow -Job $j -First $first)) | Out-Null; $first = $false }
    if (-not $open.Count) { $stack.Children.Add((New-GuiText 'Nothing to continue: every job is finished.' 'Hint')) | Out-Null }
    if ($done.Count) {
        $rows = @(for ($i = 0; $i -lt $done.Count; $i++) { New-MbtGuiJobRow -Job $done[$i] -First ($i -eq 0) })
        $fold = New-GuiMore -Summary "Finished ($($done.Count))" -Content $rows
        $fold.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
        $stack.Children.Add($fold) | Out-Null
    }
    $scroll = [Windows.Controls.ScrollViewer]::new()
    $scroll.VerticalScrollBarVisibility = 'Auto'; $scroll.HorizontalScrollBarVisibility = 'Disabled'; $scroll.MaxHeight = 380; $scroll.Focusable = $false; $scroll.FontSize = 14
    $scroll.Content = $stack
    $scroll.Tag = 'MbtJobList'
    # the wheel scrolls the list, and the page when the list is at its end
    Add-MbtGuiHandler $scroll 'PreviewMouseWheel' {
        param($sender, $e)
        $atTop = ($e.Delta -gt 0 -and $sender.VerticalOffset -le 0)
        $atBottom = ($e.Delta -lt 0 -and $sender.VerticalOffset -ge $sender.ScrollableHeight)
        if (-not ($atTop -or $atBottom)) { return }
        $e.Handled = $true
        $parent = [Windows.Media.VisualTreeHelper]::GetParent($sender) -as [Windows.UIElement]
        if ($parent) {
            $forward = [Windows.Input.MouseWheelEventArgs]::new($e.MouseDevice, $e.Timestamp, $e.Delta)
            $forward.RoutedEvent = [Windows.UIElement]::MouseWheelEvent; $forward.Source = $sender
            $parent.RaiseEvent($forward)
        }
    }
    return $scroll
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
    try { $lights = @((& $g.Engine.Readiness $g.S.WorkFolder).Lights) } catch { $lights = @(@{ Name = 'The engine could not be checked'; Detail = $_.Exception.Message; Ok = $false }) }
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
    $content.Add((New-GuiOption -Type Radio -Group authmode -Title 'Me (delegated, recommended)' -Hint 'A Microsoft sign-in window opens over the console. You see the mailboxes you are allowed to open.' -Checked (-not $app) -Margin '0,4,0,0' -OnChange { & $switchTo 'user' }.GetNewClosure()))
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
    $right = New-GuiStack @((New-MbtGuiSigninSetupCard))
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
    $grid = [Windows.Controls.Primitives.UniformGrid]::new(); $grid.Columns = @($script:Tasks.Keys | Where-Object { -not (Get-MbtGuiProp $script:Tasks[$_] 'Hidden') }).Count
    foreach ($k in @($script:Tasks.Keys | Where-Object { -not (Get-MbtGuiProp $script:Tasks[$_] 'Hidden') })) {
        $card = New-MbtGuiTaskCard -Key $k -Selected ($sel -eq $k)
        $card.Margin = [Windows.Thickness]::new(0, 0, $(if ($k -eq 'open') { 0 } else { 12 }), 0)
        Add-MbtGuiHandler $card 'Click' { param($sender, $e) $script:Gui.S.Sel = [string]$sender.Tag; Show-MbtGuiPage }
        $grid.Children.Add($card) | Out-Null
    }
    $left = New-GuiCard -Title 'Continue a job' -Info ('Found in your work folder ' + $s.WorkFolder) -Right (New-GuiButton -Text 'All jobs' -Small -Icon doc -OnClick { Open-MbtGuiJobOverview }) -Content @((New-MbtGuiJobList))
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
        $s.Step = $script:Flows[$s.Task].IndexOf($(switch ($s.Task) { 'export' { 'exFrom' } 'import' { 'imFrom' } 'repair' { 'rpFrom' } 'mailcopy' { 'exFrom' } default { 'opFrom' } }))
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

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBEW31WcESq3dNa
# YUM63U94ru1Aj8wrpZXwZWvupnzonaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYETCCGA0CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDBk4Vg7h26
# 1z21c6n794hMucUIkvJiQEymSp0ZaULDcjANBgkqhkiG9w0BAQEFAASCAgBb4fGH
# eda73UClB0FLPqoD4t4ZuSKpR5fuoT0pV3XuZHd3ABrehUyrhzwbXI6EtijmRXhR
# 0/lcUeL3tEBH+dcDSExtkApAhXISKv7+d2LGGg9rWhwwQsFZ/ANeflyZ7tDclncZ
# 6uJ3k8L7X+SFtDerDfp/yKPa/VuwyC9h5lgXr0EgDaQvU+wdF7Pve72jxm0tIQ4c
# m6bj5ww7iBDYcJYth3jdOWA5xgaAoOgfJ4pKAuZRyyfQFtjRGvW9DeOGq2Z4Blca
# qule2cuXzOARxEszd1pcmUu3cZl0PhNWpLUSqktMB2GDR2riGNnr0dUF7AVg7Hs8
# ex8Byo1yNHLMxPhI+HdCAPgBlkRFCHAZZJeNjjCPQSm1mNy29gevn71RQwmBT7ml
# ERF1F1PAoJ1QdGLNupM3zInqhLvaAXJ1Bw4V0TX7jvnzsabHMHeBl7W8NEBOTBO3
# dFgJJpbzn8eBWT/brpgsHgkpTd4m+/OyVM+eDW28TYdfjFagjdYkkyF+5yB/ZRmm
# imffIZ3vPh7TQhwVSJqlZxM9p2Bkn8ZGNBXVq9HGFI8Mr8DXwK6jilx0JHBVGw2j
# v3IACLgOGErrTA3mP3DFgwbY6G4Fuo1hiCDlHz5/1fbZRBFERMgdlxr8IhcdbN2V
# MCN6VZBX6ax6uJL/KO7RV9+6IcRTMIDdXhVde6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIILa6HjJY1xXGJBCawDb+pPhEMFEZi4w+NUTR9iGy5jBAggP4DgZnKMf
# phgPMjAyNjExMTUxNDM1MzVaMAMCAQECEEE+EjWT1Ysdw/vWVQbv9xCgghGnMIIE
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
# wzSPOzGCAlYwggJSAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTE0MzUzNVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIHqG
# FTi5qNn8RZb1mDnH67KSk2wFBJYac9c6ZX/7+tX1MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiAdAWPa9RtxrfwsFHFqmq/6UdiOOQ8uGvGki+Xi/2FIqwIhAKm7
# TEeDGCUxu/DM/3ExBGOkLjTpW4DjvT2LahTsgraE
# SIG # End signature block
