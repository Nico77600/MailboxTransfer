#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - window, flow and state (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    The window of the console: header, the stepper (Task > From > To > Plan > Run > Done), the page, the action bar, the dialog layer.
    The state machine is the one of the mockup: a task chooses a flow; only the steps already passed are clickable; while a job runs nothing is.
    Pages are functions (MailboxTransfer.Gui.Pages.ps1) that return the element of the page and set the action bar; the page is built again
    when the state changes, except the browse cards, which live as long as the window.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

$script:Flows = @{
    export = @('start', 'exFrom', 'exTo', 'exPlan', 'run', 'done')
    import = @('start', 'imFrom', 'imTo', 'imPlan', 'run', 'done')
    repair = @('start', 'rpFrom', 'rpTo', 'rpPlan', 'run', 'done')
    open   = @('start', 'opFrom')
    copy   = @('start', 'opFrom', 'cpTo', 'cpPlan', 'run', 'done')
}
$script:StepLabels = @{ start = 'Task'; exFrom = 'From'; imFrom = 'From'; rpFrom = 'From'; opFrom = 'From'; exTo = 'To'; imTo = 'To'; rpTo = 'To'; cpTo = 'To'; exPlan = 'Plan'; imPlan = 'Plan'; rpPlan = 'Plan'; cpPlan = 'Plan'; run = 'Run'; done = 'Done' }
$script:Tasks = [ordered]@{
    export = @{ Title = 'Browse and export'; Eyebrow = 'BROWSE ' + [char]0x00B7 + ' EXPORT'; Icon = 'down'; Context = 'Browse and export'
        Text = 'Look inside a mailbox, tick what to keep, and save it on this computer as PST files or as an export store. Nothing is changed in the mailbox.'
        Will = @('Sign in and pick the mailbox.', 'Browse it like a mailbox and tick what to keep.', 'Choose where to save it.', 'Read the plan, then press Start.')
        Safe = 'Nothing is changed in the mailbox.' }
    import = @{ Title = 'Import'; Eyebrow = 'IMPORT'; Icon = 'up'; Context = 'Import'
        Text = 'Put a PST file or an export into a mailbox. Items are added. Nothing is deleted and nothing is overwritten.'
        Will = @('Pick the PST file or the export.', 'Pick the mailbox (and where in it).', 'Read the plan, type IMPORT, press Start.', 'Wait. A big import takes hours; you can stop and continue.')
        Safe = 'This task adds items to a mailbox. It never deletes or overwrites.' }
    repair = @{ Title = 'Check and repair a PST'; Eyebrow = 'PST'; Icon = 'wrench'; Context = 'Check and repair a PST'
        Text = 'Test a PST file for damage. If it is damaged, write a repaired copy. Your original file is never changed.'
        Will = @('Pick the PST file. The check starts at once and only reads.', 'Read the result in plain words.', 'If needed, choose to write a repaired copy.', 'Compare before and after.')
        Safe = 'Your original file is never changed. A repaired copy is a new file.' }
    open   = @{ Title = 'Open an export or a PST'; Eyebrow = 'OPEN ' + [char]0x00B7 + ' READ ONLY'; Icon = 'open'; Context = 'Open an export or a PST'
        Text = 'Read a PST file or an export store like a mailbox: folders, messages, attachments. Read only.'
        Will = @('Pick the PST file or the export folder.', 'Browse it. Tick items if you want to move them later.')
        Safe = 'Nothing is written anywhere.' }
    copy   = @{ Title = 'Copy part of a PST'; Eyebrow = 'COPY ' + [char]0x00B7 + ' PST'; Icon = 'down'; Context = 'Copy part of a PST'; Hidden = $true
        Text = 'Write the folders and messages you ticked in an open PST into new PST files.'
        Will = @('Tick folders and messages in the open PST.', 'Choose where to save the new files.', 'Read the plan, then press Start.')
        Safe = 'Your original PST is only read.' }
}

function New-MbtGuiState {
    param([string]$WorkFolder)
    @{
        Task = $null; Sel = $null; Step = 0
        Mailbox = ''; AuthMode = 'user'; AppTenant = ''; AppClient = ''; AppThumb = ''; TenantId = ''; ClientId = ''; Shared = ''; ImShared = ''; Signed = $false; SigningIn = $false; CodeShown = $false; Code = $null; SignInError = ''; Account = ''; Token = $null
        Locations = @(); LocationChecked = @{}; Reading = $false; LocationsFile = ''; RecoverableItems = $false
        Dest = 'pst'; KeepStore = $true; SplitGB = 50; CheckEach = $true; SepAux = $false; OutFolder = ''; OutName = ''
        Conc = 4; Retry = 2; Tries = 6; Reconcile = 5; Awake = $true; WorkFolder = $WorkFolder; Changed = @{}
        ImSrc = 'store'; ImPath = ''; ImInfo = $null; ImChecking = $false; TgtLoc = 'Primary'; TgtMode = 'new'; TgtName = ''; Mapping = ''; Confirm = ''
        RpPath = ''; RpCheck = $null; RpChecking = $false; RpLevel = 'Automatic'; RpDest = ''; RpKeep = $false
        OpSrc = 'pst'; OpPath = ''; OpInfo = $null
        PlanState = 'idle'; PlanMsg = ''; PlanKey = ''; PlanTotals = $null; CpFor = ''
        RStep = $null; RObj = $null; Outcome = $null; LogOpen = $false; SignInTech = ''
    }
}

#region Window --------------------------------------------------------------------------------

function Open-MbtGuiHelp {
    <# Opens the user guide (HTML, in the package) or the page of the latest release in the default browser. A guide that is not there is said in a dialog. #>
    param([Parameter(Mandatory)][ValidateSet('Guide', 'Release')][string]$Target)
    $url = 'https://github.com/Nico77600/MailboxTransfer/releases/latest'
    if ($Target -eq 'Guide') {
        $guide = Join-Path $script:GuiRoot 'docs\MailboxTransfer-UserGuide.html'
        if (-not (Test-Path -LiteralPath $guide -PathType Leaf)) { $guide = Join-Path $script:GuiRoot 'docs\MailboxTransfer-UserGuide.md' }
        if (-not (Test-Path -LiteralPath $guide -PathType Leaf)) { $url = 'https://github.com/Nico77600/MailboxTransfer/blob/main/package/docs/MailboxTransfer-UserGuide.md' } else { $url = $guide }
    }
    try { Start-Process -FilePath $url } catch { Write-Verbose "Cannot open $url : $($_.Exception.Message)" }
}

function Set-MbtGuiThemeChoice {
    <# A button of the theme switch was checked: apply System, Light or Dark to the application and the window, then draw the page again. #>
    param([Parameter(Mandatory)][ValidateSet('System', 'Light', 'Dark')][string]$Name)
    $g = $script:Gui
    if (-not $g) { return }
    if ($g.ThemeName -eq $Name) { return }
    $g.ThemeName = $Name
    $look = Initialize-MbtGuiTheme -Theme $Name
    $g.Theme = $look
    Set-MbtGuiTheme -Window $g.Window -Theme $look
    Show-MbtGuiPage
}

function New-MbtGuiWindow {
    <#
        Builds the window (without showing it). Used by Show-MbtGuiConsole, the headless tests and the screenshots.
        -Engine: the adapter (real or fake). -Theme: System (like Windows), Light, Dark.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Engine, [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System', [string]$WorkFolder, [string]$EntryScript, [string]$Version = '1.0.0')

    $look = Initialize-MbtGuiTheme -Theme $Theme
    $xaml = Get-Content -LiteralPath (Join-Path $script:GuiRoot 'module\MailboxTransfer.Gui.xaml') -Raw
    $window = [Windows.Markup.XamlReader]::Parse($xaml)
    Set-MbtGuiTheme -Window $window -Theme $look
    $names = 'Root', 'Header', 'Version', 'Th_System', 'Th_Light', 'Th_Dark', 'HelpGuide', 'HelpRelease', 'StepperCard', 'Stepper', 'StepperContext', 'PageScroll', 'Page', 'Actions', 'BackButton', 'ActionsLeft', 'Footer', 'ActionsExtra', 'Primary', 'PrimaryIcon', 'PrimaryLabel',
        'Overlay', 'DialogCard', 'DialogTitle', 'DialogBody', 'DialogButtons'
    $c = @{}
    foreach ($n in $names) { $c[$n] = $window.FindName($n) }
    $c.Version.Text = "v$Version  " + [char]0x00B7 + '  Nicolas Fabert'
    foreach ($name in 'System', 'Light', 'Dark') { $c["Th_$name"].IsChecked = ($name -eq $Theme) }
    $c.Primary.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'AccentButtonStyle')
    $window.TaskbarItemInfo = [Windows.Shell.TaskbarItemInfo]::new()

    if (-not $WorkFolder) { $WorkFolder = Join-Path $env:LOCALAPPDATA 'MailboxTransfer' }
    $timer = [Windows.Threading.DispatcherTimer]::new([Windows.Threading.DispatcherPriority]::Background)
    $timer.Interval = [TimeSpan]::FromMilliseconds(100)
    $script:Gui = @{
        Window = $window; C = $c; Theme = $look; ThemeName = $Theme; Engine = $Engine; S = (New-MbtGuiState -WorkFolder $WorkFolder); Version = $Version
        EntryScript = $EntryScript; Timer = $timer; Job = $null; Run = $null; Browsers = @{}; PrimaryAction = $null; BackAction = $null
        OnSelectionChanged = $null; Reports = @(); ExportSelection = [pscustomobject]@{ Entries = @(); Items = 0L; Bytes = 0L; Folders = 0; UnknownSize = 0; SubtreeUnread = $false }; ImportTotals = $null; PlanJob = $null; PlanRun = $null; PlanHolder = $null
        ExportToHolder = $null; ExportDiskHolder = $null; ExportDiskBad = $false; RepairHolder = $null; CloseWhenStopped = $false; LastTickError = $null; PageName = ''
        RunHeadHolder = $null; RunDecisionHolder = $null; RunCols = $null; RunStepsHolder = $null; RunMidHolder = $null; RunDetailHolder = $null
        DialogOpen = $false; Pollers = [Collections.Generic.List[scriptblock]]::new(); Tick = 0; LastEtaPaint = [datetime]::MinValue; Closing = $false; ForceClose = $false
    }
    # The theme can be changed in the window (System follows Windows): the window takes the new look and the page is built again.
    foreach ($name in 'System', 'Light', 'Dark') {
        $c["Th_$name"].Tag = $name
        $c["Th_$name"].Add_Checked({ param($s, $e) Set-MbtGuiThemeChoice -Name ([string]$s.Tag) })
    }
    # Help: the HTML user guide of the package, and the page of the latest release (the tool never checks for updates by itself; this opens the browser on a click).
    $c.HelpGuide.Add_Click({ Open-MbtGuiHelp -Target Guide })
    $c.HelpRelease.Add_Click({ Open-MbtGuiHelp -Target Release })
    $c.BackButton.Add_Click({ Invoke-MbtGuiBack })
    $c.Primary.Add_Click({ if ($script:Gui.PrimaryAction) { & $script:Gui.PrimaryAction } })
    $timer.Add_Tick({ Step-MbtGuiTick })
    $window.Add_PreviewKeyDown({
            param($s, $e)
            if ($e.Key -eq 'Escape' -and $script:Gui.DialogOpen) { Close-MbtGuiDialog; $e.Handled = $true }
            elseif ($e.Key -eq 'Return' -and -not $script:Gui.DialogOpen -and $e.OriginalSource -isnot [Windows.Controls.TextBox] -and $c.Primary.IsEnabled -and $e.OriginalSource -isnot [Windows.Controls.Button]) { & $script:Gui.PrimaryAction; $e.Handled = $true }
        })
    $window.Add_Closing({ param($s, $e) Invoke-MbtGuiClosing $e })
    # The window fits the screen where it opens (small laptop screen at 150 %).
    $area = [Windows.SystemParameters]::WorkArea
    $window.Width = [Math]::Min($window.Width, $area.Width); $window.Height = [Math]::Min($window.Height, $area.Height)
    $window.MinWidth = [Math]::Min($window.MinWidth, $area.Width); $window.MinHeight = [Math]::Min($window.MinHeight, $area.Height)
    Show-MbtGuiPage
    return $window
}

#endregion

#region Pages, stepper, actions ------------------------------------------------------------------

function Get-MbtGuiScreen {
    $s = $script:Gui.S
    if (-not $s.Task) { return 'start' }
    return $script:Flows[$s.Task][$s.Step]
}

function Show-MbtGuiPage {
    <# Builds the page of the current step again. #>
    $g = $script:Gui
    $screen = Get-MbtGuiScreen
    $fn = $script:PageBuilders[$screen]
    $g.PageName = $screen
    $page = & $fn
    $g.C.Page.Content = $page
    Update-MbtGuiStepper
    $g.C.PageScroll.ScrollToTop()
    if ($g.Timer -and -not $g.Timer.IsEnabled) { $g.Timer.Start() }
}

function Update-MbtGuiStepper {
    $g = $script:Gui; $s = $g.S; $panel = $g.C.Stepper
    $panel.Children.Clear()
    $flow = if ($s.Task) { $script:Flows[$s.Task] } else { $script:Flows.export }
    $screen = Get-MbtGuiScreen
    $locked = ($screen -eq 'run' -and $g.Run -and -not $g.Run.Result)
    for ($i = 0; $i -lt $flow.Count; $i++) {
        if ($i -gt 0) { $panel.Children.Add((New-GuiIcon 'chev' 10 'TextFillColorTertiaryBrush')) | Out-Null; $panel.Children[$panel.Children.Count - 1].Margin = [Windows.Thickness]::new(4, 0, 4, 0) }
        $cur = $i -eq $s.Step; $done = $i -lt $s.Step
        $btn = [Windows.Controls.Button]::new()
        $btn.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtStepButton')
        $btn.IsEnabled = ($i -lt $s.Step) -and -not $locked
        $btn.Tag = $i
        $sp = [Windows.Controls.StackPanel]::new(); $sp.Orientation = 'Horizontal'
        $circle = [Windows.Controls.Border]::new(); $circle.Width = 24; $circle.Height = 24; $circle.CornerRadius = [Windows.CornerRadius]::new(12)
        $bgKey = if ($cur) { 'AccentFillColorDefaultBrush' } elseif ($done) { 'MbtSuccessBackground' } else { 'MbtInfoBackground' }
        Set-MbtGuiResource $circle ([Windows.Controls.Border]::BackgroundProperty) $bgKey
        if ($done) { $num = New-GuiIcon 'check' 12 'MbtSuccessInk'; $num.HorizontalAlignment = 'Center' }
        else {
            $num = New-GuiText ([string]($i + 1)) 'Small' -NoWrap; $num.FontWeight = [Windows.FontWeights]::SemiBold; $num.HorizontalAlignment = 'Center'; $num.VerticalAlignment = 'Center'
            Set-MbtGuiResource $num ([Windows.Controls.TextBlock]::ForegroundProperty) $(if ($cur) { 'TextOnAccentFillColorPrimaryBrush' } else { 'TextFillColorSecondaryBrush' })
        }
        $circle.Child = $num
        $lbl = New-GuiText $script:StepLabels[$flow[$i]] 'Body' -NoWrap
        $lbl.Margin = [Windows.Thickness]::new(8, 0, 0, 0); $lbl.VerticalAlignment = 'Center'
        if ($cur) { $lbl.FontWeight = [Windows.FontWeights]::SemiBold } else { Set-MbtGuiResource $lbl ([Windows.Controls.TextBlock]::ForegroundProperty) 'TextFillColorSecondaryBrush' }
        $sp.Children.Add($circle) | Out-Null; $sp.Children.Add($lbl) | Out-Null
        $btn.Content = $sp
        if ($cur) { Set-MbtGuiResource $btn ([Windows.Controls.Control]::BackgroundProperty) 'MbtAccentSoft' }
        $btn.Add_Click({ param($b, $e) Set-MbtGuiStep ([int]$b.Tag) })
        $panel.Children.Add($btn) | Out-Null
    }
    $ctx = if ($s.Task) { $script:Tasks[$s.Task].Context } else { 'Start' }
    if ($s.Task -in 'export', 'import' -and $s.Signed) { $ctx += ' ' + [char]0x00B7 + ' ' + $s.Mailbox }
    $g.C.StepperContext.Text = $ctx
}

function Set-MbtGuiActions {
    <# The action bar: Back at the left, optional left and extra elements, the footer, the primary button (grey until the page can continue). #>
    param([bool]$Back = $true, [string]$Footer = '', $Left, $Extra, [string]$Label, [string]$Icon, [bool]$Enabled = $true, [scriptblock]$OnClick)
    $c = $script:Gui.C
    $c.BackButton.Visibility = $(if ($Back) { 'Visible' } else { 'Collapsed' })
    $c.ActionsLeft.Content = $Left
    $c.ActionsExtra.Content = $Extra
    ConvertTo-MbtGuiRuns $c.Footer $Footer
    $c.Footer.Tag = $Footer
    if ($Label) {
        $c.Primary.Visibility = 'Visible'
        $c.PrimaryLabel.Text = $Label
        if ($Icon) { $c.PrimaryIcon.Text = Get-MbtGlyph $Icon; $c.PrimaryIcon.Visibility = 'Visible' } else { $c.PrimaryIcon.Visibility = 'Collapsed' }
        $c.Primary.IsEnabled = $Enabled
        $script:Gui.PrimaryAction = $OnClick
    } else { $c.Primary.Visibility = 'Collapsed'; $script:Gui.PrimaryAction = $null }
}

function Update-MbtGuiPrimary {
    <# Enables or disables the primary button and rewrites the footer without building the page again. #>
    param([bool]$Enabled, [string]$Footer)
    $c = $script:Gui.C
    $c.Primary.IsEnabled = $Enabled
    if ($PSBoundParameters.ContainsKey('Footer')) { ConvertTo-MbtGuiRuns $c.Footer $Footer }
}

function Set-MbtGuiStep {
    param([int]$Index)
    $s = $script:Gui.S
    if ($Index -lt 0) { $Index = 0 }
    $s.Step = $Index
    if ($Index -eq 0) { $s.Task = $null; $s.Sel = $null }
    Show-MbtGuiPage
}

function Move-MbtGuiStep {
    param([int]$Delta)
    $s = $script:Gui.S
    $max = $script:Flows[$s.Task].Count - 1
    Set-MbtGuiStep ([math]::Max(0, [math]::Min($max, $s.Step + $Delta)))
}

function Invoke-MbtGuiBack {
    $g = $script:Gui
    if ($g.BackAction) { & $g.BackAction } else { Move-MbtGuiStep -1 }
}

#endregion

#region Dialogs ---------------------------------------------------------------------------------

function Show-MbtGuiDialog {
    <# The console's own dialog. Buttons: @{ Text; Accent; Danger; Action } (Action runs after the dialog closes). #>
    param([Parameter(Mandatory)][string]$Title, $Body, [object[]]$Buttons)
    $g = $script:Gui; $c = $g.C
    $c.DialogTitle.Text = $Title
    $c.DialogBody.Content = $(if ($Body -is [string]) { New-GuiText $Body 'Body' } else { $Body })
    $c.DialogButtons.Children.Clear()
    foreach ($b in $Buttons) {
        $btn = New-GuiButton -Text $b.Text -Accent:([bool](Get-MbtGuiProp $b 'Accent')) -Danger:([bool](Get-MbtGuiProp $b 'Danger'))
        $btn.Margin = [Windows.Thickness]::new(8, 0, 0, 0)
        $btn.Tag = (Get-MbtGuiProp $b 'Action')
        $btn.Add_Click({ param($x, $e) $act = $x.Tag; Close-MbtGuiDialog; if ($act) { & $act } })
        $c.DialogButtons.Children.Add($btn) | Out-Null
    }
    $c.Overlay.Visibility = 'Visible'
    $g.DialogOpen = $true
    if ($c.DialogButtons.Children.Count) { [void]$c.DialogButtons.Children[0].Focus() }
}

function Close-MbtGuiDialog {
    $g = $script:Gui
    $g.C.Overlay.Visibility = 'Collapsed'
    $g.DialogOpen = $false
}

function Invoke-MbtGuiClosing {
    <# Closing the window during a run: Stop and keep progress (default) or Keep running. #>
    param($EventArgs)
    $g = $script:Gui
    if ($g.ForceClose) { Stop-MbtGuiShutdown; return }
    $running = $g.Run -and -not $g.Run.Result -and $g.Job -and -not $g.Job.Fake -and $g.Job.Process -and -not $g.Job.Exited
    if (-not $running) { Stop-MbtGuiShutdown; return }
    $EventArgs.Cancel = $true
    Show-MbtGuiDialog -Title 'A job is running' -Body 'Stop it before closing? Everything done so far is kept; press Continue the next time you open the console and only what is left is done. The job does not go on in the background.' -Buttons @(
        @{ Text = 'Keep running'; Action = $null }
        @{ Text = 'Stop and keep progress'; Accent = $true; Action = { Request-MbtGuiStopAndClose } })
}

function Request-MbtGuiStopAndClose {
    $g = $script:Gui
    Request-MbtGuiStop -Job $g.Job
    $g.Run.StopRequested = $true
    $g.CloseWhenStopped = $true
    Show-MbtGuiPage
}

function Stop-MbtGuiShutdown {
    $g = $script:Gui
    if ($g.Timer) { $g.Timer.Stop() }
    Stop-MbtGuiWorker $g.Engine.Worker
    Release-MbtGuiAwake
}

function Release-MbtGuiAwake { }

#endregion

#region Tick ---------------------------------------------------------------------------------------

function Step-MbtGuiTick {
    <# Every 100 ms: the worker, the events of the job, the end of the child, the painting of the run page. #>
    $g = $script:Gui
    if (-not $g) { return }
    $g.Tick++
    try {
        Step-MbtGuiWork $g.Engine.Worker
        if ($g.Job -and $g.Run) {
            $n = Step-MbtGuiJob -Job $g.Job -Run $g.Run
            if ($n -gt 0 -or $g.Run.Dirty) { Sync-MbtGuiRunPage }
            if ($g.Run.Result -and (Get-MbtGuiScreen) -eq 'run') { Complete-MbtGuiRun }
            elseif ((Get-MbtGuiScreen) -eq 'run' -and $g.Tick % 10 -eq 0) { Update-MbtGuiRunClock }
        }
        if ($g.CloseWhenStopped -and $g.Job -and $g.Job.Exited) { $g.ForceClose = $true; $g.Window.Close() }
        Step-MbtGuiSignInCode
        Step-MbtGuiPlanJob
        for ($i = $g.Pollers.Count - 1; $i -ge 0; $i--) { if (& $g.Pollers[$i]) { $g.Pollers.RemoveAt($i) } }
    } catch { Write-Verbose "tick: $_" ; $g.LastTickError = $_.Exception.Message }
}

#endregion

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDCmQZq+6uft5la
# IPhs4r3WnSdukyjsv1Py/KCtMYLITaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCC0DlJ3/OhM
# EbT5r7QKFv+h8ouqzyFs7SFXrqK8AYgU9zANBgkqhkiG9w0BAQEFAASCAgAhpEpV
# +CRssUFWKybGlt5Ge9wk7uX570bMOQcTVJp5acWbvfKYe86OOaAJbNPCQ1VN5fKA
# CVbxgkqGnoAUzaX507LInQdJDxWSLMQF69XHMJzUllGbZZfwNhLXFC2OSu7/Hqf7
# nBTFuQojNoyCyUdaQRRqU+upZ7A9ps92wtgcUUKiMTg//2xxFjz46KcM9bjOtpGC
# AJlk0iASXM6VN2R3CpQi5tVKVvlok0c6yOSQlANbGuoxjvXrJG52S+8QIZCohqV8
# /ITS9ZZAHejd0I0xwrzcYQCOo4CFyEZL2tXCdcQPxQ4k6+Tn8lizZy7T1a1G22eS
# YkLhwkLOL+d9fu1fSzXqp8eOA1WmYqwtACo1pkAbcI3oSxP3OJnFKoYnWQiVDTgi
# zw4U1iWO2pmqtPGM3vc8mjV41qR58GnvAj008Re/GN3srDCMZKt/R9qtpsOVkG8+
# P3e8mUZ4WNhziM5c9Iyz2wYgfGRw2hxACWxDEGDQWT2UJlM/MNSsrERsYCiqYUNQ
# Z8LjT8jnl+qSwREDIGbvbkGjm4JRnLklBNdCI6n6GTNIitS42+FN7/ARNbUtkBu7
# /hwQKxXz2F7dKrx5fsyohLL2B9kOfIGH5cgKzh++/gu1nYKhDuDNHSJ21LpDwmqE
# UaIzzmV8imprstsqxrLeRQTDlOsAswjE8K8ynKGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIKWR5qvVIXzVB0BfDJYg2w5dpoNiHZfbYYLFkmybYJ6vAghvgc1h0fS4
# CRgPMjAyNjExMTAwODIyNDhaMAMCAQECEGUVr9VYhrAtgkj+jhPkgPKgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjI0OFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIP2J
# VQIlUjjEwbf32Qzc0ULpyXEHJAt1mMZHynB1eNCDMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiBr9fH7V2L/6Q/uuwjj5v77KVYojdbuMc0Qex+yFx1zbgIgaWMt
# h3ZLo6PY1Jj+NRVO+2k3NISiDwtpUnViPWXbjwE=
# SIG # End signature block
