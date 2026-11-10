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
    param([Parameter(Mandatory)]$Engine, [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System', [string]$WorkFolder, [string]$EntryScript, [string]$Version = '1.0.1')

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
