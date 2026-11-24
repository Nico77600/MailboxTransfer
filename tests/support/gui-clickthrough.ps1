#Requires -Version 7.4
<#
    Click-through of the console. The real window is built on the fake engine, shown without being visible, and run by ShowDialog like the real console; a timer
    of the dispatcher then walks every page of every scenario and, one control at a time, clicks every button, ticks and unticks every check box and radio button,
    chooses every entry of every list, types in every text box, opens and closes every fold, selects folders and messages, and answers every dialog that opens.
    A control whose handler raises an exception is reported with the scenario and the control. This is what the first click of a page does in a real window; a
    handler that uses a variable of the function that built the page (which is gone when the control raises its event) fails here as it does there.
    -Part and -Parts share the scenarios between several runs of this script (one process each); -Only names the scenarios to walk (and extras, the states they do not reach).
    Prints one JSON object: { Errors: [..], Invoked: { <type>: n }, Scenarios: n, Controls: n, Dialogs: n, Walked: [names of the scenarios walked], Edge: [names of the scenarios with exactly one of something, or none] }. Nothing contacts the cloud: the engine is the fake one, the
    jobs are not started, no file is opened. Run by tests\Gui.ClickThrough.Tests.ps1 in its own single-threaded pwsh (WPF needs STA).
    The names in this file start with ct and the script keeps its variables in the script scope: the block of a handler sees the scope it is raised from, as in the console.
#>
param([string]$Package, [string]$Work, [string[]]$Only, [int]$Part = 0, [int]$Parts = 1)
Set-StrictMode -Version Latest
[AppDomain]::CurrentDomain.SetData('MailboxTransfer.UnverifiedTreeWarned', $true)
$ErrorActionPreference = 'Stop'
$Only = @($Only | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
foreach ($f in 'Core', 'Run', 'Jobs', 'Browse', 'State', 'Pages', 'Pages.Export', 'Pages.Run', 'Pages.Source', 'Fake', 'Engine', 'Scenarios') { . (Join-Path $Package "module\MailboxTransfer.Gui.$f.ps1") }
$script:GuiRoot = $Package
$script:ctErrors = [Collections.Generic.List[string]]::new()
$script:ctInvoked = @{}
$script:ctScenarios = 0; $script:ctControls = 0; $script:ctDialogs = 0
$script:ctWalked = [Collections.Generic.List[string]]::new()
$script:ctChromeSeen = @{}
$script:ctStepControls = @{}

# nothing is opened, picked or started outside the window
function Start-Process { param($FilePath, $ArgumentList) $script:ctOpened = $FilePath }
function Select-MbtGuiFile { param($Title, $Filter, $InitialPath, [switch]$Save) return 'D:\Picked\file.pst' }
function Select-MbtGuiFolder { param($Title, $InitialPath) return 'D:\Picked' }
function Start-MbtGuiProcess { param($Job) $Job.Queue.Clear(); return $true }
function Start-MbtGuiDetached { param($Argument) }

function Wait-CtDispatcher {
    $ctFrame = [Windows.Threading.DispatcherFrame]::new()
    [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::ContextIdle, [Action] { $script:ctFrame.Continue = $false })
    $script:ctFrame = $ctFrame
    [Windows.Threading.Dispatcher]::PushFrame($ctFrame)
}
function Measure-CtWindow { $ctRoot = $script:ctWindow.Content; $ctRoot.UpdateLayout(); Wait-CtDispatcher; $ctRoot.UpdateLayout() }

function Get-CtText($ctNode) {
    if ($ctNode -is [Windows.Controls.TextBox]) { return "text box '$([Windows.Automation.AutomationProperties]::GetName($ctNode))'" }
    $ctName = [Windows.Automation.AutomationProperties]::GetName($ctNode)
    if ($ctName) { return $ctName }
    $ctParts = [Collections.Generic.List[string]]::new()
    $ctFind = $null
    $ctFind = { param($ctX) if ($ctX -is [Windows.Controls.TextBlock]) { if ($ctX.Text) { $ctParts.Add($ctX.Text) } } elseif ($ctX -is [Windows.DependencyObject]) { foreach ($ctC in [Windows.LogicalTreeHelper]::GetChildren($ctX)) { & $ctFind $ctC } } }
    if ($ctNode -is [Windows.Controls.ContentControl]) { if ($ctNode.Content -is [string]) { $ctParts.Add($ctNode.Content) } elseif ($ctNode.Content -is [Windows.DependencyObject]) { & $ctFind $ctNode.Content } }
    if ($ctNode -is [Windows.Controls.TreeViewItem] -or $ctNode -is [Windows.Controls.ListBoxItem]) { foreach ($ctP in 'Name', 'Subject') { if ($ctNode.DataContext.PSObject.Properties[$ctP]) { $ctParts.Add([string]$ctNode.DataContext.$ctP) } } }
    return ($ctParts -join ' / ')
}

function Get-CtControl {
    <# What a person can use now: visible, enabled buttons, check boxes, radio buttons, lists, text boxes, folds, folders of the tree and messages of the list, in the order of the page. #>
    $ctList = [Collections.Generic.List[object]]::new()
    $ctWalk = $null
    $ctWalk = {
        param($ctNode)
        if ($ctNode -isnot [Windows.DependencyObject]) { return }
        if ($ctNode -is [Windows.UIElement] -and -not $ctNode.IsVisible) { return }
        $ctUse = ($ctNode -is [Windows.Controls.Primitives.ButtonBase]) -or ($ctNode -is [Windows.Controls.ComboBox]) -or ($ctNode -is [Windows.Controls.TextBox] -and -not $ctNode.IsReadOnly) -or ($ctNode -is [Windows.Controls.Expander]) -or ($ctNode -is [Windows.Controls.TreeViewItem]) -or ($ctNode -is [Windows.Controls.ListBoxItem])
        if ($ctUse -and $ctNode.IsEnabled) { $ctList.Add($ctNode) }
        $ctN = [Windows.Media.VisualTreeHelper]::GetChildrenCount($ctNode)
        for ($ctI = 0; $ctI -lt $ctN; $ctI++) { & $ctWalk ([Windows.Media.VisualTreeHelper]::GetChild($ctNode, $ctI)) }
    }
    & $ctWalk $script:ctWindow.Content
    return , $ctList
}

function Format-CtError($ctRecord) {
    <# The message of the exception that was raised first, and where: the handler of a control wraps the error of its block. #>
    $ctEx = $ctRecord.Exception
    $ctInner = $ctEx
    while ($ctInner.InnerException) { $ctInner = $ctInner.InnerException }
    $ctWhere = $null
    foreach ($ctE in $ctEx, $ctInner) { $ctR = $ctE.PSObject.Properties['ErrorRecord']; if ($ctR -and $ctR.Value -and $ctR.Value.ScriptStackTrace) { $ctWhere = $ctR.Value.ScriptStackTrace; break } }
    if (-not $ctWhere) { $ctWhere = $ctRecord.ScriptStackTrace }
    return ("$($ctInner.Message) @ " + (($ctWhere -split "`n" | Select-Object -First 3) -join ' < '))
}

function Add-CtCount([string]$ctType) { if (-not $script:ctInvoked.ContainsKey($ctType)) { $script:ctInvoked[$ctType] = 0 }; $script:ctInvoked[$ctType]++ }

function Invoke-CtControl($ctControl) {
    <# One use of a control, the way a person does it. The exception of a handler comes out of here. #>
    $ctClick = [Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent, $ctControl)
    if ($ctControl -is [Windows.Controls.RadioButton]) {
        Add-CtCount 'RadioButton'
        if (-not [bool]$ctControl.IsChecked) { $ctControl.IsChecked = $true }
    } elseif ($ctControl -is [Windows.Controls.CheckBox]) {
        Add-CtCount 'CheckBox'
        $ctControl.IsChecked = -not [bool]$ctControl.IsChecked
        $ctControl.RaiseEvent($ctClick)
        $ctControl.IsChecked = -not [bool]$ctControl.IsChecked
    } elseif ($ctControl -is [Windows.Controls.Primitives.ToggleButton]) {
        Add-CtCount 'ToggleButton'
        $ctControl.IsChecked = -not [bool]$ctControl.IsChecked
        $ctControl.RaiseEvent($ctClick)
    } elseif ($ctControl -is [Windows.Controls.Primitives.ButtonBase]) {
        Add-CtCount 'Button'
        $ctControl.RaiseEvent($ctClick)
    } elseif ($ctControl -is [Windows.Controls.ComboBox]) {
        Add-CtCount 'ComboBox'
        for ($ctI = 0; $ctI -lt $ctControl.Items.Count; $ctI++) { $ctControl.SelectedIndex = $ctI }
    } elseif ($ctControl -is [Windows.Controls.TextBox]) {
        Add-CtCount 'TextBox'
        $ctOld = $ctControl.Text
        $ctControl.Text = $ctOld + 'x'
        $ctControl.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.UIElement]::LostFocusEvent, $ctControl))
        $ctControl.Text = $ctOld
    } elseif ($ctControl -is [Windows.Controls.Expander]) {
        Add-CtCount 'Expander'
        $ctControl.IsExpanded = -not $ctControl.IsExpanded
        $ctControl.IsExpanded = -not $ctControl.IsExpanded
    } elseif ($ctControl -is [Windows.Controls.TreeViewItem]) {
        Add-CtCount 'TreeViewItem'
        $ctControl.IsSelected = $true
    } elseif ($ctControl -is [Windows.Controls.ListBoxItem]) {
        Add-CtCount 'ListBoxItem'
        $ctControl.IsSelected = $true
    }
    Wait-CtDispatcher
}

function Restore-CtScenario([string]$ctName, [scriptblock]$ctExtra) {
    <# A clean state, the scenario, the folds opened, the layout done. #>
    $ctG = $script:Gui
    if ($ctG.DialogOpen) { Close-MbtGuiDialog }
    $ctG.S = New-MbtGuiState -WorkFolder 'D:\MailboxTransfer'; $ctG.Run = $null; $ctG.Job = $null
    foreach ($ctB in $ctG.Browsers.Values) { Set-MbtGuiBrowserRoots -Browser $ctB -Roots @() }
    Set-MbtGuiScenario -Name $ctName
    $ctG.S.WorkFolder = $Work
    if ($ctExtra) { & $ctExtra }
    # the console never shows a run, or a result, without what led to it: the source of an import was opened, the run has its job
    if ($ctG.S.Task -eq 'import' -and $ctG.S.ImPath -and -not $ctG.S.ImInfo) { Open-MbtGuiSource -Use import }
    # a run on screen has its job (the console never shows a run without one)
    if ($ctG.Run -and -not $ctG.Job -and $ctG.S.Task) { $ctG.Job = New-MbtGuiJob -WorkFolder $Work -Task $ctG.S.Task -EntryScript 'x.ps1' }
    Measure-CtWindow
    foreach ($ctE in @(Get-CtControl | Where-Object { $_ -is [Windows.Controls.Expander] -and -not $_.IsExpanded })) { $ctE.IsExpanded = $true }
    Measure-CtWindow
}

function Invoke-CtScenario([string]$ctName, [scriptblock]$ctExtra) {
    $script:ctScenarios++
    $script:ctWalked.Add($ctName)
    $ctG = $script:Gui
    Restore-CtScenario $ctName $ctExtra
    $ctScreen = Get-MbtGuiScreen
    $ctCount = (Get-CtControl).Count
    $ctStale = $false
    for ($ctK = 0; $ctK -lt $ctCount; $ctK++) {
        if ($ctStale) { Restore-CtScenario $ctName $ctExtra; $ctStale = $false }
        $ctControls = Get-CtControl
        if ($ctK -ge $ctControls.Count) { continue }
        $ctControl = $ctControls[$ctK]
        $ctLabel = "$ctName #$ctK $($ctControl.GetType().Name) '$(Get-CtText $ctControl)'"
        # the header, the stepper and the action bar are the same on every page of a kind: each of their controls is used once per page of the flow
        $ctInPage = $false
        for ($ctUp = [Windows.Media.VisualTreeHelper]::GetParent($ctControl); $ctUp; $ctUp = [Windows.Media.VisualTreeHelper]::GetParent($ctUp)) { if ([object]::ReferenceEquals($ctUp, $ctG.C.Page)) { $ctInPage = $true; break } }
        if (-not $ctInPage) {
            $ctKey = "$ctScreen|$($ctControl.GetType().Name)|$(Get-CtText $ctControl)"
            if ($script:ctChromeSeen.ContainsKey($ctKey) -and -not $ctG.DialogOpen) { continue }
            $script:ctChromeSeen[$ctKey] = $true
        }
        $ctId = [Windows.Automation.AutomationProperties]::GetAutomationId($ctControl)
        if ($ctId -like "RunStep*") { $script:ctStepControls[$ctId] = 1 + [int]$script:ctStepControls[$ctId] }
        $ctPage = $ctG.C.Page.Content
        $script:ctControls++
        try { Invoke-CtControl $ctControl } catch { $script:ctErrors.Add("$ctLabel : $(Format-CtError $_)") }
        if ($ctG.DialogOpen) {
            # every answer of the dialog this control opened
            $ctAnswers = $ctG.C.DialogButtons.Children.Count
            for ($ctJ = 0; $ctJ -lt $ctAnswers; $ctJ++) {
                Restore-CtScenario $ctName $ctExtra
                $ctControls = Get-CtControl
                $script:ctDialogs++
                try {
                    Invoke-CtControl $ctControls[$ctK]
                    $ctAnswer = $ctG.C.DialogButtons.Children[$ctJ]
                    $ctLabelAnswer = "$ctLabel -> dialog '$($ctG.C.DialogTitle.Text)' answer '$($ctAnswer.Content)'"
                    $ctAnswer.RaiseEvent([Windows.RoutedEventArgs]::new([Windows.Controls.Primitives.ButtonBase]::ClickEvent, $ctAnswer))
                    Wait-CtDispatcher
                } catch { $script:ctErrors.Add("$ctLabelAnswer : $(Format-CtError $_)") }
            }
            $ctStale = $true
        }
        if (-not [object]::ReferenceEquals($ctPage, $ctG.C.Page.Content)) { $ctStale = $true }
    }
    # the timer of the console ticks (events, worker) while the page is shown
    $ctUntil = [datetime]::UtcNow.AddMilliseconds(150)
    while ([datetime]::UtcNow -lt $ctUntil) { Wait-CtDispatcher; Start-Sleep -Milliseconds 20 }
    if ($ctG.LastTickError) { $script:ctErrors.Add("$ctName : the tick of the console raised: $($ctG.LastTickError)"); $ctG.LastTickError = $null }
}

function Invoke-CtClickThrough {
    $ctIndex = 0
    foreach ($ctName in (@(Get-MbtGuiScenarioName) + @(Get-MbtGuiEdgeScenarioName))) {
        if ((-not $Only -or $ctName -in $Only) -and ($ctIndex % $Parts) -eq $Part) { Invoke-CtScenario $ctName $null }
        $ctIndex++
    }
    # the extra states are walked by the last part (or by -Only extras)
    if ($Only -and 'extras' -notin $Only) { return }
    if (-not $Only -and $Part -ne ($Parts - 1)) { return }
    # states the scenarios do not reach by themselves
    Invoke-CtScenario 'Setup' { $script:Gui.S.AuthEdit = $true; Show-MbtGuiPage }
    Invoke-CtScenario 'SetupApplication' { $script:Gui.S.AuthEdit = $true; Show-MbtGuiPage }
    Invoke-CtScenario 'ExportFrom' { $script:Gui.S.Shared = 'projects@contoso.com'; Show-MbtGuiPage }
    Invoke-CtScenario 'ImportTo' { $script:Gui.S.ImShared = 'projects@contoso.com'; $script:Gui.S.TgtMode = 'same'; Show-MbtGuiPage }
    Invoke-CtScenario 'ImportPlan' { $script:Gui.S.ImSrc = 'store'; $script:Gui.S.ImPath = 'D:\Exports\adele'; Show-MbtGuiPage }
    Invoke-CtScenario 'ExportTo' { $script:Gui.S.Dest = 'store'; Show-MbtGuiPage }
    Invoke-CtScenario 'ExportTo' { $script:Gui.S.Dest = 'eml'; Show-MbtGuiPage }
    Invoke-CtScenario 'Start' { $script:Gui.S.Sel = 'copy'; Show-MbtGuiPage }
    foreach ($ctTheme in 'Dark', 'System', 'Light') { $script:Gui.C["Th_$ctTheme"].IsChecked = $true; Wait-CtDispatcher; Invoke-CtScenario 'ExportPlan' $null }
    # closing the window
    $script:Gui.ForceClose = $false; $script:Gui.Job = $null; $script:Gui.Run = $null
}

Initialize-MbtGuiTypes
$null = New-Item -ItemType Directory -Force -Path $Work
$ctEngine = New-MbtGuiFakeEngine -Capability @{ PstImport = $true; EmlExport = $true; EmlImport = $true; SharedMailbox = $true; FolderChoice = $true; PstCopy = $true; PstPreview = $true; AppOnly = $true }
$script:ctWindow = New-MbtGuiWindow -Engine $ctEngine -Theme Light -WorkFolder $Work -EntryScript (Join-Path $Package 'module\MailboxTransfer.Gui.DemoJob.ps1') -Version '1.0.0'
$script:ctWindow.WindowStartupLocation = 'Manual'; $script:ctWindow.Left = 0; $script:ctWindow.Top = 0
$script:ctWindow.Width = 1496; $script:ctWindow.Height = 979; $script:ctWindow.MinWidth = 0; $script:ctWindow.MinHeight = 0
$script:ctWindow.Opacity = 0; $script:ctWindow.ShowInTaskbar = $false; $script:ctWindow.ShowActivated = $false
# the walk is started from inside ShowDialog, so every handler is raised under the dispatcher of the window, as when a person clicks
$script:ctDriver = [Windows.Threading.DispatcherTimer]::new([Windows.Threading.DispatcherPriority]::Background)
$script:ctDriver.Interval = [TimeSpan]::FromMilliseconds(300)
$script:ctDriver.Add_Tick({
        $script:ctDriver.Stop()
        try { Invoke-CtClickThrough } catch { $script:ctErrors.Add("the walk stopped: $(Format-CtError $_)") }
        finally { $script:Gui.ForceClose = $true; $script:ctWindow.Close() }
    })
$script:ctWindow.Add_ContentRendered({ $script:ctDriver.Start() })
[void]$script:ctWindow.ShowDialog()
[pscustomobject]@{ Errors = @($script:ctErrors); Invoked = $script:ctInvoked; Scenarios = $script:ctScenarios; Controls = $script:ctControls; Dialogs = $script:ctDialogs; StepControls = $script:ctStepControls; Walked = @($script:ctWalked); Edge = @(Get-MbtGuiEdgeScenarioName) } | ConvertTo-Json -Depth 4 -Compress
