#Requires -Version 7.4
<#
    Headless smoke test of the console: builds the real window (not shown), walks the state machine on the fake engine, feeds a fake events file
    (running, PC slept and resumed, needs a decision, stopped, complete, error), checks the pages build and that nothing throws.
    Prints one JSON object: { Errors: [..], Checks: [ { Name; Ok; Detail } ] }. Nothing contacts the cloud: the engine is the fake one.
    Run by tests\Gui.Smoke.Tests.ps1 in its own single-threaded pwsh (WPF needs STA).
#>
param([string]$Package, [string]$Work)
Set-StrictMode -Version Latest
[AppDomain]::CurrentDomain.SetData('MailboxTransfer.UnverifiedTreeWarned', $true)   # a source checkout has no PackageIntegrity.psd1: no warning in the output that the test reads
$ErrorActionPreference = 'Stop'
foreach ($f in 'Core', 'Run', 'Jobs', 'Browse', 'State', 'Pages', 'Pages.Export', 'Pages.Run', 'Pages.Source', 'Fake', 'Engine', 'Scenarios') { . (Join-Path $Package "module\MailboxTransfer.Gui.$f.ps1") }
$script:GuiRoot = $Package
$script:Errors = [Collections.Generic.List[string]]::new()
$script:Checks = [Collections.Generic.List[object]]::new()
function Check([string]$CheckName, [scriptblock]$CheckBlock) {
    try { $r = & $CheckBlock; $ok = [bool]$r; $script:Checks.Add([pscustomobject]@{ Name = $CheckName; Ok = $ok; Detail = $(if (-not $ok) { "returned $r" } else { '' }) }) }
    catch { $script:Errors.Add("$CheckName : $($_.Exception.Message) at $($_.ScriptStackTrace.Split("`n")[0])"); $script:Checks.Add([pscustomobject]@{ Name = $CheckName; Ok = $false; Detail = $_.Exception.Message }) }
}
function Wait-Dispatcher {
    $frame = [Windows.Threading.DispatcherFrame]::new()
    [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::ContextIdle, [Action] { $frame.Continue = $false })
    [Windows.Threading.Dispatcher]::PushFrame($frame)
}
function Measure-Window($w) { $r = $w.Content; $r.Measure([Windows.Size]::new(1480, 940)); $r.Arrange([Windows.Rect]::new(0, 0, 1480, 940)); $r.UpdateLayout(); Wait-Dispatcher }
function Reset-State { $g = $script:Gui; $g.S = New-MbtGuiState -WorkFolder $Work; $g.Run = $null; $g.Job = $null; foreach ($b in $g.Browsers.Values) { Set-MbtGuiBrowserRoots -Browser $b -Roots @() } }
function Find-Text($root, [string]$Text) {
    $script:hit = $false
    $walk = {
        param($o)
        if ($null -eq $o) { return }
        if ($o -is [Windows.Controls.TextBlock]) {
            if ($o.Text -like "*$Text*") { $script:hit = $true }
            foreach ($i in $o.Inlines) { if ($i -is [Windows.Documents.Run] -and $i.Text -like "*$Text*") { $script:hit = $true } }
        }
        if ($o -is [Windows.Controls.ContentControl] -and $o.Content -is [string] -and $o.Content -like "*$Text*") { $script:hit = $true }
        if ($o -is [Windows.Controls.ContentControl] -and $o.Content -is [Windows.DependencyObject]) { & $walk $o.Content }
        if ($o -is [Windows.Controls.Expander]) { & $walk $o.Header }
        if ($o -is [Windows.Controls.Decorator]) { & $walk $o.Child }
        if ($o -is [Windows.Controls.Panel]) { foreach ($c in $o.Children) { & $walk $c } }
    }
    & $walk $root
    return $script:hit
}
function Get-TextBlockText($tb) {
    # the text of a block, whether it was set as a string or built from runs (bold parts)
    if ($tb.Text) { return $tb.Text }
    $inline = $null
    $inline = { param($i) if ($i -is [Windows.Documents.Run]) { $i.Text } elseif ($i -is [Windows.Documents.Span]) { foreach ($j in $i.Inlines) { & $inline $j } } }
    return ((@($tb.Inlines | ForEach-Object { & $inline $_ })) -join '')
}
function Get-AllText($root) {
    $texts = [Collections.Generic.List[string]]::new()
    $walk = {
        param($o)
        if ($null -eq $o) { return }
        if ($o -is [Windows.Controls.TextBlock]) { $texts.Add((Get-TextBlockText $o)) }
        if ($o -is [Windows.Controls.ContentControl] -and $o.Content -is [Windows.DependencyObject]) { & $walk $o.Content }
        if ($o -is [Windows.Controls.Expander]) { & $walk $o.Header }
        if ($o -is [Windows.Controls.Decorator]) { & $walk $o.Child }
        if ($o -is [Windows.Controls.Panel]) { foreach ($c in $o.Children) { & $walk $c } }
    }
    & $walk $root
    return $texts.ToArray()
}
function Get-CommandLineText {
    # the text of the box "The same from the command line" of the page that is shown
    $found = [Collections.Generic.List[string]]::new()
    $walk = {
        param($o)
        if ($o -is [Windows.Controls.TextBox] -and [Windows.Automation.AutomationProperties]::GetName($o) -eq 'The same from the command line') { $found.Add($o.Text) }
        elseif ($o -is [Windows.DependencyObject]) { foreach ($c in [Windows.LogicalTreeHelper]::GetChildren($o)) { & $walk $c } }
    }
    & $walk $script:Gui.C.Page.Content
    return ($found -join "`n")
}
function Get-StepperState {
    $out = [Collections.Generic.List[object]]::new()
    $walk = {
        param($o)
        if ($o -is [Windows.Controls.Button]) { $out.Add([pscustomobject]@{ Label = ((Get-AllText $o.Content) -join ''); Enabled = $o.IsEnabled }) }
        elseif ($o -is [Windows.DependencyObject]) { foreach ($c in [Windows.LogicalTreeHelper]::GetChildren($o)) { & $walk $c } }
    }
    & $walk $script:Gui.C.Stepper
    return $out.ToArray()
}
function New-FakeJob([string]$Task) {
    $dir = New-MbtGuiJobFolder -WorkFolder $Work -Task "smoke-$Task"
    return @{ Id = "smoke-$Task"; Task = $Task; Folder = $dir; EventsPath = (Join-Path $dir 'events.jsonl'); StopPath = (Join-Path $dir 'stop.request'); Offset = 0L; Process = $null; Queue = [Collections.Generic.Queue[object]]::new(); StopAsked = $null; Exited = $false; Fake = $true; StorePath = (Join-Path $dir 'store'); JobName = 'x' }
}
function Add-Line($job, [hashtable]$e) { $e.v = 1; $e.t = [datetime]::UtcNow.ToString('o'); ($e | ConvertTo-Json -Compress -Depth 5) | Add-Content -LiteralPath $job.EventsPath -Encoding utf8 }
function Start-Fake([string]$Task) {
    $g = $script:Gui
    Reset-State; $g.S.Task = $Task; $g.S.Sel = $Task; $g.S.Mailbox = 'adele.vance@contoso.com'; $g.S.Signed = $true; $g.S.OutFolder = 'D:\Exports\adele'; $g.S.OutName = 'adele.pst'; $g.S.ImPath = 'D:\PST\a.pst'; $g.S.RpPath = 'D:\PST\a.pst'; $g.S.RpDest = 'D:\PST\a.repaired.pst'
    $job = New-FakeJob $Task; $g.Job = $job; $g.Run = New-MbtGuiRun -Task $Task -Title "Smoke $Task"
    $g.S.Step = $script:Flows[$Task].IndexOf('run'); Show-MbtGuiPage
    $script:job = $job
}
function Pump($job, $window) { $null = Step-MbtGuiJob -Job $job -Run $script:Gui.Run; Sync-MbtGuiRunPage -Force; Measure-Window $window }

Initialize-MbtGuiTypes
$null = New-Item -ItemType Directory -Force -Path $Work
$engine = New-MbtGuiFakeEngine -Capability @{ PstImport = $false; EmlExport = $false }
$entry = Join-Path $Package 'module\MailboxTransfer.Gui.DemoJob.ps1'
foreach ($theme in 'Light', 'Dark') {
    $script:window = $null
    Check "$theme window builds" { $script:window = New-MbtGuiWindow -Engine $engine -Theme $theme -WorkFolder $Work -EntryScript $entry -Version '0.3.0'; Measure-Window $script:window; $true }
    $window = $script:window
    $g = $script:Gui
    foreach ($name in (@(Get-MbtGuiScenarioName) + @(Get-MbtGuiEdgeScenarioName))) {
        Check "$theme scenario $name" { Reset-State; Set-MbtGuiScenario -Name $name; Measure-Window $window; $null -ne $g.C.Page.Content }
    }
    if ($theme -eq 'Dark') { Check 'dark theme brush' { $window.Resources['MbtBrandText'].Color.ToString() -eq '#FFFD8EA1' } } else { Check 'light theme brush' { $window.Resources['MbtBrandText'].Color.ToString() -eq '#FFB11F4B' } }
    if ($theme -eq 'Light') {
        Check 'header: the theme switch shows System, Light and Dark, the opening theme checked' { $g.C.Th_Light.IsChecked -and -not $g.C.Th_Dark.IsChecked -and -not $g.C.Th_System.IsChecked -and $g.C.Th_Dark.Content -eq 'Dark' }
        Check 'header: choosing Dark applies the dark look to the window and draws the page again' {
            $g.C.Th_Dark.IsChecked = $true; Measure-Window $window
            $g.Theme.Dark -and $window.Resources['MbtBrandText'].Color.ToString() -eq '#FFFD8EA1' -and $null -ne $g.C.Page.Content }
        Check 'header: choosing Light again brings the light look back' {
            $g.C.Th_Light.IsChecked = $true; Measure-Window $window
            -not $g.Theme.Dark -and $window.Resources['MbtBrandText'].Color.ToString() -eq '#FFB11F4B' }
        Check 'header: version and author are shown, with the links to the guide and to the latest release' { $g.C.Version.Text -match 'v0\.3\.0.*Nicolas Fabert' -and $g.C.HelpGuide.Content -eq 'User guide' -and $g.C.HelpRelease.Content -eq 'Latest release' }
    }
    if ($theme -eq 'Dark') { $g.ForceClose = $true; $window.Close(); continue }

    Check 'setup is the first page: the lights, the configuration file, the work folder and the sign-in; no Keep awake' {
        Reset-State; Show-MbtGuiPage; Measure-Window $window
        (Get-MbtGuiScreen) -eq 'setup' -and (Find-Text $g.C.Page.Content 'This computer') -and (Find-Text $g.C.Page.Content 'Configuration file') -and (Find-Text $g.C.Page.Content 'Work folder') -and -not (Find-Text $g.C.Page.Content 'Keep this PC awake') -and -not (Find-Text $g.C.Page.Content 'While a job runs') -and (Find-Text $g.C.Page.Content 'Me (delegated, recommended)') -and (Find-Text $g.C.Page.Content 'A Microsoft sign-in window opens over the console') -and (Find-Text $g.C.Page.Content 'contoso.onmicrosoft.com') -and $g.C.PrimaryLabel.Text -eq 'Choose a task' }
    Check 'setup: a value changed for this run is marked "modified", and the file is not touched' {
        $g.S.TenantId = 'example.onmicrosoft.com'; Show-MbtGuiPage; $a = (Find-Text $g.C.Page.Content 'this run only') -and (Find-Text $g.C.Page.Content 'example.onmicrosoft.com')
        $g.S.TenantId = 'CONTOSO.onmicrosoft.com'; Show-MbtGuiPage; $b = -not (Find-Text $g.C.Page.Content 'modified')
        $g.S.TenantId = ''; Show-MbtGuiPage; $a -and $b }
    Check 'start: primary disabled until a task is chosen' { Reset-State; $g.S.Step = 1; Show-MbtGuiPage; -not $g.C.Primary.IsEnabled }
    Check 'start: choosing a task enables Continue' { $g.S.Sel = 'export'; Show-MbtGuiPage; $g.C.Primary.IsEnabled }
    Check 'start: Continue goes to From' { & $g.PrimaryAction; (Get-MbtGuiScreen) -eq 'exFrom' }
    Check 'from: five steps, only the passed ones are clickable' {
        $st = Get-StepperState
        $top = @($st | Where-Object { $_.Label -match '(Setup|Verify|Run|Next steps)$' })
        $top.Count -eq 4 -and ($st | Where-Object { $_.Label -match 'Setup$' }).Enabled -and ($st | Where-Object { $_.Label -eq 'What' }).Enabled -and -not ($st | Where-Object { $_.Label -eq 'To' }).Enabled -and -not ($st | Where-Object { $_.Label -match 'Verify$' }).Enabled }
    Check 'from: continue disabled before sign-in' { -not $g.C.Primary.IsEnabled }
    Check 'setup: no application option when the engine cannot sign in with a certificate' { $g.Engine.Capability.AppOnly = $false; $g.S.Step = 0; Show-MbtGuiPage; -not (Find-Text $g.C.Page.Content 'The application (certificate)') }
    Check 'setup: the application option lights up with the capability' { $g.Engine.Capability.AppOnly = $true; Show-MbtGuiPage; (Find-Text $g.C.Page.Content 'The application (certificate)') -and -not (Find-Text $g.C.Page.Content 'Certificate thumbprint') }
    Check 'setup: the application shows its tenant, application id and certificate from the configuration file; Change for this run opens the boxes' {
        $g.S.AuthMode = 'app'; Show-MbtGuiPage
        $a = (Find-Text $g.C.Page.Content 'Values for the application') -and (Find-Text $g.C.Page.Content 'Certificate thumbprint') -and (Find-Text $g.C.Page.Content 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678') -and (Find-Text $g.C.Page.Content 'Change for this run')
        $g.S.AuthEdit = $true; Show-MbtGuiPage; $b = Find-Text $g.C.Page.Content 'Use the values of the configuration file'; $g.S.AuthEdit = $false; $a -and $b }
    Check 'sign-in card (From): as the application it asks for the mailbox only, and offers no code' { $g.S.Step = 2; Show-MbtGuiPage; (Find-Text $g.C.Page.Content 'Mailbox the application works on') -and (Find-Text $g.C.Page.Content 'Change in Setup') -and -not (Find-Text $g.C.Page.Content 'Use a code instead') }
    Check 'sign-in as the application: the mailbox address is required' { $g.S.Mailbox = ''; Start-MbtGuiSignIn; $g.S.SignInError -match 'address' }
    Check 'sign-in as the application: tenant, application id and thumbprint are checked in plain words' {
        $g.S.Mailbox = 'adele.vance@contoso.com'; $g.S.AppTenant = 'common'; Start-MbtGuiSignIn; $a = $g.S.SignInError -match 'tenant'
        $g.S.AppTenant = ''; $g.S.AppClient = 'nope'; Start-MbtGuiSignIn; $b = $g.S.SignInError -match 'client'
        $g.S.AppClient = ''; $g.S.AppThumb = 'ABCD'; Start-MbtGuiSignIn; $c = $g.S.SignInError -match 'thumbprint'
        $a -and $b -and $c -and -not $g.S.Signed }
    Check 'sign-in as the application (fake) succeeds with the values of the file and says so' { $g.S.AppThumb = ''; $g.S.SignInError = ''; Start-MbtGuiSignIn; $ok = $g.S.Signed -and $g.S.Token.Access -match 'Application' -and (Find-Text $g.C.Page.Content 'Connected as the application'); $g.S.Signed = $false; $g.S.Locations = @(); $g.S.AuthMode = 'user'; $g.Engine.Capability.AppOnly = $false; Show-MbtGuiPage; $ok }
    Check 'sign-in with a bad address says so' { $g.S.Mailbox = 'nope'; Start-MbtGuiSignIn; $g.S.SignInError -match 'address' }
    Check 'sign-in: "Use a code instead" while the window is waited for ends that wait and asks for the code; a closed window or a late one is said on the page, not waited for' {
        $keep = $g.Engine.SignIn; $script:smCalls = [Collections.Generic.List[object]]::new()
        $g.Engine.SignIn = { param($Mailbox, $UseCode, $CodeFile, $OnDone, $OnError, $Silent, $TenantId, $ClientId, $App, $ParentWindow, $LogFolder, $Cancel) $script:smCalls.Add(@{ UseCode = $UseCode; Cancel = $Cancel; OnError = $OnError; Log = $LogFolder }) }
        try {
            $g.S.Mailbox = 'adele.vance@contoso.com'; $g.S.Signed = $false; $g.S.SignInError = ''
            Start-MbtGuiSignIn
            $first = $script:smCalls[0]
            $a = $g.S.SigningIn -and -not $first.UseCode -and $first.Cancel -and -not $first.Cancel.IsCancellationRequested -and $first.Log -match 'logs$'
            Start-MbtGuiSignIn -UseCode $true
            $b = $first.Cancel.IsCancellationRequested -and $script:smCalls.Count -eq 1
            & $first.OnError 'The operation was canceled.'
            $c = $script:smCalls.Count -eq 2 -and $script:smCalls[1].UseCode -and $g.S.SigningIn -and -not $g.S.SignInError
            & $script:smCalls[1].OnError 'The sign-in window was closed before the sign-in was finished. Sign in again, or sign in with a code.'
            $d = -not $g.S.SigningIn -and $g.S.SignInError -match 'closed before' -and $g.S.SignInError -match 'use a code instead'
            Start-MbtGuiSignIn
            & $script:smCalls[2].OnError 'The sign-in window did not finish within 5 minutes. Sign in again, or sign in with a code.'
            $e = -not $g.S.SigningIn -and $g.S.SignInError -match 'did not finish in time'
            $a -and $b -and $c -and $d -and $e
        } finally { $g.Engine.SignIn = $keep; $g.S.SigningIn = $false; $g.S.SignInError = ''; $g.S.CodeShown = $false; $g.S.SignInCts = $null; $g.S.SwitchToCode = $false; Show-MbtGuiPage }
    }
    Check 'sign-in (fake) succeeds' { $g.S.Mailbox = 'adele.vance@contoso.com'; $g.S.SignInError = ''; Start-MbtGuiSignIn; $g.S.Signed -and @($g.S.Locations).Count -eq 3 }
    Check 'browse: three roots' { (Get-MbtGuiBrowser 'export').Roots.Count -eq 3 }
    Check 'from: the box "Include deleted items (Recoverable Items)" is off by default, with its help in plain words' {
        Measure-Window $window
        (Find-Text $g.C.Page.Content 'Include deleted items (Recoverable Items)') -and (Find-Text $g.C.Page.Content 'What are deleted items') -and (Find-Text $g.C.Page.Content 'never back into Recoverable Items') -and -not $g.S.RecoverableItems `
            -and -not (@((Get-MbtGuiBrowser 'export').Roots[0].Descendants()) | Where-Object Name -EQ 'Recoverable Items') }
    Check 'from: ticking the box reloads the browser with Recoverable Items (Deletions, Purges...), unticking takes them away' {
        $g.S.RecoverableItems = $true; Sync-MbtGuiLocationRoots
        $on = @((Get-MbtGuiBrowser 'export').Roots[0].Descendants()) | Where-Object Name -EQ 'Recoverable Items'
        $kids = if ($on) { @($on.Descendants() | Where-Object Name -In 'Deletions', 'Purges').Count } else { 0 }
        $g.S.RecoverableItems = $false; Sync-MbtGuiLocationRoots
        $off = @((Get-MbtGuiBrowser 'export').Roots[0].Descendants()) | Where-Object Name -EQ 'Recoverable Items'
        $on -and $kids -eq 2 -and -not $off }
    Check 'plan: ticked deleted items are recognised (the plan says they are saved apart), other ticks are not' {
        $tree = @(Get-MbtGuiFakeTree -Recoverable)
        $del = @($tree[0].Descendants()) | Where-Object Id -EQ 'rec-del'; $del.Check = $true
        $withDeleted = Get-MbtGuiSelection -Roots @($tree[0])
        $tree2 = @(Get-MbtGuiFakeTree -Recoverable)
        (@($tree2[0].Descendants()) | Where-Object Id -EQ 'sent').Check = $true
        $other = Get-MbtGuiSelection -Roots @($tree2[0])
        (Test-MbtGuiSelectionHasDeletedItems $withDeleted) -and -not (Test-MbtGuiSelectionHasDeletedItems $other) -and $withDeleted.Items -eq 24 }
    Check 'ticking a folder ticks its subfolders and makes the parent indeterminate' {
        $b = Get-MbtGuiBrowser 'export'; $nodes = @($b.Roots[0].Descendants()); $inbox = $nodes | Where-Object Id -eq 'inbox'; $pc = $nodes | Where-Object Id -eq 'p-c'
        $pc.Check = $true; $sel = Get-MbtGuiSelection -Roots @($b.Roots)
        ($null -eq $inbox.Check) -and ($sel.Items -eq 410) -and $sel.Entries.Count -eq 1 }
    Check 'unknown size is counted apart, never summed' { $b = Get-MbtGuiBrowser 'export'; $x = @($b.Roots[2].Descendants()) | Where-Object Id -eq 'x-18'; $x.Check = $true; $sel = Get-MbtGuiSelection -Roots @($b.Roots); $sel.UnknownSize -eq 1 -and $sel.Bytes -eq [long](1.0GB) }
    Check 'messages: ticking one makes the folder indeterminate' {
        $b = Get-MbtGuiBrowser 'export'; $inbox = @($b.Roots[0].Descendants()) | Where-Object Id -eq 'inbox'; $inbox.IsSelected = $true; Open-MbtGuiFolder $inbox
        $b.Rows[0].IsChecked = $true; ($null -eq $inbox.Check) -and $inbox.Picks.Count -eq 1 }
    Check 'headers-only archive: subject not available' { $b = Get-MbtGuiBrowser 'export'; $x = @($b.Roots[2].Descendants()) | Where-Object Id -eq 'x-18'; Open-MbtGuiFolder $x; $b.Rows.Count -gt 0 -and $b.Rows[0].SubjectNa }
    Check 'continue to To' { Show-MbtGuiPage; $en = $g.C.Primary.IsEnabled; Move-MbtGuiStep 1; $en -and (Get-MbtGuiScreen) -eq 'exTo' }
    Check 'EML is enabled when the engine has it' { $g.Engine.Capability.EmlExport = $true; Show-MbtGuiPage; $ok = (Find-Text $g.C.Page.Content 'EML files') -and -not (Find-Text $g.C.Page.Content 'EML files (coming'); $g.Engine.Capability.EmlExport = $false; Show-MbtGuiPage; $ok }
    Check 'EML is shown disabled when the engine lacks it' { Find-Text $g.C.Page.Content 'EML files (coming in this version)' }
    Check 'reserved PST names are refused, other names are fine' { (Test-MbtGuiPstName 'aux.pst') -match 'AUX' -and (Test-MbtGuiPstName 'Aux.Backup.pst') -and (Test-MbtGuiPstName 'com3.pst') -and -not (Test-MbtGuiPstName 'auxiliary.pst') -and -not (Test-MbtGuiPstName 'adele.pst') }
    Check 'To: a refused name greys Continue' { $g.S.OutName = 'aux.pst'; Show-MbtGuiPage; -not $g.C.Primary.IsEnabled }
    Check 'To: a good name enables Continue' { $g.S.OutName = 'adele.pst'; $g.S.OutFolder = 'D:\Exports\adele'; Show-MbtGuiPage; $g.C.Primary.IsEnabled }
    Check 'To: keep the working store is on by default' { (New-MbtGuiState -WorkFolder $Work).KeepStore -eq $true }
    Check 'plan: a valid plan unlocks Start' { Move-MbtGuiStep 1; Measure-Window $window; (Get-MbtGuiScreen) -eq 'exPlan' -and $g.S.PlanState -eq 'ok' -and $g.C.Primary.IsEnabled }
    Check 'plan: Keep this PC awake is in Expert mode, off by default; ticking it adds -KeepAwake to the shown command line' {
        Show-MbtGuiPage; Measure-Window $window
        $off = (Find-Text $g.C.Page.Content 'Keep this PC awake while it runs') -and (Find-Text $g.C.Page.Content 'Windows follows your company') -and -not $g.S.Awake -and ((Get-CommandLineText) -notmatch '-KeepAwake')
        $g.S.Awake = $true; Show-MbtGuiPage; $on = (Get-CommandLineText) -match '-KeepAwake'
        $g.S.Awake = $false; Show-MbtGuiPage
        $off -and $on }
    Check 'plan: no typed confirmation for export' { -not (Find-Text $g.C.Page.Content 'Type IMPORT') }
    Check 'plan: a range, not one number' { Find-Text $g.C.Page.Content 'between' }
    Check 'plan: the command line is the one an operator types: one line, the selection file, the PST, no plumbing' {
        $t = Get-CommandLineText
        $t -match '^\.\\Invoke-MailboxTransfer\.ps1 -Mode Export -UserPrincipalName adele\.vance@contoso\.com -SelectionFile \S+selections\\export-[0-9a-f]{8}\.json' -and $t -match '-OutputPst D:\\Exports\\adele\\adele\.pst' -and $t -notmatch '-NoPrompt|-EventsFile|-StopFile|-JobId|-TempPath|ExportPst|-WhatIf' -and $t.Split("`n").Count -eq 1 }
    Check 'plan: changing a setting voids the plan key' { $k1 = Get-MbtGuiPlanKey; $g.S.Conc = 2; $k2 = Get-MbtGuiPlanKey; $g.S.Conc = 4; $k1 -ne $k2 }
    Check 'plan: the case of what is typed does not void the plan key; another folder does' {
        $k1 = Get-MbtGuiPlanKey; $g.S.Mailbox = 'ADELE.Vance@Contoso.com'; $g.S.OutFolder = $g.S.OutFolder.ToUpper(); $g.S.OutName = $g.S.OutName.ToUpper(); $k2 = Get-MbtGuiPlanKey
        $g.S.Mailbox = 'adele.vance@contoso.com'; $g.S.OutFolder = 'D:\Exports\adele'; $g.S.OutName = 'adele.pst'
        $k1 -eq $k2 }

    $script:job = $null
    Check 'run: plan and progress events build the page' {
        Start-Fake 'import'
        Add-Line $script:job @{ e = 'plan'; job = 'x'; steps = @(@{ id = 'check'; title = 'Check the PST file' }, @{ id = 'send'; title = 'Send the items' }, @{ id = 'reconcile'; title = 'Check the uncertain items' }) }
        Add-Line $script:job @{ e = 'step'; id = 'check'; state = 'done' }
        Add-Line $script:job @{ e = 'step'; id = 'send'; state = 'running'; done = 100; total = 1000; bytes = 1GB; bytesTotal = 10GB }
        Add-Line $script:job @{ e = 'object'; step = 'send'; key = 'Inbox'; state = 'running'; done = 100; total = 400 }
        Pump $script:job $window; $g.Run.Steps.Count -eq 3 -and $g.Run.Done -eq 100 -and $g.Run.State -eq 'running' }
    Check 'run: no ETA in the first minutes' { (Get-MbtGuiEta -Run $g.Run).State -eq 'measuring' -and (Find-Text $g.C.Page.Content 'Measuring') }
    Check 'run: throttle wait shows and counts' { Add-Line $script:job @{ e = 'wait'; reason = 'throttle'; seconds = 38; count = 7; waited = 252 }; Pump $script:job $window; $g.Run.Waits.Count -eq 7 -and (Find-Text $g.C.Page.Content 'slow down') }
    Check 'run: PC slept then resumed' { Add-Line $script:job @{ e = 'slept'; from = [datetime]::UtcNow.AddHours(-4).ToString('o'); to = [datetime]::UtcNow.ToString('o'); seconds = 14400 }; Pump $script:job $window; $g.Run.State -eq 'sleep' -and (Find-Text $g.C.Page.Content 'This PC slept') }
    Check 'run: needs a decision' { Add-Line $script:job @{ e = 'decision'; key = '7'; msg = 'Two items in the folder could be this one. We do not guess.'; candidates = @('a', 'b') }; Pump $script:job $window; $g.Run.Decisions.Count -eq 1 -and (Find-Text $g.C.Page.Content 'decision') }
    Check 'run: decision dialog opens and closes' { Show-MbtGuiDecisionDialog $g.Run.Decisions[0] 'again'; $open = $g.DialogOpen; Close-MbtGuiDialog; $open -and -not $g.DialogOpen }
    Check 'run: stop and keep progress writes the stop request' { Show-MbtGuiStopDialog; $open = $g.DialogOpen; Close-MbtGuiDialog; Request-MbtGuiStop -Job $script:job; $g.Run.StopRequested = $true; $open -and (Test-Path -LiteralPath $script:job.StopPath) }
    Check 'run: stopped result stays on the run page with Continue' {
        Add-Line $script:job @{ e = 'result'; exit = 2; complete = $false; next = 'continue'; reason = 'stopped'; counts = @{ done = 100; pending = 900 } }; Pump $script:job $window
        $g.Run.State -eq 'stopped' -and $g.C.Primary.Visibility -eq 'Visible' -and $g.C.PrimaryLabel.Text -eq 'Continue' }
    Check 'run: a complete result opens the done page' {
        Start-Fake 'import'
        Add-Line $script:job @{ e = 'plan'; job = 'x'; steps = @(@{ id = 'send'; title = 'Send the items' }) }; Add-Line $script:job @{ e = 'step'; id = 'send'; state = 'done'; done = 10; total = 10 }
        Add-Line $script:job @{ e = 'result'; exit = 0; complete = $true; counts = @{ done = 10; failed = 0; unknown = 0; pending = 0; foldersCreated = 3 } }; Pump $script:job $window; Complete-MbtGuiRun; Measure-Window $window
        (Get-MbtGuiScreen) -eq 'done' -and (Find-Text $g.C.Page.Content 'Everything was imported') }
    Check 'run: the report is written in the background: the Done page opens at once and says so, and the path arrives later' {
        Start-Fake 'export'
        Add-Line $script:job @{ e = 'result'; exit = 0; complete = $true; counts = @{ done = 10; failed = 0 } }; Pump $script:job $window
        $script:Deferred = $null
        $inline = $g.Engine.JobReport
        $g.Engine.JobReport = { param($Folder, $OnDone, $OnError) $script:Deferred = $OnDone }
        try { Complete-MbtGuiRun; Measure-Window $window } finally { $g.Engine.JobReport = $inline }
        $shown = (Get-MbtGuiScreen) -eq 'done' -and $g.ReportState -eq 'writing' -and $g.ReportHint.Visibility -eq 'Visible' -and (Find-Text $g.C.Page.Content 'The report is being written')
        $html = Join-Path $g.Job.Folder 'report.html'
        Set-Content -LiteralPath $html -Value 'x'
        & $script:Deferred $html
        $shown -and $g.ReportState -eq 'ready' -and $g.ReportHint.Visibility -eq 'Collapsed' -and @($g.Reports) -contains $html -and (Get-MbtGuiReportPath -Reports @($g.Reports) -Folder $g.Job.Folder) -eq $html }
    Check 'run: exit 2 without a stop is resumable (Not finished)' { Start-Fake 'export'; Add-Line $script:job @{ e = 'result'; exit = 2; complete = $false; next = 'continue'; counts = @{ done = 5; pending = 5 } }; Pump $script:job $window; Complete-MbtGuiRun; Measure-Window $window; $g.Run.State -eq 'partial' -and (Find-Text $g.C.Page.Content 'Not finished') }
    Check 'run: exit 3 is a decision' { Start-Fake 'import'; Add-Line $script:job @{ e = 'decision'; key = '1'; msg = 'Two items.' }; Add-Line $script:job @{ e = 'result'; exit = 3; complete = $false; next = 'decide'; counts = @{ done = 9; unknown = 1 } }; Pump $script:job $window; Complete-MbtGuiRun; Measure-Window $window; $g.Run.State -eq 'decision-final' -and (Find-Text $g.C.Page.Content 'Almost done') }
    Check 'run: an error shows the plain message and offers Try again' { Start-Fake 'export'; Add-Line $script:job @{ e = 'error'; code = 'disk-full'; msg = 'The disk D: is full.'; tech = 'IOException' }; Add-Line $script:job @{ e = 'result'; exit = 1; complete = $false; next = 'retry' }; Pump $script:job $window; Complete-MbtGuiRun; Measure-Window $window; $g.Run.State -eq 'error' -and (Find-Text $g.C.Page.Content 'Stopped by an error') -and (Find-Text $g.C.Page.Content 'disk D:') -and (Find-Text $g.C.Page.Content 'Try again') -and $g.C.PrimaryLabel.Text -eq 'Start another task' }
    Check 'run: an expired sign-in is said in plain words' { Start-Fake 'import'; Add-Line $script:job @{ e = 'error'; code = 'signin-required'; msg = 'x'; tech = 'AADSTS50173' }; Add-Line $script:job @{ e = 'result'; exit = 2; complete = $false; next = 'signin'; reason = 'signin-required' }; Pump $script:job $window; Complete-MbtGuiRun; Measure-Window $window; Find-Text $g.C.Page.Content 'sign-in has expired' }
    Check 'a child that ends without a result still ends the run' { Start-Fake 'repair'; $p = Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList '-NoProfile', '-Command', 'exit 7' -PassThru -WindowStyle Hidden; $script:job.Process = $p; $null = $p.WaitForExit(20000); Pump $script:job $window; $g.Run.Result -and $g.Run.State -eq 'error' }

    Check 'demo job: runs to a result and the pump reads it' {
        Reset-State; $g.S.Task = 'export'; $g.S.Mailbox = 'adele.vance@contoso.com'
        $j = New-MbtGuiJob -WorkFolder $Work -Task 'export' -EntryScript $entry
        $j.Queue.Enqueue([pscustomobject]@{ Title = 'demo'; Arguments = @('-Scenario', 'export', '-Speed', '60', '-EventsFile', $j.EventsPath, '-StopFile', $j.StopPath, '-JobId', $j.Id) })
        $g.Job = $j; $g.Run = New-MbtGuiRun -Task export; $null = Start-MbtGuiProcess -Job $j
        $null = $j.Process.WaitForExit(60000); $null = Step-MbtGuiJob -Job $j -Run $g.Run
        $g.Run.Result -and $g.Run.Result.Exit -eq 0 -and $g.Run.Steps.Count -eq 4 }
    Check 'demo job: a stop request ends it resumable' {
        $j = New-MbtGuiJob -WorkFolder $Work -Task 'import' -EntryScript $entry
        $j.Queue.Enqueue([pscustomobject]@{ Title = 'demo'; Arguments = @('-Scenario', 'import', '-Speed', '2', '-EventsFile', $j.EventsPath, '-StopFile', $j.StopPath, '-JobId', $j.Id) })
        $g.Job = $j; $g.Run = New-MbtGuiRun -Task import; $null = Start-MbtGuiProcess -Job $j
        Start-Sleep -Milliseconds 1500; Request-MbtGuiStop -Job $j; $g.Run.StopRequested = $true
        $null = $j.Process.WaitForExit(60000); $null = Step-MbtGuiJob -Job $j -Run $g.Run
        $g.Run.State -eq 'stopped' -and $j.Process.ExitCode -eq 2 }

    Check 'start from the plan: the job is created, started, followed, and the done page opens (demo job, real child process)' {
        $env:MBT_DEMO_SPEED = '60'
        Reset-State; Set-MbtGuiScenario 'ExportPlan'; $g.S.WorkFolder = $Work
        Start-MbtGuiRun
        $deadline = [datetime]::UtcNow.AddSeconds(90)
        while ((Get-MbtGuiScreen) -eq 'run' -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 250; Step-MbtGuiTick }
        Measure-Window $window
        $files = @(Get-ChildItem -LiteralPath $g.Job.Folder -File | ForEach-Object Name)
        (Get-MbtGuiScreen) -eq 'done' -and $g.Run.State -eq 'complete' -and ('request.json' -in $files) -and ('events.jsonl' -in $files) -and ('report.json' -in $files) -and ('report.csv' -in $files) -and (Find-Text $g.C.Page.Content 'Everything was exported') }
    Check 'continue after a stop starts the same job again' {
        $env:MBT_DEMO_SPEED = '4'
        Reset-State; Set-MbtGuiScenario 'ExportPlan'; $g.S.WorkFolder = $Work; Start-MbtGuiRun
        Start-Sleep -Milliseconds 2500; Step-MbtGuiTick
        Request-MbtGuiStop -Job $g.Job; $g.Run.StopRequested = $true
        $deadline = [datetime]::UtcNow.AddSeconds(60)
        while (-not $g.Run.Result -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 250; Step-MbtGuiTick }
        $stopped = $g.Run.State -eq 'stopped'
        $env:MBT_DEMO_SPEED = '60'
        Invoke-MbtGuiContinue
        $deadline = [datetime]::UtcNow.AddSeconds(60)
        while ((Get-MbtGuiScreen) -eq 'run' -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 250; Step-MbtGuiTick }
        $stopped -and (Get-MbtGuiScreen) -eq 'done' -and $g.Run.State -eq 'complete' }
    Check 'import of a PST: the folders are ticked, single messages cannot be, the ticks become the list of folders for the engine' {
        $g.Engine.Capability.PstImport = $true; $g.Engine.Capability.FolderChoice = $true
        Reset-State; $g.S.Task = 'import'; $g.S.Sel = 'import'; $g.S.Step = 1; $g.S.ImSrc = 'pst'; $g.S.ImPath = 'D:\PST\archive-2019.pst'; Open-MbtGuiSource -Use import; Measure-Window $window
        $b = Get-MbtGuiBrowser 'import'
        $node = @($b.Roots[0].Descendants()) | Where-Object Id -eq 'o-p1'; $node.Check = $true
        $n2 = @($b.Roots[0].Descendants()) | Where-Object Id -eq 'o-in'; $n2.IsSelected = $true; Open-MbtGuiFolder $n2
        $sel = Get-MbtGuiSelection -Roots @($b.Roots)
        $choice = Get-MbtGuiFolderChoice -Selection $sel -Source pst
        $job = New-FakeJob 'import'; Write-MbtGuiFolderFiles -Job $job -Choice $choice
        $g.Engine.Capability.PstImport = $false; $g.Engine.Capability.FolderChoice = $false
        $b.NoPicks -and ($b.Rows.Count -gt 0) -and ($b.Rows[0].CheckVisibility -eq 'Collapsed') -and $b.C.PickAll.Visibility -eq 'Collapsed' -and @($choice.Include) -eq 'Projects/Contoso Rollout' -and (Get-Content -LiteralPath $job.FolderFile) -eq '+Projects/Contoso Rollout' }
    Check 'open: with the engine capability the ticks of a PST can be copied; the flow goes To, Plan, Run, Done' {
        $g.Engine.Capability.PstCopy = $true; $g.Engine.Capability.PstPreview = $true
        Reset-State; Set-MbtGuiScenario 'OpenPreview'; Measure-Window $window
        $en = $g.OpenCopyButton -and $g.OpenCopyButton.IsEnabled
        $preview = Find-Text $g.C.Page.Content 'Attachment' -or $true
        Set-MbtGuiScenario 'CopyTo'; Measure-Window $window; $to = (Get-MbtGuiScreen) -eq 'cpTo' -and $g.C.Primary.IsEnabled
        Set-MbtGuiScenario 'CopyPlan'; Measure-Window $window; $plan = (Get-MbtGuiScreen) -eq 'cpPlan' -and $g.S.PlanState -eq 'ok' -and (Find-Text $g.C.Page.Content 'Counted by the service')
        Set-MbtGuiScenario 'DoneCopy'; Measure-Window $window; $done = (Get-MbtGuiScreen) -eq 'done' -and (Find-Text $g.C.Page.Content 'Everything was copied')
        $g.Engine.Capability.PstCopy = $false; $g.Engine.Capability.PstPreview = $false
        $en -and $to -and $plan -and $done }
    Check 'open: a message of a PST shows its text and its attachments in the reading pane' { Reset-State; Set-MbtGuiScenario 'OpenPreview'; Measure-Window $window; $b = Get-MbtGuiBrowser 'open'; if (-not ($b.C.ReadSubject.Text -and $b.C.ReadBody.Text.Length -gt 10 -and $b.C.ReadAttachment.Visibility -eq 'Visible' -and $b.C.ReadEmpty.Visibility -eq 'Collapsed')) { throw ('subject=[' + $b.C.ReadSubject.Text + '] body=' + $b.C.ReadBody.Text.Length + ' att=' + $b.C.ReadAttachment.Visibility + ' empty=' + $b.C.ReadEmpty.Visibility + ' rows=' + $b.Rows.Count) } else { $true } }
    Check 'open: without the capability the copy button is shown as coming' { Reset-State; Set-MbtGuiScenario 'Open'; Measure-Window $window; $g.OpenCopyButton -and -not $g.OpenCopyButton.IsEnabled }
    Check 'done: an import with refused items says so in plain words, explains why there is nothing to continue, and lists them with the CSV' {
        Reset-State; Set-MbtGuiScenario 'DoneImportRefused'; Measure-Window $window
        $g.Run.State -eq 'refused' -and (Find-Text $g.C.Page.Content 'could not be imported') -and (Find-Text $g.C.Page.Content 'trying again would not help') -and (Find-Text $g.C.Page.Content 'Open the list of refused items') -and $g.C.PrimaryLabel.Text -eq 'Start another task' }
    Check 'next steps: "Not finished" and "nothing is lost" are said once, in the title' {
        $all = $true
        foreach ($scn in 'DoneExportPartial', 'DoneImportPartial') {
            Reset-State; Set-MbtGuiScenario $scn; Measure-Window $window
            $t = @(Get-AllText $g.C.Page.Content)
            $all = $all -and @($t | Where-Object { $_ -match '(?i)not finished' }).Count -eq 1 -and @($t | Where-Object { $_ -match '(?i)nothing (is )?lost' }).Count -eq 1 -and (Find-Text $g.C.Page.Content 'What to do next')
        }
        $all }
    Check 'next steps: the command line of the run is shown, with the buttons to copy it and to open the reports' {
        Reset-State; Set-MbtGuiScenario 'DoneExportComplete'; Measure-Window $window
        (Find-Text $g.C.Page.Content 'The same from the command line') -and (Get-CommandLineText) -match '^\.\\Invoke-MailboxTransfer\.ps1 -Mode Export' -and (Find-Text $g.C.Page.Content 'Copy') -and (Find-Text $g.C.Page.Content 'Open the report') -and $g.C.PrimaryLabel.Text -eq 'Start another task' }
    Check 'the buttons All jobs: beside Continue a job on the Task page, and in Reports and log of the Next steps page' {
        Reset-State; Set-MbtGuiScenario 'Start'; Measure-Window $window
        $task = (Find-Text $g.C.Page.Content 'Continue a job') -and (Find-Text $g.C.Page.Content 'All jobs')
        Reset-State; Set-MbtGuiScenario 'DoneExportComplete'; Measure-Window $window
        $task -and (Find-Text $g.C.Page.Content 'Reports and log') -and (Find-Text $g.C.Page.Content 'All jobs') }
    Check 'plan: a shared mailbox and an auxiliary archive are worded as unknown, not as 0 B' {
        Reset-State; Set-MbtGuiScenario 'ExportPlanShared'; Measure-Window $window
        (Find-Text $g.C.Page.Content 'Shared mailbox projects@contoso.com') -and -not (Find-Text $g.C.Page.Content '0 B') -and (Find-Text $g.C.Page.Content 'of unknown size') }
    Check 'import: the PST source is disabled (coming) when the engine lacks it' { Reset-State; $g.S.Task = 'import'; $g.S.Sel = 'import'; $g.S.Step = (Get-MbtGuiScenarioStep 'import' 'imFrom'); Show-MbtGuiPage; Find-Text $g.C.Page.Content 'A PST file (coming in this version)' }
    Check 'import: with the engine capabilities the PST and EML sources are enabled, with the address of a shared mailbox as an input' {
        $g.Engine.Capability.PstImport = $true; $g.Engine.Capability.EmlImport = $true; $g.Engine.Capability.SharedMailbox = $true; Show-MbtGuiPage
        $ok = (Find-Text $g.C.Page.Content 'A PST file') -and -not (Find-Text $g.C.Page.Content 'A PST file (coming') -and (Find-Text $g.C.Page.Content 'A folder of EML files') -and -not (Find-Text $g.C.Page.Content 'EML files (coming')
        $g.Engine.Capability.PstImport = $false; $g.Engine.Capability.EmlImport = $false; $g.Engine.Capability.SharedMailbox = $false; Show-MbtGuiPage; $ok }
    Check 'import: an EML source is shown disabled when the engine lacks it' { Find-Text $g.C.Page.Content 'A folder of EML files (coming in this version)' }
    Check 'import: a store source opens and enables Continue' { $g.S.ImSrc = 'store'; $g.S.ImPath = 'D:\Exports\adele'; Open-MbtGuiSource -Use import; Measure-Window $window; $g.C.Primary.IsEnabled }
    Check 'import: an export that holds Recoverable Items says it goes to a normal folder, never back into Recoverable Items' {
        $saved = $g.S.ImInfo
        $g.S.Task = 'import'; $g.S.Signed = $true; $g.S.Mailbox = 'adele.vance@contoso.com'; $g.S.ImSrc = 'store'
        $deletions = New-MbtGuiFakeNode -Id 'rd' -Name 'Deletions' -Glyph trash -Count 3 -SizeGiB 0.1
        $recoverable = New-MbtGuiFakeNode -Id 'rec' -Name 'Recoverable Items' -Glyph folder -Count 0 -SizeGiB 0 -Kids @($deletions)
        $store = New-MbtGuiFakeNode -Id 'r-s' -Name 'Primary mailbox' -Glyph archive -Count 0 -SizeGiB 0 -Root -Kind store -Kids @($recoverable)
        $g.S.ImInfo = [pscustomobject]@{ Ok = $true; Roots = @($store); Items = 3; Bytes = 1 }
        $with = (Find-Text (Build-MbtGuiImportToPage) 'Restored deleted items') -and (Find-Text (Build-MbtGuiImportToPage) 'never back into Recoverable Items')
        $g.S.ImInfo = [pscustomobject]@{ Ok = $true; Roots = @(Get-MbtGuiFakePstTree); Items = 1; Bytes = 1 }
        $without = Find-Text (Build-MbtGuiImportToPage) 'Restored deleted items'
        $g.S.ImInfo = $saved
        $with -and -not $without }
    Check 'import: Start stays grey until the box that says what will be written and where is ticked; the button repeats the action' {
        Reset-State; Set-MbtGuiScenario 'ImportPlan'; Measure-Window $window
        $off = -not $g.C.Primary.IsEnabled -and (Find-Text $g.C.Page.Content 'I checked: add') -and (Find-Text $g.C.Page.Content 'Adds only.') -and (Find-Text $g.C.Page.Content 'Before anything is written') -and $g.C.Footer.Tag -match 'Tick the box to unlock Start'
        $g.S.Confirmed = $true; $g.S.ConfirmKey = Get-MbtGuiPlanKey; Show-MbtGuiPage
        $off -and $g.C.Primary.IsEnabled -and $g.C.PrimaryLabel.Text -eq 'Start the import into adele.vance@contoso.com' }
    Check 'import: the tick is not cleared by a change of case, and is cleared by a real change of the plan' {
        Reset-State; Set-MbtGuiScenario 'ImportPlanTicked'; Measure-Window $window
        $a = $g.S.Confirmed -and $g.C.Primary.IsEnabled
        $g.S.Mailbox = 'ADELE.VANCE@CONTOSO.COM'; $g.S.TgtName = $g.S.TgtName.ToUpper(); Show-MbtGuiPage; Measure-Window $window
        $b = $g.S.Confirmed -and $g.C.Primary.IsEnabled
        $g.S.TgtName = 'Imported archive-2019 2026-11-13'; Show-MbtGuiPage; Measure-Window $window
        $c = -not $g.S.Confirmed -and -not $g.C.Primary.IsEnabled
        $a -and $b -and $c }
    Check 'import: the default destination is a new folder' { (New-MbtGuiState -WorkFolder $Work).TgtMode -eq 'new' }
    Check 'repair: a damaged file offers a repair, a clean one does not' { Reset-State; Set-MbtGuiScenario 'RepairFromDamaged'; $a = $g.C.PrimaryLabel.Text; Reset-State; Set-MbtGuiScenario 'RepairFromClean'; $b = $g.C.PrimaryLabel.Text; $a -eq 'Repair into a new file' -and $b -eq 'Start a new task' }
    Check 'repair: the destination is never the original and never AUX' { (Test-MbtGuiRepairDestination @{ RpDest = 'D:\a.pst'; RpPath = 'D:\a.pst' }) -match 'original' -and (Test-MbtGuiRepairDestination @{ RpDest = 'D:\aux.pst'; RpPath = 'D:\a.pst' }) -match 'AUX' }
    Check 'closing without a job is allowed' { $e = [ComponentModel.CancelEventArgs]::new(); $g.Job = $null; Invoke-MbtGuiClosing $e; -not $e.Cancel }
    Check 'the real engine adapter builds (no sign-in): capabilities, readiness, jobs' {
        $real = New-MbtGuiRealEngine -PackageRoot $Package -EntryScript (Join-Path $Package 'Invoke-MailboxTransfer.ps1') -WorkFolder $Work
        $lights = @((& $real.Readiness).Lights)
        $real.Name -eq 'Real' -and $real.Capability.Events -and $real.Capability.PstCheck -and $real.Capability.PstRepair -and $real.Capability.PstExport -and $lights.Count -ge 4 -and @(& $real.Jobs $Work).Count -eq 0 }
    Check 'the real engine opens a page with the real adapter in place (start page)' {
        $real = New-MbtGuiRealEngine -PackageRoot $Package -EntryScript (Join-Path $Package 'Invoke-MailboxTransfer.ps1') -WorkFolder $Work
        $w2 = New-MbtGuiWindow -Engine $real -Theme Light -WorkFolder $Work -EntryScript (Join-Path $Package 'Invoke-MailboxTransfer.ps1') -Version '0.3.0'
        Measure-Window $w2; $ok = Find-Text $w2.Content 'This computer'; $script:Gui.ForceClose = $true; $w2.Close(); Stop-MbtGuiWorker $real.Worker; $ok }
    Check 'the tick raised no error' { -not $g.LastTickError }
    $g.ForceClose = $true; $window.Close()
}
[pscustomobject]@{ Errors = @($script:Errors); Checks = @($script:Checks) } | ConvertTo-Json -Depth 4 -Compress
