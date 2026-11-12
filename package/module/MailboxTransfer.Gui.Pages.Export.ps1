#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - sign-in card and the pages of Browse and export (dot-sourced by Start-MailboxTransferConsole.ps1).

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

#region Plain-words errors ---------------------------------------------------------

function ConvertTo-MbtGuiMessage {
    <# What happened, in one sentence, from a raw engine message. Returns Text and Tech (the raw message, for the collapsed technical line). #>
    param([string]$Raw)
    $r = [string]$Raw
    $text = switch -Regex ($r) {
        'AADSTS50011|AADSTS500113|does not accept the Windows sign-in window' {
            $u = [regex]::Match($r, 'ms-appx-web://microsoft\.aad\.brokerplugin/[0-9A-Za-z-]+').Value
            if (-not $u) { $u = 'ms-appx-web://microsoft.aad.brokerplugin/<application id>' }
            "Microsoft does not accept the sign-in window for this application yet. Ask your administrator to add this redirect URI to the app registration (Authentication, Mobile and desktop applications): $u. Until then, use a code instead."; break
        }
        'sign-in window was closed before' { 'The sign-in window was closed before the sign-in was finished. Press Sign in to try again, or use a code instead.'; break }
        'sign-in window did not finish within' { 'The sign-in window did not finish in time. Press Sign in to try again, or use a code instead.'; break }
        'AADSTS50173|AADSTS700082|grant has expired|signin-required' { 'Your sign-in has expired. Sign in again to continue. Nothing is lost.'; break }
        'device.?code.*expired|AADSTS70019|code has expired' { 'The code has expired. Get a new code.'; break }
        'No certificate with the thumbprint|has no private key|certificate .* expired|is not valid before|Cannot open the certificate' { 'The certificate cannot be used. Check the thumbprint, that the certificate is installed on this PC with its private key, and that it has not expired.'; break }
        'AADSTS700027|certificate was refused|AADSTS700016|AADSTS90002|not found in the tenant' { 'Microsoft does not accept this application or this certificate. Check the tenant and the application id, and that the certificate was added to the application.'; break }
        'Cannot find the mailbox id|FullAccess|Full Access' { 'We cannot open this mailbox. Check the address, and that you have the Full Access right on it (a new right can take up to an hour to apply).'; break }
        'MailboxNotFound|mailbox.*not found|ResourceNotFound|ErrorNonExistentMailbox' { 'We cannot find this mailbox. Check the address. Mailboxes that are not in Exchange Online cannot be used.'; break }
        '403|ErrorAccessDenied|Access is denied|Forbidden' { 'This sign-in is not allowed to do this. Ask your administrator for the permission (MailboxItem.ImportExport for an import, MailboxItem.Read and MailboxFolder.Read to browse and export).'; break }
        'not enough space|disk.*full|IOException.*space' { 'The disk is full. Free some space and press Try again. What was downloaded is kept.'; break }
        'already exists' { 'This file already exists. The tool never overwrites a file. Choose another name.'; break }
        'cannot reach|NameResolution|No such host|network|SocketException' { 'We cannot reach Microsoft. Trying again every 30 seconds. Check your internet connection.'; break }
        default { 'Something went wrong. The technical message below says what.' }
    }
    [pscustomobject]@{ Text = $text; Tech = $r }
}

#endregion

#region Sign-in -----------------------------------------------------------------------------------------

function Get-MbtGuiDeviceCodeFile {
    $g = $script:Gui
    $root = if ($g.S.WorkFolder) { $g.S.WorkFolder } else { $env:TEMP }
    return (Join-Path $root 'device-code.txt')
}

function Get-MbtGuiSignInSettings {
    <# The tenant and the application for a sign-in as the user: the configuration, or what was typed on the Setup page. Asks = something is missing. #>
    $e = Get-MbtGuiEffectiveSignIn -S $script:Gui.S
    [pscustomobject]@{ TenantId = $e.Tenant; ClientId = $e.Client; Asks = [bool]($e.Missing.Count -or $e.TenantChanged -or $e.ClientChanged) }
}

function Test-MbtGuiAppSignIn {
    <# The three values of "the application" (tenant, application id, certificate thumbprint): the message in plain words, or $null when they can be used. #>
    param([hashtable]$S)
    $e = Get-MbtGuiEffectiveSignIn -S $S
    if (-not $e.Tenant) { return 'The tenant is missing. Open Setup and type it: its id (a guid) or its domain, like contoso.onmicrosoft.com.' }
    if ($e.Tenant -in 'common', 'organizations', 'consumers') { return 'Type the id or the domain of your tenant, not a general name (Setup).' }
    if ($e.Client -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { return 'The application (client) id is missing or is not a guid. Open Setup and type it, like 00000000-0000-0000-0000-000000000000.' }
    if (($e.Thumb -replace '[^0-9A-Fa-f]', '').Length -ne 40) { return 'The thumbprint of the certificate is missing or is not 40 characters (letters a-f and digits). Open Setup and type it; it is shown in the certificate details.' }
    return $null
}

function Get-MbtGuiWindowHandle {
    <# The handle of the console window: the Windows sign-in window opens over it. 0 when the window has none yet (the engine then uses the console window of the process). #>
    $w = Get-MbtGuiProp $script:Gui 'Window'
    if (-not $w) { return [long]0 }
    try { return [long]([Windows.Interop.WindowInteropHelper]::new($w).Handle) } catch { return [long]0 }
}

function Start-MbtGuiSignIn {
    param([bool]$UseCode = $false)
    $g = $script:Gui; $s = $g.S
    if ($UseCode -and $s.SigningIn -and $s.SignInCts) {
        # the sign-in window is still waited for: it is ended, and the code is asked for as soon as the worker is free
        $s.SwitchToCode = $true
        try { $s.SignInCts.Cancel() } catch { Write-Verbose "cancel: $_" }
        $s.CodeShown = $true; $s.Code = $null
        Show-MbtGuiPage
        return
    }
    $addr = $s.Mailbox.Trim()
    if ($addr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $s.SignInError = 'Type the address of the mailbox, like adele.vance@contoso.com.'
        Show-MbtGuiPage; return
    }
    # "the application": tenant, application (client) id and the thumbprint of its certificate, from the configuration or typed on the Setup page for this run only (nothing is saved)
    $appArg = $null
    $settings = $null
    if ($s.AuthMode -eq 'app' -and (Get-MbtGuiCapability 'AppOnly')) {
        $UseCode = $false
        $bad = Test-MbtGuiAppSignIn -S $s
        if ($bad) { $s.SignInError = $bad; Show-MbtGuiPage; return }
        $eff = Get-MbtGuiEffectiveSignIn -S $s
        $appArg = @{ Tenant = $eff.Tenant; ClientId = $eff.Client; Thumbprint = ($eff.Thumb -replace '[^0-9A-Fa-f]', '') }
    } else {
        $settings = Get-MbtGuiSignInSettings
        if ($settings.Asks) {
            $guid = '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$'
            if ($settings.TenantId -notmatch $guid -and $settings.TenantId -notmatch '^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') { $s.SignInError = 'The tenant is missing. Open Setup and type it: its directory id (a GUID) or its domain name, like contoso.onmicrosoft.com.'; $s.SignInTech = ''; Show-MbtGuiPage; return }
            if ($settings.ClientId -notmatch $guid) { $s.SignInError = 'The application (client) id is missing. Open Setup and type the id of the app registration: a GUID.'; $s.SignInTech = ''; Show-MbtGuiPage; return }
        }
    }
    # the Windows sign-in window needs a component of Microsoft.Graph.Authentication: without it the sign-in uses a code, and says so
    $s.SignInNote = ''
    if (-not $appArg -and -not $UseCode) {
        $probe = Get-MbtGuiProp $g.Engine 'WindowSignIn'
        $win = $null; if ($probe) { try { $win = & $probe } catch { $win = $null } }
        if ($win -and -not $win.Available) { $UseCode = $true; $s.SignInNote = "The Windows sign-in window cannot be used on this PC ($($win.Reason)). A code is used instead." }
    }
    $s.Mailbox = $addr; $s.SigningIn = $true; $s.SignInError = ''; $s.CodeShown = $UseCode; $s.Code = $null; $s.SwitchToCode = $false
    $s.SignInCts = [Threading.CancellationTokenSource]::new()
    $codeFile = Get-MbtGuiDeviceCodeFile
    if (Test-Path -LiteralPath $codeFile) { Remove-Item -LiteralPath $codeFile -Force -ErrorAction SilentlyContinue }
    Show-MbtGuiPage
    & $g.Engine.SignIn $addr $UseCode $codeFile {
        param($info)
        $s.SigningIn = $false; $s.Signed = $true; $s.Token = $info; $s.Account = $info.Account; $s.CodeShown = $false; $s.SignInCts = $null
        Read-MbtGuiLocations
        Show-MbtGuiPage
    }.GetNewClosure() {
        param($msg)
        $s.SigningIn = $false; $s.Signed = $false; $s.CodeShown = $false; $s.Code = $null; $s.SignInCts = $null
        if ($s.SwitchToCode) { $s.SwitchToCode = $false; Start-MbtGuiSignIn -UseCode $true; return }
        $m = ConvertTo-MbtGuiMessage $msg
        $s.SignInError = $m.Text; $s.SignInTech = $m.Tech
        Show-MbtGuiPage
    }.GetNewClosure() $false $(if ($settings) { $settings.TenantId } else { '' }) $(if ($settings) { $settings.ClientId } else { '' }) $appArg (Get-MbtGuiWindowHandle) $(if ($s.WorkFolder) { Join-Path $s.WorkFolder 'logs' } else { '' }) $s.SignInCts
}

function Step-MbtGuiSignInCode {
    <# While a device-code sign-in waits, the engine writes "DEVICE CODE: <code> URL: <url>" to a file: show it in the card. #>
    $g = $script:Gui
    if (-not $g -or -not $g.S.SigningIn -or $g.S.Code) { return }
    $f = Get-MbtGuiDeviceCodeFile
    if (-not (Test-Path -LiteralPath $f)) { return }
    $line = (Get-Content -LiteralPath $f -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($line -match 'DEVICE CODE:\s*(\S+)\s+URL:\s*(\S+)') {
        $g.S.CodeShown = $true
        $g.S.Code = @{ Code = $Matches[1]; Url = $Matches[2]; Until = (Get-Date).AddMinutes(15) }
        Show-MbtGuiPage
    }
}

function Read-MbtGuiLocations {
    $g = $script:Gui; $s = $g.S
    $s.Reading = $true
    & $g.Engine.Locations $s.Mailbox $s.LocationsFile {
        param($locs)
        $s.Locations = @($locs); $s.Reading = $false
        foreach ($l in $s.Locations) { if (-not $s.LocationChecked.ContainsKey($l.Kind)) { $s.LocationChecked[$l.Kind] = $true } }
        Sync-MbtGuiLocationRoots
    }.GetNewClosure() {
        param($msg) $s.Reading = $false; $s.Locations = @(); $m = ConvertTo-MbtGuiMessage $msg; $s.SignInError = $m.Text; $s.SignInTech = $m.Tech; $s.Shared = ''; Show-MbtGuiPage
    }.GetNewClosure() $s.Shared
}

function Sync-MbtGuiLocationRoots {
    <# The browser shows the locations that are ticked in "Where to look". #>
    $g = $script:Gui; $s = $g.S
    $browser = Get-MbtGuiBrowser 'export'
    $want = @($s.Locations | Where-Object { $s.LocationChecked[$_.Kind] })
    & $g.Engine.Roots $want {
        param($roots)
        Set-MbtGuiBrowserRoots -Browser $browser -Roots @($roots)
        if ((Get-MbtGuiScreen) -eq 'exFrom') { Update-MbtGuiExportFromFooter }
    }.GetNewClosure() { param($msg) $browser.C.ListEmpty.Text = "The folders could not be read: $msg" }.GetNewClosure() ([bool]$s.RecoverableItems)
}

function New-MbtGuiSignInCard {
    <# Sign in (not yet), or Mailbox (signed in). Wording is generic on purpose: who may sign in for another mailbox depends on the rights the account has. #>
    param([string]$Use = 'export')
    $g = $script:Gui; $s = $g.S
    if ($s.Signed) {
        $box = New-GuiTextBox -Text $s.Mailbox -ReadOnly -Automation 'Mailbox address'
        $out = New-GuiButton -Text 'Sign out' -OnClick { & $g.Engine.SignOut; $s.Signed = $false; $s.Shared = ''; $s.ImShared = ''; $s.Locations = @(); $s.Code = $null; if ($g.Browsers.ContainsKey('export')) { Set-MbtGuiBrowserRoots -Browser $g.Browsers['export'] -Roots @() }; Show-MbtGuiPage }
        $out.Margin = [Windows.Thickness]::new(0)
        $row = New-GuiRow -Box $box -Button $out
        $tok = $s.Token
        $gaps = @(Get-MbtGuiProp $tok 'Missing' @())
        $missing = if ($gaps.Count) { ($gaps -join ', ') } else { 'Nothing' }
        $facts = New-GuiFacts @(@('Access', $(if ($tok) { [string](Get-MbtGuiProp $tok 'Access' 'Delegated (you, with your rights)') } else { 'Delegated (you, with your rights)' })), @('Can read', $(if ($tok -and (Get-MbtGuiProp $tok 'CanReadMail' $false)) { 'Mail, folders, items (live view)' } else { 'Folders and items (listing without subjects)' })), @('Missing', $(if ($missing -eq 'Nothing') { New-GuiText 'Nothing' 'Muted' } else { $missing })))
        $more = New-GuiMore -Summary 'Sign-in details' -Content @($facts)
        $msgs = @((New-GuiMessage -Kind ok -Icon check -Text $(if ($s.AuthMode -eq 'app') { "Connected as the application (chosen in Setup), working on **$($s.Account)**. Nothing is kept on this PC." } else { "Signed in as **$($s.Account)** (delegated, chosen in Setup). The sign-in is kept on this PC, encrypted for your Windows account." })))
        if ($tok -and -not $tok.CanReadMail) { $msgs += New-GuiMessage -Kind warn -Icon warn -Text 'Subjects need the Mail.Read permission. You can still export.' }
        return (New-GuiCard -Title 'Mailbox' -Right (New-GuiPill $(if ($s.AuthMode -eq 'app') { 'Connected' } else { 'Signed in' }) 'ok') -Content (@((New-GuiText 'Mailbox address' 'Label'), $row) + $msgs + @($more)))
    }
    $app = ($s.AuthMode -eq 'app') -and (Get-MbtGuiCapability 'AppOnly')
    $box = New-GuiTextBox -Text $s.Mailbox -Automation 'Mailbox address' -OnChange { param($x, $e) $s.Mailbox = $x.Text }
    $eff = Get-MbtGuiEffectiveSignIn -S $s
    $content = @()
    $content += New-GuiText $(if ($app) { 'Mailbox the application works on' } else { 'Mailbox to look at' }) 'Label'
    $content += $box
    $content += New-GuiText $(if ($app) { 'Required: the application has no mailbox of its own to fall back on.' } else { 'Use an account that is allowed to read this mailbox. A Microsoft sign-in window opens over the console. Type the account that may open the mailbox.' }) 'Hint'
    $change = New-GuiLink 'Change in Setup' { Open-MbtGuiSetup }; $change.FontSize = 12; $change.HorizontalAlignment = 'Left'; $change.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
    $facts = New-GuiFacts @(@('Sign in as', $(if ($app) { 'The application (certificate)' } else { 'Me (delegated)' })), @('Tenant', $(if ($eff.Tenant) { $eff.Tenant } else { 'not set (open Setup)' })))
    $content += $facts
    $content += $change
    if ($eff.Missing.Count) { $content += New-GuiMessage -Kind warn -Icon warn -Text ("**Not set: " + ($eff.Missing -join ', ') + '.** Open Setup and type the missing value for this run, or put it in the configuration file.') }
    $btnText = if ($s.SigningIn -and -not $s.CodeShown) { if ($app) { 'Connecting...' } else { 'Waiting for the sign-in window...' } } else { if ($app) { 'Connect' } else { 'Sign in...' } }
    $signin = New-GuiButton -Text $btnText -Accent -Icon key -Disabled:($s.SigningIn) -OnClick { Start-MbtGuiSignIn -UseCode $false }
    $signin.Margin = [Windows.Thickness]::new(0, 12, 8, 0); $signin.HorizontalAlignment = 'Left'
    $content += $signin
    if (-not $app) {
        $first = New-GuiText 'First sign-in of an account Windows does not know? If Windows asks to stay signed in to all apps, choose "No, sign in to this app only".' 'Hint'
        $first.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
        $content += $first
    }
    if ($s.ContainsKey('SignInNote') -and $s.SignInNote -and -not $app) { $content += New-GuiMessage -Kind info -Icon info -Text $s.SignInNote }
    if ($s.SignInError) {
        $tech = if ($s.ContainsKey('SignInTech') -and $s.SignInTech) { New-GuiMore -Summary 'Technical message' -Content @((New-GuiText $s.SignInTech 'Small')) } else { $null }
        $content += New-GuiMessage -Kind bad -Icon x -Text $s.SignInError -After @($tech)
    }
    if ($s.CodeShown) {
        $rec = [Windows.Controls.Border]::new(); $rec.BorderThickness = [Windows.Thickness]::new(1); $rec.CornerRadius = [Windows.CornerRadius]::new(8); $rec.Padding = [Windows.Thickness]::new(12, 10, 12, 10); $rec.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
        Set-MbtGuiResource $rec ([Windows.Controls.Border]::BorderBrushProperty) 'DividerStrokeColorDefaultBrush'; Set-MbtGuiResource $rec ([Windows.Controls.Border]::BackgroundProperty) 'LayerFillColorDefaultBrush'
        $head = New-GuiText 'Sign in with a code' 'Body'; $head.FontWeight = [Windows.FontWeights]::SemiBold
        $inner = @($head)
        if ($s.Code) {
            $inner += New-GuiText ("Open **$($s.Code.Url)** on any device and enter the code:") 'Body'
            $code = New-GuiText $s.Code.Code 'Body' -NoWrap; $code.FontSize = 22; $code.FontWeight = [Windows.FontWeights]::SemiBold; $code.Margin = [Windows.Thickness]::new(0, 4, 0, 4)
            $inner += $code
            $inner += New-GuiText ("This code is valid until " + $s.Code.Until.ToString('HH:mm') + '.') 'Small'
            $btns = [Windows.Controls.StackPanel]::new(); $btns.Orientation = 'Horizontal'; $btns.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
            $btns.Children.Add((New-GuiButton -Text 'Copy the code' -Small -OnClick { $null = Set-MbtGuiClipboard $s.Code.Code })) | Out-Null
            $btns.Children.Add((New-GuiButton -Text 'Open the page' -Small -OnClick { Start-Process $s.Code.Url })) | Out-Null
            $inner += $btns
        } else { $inner += New-GuiText 'Getting a code...' 'Hint' }
        $inner += New-GuiText 'Waiting for you to finish... the console continues by itself.' 'Hint'
        $rec.Child = (New-GuiStack $inner)
        $content += $rec
    } elseif (-not $app) {
        $line = [Windows.Controls.WrapPanel]::new(); $line.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
        $line.Children.Add((New-GuiText 'The window does not appear (look behind the console)? ' 'Small')) | Out-Null
        $line.Children[0].Margin = [Windows.Thickness]::new(0, 0, 4, 0)
        $lk = New-GuiLink 'Use a code instead' { Start-MbtGuiSignIn -UseCode $true }; $lk.FontSize = 12
        $line.Children.Add($lk) | Out-Null
        $content += $line
    }
    return (New-GuiCard -Title 'Sign in' -Content $content)
}

function New-MbtGuiSharedMailboxSection {
    <#
        A shared mailbox, or the mailbox of another user, given by its address. Use export: the browse looks into it (Read-MbtGuiLocations);
        Use import: the items go into it (-MailboxId of the engine). Without the capability: shown as coming.
    #>
    param([ValidateSet('export', 'import')][string]$Use)
    $g = $script:Gui; $s = $g.S
    $key = if ($Use -eq 'export') { 'Shared' } else { 'ImShared' }
    $on = Get-MbtGuiCapability 'SharedMailbox'
    $title = if ($on) { 'A shared mailbox, or another user''s mailbox' } else { 'A shared mailbox (coming in this version)' }
    if (-not $on) { return (New-GuiMore -Summary $title -Content @((New-GuiText 'This version of the engine cannot open a mailbox by its address.' 'Hint'))) }
    $hold = @{ Text = [string]$s[$key] }
    $box = New-GuiTextBox -Text $hold.Text -Automation 'Shared mailbox address' -OnChange { param($x, $e) $hold.Text = $x.Text }.GetNewClosure()
    $open = New-GuiButton -Text $(if ($Use -eq 'export') { 'Open it' } else { 'Use it' }) -Small -Icon key -OnClick {
        $addr = $hold.Text.Trim()
        if ($addr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { $s.SignInError = 'Type the address of the mailbox, like projects@contoso.com.'; $s.SignInTech = ''; Show-MbtGuiPage; return }
        $s.SignInError = ''
        $s[$key] = $addr
        if ($Use -eq 'export') {
            $g.ExportSelection = [pscustomobject]@{ Entries = @(); Items = 0L; Bytes = 0L; Folders = 0; UnknownSize = 0; SubtreeUnread = $false }
            if ($g.Browsers.ContainsKey('export')) { Set-MbtGuiBrowserRoots -Browser $g.Browsers['export'] -Roots @() }
            $s.Locations = @(); $s.LocationChecked = @{}
            Read-MbtGuiLocations
        } else { $s.TgtLoc = 'Primary'; $s.PlanState = 'idle' }
        Show-MbtGuiPage
    }.GetNewClosure()
    $open.Margin = [Windows.Thickness]::new(0, 6, 8, 0); $open.HorizontalAlignment = 'Left'
    $content = [Collections.Generic.List[object]]::new()
    $content.Add((New-GuiText 'Mailbox address' 'Label')); $content.Add($box); $content.Add($open)
    $content.Add((New-GuiText 'You need the **Full Access** right on that mailbox (a new right can take up to an hour to apply). The mailbox is only read unless you import into it.' 'Hint'))
    if ($s[$key]) {
        $back = New-GuiLink 'Back to my own mailbox' {
            $s[$key] = ''
            if ($Use -eq 'export') {
                $g.ExportSelection = [pscustomobject]@{ Entries = @(); Items = 0L; Bytes = 0L; Folders = 0; UnknownSize = 0; SubtreeUnread = $false }
                if ($g.Browsers.ContainsKey('export')) { Set-MbtGuiBrowserRoots -Browser $g.Browsers['export'] -Roots @() }
                $s.Locations = @(); $s.LocationChecked = @{}
                Read-MbtGuiLocations
            } else { $s.PlanState = 'idle' }
            Show-MbtGuiPage
        }.GetNewClosure()
        $content.Add((New-GuiMessage -Kind info -Icon info -Text "**Looking at $($s[$key]).** Signed in as $($s.Account)."))
        $content.Add($back)
    }
    return (New-GuiMore -Summary $title -Content $content.ToArray() -Open:([bool]$s[$key]))
}

function New-MbtGuiLocationsCard {
    $g = $script:Gui; $s = $g.S
    $items = @()
    foreach ($l in $s.Locations) {
        $kind = $l.Kind
        $opt = New-GuiOption -Type Check -Title $l.Label -Hint $l.Hint -Checked ([bool]$s.LocationChecked[$kind]) -OnChange { param($x, $e) $s.LocationChecked[$kind] = [bool]$x.IsChecked; Sync-MbtGuiLocationRoots }.GetNewClosure()
        $items += $opt
    }
    if (-not $items.Count) { $items += New-GuiText $(if ($s.Reading) { 'Reading the locations...' } else { 'No location found.' }) 'Hint' }
    else {
        $items += New-GuiOption -Type Check -Title 'Include deleted items (Recoverable Items)' -Margin '0,10,0,0' -Checked ([bool]$s.RecoverableItems) `
            -Hint 'Off: only the folders you see in Outlook. On: the browser also shows the hidden folder where deleted mail waits. Changing this reloads the browser.' `
            -OnChange { param($x, $e) $s.RecoverableItems = [bool]$x.IsChecked; Sync-MbtGuiLocationRoots }.GetNewClosure()
        $what = [Collections.Generic.List[object]]::new()
        foreach ($t in @(
                'When you empty **Deleted Items** in Outlook, the messages are not gone at once: Exchange keeps them for a while in a hidden folder (Outlook''s "Recover deleted items" looks there).'
                'The same hidden area holds items purged for good, older versions of items that were changed, and items kept by a hold.'
                'They are exported with their full content and attachments, and saved apart, under a folder named **Recoverable Items**, in the export and in the PST. They are never mixed with your normal folders.'
                'Calendar Logging and SubstrateHolds, which are Microsoft''s own copies, are not shown here.'
                'When you import them later they go into a normal folder named **Restored deleted items**, never back into Recoverable Items.')) {
            $r = New-GuiText ([char]0x2022 + '  ' + ($t -replace '\*\*', '')) 'Body'; $r.FontSize = 13; $r.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $what.Add($r)
        }
        $items += New-GuiMore -Summary 'What are deleted items (Recoverable Items)?' -Content $what.ToArray()
    }
    if (@($s.Locations | Where-Object { $_.PSObject.Properties['Filling'] -and $_.Filling }).Count) {
        $f = @($s.Locations | Where-Object { $_.PSObject.Properties['Filling'] -and $_.Filling })[0]
        $items += New-GuiMessage -Kind warn -Icon warn -Text "**$($f.Label) is still being filled by Microsoft.** Some folders may be `"not ready`". The tool tries them again later and lists them; nothing is exported as empty."
    }
    $pick = New-GuiButton -Text 'Load a locations file...' -Small -Icon open -OnClick {
        $f = Select-MbtGuiFile -Title 'Locations file (Get-MailboxLocation output)' -Filter 'Locations (*.csv;*.json;*.txt)|*.csv;*.json;*.txt|All files (*.*)|*.*'
        if ($f) { $s.LocationsFile = $f; Read-MbtGuiLocations; Show-MbtGuiPage }
    }
    $pick.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
    $items += New-GuiMore -Summary 'The archive is missing from this list' -Content @((New-GuiText 'Microsoft does not tell us every archive. Ask your administrator to run **Get-MailboxLocation** for this user and save the result, then load it here.' 'Hint'), $pick)
    $items += New-MbtGuiSharedMailboxSection -Use export
    return (New-GuiCard -Title 'Where to look' -Info "$(@($s.Locations).Count) locations found" -Content $items)
}

#endregion

#region Export: From -------------------------------------------------------------------------------------

function Update-MbtGuiExportFromFooter {
    $g = $script:Gui
    if ((Get-MbtGuiScreen) -ne 'exFrom') { return }
    $browser = Get-MbtGuiBrowser 'export'
    $sel = Get-MbtGuiSelection -Roots @($browser.Roots)
    Update-MbtGuiPrimary -Enabled ($g.S.Signed -and $sel.Items -gt 0) -Footer (Format-MbtGuiSelectionFooter $sel $g.S.Signed)
    $g.ExportSelection = $sel
}

function Build-MbtGuiExportFromPage {
    $g = $script:Gui; $s = $g.S
    $browser = Get-MbtGuiBrowser 'export'
    $sel = Get-MbtGuiSelection -Roots @($browser.Roots)
    $g.ExportSelection = $sel
    $left = New-GuiStack @((New-MbtGuiSignInCard 'export'))
    if ($s.Signed) { $left.Children.Add((New-MbtGuiLocationsCard)) | Out-Null; $left.Children.Add((New-MbtGuiSelectionCard $sel)) | Out-Null }
    else { $left.Children.Add((New-MbtGuiSelectionCard $sel)) | Out-Null }
    if ($s.Signed) {
        $right = [Windows.Controls.Grid]::new()
        Place-MbtGuiBrowser $browser $right
    } else {
        $ico = New-GuiIcon mail 22 'MbtBrandText'; $ico.HorizontalAlignment = 'Center'
        $t1 = New-GuiText 'Sign in to look inside the mailbox.' 'Muted'; $t1.TextAlignment = 'Center'; $t1.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
        $t2 = New-GuiText 'Nothing is read before you sign in.' 'Hint'; $t2.TextAlignment = 'Center'
        $body = New-GuiStack @($ico, $t1, $t2); $body.Margin = [Windows.Thickness]::new(10, 60, 10, 60)
        $right = New-GuiCard -Title 'Browse' -Content @($body)
    }
    $g.OnSelectionChanged = { Update-MbtGuiExportFromFooter }
    Set-MbtGuiActions -Footer (Format-MbtGuiSelectionFooter $sel $s.Signed) -Label 'Continue to destination' -Enabled ($s.Signed -and $sel.Items -gt 0) -OnClick { Move-MbtGuiStep 1 }
    return (New-GuiTwoColumns -Left $left -Right $right)
}

#endregion

#region Export: To -----------------------------------------------------------------------------------------

function Get-MbtGuiDiskNeed {
    param($Selection, [hashtable]$S)
    $size = [double]$Selection.Bytes
    # EML files: the working store stays, and a .eml is base64 inside (a third larger)
    $need = if ($S.Dest -eq 'store') { $size } elseif ($S.Dest -eq 'eml') { $size * 2.5 } elseif ($S.KeepStore) { $size * 2 } else { $size * 1.1 }
    return [pscustomobject]@{ Size = $size; Need = $need; Unknown = ($Selection.UnknownSize -gt 0) }
}

function Test-MbtGuiExportDestination {
    <# Returns the problem in plain words, or $null. #>
    param([hashtable]$S)
    if ([string]::IsNullOrWhiteSpace($S.OutFolder)) { return 'Choose a folder on this computer.' }
    if ($S.Dest -eq 'pst') {
        $bad = Test-MbtGuiPstName $S.OutName
        if ($bad) { return $bad }
        $full = [IO.Path]::Combine($S.OutFolder, $S.OutName)
        if (Test-Path -LiteralPath $full) { return 'This file already exists. The tool never overwrites a file. Choose another name.' }
    }
    return $null
}

function Build-MbtGuiExportToPage {
    $g = $script:Gui; $s = $g.S
    $sel = $g.ExportSelection
    if (-not $s.OutName) { $stem = ($s.Mailbox -split '@')[0] -replace '[^\w\.\-]', '-'; $s.OutName = "$stem.pst" }
    $emlOn = Get-MbtGuiCapability 'EmlExport'
    $options = @(
        (New-GuiOption -Type Radio -Group dest -Title 'PST files' -Hint 'Open them in Outlook. Big mailboxes are split into several files.' -Checked ($s.Dest -eq 'pst') -OnChange { $s.Dest = 'pst'; Show-MbtGuiPage })
        (New-GuiOption -Type Radio -Group dest -Title 'Export store only' -Hint 'Our own format: checked, resumable, quick to import into a mailbox later.' -Checked ($s.Dest -eq 'store') -OnChange { $s.Dest = 'store'; Show-MbtGuiPage })
        (New-GuiOption -Type Radio -Group dest -Title $(if ($emlOn) { 'EML files' } else { 'EML files (coming in this version)' }) -Hint 'One .eml file per message, in the real folder tree.' -Checked ($s.Dest -eq 'eml') -Disabled:(-not $emlOn) -OnChange { $s.Dest = 'eml'; Show-MbtGuiPage }))
    $problem = Test-MbtGuiExportDestination $s
    $content = [Collections.Generic.List[object]]::new()
    foreach ($o in $options) { $content.Add($o) }
    $content.Add((New-GuiText 'Folder on this computer' 'Label'))
    $content.Add((New-MbtGuiBrowseFolderRow -Value $s.OutFolder -Kind Folder -Title 'Folder for the export' -Automation 'Folder' -OnChange { param($v) $s.OutFolder = $v; Refresh-MbtGuiExportTo }))
    $nameError = $null
    if ($s.Dest -eq 'pst') {
        $content.Add((New-GuiText 'File name' 'Label'))
        $name = New-GuiTextBox -Text $s.OutName -Automation 'File name' -OnChange { param($x, $e) $s.OutName = $x.Text; Refresh-MbtGuiExportTo }
        $content.Add($name)
        $stem = [IO.Path]::GetFileNameWithoutExtension($s.OutName)
        $content.Add((New-GuiText "Parts are named $stem.pst, $stem-002.pst, ... Names such as AUX, CON or NUL are refused (Windows reserves them)." 'Hint'))
        $split = [Windows.Controls.ComboBox]::new(); $split.HorizontalAlignment = 'Stretch'
        foreach ($t in '50 GB (Outlook''s default limit)', '20 GB', '10 GB', '5 GB') { [void]$split.Items.Add($t) }
        $split.SelectedIndex = switch ([int]$s.SplitGB) { 50 { 0 } 20 { 1 } 10 { 2 } 5 { 3 } default { 0 } }
        Add-MbtGuiHandler $split 'SelectionChanged' { param($sender, $e) $s.SplitGB = @(50, 20, 10, 5)[$sender.SelectedIndex]; $s.Changed['SplitGB'] = $true }
        $more = New-GuiMore -Summary 'More options' -Content @(
            (New-GuiText 'Largest PST file' 'Label'), $split,
            (New-GuiText 'A message is never cut in two. A file may be a little larger than the limit if one message does not fit. This choice applies to this run only.' 'Hint'),
            (New-GuiOption -Type Check -Title 'Keep the working store' -Hint 'Recommended. You can write the PST again, or import into a mailbox, without downloading again. Uses more disk space.' -Checked $s.KeepStore -OnChange { param($x, $e) $s.KeepStore = [bool]$x.IsChecked; Refresh-MbtGuiExportTo }),
            (New-GuiOption -Type Check -Title 'Check every PST file when it is written' -Hint 'Recommended.' -Checked $s.CheckEach -OnChange { param($x, $e) $s.CheckEach = [bool]$x.IsChecked }),
            (New-GuiOption -Type Check -Title 'One top folder per auxiliary archive' -Hint 'Off: auxiliary archives are merged into the Archive tree.' -Checked $s.SepAux -OnChange { param($x, $e) $s.SepAux = [bool]$x.IsChecked }))
        $content.Add($more)
    }
    $holder = [Windows.Controls.ContentControl]::new(); $holder.Name = 'ProblemHolder'
    $content.Add($holder)
    $g.ExportToHolder = $holder
    $left = New-GuiCard -Title 'Save as' -Content $content
    $diskHolder = [Windows.Controls.ContentControl]::new()
    $g.ExportDiskHolder = $diskHolder
    $good = New-GuiCard -Title 'Good to know' -Content @((New-GuiText 'Exporting is read only: the mailbox is not changed. If the PC sleeps or the window closes, nothing is lost: **Continue** does only what is left.' 'Body'))
    Refresh-MbtGuiExportTo
    Set-MbtGuiActions -Footer 'Choose where to save it. You can still change everything on the next page.' -Label 'Continue to Verify' -Enabled (-not $problem -and -not $g.ExportDiskBad) -OnClick { Move-MbtGuiStep 1 }
    Refresh-MbtGuiExportTo
    return (New-GuiTwoColumns -Left (New-GuiStack @($left)) -Right (New-GuiStack @($diskHolder, $good)))
}

function Refresh-MbtGuiExportTo {
    <# The problem line, the disk card and the primary button, without building the page again. #>
    $g = $script:Gui; $s = $g.S
    if (-not $g.ContainsKey('ExportToHolder') -or -not $g.ExportToHolder) { return }
    $problem = Test-MbtGuiExportDestination $s
    $g.ExportToHolder.Content = $(if ($problem -and ($s.OutName -or $s.Dest -ne 'pst')) { New-GuiMessage -Kind bad -Icon x -Text $problem } else { $null })
    $need = Get-MbtGuiDiskNeed $g.ExportSelection $s
    $free = $null
    try { $free = [double](& $g.Engine.FreeSpace $s.OutFolder) } catch { $free = $null }
    $enough = ($null -eq $free) -or ($need.Need -le $free)
    $g.ExportDiskBad = -not $enough
    $drive = if ($s.OutFolder -match '^([A-Za-z]:)') { $Matches[1].ToUpperInvariant() } else { 'the disk' }
    $needText = if ($need.Size -le 0 -and $need.Unknown) { 'unknown' } else { "about **$(Format-MbtGuiBytes $need.Need)**" + $(if ($s.Dest -eq 'pst' -and $s.KeepStore) { ' (store + PST)' } else { '' }) }
    $facts = New-GuiFacts @(@('Selection', $(if ($need.Size -le 0 -and $need.Unknown) { 'unknown' } else { Format-MbtGuiBytes $need.Size })), @('Needed on disk', (New-GuiText $needText 'Body')), @("Free on $drive", $(if ($null -ne $free) { Format-MbtGuiBytes $free } else { 'unknown' })))
    $facts.Margin = [Windows.Thickness]::new(0)
    $msg = if ($enough) { New-GuiMessage -Kind ok -Icon check -Text 'There is enough room.' } else { New-GuiMessage -Kind bad -Icon x -Text "**Not enough room on $drive.** Free about $(Format-MbtGuiBytes ($need.Need - $free)) or choose another folder." }
    $g.ExportDiskHolder.Content = New-GuiCard -Title 'Disk space' -Content @($facts, $msg, (New-GuiText 'If the disk fills up, the run stops by itself and can be continued once you free some space.' 'Hint'))
    if ((Get-MbtGuiScreen) -eq 'exTo') { $g.C.Primary.IsEnabled = (-not $problem) -and $enough }
}

#endregion

#region Plan (shared pieces) -------------------------------------------------------------------------------------

function Get-MbtGuiNodePath {
    param([MailboxTransferGui.FolderNode]$Node)
    $names = [Collections.Generic.List[string]]::new()
    $n = $Node
    while ($n) {
        $label = if ($n.IsRoot) { switch ($n.Kind) { 'primary' { $(if ($script:Gui.S.Shared) { "Shared mailbox $($script:Gui.S.Shared)" } else { 'Primary' }) } 'archive' { 'Archive' } 'aux' { $n.Name } default { $n.Name } } } else { $n.Name }
        $names.Insert(0, $label); $n = $n.Parent
    }
    return ($names -join ' / ')
}

function Test-MbtGuiSelectionHasDeletedItems {
    <# Is something below the folder Recoverable Items ticked? (The plan page says where those items are saved and where an import puts them.) #>
    param($Selection)
    return [bool]@($Selection.Entries | Where-Object { (Get-MbtGuiNodePath $_.Node) -match ' / Recoverable Items( / |$)' }).Count
}

function New-MbtGuiPlanTable {
    param($Selection)
    $rows = foreach ($e in $Selection.Entries) {
        $n = $e.Node; $root = Get-MbtGuiRootNode $n
        $note = switch ($e.Kind) {
            'tree' { $below = @($n.Descendants() | Where-Object { $_ -ne $n -and $_.Count -gt 0 }).Count; $(if ($below) { "and $below folder$(if ($below -gt 1) { 's' }) below it" } else { '' }) }
            'own' { 'this folder only' }
            default { "$($n.Picks.Count) message$(if ($n.Picks.Count -gt 1) { 's' }) ticked" }
        }
        if ($root.HeadersOnly) { $note = ($note + $(if ($note) { "`n" } else { '' }) + 'may be "not ready"; tried again later') }
        $bytes = 0L; $unknown = $false
        switch ($e.Kind) {
            'tree' { foreach ($d in $n.Descendants()) { if ($d.Count -gt 0) { if ($d.SizeBytes -lt 0) { $unknown = $true } else { $bytes += $d.SizeBytes } } } }
            'own' { if ($n.SizeBytes -lt 0) { $unknown = $true } else { $bytes = $n.SizeBytes } }
            default { foreach ($v in $n.Picks.Values) { $bytes += $v } }
        }
        $cell = [Windows.Controls.StackPanel]::new()
        $cell.Children.Add((New-GuiText (Get-MbtGuiNodePath $n) 'Body')) | Out-Null
        if ($note) { $sm = New-GuiText $note 'Small'; $sm.Margin = [Windows.Thickness]::new(0); $cell.Children.Add($sm) | Out-Null }
        , @($cell, (Format-MbtGuiCount $e.Items), $(if ($unknown -and $bytes -le 0) { 'size unknown' } elseif ($unknown) { (Format-MbtGuiBytes $bytes) + ' + unknown' } else { Format-MbtGuiBytes $bytes }))
    }
    $tbl = New-GuiTable -Headers @('Where', 'Items', 'Size') -Rows @($rows) -Widths @('*', '90', '110') -Right @(1, 2)
    if (-not @($rows).Count) { $tbl = New-GuiText 'Nothing ticked.' 'Muted' }
    return (New-GuiStack @($tbl, (New-GuiText 'Every folder is checked against the export service just before the start. A folder that cannot be matched stops the plan; nothing is guessed.' 'Hint')))
}

function New-MbtGuiSleepInfo {
    $items = @(
        '**The PC goes to sleep:** the job pauses and goes on when the PC wakes up. The console says how long it was asleep and re-computes the time left.'
        '**You close the window:** the console asks. *Stop and keep progress* finishes the items in flight, then stops. Reopen the console and press **Continue**.'
        '**The PC restarts or loses power:** nothing is lost. Items that were being sent are checked first when you continue.'
        '**The internet drops:** the tool waits and tries again by itself.')
    $stack = [Windows.Controls.StackPanel]::new()
    foreach ($t in $items) { $row = New-GuiText ([char]0x2022 + '  ' + ($t -replace '\*', '')) 'Body'; $row.FontSize = 13; $row.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $stack.Children.Add($row) | Out-Null }
    return (New-GuiMore -Summary 'What if this PC sleeps, or I close the window?' -Content @($stack))
}

function New-MbtGuiSettingsSection {
    <# Expert mode: the run settings, for this run only, never written to the configuration file. #>
    param([string[]]$Include = @('conc'))
    $g = $script:Gui; $s = $g.S
    $rows = [Collections.Generic.List[object]]::new()
    $rows.Add((New-GuiText '**What you change here applies to this run only and is never saved.** The defaults come from the configuration file; they suit almost every job.' 'Hint'))
    $label = { param($text, $key) $st = $s; $p = [Windows.Controls.StackPanel]::new(); $p.Orientation = 'Horizontal'; $p.Children.Add((New-GuiText $text 'Label')) | Out-Null; if ($s.Changed.ContainsKey($key)) { $pill = New-GuiPill 'modified' 'warn' -Light; $pill.Margin = [Windows.Thickness]::new(8, 6, 8, 0); $p.Children.Add($pill) | Out-Null; $p.Children.Add((New-GuiLink 'Reset' { $st.Changed.Remove($key); $defaults = @{ Conc = 4; Retry = 2; Tries = 6; Reconcile = 5; Awake = $false }; if ($defaults.ContainsKey($key)) { $st[$key] = $defaults[$key] }; $st.PlanState = 'idle'; Show-MbtGuiPage }.GetNewClosure())) | Out-Null }; $p }
    if ('conc' -in $Include) {
        $rows.Add((& $label 'Requests in parallel' 'Conc'))
        $cb = [Windows.Controls.ComboBox]::new(); $cb.HorizontalAlignment = 'Stretch'
        foreach ($t in '4 (default, the Microsoft limit)', '2', '1') { [void]$cb.Items.Add($t) }
        $cb.SelectedIndex = switch ([int]$s.Conc) { 4 { 0 } 2 { 1 } 1 { 2 } default { 0 } }
        Add-MbtGuiHandler $cb 'SelectionChanged' { param($sender, $e) $new = @(4, 2, 1)[$sender.SelectedIndex]; if ($new -ne $s.Conc) { $s.Conc = $new; $s.Changed['Conc'] = $true; Show-MbtGuiPage } }
        $rows.Add($cb); $rows.Add((New-GuiText 'The tool lowers this by itself when Microsoft slows it down.' 'Hint'))
    }
    if ('retry' -in $Include) {
        $rows.Add((& $label 'Wait before a failed item is tried again (seconds)' 'Retry'))
        $tb = New-GuiTextBox -Text ([string]$s.Retry) -Automation 'Seconds before a retry' -OnChange { param($x, $e) $v = 0; if ([int]::TryParse($x.Text, [ref]$v) -and $v -ge 1 -and $v -le 60 -and $v -ne $s.Retry) { $s.Retry = $v; $s.Changed['Retry'] = $true } }.GetNewClosure()
        $rows.Add($tb); $rows.Add((New-GuiText 'It doubles each time, up to 60 seconds. Microsoft''s own waiting requests are always obeyed.' 'Hint'))
    }
    if ('reconcile' -in $Include) {
        $rows.Add((& $label 'Wait before checking an uncertain item (minutes)' 'Reconcile'))
        $tb = New-GuiTextBox -Text ([string]$s.Reconcile) -Automation 'Minutes before a check' -OnChange { param($x, $e) $v = 0; if ([int]::TryParse($x.Text, [ref]$v) -and $v -ge 1 -and $v -le 60 -and $v -ne $s.Reconcile) { $s.Reconcile = $v; $s.Changed['Reconcile'] = $true } }.GetNewClosure()
        $rows.Add($tb); $rows.Add((New-GuiText 'Microsoft may still be creating the item. Default 5.' 'Hint'))
    }
    if ('tries' -in $Include) {
        $rows.Add((& $label 'Tries per item' 'Tries'))
        $tb = New-GuiTextBox -Text ([string]$s.Tries) -Automation 'Tries per item' -OnChange { param($x, $e) $v = 0; if ([int]::TryParse($x.Text, [ref]$v) -and $v -ge 1 -and $v -le 20 -and $v -ne $s.Tries) { $s.Tries = $v; $s.Changed['Tries'] = $true } }.GetNewClosure()
        $rows.Add($tb); $rows.Add((New-GuiText 'How many times an item is tried before it is marked failed. Default 3.' 'Hint'))
    }
    if ('awake' -in $Include) {
        $rows.Add((& $label 'While the job runs' 'Awake'))
        $rows.Add((New-GuiOption -Type Check -Title 'Keep this PC awake while it runs' -Hint 'Windows follows your company''s power settings unless you tick this. If the PC sleeps, the job stops and Continue does what is left.' -Checked ([bool]$s.Awake) -Margin '0' -OnChange { param($x, $e) $on = [bool]$x.IsChecked; if ($on -ne [bool]$s.Awake) { $s.Awake = $on; if ($on) { $s.Changed['Awake'] = $true } else { $s.Changed.Remove('Awake') }; Show-MbtGuiPage } }.GetNewClosure()))
    }
    $modified = @($s.Changed.Keys | Where-Object { $_ -in 'Conc', 'Retry', 'Tries', 'Reconcile', 'Awake' })
    return (New-GuiMore -Summary $(if ($modified.Count) { 'Expert mode (modified for this run)' } else { 'Expert mode' }) -Content $rows.ToArray() -Open:($modified.Count -gt 0))
}

function New-MbtGuiSettingsCard {
    <# The Expert mode fold at the bottom of a Verify page; nothing for a task that has no run settings. #>
    param([string[]]$Include)
    if (-not @($Include).Count) { return $null }
    return (New-GuiCard -Content @((New-MbtGuiSettingsSection -Include $Include)))
}

#endregion

#region The same from the command line --------------------------------------------------------------------------

function Get-MbtGuiCommandLine {
    <#
        The command line(s) of the task as an operator would type them (ConvertTo-MbtGuiCommandLine): the parameters that are a choice, no plumbing of the
        console, no value that is the default. With -Job: the job that ran; without: the one the Start button would begin.
    #>
    param([Parameter(Mandatory)][ValidateSet('export', 'import', 'repair', 'copy')][string]$Task, $Job)
    $s = $script:Gui.S
    $stub = if ($Job) { $Job } else { New-MbtGuiPreviewJob -Task $Task }
    try { $commands = @(Get-MbtGuiCommand -Task $Task -S $s -Job $stub) } catch { return '' }
    return (ConvertTo-MbtGuiCommandLine -Command $commands -S $s)
}

function New-MbtGuiCommandLineCard {
    <# "The same from the command line": the exact line, read only, with Copy. #>
    param([Parameter(Mandatory)][string]$Task, $Job)
    $text = Get-MbtGuiCommandLine -Task $Task -Job $Job
    $box = [Windows.Controls.TextBox]::new()
    $box.Text = $text; $box.IsReadOnly = $true; $box.TextWrapping = 'Wrap'; $box.FontFamily = [Windows.Media.FontFamily]::new('Cascadia Mono, Consolas'); $box.FontSize = 12
    $box.Padding = [Windows.Thickness]::new(8, 6, 8, 6); $box.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
    [Windows.Automation.AutomationProperties]::SetName($box, 'The same from the command line')
    $state = @{ Text = $text }
    $copy = New-GuiButton -Text 'Copy' -Small -Icon clip -OnClick { param($b, $e) if ($state.Text -and (Set-MbtGuiClipboard $state.Text)) { $b.Content = 'Copied' } }.GetNewClosure()
    $copy.Margin = [Windows.Thickness]::new(0)
    $hint = $(if ($Task -in 'export', 'copy' -or ($Task -eq 'import' -and $script:Gui.S.ImSrc -ne 'store')) { 'Run it in PowerShell 7 from the tool folder. The files it names are saved when the job starts.' } else { 'Run it in PowerShell 7 from the tool folder.' })
    return (New-GuiCard -Title 'The same from the command line' -Info $hint -Right $copy -Content @($box))
}

function New-MbtGuiVerifyBottom {
    <# Under the two columns of a Verify page: the command line and the Expert mode fold. #>
    param([Parameter(Mandatory)]$Columns, [Parameter(Mandatory)][string]$Task, [string[]]$Include)
    return (New-GuiStack @($Columns, (New-MbtGuiCommandLineCard -Task $Task), (New-MbtGuiSettingsCard -Include $Include)))
}

function New-MbtGuiDiskCheckRow {
    $g = $script:Gui; $s = $g.S
    $need = Get-MbtGuiDiskNeed $g.ExportSelection $s
    $free = $null; try { $free = [double](& $g.Engine.FreeSpace $s.OutFolder) } catch { $free = $null }
    $drive = if ($s.OutFolder -match '^([A-Za-z]:)') { $Matches[1].ToUpperInvariant() } else { 'the disk' }
    $enough = ($null -eq $free) -or ($need.Need -le $free)
    $detail = if ($need.Size -le 0 -and $need.Unknown) { "$(if ($null -ne $free) { Format-MbtGuiBytes $free } else { 'unknown' }) free on $drive" } else { "about $(Format-MbtGuiBytes $need.Need) needed, $(if ($null -ne $free) { Format-MbtGuiBytes $free } else { 'unknown' }) free on $drive" }
    return (New-MbtGuiCheckRow -State $(if ($enough) { 'ok' } else { 'bad' }) -Text $(if ($enough) { 'Enough disk space' } else { 'Not enough disk space' }) -Detail $detail)
}

function New-MbtGuiSignInCheckRow {
    $s = $script:Gui.S
    if ($s.Signed) { return (New-MbtGuiCheckRow -State ok -Text 'Sign-in is valid' -Detail $(if ($s.AuthMode -eq 'app') { "the application, working on $($s.Account)" } else { "$($s.Account), delegated" })) }
    return (New-MbtGuiCheckRow -State bad -Text 'Not signed in' -Detail 'go back to From and sign in')
}

#endregion

#region Export: Plan -----------------------------------------------------------------------------------------

function Build-MbtGuiExportPlanPage {
    $g = $script:Gui; $s = $g.S
    $sel = $g.ExportSelection
    $pst = $s.Dest -eq 'pst'
    $steps = @('Read the folder tree and list the items you chose.', 'Download them into the export store, checking each one.')
    if ($pst) { $steps += "Write the PST files ($($s.SplitGB) GB at most each) with the real folder tree."; $steps += 'Check every PST file.' }
    elseif ($s.Dest -eq 'eml') { $steps += 'Write one .eml file per message, in the real folder tree.' }
    else { $steps += 'Check the store.' }
    $ol = [Windows.Controls.StackPanel]::new(); $i = 1
    foreach ($t in $steps) { $r = New-GuiText ("$i.  $t") 'Body'; $r.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $ol.Children.Add($r) | Out-Null; $i++ }
    $eta = Get-MbtGuiPlanEta -Task export -Bytes $sel.Bytes
    $etaText = if ($sel.Bytes -le 0 -and $sel.UnknownSize) { 'cannot be estimated: Microsoft does not give the size of these folders. It shows once the export runs.' } elseif ($sel.UnknownSize) { "**$($eta.Text)** for the folders whose size is known; the others add to it" } else { "**$($eta.Text)** (Microsoft may slow it; we refine as it runs)" }
    $shared = @(); if ($s.Shared) { $shared = @(, @('Mailbox', "$($s.Shared) (shared, opened as $($s.Mailbox))")) }
    $facts = New-GuiFacts ($shared + @(@('Saved to', "$($s.OutFolder) ($(if ($pst) { 'PST files + working store' } elseif ($s.Dest -eq 'eml') { 'EML files + working store' } else { 'export store' }))"),
        @('Size', "$(Format-MbtGuiSelectionItems $sel) items $([char]0x00B7) $(Format-MbtGuiSelectionSize $sel)"), @('Time', (New-GuiText $etaText 'Body'))))
    $g.PlanHolder = [Windows.Controls.ContentControl]::new(); $g.PlanHolder.Content = New-MbtGuiPlanStateMessage
    $destRow = New-MbtGuiCheckRow -State $(if (Test-MbtGuiExportDestination $s) { 'bad' } else { 'ok' }) -Text 'Nothing is overwritten' -Detail $(if ($pst) { "$($s.OutName) does not exist yet" } else { 'the destination is free' })
    $checks = New-GuiCard -Title 'Checks' -Content @($g.PlanHolder, (New-MbtGuiSignInCheckRow), (New-MbtGuiDiskCheckRow), $destRow)
    $what = New-GuiCard -Title 'What will be done' -Content @($ol, $facts)
    $left = New-GuiStack @($checks, $what)
    $inItContent = [Collections.Generic.List[object]]::new()
    $inItContent.Add((New-MbtGuiPlanTable $sel))
    if (Test-MbtGuiSelectionHasDeletedItems $sel) {
        $inItContent.Add((New-GuiMessage -Kind info -Icon shield -Text '**Deleted items (Recoverable Items) are in this plan.** They are saved apart, under a folder named Recoverable Items, in the export and in the PST. If you import them later, they go into a normal folder named Restored deleted items, never back into Recoverable Items.'))
    }
    $inIt = New-GuiCard -Title 'What is in it' -Content $inItContent.ToArray()
    $before = New-GuiCard -Title 'Before you start' -Content @((New-GuiText 'The export only reads the mailbox. You can stop at any time and continue later.' 'Body'), (New-MbtGuiSleepInfo))
    $right = New-GuiStack @($inIt, $before)
    $key = Get-MbtGuiPlanKey
    if ($s.PlanKey -ne $key) { $s.PlanState = 'idle'; $s.PlanKey = $key; Start-MbtGuiPlanCheck }
    Set-MbtGuiActions -Footer 'Read the checks and the plan. Nothing starts before you press Start.' -Label 'Start the export' -Icon play -Enabled ($s.PlanState -eq 'ok') -OnClick { Start-MbtGuiRun }
    return (New-MbtGuiVerifyBottom -Columns (New-GuiTwoColumns -Left $left -Right $right -LeftWidth 520) -Task export -Include @('conc', 'retry', 'tries', 'awake'))
}

#endregion
