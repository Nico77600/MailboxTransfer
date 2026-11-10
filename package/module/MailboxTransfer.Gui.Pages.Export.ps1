#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - sign-in card and the pages of Browse and export (dot-sourced by Start-MailboxTransferConsole.ps1).

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

#region Plain-words errors (CONSOLE-SPEC 11) ---------------------------------------------------------

function ConvertTo-MbtGuiMessage {
    <# What happened, in one sentence, from a raw engine message. Returns Text and Tech (the raw message, for the collapsed technical line). #>
    param([string]$Raw)
    $r = [string]$Raw
    $text = switch -Regex ($r) {
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
    <# The tenant and the application of the configuration: what the engine has, completed by what was typed in the Sign-in card. Missing = the card asks. #>
    $g = $script:Gui; $s = $g.S
    $has = if ($g.Engine -is [Collections.IDictionary]) { $g.Engine.Contains('SignInSettings') } else { [bool]$g.Engine.PSObject.Properties['SignInSettings'] }
    $configured = if ($has) { & $g.Engine.SignInSettings } else { $null }
    $tenant = if ($s.TenantId) { $s.TenantId.Trim() } elseif ($configured) { [string]$configured.TenantId } else { '' }
    $client = if ($s.ClientId) { $s.ClientId.Trim() } elseif ($configured) { [string]$configured.ClientId } else { '' }
    [pscustomobject]@{ TenantId = $tenant; ClientId = $client; Asks = [bool]($configured -and -not ($configured.TenantId -and $configured.ClientId)) }
}

function Test-MbtGuiAppSignIn {
    <# The three run-only values of "the application" (tenant, application id, certificate thumbprint): the message in plain words, or $null when they can be used. #>
    param([hashtable]$S)
    if (-not $S.AppTenant.Trim()) { return 'Type the tenant: its id (a guid) or its domain, like contoso.onmicrosoft.com.' }
    if ($S.AppTenant.Trim() -in 'common', 'organizations', 'consumers') { return 'Type the id or the domain of your tenant, not a general name.' }
    if ($S.AppClient.Trim() -notmatch '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$') { return 'Type the application (client) id: a guid, like 00000000-0000-0000-0000-000000000000.' }
    if (($S.AppThumb -replace '[^0-9A-Fa-f]', '').Length -ne 40) { return 'Type the thumbprint of the certificate: 40 characters (letters a-f and digits), shown in the certificate details.' }
    return $null
}

function Start-MbtGuiSignIn {
    param([bool]$UseCode = $false)
    $g = $script:Gui; $s = $g.S
    $addr = $s.Mailbox.Trim()
    if ($addr -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
        $s.SignInError = 'Type the address of the mailbox, like adele.vance@contoso.com.'
        Show-MbtGuiPage; return
    }
    # "the application": tenant, application (client) id and the thumbprint of its certificate, asked for this run only (nothing is saved)
    $appArg = $null
    $settings = $null
    if ($s.AuthMode -eq 'app' -and (Get-MbtGuiCapability 'AppOnly')) {
        $UseCode = $false
        $bad = Test-MbtGuiAppSignIn -S $s
        if ($bad) { $s.SignInError = $bad; Show-MbtGuiPage; return }
        $appArg = @{ Tenant = $s.AppTenant.Trim(); ClientId = $s.AppClient.Trim(); Thumbprint = ($s.AppThumb -replace '[^0-9A-Fa-f]', '') }
    } else {
        $settings = Get-MbtGuiSignInSettings
        if ($settings.Asks) {
            $guid = '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$'
            if ($settings.TenantId -notmatch $guid -and $settings.TenantId -notmatch '^[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+$') { $s.SignInError = 'Type the tenant: its directory id (a GUID) or its domain name, like contoso.onmicrosoft.com.'; $s.SignInTech = ''; Show-MbtGuiPage; return }
            if ($settings.ClientId -notmatch $guid) { $s.SignInError = 'Type the application (client) id of the app registration: a GUID.'; $s.SignInTech = ''; Show-MbtGuiPage; return }
        }
    }
    $s.Mailbox = $addr; $s.SigningIn = $true; $s.SignInError = ''; $s.CodeShown = $UseCode; $s.Code = $null
    $codeFile = Get-MbtGuiDeviceCodeFile
    if ($UseCode -and (Test-Path -LiteralPath $codeFile)) { Remove-Item -LiteralPath $codeFile -Force -ErrorAction SilentlyContinue }
    Show-MbtGuiPage
    & $g.Engine.SignIn $addr $UseCode $codeFile {
        param($info)
        $s.SigningIn = $false; $s.Signed = $true; $s.Token = $info; $s.Account = $info.Account; $s.CodeShown = $false
        Read-MbtGuiLocations
        Show-MbtGuiPage
    }.GetNewClosure() {
        param($msg)
        $s.SigningIn = $false; $s.Signed = $false; $s.CodeShown = $false; $s.Code = $null
        $m = ConvertTo-MbtGuiMessage $msg
        $s.SignInError = $m.Text; $s.SignInTech = $m.Tech
        Show-MbtGuiPage
    }.GetNewClosure() $false $(if ($settings) { $settings.TenantId } else { '' }) $(if ($settings) { $settings.ClientId } else { '' }) $appArg
}

function Step-MbtGuiSignInCode {
    <# While a device-code sign-in waits, the engine writes "DEVICE CODE: <code> URL: <url>" to a file: show it in the card. #>
    $g = $script:Gui
    if (-not $g -or -not $g.S.SigningIn -or -not $g.S.CodeShown -or $g.S.Code) { return }
    $f = Get-MbtGuiDeviceCodeFile
    if (-not (Test-Path -LiteralPath $f)) { return }
    $line = (Get-Content -LiteralPath $f -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($line -match 'DEVICE CODE:\s*(\S+)\s+URL:\s*(\S+)') {
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
    <# Sign in (not yet), or Mailbox (signed in). Wording is generic on purpose: who may sign in for another mailbox is confirmed in the lab. #>
    param([string]$Use = 'export')
    $g = $script:Gui; $s = $g.S
    if ($s.Signed) {
        $box = New-GuiTextBox -Text $s.Mailbox -ReadOnly -Automation 'Mailbox address'
        $out = New-GuiButton -Text 'Sign out' -OnClick { & $g.Engine.SignOut; $s.Signed = $false; $s.Shared = ''; $s.ImShared = ''; $s.Locations = @(); $s.Code = $null; if ($g.Browsers.ContainsKey('export')) { Set-MbtGuiBrowserRoots -Browser $g.Browsers['export'] -Roots @() }; Show-MbtGuiPage }
        $out.Margin = [Windows.Thickness]::new(0)
        $row = New-GuiRow -Box $box -Button $out
        $tok = $s.Token
        $missing = if ($tok -and @($tok.Missing).Count) { (@($tok.Missing) -join ', ') } else { 'Nothing' }
        $facts = New-GuiFacts @(@('Access', $(if ($tok) { [string]$tok.Access } else { 'Delegated (you, with your rights)' })), @('Can read', $(if ($tok -and $tok.CanReadMail) { 'Mail, folders, items (live view)' } else { 'Folders and items (listing without subjects)' })), @('Missing', $(if ($missing -eq 'Nothing') { New-GuiText 'Nothing' 'Muted' } else { $missing })))
        $more = New-GuiMore -Summary 'Sign-in details' -Content @($facts)
        $msgs = @((New-GuiMessage -Kind ok -Icon check -Text $(if ($s.AuthMode -eq 'app') { "Connected as the application, working on **$($s.Account)**. Nothing is kept on this PC." } else { "Signed in as **$($s.Account)**. The sign-in is kept on this PC, encrypted for your Windows account." })))
        if ($tok -and -not $tok.CanReadMail) { $msgs += New-GuiMessage -Kind warn -Icon warn -Text 'Subjects need the Mail.Read permission. You can still export.' }
        return (New-GuiCard -Title 'Mailbox' -Right (New-GuiPill $(if ($s.AuthMode -eq 'app') { 'Connected' } else { 'Signed in' }) 'ok') -Content (@((New-GuiText 'Mailbox address' 'Label'), $row) + $msgs + @($more)))
    }
    $app = ($s.AuthMode -eq 'app') -and (Get-MbtGuiCapability 'AppOnly')
    $box = New-GuiTextBox -Text $s.Mailbox -Automation 'Mailbox address' -OnChange { param($x, $e) $s.Mailbox = $x.Text }
    $content = @()
    if (Get-MbtGuiCapability 'AppOnly') {
        $content += New-GuiText 'Sign in as' 'Label'
        $content += New-GuiOption -Type Radio -Group authmode -Title 'Me' -Hint 'A window from Microsoft opens. You see the mailboxes you are allowed to open.' -Checked (-not $app) -Margin '0,4,0,0' -OnChange { $s.AuthMode = 'user'; Show-MbtGuiPage }
        $content += New-GuiOption -Type Radio -Group authmode -Title 'The application (certificate)' -Hint 'No window and no user: for scheduled or bulk jobs, and to write into a shared mailbox. Your administrator sets the application up first (see the app-only guide).' -Checked $app -Margin '0,4,0,8' -OnChange { $s.AuthMode = 'app'; Show-MbtGuiPage }
    }
    if ($app) {
        $content += @((New-GuiText 'Mailbox the application works on' 'Label'), $box,
            (New-GuiText 'Required: the application has no mailbox of its own to fall back on.' 'Hint'),
            (New-GuiText 'Tenant (id or domain)' 'Label'), (New-GuiTextBox -Text $s.AppTenant -Automation 'Tenant' -OnChange { param($x, $e) $s.AppTenant = $x.Text }),
            (New-GuiText 'Application (client) id' 'Label'), (New-GuiTextBox -Text $s.AppClient -Automation 'Application id' -OnChange { param($x, $e) $s.AppClient = $x.Text }),
            (New-GuiText 'Certificate thumbprint' 'Label'), (New-GuiTextBox -Text $s.AppThumb -Automation 'Certificate thumbprint' -OnChange { param($x, $e) $s.AppThumb = $x.Text }),
            (New-GuiText 'The certificate is looked up on this PC (your certificates, then the computer ones). These three values are used for this run only; nothing is saved.' 'Hint'))
    } else {
        $content += @((New-GuiText 'Mailbox to look at' 'Label'), $box,
            (New-GuiText 'Use an account that is allowed to read this mailbox. A window from Microsoft opens.' 'Hint'))
        $settings = Get-MbtGuiSignInSettings
        if ($settings.Asks) {
            # the configuration names no tenant or no application: they are typed here (kept for this session)
            $tenantBox = New-GuiTextBox -Text $s.TenantId -Automation 'Tenant' -OnChange { param($x, $e) $s.TenantId = $x.Text }
            $clientBox = New-GuiTextBox -Text $s.ClientId -Automation 'Application id' -OnChange { param($x, $e) $s.ClientId = $x.Text }
            $content += @((New-GuiText 'Tenant' 'Label'), $tenantBox, (New-GuiText 'Application (client) id' 'Label'), $clientBox,
                (New-GuiText 'The configuration file has no tenant or application. Type the directory id (or domain) of your Microsoft 365 tenant and the application id of the app registered for this tool.' 'Hint'))
        }
    }
    $btnText = if ($s.SigningIn -and -not $s.CodeShown) { if ($app) { 'Connecting...' } else { 'Waiting for the Microsoft window...' } } else { if ($app) { 'Connect' } else { 'Sign in...' } }
    $signin = New-GuiButton -Text $btnText -Accent -Icon key -Disabled:($s.SigningIn) -OnClick { Start-MbtGuiSignIn -UseCode $false }
    $signin.Margin = [Windows.Thickness]::new(0, 12, 8, 0); $signin.HorizontalAlignment = 'Left'
    $content += $signin
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
            $btns.Children.Add((New-GuiButton -Text 'Copy the code' -Small -OnClick { [Windows.Clipboard]::SetText($s.Code.Code) })) | Out-Null
            $btns.Children.Add((New-GuiButton -Text 'Open the page' -Small -OnClick { Start-Process $s.Code.Url })) | Out-Null
            $inner += $btns
        } else { $inner += New-GuiText 'Getting a code...' 'Hint' }
        $inner += New-GuiText 'Waiting for you to finish... the console continues by itself.' 'Hint'
        $rec.Child = (New-GuiStack $inner)
        $content += $rec
    } elseif (-not $app) {
        $line = [Windows.Controls.WrapPanel]::new(); $line.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
        $line.Children.Add((New-GuiText 'The window does not open? ' 'Small')) | Out-Null
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
        $content.Add((New-GuiOption -Type Check -Title 'Keep the working store' -Hint 'Recommended. You can write the PST again, or import into a mailbox, without downloading again. Uses more disk space.' -Checked $s.KeepStore -Margin '0,12,0,0' -OnChange { param($x, $e) $s.KeepStore = [bool]$x.IsChecked; Refresh-MbtGuiExportTo }))
        $split = [Windows.Controls.ComboBox]::new(); $split.HorizontalAlignment = 'Stretch'
        foreach ($t in '50 GB (Outlook''s default limit)', '20 GB', '10 GB', '5 GB') { [void]$split.Items.Add($t) }
        $split.SelectedIndex = switch ([int]$s.SplitGB) { 50 { 0 } 20 { 1 } 10 { 2 } 5 { 3 } default { 0 } }
        $split.Add_SelectionChanged({ param($x, $e) $s.SplitGB = @(50, 20, 10, 5)[$x.SelectedIndex]; $s.Changed['SplitGB'] = $true })
        $more = New-GuiMore -Summary 'More options' -Content @(
            (New-GuiText 'Largest PST file' 'Label'), $split,
            (New-GuiText 'A message is never cut in two. A file may be a little larger than the limit if one message does not fit. This choice applies to this run only.' 'Hint'),
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
    Set-MbtGuiActions -Footer 'Choose where to save it. You can still change everything on the next page.' -Label 'Continue to the plan' -Enabled (-not $problem -and -not $g.ExportDiskBad) -OnClick { Move-MbtGuiStep 1 }
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
    foreach ($t in $items) { $row = New-GuiText ([char]0x2022 + '  ' + ($t -replace '\*(?!\*)', '')) 'Body'; $row.FontSize = 13; $row.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $stack.Children.Add($row) | Out-Null }
    return (New-GuiMore -Summary 'What if this PC sleeps, or I close the window?' -Content @($stack))
}

function New-MbtGuiSettingsSection {
    <# Run settings (this run only): never written to the configuration file. Extra: elements specific to the task. #>
    param([string[]]$Include = @('conc'), [bool]$Modified = $false)
    $g = $script:Gui; $s = $g.S
    $rows = [Collections.Generic.List[object]]::new()
    $rows.Add((New-GuiText '**What you change here applies to this run only and is never saved.** The defaults come from the configuration file.' 'Hint'))
    $changed = { param($key) $s.Changed[$key] = $true }
    $label = { param($text, $key) $st = $s; $p = [Windows.Controls.StackPanel]::new(); $p.Orientation = 'Horizontal'; $p.Children.Add((New-GuiText $text 'Label')) | Out-Null; if ($s.Changed.ContainsKey($key)) { $pill = New-GuiPill 'modified' 'warn' -Light; $pill.Margin = [Windows.Thickness]::new(8, 6, 8, 0); $p.Children.Add($pill) | Out-Null; $p.Children.Add((New-GuiLink 'Reset' { $st.Changed.Remove($key); $defaults = @{ Conc = 4; Retry = 2; Tries = 6; Reconcile = 5 }; if ($defaults.ContainsKey($key)) { $st[$key] = $defaults[$key] }; $st.PlanState = 'idle'; Show-MbtGuiPage }.GetNewClosure())) | Out-Null }; $p }
    if ('conc' -in $Include) {
        $rows.Add((& $label 'Requests in parallel' 'Conc'))
        $cb = [Windows.Controls.ComboBox]::new(); $cb.HorizontalAlignment = 'Stretch'
        foreach ($t in '4 (default, the Microsoft limit)', '2', '1') { [void]$cb.Items.Add($t) }
        $cb.SelectedIndex = switch ([int]$s.Conc) { 4 { 0 } 2 { 1 } 1 { 2 } default { 0 } }
        $cb.Add_SelectionChanged({ param($x, $e) $new = @(4, 2, 1)[$x.SelectedIndex]; if ($new -ne $s.Conc) { $s.Conc = $new; $s.Changed['Conc'] = $true; Show-MbtGuiPage } }.GetNewClosure())
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
        $rows.Add($tb)
    }
    $rows.Add((New-GuiOption -Type Check -Title 'Keep this PC awake while it runs' -Hint 'Stops Windows from going to sleep by itself. Closing the lid or choosing Sleep still sleeps.' -Checked $s.Awake -Margin '0,10,0,0' -OnChange { param($x, $e) $s.Awake = [bool]$x.IsChecked }))
    $rows.Add((& $label 'Work folder' 'WorkFolder'))
    $rows.Add((New-MbtGuiBrowseFolderRow -Value $s.WorkFolder -Kind Folder -Title 'Work folder' -Automation 'Work folder' -OnChange { param($v) if ($v -ne $s.WorkFolder) { $s.WorkFolder = $v; $s.Changed['WorkFolder'] = $true } }))
    return (New-GuiMore -Summary 'Run settings (this run only)' -Content $rows.ToArray())
}

function New-MbtGuiSettingsCard {
    param([string[]]$Include)
    $g = $script:Gui
    $note = if ($g.S.Changed.Count) { 'Modified for this run: ' + (($g.S.Changed.Keys | Sort-Object) -join ', ') + '. Nothing is saved.' } else { 'Defaults from the configuration file. Nothing modified.' }
    $n = New-GuiText $note 'Hint'; $n.Margin = [Windows.Thickness]::new(0)
    return (New-GuiCard -Title 'Settings used' -Content @($n, (New-MbtGuiSettingsSection -Include $Include)))
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
    $facts = New-GuiFacts @(@('Mailbox', $(if ($s.Shared) { "$($s.Shared) (shared, opened as $($s.Mailbox))" } else { $s.Mailbox })), @('Saved to', "$($s.OutFolder) ($(if ($pst) { 'PST files + working store' } elseif ($s.Dest -eq 'eml') { 'EML files + working store' } else { 'export store' }))"),
        @('Size', "$(Format-MbtGuiSelectionItems $sel) items $([char]0x00B7) $(Format-MbtGuiSelectionSize $sel)"), @('Time', (New-GuiText $etaText 'Body')))
    $planState = New-MbtGuiPlanStateMessage
    $g.PlanHolder = [Windows.Controls.ContentControl]::new(); $g.PlanHolder.Content = $planState
    $what = New-GuiCard -Title 'What will be done' -Content @($ol, $facts, $g.PlanHolder)
    $left = New-GuiStack @($what, (New-MbtGuiSettingsCard -Include @('conc', 'retry')))
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
    Set-MbtGuiActions -Footer 'Read the plan. Nothing starts before you press Start.' -Label 'Start the export' -Icon play -Enabled ($s.PlanState -eq 'ok') -OnClick { Start-MbtGuiRun }
    return (New-GuiTwoColumns -Left $left -Right $right)
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD95cPKTHkNDTlY
# jW2yhQDld55lgvEtvx7WXMxYGBync6CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCCGdCw1lLk
# J8FpOrflhwwzh0+UytcnZFupv1us/DMRqzANBgkqhkiG9w0BAQEFAASCAgAsSAcQ
# zT9qEaYFVUtDuHBqcFWCbZG8LvVMNVcet97lJ/PjAmNi+h+3+P2ocgjy/AkD/Y0j
# G4k52U5lh8jnAoyVLhmWmKgivuNtLRE5LTCDUKk6IVHyq91XF1dpjiyX1A8aWygv
# gk5MtCIP1bDnMgghVPJhzZiQ2roX1+HQktFva1hdWIkmoVppfnHrZxfHVKxDy4CJ
# uEB9gZKFvYtvXIVhLQmhceq8hF1Yw5qiSEhHRaOXXVKniW52MB8lO4SueokKIKE6
# iK4yJz9gTdTk3lWtDFu4xk4/DG8y/9WjYQFImBXpLc6MueIVrDI4jE6P+/RNkeVS
# wVaTKfH37CyecSl7CAMWWMeYd5Xh2GScEzSmtk0BCLDufpDk+1i37n2mQ/oeORnY
# 5RRL2Ihm7j8EmJFt1NuujoLK+B07Shx605L+y7kRC+nIBbInjNPYSVATHQPC7qMW
# wVdWL91qK2Hw4bzLQyZalKIUUXsjtgiclhvTN6JCRsnR+o15xxHJNKus/znFz5Dt
# fadNxrLior0C6dFjU2zXMFHqIc4yTyefBQhnJOOomMWN9hbrl6r9G/bxjoC4MGiJ
# XD3V8/FuOmzolGQ17IAPWzX85pSXPWj0yg3nzfG9DpaIqGDYlWv5Wcc2P5WqLw/3
# tOdCzmXurne00zmzc0qHYURBiEQsxl7cbTxSh6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIOH2jczzYnqlz5+qNX1wOCNvdQPF2FrB63SMzR0weIT3AghBmPwp/cBq
# DhgPMjAyNjExMTAwODIyMjlaMAMCAQECEBWPCaUdFiegUlSH4L/irTugghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjIyOVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIBfi
# QthcJFi/Siw/i1QT6GI3jfzYBEX5b3pxuLuD1yqDMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiA+/78njYfRpZLfN3+HW+JhpktUSBdu/hzX0XwRQsjKzgIhAPVj
# CRMVwM/c/aEH0zN4KhAgnHTKvmUgoVf6IhyBwL+W
# SIG # End signature block
