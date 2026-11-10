#Requires -Version 7.4
<#
.SYNOPSIS
    Lab tool: creates varied mail, calendar and contact items in the PRIMARY mailbox of the lab user with Microsoft Graph v1.0
    (delegated), to export them afterwards and validate the FastTransfer codec on real Exchange Online streams.

.DESCRIPTION
    The real corpus of the first exports only holds simple IPM.Note items. This script adds what is missing: HTML with an inline image,
    many recipients with Unicode names, embedded messages (1 and 2 levels), event and contact attachments, many attachment types,
    categories / flag / importance, extended properties of every type (single and multi-valued, named by name and by id), custom
    headers, very long lines and subject, empty bodies, a message created from MIME (embedded rfc822 in rfc822), calendar events
    (attendees, all-day, recurring with one modified and one cancelled occurrence, attachments, private) and contacts (full, minimal,
    with a photo). The plan is listed by -WhatIf (nothing is sent, no sign-in); see tools\LabFxCorpus.Lib.ps1 for the content.

    Where the items go (all created below one name, -RootName, default MBT-FXICS-Corpus; everything is found again by name and reused):

        mail folder      MBT-FXICS-Corpus\Mail           messages (created as drafts; fxc-14 is made to look received)
        calendar         MBT-FXICS-Corpus                a secondary calendar (the default calendar is not touched)
        contact folder   MBT-FXICS-Corpus                a contact folder

    Idempotent: every item has a subject (display name for a contact) that starts with [fxc-NN]. An item whose marker is already in
    its folder is skipped, so the script can be run again (after a failure, or to complete it). Nothing existing is ever deleted or
    replaced: items are created; the follow-up writes (attachments, extended properties, the modified and the cancelled occurrence of
    the recurring event, the contact photo) only touch the item the same run has just created. The recurring event is completed on
    a later run only if one of its two follow-up steps is missing.

    Graph does not accept everything everywhere, and this could not be tried offline, so each item is built defensively: a body that
    the service refuses (HTTP 400 ...) is replaced by the next, weaker variant of the same item; every attachment and every extended
    property is sent on its own, and a refusal is reported on the item (status Partial) without stopping the run. The final table
    says, per item, what was created and what the service refused. A failed item does not stop the others.

    Side effects to know: an event created with attendees makes Exchange send the invitations; the addresses end in
    @contoso.example / @fabrikam.example, so the lab mailbox may receive delivery failure reports in its Inbox (not in the corpus
    folders). -NoAttendees creates the event without attendees.

    Sign-in is delegated as the lab user with MSAL.NET (public client "Microsoft Graph Command Line Tools"), device code by
    default when there is no interactive console: the line "DEVICE CODE: <code> URL: <url>" is printed at once (flushed) and
    written to a text file. Scopes: Mail.ReadWrite, Calendars.ReadWrite, Contacts.ReadWrite (offline_access is added by MSAL). The
    first sign-in asks for the consent of the three; the token cache (DPAPI, current user) makes later runs silent. It is not the
    cache of the lab fill tool (other scopes).

    At the end: the path of each folder when Graph tells it (to give to the export; for a secondary calendar it does not: the manifest
    then holds the id, the name and how to find the path), the counts, and a manifest written to -OutputPath even when items failed.

.PARAMETER UserPrincipalName
    The lab user (default: Lab.UserPrincipalName of package\config\MailboxTransfer.config.psd1).

.PARAMETER TenantId
    Default: Lab.TenantId of the configuration.

.PARAMETER ClientId
    Public client application (default 14d82eec-204b-4c2f-b7e8-296a70dab67e, Microsoft Graph Command Line Tools).

.PARAMETER RootName
    Name of the root mail folder, of the secondary calendar and of the contact folder (default MBT-FXICS-Corpus).

.PARAMETER Only
    Limits the run to some items: keys (fxc-07, 7) or kinds (Mail, Event, Contact).

.PARAMETER NoAttendees
    Creates the event fxc-20 without attendees (no invitation is sent).

.PARAMETER DeviceCode
    Device code instead of the browser. Default: on when the console is redirected, off otherwise.

.PARAMETER DeviceCodeFile
    File that receives the device code line (default <repository>\out\device-code-fxcorpus.txt).

.PARAMETER TokenCachePath
    DPAPI token cache (default %LOCALAPPDATA%\MailboxTransfer\lab-fxcorpus-token.cache).

.PARAMETER OutputPath
    Manifest of the run, JSON (default <repository>\out\fxics-corpus-created.json; out\ is git-ignored).

.PARAMETER ConfigPath
    Configuration file read for the Lab defaults (default package\config\MailboxTransfer.config.psd1).

.PARAMETER AccessToken
    A Graph access token to use as is instead of signing in (no renewal). Meant for the tests against a local fake service.

.PARAMETER GraphBaseUrl
    Graph root (default https://graph.microsoft.com/v1.0). Another value is accepted only for http://127.0.0.1 or http://localhost
    (tests against a local fake service).

.EXAMPLE
    .\tools\New-LabFxCorpus.ps1 -WhatIf

.EXAMPLE
    .\tools\New-LabFxCorpus.ps1 -DeviceCode

.EXAMPLE
    # only the mail items 7 to 9, and the contacts
    .\tools\New-LabFxCorpus.ps1 -Only 7, 8, 9, Contact

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

    As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$UserPrincipalName,
    [string]$TenantId,
    [string]$ClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e',
    [string]$RootName = 'MBT-FXICS-Corpus',
    [string[]]$Only,
    [switch]$NoAttendees,
    [switch]$DeviceCode,
    [string]$DeviceCodeFile,
    [string]$TokenCachePath,
    [string]$OutputPath,
    [string]$ConfigPath,
    [string]$AccessToken,
    [string]$GraphBaseUrl = 'https://graph.microsoft.com/v1.0'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = Split-Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'LabFxCorpus.Lib.ps1')

#region Defaults and plan ---------------------------------------------------------------------------------------------------------

if (-not $ConfigPath) { $ConfigPath = Join-Path $repo 'package\config\MailboxTransfer.config.psd1' }
$lab = if (Test-Path -LiteralPath $ConfigPath) { (Import-PowerShellDataFile -LiteralPath $ConfigPath).Lab } else { $null }
# the private local file (see Import-MbtConfiguration) overrides the Lab defaults
$localConfig = if ($null -ne $env:MAILBOXTRANSFER_LOCAL_CONFIG) { $env:MAILBOXTRANSFER_LOCAL_CONFIG } else { Join-Path $repo 'package\config\MailboxTransfer.config.local.psd1' }
if ($localConfig -and (Test-Path -LiteralPath $localConfig)) {
    $localLab = (Import-PowerShellDataFile -LiteralPath $localConfig).Lab
    if ($localLab) { $merged = @{}; if ($lab) { foreach ($k in $lab.Keys) { $merged[$k] = $lab[$k] } }; foreach ($k in $localLab.Keys) { if ($localLab[$k]) { $merged[$k] = $localLab[$k] } }; $lab = $merged }
}
if (-not $UserPrincipalName -and $lab) { $UserPrincipalName = $lab.UserPrincipalName }
if (-not $TenantId -and $lab) { $TenantId = $lab.TenantId }
if (-not $UserPrincipalName) { throw 'No -UserPrincipalName and no Lab.UserPrincipalName in the configuration.' }
if (-not $TenantId) { throw 'No -TenantId and no Lab.TenantId in the configuration.' }
if (-not $DeviceCodeFile) { $DeviceCodeFile = Join-Path $repo 'out\device-code-fxcorpus.txt' }
if (-not $TokenCachePath) { $TokenCachePath = Join-Path $env:LOCALAPPDATA 'MailboxTransfer\lab-fxcorpus-token.cache' }
if (-not $OutputPath) { $OutputPath = Join-Path $repo 'out\fxics-corpus-created.json' }
if ($GraphBaseUrl -ne 'https://graph.microsoft.com/v1.0' -and $GraphBaseUrl -notmatch '^http://(127\.0\.0\.1|localhost)(:\d+)?(/|$)') {
    throw "-GraphBaseUrl: only https://graph.microsoft.com/v1.0 or a local http://127.0.0.1 / http://localhost address is accepted."
}
$useDeviceCode = if ($PSBoundParameters.ContainsKey('DeviceCode')) { [bool]$DeviceCode } else { [Console]::IsInputRedirected -or [Console]::IsOutputRedirected }

function Write-Fxc {
    param([ValidateSet('INFO', 'OK', 'WARN', 'FAIL', 'STEP')][string]$Level, [string]$Text)
    $color = @{ INFO = 'Gray'; OK = 'Green'; WARN = 'Yellow'; FAIL = 'Red'; STEP = 'Cyan' }[$Level]
    Write-Host ('{0,-5} {1}' -f $Level, $Text) -ForegroundColor $color
}

$plan = Get-FxcPlan -Mailbox $UserPrincipalName
if ($NoAttendees) {
    foreach ($s in $plan | Where-Object Key -eq 'fxc-20') { foreach ($v in $s.Variants) { [void]$v.Remove('attendees') } }
}
if ($Only) {
    $wanted = @($Only | ForEach-Object { $_ -split '[,; ]+' } | Where-Object { $_ } | ForEach-Object { if ($_ -match '^\d+$') { 'fxc-{0:D2}' -f [int]$_ } else { $_ } })
    $plan = @($plan | Where-Object { $_.Key -in $wanted -or $_.Kind -in $wanted })
    if (-not $plan) { throw "-Only '$($Only -join ', ')' matches no item." }
}

Write-Fxc STEP "Lab FastTransfer corpus for $UserPrincipalName (tenant $TenantId), root '$RootName'"
Write-Host ((Get-FxcPlanSummary -Plan $plan | Format-Table Key, Kind, Via, BodyVariants, Attachments, ExtProps, AttachmentBytes, Extra, Title -AutoSize | Out-String -Width 220).TrimEnd())
$folderPlan = @(
    if ($plan | Where-Object Kind -eq 'Mail') { "mail folder      $RootName\Mail" }
    if ($plan | Where-Object Kind -eq 'Event') { "calendar         $RootName (secondary calendar)" }
    if ($plan | Where-Object Kind -eq 'Contact') { "contact folder   $RootName" }
)
Write-Host ("Folders (created if missing, else reused):`n  " + ($folderPlan -join "`n  "))

if (-not $PSCmdlet.ShouldProcess("$UserPrincipalName ($($plan.Count) items)", "Create the corpus below '$RootName' with Microsoft Graph")) {
    Write-Fxc WARN 'WhatIf: nothing was sent, nobody signed in.'
    return
}

#endregion

#region Sign-in and Graph ---------------------------------------------------------------------------------------------------------

function Import-FxcAuthType {
    # Compiles LabFxAuth.cs once per process with the MSAL.NET of the module Microsoft.Graph.Authentication.
    if ('LabFx.Auth' -as [type]) { return }
    $loaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Microsoft.Identity.Client' } | Select-Object -First 1
    if ($loaded) { $msal = $loaded.Location }
    else {
        $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module) { throw 'MSAL.NET not found: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser.' }
        $deps = Join-Path $module.ModuleBase 'Dependencies'
        $msal = Join-Path $deps 'Core\Microsoft.Identity.Client.dll'
        $abstractions = Join-Path $deps 'Microsoft.IdentityModel.Abstractions.dll'
        if (-not ([AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'Microsoft.IdentityModel.Abstractions' }) -and (Test-Path -LiteralPath $abstractions)) { Add-Type -LiteralPath $abstractions }
        Add-Type -LiteralPath $msal
    }
    Add-Type -AssemblyName System.Security.Cryptography.ProtectedData
    $references = @($msal, 'System.Security.Cryptography.ProtectedData', 'System.Linq', 'System.Runtime', 'System.Threading', 'System.Threading.Tasks',
        'System.Threading.Thread', 'System.Collections', 'System.Memory', 'netstandard', 'System.Security.Cryptography', 'System.Text.Encoding.Extensions',
        'System.Runtime.Extensions', 'System.IO.FileSystem', 'System.Console', 'System.Private.Uri', 'System.ComponentModel.Primitives')
    Add-Type -LiteralPath (Join-Path $PSScriptRoot 'LabFxAuth.cs') -ReferencedAssemblies $references -IgnoreWarnings -WarningAction SilentlyContinue
}

$script:GraphRoot = $GraphBaseUrl.TrimEnd('/')
$script:Http = $null
$script:Auth = $null

function Invoke-FxcGraph {
    # One Graph call with a fresh Bearer token, retried on 429 / 502 / 503 / 504 and network errors (Retry-After honoured).
    # Returns the parsed JSON (or $null for an empty answer). Error: exception with Data['Status'] (HTTP code) and the Graph error text.
    param(
        [Parameter(Mandatory)][string]$Method, [Parameter(Mandatory)][string]$Uri, $Body, [byte[]]$BodyBytes, [string]$ContentType = 'application/json',
        [switch]$RetryNotFound, [int]$MaxAttempts = 6
    )
    if ($Uri -notmatch '^https?://') { $Uri = "$script:GraphRoot/$($Uri.TrimStart('/'))" }
    $bytes = if ($null -ne $BodyBytes) { $BodyBytes } elseif ($null -ne $Body) { [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 30 -Compress)) }
    for ($attempt = 1; ; $attempt++) {
        $req = [Net.Http.HttpRequestMessage]::new([Net.Http.HttpMethod]::new($Method), $Uri)
        $req.Headers.Authorization = [Net.Http.Headers.AuthenticationHeaderValue]::new('Bearer', $script:Auth.GetToken())
        if ($null -ne $bytes) {
            $req.Content = [Net.Http.ByteArrayContent]::new($bytes)
            $req.Content.Headers.ContentType = [Net.Http.Headers.MediaTypeHeaderValue]::Parse($(if ($ContentType -eq 'application/json') { 'application/json; charset=utf-8' } else { $ContentType }))
        }
        $status = $null; $text = $null; $retryAfter = $null
        try {
            $resp = $script:Http.SendAsync($req).GetAwaiter().GetResult()
            $status = [int]$resp.StatusCode
            $text = $resp.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            if ($resp.Headers.RetryAfter -and $resp.Headers.RetryAfter.Delta) { $retryAfter = $resp.Headers.RetryAfter.Delta.TotalSeconds }
            $resp.Dispose()
        } catch { $text = $_.Exception.Message } finally { $req.Dispose() }
        if ($null -ne $status -and $status -ge 200 -and $status -lt 300) {
            if ([string]::IsNullOrWhiteSpace($text)) { return $null }
            return ($text | ConvertFrom-Json -Depth 30)
        }
        $transient = ($null -eq $status) -or $status -in 429, 502, 503, 504 -or ($RetryNotFound -and $status -eq 404 -and $attempt -le 4)
        if (-not $transient -or $attempt -ge $MaxAttempts) {
            $detail = $text; try { $e = ($text | ConvertFrom-Json).error; if ($e) { $detail = "$($e.code): $($e.message)" } } catch { }
            $ex = [Exception]::new("Graph $Method $($Uri -replace '^https?://[^/]+/v1.0/', '') -> HTTP $status : $detail")
            if ($null -ne $status) { $ex.Data['Status'] = $status }
            throw $ex
        }
        $wait = if ($retryAfter) { [Math]::Min(120, $retryAfter) } else { [Math]::Min(30, [Math]::Pow(2, $attempt)) }
        Write-Fxc WARN "HTTP $status, again in $wait s ($attempt/$MaxAttempts)"
        Start-Sleep -Seconds $wait
    }
}

function Get-FxcAll {
    # Every element of a collection, following @odata.nextLink.
    param([Parameter(Mandatory)][string]$Uri)
    while ($Uri) {
        $page = Invoke-FxcGraph -Method GET -Uri $Uri
        if ($null -eq $page) { return }
        foreach ($v in @($page.value)) { $v }
        $Uri = $page.PSObject.Properties['@odata.nextLink']?.Value
    }
}

function Test-FxcRefused {
    # True when the service refused THIS payload (a weaker variant can be tried), false for anything else (permission, throttling...).
    param($ErrorRecord)
    $s = $ErrorRecord.Exception.Data['Status']
    $null -ne $s -and [int]$s -in 400, 404, 405, 409, 413, 415, 422
}

#endregion

#region Folders -------------------------------------------------------------------------------------------------------------------

function Get-FxcMailFolder {
    param([Parameter(Mandatory)][string]$ParentId, [Parameter(Mandatory)][string]$Name)
    $esc = [uri]::EscapeDataString("displayName eq '" + $Name.Replace("'", "''") + "'")
    $hit = @(Get-FxcAll "me/mailFolders/$ParentId/childFolders?`$filter=$esc&`$select=id,displayName&`$top=50") | Where-Object displayName -ceq $Name | Select-Object -First 1
    if ($hit) { return [pscustomobject]@{ Id = $hit.id; Created = $false } }
    [pscustomobject]@{ Id = (Invoke-FxcGraph -Method POST -Uri "me/mailFolders/$ParentId/childFolders" -Body @{ displayName = $Name }).id; Created = $true }
}

function Get-FxcCalendar {
    param([Parameter(Mandatory)][string]$Name)
    $hit = @(Get-FxcAll 'me/calendars?$top=100') | Where-Object name -ceq $Name | Select-Object -First 1
    if ($hit) { return [pscustomobject]@{ Id = $hit.id; Created = $false } }
    [pscustomobject]@{ Id = (Invoke-FxcGraph -Method POST -Uri 'me/calendars' -Body @{ name = $Name }).id; Created = $true }
}

function Get-FxcContactFolder {
    param([Parameter(Mandatory)][string]$Name)
    $hit = @(Get-FxcAll 'me/contactFolders?$top=100&$select=id,displayName') | Where-Object displayName -ceq $Name | Select-Object -First 1
    if ($hit) { return [pscustomobject]@{ Id = $hit.id; Created = $false } }
    [pscustomobject]@{ Id = (Invoke-FxcGraph -Method POST -Uri 'me/contactFolders' -Body @{ displayName = $Name }).id; Created = $true }
}

function Get-FxcExisting {
    # Markers already present in a folder: key (fxc-NN) -> id.
    param([Parameter(Mandatory)][string]$Uri, [Parameter(Mandatory)][string]$Property)
    $found = @{}
    foreach ($item in Get-FxcAll $Uri) {
        $key = Get-FxcMarkerFromSubject $item.$Property
        if ($key -and -not $found.ContainsKey($key)) { $found[$key] = $item.id }
    }
    $found
}

function Resolve-FxcFolderPath {
    # Path of a folder as the export tree names it (below the mailbox root), or $null when Graph does not tell.
    param([Parameter(Mandatory)][string]$Id, [string]$ContactFolder)
    try {
        $rootId = (Invoke-FxcGraph -Method GET -Uri 'me/mailFolders/msgfolderroot?$select=id').id
        $names = [Collections.Generic.List[string]]::new()
        $current = $Id
        for ($i = 0; $i -lt 12 -and $current -and $current -ne $rootId; $i++) {
            $f = if ($ContactFolder -and $i -eq 0) { Invoke-FxcGraph -Method GET -Uri "me/contactFolders/${current}?`$select=id,displayName,parentFolderId" }
            else { Invoke-FxcGraph -Method GET -Uri "me/mailFolders/${current}?`$select=id,displayName,parentFolderId" }
            $names.Insert(0, $f.displayName); $current = $f.parentFolderId
        }
        if ($names.Count) { return ($names -join '/') }
    } catch { }
    $null
}

#endregion

#region Items ---------------------------------------------------------------------------------------------------------------------

function Get-FxcItemUri {
    param([string]$Kind, [string]$Id)
    switch ($Kind) { 'Mail' { "me/messages/$Id" } 'Event' { "me/events/$Id" } 'Contact' { "me/contacts/$Id" } }
}

function New-FxcItem {
    # Creates one step. Returns Status Created | Partial | Failed, the id and the notes.
    param([Parameter(Mandatory)]$Step, [Parameter(Mandatory)][string]$FolderId)
    $notes = [Collections.Generic.List[string]]::new()
    $result = [pscustomobject]@{ Key = $Step.Key; Kind = $Step.Kind; Status = 'Failed'; Id = $null; Notes = $notes; Title = $Step.Title }
    $collection = switch ($Step.Kind) { 'Mail' { "me/mailFolders/$FolderId/messages" } 'Event' { "me/calendars/$FolderId/events" } 'Contact' { "me/contactFolders/$FolderId/contacts" } }

    # 1. creation: the first body the service accepts
    $created = $null
    $weaker = 0
    if ($Step.Mime) {
        $created = Invoke-FxcGraph -Method POST -Uri $collection -BodyBytes ([Text.Encoding]::ASCII.GetBytes([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Step.Mime)))) -ContentType 'text/plain'
    } else {
        for ($i = 0; $i -lt $Step.Variants.Count -and -not $created; $i++) {
            try {
                $created = Invoke-FxcGraph -Method POST -Uri $collection -Body $Step.Variants[$i]
                $weaker = $i
            } catch {
                if (-not (Test-FxcRefused $_) -or $i -eq $Step.Variants.Count - 1) { throw }
                $notes.Add("body variant $($i + 1) of $($Step.Variants.Count) refused: $($_.Exception.Message)")
            }
        }
    }
    $result.Id = $created.id
    $result.Status = if ($weaker -gt 0) { 'Partial' } else { 'Created' }
    if ($weaker -gt 0) { $notes.Add("created with the weaker body variant $($weaker + 1) of $($Step.Variants.Count)") }
    if ($Step.ReportSubject) { $notes.Add("subject sent: $($Step.Variants[0].subject.Length) characters; stored: $(([string]$created.subject).Length)") }
    $item = Get-FxcItemUri $Step.Kind $created.id

    # The item exists from here on: whatever goes wrong next makes it Partial, not Failed (a rerun would skip it).
    try {
        # 2. attachments, one by one
        foreach ($a in $Step.Attachments) {
            $ok = $false
            for ($i = 0; $i -lt $a.Variants.Count -and -not $ok; $i++) {
                try {
                    [void](Invoke-FxcGraph -Method POST -Uri "$item/attachments" -Body $a.Variants[$i] -RetryNotFound)
                    $ok = $true
                    if ($i -gt 0) { $notes.Add("attachment '$($a.Name)': weaker variant $($i + 1) used"); $result.Status = 'Partial' }
                } catch {
                    if (-not (Test-FxcRefused $_)) { throw }
                    $notes.Add("attachment '$($a.Name)' variant $($i + 1) refused: $($_.Exception.Message)")
                }
            }
            if (-not $ok) { $result.Status = 'Partial' }
        }

        # 3. extended properties, one by one (a type the service does not know does not spoil the others)
        foreach ($e in $Step.Ext) {
            $ok = $false
            for ($i = 0; $i -lt $e.Variants.Count -and -not $ok; $i++) {
                try {
                    [void](Invoke-FxcGraph -Method PATCH -Uri $item -Body (Get-FxcExtBody -Ext $e -Variant $i) -RetryNotFound)
                    $ok = $true
                    if ($i -gt 0) { $notes.Add("extended property '$($e.Id)': value variant $($i + 1) accepted") }
                } catch {
                    if (-not (Test-FxcRefused $_)) { throw }
                    $notes.Add("extended property '$($e.Id)' value $($i + 1) refused: $($_.Exception.Message)")
                }
            }
            if (-not $ok) { $result.Status = 'Partial' }
        }

        # 4. photo
        if ($Step.Photo) {
            try { [void](Invoke-FxcGraph -Method PUT -Uri "me/contactFolders/$FolderId/contacts/$($created.id)/photo/`$value" -BodyBytes $Step.Photo -ContentType 'image/jpeg' -RetryNotFound) }
            catch { $notes.Add("photo refused: $($_.Exception.Message)"); $result.Status = 'Partial' }
        } elseif ($Step.Kind -eq 'Contact' -and $Step.Key -eq 'fxc-32') { $notes.Add('no photo: System.Drawing is not available'); $result.Status = 'Partial' }

        # 5. the recurring event: one modified and one cancelled occurrence
        if ($Step.Special -eq 'RecurringExceptions') { Complete-FxcRecurring -Result $result -EventId $created.id }
    } catch {
        $notes.Add("follow-up stopped: $($_.Exception.Message)"); $result.Status = 'Partial'
    }
    $result
}

function ConvertTo-FxcDateTime {
    # ConvertFrom-Json turns an ISO date string into a [datetime] (kind Unspecified when there is no Z): keep the wall clock as is.
    param($Value)
    if ($Value -is [datetime]) { return $Value }
    [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None)
}

function Complete-FxcRecurring {
    # Modifies the occurrence of 2026-11-16 and cancels the one of 2026-11-30 of the series just created (or found without them).
    param([Parameter(Mandatory)]$Result, [Parameter(Mandatory)][string]$EventId)
    $window = 'startDateTime=2026-11-01T00:00:00Z&endDateTime=2026-12-31T00:00:00Z&$select=id,subject,type,start,end&$top=50'
    $list = { @(Get-FxcAll "me/events/$EventId/instances?$window") }
    try {
        $instances = & $list
        $expected = 6
        if (-not ($instances | Where-Object type -eq 'exception')) {
            $target = $instances | Where-Object { (ConvertTo-FxcDateTime $_.start.dateTime).ToString('yyyy-MM-dd') -eq '2026-11-16' } | Select-Object -First 1
            if ($target) {
                $start = (ConvertTo-FxcDateTime $target.start.dateTime).AddHours(1)
                $end = (ConvertTo-FxcDateTime $target.end.dateTime).AddHours(1)
                [void](Invoke-FxcGraph -Method PATCH -Uri "me/events/$($target.id)" -Body @{
                    subject = '[fxc-22] Weekly recurring event (modified occurrence)'; location = @{ displayName = 'Room 3 (moved)' }
                    start = (New-FxcDateTime $start.ToString('yyyy-MM-ddTHH:mm:ss.0000000') 'UTC'); end = (New-FxcDateTime $end.ToString('yyyy-MM-ddTHH:mm:ss.0000000') 'UTC') })
                $Result.Notes.Add('occurrence 2026-11-16 modified (exception)')
            } else { $Result.Notes.Add('occurrence 2026-11-16 not found: no exception'); $Result.Status = 'Partial' }
            $instances = & $list
        }
        if (@($instances).Count -ge $expected) {
            $target = $instances | Where-Object { (ConvertTo-FxcDateTime $_.start.dateTime).ToString('yyyy-MM-dd') -eq '2026-11-30' } | Select-Object -First 1
            if ($target) { [void](Invoke-FxcGraph -Method DELETE -Uri "me/events/$($target.id)"); $Result.Notes.Add('occurrence 2026-11-30 cancelled') }
            else { $Result.Notes.Add('occurrence 2026-11-30 not found: nothing cancelled'); $Result.Status = 'Partial' }
        }
    } catch { $Result.Notes.Add("recurring follow-up failed: $($_.Exception.Message)"); $Result.Status = 'Partial' }
}

#endregion

#region Run -----------------------------------------------------------------------------------------------------------------------

if ($AccessToken) {
    $script:Auth = [pscustomobject]@{ Account = $UserPrincipalName; Token = $AccessToken }
    $script:Auth | Add-Member -MemberType ScriptMethod -Name GetToken -Value { $this.Token }
    Write-Fxc INFO 'using the access token given, no sign-in'
} else {
    Import-FxcAuthType
    $scopes = [string[]]@('https://graph.microsoft.com/Mail.ReadWrite', 'https://graph.microsoft.com/Calendars.ReadWrite', 'https://graph.microsoft.com/Contacts.ReadWrite')
    $script:Auth = [LabFx.Auth]::new($ClientId, $TenantId, $scopes, $TokenCachePath, $UserPrincipalName)
    $script:Auth.DeviceCode = $useDeviceCode
    $script:Auth.DeviceCodeFile = $DeviceCodeFile
    if (Test-Path -LiteralPath $DeviceCodeFile) { Remove-Item -LiteralPath $DeviceCodeFile -Force -ErrorAction SilentlyContinue }
    Write-Fxc INFO ("signing in as {0} ({1})" -f $UserPrincipalName, $(if ($useDeviceCode) { "device code: read the line 'DEVICE CODE:' (also in $DeviceCodeFile)" } else { 'browser or cache' }))
    [void]$script:Auth.SignIn()
    Write-Fxc OK "signed in as $($script:Auth.Account)"
}
$script:Http = [Net.Http.HttpClient]::new()
$script:Http.Timeout = [TimeSpan]::FromMinutes(5)

$results = [Collections.Generic.List[object]]::new()
$folders = [ordered]@{}
$folderNames = [ordered]@{}
$paths = [ordered]@{}
$runError = $null
try {
    foreach ($kind in 'Mail', 'Event', 'Contact') {
        $steps = @($plan | Where-Object Kind -eq $kind)
        if (-not $steps) { continue }
        try {
            # folder
            $f = switch ($kind) {
                'Mail' { $r = Get-FxcMailFolder -ParentId 'msgfolderroot' -Name $RootName; Get-FxcMailFolder -ParentId $r.Id -Name 'Mail' }
                'Event' { Get-FxcCalendar -Name $RootName }
                'Contact' { Get-FxcContactFolder -Name $RootName }
            }
            $folders[$kind] = $f.Id
            $folderNames[$kind] = if ($kind -eq 'Mail') { 'Mail' } else { $RootName }
            Write-Fxc STEP ("{0}: folder {1}" -f $kind, $(if ($f.Created) { 'created' } else { 'found' }))
            $existing = switch ($kind) {
                'Mail' { Get-FxcExisting "me/mailFolders/$($f.Id)/messages?`$select=id,subject&`$top=100" 'subject' }
                'Event' { Get-FxcExisting "me/calendars/$($f.Id)/events?`$select=id,subject&`$top=100" 'subject' }
                'Contact' { Get-FxcExisting "me/contactFolders/$($f.Id)/contacts?`$select=id,displayName&`$top=100" 'displayName' }
            }
        } catch {
            # no folder, no items of this kind; the other kinds go on
            Write-Fxc FAIL "${kind}: $($_.Exception.Message)"
            foreach ($s in $steps) { $results.Add([pscustomobject]@{ Key = $s.Key; Kind = $kind; Status = 'Failed'; Id = $null; Notes = @("folder of the kind not available: $($_.Exception.Message)"); Title = $s.Title }) }
            continue
        }
        foreach ($s in $steps) {
            if ($existing.ContainsKey($s.Key)) {
                $r = [pscustomobject]@{ Key = $s.Key; Kind = $kind; Status = 'Skipped'; Id = $existing[$s.Key]; Notes = [Collections.Generic.List[string]]::new(); Title = $s.Title }
                if ($s.Special -eq 'RecurringExceptions') { Complete-FxcRecurring -Result $r -EventId $existing[$s.Key] }   # only does what is missing
                $results.Add($r); Write-Fxc INFO "$($s.Key) already there: skipped"
                continue
            }
            try { $r = New-FxcItem -Step $s -FolderId $f.Id }
            catch { $r = [pscustomobject]@{ Key = $s.Key; Kind = $kind; Status = 'Failed'; Id = $null; Notes = @($_.Exception.Message); Title = $s.Title } }
            $results.Add($r)
            Write-Fxc $(switch ($r.Status) { 'Created' { 'OK' } 'Partial' { 'WARN' } default { 'FAIL' } }) ("{0} {1}: {2}" -f $r.Key, $r.Status, $s.Title)
            foreach ($n in $r.Notes) { Write-Host "        $n" -ForegroundColor DarkGray }
        }
    }

    # paths of the folders as the export tree names them (Graph tells for mail and contacts, not for a secondary calendar)
    foreach ($kind in @($folders.Keys)) { $paths[$kind] = Resolve-FxcFolderPath -Id $folders[$kind] -ContactFolder $(if ($kind -eq 'Contact') { 'yes' }) }
} catch {
    $runError = $_.Exception.Message
    Write-Fxc FAIL "the run stopped: $runError"
} finally {
    $script:Http.Dispose()
}

# The manifest is written whatever happened above (failed items, a run that stopped): it holds what exists in the mailbox.
$manifest = New-FxcManifest -User $UserPrincipalName -Root $RootName -Results $results.ToArray() -Folders $folders -Paths $paths -FolderNames $folderNames -RunError $runError
$dir = Split-Path -Parent $OutputPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) { $null = New-Item -ItemType Directory -Path $dir }
$manifest | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding utf8

Write-Fxc STEP 'Result'
if ($results.Count) { Write-Host (($results | Select-Object Key, Kind, Status, @{ n = 'Notes'; e = { @($_.Notes).Count } }, Title | Format-Table -AutoSize | Out-String -Width 220).TrimEnd()) }
foreach ($g in $results | Group-Object Kind) {
    Write-Fxc INFO ("{0}: {1}" -f $g.Name, ((@($g.Group | Group-Object Status) | ForEach-Object { "$($_.Count) $($_.Name)" }) -join ', '))
}
Write-Fxc STEP 'Folders to export (the -Folder of Invoke-MailboxTransfer.ps1 -Mode Export, one run per folder)'
foreach ($kind in $manifest.Folders.Keys) {
    $fr = $manifest.Folders[$kind]
    if ($fr.PathResolved) { Write-Fxc OK ("{0,-8} {1}" -f $kind, $fr.ExportPath) }
    else {
        Write-Fxc WARN ("{0,-8} path NOT resolved by Graph. Name '{1}', Graph id {2}" -f $kind, $fr.Name, $fr.GraphId)
        Write-Host "         $($fr.HowToFindPath)" -ForegroundColor DarkGray
    }
}
Write-Fxc INFO "manifest: $OutputPath"
if ($runError -or ($results | Where-Object Status -eq 'Failed')) { exit 1 }

#endregion
