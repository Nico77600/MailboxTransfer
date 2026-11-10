#Requires -Version 7.4
#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    Lab only: gives the dedicated application the delegated Microsoft Graph permissions of the mailbox import and export
    API, a public redirect URI, and a consent granted to ONE user (-UserPrincipalName, adele.vance@contoso.com by default) in the lab tenant.

.DESCRIPTION
    Run it once, by an account that may manage applications and consents in the tenant (Cloud Application Administrator
    or Global Administrator). It signs in with the Microsoft Graph PowerShell module, interactively.

    What it does, and only this, on the dedicated application:
      1. Adds to the application (requiredResourceAccess, Microsoft Graph, type Scope) the permissions of -Permissions that
         it does not have yet: MailboxFolder.Read, MailboxFolder.ReadWrite, MailboxItem.Read, MailboxItem.ImportExport.
         Names are checked against the delegated permissions published by Microsoft Graph in the tenant.
      2. Adds http://localhost to the redirect URIs of the "Mobile and desktop applications" platform (public client),
         if it is not there yet.
      3. Creates or completes ONE oauth2PermissionGrant: consentType Principal, for one user only (-UserPrincipalName,
         resolved to its object id with GET /users/{upn}, or -PrincipalObjectId), resource Microsoft Graph, with the permissions above (and openid, profile, offline_access, so that the
         user is not asked anything and the refresh token works; -NoOidcScopes to leave them out).
      4. Turns on "Allow public client flows" (isFallbackPublicClient) if it is off: the device code flow of
         Invoke-Phase0.ps1 -DeviceCode needs it (error AADSTS7000218 otherwise). -NoPublicClientFlows to leave it alone.
    Nothing is removed or replaced: existing permissions, redirect URIs, scopes and grants stay. No other application, no
    AllPrincipals grant, no secret, no certificate, no application permission is touched. Running it again changes nothing.

    The state of the application is shown before and after.

.PARAMETER UserPrincipalName
    The user who consents (default: Lab.UserPrincipalName of package\config\MailboxTransfer.config.psd1, adele.vance@contoso.com). Resolved to
    an object id with GET /users/{upn} (delegated User.ReadBasic.All). If that call is refused and the user is the one of the
    configuration, the object id of the configuration (Lab.UserObjectId) is used.

.PARAMETER PrincipalObjectId
    Object id of the user, when it is already known (nothing is looked up).

.PARAMETER DeviceCode
    Signs the administrator in with a device code (MSAL.NET of the module, same engine as Invoke-Phase0.ps1) instead of
    Connect-MgGraph, whose own device code flow stops waiting after 120 seconds. The line "DEVICE CODE: <code> URL: <url>" is
    written to the console at once and to out\device-code.txt; the script waits until the code expires (about 15 minutes).
    The token is then given to Connect-MgGraph -AccessToken. The administrator's token cache is kept apart
    (%LOCALAPPDATA%\MailboxTransfer\token-admin.cache).

.PARAMETER NoPublicClientFlows
    Does not turn on "Allow public client flows" on the application (see 4. below).

.PARAMETER WhatIf
    Shows the state and what would be changed, changes nothing.

.EXAMPLE
    .\tools\Grant-LabConsent.ps1 -WhatIf
    .\tools\Grant-LabConsent.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0 (phase 0)
    Needs the delegated permissions Application.ReadWrite.All, DelegatedPermissionGrant.ReadWrite.All and
    User.ReadBasic.All for the Microsoft Graph PowerShell application (the admin is asked to consent to them the first time).
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$UserPrincipalName,
    [string]$PrincipalObjectId,
    [string[]]$Permissions = @('MailboxFolder.Read', 'MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'),
    [string]$RedirectUri = 'http://localhost',
    [switch]$NoOidcScopes,
    [switch]$NoPublicClientFlows,
    [switch]$DeviceCode
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) 'package\MailboxTransfer.psd1')
$config = Import-MbtConfiguration
if (-not $TenantId) { $TenantId = $config.Lab.TenantId }
if (-not $ClientId) { $ClientId = $config.Lab.ClientId }
if (-not $UserPrincipalName) { $UserPrincipalName = $config.Lab.UserPrincipalName }
$graphAppId = '00000003-0000-0000-c000-000000000000'
# Application "Microsoft Graph Command Line Tools": the public client the administrator signs in with (Connect-MgGraph's default).
$adminClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$oidcScopes = if ($NoOidcScopes) { @() } else { @('openid', 'profile', 'offline_access') }

function Invoke-Graph {
    param([Parameter(Mandatory)][string]$Method, [Parameter(Mandatory)][string]$Uri, $Body)
    $params = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0$Uri"; OutputType = 'PSObject' }
    if ($null -ne $Body) { $params.Body = ($Body | ConvertTo-Json -Depth 10); $params.ContentType = 'application/json' }
    Invoke-MgGraphRequest @params
}

function Test-PublicClientFlows {
    <# "Allow public client flows" of an application (isFallbackPublicClient). #>
    param([Parameter(Mandatory)]$App)
    return [bool]($App.PSObject.Properties['isFallbackPublicClient'] -and $App.isFallbackPublicClient)
}

function Get-State {
    <# Reads everything the script may change: application, service principal, Graph scopes, grants. #>
    $app = @((Invoke-Graph GET "/applications?`$filter=appId eq '$ClientId'&`$select=id,appId,displayName,signInAudience,requiredResourceAccess,publicClient,isFallbackPublicClient").value)
    if ($app.Count -ne 1) { throw "Application $ClientId not found (or not unique) in tenant $TenantId." }
    $sp = @((Invoke-Graph GET "/servicePrincipals?`$filter=appId eq '$ClientId'&`$select=id,appId,displayName").value)
    $graph = @((Invoke-Graph GET "/servicePrincipals?`$filter=appId eq '$graphAppId'&`$select=id,displayName,oauth2PermissionScopes").value)[0]
    $grants = @()
    if ($sp.Count -eq 1) { $grants = @((Invoke-Graph GET "/oauth2PermissionGrants?`$filter=clientId eq '$($sp[0].id)'").value) }
    [pscustomobject]@{ App = $app[0]; Sp = if ($sp.Count) { $sp[0] } else { $null }; Graph = $graph; Grants = $grants }
}

function Show-State {
    param([Parameter(Mandatory)]$State, [Parameter(Mandatory)][string]$Title)
    $names = @{}
    foreach ($s in $State.Graph.oauth2PermissionScopes) { $names[$s.id] = $s.value }
    Write-Host ''
    Write-Host "== $Title" -ForegroundColor Cyan
    Write-Host ("  Application   : {0} ({1}), object id {2}" -f $State.App.displayName, $State.App.appId, $State.App.id)
    Write-Host ("  Service principal : {0}" -f $(if ($State.Sp) { $State.Sp.id } else { '(none in this tenant)' }))
    $own = @($State.App.requiredResourceAccess | Where-Object { $_.resourceAppId -eq $graphAppId })
    $declared = @($own | ForEach-Object { $_.resourceAccess } | ForEach-Object { if ($_.type -eq 'Scope') { "$($names[$_.id])" } else { "Role:$($_.id)" } } | Sort-Object)
    Write-Host ("  Declared Microsoft Graph permissions ({0}): {1}" -f $declared.Count, ($declared -join ', '))
    $others = @($State.App.requiredResourceAccess | Where-Object { $_.resourceAppId -ne $graphAppId })
    if ($others.Count) { Write-Host ("  Other resources declared (untouched): {0}" -f (($others | ForEach-Object { $_.resourceAppId }) -join ', ')) }
    Write-Host ("  Redirect URIs (public client): {0}" -f $(if ($State.App.publicClient.redirectUris) { @($State.App.publicClient.redirectUris) -join ', ' } else { '(none)' }))
    Write-Host ("  Allow public client flows: {0}" -f $(if (Test-PublicClientFlows $State.App) { 'yes' } else { 'no' }))
    if (-not $State.Grants.Count) { Write-Host '  Consents (oauth2PermissionGrants): none' }
    foreach ($g in $State.Grants) {
        $who = if ($g.consentType -eq 'Principal') { "Principal $($g.principalId)" } else { 'AllPrincipals' }
        $res = if ($g.resourceId -eq $State.Graph.id) { 'Microsoft Graph' } else { "resource $($g.resourceId)" }
        Write-Host ("  Consent {0}: {1} on {2}: {3}" -f $g.id, $who, $res, $g.scope.Trim())
    }
}

# ---- sign in ---------------------------------------------------------------------------------------
$needed = 'Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All', 'User.ReadBasic.All'
$context = Get-MgContext
$adminAccount = $null
if (-not $context -or $context.TenantId -ne $TenantId -or ($needed | Where-Object { $_ -notin $context.Scopes })) {
    if ($DeviceCode) {
        $scopes = @($needed | ForEach-Object { "https://graph.microsoft.com/$_" })
        $token = Get-MbtAccessToken -TenantId $TenantId -ClientId $adminClientId -Scopes $scopes -DeviceCode -DeviceCodeFile (Join-Path (Split-Path $PSScriptRoot -Parent) 'out\device-code.txt')
        $adminAccount = $token.Account
        $null = Connect-MgGraph -AccessToken (ConvertTo-SecureString $token.AccessToken -AsPlainText -Force) -NoWelcome
    } else {
        $null = Connect-MgGraph -TenantId $TenantId -Scopes $needed -NoWelcome
    }
    $context = Get-MgContext
}
Write-Host ("Signed in as {0} in tenant {1}" -f $(if ($adminAccount) { $adminAccount } else { $context.Account }), $TenantId)

# ---- the user who consents -------------------------------------------------------------------------
if (-not $PrincipalObjectId) {
    try {
        $PrincipalObjectId = (Invoke-Graph GET "/users/$([uri]::EscapeDataString($UserPrincipalName))?`$select=id,userPrincipalName").id
    } catch {
        if ($UserPrincipalName -ieq $config.Lab.UserPrincipalName -and $config.Lab.UserObjectId) {
            Write-Host ("  WARNING: cannot read {0} ({1}); using Lab.UserObjectId of the configuration." -f $UserPrincipalName, $_.Exception.Message) -ForegroundColor Yellow
            $PrincipalObjectId = $config.Lab.UserObjectId
        } else {
            throw "Cannot resolve $UserPrincipalName to an object id ($($_.Exception.Message)). Give -PrincipalObjectId."
        }
    }
}
Write-Host ("User who consents: {0}, object id {1}" -f $UserPrincipalName, $PrincipalObjectId)

$before = Get-State
Show-State $before 'State BEFORE'

# ---- what is missing -------------------------------------------------------------------------------
$scopeIds = @{}
foreach ($s in $before.Graph.oauth2PermissionScopes) { $scopeIds[$s.value] = $s.id }
$unknown = @($Permissions + $oidcScopes | Where-Object { -not $scopeIds.ContainsKey($_) })
if ($unknown) {
    $similar = ($scopeIds.Keys | Where-Object { $_ -like 'Mailbox*' } | Sort-Object) -join ', '
    throw "These delegated permissions do not exist on Microsoft Graph in this tenant: $($unknown -join ', '). Existing Mailbox* permissions: $similar"
}

# 1. requiredResourceAccess: all existing entries are sent back unchanged, plus the missing permissions
$resources = [System.Collections.Generic.List[object]]::new()
$graphEntry = $null
foreach ($r in @($before.App.requiredResourceAccess)) {
    $entry = [ordered]@{ resourceAppId = $r.resourceAppId; resourceAccess = @($r.resourceAccess | ForEach-Object { [ordered]@{ id = $_.id; type = $_.type } }) }
    if ($r.resourceAppId -eq $graphAppId) { $graphEntry = $entry }
    $resources.Add($entry)
}
if (-not $graphEntry) {
    $graphEntry = [ordered]@{ resourceAppId = $graphAppId; resourceAccess = @() }
    $resources.Add($graphEntry)
}
$addedPermissions = @()
foreach ($p in $Permissions) {
    $id = $scopeIds[$p]
    if (-not ($graphEntry.resourceAccess | Where-Object { $_.id -eq $id -and $_.type -eq 'Scope' })) {
        $graphEntry.resourceAccess = @($graphEntry.resourceAccess) + [ordered]@{ id = $id; type = 'Scope' }
        $addedPermissions += $p
    }
}

# 2. redirect URI
$redirects = @($before.App.publicClient.redirectUris | Where-Object { $_ })
$addRedirect = $RedirectUri -notin $redirects

$patch = [ordered]@{}
if (-not $NoPublicClientFlows -and -not (Test-PublicClientFlows $before.App)) { $patch.isFallbackPublicClient = $true }
if ($addedPermissions) { $patch.requiredResourceAccess = @($resources) }
if ($addRedirect) { $patch.publicClient = [ordered]@{ redirectUris = @($redirects + $RedirectUri) } }

Write-Host ''
Write-Host '== Changes' -ForegroundColor Cyan
if ($patch.Count) {
    if ($addedPermissions) { Write-Host ("  + permissions to declare: {0}" -f ($addedPermissions -join ', ')) }
    if ($addRedirect) { Write-Host ("  + redirect URI: {0}" -f $RedirectUri) }
    if ($patch.Contains('isFallbackPublicClient')) { Write-Host '  + Allow public client flows (needed by the device code flow)' }
    if ($PSCmdlet.ShouldProcess("application $ClientId", "Update requiredResourceAccess / redirect URIs (additions only)")) {
        $null = Invoke-Graph PATCH "/applications/$($before.App.id)" $patch
        Write-Host '  application updated' -ForegroundColor Green
    }
} else {
    Write-Host '  application: nothing to change'
}

# 3. service principal and grant (Principal, one user)
$sp = $before.Sp
if (-not $sp) {
    Write-Host "  + service principal of the application (it does not exist in this tenant yet; needed for the consent)"
    if ($PSCmdlet.ShouldProcess("application $ClientId", 'Create its service principal')) {
        $sp = Invoke-Graph POST '/servicePrincipals' @{ appId = $ClientId }
        Write-Host '  service principal created' -ForegroundColor Green
    }
}
$wantedScopes = @($Permissions + $oidcScopes)
if ($sp) {
    $mine = @($before.Grants | Where-Object { $_.consentType -eq 'Principal' -and $_.principalId -eq $PrincipalObjectId -and $_.resourceId -eq $before.Graph.id })
    if ($mine.Count -gt 1) { throw "Several Principal grants exist for $PrincipalObjectId on Microsoft Graph; remove the duplicates by hand and run again." }
    if ($mine.Count -eq 1) {
        $have = @($mine[0].scope -split '\s+' | Where-Object { $_ })
        $missing = @($wantedScopes | Where-Object { $_ -notin $have })
        if ($missing) {
            Write-Host ("  + grant {0}: adding {1}" -f $mine[0].id, ($missing -join ', '))
            if ($PSCmdlet.ShouldProcess("consent $($mine[0].id) of user $PrincipalObjectId", "Add scopes: $($missing -join ' ')")) {
                $null = Invoke-Graph PATCH "/oauth2PermissionGrants/$($mine[0].id)" @{ scope = (@($have + $missing) -join ' ') }
                Write-Host '  consent completed' -ForegroundColor Green
            }
        } else { Write-Host '  consent of the user: nothing to change' }
    } else {
        Write-Host ("  + new consent for user {0}: {1}" -f $PrincipalObjectId, ($wantedScopes -join ' '))
        if ($PSCmdlet.ShouldProcess("user $PrincipalObjectId", "Create a Principal consent: $($wantedScopes -join ' ')")) {
            $null = Invoke-Graph POST '/oauth2PermissionGrants' @{
                clientId = $sp.id; consentType = 'Principal'; principalId = $PrincipalObjectId; resourceId = $before.Graph.id; scope = ($wantedScopes -join ' ')
            }
            Write-Host '  consent created' -ForegroundColor Green
        }
    }
}

$after = Get-State
Show-State $after $(if ($WhatIfPreference) { 'State AFTER (-WhatIf: nothing was changed)' } else { 'State AFTER' })

# ---- verdict ---------------------------------------------------------------------------------------
$ok = $true
$declaredIds = @($after.App.requiredResourceAccess | Where-Object { $_.resourceAppId -eq $graphAppId } | ForEach-Object { $_.resourceAccess } | ForEach-Object { $_.id })
foreach ($p in $Permissions) { if ($scopeIds[$p] -notin $declaredIds) { $ok = $false; Write-Host "  MISSING in the application: $p" -ForegroundColor Yellow } }
if ($RedirectUri -notin @($after.App.publicClient.redirectUris)) { $ok = $false; Write-Host "  MISSING redirect URI: $RedirectUri" -ForegroundColor Yellow }
if (-not $NoPublicClientFlows -and -not (Test-PublicClientFlows $after.App)) { $ok = $false; Write-Host '  MISSING: Allow public client flows is off' -ForegroundColor Yellow }
$grant = @($after.Grants | Where-Object { $_.consentType -eq 'Principal' -and $_.principalId -eq $PrincipalObjectId -and $_.resourceId -eq $after.Graph.id })
if ($grant.Count -ne 1) { $ok = $false; Write-Host "  MISSING consent for user $PrincipalObjectId" -ForegroundColor Yellow }
else { foreach ($p in $Permissions) { if ($p -notin @($grant[0].scope -split '\s+')) { $ok = $false; Write-Host "  MISSING in the consent: $p" -ForegroundColor Yellow } } }
Write-Host ''
if ($ok) { Write-Host 'OK: the application declares the permissions, has the redirect URI, and the user has consented.' -ForegroundColor Green }
elseif ($WhatIfPreference) { Write-Host 'WhatIf: the missing items above would be added.' -ForegroundColor Yellow }
else { Write-Host 'Something is still missing, see above.' -ForegroundColor Red; exit 1 }
