#Requires -Version 7.4
#Requires -Modules Microsoft.Graph.Authentication
<#
.SYNOPSIS
    OPTIONAL lab setup for the app-only sign-in (docs\MailboxTransfer-Guide.md, chapter 4.6): a dedicated application registration, a self-signed
    certificate in CurrentUser\My whose private key cannot be exported, its public key uploaded to the application, and the application
    permissions requested. It prints the steps the administrator takes next. Nothing in the product runs it.

.DESCRIPTION
    With -WhatIf it works OFFLINE: no sign-in, no certificate, nothing created; it prints what it would do and the steps after it.
    Without -WhatIf it asks to confirm each change (ConfirmImpact High) and does only this:
      1. Signs in to the tenant with the Microsoft Graph PowerShell module (delegated Application.ReadWrite.All, interactive).
      2. Looks the application permissions up by name in the Microsoft Graph service principal of the tenant (a name that does not exist stops it).
      3. Creates the self-signed certificate in Cert:\CurrentUser\My: RSA 2048, SHA-256, key NOT exportable, valid -CertificateMonths.
      4. Creates ONE application registration (-DisplayName, this tenant only) with the public key of the certificate and the application
         permissions in its requiredResourceAccess, and its service principal. It stops if an application with that name exists.
    It does NOT grant admin consent (that needs another permission and a human decision), does NOT create a secret, does NOT touch another
    application, and never modifies or deletes a mailbox, a user or a group.

    The requested permissions are those of docs\MailboxTransfer-Guide.md, chapter 4.6 for browse, export and import (Get-MbtAppPermission), plus
    -IncludeUserRead for User.Read.All (find a mailbox id from an address; reads every user: leave it out and give mailbox ids).

    Afterwards it prints the two decisions that stay with the administrator: how the permissions are granted (admin consent in Entra =
    every mailbox of the tenant, OR Exchange RBAC for Applications with a management scope = chosen mailboxes only, recommended), with
    the exact commands, and how to test the sign-in with Invoke-MailboxTransfer.ps1.

.PARAMETER TenantId
    The tenant (id or domain). Required: there is no lab default.

.PARAMETER DisplayName
    Name of the application registration (default 'Mailbox Transfer (app-only)').

.PARAMETER CertificateSubject
    Subject of the certificate (default 'CN=MailboxTransfer-AppOnly').

.PARAMETER CertificateMonths
    Validity of the certificate in months (1 to 24, default 12).

.PARAMETER Permissions
    Application permissions of Microsoft Graph to request (default: MailboxFolder.ReadWrite.All, MailboxItem.Read.All, MailboxItem.ImportExport.All, Mail.Read).

.PARAMETER IncludeUserRead
    Also requests User.Read.All.

.PARAMETER ScopeName
    Name of the management scope shown in the RBAC steps (default MbtMailboxes).

.PARAMETER ScopeFilter
    Recipient filter shown in the RBAC steps (default "CustomAttribute1 -eq 'MailboxTransfer'").

.EXAMPLE
    .\tools\New-MbtAppOnlyLab.ps1 -TenantId contoso.onmicrosoft.com -WhatIf
    .\tools\New-MbtAppOnlyLab.ps1 -TenantId contoso.onmicrosoft.com

.NOTES
    Author  : Nicolas Fabert
    Not run by the product or by the tests with the tenant: the tests only run it with -WhatIf.
#>
[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)][string]$TenantId,
    [string]$DisplayName = 'Mailbox Transfer (app-only)',
    [string]$CertificateSubject = 'CN=MailboxTransfer-AppOnly',
    [ValidateRange(1, 24)][int]$CertificateMonths = 12,
    [string[]]$Permissions = @('MailboxFolder.ReadWrite.All', 'MailboxItem.Read.All', 'MailboxItem.ImportExport.All', 'Mail.Read'),
    [switch]$IncludeUserRead,
    [string]$ScopeName = 'MbtMailboxes',
    [string]$ScopeFilter = "CustomAttribute1 -eq 'MailboxTransfer'"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$graphAppId = '00000003-0000-0000-c000-000000000000'
if ($IncludeUserRead -and 'User.Read.All' -notin $Permissions) { $Permissions += 'User.Read.All' }

# Exchange application role of each Graph permission that has one (docs\MailboxTransfer-Guide.md, chapter 4.6)
$exchangeRole = @{
    'MailboxFolder.Read.All' = 'Application MailboxFolder.Read'; 'MailboxFolder.ReadWrite.All' = 'Application MailboxFolder.ReadWrite'
    'MailboxItem.Read.All' = 'Application MailboxItem.Read'; 'MailboxItem.ReadWrite.All' = 'Application MailboxItem.ReadWrite'
    'MailboxItem.Export.All' = 'Application MailboxItem.Export'; 'MailboxItem.ImportExport.All' = 'Application MailboxItem.ImportExport'
    'Mail.Read' = 'Application Mail.Read'; 'Mail.ReadWrite' = 'Application Mail.ReadWrite'
}

function Write-NextStep {
    param([string]$AppId = '<application (client) id>', [string]$SpObjectId = '<service principal object id>', [string]$Thumbprint = '<certificate thumbprint>')
    $roles = @($Permissions | Where-Object { $exchangeRole.ContainsKey($_) } | ForEach-Object { "'$($exchangeRole[$_])'" })
    Write-Host ''
    Write-Host '== Steps for the administrator (nothing below has been done)' -ForegroundColor Cyan
    Write-Host ''
    Write-Host 'A. Recommended: only the chosen mailboxes, Exchange RBAC for Applications. Do NOT grant tenant-wide admin consent to these permissions (the effective access is the union of both).'
    Write-Host '   Connect-ExchangeOnline'
    Write-Host "   New-ServicePrincipal -AppId $AppId -ObjectId $SpObjectId -DisplayName '$DisplayName'"
    Write-Host "   New-ManagementScope -Name '$ScopeName' -RecipientRestrictionFilter `"$ScopeFilter`""
    Write-Host "   foreach (`$role in $($roles -join ', ')) { New-ManagementRoleAssignment -App $SpObjectId -Role `$role -CustomResourceScope '$ScopeName' }"
    Write-Host "   Test-ServicePrincipalAuthorization -Identity $AppId -Resource <mailbox address> | Format-Table     # InScope must be True; allow 30 minutes to 2 hours"
    if ('User.Read.All' -in $Permissions) { Write-Host '   (User.Read.All has no Exchange role: it needs admin consent in Entra and reads every user.)' }
    Write-Host ''
    Write-Host 'B. Or the whole tenant: admin consent in Entra (Cloud Application Administrator or Global Administrator):'
    Write-Host "   https://login.microsoftonline.com/$TenantId/adminconsent?client_id=$AppId"
    Write-Host '   (or Entra admin center > Enterprise applications > the application > Permissions > Grant admin consent)'
    Write-Host ''
    Write-Host 'C. Test, from a PC that has the certificate (CurrentUser\My):'
    Write-Host "   .\Invoke-MailboxTransfer.ps1 -Mode Browse -CertificateThumbprint $Thumbprint -TenantId $TenantId -ClientId $AppId -UserPrincipalName <mailbox address> -Folder Inbox"
    Write-Host ''
    Write-Host 'D. When the lab is over: remove the application registration (Entra) and the certificate (Remove-Item Cert:\CurrentUser\My\<thumbprint>).'
}

Write-Host ("Tenant                : {0}" -f $TenantId)
Write-Host ("Application           : {0} (this tenant only)" -f $DisplayName)
Write-Host ("Certificate           : {0}, RSA 2048 / SHA-256, valid {1} month(s), private key NOT exportable, Cert:\CurrentUser\My" -f $CertificateSubject, $CertificateMonths)
Write-Host ("Application permissions requested: {0}" -f ($Permissions -join ', '))
Write-Host 'Admin consent         : NOT granted by this script'

if ($WhatIfPreference) {
    Write-Host ''
    Write-Host 'WhatIf: no sign-in, no certificate, no application was created.' -ForegroundColor Yellow
    Write-NextStep
    return
}

if (-not $PSCmdlet.ShouldProcess("tenant $TenantId", "Create the application '$DisplayName', the certificate '$CertificateSubject' and request $($Permissions.Count) application permission(s) (no admin consent)")) { return }

$null = Connect-MgGraph -TenantId $TenantId -Scopes 'Application.ReadWrite.All' -NoWelcome
function Invoke-Graph {
    param([Parameter(Mandatory)][string]$Method, [Parameter(Mandatory)][string]$Uri, $Body)
    $params = @{ Method = $Method; Uri = "https://graph.microsoft.com/v1.0$Uri"; OutputType = 'PSObject' }
    if ($null -ne $Body) { $params.Body = ($Body | ConvertTo-Json -Depth 10); $params.ContentType = 'application/json' }
    Invoke-MgGraphRequest @params
}

# 2. the permissions by name -> role ids
$graphSp = @((Invoke-Graph GET "/servicePrincipals?`$filter=appId eq '$graphAppId'&`$select=id,appRoles").value)[0]
$roleIds = @{}
foreach ($r in $graphSp.appRoles) { if ($r.allowedMemberTypes -contains 'Application') { $roleIds[$r.value] = $r.id } }
$unknown = @($Permissions | Where-Object { -not $roleIds.ContainsKey($_) })
if ($unknown) { throw "These application permissions do not exist on Microsoft Graph in this tenant: $($unknown -join ', ')." }

# the application must not exist yet (nothing existing is modified)
$existing = @((Invoke-Graph GET "/applications?`$filter=displayName eq '$($DisplayName.Replace("'", "''"))'&`$select=id,appId").value)
if ($existing.Count) { throw "An application named '$DisplayName' already exists ($($existing[0].appId)): nothing was created. Choose another -DisplayName." }

# 3. the certificate: the private key stays in the key store of the user, not exportable
$notAfter = (Get-Date).AddMonths($CertificateMonths)
if (-not $PSCmdlet.ShouldProcess('Cert:\CurrentUser\My', "Create the self-signed certificate $CertificateSubject")) { return }
$cert = New-SelfSignedCertificate -Subject $CertificateSubject -CertStoreLocation 'Cert:\CurrentUser\My' -KeyExportPolicy NonExportable `
    -KeyUsage DigitalSignature -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256 -NotAfter $notAfter -Provider 'Microsoft Software Key Storage Provider'
Write-Host ("Certificate created   : {0} (expires {1:yyyy-MM-dd})" -f $cert.Thumbprint, $cert.NotAfter)

# 4. the application and its service principal
$body = [ordered]@{
    displayName            = $DisplayName
    signInAudience         = 'AzureADMyOrg'
    requiredResourceAccess = @([ordered]@{ resourceAppId = $graphAppId; resourceAccess = @($Permissions | ForEach-Object { [ordered]@{ id = $roleIds[$_]; type = 'Role' } }) })
    keyCredentials         = @([ordered]@{
            type = 'AsymmetricX509Cert'; usage = 'Verify'; displayName = $CertificateSubject
            key = [Convert]::ToBase64String($cert.RawData)
            startDateTime = $cert.NotBefore.ToUniversalTime().ToString('o'); endDateTime = $cert.NotAfter.ToUniversalTime().ToString('o')
        })
}
if (-not $PSCmdlet.ShouldProcess($TenantId, "Create the application registration '$DisplayName'")) { return }
$app = Invoke-Graph POST '/applications' $body
$sp = Invoke-Graph POST '/servicePrincipals' ([ordered]@{ appId = $app.appId })
Write-Host ("Application created   : {0} (client id {1})" -f $DisplayName, $app.appId)
Write-Host ("Service principal     : {0}" -f $sp.id)
Write-NextStep -AppId $app.appId -SpObjectId $sp.id -Thumbprint $cert.Thumbprint
[pscustomobject]@{ TenantId = $TenantId; ClientId = $app.appId; ServicePrincipalObjectId = $sp.id; CertificateThumbprint = $cert.Thumbprint; Permissions = $Permissions; AdminConsent = 'not granted' }
