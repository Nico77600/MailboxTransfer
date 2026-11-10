#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the app-only sign-in (client credentials with a certificate): the token source with a fake acquisition, the
    certificate lookup (thumbprint order CurrentUser then LocalMachine, .pfx with password), the claims of an application token, the
    permissions each mode needs, /users (never /me) in an app-only session, the entry script parameters and the capability.
    Nothing here reaches the network, the tenant or a mailbox; no certificate is added to a store.

    Run:  Invoke-Pester -Path .\tests\AppOnly.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    Initialize-MbtAuthEngine
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ('mbt-apponly-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp

    $script:Tenant = '11111111-2222-3333-4444-555555555555'
    $script:Client = '66666666-7777-8888-9999-000000000000'

    # A self-signed certificate in memory (never put in a store), optionally exported to a .pfx / .cer
    function New-TestCertificate {
        param([string]$Name = 'CN=mbt-apponly-test', [int]$NotBeforeDays = -1, [int]$NotAfterDays = 30)
        $rsa = [Security.Cryptography.RSA]::Create(2048)
        $req = [Security.Cryptography.X509Certificates.CertificateRequest]::new($Name, $rsa, [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1)
        return $req.CreateSelfSigned([DateTimeOffset]::UtcNow.AddDays($NotBeforeDays), [DateTimeOffset]::UtcNow.AddDays($NotAfterDays))
    }
    function ConvertTo-Base64Url([string]$Text) { [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
    function New-TestJwt([hashtable]$Claims) { (ConvertTo-Base64Url '{"alg":"RS256"}') + '.' + (ConvertTo-Base64Url ($Claims | ConvertTo-Json -Compress)) + '.sig' }
    function New-AppToken([string[]]$Roles) { New-TestJwt @{ aud = 'https://graph.microsoft.com'; tid = $script:Tenant; appid = $script:Client; oid = 'sp-oid'; idtyp = 'app'; roles = $Roles; exp = [DateTimeOffset]::UtcNow.AddHours(1).ToUnixTimeSeconds() } }

    # A GraphAppAuth with a fake token acquisition (the network is never reached)
    function New-FakeAppAuth {
        param([string[]]$Roles = @('MailboxItem.ImportExport.All'), [int]$ValidMinutes = 60)
        $global:AoCalls = [Collections.Generic.List[object]]::new()
        $jwt = New-AppToken $Roles; $minutes = $ValidMinutes
        $acquire = [Func[string[], Threading.CancellationToken, MailboxTransfer.AppToken]] {
            param($scopes, $ct)
            $global:AoCalls.Add($scopes -join ' ')
            $t = [MailboxTransfer.AppToken]::new(); $t.AccessToken = $jwt + ".$($global:AoCalls.Count)"; $t.ExpiresOn = [DateTimeOffset]::UtcNow.AddMinutes($minutes); $t
        }.GetNewClosure()
        return [MailboxTransfer.GraphAppAuth]::new($script:Client, $script:Tenant, [string[]]@('https://graph.microsoft.com/.default'), $acquire)
    }
}

AfterAll { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue; Remove-Variable -Name AoCalls, AoCert, AoAsked, AoSeen, AoAuth -Scope Global -ErrorAction SilentlyContinue }

Describe 'GraphAppAuth: the token of the application' {
    It 'asks for https://graph.microsoft.com/.default and has no account' {
        $a = New-FakeAppAuth
        $tok = $a.GetToken()
        $tok | Should -Match '\.1$'
        $global:AoCalls[0] | Should -Be 'https://graph.microsoft.com/.default'
        $a.Account | Should -BeNullOrEmpty
        $a.ClientId | Should -Be $script:Client
        $a.ExpiresOn | Should -Not -BeNullOrEmpty
    }
    It 'keeps the token in memory: the second call asks nothing, ForceRefresh asks again' {
        $a = New-FakeAppAuth
        $first = $a.GetToken(); $second = $a.GetToken()
        $second | Should -Be $first
        $a.Acquisitions | Should -Be 1
        $third = $a.GetToken([Threading.CancellationToken]::None, $true)
        $third | Should -Not -Be $first
        $a.Acquisitions | Should -Be 2
    }
    It 'asks a new token when the last one is about to expire (5 minutes)' {
        $a = New-FakeAppAuth -ValidMinutes 3
        $null = $a.GetToken(); $null = $a.GetToken()
        $a.Acquisitions | Should -Be 2
    }
    It 'is the token source of GraphClient (a C# delegate, no PowerShell runspace needed)' {
        $a = New-FakeAppAuth
        $f = $a.Provider()
        $f | Should -BeOfType ([Func[string]])
        $f.Invoke() | Should -Be $a.GetToken()
    }
    It 'SignOut forgets the token' {
        $a = New-FakeAppAuth
        $null = $a.GetToken(); $a.SignOut(); $null = $a.GetToken()
        $a.Acquisitions | Should -Be 2
    }
    It 'refuses a certificate without a private key' {
        $c = New-TestCertificate
        $pub = [Security.Cryptography.X509Certificates.X509Certificate2]::new($c.Export('Cert'))
        { [MailboxTransfer.GraphAppAuth]::new($script:Client, $script:Tenant, $pub, [string[]]@('https://graph.microsoft.com/.default')) } | Should -Throw '*no private key*'
    }
    It 'builds the MSAL confidential client from a certificate with its key (no request is made)' {
        $c = New-TestCertificate
        $a = [MailboxTransfer.GraphAppAuth]::new($script:Client, $script:Tenant, $c, [string[]]@('https://graph.microsoft.com/.default'))
        $a.CertificateThumbprint | Should -Be $c.Thumbprint
    }
}

Describe 'Get-MbtCertificate' {
    It 'opens a .pfx with its password (SecureString), key usable, nothing written to a store' {
        $c = New-TestCertificate
        $pfx = Join-Path $script:Tmp 'a.pfx'
        [IO.File]::WriteAllBytes($pfx, $c.Export('Pfx', 'S3cret!'))
        $got = Get-MbtCertificate -Path $pfx -Password (ConvertTo-SecureString 'S3cret!' -AsPlainText -Force)
        $got.Thumbprint | Should -Be $c.Thumbprint
        $got.HasPrivateKey | Should -BeTrue
        $rsa = [Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($got)
        $rsa.SignData([byte[]](1, 2, 3), [Security.Cryptography.HashAlgorithmName]::SHA256, [Security.Cryptography.RSASignaturePadding]::Pkcs1).Length | Should -Be 256
        $got.Dispose()
    }
    It 'opens a .pfx without a password' {
        $c = New-TestCertificate
        $pfx = Join-Path $script:Tmp 'nopw.pfx'
        [IO.File]::WriteAllBytes($pfx, $c.Export('Pfx'))
        (Get-MbtCertificate -Path $pfx).Thumbprint | Should -Be $c.Thumbprint
    }
    It 'says so for a wrong password, a missing file and a certificate without private key' {
        $c = New-TestCertificate
        $pfx = Join-Path $script:Tmp 'pw.pfx'
        [IO.File]::WriteAllBytes($pfx, $c.Export('Pfx', 'right'))
        { Get-MbtCertificate -Path $pfx -Password (ConvertTo-SecureString 'wrong' -AsPlainText -Force) } | Should -Throw '*Cannot open the certificate file*'
        { Get-MbtCertificate -Path (Join-Path $script:Tmp 'nope.pfx') } | Should -Throw '*not found*'
        $cer = Join-Path $script:Tmp 'pub.cer'
        [IO.File]::WriteAllBytes($cer, $c.Export('Cert'))
        { Get-MbtCertificate -Path $cer } | Should -Throw '*no private key*'
    }
    It 'refuses a certificate that has expired or is not valid yet' {
        $old = New-TestCertificate -NotBeforeDays -60 -NotAfterDays -2
        $future = New-TestCertificate -NotBeforeDays 5 -NotAfterDays 60
        $p1 = Join-Path $script:Tmp 'old.pfx'; $p2 = Join-Path $script:Tmp 'future.pfx'
        [IO.File]::WriteAllBytes($p1, $old.Export('Pfx')); [IO.File]::WriteAllBytes($p2, $future.Export('Pfx'))
        { Get-MbtCertificate -Path $p1 } | Should -Throw '*expired*'
        { Get-MbtCertificate -Path $p2 } | Should -Throw '*not valid before*'
    }
    Context 'by thumbprint' {
        It 'looks in CurrentUser\My first, then LocalMachine\My, and stops at the first hit' {
            $c = New-TestCertificate
            $global:AoCert = $c
            $global:AoAsked = [Collections.Generic.List[string]]::new()
            Mock -ModuleName MailboxTransfer Find-MbtStoreCertificate {
                $global:AoAsked.Add($StoreLocation)
                if ($StoreLocation -eq 'LocalMachine') { return $global:AoCert }
                return $null
            }
            $got = Get-MbtCertificate -Thumbprint $c.Thumbprint
            $got.Thumbprint | Should -Be $c.Thumbprint
            $global:AoAsked | Should -Be @('CurrentUser', 'LocalMachine')
            $global:AoAsked.Clear()
            Mock -ModuleName MailboxTransfer Find-MbtStoreCertificate { $global:AoAsked.Add($StoreLocation); return $global:AoCert }
            $null = Get-MbtCertificate -Thumbprint $c.Thumbprint
            $global:AoAsked | Should -Be @('CurrentUser')
        }
        It 'accepts a thumbprint as copied from the certificate console (spaces, invisible character, lower case)' {
            $c = New-TestCertificate
            $global:AoSeen = $null; $global:AoCert = $c
            Mock -ModuleName MailboxTransfer Find-MbtStoreCertificate { $global:AoSeen = $Thumbprint; return $global:AoCert }
            $messy = ([char]0x200E) + ($c.Thumbprint.ToLowerInvariant() -replace '(.{4})', '$1 ')
            $null = Get-MbtCertificate -Thumbprint $messy
            $global:AoSeen | Should -Be $c.Thumbprint.ToLowerInvariant()
        }
        It 'says where it looked when the certificate is in neither store, and refuses a malformed thumbprint' {
            Mock -ModuleName MailboxTransfer Find-MbtStoreCertificate { $null }
            { Get-MbtCertificate -Thumbprint ('AB' * 20) } | Should -Throw '*CurrentUser\My or LocalMachine\My*'
            { Get-MbtCertificate -Thumbprint 'ABC' } | Should -Throw '*not a certificate thumbprint*'
        }
        It 'refuses a certificate found without its private key' {
            $c = New-TestCertificate
            $pub = [Security.Cryptography.X509Certificates.X509Certificate2]::new($c.Export('Cert'))
            $global:AoCert = $pub
            Mock -ModuleName MailboxTransfer Find-MbtStoreCertificate { $global:AoCert }
            { Get-MbtCertificate -Thumbprint $c.Thumbprint } | Should -Throw '*no private key*'
        }
        It 'the real stores answer (read only): an unknown thumbprint is simply not found' {
            { Get-MbtCertificate -Thumbprint ('00' * 20) } | Should -Throw '*No certificate with the thumbprint*'
        }
    }
}

Describe 'Get-MbtTokenInfo: the claims of an application token' {
    It 'an app-only token: roles, no account, AppOnly; Scopes under the names of the delegated permissions' {
        $t = Get-MbtTokenInfo -AccessToken (New-AppToken @('MailboxItem.ImportExport.All', 'MailboxFolder.ReadWrite.All', 'Mail.Read'))
        $t.AppOnly | Should -BeTrue
        $t.Account | Should -BeNullOrEmpty
        $t.Roles | Should -Be @('MailboxItem.ImportExport.All', 'MailboxFolder.ReadWrite.All', 'Mail.Read')
        $t.Scopes | Should -Be @('MailboxItem.ImportExport', 'MailboxFolder.ReadWrite', 'Mail.Read')
        $t.AppId | Should -Be $script:Client
        $t.TenantId | Should -Be $script:Tenant
    }
    It 'an app-only token with no permission consented has no roles' {
        $t = Get-MbtTokenInfo -AccessToken (New-TestJwt @{ aud = 'https://graph.microsoft.com'; tid = $script:Tenant; appid = $script:Client; idtyp = 'app'; exp = [DateTimeOffset]::UtcNow.AddHours(1).ToUnixTimeSeconds() })
        $t.AppOnly | Should -BeTrue
        @($t.Roles).Count | Should -Be 0
    }
    It 'a delegated token is unchanged: scp, the account, not app-only' {
        $t = Get-MbtTokenInfo -AccessToken (New-TestJwt @{ aud = 'https://graph.microsoft.com'; upn = 'adele@contoso.com'; scp = 'MailboxItem.Read Mail.Read'; exp = [DateTimeOffset]::UtcNow.AddHours(1).ToUnixTimeSeconds() })
        $t.AppOnly | Should -BeFalse
        $t.Account | Should -Be 'adele@contoso.com'
        $t.Scopes | Should -Be @('MailboxItem.Read', 'Mail.Read')
        @($t.Roles).Count | Should -Be 0
    }
}

Describe 'Connect-MbtGraph: app-only' {
    BeforeAll {
        $c = New-TestCertificate
        $script:Pfx = Join-Path $script:Tmp 'connect.pfx'
        [IO.File]::WriteAllBytes($script:Pfx, $c.Export('Pfx', 'pw'))
        $script:Pw = ConvertTo-SecureString 'pw' -AsPlainText -Force
        $script:Cert = $c
    }
    BeforeEach {
        # the token acquisition is replaced: no network
        $global:AoAuth = New-FakeAppAuth -Roles @('MailboxItem.ImportExport.All', 'MailboxFolder.ReadWrite.All', 'MailboxItem.Read.All')
        $global:AoCert = $script:Cert
        Mock -ModuleName MailboxTransfer New-MbtAppAuth { $global:AoAuth }
    }
    AfterEach { Disconnect-MbtGraph }
    It 'with a .pfx: a session with no user, the roles, the target mailbox owner and the same GraphClient' {
        $s = Connect-MbtGraph -TenantId $script:Tenant -ClientId $script:Client -UserPrincipalName 'shared@contoso.com' -CertificatePath $script:Pfx -CertificatePassword $script:Pw
        $s.AppOnly | Should -BeTrue
        $s.UserPrincipalName | Should -Be 'shared@contoso.com'
        $s.Token.AppOnly | Should -BeTrue
        $s.Token.Account | Should -BeNullOrEmpty
        $s.Token.Roles | Should -Contain 'MailboxItem.ImportExport.All'
        $s.Graph | Should -BeOfType ([MailboxTransfer.GraphClient])
        $s.Api | Should -BeOfType ([MailboxTransfer.MailboxApi])
        (Get-MbtSession).AppOnly | Should -BeTrue
        Test-MbtAppOnly | Should -BeTrue
        Get-MbtSessionUser | Should -Be 'shared@contoso.com'
        Should -Invoke -ModuleName MailboxTransfer New-MbtAppAuth -Times 1 -Exactly -ParameterFilter { $Scopes -contains 'https://graph.microsoft.com/.default' -and $Certificate.Thumbprint -eq $global:AoCert.Thumbprint -and $ClientId -eq $global:AoAuth.ClientId }
    }
    It 'with a thumbprint: the certificate comes from the stores' {
        Mock -ModuleName MailboxTransfer Find-MbtStoreCertificate { $global:AoCert }
        $s = Connect-MbtGraph -TenantId $script:Tenant -ClientId $script:Client -UserPrincipalName 'adele@contoso.com' -CertificateThumbprint $script:Cert.Thumbprint
        $s.AppOnly | Should -BeTrue
        $s.Certificate.Thumbprint | Should -Be $script:Cert.Thumbprint
        Should -Invoke -ModuleName MailboxTransfer Find-MbtStoreCertificate -Times 1 -Exactly
    }
    It 'a delegated sign-in is still delegated (AppOnly false)' {
        Set-MbtSession -Graph ([MailboxTransfer.GraphClient]::new([Func[string]] { 't' }, [MailboxTransfer.GraphOptions]::new())) | Out-Null
        (Get-MbtSession).AppOnly | Should -BeFalse
        Test-MbtAppOnly | Should -BeFalse
    }
    It 'refuses: both thumbprint and file, -DeviceCode with a certificate, a general tenant name' {
        { Connect-MbtGraph -TenantId $script:Tenant -ClientId $script:Client -CertificateThumbprint ('AB' * 20) -CertificatePath $script:Pfx } | Should -Throw '*not both*'
        { Connect-MbtGraph -TenantId $script:Tenant -ClientId $script:Client -CertificatePath $script:Pfx -CertificatePassword $script:Pw -DeviceCode } | Should -Throw '*delegated*'
        { Connect-MbtGraph -TenantId 'common' -ClientId $script:Client -CertificatePath $script:Pfx -CertificatePassword $script:Pw } | Should -Throw '*not ''common''*'
    }
    It 'a wrong password stops before any token is asked' {
        { Connect-MbtGraph -TenantId $script:Tenant -ClientId $script:Client -CertificatePath $script:Pfx -CertificatePassword (ConvertTo-SecureString 'bad' -AsPlainText -Force) } | Should -Throw '*Cannot open the certificate file*'
        Should -Invoke -ModuleName MailboxTransfer New-MbtAppAuth -Times 0 -Exactly
    }
}

Describe 'App-only: every call names the user or the mailbox (no /me)' {
    BeforeAll {
        $script:Settings = '{"primaryMailboxId":"MBX:bbbbbbbb-1111-2222-3333-444444444444@11111111-2222-3333-4444-555555555555"}'
        function New-SessionWith([bool]$AppOnly, [string]$Account) {
            $route = [MailboxTransfer.Tests.RouteHandler]::new()
            $route.Add('GET', '/users/adele%40contoso.com/settings/exchange', 200, $script:Settings)
            $route.Add('GET', '/me/settings/exchange', 200, $script:Settings)
            $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
            $client = [MailboxTransfer.GraphClient]::new([Func[string]] { 'test-token' }, $o, $route)
            $s = Set-MbtSession -Graph $client
            $s.AppOnly = $AppOnly
            $s.UserPrincipalName = 'adele@contoso.com'
            $s.TenantId = $script:Tenant
            $s.Token = [pscustomobject]@{ Account = $Account; Scopes = @(); Roles = @(); AppOnly = $AppOnly }
            return $route
        }
    }
    It 'Get-MbtExchangeSettingsPath: only /users/{id} in an app-only session, even if a token had an account' {
        $null = New-SessionWith $true 'adele@contoso.com'
        Get-MbtExchangeSettingsPath -UserPrincipalName 'adele@contoso.com' | Should -Be @('/users/adele%40contoso.com/settings/exchange')
    }
    It 'Get-MbtExchangeSettingsPath: a delegated session of that very user also tries /me (unchanged)' {
        $null = New-SessionWith $false 'adele@contoso.com'
        Get-MbtExchangeSettingsPath -UserPrincipalName 'adele@contoso.com' | Should -Be @('/users/adele%40contoso.com/settings/exchange', '/me/settings/exchange')
        Get-MbtExchangeSettingsPath -UserPrincipalName 'other@contoso.com' | Should -Be @('/users/other%40contoso.com/settings/exchange')
    }
    It 'the primary mailbox id is read from /users/{upn}/settings/exchange and /me is never requested' {
        $route = New-SessionWith $true $null
        Get-MbtPrimaryMailboxId -UserPrincipalName 'adele@contoso.com' -TenantId $script:Tenant | Should -Be 'MBX:bbbbbbbb-1111-2222-3333-444444444444@11111111-2222-3333-4444-555555555555'
        @($route.Requests | Where-Object { $_.Url -match '/me[/?]' }).Count | Should -Be 0
        @($route.Requests | Where-Object { $_.Url -match '/users/adele%40contoso.com/settings/exchange' }).Count | Should -Be 1
    }
    It 'when /users answers 403 an app-only session does not fall back to /me' {
        $route = New-SessionWith $true $null
        $route2 = [MailboxTransfer.Tests.RouteHandler]::new()
        $route2.Add('GET', '/users/', 403, '{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}')
        $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
        $s = Set-MbtSession -Graph ([MailboxTransfer.GraphClient]::new([Func[string]] { 'test-token' }, $o, $route2))
        $s.AppOnly = $true; $s.Token = [pscustomobject]@{ Account = $null; Scopes = @(); Roles = @(); AppOnly = $true }
        { Get-MbtPrimaryMailboxId -UserPrincipalName 'adele@contoso.com' -TenantId $script:Tenant } | Should -Throw '*Give it with -PrimaryMailboxId*'
        @($route2.Requests | Where-Object { $_.Url -match '/me/' }).Count | Should -Be 0
        @($route2.Requests).Count | Should -Be 1
    }
    It 'a mailbox in the RBAC scope is found by its address without User.Read.All: settings/exchange 403, then the import-export API (parentMailboxUrl), never /me' {
        $route = [MailboxTransfer.Tests.RouteHandler]::new()
        $route.Add('GET', '/settings/exchange', 403, '{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}')
        $route.Add('GET', '/admin/exchange/mailboxes/shared%40contoso.com/folders', 200, '{"value":[{"id":"F1","displayName":"Inbox","parentMailboxUrl":"https://graph.microsoft.com/beta/admin/exchange/mailboxes/MBX:aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee@11111111-2222-3333-4444-555555555555"}]}')
        $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
        $s = Set-MbtSession -Graph ([MailboxTransfer.GraphClient]::new([Func[string]] { 'test-token' }, $o, $route))
        $s.AppOnly = $true; $s.UserPrincipalName = 'shared@contoso.com'; $s.TenantId = $script:Tenant; $s.Token = [pscustomobject]@{ Account = $null; Scopes = @(); Roles = @(); AppOnly = $true }
        Resolve-MbtMailboxAddress -Address 'shared@contoso.com' -TenantId $script:Tenant | Should -Be 'MBX:aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee@11111111-2222-3333-4444-555555555555'
        @($route.Requests | Where-Object { $_.Url -match '/me[/?]' }).Count | Should -Be 0
    }
    It 'an address out of the RBAC scope fails cleanly: both lookups 403, the error says to check the scope' {
        $route = [MailboxTransfer.Tests.RouteHandler]::new()
        $route.Add('GET', '/settings/exchange', 403, '{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}')
        $route.Add('GET', '/admin/exchange/mailboxes/', 403, '{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}')
        $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
        $s = Set-MbtSession -Graph ([MailboxTransfer.GraphClient]::new([Func[string]] { 'test-token' }, $o, $route))
        $s.AppOnly = $true; $s.TenantId = $script:Tenant; $s.Token = [pscustomobject]@{ Account = $null; Scopes = @(); Roles = @(); AppOnly = $true }
        { Resolve-MbtMailboxAddress -Address 'other@contoso.com' -TenantId $script:Tenant } | Should -Throw '*RBAC for Applications*'
    }
    It 'the folders of a mailbox are addressed by the mailbox id, not by a user' {
        $id = 'MBX:bbbbbbbb-1111-2222-3333-444444444444@11111111-2222-3333-4444-555555555555'
        [MailboxTransfer.MailboxApi]::MailboxPath($id) | Should -Be "/admin/exchange/mailboxes/$id"
        [MailboxTransfer.MailboxApi]::MailboxPath('adele@contoso.com') | Should -Be '/admin/exchange/mailboxes/adele%40contoso.com'
    }
    It 'no code of the package builds a /me path (only the guarded fallback of Get-MbtExchangeSettingsPath)' {
        $hits = @(Get-ChildItem -LiteralPath $script:Root -Include *.ps1, *.psm1, *.cs -Recurse | Where-Object { $_.FullName -notmatch '\\lib\\' } |
                Select-String -Pattern "['""]/me[/'""]|users/me\b" | Where-Object { $_.Line -notmatch '^\s*#|^\s*//' })
        @($hits).Count | Should -Be 1
        $hits[0].Path | Should -BeLike '*MailboxTransfer.psm1'
        $hits[0].Line | Should -Match 'Test-MbtAppOnly'
    }
}

Describe 'Application permissions of each mode' {
    It 'Import and ImportPst need folders ReadWrite, items Read and ImportExport (User.Read.All optional)' {
        foreach ($m in 'Import', 'ImportPst') {
            $rows = Get-MbtAppPermission -Mode $m
            ($rows | Where-Object Required).Permission | Should -Be @('MailboxFolder.ReadWrite.All', 'MailboxItem.Read.All', 'MailboxItem.ImportExport.All')
            ($rows | Where-Object { -not $_.Required }).Permission | Should -Be @('User.Read.All')
        }
    }
    It 'Export needs folders Read, items Read and Export (ImportExport also satisfies it)' {
        ($rows = Get-MbtAppPermission -Mode Export | Where-Object Required).Permission | Should -Be @('MailboxFolder.Read.All', 'MailboxItem.Read.All', 'MailboxItem.Export.All')
        ($rows | Where-Object Permission -EQ 'MailboxItem.Export.All').AnyOf | Should -Contain 'MailboxItem.ImportExport.All'
    }
    It 'Browse of the primary or the archive adds Mail.Read; an auxiliary archive does not use the Mail API' {
        (Get-MbtAppPermission -Mode Browse -Mailbox Primary | Where-Object Required).Permission | Should -Be @('Mail.Read', 'MailboxFolder.Read.All', 'MailboxItem.Read.All')
        (Get-MbtAppPermission -Mode Browse -Mailbox AuxArchive | Where-Object Required).Permission | Should -Be @('MailboxFolder.Read.All', 'MailboxItem.Read.All')
        (Get-MbtAppPermission -Mode ExportEml | Where-Object Permission -EQ 'Mail.Read') | Should -Not -BeNullOrEmpty
    }
    It 'ImportEml with the Exchange conversion also needs Mail.ReadWrite' {
        (Get-MbtAppPermission -Mode ImportEml | Where-Object Permission -EQ 'Mail.ReadWrite') | Should -BeNullOrEmpty
        (Get-MbtAppPermission -Mode ImportEml -ExchangeConversion | Where-Object Permission -EQ 'Mail.ReadWrite').Required | Should -BeTrue
    }
    It 'local modes need nothing' {
        foreach ($m in 'Status', 'ExportPst', 'CheckPst', 'RepairPst') { @(Get-MbtAppPermission -Mode $m).Count | Should -Be 0 }
    }
    It 'Get-MbtMissingAppPermission compares with the roles of the token (ReadWrite satisfies Read; ImportExport satisfies Export)' {
        $tok = Get-MbtTokenInfo -AccessToken (New-AppToken @('MailboxFolder.ReadWrite.All', 'MailboxItem.ReadWrite.All', 'MailboxItem.ImportExport.All'))
        @(Get-MbtMissingAppPermission -Token $tok -Mode Export).Count | Should -Be 0
        @(Get-MbtMissingAppPermission -Token $tok -Mode Import).Count | Should -Be 0
        Get-MbtMissingAppPermission -Token $tok -Mode Browse | Should -Be @('Mail.Read')
        $none = Get-MbtTokenInfo -AccessToken (New-AppToken @('MailboxFolder.Read.All'))
        Get-MbtMissingAppPermission -Token $none -Mode Import | Should -Be @('MailboxFolder.ReadWrite.All', 'MailboxItem.Read.All', 'MailboxItem.ImportExport.All')
    }
}

Describe 'Access hint of an app-only session' {
    It 'names the application permission that is missing and the RBAC for Applications check' {
        $h = Get-MbtAccessHint -Message 'HTTP 403 ErrorAccessDenied' -Mode Import -Scopes @('MailboxFolder.ReadWrite', 'MailboxItem.Read') -AppOnly
        $h | Should -Match 'MailboxItem.ImportExport.All'
        $h | Should -Not -Match 'MailboxFolder.ReadWrite.All'
        $h | Should -Match 'Test-ServicePrincipalAuthorization'
        $h | Should -Not -Match 'FullAccess'
    }
    It 'when the permissions are in the token the refusal comes from the Exchange scope' {
        $h = Get-MbtAccessHint -Message '403' -Mode Import -Scopes @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport') -AppOnly
        $h | Should -Match 'management scope'
        $h | Should -Match '30 minutes to 2 hours'
    }
    It 'a delegated session keeps its FullAccess hint' {
        Get-MbtAccessHint -Message '403' -Mode Import -Scopes @('MailboxFolder.ReadWrite') -OtherMailbox | Should -Match 'FullAccess'
    }
}

Describe 'Entry script and capability' {
    BeforeAll {
        function Invoke-EntryChild {
            param([string[]]$Arguments)
            $out = & pwsh -NoProfile -NonInteractive -File $script:Entry -DevelopmentMode @Arguments 2>&1 | Out-String
            [pscustomobject]@{ Code = $LASTEXITCODE; Text = $out }
        }
    }
    It 'Get-MbtCapability reports AppOnly from the parameter -CertificateThumbprint' {
        (Get-MbtCapability).AppOnly | Should -BeTrue
        $bare = Join-Path $TestDrive 'bare.ps1'
        Set-Content $bare "param([ValidateSet('Status','Export')][string]`$Mode)"
        (Get-MbtCapability -EntryScript $bare).AppOnly | Should -BeFalse
    }
    It 'the entry script takes the certificate parameters for every mode' {
        $p = (Get-Command $script:Entry).Parameters
        foreach ($n in 'AppOnly', 'CertificateThumbprint', 'CertificatePath', 'CertificatePassword') { $p.ContainsKey($n) | Should -BeTrue -Because $n }
        $p['CertificatePassword'].ParameterType | Should -Be ([securestring])
    }
    It 'the configuration has the additive Auth keys, empty by default' {
        $c = Import-MbtConfiguration
        foreach ($k in 'AppTenantId', 'AppClientId', 'CertificateThumbprint', 'CertificatePath') { $c.Auth.$k | Should -Be '' -Because $k }
    }
    It 'app-only never falls back on the lab application: -ClientId is required' {
        $r = Invoke-EntryChild @('-Mode', 'Export', '-CertificateThumbprint', ('AB' * 20), '-UserPrincipalName', 'adele@contoso.com', '-TenantId', $script:Tenant, '-StorePath', (Join-Path $script:Tmp 's1'))
        $r.Code | Should -Not -Be 0
        $r.Text | Should -Match 'App-only: give -ClientId'
    }
    It 'app-only never assumes the lab user: the mailbox address is required' {
        $r = Invoke-EntryChild @('-Mode', 'Export', '-CertificateThumbprint', ('AB' * 20), '-ClientId', $script:Client, '-TenantId', $script:Tenant, '-StorePath', (Join-Path $script:Tmp 's2'))
        $r.Code | Should -Not -Be 0
        $r.Text | Should -Match 'give -UserPrincipalName, the address of the target mailbox'
    }
    It 'a certificate that is not in the stores stops the run before any sign-in request' {
        $r = Invoke-EntryChild @('-Mode', 'Browse', '-CertificateThumbprint', ('00' * 20), '-ClientId', $script:Client, '-TenantId', $script:Tenant, '-UserPrincipalName', 'adele@contoso.com', '-StorePath', (Join-Path $script:Tmp 's3'))
        $r.Code | Should -Not -Be 0
        $r.Text | Should -Match 'No certificate with the thumbprint'
        $r.Text | Should -Match 'signing in as the application'
    }
    It '-DeviceCode cannot be combined with a certificate' {
        $r = Invoke-EntryChild @('-Mode', 'Browse', '-CertificateThumbprint', ('00' * 20), '-DeviceCode', '-ClientId', $script:Client, '-TenantId', $script:Tenant, '-UserPrincipalName', 'adele@contoso.com', '-StorePath', (Join-Path $script:Tmp 's4'))
        $r.Text | Should -Match 'DeviceCode is a delegated sign-in'
    }
    It '-AppOnly without any certificate says what to give' {
        $r = Invoke-EntryChild @('-Mode', 'Status', '-AppOnly', '-StorePath', (Join-Path $script:Tmp 's5'))
        $r.Text | Should -Match '-AppOnly needs -CertificateThumbprint or -CertificatePath'
    }
}

Describe 'tools\New-MbtAppOnlyLab.ps1 (optional lab setup, never run against a tenant by the tests)' {
    BeforeAll { $script:Lab = Join-Path $script:RepoRoot 'tools\New-MbtAppOnlyLab.ps1' }
    It 'parses, supports -WhatIf and needs an explicit tenant (no lab default)' {
        $errors = $null; $null = [Management.Automation.Language.Parser]::ParseFile($script:Lab, [ref]$null, [ref]$errors)
        @($errors).Count | Should -Be 0
        $cmd = Get-Command $script:Lab
        $cmd.Parameters.ContainsKey('WhatIf') | Should -BeTrue
        $cmd.Parameters['TenantId'].Attributes.Where({ $_ -is [Management.Automation.ParameterAttribute] }).Mandatory | Should -BeTrue
    }
    It '-WhatIf is offline: plans, prints the admin consent and RBAC for Applications steps, creates no certificate' {
        $subject = 'CN=MbtLabWhatIfTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
        $out = & pwsh -NoProfile -NonInteractive -File $script:Lab -TenantId 'contoso.onmicrosoft.com' -CertificateSubject $subject -WhatIf 2>&1 | Out-String
        $LASTEXITCODE | Should -Be 0
        $out | Should -Match 'WhatIf: no sign-in, no certificate, no application was created'
        $out | Should -Match 'MailboxFolder.ReadWrite.All, MailboxItem.Read.All, MailboxItem.ImportExport.All, Mail.Read'
        $out | Should -Match 'Admin consent\s+: NOT granted'
        $out | Should -Match 'New-ServicePrincipal'
        $out | Should -Match 'New-ManagementScope'
        $out | Should -Match 'New-ManagementRoleAssignment -App'
        $out | Should -Match 'Application MailboxItem.ImportExport'
        $out | Should -Match 'Test-ServicePrincipalAuthorization'
        $out | Should -Match 'adminconsent\?client_id='
        $out | Should -Match 'NOT exportable'
        @(Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq $subject }).Count | Should -Be 0
    }
    It 'the script never grants consent or creates a secret' {
        $text = Get-Content -LiteralPath $script:Lab -Raw
        $text | Should -Not -Match 'appRoleAssignedTo|AppRoleAssignment\.ReadWrite|oauth2PermissionGrants|passwordCredentials|addPassword'
        $text | Should -Match 'KeyExportPolicy NonExportable'
        $text | Should -Not -Match '-KeySpec' -Because 'the key storage provider refuses a CAPI key spec (NTE_PROV_TYPE_NOT_DEF)'
    }
}