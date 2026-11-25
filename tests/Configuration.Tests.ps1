#Requires -Version 7.6
# Configuration: the Tenant section (and Lab, the name it had), read by Import-MbtConfiguration, the command line and the console. Nothing contacts the cloud.
BeforeAll {
    $script:Package = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    Import-Module (Join-Path $script:Package 'MailboxTransfer.psd1') -Force
    $script:SavedLocal = $env:MAILBOXTRANSFER_LOCAL_CONFIG
    function Read-TestConfiguration([string]$Text) {
        $file = Join-Path $TestDrive ('local-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.psd1')
        Set-Content -LiteralPath $file -Value $Text -Encoding UTF8
        $env:MAILBOXTRANSFER_LOCAL_CONFIG = $file
        Import-MbtConfiguration
    }
}
AfterAll { $env:MAILBOXTRANSFER_LOCAL_CONFIG = $script:SavedLocal }

Describe 'Configuration: Tenant, and Lab as the name it had' {
    It 'reads the tenant, the application and the user from the Tenant section' {
        $c = Read-TestConfiguration "@{ Tenant = @{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '11111111-2222-3333-4444-555555555555'; UserPrincipalName = 'adele.vance@contoso.com' } }"
        $c.Tenant.TenantId | Should -Be 'contoso.onmicrosoft.com'
        $c.Tenant.ClientId | Should -Be '11111111-2222-3333-4444-555555555555'
        $c.Tenant.UserPrincipalName | Should -Be 'adele.vance@contoso.com'
    }
    It 'a file of the earlier version that still says Lab is read the same way: Tenant shows what Lab says' {
        $c = Read-TestConfiguration "@{ Lab = @{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '11111111-2222-3333-4444-555555555555'; UserPrincipalName = 'adele.vance@contoso.com'; PrimaryMailboxId = 'MBX:aaaaaaaa-0000-0000-0000-000000000001@contoso.onmicrosoft.com' } }"
        $c.Tenant.TenantId | Should -Be 'contoso.onmicrosoft.com'
        $c.Tenant.ClientId | Should -Be '11111111-2222-3333-4444-555555555555'
        $c.Tenant.UserPrincipalName | Should -Be 'adele.vance@contoso.com'
        $c.Tenant.PrimaryMailboxId | Should -Be 'MBX:aaaaaaaa-0000-0000-0000-000000000001@contoso.onmicrosoft.com'
    }
    It 'the lab scripts keep reading Lab: a value written in Tenant is there too, and the keys that are only theirs stay in Lab' {
        $c = Read-TestConfiguration "@{ Tenant = @{ TenantId = 'contoso.onmicrosoft.com'; UserPrincipalName = 'adele.vance@contoso.com' }; Lab = @{ UserObjectId = '1a2b3c05-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; SourceFolder = 'Inbox' } }"
        $c.Lab.TenantId | Should -Be 'contoso.onmicrosoft.com'
        $c.Lab.UserPrincipalName | Should -Be 'adele.vance@contoso.com'
        $c.Lab.UserObjectId | Should -Be '1a2b3c05-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $c.Lab.SourceFolder | Should -Be 'Inbox'
        $c.PSObject.Properties['Tenant'].Value.Keys | Should -Not -Contain 'SourceFolder'
    }
    It 'Tenant wins over Lab when both name the same key' {
        $c = Read-TestConfiguration "@{ Lab = @{ TenantId = 'your-tenant.onmicrosoft.com' }; Tenant = @{ TenantId = 'tenant.onmicrosoft.com' } }"
        $c.Tenant.TenantId | Should -Be 'tenant.onmicrosoft.com'
        $c.Lab.TenantId | Should -Be 'tenant.onmicrosoft.com'
    }
    It 'the application that signs in with a certificate stays in Auth' {
        $c = Read-TestConfiguration "@{ Auth = @{ AppTenantId = 'contoso.onmicrosoft.com'; AppClientId = '22222222-2222-3333-4444-555555555555'; CertificateThumbprint = 'AB12' } }"
        $c.Auth.AppClientId | Should -Be '22222222-2222-3333-4444-555555555555'
        $c.Auth.CertificateThumbprint | Should -Be 'AB12'
        $c.Tenant.TenantId | Should -BeNullOrEmpty
    }
    It 'an unknown key of Tenant is refused with its name, as everywhere else' {
        { Read-TestConfiguration "@{ Tenant = @{ Nonsense = 1 } }" } | Should -Throw '*unknown key*Tenant.Nonsense*'
    }
    It 'the shipped file holds placeholders only, in Tenant, Lab and Auth, and its comments name the new section' {
        $path = Join-Path $script:Package 'config\MailboxTransfer.config.psd1'
        $shipped = Import-PowerShellDataFile -LiteralPath $path
        foreach ($section in 'Tenant', 'Lab') { foreach ($k in $shipped[$section].Keys) { $shipped[$section][$k] | Should -BeNullOrEmpty -Because "$section.$k is a placeholder" } }
        foreach ($k in 'AppTenantId', 'AppClientId', 'CertificateThumbprint', 'CertificatePath') { $shipped.Auth[$k] | Should -BeNullOrEmpty }
        (Get-Content -LiteralPath $path -Raw) | Should -Match 'Tenant section'
    }
    It 'the ids of the mailboxes of the configuration are the ones of the user of the configuration only, read from Tenant or from Lab' {
        $tenant = [pscustomobject]@{ Tenant = @{ UserPrincipalName = 'adele.vance@contoso.com'; PrimaryMailboxId = 'MBX:p'; ArchiveMailboxId = 'MBX:a' } }
        $lab = [pscustomobject]@{ Lab = @{ UserPrincipalName = 'adele.vance@contoso.com'; PrimaryMailboxId = 'MBX:p'; ArchiveMailboxId = 'MBX:a' } }
        foreach ($cfg in $tenant, $lab) {
            (Get-MbtConfiguredMailboxId -Configuration $cfg -UserPrincipalName 'ADELE.vance@contoso.com').Primary | Should -Be 'MBX:p'
            (Get-MbtConfiguredMailboxId -Configuration $cfg -UserPrincipalName 'megan.bowen@contoso.com').Primary | Should -BeNullOrEmpty
        }
    }
    It 'the command line takes the tenant, the application and the user from Tenant when the parameters are not given' {
        $text = Get-Content -LiteralPath (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1') -Raw
        $text | Should -Match '\$TenantId = \$config\.Tenant\.TenantId'
        $text | Should -Match '\$ClientId = \$config\.Tenant\.ClientId'
        $text | Should -Match '\$UserPrincipalName = \$config\.Tenant\.UserPrincipalName'
    }
}

Describe 'Configuration: what the console shows on the Setup page' {
    BeforeAll {
        foreach ($f in 'Core', 'Run', 'Jobs', 'Browse', 'State', 'Pages', 'Pages.Export', 'Pages.Run', 'Pages.Source', 'Fake', 'Engine', 'Scenarios') { . (Join-Path $script:Package "module\MailboxTransfer.Gui.$f.ps1") }
        $script:GuiRoot = $script:Package
        $script:Engine = $null
    }
    It 'the real adapter gives the tenant and the application of Tenant (or Lab) and the three values of the application of Auth, and nothing is written back' {
        $file = Join-Path $TestDrive 'setup-local.psd1'
        Set-Content -LiteralPath $file -Encoding UTF8 -Value "@{ Lab = @{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '11111111-2222-3333-4444-555555555555' }; Auth = @{ AppTenantId = 'example.onmicrosoft.com'; AppClientId = '22222222-2222-3333-4444-555555555555'; CertificateThumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678' } }"
        $before = (Get-FileHash -LiteralPath $file).Hash
        $env:MAILBOXTRANSFER_LOCAL_CONFIG = $file
        $engine = New-MbtGuiRealEngine -PackageRoot $script:Package -EntryScript (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1') -WorkFolder $TestDrive
        $v = & $engine.SignInSettings
        $v.TenantId | Should -Be 'contoso.onmicrosoft.com'
        $v.ClientId | Should -Be '11111111-2222-3333-4444-555555555555'
        $v.AppTenantId | Should -Be 'example.onmicrosoft.com'
        $v.AppClientId | Should -Be '22222222-2222-3333-4444-555555555555'
        $v.CertificateThumbprint | Should -Be 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678'
        & $engine.ConfigPath | Should -Be $file -Because 'the local file is the one Open shows when there is one'
        (Get-FileHash -LiteralPath $file).Hash | Should -Be $before
        $light = @((& $engine.Readiness).Lights | Where-Object Name -EQ 'Configuration file')
        $light.Count | Should -Be 1
        $light[0].Ok | Should -BeTrue
    }
    It 'Keep this PC awake is off by default, so that -KeepAwake is not in the shown command line' {
        (New-MbtGuiState -WorkFolder $TestDrive).Awake | Should -BeFalse
    }
    It 'the real work folder light gives the path and the free space, and creates the folder' {
        $env:MAILBOXTRANSFER_LOCAL_CONFIG = (Join-Path $TestDrive 'none-local.psd1')
        $work = Join-Path $TestDrive 'work-one\MailboxTransfer'
        $engine = New-MbtGuiRealEngine -PackageRoot $script:Package -EntryScript (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1') -WorkFolder $work
        $light = @((& $engine.Readiness).Lights | Where-Object Name -EQ 'Work folder')
        $light.Count | Should -Be 1
        $light[0].Ok | Should -BeTrue
        $light[0].Detail | Should -BeLike "$work - * free"
        $light[0].Detail | Should -Not -Match 'unknown'
        Test-Path -LiteralPath $work -PathType Container | Should -BeTrue
        $other = Join-Path $TestDrive 'work-two'
        (@((& $engine.Readiness $other).Lights | Where-Object Name -EQ 'Work folder')[0]).Detail | Should -BeLike "$other - * free" -Because 'the folder chosen on the Setup page is the one that is checked'
    }
    It 'the work folder light is red, in plain words, when the folder cannot be created, written, or has no room' {
        $file = Join-Path $TestDrive 'a-file'
        Set-Content -LiteralPath $file -Value 'x'
        $bad = Get-MbtWorkFolderStatus -Path (Join-Path $file 'below')
        $bad.Ok | Should -BeFalse
        $bad.Problem | Should -Match '^cannot be created'
        $bad.Problem | Should -Match 'Choose another folder'
        $full = Get-MbtWorkFolderStatus -Path (Join-Path $TestDrive 'room') -MinimumFreeBytes ([long]::MaxValue)
        $full.Ok | Should -BeFalse
        $full.Problem | Should -Match 'has only .* free: at least .* is needed'
        $env:MAILBOXTRANSFER_LOCAL_CONFIG = (Join-Path $TestDrive 'none-local.psd1')
        $engine = New-MbtGuiRealEngine -PackageRoot $script:Package -EntryScript (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1') -WorkFolder (Join-Path $file 'below')
        $light = @((& $engine.Readiness).Lights | Where-Object Name -EQ 'Work folder')
        $light[0].Ok | Should -BeFalse
        $light[0].Detail | Should -Match 'cannot be created'
    }
    It 'the sign-in component light says whether the Windows sign-in window is available' {
        $env:MAILBOXTRANSFER_LOCAL_CONFIG = (Join-Path $TestDrive 'none-local.psd1')
        $engine = New-MbtGuiRealEngine -PackageRoot $script:Package -EntryScript (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1') -WorkFolder $TestDrive
        $light = @((& $engine.Readiness).Lights | Where-Object Name -EQ 'Sign-in component (MSAL)')
        $light[0].Detail | Should -Match 'Windows sign-in window is (available|not available)'
        $window = & $engine.WindowSignIn
        $window.PSObject.Properties.Name | Should -Contain 'Available'
        $window.PSObject.Properties.Name | Should -Contain 'Reason'
        $script:Gui = @{ S = @{ WorkFolder = $TestDrive } }
        $fake = New-MbtGuiFakeEngine
        @((& $fake.Readiness).Lights | Where-Object Name -EQ 'Sign-in component (MSAL)')[0].Detail | Should -Match 'Windows sign-in window is available'
    }
    It 'the application tenant is the one of the configuration when Auth.AppTenantId is empty; the user and the mailbox are never borrowed' {
        $script:Gui = @{ Engine = @{ SignInSettings = { [pscustomobject]@{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '11111111-2222-3333-4444-555555555555'; AppTenantId = ''; AppClientId = '22222222-2222-3333-4444-555555555555'; CertificateThumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678'; UserPrincipalName = 'adele.vance@contoso.com'; PrimaryMailboxId = 'MBX:p' } } } }
        $s = @{ AuthMode = 'app'; AppTenant = ''; AppClient = ''; AppThumb = ''; TenantId = ''; ClientId = '' }
        $e = Get-MbtGuiEffectiveSignIn -S $s
        $e.Tenant | Should -Be 'contoso.onmicrosoft.com'
        $e.Client | Should -Be '22222222-2222-3333-4444-555555555555'
        $e.TenantChanged | Should -BeFalse
        $e.Missing | Should -Not -Contain 'Tenant'
        (Test-MbtGuiAppSignIn -S $s) | Should -BeNullOrEmpty
        $script:Gui.Engine.SignInSettings = { [pscustomobject]@{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '11111111-2222-3333-4444-555555555555'; AppTenantId = 'example.onmicrosoft.com'; AppClientId = '22222222-2222-3333-4444-555555555555'; CertificateThumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678' } }
        (Get-MbtGuiEffectiveSignIn -S $s).Tenant | Should -Be 'example.onmicrosoft.com' -Because 'the tenant of the application wins when it is given'
        $script:Gui = $null
    }
    It 'a missing redirect URI for the Windows sign-in window is said in plain words with the exact URI to add' {
        Initialize-MbtAuthEngine
        $client = '33333333-2222-3333-4444-555555555555'
        $raw = [MailboxTransfer.WindowsBroker]::MissingRedirectUriMessage('AADSTS50011: The redirect URI ''ms-appx-web://microsoft.aad.brokerplugin/x'' does not match', $client)
        $m = ConvertTo-MbtGuiMessage $raw
        $m.Text | Should -Match 'redirect URI'
        $m.Text | Should -Match ([regex]::Escape("ms-appx-web://microsoft.aad.brokerplugin/$client"))
        $m.Text | Should -Match 'use a code instead'
    }
    It 'a sign-in window that was closed, or that did not finish in time, is said in plain words with the way on' {
        $closed = ConvertTo-MbtGuiMessage 'The sign-in window was closed before the sign-in was finished. Sign in again, or sign in with a code.'
        $closed.Text | Should -Match 'closed before'
        $closed.Text | Should -Match 'use a code instead'
        $late = ConvertTo-MbtGuiMessage 'The sign-in window did not finish within 5 minutes. Sign in again, or sign in with a code.'
        $late.Text | Should -Match 'did not finish in time'
        $late.Text | Should -Match 'use a code instead'
        $late.Tech | Should -Match '5 minutes'
    }
    It 'a configuration file that cannot be read is a red light, not a crash' {
        $file = Join-Path $TestDrive 'setup-bad.psd1'
        Set-Content -LiteralPath $file -Encoding UTF8 -Value "@{ Tenant = @{ Nonsense = 1 } }"
        $env:MAILBOXTRANSFER_LOCAL_CONFIG = $file
        $engine = New-MbtGuiRealEngine -PackageRoot $script:Package -EntryScript (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1') -WorkFolder $TestDrive
        $script:Gui = @{ S = @{ WorkFolder = $TestDrive } }
        $light = @((& $engine.Readiness).Lights | Where-Object Name -EQ 'Configuration file')
        $light[0].Ok | Should -BeFalse
        $light[0].Detail | Should -Match 'Nonsense'
    }
}

Describe 'PowerShell minimum: 7.6 for the console and the command line' {
    It 'the manifest and product.json ask for 7.6' {
        (Import-PowerShellDataFile -LiteralPath (Join-Path $script:Package 'MailboxTransfer.psd1')).PowerShellVersion | Should -Be '7.6'
        $product = Get-Content -LiteralPath (Join-Path $script:Package 'product.json') -Raw | ConvertFrom-Json
        $product.powerShell.minimum | Should -Be '7.6'
        $product.powerShell.commandLineMinimum | Should -Be '7.6'
    }
    It 'every script of the package asks for 7.6 with #Requires' {
        $bad = foreach ($f in Get-ChildItem -LiteralPath $script:Package -Recurse -Include *.ps1, *.psm1 -File) {
            $line = Get-Content -LiteralPath $f.FullName -TotalCount 3 | Where-Object { $_ -match '#Requires\s+-Version' } | Select-Object -First 1
            if ($line -notmatch '-Version\s+7\.6\s*$') { $f.Name }
        }
        @($bad) | Should -BeNullOrEmpty
    }
    It 'the launcher refuses an older PowerShell and names the way to install or update it' {
        $cmd = Get-Content -LiteralPath (Join-Path $script:Package 'Start-MailboxTransferConsole.cmd') -Raw
        $cmd | Should -Match "\[version\]'7\.6'"
        $cmd | Should -Match 'winget install Microsoft\.PowerShell'
        $cmd | Should -Match 'winget upgrade Microsoft\.PowerShell'
    }
    It 'the first light of the Setup page asks for PowerShell 7.6 or later, and the running PowerShell satisfies it' {
        foreach ($f in 'Core', 'Run', 'Jobs', 'Browse', 'State', 'Pages', 'Pages.Export', 'Pages.Run', 'Pages.Source', 'Fake', 'Engine', 'Scenarios') { . (Join-Path $script:Package "module\MailboxTransfer.Gui.$f.ps1") }
        $script:GuiRoot = $script:Package
        $engine = New-MbtGuiRealEngine -PackageRoot $script:Package -EntryScript (Join-Path $script:Package 'Invoke-MailboxTransfer.ps1') -WorkFolder $TestDrive
        $light = @((& $engine.Readiness).Lights | Select-Object -First 1)
        $light[0].Name | Should -Be 'PowerShell 7.6 or later'
        $light[0].Ok | Should -BeTrue
    }
}
