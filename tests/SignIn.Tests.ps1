#Requires -Version 7.6
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the delegated sign-in: the Windows sign-in window (the Windows account broker, WAM) loaded from the Microsoft.Graph.Authentication
    module that holds MSAL, the device code when that component cannot be used, the message for a redirect URI that is not registered, and the absence of any
    system-browser sign-in. The window itself is never opened and nothing here reaches the network: the real pop-up is checked by hand.

    Run:  Invoke-Pester -Path .\tests\SignIn.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root
    Initialize-MbtAuthEngine
    $script:Window = Get-MbtWindowsSignIn
    $script:Client = '11111111-2222-3333-4444-555555555555'
    $script:Scopes = [string[]]@('https://graph.microsoft.com/.default')
    $script:MsalDll = [Microsoft.Identity.Client.PublicClientApplication].Assembly.Location
    function New-GraphLikeFolder {
        # a copy of the folder layout of Microsoft.Graph.Authentication: Dependencies\Core\Microsoft.Identity.Client.dll, with or without the broker files
        param([string]$Name, [switch]$Broker, [switch]$Native)
        $deps = Join-Path $TestDrive "$Name\Dependencies"
        $null = New-Item -ItemType Directory -Force -Path (Join-Path $deps 'Core')
        Copy-Item -LiteralPath $script:MsalDll -Destination (Join-Path $deps 'Core\Microsoft.Identity.Client.dll')
        if ($Broker) { Set-Content -LiteralPath (Join-Path $deps 'Microsoft.Identity.Client.Broker.dll') -Value 'not an assembly' }
        if ($Native) { Set-Content -LiteralPath (Join-Path $deps ([MailboxTransfer.WindowsBroker]::NativeRuntimeFile())) -Value 'not a library' }
        return (Join-Path $deps 'Core\Microsoft.Identity.Client.dll')
    }
}

Describe 'Windows sign-in window: is it available' {
    It 'answers with Available, Reason and Directory, without loading anything' {
        $script:Window.PSObject.Properties.Name | Should -Contain 'Available'
        $script:Window.PSObject.Properties.Name | Should -Contain 'Reason'
        $script:Window.PSObject.Properties.Name | Should -Contain 'Directory'
        if ($script:Window.Available) { $script:Window.Reason | Should -BeNullOrEmpty } else { $script:Window.Reason | Should -Not -BeNullOrEmpty }
    }
    It 'finds the broker and its native runtime next to the MSAL of the Microsoft.Graph.Authentication module (same version)' {
        if (-not $script:Window.Available) { Set-ItResult -Skipped -Because 'the Windows sign-in component is not installed on this machine'; return }
        $script:Window.Directory | Should -Not -BeNullOrEmpty
        Test-Path -LiteralPath (Join-Path $script:Window.Directory 'Microsoft.Identity.Client.Broker.dll') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:Window.Directory ([MailboxTransfer.WindowsBroker]::NativeRuntimeFile())) | Should -BeTrue
        $msal = [Reflection.AssemblyName]::GetAssemblyName($script:MsalDll).Version
        $broker = [Reflection.AssemblyName]::GetAssemblyName((Join-Path $script:Window.Directory 'Microsoft.Identity.Client.Broker.dll')).Version
        $broker.Major | Should -Be $msal.Major
        $broker.Minor | Should -Be $msal.Minor
    }
    It 'is not available, and says how to fix it, when the Graph module has no broker (an older module)' {
        $w = Get-MbtWindowsSignIn -MsalPath (New-GraphLikeFolder -Name 'old')
        $w.Available | Should -BeFalse
        $w.Reason | Should -Match 'no Windows sign-in component'
        $w.Reason | Should -Match 'Update-Module Microsoft.Graph.Authentication'
    }
    It 'is not available when the native runtime is missing' {
        $w = Get-MbtWindowsSignIn -MsalPath (New-GraphLikeFolder -Name 'nonative' -Broker)
        $w.Available | Should -BeFalse
        $w.Reason | Should -Match 'incomplete'
        $w.Reason | Should -Match ([regex]::Escape([MailboxTransfer.WindowsBroker]::NativeRuntimeFile()))
    }
    It 'is not available when the broker file cannot be read as an assembly' {
        $w = Get-MbtWindowsSignIn -MsalPath (New-GraphLikeFolder -Name 'broken' -Broker -Native)
        $w.Available | Should -BeFalse
        $w.Reason | Should -Match 'cannot be read'
    }
    It 'is not available when MSAL itself is not found' {
        $w = Get-MbtWindowsSignIn -MsalPath (Join-Path $TestDrive 'nowhere\Microsoft.Identity.Client.dll')
        $w.Available | Should -BeFalse
        $w.Reason | Should -Match 'MSAL'
    }
}

Describe 'GraphAuth: the Windows sign-in window and the code' {
    It 'builds a broker application with the broker redirect URI, apart from the application of the DPAPI cache and of the code' {
        if (-not $script:Window.Available) { Set-ItResult -Skipped -Because 'the Windows sign-in component is not installed on this machine'; return }
        $a = [MailboxTransfer.GraphAuth]::new($script:Client, 'organizations', $script:Scopes, (Join-Path $TestDrive 'token.cache'), 'http://localhost', $false, $null, $script:Window.Directory, $null)
        $a.BrokerEnabled | Should -BeTrue
        $a.SignInMethod | Should -Be 'window'
        $a.BrokerProblem | Should -BeNullOrEmpty
        $a.WindowApp.AppConfig.IsBrokerEnabled | Should -BeTrue
        $a.WindowApp.AppConfig.RedirectUri | Should -Be "ms-appx-web://microsoft.aad.brokerplugin/$($script:Client)"
        $a.CacheApp.AppConfig.IsBrokerEnabled | Should -BeFalse -Because 'the silent reuse of the DPAPI cache and the device code do not go through the broker'
        $a.CacheApp.AppConfig.RedirectUri | Should -Be 'http://localhost' -Because 'it is harmless and no longer needed for an interactive sign-in'
        [MailboxTransfer.WindowsBroker]::RedirectUri($script:Client) | Should -Be "ms-appx-web://microsoft.aad.brokerplugin/$($script:Client)"
    }
    It 'uses the code when it is asked for, even if the window is available' {
        if (-not $script:Window.Available) { Set-ItResult -Skipped -Because 'the Windows sign-in component is not installed on this machine'; return }
        $a = [MailboxTransfer.GraphAuth]::new($script:Client, 'organizations', $script:Scopes, $null, 'http://localhost', $true, $null, $script:Window.Directory, $null)
        $a.DeviceCode | Should -BeTrue
        $a.SignInMethod | Should -Be 'code'
    }
    It 'falls back on the device code, never on the system browser, when the window cannot be used' {
        $a = [MailboxTransfer.GraphAuth]::new($script:Client, 'organizations', $script:Scopes, $null, 'http://localhost', $false, $null, $null, 'the Windows sign-in component was not found')
        $a.BrokerEnabled | Should -BeFalse
        $a.WindowApp | Should -BeNullOrEmpty
        $a.SignInMethod | Should -Be 'code'
        $a.BrokerProblem | Should -Be 'the Windows sign-in component was not found'
        { $a.GetToken('adele.vance@contoso.com', $false) } | Should -Throw '*Not signed in*' -Because 'a silent call never opens a window nor asks for a code'
    }
    It 'in a fresh process, a broker folder that does not hold the broker gives a code and the reason' {
        $empty = Join-Path $TestDrive 'no-broker-here'
        $null = New-Item -ItemType Directory -Force -Path $empty
        $script = @"
[AppDomain]::CurrentDomain.SetData('MailboxTransfer.UnverifiedTreeWarned', `$true)
Import-Module '$($script:Root)\MailboxTransfer.psd1'
Initialize-MbtAuthEngine
`$a = [MailboxTransfer.GraphAuth]::new('$($script:Client)', 'organizations', [string[]]@('https://graph.microsoft.com/.default'), `$null, 'http://localhost', `$false, `$null, '$empty', `$null)
[pscustomobject]@{ Enabled = `$a.BrokerEnabled; Method = `$a.SignInMethod; Problem = `$a.BrokerProblem } | ConvertTo-Json -Compress
"@
        $r = (& pwsh -NoProfile -NonInteractive -Command $script 2>&1 | Out-String).Trim() | ConvertFrom-Json
        $r.Enabled | Should -BeFalse
        $r.Method | Should -Be 'code'
        $r.Problem | Should -Match 'no Windows sign-in component'
    }
    It 'New-MbtAuth: the window handle given by the caller is kept, and -DeviceCode wins' {
        if (-not $script:Window.Available) { Set-ItResult -Skipped -Because 'the Windows sign-in component is not installed on this machine'; return }
        $cfg = Import-MbtConfiguration
        $a = & (Get-Module MailboxTransfer) { param($c, $client, $scopes, $cache) New-MbtAuth -TenantId 'organizations' -ClientId $client -Scopes $scopes -CacheFile $cache -Configuration $c -ParentWindow 123456 } $cfg $script:Client $script:Scopes (Join-Path $TestDrive 'a.cache')
        $a.ParentWindow.ToInt64() | Should -Be 123456
        $a.SignInMethod | Should -Be 'window'
        $b = & (Get-Module MailboxTransfer) { param($c, $client, $scopes, $cache) New-MbtAuth -TenantId 'organizations' -ClientId $client -Scopes $scopes -CacheFile $cache -Configuration $c -DeviceCode } $cfg $script:Client $script:Scopes (Join-Path $TestDrive 'b.cache')
        $b.SignInMethod | Should -Be 'code'
        $b.BrokerEnabled | Should -BeFalse
    }
    It 'the window the sign-in opens over is found without throwing (console window, else foreground window)' {
        { [MailboxTransfer.WindowsBroker]::ConsoleOrForegroundWindow() } | Should -Not -Throw
    }
    It 'under a terminal that hosts the console in a hidden pseudo window (Windows Terminal), the sign-in opens over the window that owns it, not over the pseudo window' {
        $pseudo = [IntPtr]100; $terminal = [IntPtr]200; $foreground = [IntPtr]300; $desktop = [IntPtr]400
        $usable = [Func[IntPtr, bool]] { param($h) $h -ne [IntPtr]100 }
        [MailboxTransfer.WindowsBroker]::ChooseParentWindow($pseudo, $terminal, $foreground, $desktop, $usable) | Should -Be $terminal
    }
    It 'in a classic console window (the console is its own owner) the sign-in opens over that window' {
        $console = [IntPtr]100; $foreground = [IntPtr]300; $desktop = [IntPtr]400
        [MailboxTransfer.WindowsBroker]::ChooseParentWindow($console, $console, $foreground, $desktop, [Func[IntPtr, bool]] { param($h) $true }) | Should -Be $console
    }
    It 'when no console window can be seen (hidden, minimized, no console) the foreground window is used, then the desktop' {
        $console = [IntPtr]100; $terminal = [IntPtr]200; $foreground = [IntPtr]300; $desktop = [IntPtr]400
        [MailboxTransfer.WindowsBroker]::ChooseParentWindow($console, $terminal, $foreground, $desktop, [Func[IntPtr, bool]] { param($h) $h -eq [IntPtr]300 }) | Should -Be $foreground
        [MailboxTransfer.WindowsBroker]::ChooseParentWindow($console, $terminal, $foreground, $desktop, [Func[IntPtr, bool]] { param($h) $false }) | Should -Be $desktop
        [MailboxTransfer.WindowsBroker]::ChooseParentWindow([IntPtr]::Zero, [IntPtr]::Zero, [IntPtr]::Zero, $desktop, [Func[IntPtr, bool]] { param($h) $true }) | Should -Be $desktop
    }
    It 'a window that is not there is not usable, and is described without throwing' {
        [MailboxTransfer.WindowsBroker]::IsUsableWindow([IntPtr]::Zero) | Should -BeFalse
        [MailboxTransfer.WindowsBroker]::IsUsableWindow([IntPtr]0x7FFFFFF0) | Should -BeFalse
        [MailboxTransfer.WindowsBroker]::Describe([IntPtr]::Zero) | Should -Be 'no window'
        [MailboxTransfer.WindowsBroker]::Describe([IntPtr]0x7FFFFFF0) | Should -Match '^window 0x'
    }
    It 'the waits are bounded: a silent attempt of the broker 30 seconds, the window 5 minutes' {
        [MailboxTransfer.GraphAuth]::SilentTimeout | Should -Be ([TimeSpan]::FromSeconds(30))
        [MailboxTransfer.GraphAuth]::WindowTimeout | Should -Be ([TimeSpan]::FromMinutes(5))
    }
    It 'a silent call writes nothing to the sign-in log, and the log is only a callback of the caller' {
        $lines = [Collections.Generic.List[string]]::new()
        $a = [MailboxTransfer.GraphAuth]::new($script:Client, 'organizations', $script:Scopes, $null, 'http://localhost', $false, $null, $null, 'no broker here')
        $a.Log = [Action[string]] { param($l) $lines.Add($l) }
        { $a.GetToken('adele.vance@contoso.com', $false) } | Should -Throw '*Not signed in*'
        $lines.Count | Should -Be 0
    }
    It 'Connect-MbtGraph takes a cancellation token and the sign-in log goes to Write-MbtLog' {
        (Get-Command Connect-MbtGraph).Parameters['CancellationToken'].ParameterType | Should -Be ([Threading.CancellationToken])
        $log = Join-Path $TestDrive 'logs'
        $null = Start-MbtLog -Directory $log
        $cfg = Import-MbtConfiguration
        $a = & (Get-Module MailboxTransfer) { param($c, $client, $scopes, $cache) New-MbtAuth -TenantId 'organizations' -ClientId $client -Scopes $scopes -CacheFile $cache -Configuration $c -DeviceCode } $cfg $script:Client $script:Scopes (Join-Path $TestDrive 'log.cache')
        $a.Log | Should -Not -BeNullOrEmpty
        $a.Log.Invoke('sign-in: a line')
        Get-Content -LiteralPath (Get-MbtLogPath) -Raw | Should -Match 'sign-in: a line'
        Stop-MbtLog
    }
    It 'the entry script and Connect-MbtGraph keep -DeviceCode and take the window handle' {
        $entry = (Get-Command (Join-Path $script:Root 'Invoke-MailboxTransfer.ps1')).Parameters
        $entry['DeviceCode'].ParameterType | Should -Be ([switch])
        $connect = (Get-Command Connect-MbtGraph).Parameters
        $connect['DeviceCode'].ParameterType | Should -Be ([switch])
        $connect.ContainsKey('ParentWindow') | Should -BeTrue
    }
}

Describe 'A redirect URI that is not registered' {
    It 'names the exact URI to add, where to add it, and the code as the way meanwhile' {
        foreach ($raw in 'AADSTS50011: The redirect URI specified in the request does not match the redirect URIs configured for the application', 'AADSTS500113: No reply address is registered for the application.') {
            $m = [MailboxTransfer.WindowsBroker]::MissingRedirectUriMessage($raw, $script:Client)
            $m | Should -Match ([regex]::Escape("ms-appx-web://microsoft.aad.brokerplugin/$($script:Client)"))
            $m | Should -Match 'Mobile and desktop applications'
            $m | Should -Match '-DeviceCode'
        }
    }
    It 'says nothing for another error' {
        [MailboxTransfer.WindowsBroker]::MissingRedirectUriMessage('AADSTS50173: The provided grant has expired', $script:Client) | Should -BeNullOrEmpty
        [MailboxTransfer.WindowsBroker]::MissingRedirectUriMessage('', $script:Client) | Should -BeNullOrEmpty
        [MailboxTransfer.WindowsBroker]::MissingRedirectUriMessage($null, $script:Client) | Should -BeNullOrEmpty
    }
}

Describe 'No sign-in through the system browser' {
    BeforeAll {
        $script:Files = @(Get-ChildItem -LiteralPath $script:Root -Recurse -File -Include '*.ps1', '*.psm1', '*.cs', '*.xaml' | Where-Object { $_.FullName -notmatch '\\lib\\' })
    }
    It 'the only interactive sign-in of the engine is the one of the Windows broker application' {
        $auth = Get-Content -LiteralPath (Join-Path $script:Root 'src\Engine.Auth.cs') -Raw
        ([regex]::Matches($auth, 'AcquireTokenInteractive')).Count | Should -Be 1
        $auth | Should -Match '_brokerApp\s*\.AcquireTokenInteractive'
        $auth | Should -Match '\.WithParentActivityOrWindow\('
        $auth | Should -Match '\.WithPrompt\(Prompt\.SelectAccount\)'
        $auth | Should -Not -Match 'UseEmbeddedWebView'
        $auth | Should -Not -Match 'WithSystemWebViewOptions|SystemWebViewOptions|UseSystemBrowser'
    }
    It 'nothing in the package opens a browser to sign in, and no text promises one' {
        foreach ($f in $script:Files) {
            $text = Get-Content -LiteralPath $f.FullName -Raw
            $text | Should -Not -Match '(?i)system browser' -Because $f.Name
            $text | Should -Not -Match 'UseEmbeddedWebView|SystemWebViewOptions' -Because $f.Name
            $text | Should -Not -Match '(?i)A window from Microsoft opens' -Because $f.Name
            $text | Should -Not -Match '(?i)Start-Process[^\r\n]*login\.microsoftonline' -Because $f.Name
        }
    }
    It 'the texts of the console and of the command line describe the sign-in window' {
        $page = Get-Content -LiteralPath (Join-Path $script:Root 'module\MailboxTransfer.Gui.Pages.Export.ps1') -Raw
        $page | Should -Match 'A Microsoft sign-in window opens over the console\. Type the account that may open the mailbox\.'
        $page | Should -Match 'No, sign in to this app only'
    }
}
