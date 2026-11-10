# Loads the module and its engine once per process: SQLite from lib\sqlite, then the C# sources of package\src including
# Engine.Store.cs (all of it is Initialize-MbtEngine in package\MailboxTransfer.psm1, nothing is duplicated here).
# With -WithTestSupport the test helpers (StoreTestSupport.cs) are compiled in the same unit: an in-memory assembly cannot be
# referenced by a later Add-Type, so every test file of a Pester run goes through here (MailboxTransfer.Tests.ps1 too),
# whichever runs first.
function Import-StoreEngine {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PackageRoot,
        [switch]$WithTestSupport
    )
    Import-Module (Join-Path $PackageRoot 'MailboxTransfer.psd1')
    $extra = if ($WithTestSupport) { @((Join-Path $PSScriptRoot 'StoreTestSupport.cs'), (Join-Path $PSScriptRoot 'PstMemorySupport.cs'), (Join-Path $PSScriptRoot 'PstRepairSupport.cs')) } else { @() }
    Initialize-MbtEngine -Root $PackageRoot -AdditionalSource $extra
}

# The private local configuration of the developer (package\config\MailboxTransfer.config.local.psd1) must never reach a test: every test
# file runs against this synthetic one, which holds the lab defaults the tests expect (contoso values, made up).
function Use-TestLocalConfiguration {
    $dir = Join-Path ([IO.Path]::GetTempPath()) 'MailboxTransfer-test-config'
    $null = New-Item -ItemType Directory -Force -Path $dir
    $file = Join-Path $dir 'MailboxTransfer.config.local.psd1'
    Set-Content -LiteralPath $file -Encoding UTF8 -Value @'
@{
    Lab = @{
        TenantId          = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        ClientId          = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        UserPrincipalName = 'adele.vance@contoso.com'
        UserObjectId      = '1a2b3c05-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        PrimaryMailboxId  = 'MBX:1a2b3c03-5e7f-4a1b-9c2d-3e4f5a6b7c8d@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        ArchiveMailboxId  = 'MBX:1a2b3c04-5e7f-4a1b-9c2d-3e4f5a6b7c8d@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        SourceFolder      = 'MBT-Fill-20261028'
        TargetFolder      = 'MBT-Import-20261028'
    }
}
'@
    $env:MAILBOXTRANSFER_LOCAL_CONFIG = $file
}
Use-TestLocalConfiguration

# A source checkout has no PackageIntegrity.psd1: the entry script is started with -DevelopmentMode (the module then does not repeat its warning).
[AppDomain]::CurrentDomain.SetData('MailboxTransfer.UnverifiedTreeWarned', $true)
$global:PSDefaultParameterValues['Invoke-MailboxTransfer.ps1:DevelopmentMode'] = $true
$global:PSDefaultParameterValues['Start-MailboxTransferConsole.ps1:DevelopmentMode'] = $true
# The log file of the runs goes to a folder of the temp, never to out\logs of the repository (the jobs that a test starts inherit the variable).
$env:MAILBOXTRANSFER_LOG_FOLDER = Join-Path ([IO.Path]::GetTempPath()) 'MailboxTransfer-test-logs'
# The icons of the console are the ASCII ones, whatever the terminal of the developer is: the tests look for them (+ ok, ! warning, x failure).
$env:MBT_ICONS = 'Ascii'
