#Requires -Version 7.4
# Console: the command line shown on the Verify and Next steps pages. Only what is a choice of the operator, nothing of the plumbing of the console, no value that is
# the default, runnable as it is, and the same result as the run. No window, nothing contacts the cloud.
BeforeAll {
    $script:Package = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    $script:Entry = Join-Path $script:Package 'Invoke-MailboxTransfer.ps1'
    foreach ($f in 'Core', 'Run', 'Jobs') { . (Join-Path $script:Package "module\MailboxTransfer.Gui.$f.ps1") }
    # the configuration file of this machine, as the engine reads it: the tenant and the application, and the application that signs in with a certificate
    $script:FileTenant = 'contoso.onmicrosoft.com'; $script:FileClient = '11111111-2222-3333-4444-555555555555'
    $script:FileAppTenant = 'example.onmicrosoft.com'; $script:FileAppClient = '22222222-2222-3333-4444-555555555555'; $script:FileThumb = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678'
    $script:Gui = @{ Engine = [pscustomobject]@{ SignInSettings = { [pscustomobject]@{ TenantId = $script:FileTenant; ClientId = $script:FileClient; AppTenantId = $script:FileAppTenant; AppClientId = $script:FileAppClient; CertificateThumbprint = $script:FileThumb } } } }

    function New-TestState {
        @{ Task = 'export'; Mailbox = 'adele.vance@contoso.com'; Shared = ''; ImShared = ''; Dest = 'pst'; OutFolder = 'D:\Exports\adele'; OutName = 'adele.pst'; SplitGB = 50; CheckEach = $true; SepAux = $false; Conc = 4; Tries = 3; Retry = 2; Reconcile = 5; Awake = $false
            ImSrc = 'pst'; ImPath = 'D:\PST\archive-2019.pst'; TgtLoc = 'Primary'; TgtMode = 'new'; TgtName = 'Imported archive'; Mapping = ''; RpPath = 'D:\PST\a.pst'; RpDest = 'D:\PST\a.repaired.pst'; RpLevel = 'Automatic'; RpKeep = $false
            OpPath = 'D:\PST\archive-2019.pst'; AuthMode = 'user'; TenantId = ''; ClientId = ''; AppTenant = ''; AppClient = ''; AppThumb = ''; Changed = @{} }
    }
    function New-TestJob {
        @{ Id = 'job-1'; Folder = 'D:\MailboxTransfer\jobs\20261112-1030-export-ab12'; EventsPath = 'D:\MailboxTransfer\jobs\20261112-1030-export-ab12\events.jsonl'; StopPath = 'D:\MailboxTransfer\jobs\20261112-1030-export-ab12\stop.request'
            StorePath = 'D:\MailboxTransfer\stores\adele.vance@contoso.com'; SelectionFiles = @('D:\MailboxTransfer\selections\export-7b843299.json'); JobName = 'import-ab12'; FolderFile = $null; ItemsFile = $null }
    }
    # a command line as the PowerShell parser reads it: the script, then the parameters with their values
    function ConvertTo-ParameterMap([string]$Line) {
        $errors = $null; $tokens = $null
        $ast = [Management.Automation.Language.Parser]::ParseInput($Line, [ref]$tokens, [ref]$errors)
        $errors | Should -BeNullOrEmpty -Because "the line parses: $Line"
        $command = $ast.Find({ param($n) $n -is [Management.Automation.Language.CommandAst] }, $true)
        $map = [ordered]@{}
        $elements = @($command.CommandElements)
        $elements[0].Extent.Text | Should -Be '.\Invoke-MailboxTransfer.ps1'
        for ($i = 1; $i -lt $elements.Count; $i++) {
            $e = $elements[$i]
            if ($e -isnot [Management.Automation.Language.CommandParameterAst]) { continue }
            $value = $null
            if (($i + 1) -lt $elements.Count -and $elements[$i + 1] -isnot [Management.Automation.Language.CommandParameterAst]) { $value = [string]$elements[$i + 1].Value; $i++ }
            $map[$e.ParameterName] = $value
        }
        return $map
    }
    function Get-RunMap($Command) {
        $map = [ordered]@{}
        foreach ($p in (Split-MbtGuiArgument -Argument @($Command.Arguments))) { $map[$p.Name] = $p.Value }
        return $map
    }
}

Describe 'The command line of a task' {
    It 'shows nothing of the plumbing of the console: no silent sign-in, events, stop file, job id, temporary folder, and no path inside the folder of the job' {
        foreach ($task in 'export', 'import', 'repair', 'copy') {
            $s = New-TestState; $s.Task = $task
            $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task $task -S $s -Job (New-TestJob)) -S $s
            foreach ($internal in '-NoPrompt', '-EventsFile', '-StopFile', '-JobId', '-TempPath', '-ConfigPath', '-LogFolder', '-DevelopmentMode', '-WhatIf') { $line | Should -Not -Match ([regex]::Escape($internal) + '(\s|$)') -Because "${task}: $internal is the console's own" }
            $line | Should -Not -Match '\\jobs\\' -Because "${task}: the folder of the job is the console's own"
        }
    }
    It 'leaves out every value that is what the script does without it, and keeps the others' {
        $s = New-TestState; $s.Tries = 3; $s.SplitGB = 50
        $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob)) -S $s
        $line | Should -Not -Match '-Concurrency'; $line | Should -Not -Match '-MaxAttempts'; $line | Should -Not -Match '-PstMaxSizeGB'; $line | Should -Not -Match '-KeepAwake'
        $s.Conc = 2; $s.Tries = 6; $s.SplitGB = 20; $s.Awake = $true
        $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob)) -S $s
        $line | Should -Match '-Concurrency 2'; $line | Should -Match '-MaxAttempts 6'; $line | Should -Match '-PstMaxSizeGB 20'; $line | Should -Match '-KeepAwake'
        $s = New-TestState; $s.Task = 'import'; $s.TgtLoc = 'Primary'; $s.Reconcile = 5
        $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task import -S $s -Job (New-TestJob)) -S $s
        $line | Should -Not -Match '-Mailbox '; $line | Should -Not -Match '-ReconcileDelaySeconds'
        $s.TgtLoc = 'Archive'; $s.Reconcile = 10
        $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task import -S $s -Job (New-TestJob)) -S $s
        $line | Should -Match '-Mailbox Archive'; $line | Should -Match '-ReconcileDelaySeconds 600'
    }
    It 'compares the values without regard to case or to the way a number is written' {
        foreach ($case in @{ Name = 'Mailbox'; Value = 'primary' }, @{ Name = 'Level'; Value = 'AUTOMATIC' }, @{ Name = 'PstMaxSizeGB'; Value = '50.0' }, @{ Name = 'ReconcileDelaySeconds'; Value = '300.00' }) {
            (ConvertTo-MbtGuiShownArgument -Argument @("-$($case.Name)", $case.Value)) | Should -BeNullOrEmpty -Because "$($case.Name) $($case.Value) is the default"
        }
        (ConvertTo-MbtGuiShownArgument -Argument @('-MAILBOX', 'archive')).Value | Should -Be 'archive'
        (ConvertTo-MbtGuiShownArgument -Argument @('-NOPROMPT', '-EVENTSFILE', 'x')) | Should -BeNullOrEmpty
    }
    It 'an export of several locations is one command: one selection file, the PST written by the same command, and nothing to chain' {
        $s = New-TestState
        $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob)) -S $s
        $line.Split("`n").Count | Should -Be 1
        $line | Should -Be '.\Invoke-MailboxTransfer.ps1 -Mode Export -UserPrincipalName adele.vance@contoso.com -SelectionFile D:\MailboxTransfer\selections\export-7b843299.json -StorePath D:\MailboxTransfer\stores\adele.vance@contoso.com -OutputPst D:\Exports\adele\adele.pst'
        $line | Should -Not -Match 'ExportPst'
    }
    It 'EML files are two lines (the export, then the files written from the store), a store alone is one' {
        $s = New-TestState; $s.Dest = 'eml'
        (ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob)) -S $s).Split("`n").Count | Should -Be 2
        $s.Dest = 'store'
        $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob)) -S $s
        $line.Split("`n").Count | Should -Be 1
        $line | Should -Not -Match '-OutputPst'
    }
    It 'imports a PST into the archive and repairs a PST as one short line each' {
        $s = New-TestState; $s.Task = 'import'; $s.TgtLoc = 'Archive'; $s.Awake = $true
        (ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task import -S $s -Job (New-TestJob)) -S $s) |
            Should -Be ".\Invoke-MailboxTransfer.ps1 -Mode ImportPst -UserPrincipalName adele.vance@contoso.com -JobName import-ab12 -PstPath D:\PST\archive-2019.pst -Mailbox Archive -TargetRoot 'Imported archive' -KeepAwake"
        $s = New-TestState; $s.Task = 'repair'
        (ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task repair -S $s -Job (New-TestJob)) -S $s) |
            Should -Be '.\Invoke-MailboxTransfer.ps1 -Mode RepairPst -Path D:\PST\a.pst -Destination D:\PST\a.repaired.pst'
    }
    It 'the folders ticked in a PST are named by a file that stays after the job, never one inside the folder of the job' {
        $s = New-TestState; $s.Task = 'import'
        $j = New-TestJob; $j.FolderFile = 'D:\MailboxTransfer\selections\import-1a2b3c4d.folders.txt'
        $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task import -S $s -Job $j) -S $s
        $line | Should -Match '-FolderFile D:\\MailboxTransfer\\selections\\import-1a2b3c4d\.folders\.txt'
        $base = Get-MbtGuiSelectionBase -WorkFolder 'D:\MailboxTransfer' -Task export -Key 'ABCDEF0123456789'
        $base | Should -Be 'D:\MailboxTransfer\selections\export-abcdef01'
        $folder = Join-Path $TestDrive 'work'
        $job = @{ Folder = (Join-Path $TestDrive 'job'); SelectionBase = (Join-Path $folder 'selections\import-1a2b3c4d') }
        Write-MbtGuiFolderFiles -Job $job -Choice ([pscustomobject]@{ Include = @('Inbox'); Exclude = @('Inbox/*'); Nids = @([uint32]5) })
        $job.FolderFile | Should -Be "$($job.SelectionBase).folders.txt"
        $job.ItemsFile | Should -Be "$($job.SelectionBase).items.txt"
        Get-Content -LiteralPath $job.FolderFile | Should -Be @('+Inbox', '-Inbox/*')
    }
    It 'a value with a space or a quote is quoted so that it is read back as it was' {
        $s = New-TestState; $s.Task = 'import'; $s.TgtName = "Imported Adele's mail 2026-11-12"
        $line = ConvertTo-MbtGuiCommandLine -Command @(Get-MbtGuiCommand -Task import -S $s -Job (New-TestJob)) -S $s
        (ConvertTo-ParameterMap $line)['TargetRoot'] | Should -Be "Imported Adele's mail 2026-11-12"
    }
    Context 'signing in' {
        It 'as me: the tenant and the application only when they were changed for this run' {
            $s = New-TestState
            $run = (Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob))[0]
            ConvertTo-MbtGuiCommandLine -Command @($run) -S $s | Should -Not -Match '-TenantId|-ClientId|-AppOnly|-CertificateThumbprint'
            $s.TenantId = 'CONTOSO.onmicrosoft.com'
            ConvertTo-MbtGuiCommandLine -Command @((Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob))[0]) -S $s | Should -Not -Match '-TenantId' -Because 'the case of a tenant is not a change'
            $s.TenantId = 'yourtenant.onmicrosoft.com'; $s.ClientId = '33333333-2222-3333-4444-555555555555'
            $line = ConvertTo-MbtGuiCommandLine -Command @((Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob))[0]) -S $s
            $line | Should -Match '-TenantId yourtenant\.onmicrosoft\.com'; $line | Should -Match '-ClientId 33333333-2222-3333-4444-555555555555'
        }
        It 'as the application with the values of the configuration file: -AppOnly, the file gives the rest' {
            $s = New-TestState; $s.AuthMode = 'app'; $s.AppTenant = $script:FileAppTenant.ToUpper(); $s.AppClient = $script:FileAppClient.ToUpper(); $s.AppThumb = ($script:FileThumb.ToLower() -replace '(..)', '$1 ').Trim()
            $run = (Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob))[0]
            (Get-RunMap $run).Keys | Should -Contain 'CertificateThumbprint' -Because 'the run gives the values it uses'
            $line = ConvertTo-MbtGuiCommandLine -Command @($run) -S $s
            $line | Should -Match '-AppOnly'
            $line | Should -Not -Match '-CertificateThumbprint|-TenantId|-ClientId'
        }
        It 'as the application with a value changed for this run: that value is on the line, the others still come from the file' {
            $s = New-TestState; $s.AuthMode = 'app'; $s.AppThumb = ('AB' * 20)
            $line = ConvertTo-MbtGuiCommandLine -Command @((Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob))[0]) -S $s
            $line | Should -Match "-CertificateThumbprint $('AB' * 20)"
            $line | Should -Not -Match '-AppOnly|-TenantId|-ClientId'
            $s = New-TestState; $s.AuthMode = 'app'; $s.AppClient = '33333333-2222-3333-4444-555555555555'
            $line = ConvertTo-MbtGuiCommandLine -Command @((Get-MbtGuiCommand -Task export -S $s -Job (New-TestJob))[0]) -S $s
            $line | Should -Match '-AppOnly'; $line | Should -Match '-ClientId 33333333-2222-3333-4444-555555555555'; $line | Should -Not -Match '-TenantId'
        }
    }
    Context 'the shown line and the run' {
        BeforeAll {
            $script:Cases = foreach ($name in 'export-pst', 'export-store', 'export-eml', 'import-pst', 'import-store', 'import-eml-archive', 'repair', 'copy', 'app') {
                $s = New-TestState; $j = New-TestJob
                switch ($name) {
                    'export-pst' { $s.Conc = 2; $s.Tries = 6; $s.SplitGB = 10; $s.CheckEach = $false; $s.SepAux = $true; $s.Awake = $true; $s.Changed['Retry'] = $true; $s.Retry = 7 }
                    'export-store' { $s.Dest = 'store' }
                    'export-eml' { $s.Dest = 'eml' }
                    'import-pst' { $s.Task = 'import'; $s.TgtLoc = 'Archive'; $s.Reconcile = 10; $j.FolderFile = 'D:\MailboxTransfer\selections\import-1a2b3c4d.folders.txt' }
                    'import-store' { $s.Task = 'import'; $s.ImSrc = 'store'; $s.ImPath = 'D:\Store'; $s.ImShared = 'projects@contoso.com'; $s.TgtMode = 'same' }
                    'import-eml-archive' { $s.Task = 'import'; $s.ImSrc = 'eml'; $s.ImPath = 'D:\eml'; $s.TgtLoc = 'Archive' }
                    'repair' { $s.Task = 'repair'; $s.RpLevel = 'Salvage'; $s.RpKeep = $true }
                    'copy' { $s.Task = 'copy'; $s.CheckEach = $false; $s.SplitGB = 5; $j.FolderFile = 'D:\MailboxTransfer\selections\copy-1a2b3c4d.folders.txt'; $j.ItemsFile = 'D:\MailboxTransfer\selections\copy-1a2b3c4d.items.txt' }
                    'app' { $s.AuthMode = 'app'; $s.AppThumb = ('CD' * 20); $s.AppClient = '33333333-2222-3333-4444-555555555555' }
                }
                @{ Name = $name; S = $s; Job = $j; Task = $s.Task }
            }
            $script:Parameters = (Get-Command $script:Entry).Parameters
        }
        It 'every parameter of the line is a parameter of the script: <Name>' -ForEach @(@{ Name = 'export-pst' }, @{ Name = 'export-store' }, @{ Name = 'export-eml' }, @{ Name = 'import-pst' }, @{ Name = 'import-store' }, @{ Name = 'import-eml-archive' }, @{ Name = 'repair' }, @{ Name = 'copy' }, @{ Name = 'app' }) {
            $case = $script:Cases | Where-Object Name -EQ $Name
            $commands = @(Get-MbtGuiCommand -Task $case.Task -S $case.S -Job $case.Job)
            foreach ($line in (ConvertTo-MbtGuiCommandLine -Command $commands -S $case.S).Split("`n")) {
                foreach ($parameter in (ConvertTo-ParameterMap $line).Keys) { $script:Parameters.ContainsKey($parameter) | Should -BeTrue -Because "-$parameter is a parameter of Invoke-MailboxTransfer.ps1" }
            }
        }
        It 'the line read back gives the same parameters and values as the run, minus the plumbing and what is the default: <Name>' -ForEach @(@{ Name = 'export-pst' }, @{ Name = 'export-store' }, @{ Name = 'export-eml' }, @{ Name = 'import-pst' }, @{ Name = 'import-store' }, @{ Name = 'import-eml-archive' }, @{ Name = 'repair' }, @{ Name = 'copy' }, @{ Name = 'app' }) {
            $case = $script:Cases | Where-Object Name -EQ $Name
            $commands = @(Get-MbtGuiCommand -Task $case.Task -S $case.S -Job $case.Job)
            $lines = @((ConvertTo-MbtGuiCommandLine -Command $commands -S $case.S).Split("`n"))
            $lines.Count | Should -Be $commands.Count
            for ($i = 0; $i -lt $commands.Count; $i++) {
                $shown = ConvertTo-ParameterMap $lines[$i]
                $expected = [ordered]@{}
                foreach ($p in (Get-RunMap $commands[$i]).GetEnumerator()) {
                    if ($p.Key -in $script:MbtGuiInternalParameter) { continue }
                    if ($null -ne $p.Value -and (Test-MbtGuiDefaultValue $p.Key $p.Value)) { continue }
                    $expected[$p.Key] = $p.Value
                }
                if ($case.Name -eq 'app') {
                    # the application of the file: the tenant and the thumbprint come from the configuration, the changed values are on the line
                    $expected.Remove('TenantId'); $expected['CertificateThumbprint'] | Should -Be ('CD' * 20)
                }
                if ($case.Name -in 'app') { $expected.Keys | Should -Contain 'ClientId' }
                @($shown.Keys | Sort-Object) | Should -Be @($expected.Keys | Sort-Object) -Because "$($case.Name): $($lines[$i])"
                foreach ($key in $expected.Keys) { $shown[$key] | Should -Be $expected[$key] -Because "-$key of $($case.Name)" }
            }
        }
        It 'the table of defaults is the one of the parameter block of the script' {
            $errors = $null; $tokens = $null
            $ast = [Management.Automation.Language.Parser]::ParseFile($script:Entry, [ref]$tokens, [ref]$errors)
            foreach ($name in $script:MbtGuiCliDefault.Keys) {
                $parameter = $ast.ParamBlock.Parameters | Where-Object { $_.Name.VariablePath.UserPath -eq $name }
                $parameter | Should -Not -BeNullOrEmpty -Because "-$name is a parameter"
                $parameter.DefaultValue | Should -Not -BeNullOrEmpty -Because "-$name has a default in the script"
                $text = $parameter.DefaultValue.Extent.Text.Trim("'")
                Test-MbtGuiDefaultValue $name $text | Should -BeTrue -Because "-$name is $text in the script and $($script:MbtGuiCliDefault[$name]) in the table"
            }
        }
        It 'the parameters the console keeps to itself are the ones that only make sense inside the console' {
            $script:MbtGuiInternalParameter | Should -Be @('NoPrompt', 'EventsFile', 'StopFile', 'JobId', 'TempPath', 'ConfigPath', 'LogFolder', 'DevelopmentMode')
        }
    }
}
