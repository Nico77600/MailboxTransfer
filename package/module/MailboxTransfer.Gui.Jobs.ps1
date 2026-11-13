#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - jobs (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    A job is what the console starts in a child pwsh: export, PST write, import, check, repair. Here: the job folder
    (<work folder>\jobs\<yyyyMMdd-HHmm>-<task>-<id>\ with request.json, events.jsonl, status.json, stop.request, stdout.txt, stderr.txt), the command lines
    (every value comes from a validated control, never from free text), the child process, the stop request, and the pump that reads the events file
    into the run model. The pump is also what the tests drive with a fake events file. What was chosen in the browse view (the selection of an export, the
    folders and messages ticked in a PST) is kept in <work folder>\selections\, outside the job folder, because the command line names those files.
    What is changed for one run (tenant, application, certificate, parallel requests, waits, tries) goes to the command line as parameters: nothing is
    written to the configuration file, and the line the console shows gives the same result as the one it runs.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

# The HTML report of a job (the same writer as the command line: report.html in the job folder).
. (Join-Path $PSScriptRoot 'MailboxTransfer.Report.ps1')

function New-MbtGuiJobId { return -join ((1..4) | ForEach-Object { '0123456789abcdef'[(Get-Random -Maximum 16)] }) }

function New-MbtGuiJobFolder {
    <# Creates the private job folder. Returns its path. #>
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)][string]$Task, [string]$Id = (New-MbtGuiJobId))
    $name = '{0}-{1}-{2}' -f (Get-Date).ToString('yyyyMMdd-HHmm'), $Task, $Id
    $dir = Join-Path (Join-Path $WorkFolder 'jobs') $name
    $null = New-Item -ItemType Directory -Force -Path $dir
    return $dir
}

function Test-MbtGuiPstName {
    <# A PST FILE name Windows or Outlook cannot take: a device name (CON, PRN, AUX, NUL, COM1-9, LPT1-9, with any extension), characters Windows refuses, a trailing dot or space. Returns the message in plain words, or $null when the name is fine. #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'Choose a file name.' }
    $n = $Name.Trim()
    if ($n -match '[<>:"/\\|?*\x00-\x1f]') { return 'This name has a character Windows does not accept (< > : " / \ | ? *). Choose another file name.' }
    if ($n.EndsWith('.') -or $Name.EndsWith(' ')) { return 'A file name cannot end with a dot or a space. Choose another file name.' }
    $stem = $n.Split('.')[0].TrimEnd(' ')
    if ($stem -match '^(?i)(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$') { return ('"{0}" is a name Windows keeps for itself. Choose another file name.' -f $stem.ToUpperInvariant()) }
    return $null
}

function Get-MbtGuiConfiguredSignIn {
    <# What the configuration file says about the tenant, the application and the certificate (the engine reads it); $null when the engine does not tell. #>
    $g = Get-Variable -Name Gui -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if (-not $g) { return $null }
    $has = if ($g.Engine -is [Collections.IDictionary]) { $g.Engine.Contains('SignInSettings') } else { [bool]$g.Engine.PSObject.Properties['SignInSettings'] }
    # a configuration file that cannot be read gives no values: the Setup page says so (a light), and the sign-in asks for them
    if ($has) { try { return (& $g.Engine.SignInSettings) } catch { return $null } }
    return $null
}

function Get-MbtGuiEffectiveSignIn {
    <#
        The values the sign-in uses for this run: the configuration file, over which a value typed on the Setup page (this run only) wins. Tenant, Client
        (and Thumb for the application) are the ones used; ...Changed says that the value differs from the file; Missing names what is empty.
    #>
    param([hashtable]$S)
    $cfg = Get-MbtGuiConfiguredSignIn
    $app = ($S['AuthMode'] -eq 'app')
    $typed = if ($app) { @([string]$S['AppTenant'], [string]$S['AppClient'], [string]$S['AppThumb']) } else { @([string]$S['TenantId'], [string]$S['ClientId'], '') }
    # the application is in the same tenant as the user in practice: without its own tenant, the tenant of the configuration is used (never the user, the mailbox or the mailbox ids)
    $appTenant = [string](Get-MbtGuiProp $cfg 'AppTenantId' ''); if (-not $appTenant.Trim()) { $appTenant = [string](Get-MbtGuiProp $cfg 'TenantId' '') }
    $file = if ($app) { @($appTenant, [string](Get-MbtGuiProp $cfg 'AppClientId' ''), [string](Get-MbtGuiProp $cfg 'CertificateThumbprint' '')) } else { @([string](Get-MbtGuiProp $cfg 'TenantId' ''), [string](Get-MbtGuiProp $cfg 'ClientId' ''), '') }
    $use = for ($i = 0; $i -lt 3; $i++) { if ($typed[$i].Trim()) { $typed[$i].Trim() } else { $file[$i].Trim() } }
    # not case sensitive; a thumbprint is compared by its digits (it is written with spaces or colons in the certificate details)
    $same = { param($a, $b, $thumbprint) if ($thumbprint) { ($a -replace '[^0-9A-Fa-f]', '') -eq ($b -replace '[^0-9A-Fa-f]', '') } else { $a.Trim() -eq $b.Trim() } }
    $changed = for ($i = 0; $i -lt 3; $i++) { [bool]($typed[$i].Trim() -and -not (& $same $typed[$i] $file[$i] ($app -and $i -eq 2))) }
    # an engine that does not say what its configuration holds is not asked questions: nothing is reported missing
    $missing = [Collections.Generic.List[string]]::new()
    if ($null -ne $cfg) {
        if (-not $use[0]) { $missing.Add('Tenant') }
        if (-not $use[1]) { $missing.Add('Application id') }
        if ($app -and -not $use[2]) { $missing.Add('Certificate thumbprint') }
    }
    [pscustomobject]@{ App = $app; Tenant = $use[0]; Client = $use[1]; Thumb = $use[2]; TenantChanged = $changed[0]; ClientChanged = $changed[1]; ThumbChanged = $changed[2]
        FromFile = $file; Missing = $missing.ToArray(); Changed = [bool]($changed -contains $true) }
}

function Get-MbtGuiAuthArguments {
    <# The sign-in of the child job: nothing for "me" (the cached token; the tenant and the application go too when they were changed for this run), the application's certificate for "the application" (run-only: kept in the console, never written to a file). #>
    param([hashtable]$S)
    $e = Get-MbtGuiEffectiveSignIn -S $S
    if (-not $e.App) {
        if ($e.TenantChanged -or $e.ClientChanged) { '-TenantId'; $e.Tenant; '-ClientId'; $e.Client }
        return
    }
    '-CertificateThumbprint'; ($e.Thumb -replace '[^0-9A-Fa-f]', ''); '-TenantId'; $e.Tenant; '-ClientId'; $e.Client
}
function ConvertTo-MbtGuiFolderPattern {
    <# A folder path as a pattern of the engine (-IncludeFolder / -ExcludeFolder): a * or ? in a folder name is written literally (backtick), slashes stay. #>
    param([string]$Path)
    return ([regex]::Replace($Path, '([`*?])', '`$1'))
}

function Get-MbtGuiFolderChoice {
    <#
        What is ticked in the tree of a PST or of a folder of .eml files, as the engine takes it: Include and Exclude (folder paths below the top of the source) and Nids (messages ticked
        one by one in a PST, by node id). A folder ticked as a whole is its path (the engine takes what is below a path); a folder's own messages only is its path with its subfolders
        excluded. Nothing ticked: all three empty (everything is taken).
    #>
    param([Parameter(Mandatory)]$Selection, [Parameter(Mandatory)][ValidateSet('pst', 'eml')][string]$Source)
    $inc = [Collections.Generic.List[string]]::new(); $exc = [Collections.Generic.List[string]]::new(); $nids = [Collections.Generic.List[uint32]]::new()
    foreach ($e in @($Selection.Entries)) {
        $tag = $e.Node.Tag
        $path = [string](Get-MbtGuiProp $tag $(if ($Source -eq 'pst') { 'FolderPath' } else { 'Rel' }) '')
        $pattern = ConvertTo-MbtGuiFolderPattern $path
        switch ($e.Kind) {
            'tree' { $inc.Add($pattern) }
            'own' { $inc.Add($pattern); $exc.Add("$pattern/*") }
            default { foreach ($row in @($e.Node.PickRows.Values)) { $n = Get-MbtGuiProp $row 'Nid'; if ($null -ne $n) { $nids.Add([uint32]$n) } } }
        }
    }
    [pscustomobject]@{ Include = $inc.ToArray(); Exclude = $exc.ToArray(); Nids = $nids.ToArray() }
}

function Get-MbtGuiSelectionBase {
    <# Where the choice of an operation is kept, without extension: <work folder>\selections\<task>-<8 characters of what was chosen>. The command line names these files. #>
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)][string]$Task, [Parameter(Mandatory)][string]$Key)
    return [IO.Path]::Combine($WorkFolder, 'selections', ('{0}-{1}' -f $Task, $Key.Substring(0, 8).ToLowerInvariant()))
}

function Write-MbtGuiFolderFiles {
    <# Writes the choice next to the selection: <base>.folders.txt (lines "+path" take, "-path" leave out; -FolderFile of the entry script) and <base>.items.txt (node ids; -IncludeItemFile). Sets FolderFile / ItemsFile of the job. #>
    param([Parameter(Mandatory)][hashtable]$Job, [Parameter(Mandatory)]$Choice)
    $Job.FolderFile = $null; $Job.ItemsFile = $null
    $base = if ($Job.ContainsKey('SelectionBase') -and $Job.SelectionBase) { $Job.SelectionBase } else { Join-Path $Job.Folder 'choice' }
    $lines = @($Choice.Include | ForEach-Object { "+$_" }) + @($Choice.Exclude | ForEach-Object { "-$_" })
    if ($lines.Count -or @($Choice.Nids).Count) { $null = New-Item -ItemType Directory -Force -Path (Split-Path $base -Parent) }
    if ($lines.Count) { $Job.FolderFile = "$base.folders.txt"; Set-Content -LiteralPath $Job.FolderFile -Value $lines -Encoding utf8 }
    if (@($Choice.Nids).Count) { $Job.ItemsFile = "$base.items.txt"; Set-Content -LiteralPath $Job.ItemsFile -Value @($Choice.Nids | ForEach-Object { [string]$_ }) -Encoding utf8 }
}

function Get-MbtGuiCommonArguments {
    param([hashtable]$S, [hashtable]$Job, [switch]$WhatIf)
    $a = [Collections.Generic.List[string]]::new()
    $a.AddRange([string[]]@(Get-MbtGuiAuthArguments -S $S))
    $a.AddRange([string[]]@('-EventsFile', $Job.EventsPath, '-StopFile', $Job.StopPath, '-JobId', $Job.Id))
    if ($S.Awake) { $a.Add('-KeepAwake') }
    if ($WhatIf) { $a.Add('-WhatIf') }
    return , $a
}

function Get-MbtGuiCommand {
    <#
        The command lines of a task (one or several, run one after the other). Returns objects with Arguments (for Invoke-MailboxTransfer.ps1) and
        Title. Every value comes from a validated control of the console. -Plan: the read-only run (-WhatIf) of the same request. An export of
        several locations is one command: the selection file holds one selection per location, and the PST files are written by the same command.
    #>
    param(
        [Parameter(Mandatory)][ValidateSet('export', 'import', 'repair', 'check', 'copy')][string]$Task,
        [Parameter(Mandatory)][hashtable]$S, [Parameter(Mandatory)][hashtable]$Job, [switch]$Plan
    )
    $common = Get-MbtGuiCommonArguments -S $S -Job $Job -WhatIf:$Plan
    $out = [Collections.Generic.List[object]]::new()
    $tuning = [Collections.Generic.List[string]]::new()
    $tuning.AddRange([string[]]@('-Concurrency', [string][int]$S.Conc, '-MaxAttempts', [string][int]$S.Tries))
    if ($S.Changed.ContainsKey('Retry')) { $tuning.AddRange([string[]]@('-RetryWaitSeconds', [string][int]$S.Retry)) }
    switch ($Task) {
        'export' {
            $pst = $S.Dest -eq 'pst'
            $eml = $S.Dest -eq 'eml'
            $a = [Collections.Generic.List[string]]::new()
            # the temporary files of the export are in the job folder, not below the folder of the script
            $a.AddRange([string[]]@('-Mode', 'Export', '-UserPrincipalName', $S.Mailbox))
            foreach ($file in @($Job.SelectionFiles)) { $a.AddRange([string[]]@('-SelectionFile', $file)) }
            $a.AddRange([string[]]@('-StorePath', $Job.StorePath, '-TempPath', [IO.Path]::Combine($Job.Folder, 'temp'), '-NoPrompt'))
            $a.AddRange([string[]]$tuning)
            if ($pst) {
                $a.AddRange([string[]]@('-OutputPst', [IO.Path]::Combine($S.OutFolder, $S.OutName), '-PstMaxSizeGB', [string][int]$S.SplitGB))
                if (-not $S.CheckEach) { $a.Add('-SkipCheck') }
                if ($S.SepAux) { $a.Add('-SeparateAuxArchives') }
            }
            $a.AddRange([string[]]$common)
            $out.Add([pscustomobject]@{ Title = 'Export'; Arguments = $a.ToArray() })
            if ($eml) {
                # the items come to the export store first (resumable, checked), then each message is written as an .eml file below the folder chosen
                $a = [Collections.Generic.List[string]]::new()
                $a.AddRange([string[]]@('-Mode', 'ExportEml', '-StorePath', $Job.StorePath, '-EmlPath', $S.OutFolder, '-Concurrency', [string][int]$S.Conc))
                $a.AddRange([string[]]$common)
                $out.Add([pscustomobject]@{ Title = 'Write the EML files'; Arguments = $a.ToArray() })
            }
        }
        'import' {
            $a = [Collections.Generic.List[string]]::new()
            $mode = switch ($S.ImSrc) { 'pst' { 'ImportPst' } 'eml' { 'ImportEml' } default { 'Import' } }
            $a.AddRange([string[]]@('-Mode', $mode, '-UserPrincipalName', $S.Mailbox, '-JobName', $Job.JobName, '-NoPrompt'))
            switch ($S.ImSrc) {
                # the journal of a PST import lives in the job folder: Continue (same folder) finds it again
                'pst' { $a.AddRange([string[]]@('-PstPath', $S.ImPath, '-TempPath', [IO.Path]::Combine($Job.Folder, 'work'))) }
                'eml' { $a.AddRange([string[]]@('-EmlPath', $S.ImPath)) }
                default { $a.AddRange([string[]]@('-StorePath', $S.ImPath)) }
            }
            $a.AddRange([string[]]@('-Mailbox', $S.TgtLoc))
            # a shared mailbox (or another user's) is given by its address
            if ($S.ImShared) { $a.AddRange([string[]]@('-MailboxId', $S.ImShared)) }
            if ($S.TgtMode -eq 'new' -and $S.TgtName) {
                $problem = Test-MbtGuiFolderName $S.TgtName
                if ($problem) { throw "The target folder name is refused: $problem" }
                $a.AddRange([string[]]@('-TargetRoot', $S.TgtName))
            }
            elseif ($S.ImSrc -eq 'pst') { $a.Add('-MergeWellKnown') }
            if ($S.Mapping -and $S.ImSrc -ne 'eml') { $a.AddRange([string[]]@('-MappingFile', $S.Mapping)) }
            if ($Job.SelectionFiles -and $S.ImSrc -eq 'store') { $a.AddRange([string[]]@('-SelectionFile', @($Job.SelectionFiles)[0])) }
            # the folders ticked in a PST or in a folder of .eml files
            if ((Get-MbtGuiProp $Job 'FolderFile') -and $S.ImSrc -in 'pst', 'eml') { $a.AddRange([string[]]@('-FolderFile', (Get-MbtGuiProp $Job 'FolderFile'))) }
            $a.AddRange([string[]]@('-ReconcileDelaySeconds', [string]([int]$S.Reconcile * 60)))
            $a.AddRange([string[]]$tuning)
            $a.AddRange([string[]]$common)
            $out.Add([pscustomobject]@{ Title = 'Import'; Arguments = $a.ToArray() })
        }
        'copy' {
            # the ticks of an open PST written into new PST files: -Mode ExportPst reads a PST instead of a store
            $a = [Collections.Generic.List[string]]::new()
            $a.AddRange([string[]]@('-Mode', 'ExportPst', '-PstPath', $S.OpPath, '-OutputPst', [IO.Path]::Combine($S.OutFolder, $S.OutName), '-PstMaxSizeGB', [string][int]$S.SplitGB, '-TempPath', [IO.Path]::Combine($Job.Folder, 'temp')))
            if (-not $S.CheckEach) { $a.Add('-SkipCheck') }
            if (Get-MbtGuiProp $Job 'FolderFile') { $a.AddRange([string[]]@('-FolderFile', (Get-MbtGuiProp $Job 'FolderFile'))) }
            if (Get-MbtGuiProp $Job 'ItemsFile') { $a.AddRange([string[]]@('-IncludeItemFile', (Get-MbtGuiProp $Job 'ItemsFile'))) }
            $a.AddRange([string[]]$common)
            $out.Add([pscustomobject]@{ Title = 'Write the PST files'; Arguments = $a.ToArray() })
        }
        'repair' {
            $a = [Collections.Generic.List[string]]::new()
            $a.AddRange([string[]]@('-Mode', 'RepairPst', '-Path', $S.RpPath, '-Destination', $S.RpDest, '-Level', $S.RpLevel))
            if ($S.RpKeep) { $a.Add('-KeepDamagedAttachments') }
            $a.AddRange([string[]]$common)
            $out.Add([pscustomobject]@{ Title = 'Repair'; Arguments = $a.ToArray() })
        }
        'check' {
            $a = [Collections.Generic.List[string]]::new()
            $a.AddRange([string[]]@('-Mode', 'CheckPst', '-Path', $S.RpPath))
            $a.AddRange([string[]]$common)
            $out.Add([pscustomobject]@{ Title = 'Check'; Arguments = $a.ToArray() })
        }
    }
    return $out.ToArray()
}

#region The command line an operator can run ------------------------------------------------------------------------------------------------

# What the console passes for its own sake: not an operator's choice, so not in the line that is shown.
$script:MbtGuiInternalParameter = @('NoPrompt', 'EventsFile', 'StopFile', 'JobId', 'TempPath', 'ConfigPath', 'LogFolder', 'DevelopmentMode')
# The value a parameter of Invoke-MailboxTransfer.ps1 has when it is not given (the tests compare this table with the parameter block of the script).
$script:MbtGuiCliDefault = @{ Mailbox = 'Primary'; Concurrency = '4'; MaxAttempts = '3'; ReconcileDelaySeconds = '300'; ReconcileRounds = '3'; PstMaxSizeGB = '50'; Level = 'Automatic'; RepairLevel = 'Rebuild' }

function Split-MbtGuiArgument {
    <# The flat arguments of a command as pairs: Name (without the dash), Value ($null for a switch). #>
    param([string[]]$Argument)
    $pairs = [Collections.Generic.List[object]]::new()
    $list = @($Argument)
    for ($i = 0; $i -lt $list.Count; $i++) {
        $token = [string]$list[$i]
        if ($token -notmatch '^-[A-Za-z]+$') { continue }
        $value = $null
        if (($i + 1) -lt $list.Count -and [string]$list[$i + 1] -notmatch '^-[A-Za-z]+$') { $value = [string]$list[$i + 1]; $i++ }
        $pairs.Add([pscustomobject]@{ Name = $token.Substring(1); Value = $value })
    }
    return $pairs.ToArray()
}

function Test-MbtGuiDefaultValue {
    param([string]$Name, [string]$Value)
    if (-not $script:MbtGuiCliDefault.ContainsKey($Name)) { return $false }
    $default = $script:MbtGuiCliDefault[$Name]
    $x = 0.0; $y = 0.0
    if ([double]::TryParse($Value, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$x) -and [double]::TryParse($default, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$y)) { return ($x -eq $y) }
    return ($Value -ieq $default)
}

function ConvertTo-MbtGuiShownArgument {
    <#
        The arguments the console runs, as the arguments an operator would type: without the plumbing of the console (silent sign-in, events, stop file, job id,
        temporary folder) and without a value that is what the script does anyway. The sign-in of the application is written the short way when it is the one of
        the configuration file (-AppOnly). Everything else is kept, so the line gives the same result as the run.
    #>
    param([Parameter(Mandatory)][string[]]$Argument, [hashtable]$S)
    $sign = if ($S) { Get-MbtGuiEffectiveSignIn -S $S } else { $null }
    $out = [Collections.Generic.List[object]]::new()
    foreach ($p in (Split-MbtGuiArgument $Argument)) {
        if ($p.Name -in $script:MbtGuiInternalParameter) { continue }
        if ($null -ne $p.Value -and (Test-MbtGuiDefaultValue $p.Name $p.Value)) { continue }
        if ($sign -and $p.Name -eq 'CertificateThumbprint' -and -not $sign.ThumbChanged) { $out.Add([pscustomobject]@{ Name = 'AppOnly'; Value = $null }); continue }
        if ($sign -and $p.Name -eq 'TenantId' -and -not $sign.TenantChanged) { continue }
        if ($sign -and $p.Name -eq 'ClientId' -and -not $sign.ClientChanged) { continue }
        $out.Add($p)
    }
    return $out.ToArray()
}

function Format-MbtGuiCommandToken {
    param([string]$Text)
    if ($Text -match '[\s;&()$''`"]') { return ("'" + $Text.Replace("'", "''") + "'") }
    return $Text
}

function ConvertTo-MbtGuiCommandLine {
    <# One line of PowerShell per command of the task, as an operator would type it in the folder of the tool. #>
    param([Parameter(Mandatory)][object[]]$Command, [hashtable]$S)
    $lines = foreach ($c in $Command) {
        $tokens = [Collections.Generic.List[string]]::new(); $tokens.Add('.\Invoke-MailboxTransfer.ps1')
        foreach ($p in (ConvertTo-MbtGuiShownArgument -Argument @($c.Arguments) -S $S)) {
            $tokens.Add('-' + $p.Name)
            if ($null -ne $p.Value) { $tokens.Add((Format-MbtGuiCommandToken $p.Value)) }
        }
        $tokens -join ' '
    }
    return ($lines -join "`n")
}

#endregion

function New-MbtGuiJob {
    <# The job object: folder, paths of its files, the commands still to run, the child process. #>
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)][string]$Task, [string]$EntryScript, [string]$StorePath, [string[]]$SelectionFiles = @(), [string]$JobName, [string]$SelectionBase)
    $id = New-MbtGuiJobId
    $dir = New-MbtGuiJobFolder -WorkFolder $WorkFolder -Task $Task -Id $id
    $job = @{
        Id = "$Task-$id"; Task = $Task; Folder = $dir; EntryScript = $EntryScript
        EventsPath = Join-Path $dir 'events.jsonl'; StopPath = Join-Path $dir 'stop.request'; StatusPath = Join-Path $dir 'status.json'
        StdOut = Join-Path $dir 'stdout.txt'; StdErr = Join-Path $dir 'stderr.txt'; SelectionBase = $SelectionBase
        StorePath = $(if ($StorePath) { $StorePath } else { Join-Path $WorkFolder 'stores\default' }); SelectionFiles = $SelectionFiles
        JobName = $(if ($JobName) { $JobName } else { "import-$id" })
        FolderFile = $null; ItemsFile = $null
        Offset = 0L; Process = $null; Queue = [Collections.Generic.Queue[object]]::new(); Started = [datetime]::UtcNow; StopAsked = $null; Exited = $false; ExitCode = $null; Fake = $false
    }
    return $job
}

function New-MbtGuiProcessStartInfo {
    <#
        The start of a child pwsh: every argument is one element of ProcessStartInfo.ArgumentList, so the runtime does the quoting (a space, a quote,
        a trailing backslash or a value typed by the user cannot change what the child receives). -Redirect: standard input, output and error are pipes.
    #>
    param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Argument, [string]$WorkingDirectory, [switch]$Redirect)
    $psi = [Diagnostics.ProcessStartInfo]::new((Get-Process -Id $PID).Path)
    foreach ($a in $Argument) { $psi.ArgumentList.Add($a) }
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    if ($WorkingDirectory) { $psi.WorkingDirectory = $WorkingDirectory }
    if ($Redirect) { $psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true }
    return $psi
}

function Start-MbtGuiDetached {
    <# Starts a child pwsh that nobody waits for (a short command whose output is not needed). #>
    param([Parameter(Mandatory)][AllowEmptyString()][string[]]$Argument)
    $psi = New-MbtGuiProcessStartInfo -Argument $Argument
    $psi.RedirectStandardInput = $true
    $p = [Diagnostics.Process]::Start($psi)
    $p.StandardInput.Close()
    $p.Dispose()
}

function Test-MbtGuiFolderName {
    <#
        The name typed for the new folder (a path below the mailbox, segments separated by "/"): returns the reason it is refused, or $null.
        It is given to the engine as a command line value, so a leading "-" (taken for a parameter) and control characters are refused.
    #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'The folder needs a name.' }
    if ($Name.Length -gt 1024) { return 'The folder name is too long.' }
    if ($Name.TrimStart().StartsWith('-')) { return 'The folder name cannot start with a dash.' }
    if ($Name -match '[\x00-\x1F\x7F]') { return 'The folder name has a character that cannot be used (a control character).' }
    foreach ($segment in $Name.Split('/')) {
        $t = $segment.Trim()
        if ($t -eq '') { return 'The folder name has an empty part (two slashes in a row, or a slash at the start or the end).' }
        if ($t -eq '.' -or $t -eq '..') { return 'The folder name cannot have "." or ".." as a part.' }
        if ($t.Length -gt 255) { return 'A part of the folder name is too long (255 characters at most).' }
    }
    return $null
}

function Close-MbtGuiProcessPipes {
    <# Ends the copy of the output of the last child to its files (waits a few seconds for the end of the data, then lets go). #>
    param([Parameter(Mandatory)][hashtable]$Job)
    $pipes = if ($Job.ContainsKey('Pipes')) { @($Job.Pipes) } else { @() }
    foreach ($p in $pipes) {
        if (-not $p) { continue }
        try { $null = $p.Task.Wait(3000) } catch { Write-Verbose "pipe: $_" }
        try { $p.File.Dispose() } catch { Write-Verbose "pipe: $_" }
    }
    $Job.Pipes = @()
}

function Start-MbtGuiProcess {
    <# Starts the next command of the job in a child pwsh: no window, standard input closed, output to files. #>
    param([Parameter(Mandatory)][hashtable]$Job)
    if ($Job.Queue.Count -eq 0) { return $false }
    $cmd = $Job.Queue.Dequeue()
    $exe = (Get-Process -Id $PID).Path
    $arguments = [Collections.Generic.List[string]]::new()
    $arguments.AddRange([string[]]@('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $Job.EntryScript))
    if ($script:GuiDevelopmentMode) { $arguments.Add('-DevelopmentMode') }
    # the log file of the run is in the logs folder of the work folder, next to the jobs, not in the package
    $arguments.AddRange([string[]]@('-LogFolder', (Join-Path (Split-Path (Split-Path $Job.Folder -Parent) -Parent) 'logs')))
    $arguments.AddRange([string[]]$cmd.Arguments)
    Close-MbtGuiProcessPipes -Job $Job
    $psi = New-MbtGuiProcessStartInfo -Argument $arguments.ToArray() -WorkingDirectory (Split-Path $Job.EntryScript -Parent) -Redirect
    $process = [Diagnostics.Process]::Start($psi)
    $process.StandardInput.Close()
    $pipes = foreach ($pair in @(@($process.StandardOutput.BaseStream, $Job.StdOut), @($process.StandardError.BaseStream, $Job.StdErr))) {
        $file = [IO.FileStream]::new($pair[1], [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
        @{ File = $file; Task = $pair[0].CopyToAsync($file) }
    }
    $Job.Pipes = @($pipes)
    $Job.Process = $process
    $Job.Exited = $false; $Job.ExitCode = $null
    return $true
}

function Request-MbtGuiStop {
    <# "Stop and keep progress": the engine polls the stop request file and cancels like Ctrl+C. #>
    param([Parameter(Mandatory)][hashtable]$Job)
    Set-Content -LiteralPath $Job.StopPath -Value ([datetime]::UtcNow.ToString('o'))
    $Job.StopAsked = [datetime]::UtcNow
}

function Stop-MbtGuiProcessTree {
    <# "Force stop" (offered after 10 minutes): safe, the store crash-recovers; import rows in flight become Unknown. #>
    param([Parameter(Mandatory)][hashtable]$Job)
    if ($Job.Process -and -not $Job.Process.HasExited) { try { $Job.Process.Kill($true) } catch { Write-Verbose "kill: $_" } }
}

function Merge-MbtGuiChainResult {
    <# The result of the last command of a chain, with the worst exit of the commands before it: a first command that ended resumable (2) is not complete because the last one is. #>
    param([Parameter(Mandatory)][hashtable]$Job, [Parameter(Mandatory)]$Event)
    if ((Get-MbtGuiProp $Event 'e') -ne 'result') { return $Event }
    $exit = [int](Get-MbtGuiProp $Event 'exit' 1)
    if ($exit -eq 0 -and (Get-MbtGuiProp $Job 'PartExit') -eq 2) {
        return [pscustomobject]@{ e = 'result'; exit = 2; complete = $false; counts = (Get-MbtGuiProp $Event 'counts' $null); next = 'continue'; reason = [string](Get-MbtGuiProp $Event 'reason' '') }
    }
    return $Event
}

function Step-MbtGuiJob {
    <#
        One tick of the job: new event lines into the run model, and the end of the child process. Returns the number of events applied.
        A child that ends without a result event (killed, crashed) gives a result of its exit code, so the page never waits for ever.
    #>
    param([Parameter(Mandatory)][hashtable]$Job, [Parameter(Mandatory)]$Run, [int]$MaxLines = 400)
    $n = 0
    $tail = Read-MbtGuiEventTail -Path $Job.EventsPath -Offset $Job.Offset -MaxLines $MaxLines
    $Job.Offset = $tail.Offset
    foreach ($line in $tail.Lines) {
        $evt = ConvertFrom-MbtGuiEventLine $line
        if (-not $evt) { continue }
        # a chain of commands: only the last result is the result of the job
        if ((Get-MbtGuiProp $evt 'e') -eq 'result' -and $Job.Queue.Count -gt 0 -and [int](Get-MbtGuiProp $evt 'exit' 1) -in 0, 2) { $Job.PartDone = $true; if ([int](Get-MbtGuiProp $evt 'exit' 1) -eq 2) { $Job.PartExit = 2 }; continue }
        $null = Add-MbtGuiEvent -Run $Run -Event (Merge-MbtGuiChainResult $Job $evt)
        $n++
    }
    if ($Job.Process -and -not $Job.Exited) {
        if ($Job.Process.HasExited) {
            $Job.Exited = $true; $Job.ExitCode = $Job.Process.ExitCode
            Close-MbtGuiProcessPipes -Job $Job
            # whatever the child wrote last has been read above only if it was flushed: read once more
            $tail = Read-MbtGuiEventTail -Path $Job.EventsPath -Offset $Job.Offset
            $Job.Offset = $tail.Offset
            foreach ($line in $tail.Lines) {
                $evt = ConvertFrom-MbtGuiEventLine $line
                if (-not $evt) { continue }
                if ((Get-MbtGuiProp $evt 'e') -eq 'result' -and $Job.Queue.Count -gt 0 -and [int](Get-MbtGuiProp $evt 'exit' 1) -in 0, 2) { if ([int](Get-MbtGuiProp $evt 'exit' 1) -eq 2) { $Job.PartExit = 2 }; continue }
                $null = Add-MbtGuiEvent -Run $Run -Event (Merge-MbtGuiChainResult $Job $evt); $n++
            }
            $continueChain = $Job.Queue.Count -gt 0 -and $Job.ExitCode -in 0, 2 -and -not $Job.StopAsked
            if ($continueChain) { if ($Job.ExitCode -eq 2) { $Job.PartExit = 2 }; $null = Start-MbtGuiProcess -Job $Job }
            elseif (-not $Run.Result) {
                $exit = [int]$Job.ExitCode
                $fallback = [pscustomobject]@{ e = 'result'; exit = $(if ($exit -eq 0) { 0 } elseif ($Job.StopAsked) { 2 } else { 1 }); complete = ($exit -eq 0); next = $(if ($Job.StopAsked) { 'continue' } else { '' }) }
                $null = Add-MbtGuiEvent -Run $Run -Event (Merge-MbtGuiChainResult $Job $fallback)
                $n++
            }
        }
    }
    return $n
}

function Get-MbtGuiStopWait {
    <# Seconds since "Stop and keep progress" was pressed, or 0. #>
    param([hashtable]$Job)
    if (-not $Job -or -not $Job.StopAsked) { return 0 }
    return [int]([datetime]::UtcNow - $Job.StopAsked).TotalSeconds
}

function ConvertTo-MbtGuiCsvRow {
    <# A row for the CSV: a text that starts with = + - @ (or a tab or a carriage return) is taken for a formula by a spreadsheet, so it gets a leading apostrophe. #>
    param([Parameter(Mandatory, ValueFromPipeline)]$Row)
    process {
        $copy = [ordered]@{}
        foreach ($p in $Row.PSObject.Properties) { $copy[$p.Name] = if ($p.Value -is [string] -and $p.Value -match '^[=+\-@\t\r]') { "'" + $p.Value } else { $p.Value } }
        [pscustomobject]$copy
    }
}

function Write-MbtGuiRunReport {
    <#
        The report of a finished run: report.json and report.csv in the job folder, and report.html (the chronology of the job, written from its events.jsonl, which links the CSV and the JSON);
        for a repair also <new file>.repair-report.csv / .json next to the new file. Returns the paths written.
    #>
    param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)][string]$Folder, [string]$Task, [string]$RepairDestination)
    $null = New-Item -ItemType Directory -Force -Path $Folder
    $rows = [Collections.Generic.List[object]]::new()
    foreach ($s in $Run.Steps) { $rows.Add([pscustomobject]@{ Kind = 'step'; Step = $s.Id; Key = $s.Title; State = $s.State; Done = $s.Done; Total = $s.Total; Message = $s.Msg }) }
    foreach ($o in $Run.Objects) { $rows.Add([pscustomobject]@{ Kind = 'folder'; Step = $o.Step; Key = $o.Key; State = $o.State; Done = $o.Done; Total = $o.Total; Message = $o.Msg }) }
    foreach ($i in $Run.Items) { $rows.Add([pscustomobject]@{ Kind = 'item'; Step = $i.Step; Key = $i.Key; State = $i.State; Done = ''; Total = $i.Total; Message = (($i.Msg, $i.Detail | Where-Object { $_ }) -join ' - ') }) }
    foreach ($w in $Run.Warnings) { $rows.Add([pscustomobject]@{ Kind = 'warning'; Step = $w.Step; Key = $w.Key; State = $w.Code; Done = ''; Total = ''; Message = $w.Msg }) }
    foreach ($d in $Run.Decisions) { $rows.Add([pscustomobject]@{ Kind = 'decision'; Step = ''; Key = $d.Key; State = 'needs a decision'; Done = ''; Total = ''; Message = $d.Msg }) }
    if ($Run.Error) { $rows.Add([pscustomobject]@{ Kind = 'error'; Step = ''; Key = $Run.Error.Code; State = 'error'; Done = ''; Total = ''; Message = $Run.Error.Msg }) }
    $summary = [ordered]@{
        task = $Task; title = $Run.Title; state = $Run.State; started = $(if ($Run.Started) { $Run.Started.ToString('o') } else { $null }); finished = $(if ($Run.LastEvent) { $Run.LastEvent.ToString('o') } else { $null })
        exitCode = $(if ($Run.Result) { $Run.Result.Exit } else { $null }); counts = $(if ($Run.Result) { $Run.Result.Counts } else { $null }); throttleWaits = $Run.Waits.Count; throttleSecondsWaited = $Run.Waits.Seconds
        rows = $rows.ToArray()
    }
    $written = [Collections.Generic.List[string]]::new()
    $targets = @(@{ Json = (Join-Path $Folder 'report.json'); Csv = (Join-Path $Folder 'report.csv') })
    if ($Task -eq 'repair' -and $RepairDestination) {
        $stem = [IO.Path]::Combine((Split-Path $RepairDestination -Parent), [IO.Path]::GetFileNameWithoutExtension($RepairDestination))
        if (Test-Path -LiteralPath (Split-Path $RepairDestination -Parent)) { $targets += @{ Json = "$stem.repair-report.json"; Csv = "$stem.repair-report.csv" } }
    }
    foreach ($t in $targets) {
        $summary | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $t.Json -Encoding utf8
        $rows | ConvertTo-MbtGuiCsvRow | Export-Csv -LiteralPath $t.Csv -NoTypeInformation -Encoding utf8
        $written.Add($t.Json); $written.Add($t.Csv)
    }
    if (Test-Path -LiteralPath (Join-Path $Folder 'events.jsonl') -PathType Leaf) {
        try { $written.Add((Write-MbtJobHtmlReport -Folder $Folder)) } catch { Write-Verbose "The HTML report could not be written: $($_.Exception.Message)" }
    }
    return $written.ToArray()
}

function Get-MbtGuiReportPath {
    <#
        The report the Done page opens: report.html when the job has one (written by the command line at the end of the job, and again by the console beside the CSV), else report.csv
        (a job of an older version has only the CSV). -Reports are the files the console wrote for this run, -Folder the job folder. Null without either.
    #>
    param([string[]]$Reports = @(), [string]$Folder)
    $html = @($Reports | Where-Object { $_ -like '*.html' } | Select-Object -First 1)
    if ($html.Count -and (Test-Path -LiteralPath $html[0] -PathType Leaf)) { return $html[0] }
    if ($Folder) {
        $inFolder = [IO.Path]::Combine($Folder, 'report.html')
        if (Test-Path -LiteralPath $inFolder -PathType Leaf) { return $inFolder }
    }
    $csv = @($Reports | Where-Object { $_ -like '*.csv' } | Select-Object -First 1)
    if ($csv.Count) { return $csv[0] }
    if ($Folder) { return [IO.Path]::Combine($Folder, 'report.csv') }
    return $null
}

