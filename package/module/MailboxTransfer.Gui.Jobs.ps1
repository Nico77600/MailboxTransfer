#Requires -Version 7.6
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

# The HTML report of a job (the same writer as the command line: report.html in the job folder) and the overview of all the jobs.
. (Join-Path $PSScriptRoot 'MailboxTransfer.Report.ps1')
. (Join-Path $PSScriptRoot 'MailboxTransfer.Overview.ps1')

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

function Get-MbtGuiPeriod {
    <#
        The period the operator asked for an export, from the state: Chosen (the second choice is on), Active (and it is valid: it holds a date and the first date is not after the last),
        From and To (yyyy-MM-dd, empty: no limit), Problem (in plain words, or $null) and Text ("arrived between 2026-03-01 and 2026-03-31", "all dates").
    #>
    param([hashtable]$S)
    $chosen = ((Get-MbtGuiProp $S 'Task' '') -in 'export', 'mailcopy') -and ((Get-MbtGuiProp $S 'Period' 'all') -ieq 'between')
    $from = [string](Get-MbtGuiProp $S 'ArrFrom' ''); $to = [string](Get-MbtGuiProp $S 'ArrTo' '')
    $problem = $null
    if ($chosen) {
        if (-not $from -and -not $to) { $problem = 'Pick at least one date, or choose All dates.' }
        elseif ($from -and $to -and [string]::CompareOrdinal($from, $to) -gt 0) { $problem = 'The first date is after the last date. Swap them, or pick other dates.' }
    }
    $text = if (-not $chosen -or $problem) { 'all dates' } elseif ($from -and $to) { "arrived between $from and $to" } elseif ($from) { "arrived from $from" } else { "arrived until $to" }
    [pscustomobject]@{ Chosen = $chosen; Active = ($chosen -and -not $problem); From = $(if ($chosen) { $from } else { '' }); To = $(if ($chosen) { $to } else { '' }); Problem = $problem; Text = $text }
}

function Get-MbtGuiPeriodStoreSuffix {
    <# A store keeps one period for a mailbox: an export with a period has its own store, named after it ('' for every date). #>
    param([hashtable]$S)
    $p = Get-MbtGuiPeriod $S
    if (-not $p.Active) { return '' }
    if ($p.From -and $p.To) { return "-arrived-$($p.From)_$($p.To)" }
    if ($p.From) { return "-arrived-from-$($p.From)" }
    return "-arrived-until-$($p.To)"
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
        [Parameter(Mandatory)][ValidateSet('export', 'import', 'repair', 'check', 'copy', 'mailcopy')][string]$Task,
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
            # the period: the days as ISO text (the same whatever the language of Windows); a day left empty is no limit on that side
            $period = Get-MbtGuiPeriod $S
            if ($period.Active -and $period.From) { $a.AddRange([string[]]@('-ArrivedFrom', $period.From)) }
            if ($period.Active -and $period.To) { $a.AddRange([string[]]@('-ArrivedTo', $period.To)) }
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
        'mailcopy' {
            # a mailbox into another in one job: the engine reads it into the store, then imports the store; one command, so that Pause and Continue work across the two parts
            $a = [Collections.Generic.List[string]]::new()
            $a.AddRange([string[]]@('-Mode', 'Copy', '-UserPrincipalName', $S.Mailbox))
            foreach ($file in @($Job.SelectionFiles)) { $a.AddRange([string[]]@('-SelectionFile', $file)) }
            $period = Get-MbtGuiPeriod $S
            if ($period.Active -and $period.From) { $a.AddRange([string[]]@('-ArrivedFrom', $period.From)) }
            if ($period.Active -and $period.To) { $a.AddRange([string[]]@('-ArrivedTo', $period.To)) }
            $a.AddRange([string[]]@('-TargetMailbox', $S.TgtLoc))
            # the application has no mailbox of its own: it names the one that receives; a delegated sign-in writes into its own
            if ($S.AuthMode -eq 'app' -and $S.McTarget) { $a.AddRange([string[]]@('-TargetMailboxId', $S.McTarget.Trim())) }
            if ($S.TgtMode -eq 'new' -and $S.TgtName) {
                $problem = Test-MbtGuiFolderName $S.TgtName
                if ($problem) { throw "The target folder name is refused: $problem" }
                $a.AddRange([string[]]@('-TargetRoot', $S.TgtName))
            }
            if ($S.Mapping) { $a.AddRange([string[]]@('-MappingFile', $S.Mapping)) }
            $a.AddRange([string[]]@('-StorePath', $Job.StorePath, '-TempPath', [IO.Path]::Combine($Job.Folder, 'temp'), '-NoPrompt', '-ReconcileDelaySeconds', [string]([int]$S.Reconcile * 60)))
            $a.AddRange([string[]]$tuning)
            $a.AddRange([string[]]$common)
            $out.Add([pscustomobject]@{ Title = 'Copy'; Arguments = $a.ToArray() })
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
$script:MbtGuiCliDefault = @{ Mailbox = 'Primary'; TargetMailbox = 'Primary'; Concurrency = '4'; MaxAttempts = '3'; ReconcileDelaySeconds = '300'; ReconcileRounds = '3'; PstMaxSizeGB = '50'; Level = 'Automatic'; RepairLevel = 'Rebuild' }

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
        for a repair also <new file>.repair-report.csv / .json next to the new file. Returns the paths written. -NoHtml leaves report.html out: the page of a job of 50,000 folders takes about a minute, so
        the console has the worker write it (Complete-MbtGuiRun).
    #>
    param([Parameter(Mandatory)]$Run, [Parameter(Mandatory)][string]$Folder, [string]$Task, [string]$RepairDestination, [switch]$NoHtml)
    $null = New-Item -ItemType Directory -Force -Path $Folder
    $rows = [Collections.Generic.List[object]]::new()
    foreach ($s in $Run.Steps) { $rows.Add([pscustomobject]@{ Kind = 'step'; Step = $s.Id; Key = $s.Title; State = $s.State; Done = $s.Done; Total = $s.Total; Message = $s.Msg }) }
    foreach ($o in $Run.Objects) { $rows.Add([pscustomobject]@{ Kind = 'folder'; Step = $o.Step; Key = $o.Key; State = $o.State; Done = $o.Done; Total = $o.Total; Message = $o.Msg }) }
    foreach ($i in $Run.Items) { $rows.Add([pscustomobject]@{ Kind = 'item'; Step = $i.Step; Key = $i.Key; State = $i.State; Done = ''; Total = $i.Total; Message = (($i.Msg, $i.Detail | Where-Object { $_ }) -join ' - ') }) }
    # an export with a period: the items it left out, said once in the CSV (skipped is not a failure)
    $periodCounts = if ($Run.Result) { $Run.Result.Counts } else { $null }
    $outside = Get-MbtGuiProp $periodCounts 'outsidePeriod' $null
    if ($null -ne $outside) {
        $text = '{0} items kept; {1} items outside the period were skipped (not an error)' -f (Get-MbtGuiProp $periodCounts 'kept' 0), $outside
        $undated = [long](Get-MbtGuiProp $periodCounts 'withoutDate' 0)
        if ($undated -gt 0) { $text += "; $undated items without a date were kept" }
        $rows.Add([pscustomobject]@{ Kind = 'period'; Step = 'list'; Key = 'items outside the period'; State = 'skipped'; Done = $outside; Total = ''; Message = $text })
    }
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
    if (-not $NoHtml -and (Test-Path -LiteralPath (Join-Path $Folder 'events.jsonl') -PathType Leaf)) {
        try { $written.Add((Write-MbtJobHtmlReport -Folder $Folder)) } catch { Write-Verbose "The HTML report could not be written: $($_.Exception.Message)" }
    }
    return $written.ToArray()
}

function Open-MbtGuiJobOverview {
    <#
        "All jobs": writes <work folder>\jobs\overview.html again from the status of every job (the worker does it, the window stays alive), then opens it in the browser. The
        demonstration writes nothing, and the dialog says so.
    #>
    $g = $script:Gui
    $say = { param([string]$Title, [string]$Body) Show-MbtGuiDialog -Title $Title -Body $Body -Buttons @(@{ Text = 'Close'; Accent = $true }) }
    $open = { param($path) if ($path -and (Test-Path -LiteralPath ([string]$path))) { Start-Process -FilePath ([string]$path) } else { & $say 'Nothing to open here' 'No file was written: this is the demonstration.' } }.GetNewClosure()
    $fail = { param($message) & $say 'The overview of all the jobs could not be written' ([string]$message) }.GetNewClosure()
    try { & $g.Engine.JobOverview $g.S.WorkFolder $open $fail }
    catch { & $fail $_.Exception.Message }
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

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCJ+M48ZqqfGzFz
# bjo7SHNte0DRzuZP4lX2fI6iOtPmdqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBznngP0j5n
# g68KzrHh2exsFFI8s9dfVTKku8Pc64ThfjANBgkqhkiG9w0BAQEFAASCAgBLIgOv
# BDo/F6/QLZjxLQVY4Qki4UzWQ4Q2aV8hSHwfzCeVH452VqIzHDf6e3n58WE+oGiB
# w1UhF6q4rY7NADGks4dCRmS9eJXIUVg7+F+LtGyODFCey7iqo0oJxiOfAvrLccQ7
# JXANYTeSebN53vBKoJWmkmf2Czc+o7tZjAvzzVMBtftE7ut/hKVxYRIKUidU3/EC
# lAjANxIusGXvSfUwF/QpKRA++ejwZ6WJGUzLfuBVeULr9HpQZ3pemlRWt/Y/Qdda
# yQbt+WpXjc2DFrApSf5x34ULMpLXZl4YQOgYdLXHY5eCje0DP3JYhX/0RCFgqVcA
# 2KE8lYspkKyWUV9DpZVCn7YJurtTeAMREqwnPOVvVNPMrTM8EjV6Jl1ZQtr986xz
# 2K+p70ig2qMFSnX0LK5FLj0sA1wUGS6QnKwU3UEIFXt13DjotE5un0HGYAuj9YZ7
# ZKZLvTSBHYrmBhjsIHFTJb8BSZeeMWpu3/zh3CdkdNB2mK1JaUpaUFH0n9Wk6btk
# DwC9oGw/EEG7J2f1Pwj/R97ayXMG81CAaC30Y/8Fy46+alwhfI8cpABKvaQeTMEO
# HOjB6cCSm+icgjCz+oDv3vMHUShQF16weP1GZvt5Ke5fZod69KryTc9hCGVsaFWs
# 9bPcYeZpLibqARaqrB0H7XUg5add4AUA4c1GTKGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEICILgoG/4Y9Xouy+SFIFgzT1H1CmApGX6oCWV5sOVkxqAggOhAjvoQCH
# gRgPMjAyNjExMTUxNDM1MTZaMAMCAQECEFU1E2TDP7b0mQ5f8y6R4KqgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTE0MzUxNlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEINcJ
# wdWQfpE8Fw2S5kRpOtPZI5XQieuQY3oq9yP3BvKNMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEA80Smuu1zCDgOrHpX0dlheceFP/6Jg3KvjamVI0vJugICIE0t
# /tcENh8L870osGwDkpz93etxWTGbNlYQhT6/bnFU
# SIG # End signature block
