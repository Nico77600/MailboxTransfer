#Requires -Version 7.6
<#
.SYNOPSIS
    Mailbox Transfer - the export or the import of a list of mailboxes, a batch (dot-sourced by MailboxTransfer.psm1).

.DESCRIPTION
    A batch exports the mailboxes of a CSV file, imports into them, or copies one mailbox into another, one ordinary job for each mailbox (or for each import, or each copy): its own job folder (events.jsonl, status.json with the batch name, report.html,
    request.json for the Continue of the console), run by a child process of Invoke-MailboxTransfer.ps1, so that the engine, the sign-in cache and the throttling of one mailbox never meet those of another.
    The batch itself is a folder, <work folder>\batches\<batch>\, with batch.json (the rows of the CSV, the options, the state of each row and its job), batch.log and report.html (the overview of
    the jobs of the batch). The same command again continues it: complete mailboxes are skipped, the others do only what is left.

        Read-MbtBatchCsv         reads and checks the whole CSV of an export before anything starts; every problem is returned, not the first only
        Read-MbtImportBatchCsv   the same for the CSV of an import (-Mode Import, ImportPst, ImportEml)
        Read-MbtCopyBatchCsv     the same for the CSV of a copy (-Mode Copy): one line for each source mailbox location and target mailbox location
        Read-MbtBatch            batch.json of a batch folder
        Invoke-MbtBatch          runs the rows, 1 to 4 mailboxes at a time; returns the counts and the exit code
        Write-MbtBatchReport     <batch folder>\report.html
        Get-MbtBatchList         the batches of a work folder with their counts (the console lists them)
        Get-MbtBatchProtection   the jobs and batches that the retention of the job folders must keep

    An import batch runs the rows that write into different mailboxes side by side and never two rows that write into the same mailbox location, or read the same source, at the same time. A batch of copies
    does the same for the target mailbox location and the source mailbox location of its lines.

.NOTES
    Author  : Nicolas Fabert
#>

$script:MbtBatchLocations = @('Primary', 'Archive', 'Both')
$script:MbtBatchRowStates = @('pending', 'running', 'complete', 'partial', 'failed')
$script:MbtBatchImportModes = @('Import', 'ImportPst', 'ImportEml')
$script:MbtBatchImportColumns = @('Mailbox', 'Source', 'Location', 'TargetRoot', 'SourceMailbox', 'Notes')
$script:MbtBatchCopyColumns = @('SourceAddress', 'SourceLocation', 'TargetAddress', 'TargetLocation', 'TargetLayout', 'TargetRoot', 'TargetArchiveRoot', 'SelectionFile', 'Notes')

#region Names and folders -----------------------------------------------------------------------

function Test-MbtBatchName {
    <# The reason a batch name cannot be used (it is the name of a folder and of the batch in the reports), or $null when it is fine. #>
    param([AllowEmptyString()][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'The batch name is empty.' }
    if ($Name.Length -gt 64) { return 'The batch name is too long (64 characters at most).' }
    if ($Name -notmatch '^[A-Za-z0-9][A-Za-z0-9 ._-]*$' -or $Name -match '[ .]$') { return 'The batch name can hold letters, digits, spaces, dots, dashes and underscores, starts with a letter or a digit, and does not end with a space or a dot.' }
    return $null
}

function Get-MbtBatchFolder {
    <# <work folder>\batches\<name>. The name is a folder name, never a path. #>
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)][string]$Name)
    $problem = Test-MbtBatchName $Name
    if ($problem) { throw $problem }
    return [IO.Path]::Combine($WorkFolder, 'batches', $Name)
}

function Find-MbtBatchFolder {
    <# The folder of an existing batch, found by its name whatever the case, or $null. #>
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)][string]$Name)
    $root = [IO.Path]::Combine($WorkFolder, 'batches')
    if (-not [IO.Directory]::Exists($root)) { return $null }
    foreach ($d in [IO.Directory]::EnumerateDirectories($root)) { if ([IO.Path]::GetFileName($d) -ieq $Name.Trim()) { return $d } }
    return $null
}

function ConvertTo-MbtBatchSafeName {
    <# A text as one file or folder name: characters Windows refuses become an underscore; the dot, the dash and the at sign stay, so an address stays readable. #>
    param([Parameter(Mandatory)][string]$Text)
    $safe = [regex]::Replace($Text.Trim(), '[\\/:*?"<>|\x00-\x1f\s]+', '_').TrimEnd('.', ' ')
    if (-not $safe) { return '_' }
    return $safe
}

function Test-MbtBatchFileName {
    <# The reason a file name (one name, no folder) cannot be created or opened by Outlook, or $null. #>
    param([AllowEmptyString()][string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return 'the name is empty' }
    if ($Name.IndexOfAny([IO.Path]::GetInvalidFileNameChars()) -ge 0) { return 'the name has a character Windows does not accept' }
    if ($Name.EndsWith('.') -or $Name.EndsWith(' ')) { return 'the name ends with a dot or a space' }
    if ($Name.Split('.')[0].TrimEnd(' ') -match '^(?i)(CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])$') { return 'the name is one that Windows keeps for itself' }
    return $null
}

#endregion

#region The CSV ---------------------------------------------------------------------------------

function ConvertFrom-MbtBatchCsvText {
    <# CSV text as records (Line = the line where the record starts, Fields). The delimiter is the one, comma or semicolon, that the first line holds most of; a field may be quoted. Blank lines are left out. #>
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    $Text = $Text.TrimStart([char]0xFEFF)
    $commas = 0; $semicolons = 0; $quoted = $false
    foreach ($c in $Text.ToCharArray()) {
        if ($c -eq '"') { $quoted = -not $quoted }
        elseif (-not $quoted) {
            if ($c -eq ',') { $commas++ } elseif ($c -eq ';') { $semicolons++ }
            elseif ($c -eq "`n" -and ($commas + $semicolons) -gt 0) { break }
        }
    }
    $delimiter = if ($semicolons -gt $commas) { ';' } else { ',' }
    $records = [Collections.Generic.List[object]]::new()
    $fields = [Collections.Generic.List[string]]::new()
    $field = [Text.StringBuilder]::new()
    $line = 1; $startLine = 1; $inQuote = $false;
    $end = {
        $fields.Add($field.ToString()); [void]$field.Clear()
        if (@($fields | Where-Object { $_.Trim() }).Count) { $records.Add([pscustomobject]@{ Line = $startLine; Fields = $fields.ToArray() }) }
        $fields.Clear()
    }
    for ($i = 0; $i -lt $Text.Length; $i++) {
        $c = $Text[$i]
        if ($inQuote) {
            if ($c -eq '"') { if (($i + 1) -lt $Text.Length -and $Text[$i + 1] -eq '"') { [void]$field.Append('"'); $i++ } else { $inQuote = $false } }
            else { if ($c -eq "`n") { $line++ }; [void]$field.Append($c) }
        }
        elseif ($c -eq '"' -and $field.ToString().Trim().Length -eq 0) { $inQuote = $true; [void]$field.Clear() }
        elseif ($c -eq $delimiter) { $fields.Add($field.ToString()); [void]$field.Clear() }
        elseif ($c -eq "`r") { }
        elseif ($c -eq "`n") { & $end; $line++; $startLine = $line }
        else { [void]$field.Append($c) }
    }
    if ($field.Length -gt 0 -or $fields.Count -gt 0) { & $end }
    [pscustomobject]@{ Delimiter = $delimiter; Records = $records.ToArray(); Unclosed = $inQuote }
}

function Resolve-MbtBatchOutput {
    <#
        The PST file a mailbox is written to: the OutputPst of its row (a .pst file, or a folder: the file is named after the address), else the global -OutputPst (a folder: one file named after the
        address; a .pst file: its name with the address added), else nothing (the export store only). A relative path of the row is relative to the folder of the CSV file. -Suffix is added to the name
        taken from the address (an address that is on several lines).
    #>
    param([Parameter(Mandatory)][string]$UserPrincipalName, [string]$RowOutput, [string]$GlobalOutput, [string]$CsvFolder, [string]$Suffix)
    $safe = ConvertTo-MbtBatchSafeName $UserPrincipalName
    if ($Suffix) { $safe += '-' + (ConvertTo-MbtBatchSafeName $Suffix) }
    $isFile = { param($p) [IO.Path]::GetExtension($p.TrimEnd('\', '/')) -ieq '.pst' }
    if ($RowOutput) {
        $p = if ([IO.Path]::IsPathRooted($RowOutput)) { $RowOutput } else { [IO.Path]::Combine($CsvFolder, $RowOutput) }
        if (& $isFile $p) { return [IO.Path]::GetFullPath($p) }
        return [IO.Path]::GetFullPath([IO.Path]::Combine($p, "$safe.pst"))
    }
    if ($GlobalOutput) {
        $p = if ([IO.Path]::IsPathRooted($GlobalOutput)) { $GlobalOutput } else { [IO.Path]::Combine((Get-Location).Path, $GlobalOutput) }
        if (& $isFile $p) {
            $dir = [IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($p))
            return [IO.Path]::Combine($dir, ('{0}-{1}.pst' -f [IO.Path]::GetFileNameWithoutExtension($p), $safe))
        }
        return [IO.Path]::GetFullPath([IO.Path]::Combine($p, "$safe.pst"))
    }
    return ''
}

function Read-MbtBatchCsv {
    <#
    .SYNOPSIS
        Reads the CSV file of a batch and checks all of it: the problems of every row are returned together, so that one correction of the file is enough.
    .DESCRIPTION
        One line for each mailbox and Location: the same address may be on several lines for its Primary mailbox, its Archive or different SelectionFiles (the store and the PST of those lines get the Location, or
        the number of the line, added to their name), never twice for the same Location (Both covers Primary and Archive) or the same SelectionFile.
        Columns (the names do not depend on upper or lower case; the delimiter is a comma or a semicolon; UTF-8 with or without BOM; other columns are ignored):
        UserPrincipalName (required), Location (Primary, Archive or Both; empty: -DefaultLocation), OutputPst (optional: a .pst file or a folder), SelectionFile (optional: a selection written by the
        browse view or the console; it names the mailbox and the folders, the Location is then not used), Notes (ignored).
        Returns Rows (UserPrincipalName, Location, SelectionFile, Output, Store, Line), Problems (text), Delimiter. Checked: a missing or wrong address, the same address twice (upper and lower case are the
        same), a wrong Location, a SelectionFile that does not exist, a file name Windows cannot take, two rows that write the same PST file (or a part of it) or use the same store folder, and, with a delegated
        sign-in, an archive asked of a mailbox that is not the signed-in account's.
    .PARAMETER OutputPst
        The global -OutputPst: a folder, or a .pst file whose name gets the address added.
    .PARAMETER StoreBase
        The folder that holds one export store for each mailbox (<StoreBase>\<address>).
    .PARAMETER SignedInUser
        With a delegated sign-in: the account that signs in. Only its own mailbox has an archive that can be exported; other mailboxes only their primary mailbox.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('Primary', 'Archive', 'Both', 'All')][string]$DefaultLocation = 'Primary',
        [string]$OutputPst, [string]$StoreBase, [string]$SignedInUser, [switch]$AppOnly
    )
    Set-StrictMode -Version Latest
    $problems = [Collections.Generic.List[string]]::new()
    $rows = [Collections.Generic.List[object]]::new()
    $result = { [pscustomobject]@{ Rows = $rows.ToArray(); Problems = $problems.ToArray(); Delimiter = $delimiter } }
    $delimiter = ','
    if (-not [IO.File]::Exists($Path)) { $problems.Add("The CSV file was not found: $Path"); return (& $result) }
    $text = $null
    try {
        $reader = [IO.StreamReader]::new($Path, [Text.UTF8Encoding]::new($false), $true)
        try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } catch { $problems.Add("The CSV file cannot be read ($Path): $($_.Exception.Message)"); return (& $result) }
    $csv = ConvertFrom-MbtBatchCsvText $text
    $delimiter = $csv.Delimiter
    $csvFolder = Split-Path ([IO.Path]::GetFullPath($Path)) -Parent
    if ($csv.Unclosed) { $problems.Add('A quote is not closed: the end of the file is inside a quoted text.') }
    if (-not $csv.Records.Count) { $problems.Add('The CSV file is empty: the first line names the columns, then one mailbox on each line.'); return (& $result) }

    $header = $csv.Records[0]
    $column = @{}
    for ($i = 0; $i -lt $header.Fields.Count; $i++) {
        $name = $header.Fields[$i].Trim()
        if (-not $name) { continue }
        if ($column.ContainsKey($name.ToLowerInvariant())) { $problems.Add("The column $name is there twice (the first line): keep one."); continue }
        $column[$name.ToLowerInvariant()] = $i
    }
    if (-not $column.ContainsKey('userprincipalname')) {
        $problems.Add('The first line must name the columns, and UserPrincipalName is required (found: ' + $(if ($column.Count) { (($header.Fields | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ', ') } else { 'nothing' }) + ').')
        return (& $result)
    }
    $cell = { param($record, [string]$name) $at = $column[$name]; if ($null -ne $at -and $at -lt $record.Fields.Count) { $record.Fields[$at].Trim() } else { '' } }
    $default = if ($DefaultLocation -eq 'All') { 'Both' } else { $DefaultLocation }
    $seen = @{}
    $files = @{}
    $stores = @{}
    $number = 0
    foreach ($record in @($csv.Records | Select-Object -Skip 1)) {
        $number++
        $line = $record.Line
        $upn = & $cell $record 'userprincipalname'
        if (-not $upn) { $problems.Add("line ${line}: UserPrincipalName is empty."); continue }
        if ($upn -notmatch '^[^@\s,;"<>]+@[^@\s,;"<>]+$') { $problems.Add("line ${line}: '$upn' is not a mailbox address (user@domain)."); continue }
        $key = $upn.ToLowerInvariant()
        $given = & $cell $record 'location'
        $location = $default
        if ($given) {
            $match = @($script:MbtBatchLocations + 'All' | Where-Object { $_ -ieq $given })
            if (-not $match.Count) { $problems.Add("line ${line}: Location '$given' is not Primary, Archive or Both."); continue }
            $location = if ($match[0] -ieq 'All') { 'Both' } else { $match[0] }
        }
        $selection = & $cell $record 'selectionfile'
        if ($selection) {
            if (-not [IO.Path]::IsPathRooted($selection)) { $selection = [IO.Path]::Combine($csvFolder, $selection) }
            if (-not [IO.File]::Exists($selection)) { $problems.Add("line ${line}: the SelectionFile was not found: $selection"); continue }
            $selection = [IO.Path]::GetFullPath($selection)
        }
        # the same mailbox again: only for another Location (Both covers the two) or another selection
        $clash = $null
        foreach ($earlier in $(if ($seen.ContainsKey($key)) { $seen[$key] } else { @() })) {
            if ($selection -or $earlier.Selection) { if ($selection -and $earlier.Selection -ieq $selection) { $clash = $earlier } }
            elseif ($earlier.Location -eq $location -or $earlier.Location -eq 'Both' -or $location -eq 'Both') { $clash = $earlier }
        }
        if ($clash) { $problems.Add("line ${line}: $upn is already on line $($clash.Line) (the same mailbox for the same Location, or the same SelectionFile; upper and lower case are the same; Both covers Primary and Archive)."); continue }
        if (-not $seen.ContainsKey($key)) { $seen[$key] = [Collections.Generic.List[object]]::new() }
        $seen[$key].Add([pscustomobject]@{ Line = $line; Location = $location; Selection = $selection })
        if (-not $selection -and $SignedInUser -and -not $AppOnly -and $location -ne 'Primary' -and $upn -ine $SignedInUser) {
            $problems.Add("line ${line}: $upn is not the signed-in account ($SignedInUser). With a delegated sign-in only its primary mailbox can be exported: use Location Primary, or sign in as the application (-AppOnly).")
            continue
        }
        $rows.Add([pscustomobject]@{ Number = $number; Line = $line; UserPrincipalName = $upn; Location = $location; SelectionFile = $selection; Output = ''; Store = ''; RowOutput = (& $cell $record 'outputpst') })
    }
    if (-not $rows.Count -and -not $problems.Count) { $problems.Add('The CSV file has no mailbox: the first line names the columns, then one mailbox on each line.') }

    # the PST file and the store folder of each row: named after the address; an address that is on several lines gets the Location of the line, or its number, added
    $perAddress = @{}
    foreach ($r in $rows) { $k = $r.UserPrincipalName.ToLowerInvariant(); if (-not $perAddress.ContainsKey($k)) { $perAddress[$k] = [Collections.Generic.List[object]]::new() }; $perAddress[$k].Add($r) }
    $valid = [Collections.Generic.List[object]]::new()
    foreach ($r in $rows) {
        $group = @($perAddress[$r.UserPrincipalName.ToLowerInvariant()])
        $suffix = ''
        if ($group.Count -gt 1) { $suffix = if (-not $r.SelectionFile -and @($group | Where-Object { -not $_.SelectionFile -and $_.Location -eq $r.Location }).Count -eq 1) { $r.Location.ToLowerInvariant() } else { "row$($r.Number)" } }
        if ($r.RowOutput -or $OutputPst) {
            # the name as written in the CSV: resolving a path ending in a device name such as aux.pst gives the device itself on Windows 10 and Server 2022
            $rowFile = ([string]$r.RowOutput).TrimEnd('\', '/')
            if ($rowFile -and [IO.Path]::GetExtension($rowFile) -ieq '.pst') {
                $leaf = [IO.Path]::GetFileName($rowFile)
                $why = Test-MbtBatchFileName $leaf
                if ($why) { $problems.Add("line $($r.Line): the PST file name $leaf cannot be used: $why."); continue }
            }
            try { $r.Output = Resolve-MbtBatchOutput -UserPrincipalName $r.UserPrincipalName -RowOutput $r.RowOutput -GlobalOutput $OutputPst -CsvFolder $csvFolder -Suffix $suffix } catch { $problems.Add("line $($r.Line): the OutputPst is not a path ($($r.RowOutput)): $($_.Exception.Message)"); continue }
            $why = Test-MbtBatchFileName ([IO.Path]::GetFileName($r.Output))
            if ($why) { $problems.Add("line $($r.Line): the PST file name $([IO.Path]::GetFileName($r.Output)) cannot be used: $why."); continue }
        }
        if ($StoreBase) { $r.Store = [IO.Path]::GetFullPath([IO.Path]::Combine($StoreBase, (ConvertTo-MbtBatchSafeName ($r.UserPrincipalName + $(if ($suffix) { '-' + $suffix } else { '' }))))) }
        $valid.Add($r)
    }
    $rows = $valid

    # two rows must never write the same file: the file, and the parts of a split PST (name-002.pst, name-003.pst...) count
    $stems = @{}
    foreach ($r in $rows) {
        if ($r.Output) {
            $stem = $r.Output.Substring(0, $r.Output.Length - 4).ToLowerInvariant()
            if ($stems.ContainsKey($stem)) { $problems.Add("line $($r.Line): $($r.Output) is also the file of line $($stems[$stem]): two mailboxes would write the same file."); continue }
            $stems[$stem] = $r.Line
        }
    }
    foreach ($r in $rows) {
        if (-not $r.Output) { continue }
        $stem = $r.Output.Substring(0, $r.Output.Length - 4).ToLowerInvariant()
        $m = [regex]::Match($stem, '^(.*)-\d{3}$')
        if ($m.Success -and $stems.ContainsKey($m.Groups[1].Value) -and $stems[$m.Groups[1].Value] -ne $r.Line) {
            $problems.Add("line $($r.Line): $($r.Output) has the name of a part of the PST file of line $($stems[$m.Groups[1].Value]) (name-002.pst, name-003.pst...).")
        }
    }
    foreach ($r in $rows) {
        if (-not $r.Store) { continue }
        $k = $r.Store.ToLowerInvariant()
        if ($stores.ContainsKey($k)) { $problems.Add("line $($r.Line): the store folder $($r.Store) is also the one of line $($stores[$k]) (the two addresses give the same folder name)."); continue }
        $stores[$k] = $r.Line
    }
    return (& $result)
}

#endregion

#region The CSV of an import ---------------------------------------------------------------------

function Get-MbtBatchMode {
    <# Export, Import, ImportPst or ImportEml: what the batch does (a batch written before import existed says nothing: it is an export). #>
    param($Batch)
    if ($Batch -is [System.Collections.IDictionary] -and $Batch.Contains('mode') -and $Batch['mode']) { return [string]$Batch['mode'] }
    return 'Export'
}

function ConvertTo-MbtBatchTargetRoot {
    <# A folder path of a mailbox as the import takes it: slashes, no slash at either end. #>
    param([AllowEmptyString()][string]$Text)
    return (([string]$Text).Trim() -replace '\\', '/').Trim('/', ' ')
}

function Get-MbtImportBatchKey {
    <#
        Twelve hexadecimal digits that are the same for the same target mailbox, Location, Source, folder and source mailbox, whatever the case: the journal of a PST import and the job name of a store
        import come from it, so the same row never imports an item twice, whichever batch (or single command) it is in.
    #>
    param([string]$Mailbox, [string]$Location, [string]$Source, [string]$TargetRoot, [string]$SourceMailbox)
    $text = (@($Mailbox, $Location, $Source, $TargetRoot, $SourceMailbox) | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() }) -join '|'
    $hash = [Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes($text))
    return ([BitConverter]::ToString($hash).Replace('-', '').Substring(0, 12).ToLowerInvariant())
}

function Test-MbtImportBatchSource {
    <# The reason a Source cannot be imported by the mode (not found, the wrong kind of path, nothing to import in it), or $null. #>
    param([Parameter(Mandatory)][ValidateSet('Import', 'ImportPst', 'ImportEml')][string]$Mode, [Parameter(Mandatory)][string]$Path)
    $isFolder = [IO.Directory]::Exists($Path)
    $isFile = [IO.File]::Exists($Path)
    if (-not $isFolder -and -not $isFile) { return "the Source was not found: $Path" }
    try {
        switch ($Mode) {
            'Import' {
                if (-not $isFolder) { return "the Source is a file: an export store is a folder (it holds catalog.db): $Path" }
                if (-not [IO.File]::Exists([IO.Path]::Combine($Path, 'catalog.db'))) { return "the Source is not an export store (no catalog.db in it): $Path" }
            }
            'ImportPst' {
                if ($isFile -and [IO.Path]::GetExtension($Path) -ine '.pst') { return "the Source is not a .pst file: $Path" }
                if ($isFolder -and -not [IO.Directory]::EnumerateFiles($Path, '*.pst', [IO.SearchOption]::TopDirectoryOnly).GetEnumerator().MoveNext()) { return "the Source folder holds no .pst file: $Path" }
            }
            'ImportEml' {
                if ($isFile -and [IO.Path]::GetExtension($Path) -ine '.eml') { return "the Source is not an .eml file or a folder of .eml files: $Path" }
                if ($isFolder -and -not [IO.Directory]::EnumerateFiles($Path, '*.eml', [IO.SearchOption]::AllDirectories).GetEnumerator().MoveNext()) { return "the Source folder holds no .eml file: $Path" }
            }
        }
    } catch { return "the Source cannot be read ($Path): $($_.Exception.Message)" }
    return $null
}

function Read-MbtImportBatchCsv {
    <#
    .SYNOPSIS
        Reads the CSV file of an import batch and checks all of it: the problems of every row are returned together, so that one correction of the file is enough.
    .DESCRIPTION
        One line for each import: the mailbox it goes into, where it comes from and where it goes. Columns (the names do not depend on upper or lower case; the delimiter is a comma or a semicolon; UTF-8 with
        or without BOM): Mailbox (the target address, required), Source (required: the export store for -Mode Import, a .pst file or a folder of .pst files for -Mode ImportPst, a folder of .eml files or one .eml
        file for -Mode ImportEml; relative to the folder of the CSV file), Location (Primary or Archive; empty: -DefaultLocation), TargetRoot (the folder everything is created below; empty: the global one; an
        export store may go to the same folder names, a PST needs one or -MergeWellKnown, EML files need one), SourceMailbox (-Mode Import only: Primary or Archive of a store that holds both; empty: the
        global one), Notes (ignored). Any other column is a problem.
        Returns Rows (Number, Line, Mailbox, Location, Source, TargetRoot, SourceMailbox, Key, Journal), Problems (text), Delimiter. Checked: a missing or wrong address, a wrong Location, a Source that is
        missing, not found or of the wrong kind, the same target, Location and Source twice (and the same SourceMailbox), a folder or a source mailbox that does not fit the mode, and, with a delegated sign-in,
        a target that is not the signed-in account (the service refuses a delegated write into another mailbox).
    .PARAMETER WorkFolder
        The work folder: the journal of a PST import is <work folder>\imports\<key>, outside the job folders, so that the retention of the job folders never removes what stops an item from being imported twice.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][ValidateSet('Import', 'ImportPst', 'ImportEml')][string]$Mode,
        [ValidateSet('Primary', 'Archive')][string]$DefaultLocation = 'Primary',
        [string]$DefaultTargetRoot, [string]$DefaultSourceMailbox, [switch]$MergeWellKnown,
        [string]$SignedInUser, [switch]$AppOnly, [string]$WorkFolder
    )
    Set-StrictMode -Version Latest
    $problems = [Collections.Generic.List[string]]::new()
    $rows = [Collections.Generic.List[object]]::new()
    $delimiter = ','
    $result = { [pscustomobject]@{ Rows = $rows.ToArray(); Problems = $problems.ToArray(); Delimiter = $delimiter } }
    if (-not [IO.File]::Exists($Path)) { $problems.Add("The CSV file was not found: $Path"); return (& $result) }
    $text = $null
    try {
        $reader = [IO.StreamReader]::new($Path, [Text.UTF8Encoding]::new($false), $true)
        try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } catch { $problems.Add("The CSV file cannot be read ($Path): $($_.Exception.Message)"); return (& $result) }
    $csv = ConvertFrom-MbtBatchCsvText $text
    $delimiter = $csv.Delimiter
    $csvFolder = Split-Path ([IO.Path]::GetFullPath($Path)) -Parent
    if ($csv.Unclosed) { $problems.Add('A quote is not closed: the end of the file is inside a quoted text.') }
    if (-not $csv.Records.Count) { $problems.Add('The CSV file is empty: the first line names the columns, then one import on each line.'); return (& $result) }

    $known = @{}
    foreach ($c in $script:MbtBatchImportColumns) { $known[$c.ToLowerInvariant()] = $c }
    $header = $csv.Records[0]
    $column = @{}
    for ($i = 0; $i -lt $header.Fields.Count; $i++) {
        $name = $header.Fields[$i].Trim()
        if (-not $name) { continue }
        $k = $name.ToLowerInvariant()
        if ($column.ContainsKey($k)) { $problems.Add("The column $name is there twice (the first line): keep one."); continue }
        $column[$k] = $i
        if (-not $known.ContainsKey($k)) { $problems.Add("The column $name is not known (the columns are $($script:MbtBatchImportColumns -join ', ')): check the spelling, or remove it.") }
    }
    if (-not $column.ContainsKey('mailbox') -or -not $column.ContainsKey('source')) {
        $problems.Add('The first line must name the columns, and Mailbox and Source are required (found: ' + $(if ($column.Count) { (($header.Fields | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ', ') } else { 'nothing' }) + ').')
        return (& $result)
    }
    $cell = { param($record, [string]$name) $at = $column[$name]; if ($null -ne $at -and $at -lt $record.Fields.Count) { $record.Fields[$at].Trim() } else { '' } }
    $seen = @{}
    $number = 0
    foreach ($record in @($csv.Records | Select-Object -Skip 1)) {
        $number++
        $line = $record.Line
        $before = $problems.Count

        $mailbox = & $cell $record 'mailbox'
        if (-not $mailbox) { $problems.Add("line ${line}: Mailbox is empty.") }
        elseif ($mailbox -notmatch '^[^@\s,;"<>]+@[^@\s,;"<>]+$') { $problems.Add("line ${line}: '$mailbox' is not a mailbox address (user@domain)."); $mailbox = '' }
        elseif ($SignedInUser -and -not $AppOnly -and $mailbox -ine $SignedInUser) {
            $problems.Add("line ${line}: $mailbox is not the signed-in account ($SignedInUser). With a delegated sign-in only the mailbox of the signed-in account can be written to (the service refuses a delegated write into another person's or a shared mailbox): sign in as the application (-AppOnly), or list only that mailbox.")
        }

        $location = $DefaultLocation
        $given = & $cell $record 'location'
        if ($given) {
            $match = @('Primary', 'Archive' | Where-Object { $_ -ieq $given })
            if ($match.Count) { $location = $match[0] } else { $problems.Add("line ${line}: Location '$given' is not Primary or Archive (an import goes into one mailbox)."); $location = '' }
        }

        $source = & $cell $record 'source'
        $sourcePath = ''
        if (-not $source) { $problems.Add("line ${line}: Source is empty.") }
        else {
            $full = $null
            try { $full = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($(if ([IO.Path]::IsPathRooted($source)) { $source } else { [IO.Path]::Combine($csvFolder, $source) }))) }
            catch { $problems.Add("line ${line}: the Source is not a path ($source): $($_.Exception.Message)") }
            if ($full) {
                $why = Test-MbtImportBatchSource -Mode $Mode -Path $full
                if ($why) { $problems.Add("line ${line}: $why") } else { $sourcePath = $full }
            }
        }

        $rootText = & $cell $record 'targetroot'
        if (-not $rootText) { $rootText = $DefaultTargetRoot }
        $root = ConvertTo-MbtBatchTargetRoot $rootText
        if ($rootText -and -not $root) { $problems.Add("line ${line}: TargetRoot '$rootText' is not a folder path.") }
        elseif ($root -and @($root -split '/' | Where-Object { -not $_.Trim() }).Count) { $problems.Add("line ${line}: TargetRoot '$rootText' has an empty folder name in it.") }
        elseif (-not $root -and $Mode -eq 'ImportEml') { $problems.Add("line ${line}: no TargetRoot: give the column, or -TargetRoot (the .eml files are created in folders below one top-level folder).") }
        elseif (-not $root -and $Mode -eq 'ImportPst' -and -not $MergeWellKnown) { $problems.Add("line ${line}: no TargetRoot: give the column, or -TargetRoot, or -MergeWellKnown (the folders of the PST are then merged into the standard folders of the mailbox).") }

        $sourceMailbox = & $cell $record 'sourcemailbox'
        if ($Mode -ne 'Import') {
            if ($sourceMailbox) { $problems.Add("line ${line}: SourceMailbox applies to an export store (-Mode Import) only: leave it empty.") }
            $sourceMailbox = ''
        } else {
            if (-not $sourceMailbox) { $sourceMailbox = $DefaultSourceMailbox }
            if ($sourceMailbox) {
                $match = @('Primary', 'Archive' | Where-Object { $_ -ieq $sourceMailbox })
                if ($match.Count) { $sourceMailbox = $match[0] } else { $problems.Add("line ${line}: SourceMailbox '$sourceMailbox' is not Primary or Archive."); $sourceMailbox = '' }
            }
        }
        if ($problems.Count -gt $before) { continue }

        $twice = (@($mailbox, $location, $sourcePath, $sourceMailbox) | ForEach-Object { $_.ToLowerInvariant() }) -join '|'
        if ($seen.ContainsKey($twice)) { $problems.Add("line ${line}: $mailbox ($location) already imports $sourcePath on line $($seen[$twice]) (the same target, Location and Source; upper and lower case are the same)."); continue }
        $seen[$twice] = $line
        $key = Get-MbtImportBatchKey -Mailbox $mailbox -Location $location -Source $sourcePath -TargetRoot $root -SourceMailbox $sourceMailbox
        $rows.Add([pscustomobject]@{
                Number = $number; Line = $line; Mailbox = $mailbox; Location = $location; Source = $sourcePath; TargetRoot = $root; SourceMailbox = $sourceMailbox; Key = $key
                Journal = $(if ($WorkFolder) { [IO.Path]::Combine($WorkFolder, 'imports', $key) } else { '' })
            })
    }
    if (-not $rows.Count -and -not $problems.Count) { $problems.Add('The CSV file has no import: the first line names the columns, then one import on each line.') }
    return (& $result)
}

#endregion

#region The CSV of a copy ------------------------------------------------------------------------

function Read-MbtCopyBatchCsv {
    <#
    .SYNOPSIS
        Reads the CSV file of a batch of copies and checks all of it: the problems of every row are returned together, so that one correction of the file is enough.
    .DESCRIPTION
        One line for each copy: the mailbox location it reads and the mailbox location it writes into. Columns (the names do not depend on upper or lower case; the delimiter is a comma or a semicolon; UTF-8 with
        or without BOM): SourceAddress (the mailbox read, required), SourceLocation (Primary, Archive or All, the Primary and the Archive together; empty: -DefaultSourceLocation), TargetAddress (the mailbox written into,
        required), TargetLocation (Primary or Archive; empty: -DefaultTargetLocation; empty for a line that copies both places), TargetLayout (Mirror, Primary or Archive, for a line that copies both places: where the two
        parts go; empty: -DefaultTargetLayout, else Mirror), TargetRoot (the folder everything is created below; empty: the global one; empty everywhere: the folder names of the source), TargetArchiveRoot (the folder the
        Archive part of a line that copies both places is created below; empty: -DefaultTargetArchiveRoot, else TargetRoot), SelectionFile (optional: what to copy, written by the browse view or the console, of one mailbox
        location or of the Primary and the Archive; relative to the folder of the CSV file), Notes (ignored). Any other column is a problem. A file of an earlier version, without the new columns, is read as before.
        Returns Rows (Number, Line, SourceAddress, SourceLocation, TargetAddress, TargetLocation, TargetLayout, TargetRoot, TargetArchiveRoot, SelectionFile, Store), Problems (text), Delimiter. Checked: a missing or wrong address, a wrong Location,
        a SelectionFile that is missing or of several mailbox locations, a source that is the target (the same address and Location), the same copy twice (the same source, target, Locations and selection), and, with a
        delegated sign-in, a target that is not the signed-in account (the service refuses a delegated write into another mailbox) and an archive read from a mailbox that is not the signed-in account's. Lines that
        write into the same target location, or read the same source location, are fine: they run one after the other.
    .PARAMETER StoreBase
        The folder that holds one export store for each source location and choice (<StoreBase>\copy-<address>-<location>-<8 digits of the choice>): a source copied into two targets is read once.
    .PARAMETER Choice
        What every line of the batch reads besides its selection (the options of the command): Folder, Period (a period of the engine), IncludeRecoverableItems, RecoverableItemsFolders. They make the store of a line.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [ValidateSet('Primary', 'Archive', 'All')][string]$DefaultSourceLocation = 'Primary',
        [ValidateSet('Primary', 'Archive')][string]$DefaultTargetLocation = 'Primary',
        [ValidateSet('', 'Mirror', 'Primary', 'Archive')][string]$DefaultTargetLayout = '',
        [string]$DefaultTargetRoot, [string]$DefaultTargetArchiveRoot, [string]$StoreBase, [hashtable]$Choice = @{}, [string]$SignedInUser, [switch]$AppOnly
    )
    Set-StrictMode -Version Latest
    $problems = [Collections.Generic.List[string]]::new()
    $rows = [Collections.Generic.List[object]]::new()
    $delimiter = ','
    $result = { [pscustomobject]@{ Rows = $rows.ToArray(); Problems = $problems.ToArray(); Delimiter = $delimiter } }
    if (-not [IO.File]::Exists($Path)) { $problems.Add("The CSV file was not found: $Path"); return (& $result) }
    $text = $null
    try {
        $reader = [IO.StreamReader]::new($Path, [Text.UTF8Encoding]::new($false), $true)
        try { $text = $reader.ReadToEnd() } finally { $reader.Dispose() }
    } catch { $problems.Add("The CSV file cannot be read ($Path): $($_.Exception.Message)"); return (& $result) }
    $csv = ConvertFrom-MbtBatchCsvText $text
    $delimiter = $csv.Delimiter
    $csvFolder = Split-Path ([IO.Path]::GetFullPath($Path)) -Parent
    if ($csv.Unclosed) { $problems.Add('A quote is not closed: the end of the file is inside a quoted text.') }
    if (-not $csv.Records.Count) { $problems.Add('The CSV file is empty: the first line names the columns, then one copy on each line.'); return (& $result) }

    $known = @{}
    foreach ($c in $script:MbtBatchCopyColumns) { $known[$c.ToLowerInvariant()] = $c }
    $header = $csv.Records[0]
    $column = @{}
    for ($i = 0; $i -lt $header.Fields.Count; $i++) {
        $name = $header.Fields[$i].Trim()
        if (-not $name) { continue }
        $k = $name.ToLowerInvariant()
        if ($column.ContainsKey($k)) { $problems.Add("The column $name is there twice (the first line): keep one."); continue }
        $column[$k] = $i
        if (-not $known.ContainsKey($k)) { $problems.Add("The column $name is not known (the columns are $($script:MbtBatchCopyColumns -join ', ')): check the spelling, or remove it.") }
    }
    if (-not $column.ContainsKey('sourceaddress') -or -not $column.ContainsKey('targetaddress')) {
        $problems.Add('The first line must name the columns, and SourceAddress and TargetAddress are required (found: ' + $(if ($column.Count) { (($header.Fields | ForEach-Object { $_.Trim() } | Where-Object { $_ }) -join ', ') } else { 'nothing' }) + ').')
        return (& $result)
    }
    $cell = { param($record, [string]$name) $at = $column[$name]; if ($null -ne $at -and $at -lt $record.Fields.Count) { $record.Fields[$at].Trim() } else { '' } }
    $address = '^[^@\s,;"<>]+@[^@\s,;"<>]+$'
    $seen = @{}
    $number = 0
    foreach ($record in @($csv.Records | Select-Object -Skip 1)) {
        $number++
        $line = $record.Line
        $before = $problems.Count

        $source = & $cell $record 'sourceaddress'
        if (-not $source) { $problems.Add("line ${line}: SourceAddress is empty."); $source = '' }
        elseif ($source -notmatch $address) { $problems.Add("line ${line}: '$source' is not a mailbox address (user@domain)."); $source = '' }
        $target = & $cell $record 'targetaddress'
        if (-not $target) { $problems.Add("line ${line}: TargetAddress is empty."); $target = '' }
        elseif ($target -notmatch $address) { $problems.Add("line ${line}: '$target' is not a mailbox address (user@domain)."); $target = '' }
        elseif ($SignedInUser -and -not $AppOnly -and $target -ine $SignedInUser) {
            $problems.Add("line ${line}: $target is not the signed-in account ($SignedInUser). With a delegated sign-in only the mailbox of the signed-in account can be written into (the service refuses a delegated write into another person's or a shared mailbox): sign in as the application (-AppOnly), or list only that mailbox as the target.")
        }

        $sourceLocation = $DefaultSourceLocation
        $given = & $cell $record 'sourcelocation'
        if ($given) {
            $match = @('Primary', 'Archive', 'All' | Where-Object { $_ -ieq $given })
            if ($match.Count) { $sourceLocation = $match[0] } else { $problems.Add("line ${line}: SourceLocation '$given' is not Primary, Archive or All (the Primary and the Archive together)."); $sourceLocation = '' }
        }
        $targetLocation = $DefaultTargetLocation
        $givenTarget = & $cell $record 'targetlocation'
        if ($givenTarget) {
            $match = @('Primary', 'Archive' | Where-Object { $_ -ieq $givenTarget })
            if ($match.Count) { $targetLocation = $match[0] } else { $problems.Add("line ${line}: TargetLocation '$givenTarget' is not Primary or Archive (a copy goes into one mailbox location; for both places use TargetLayout)."); $targetLocation = '' }
        }
        $layout = ''
        $givenLayout = & $cell $record 'targetlayout'
        if ($givenLayout) {
            $match = @('Mirror', 'Primary', 'Archive' | Where-Object { $_ -ieq $givenLayout })
            if ($match.Count) { $layout = $match[0] } else { $problems.Add("line ${line}: TargetLayout '$givenLayout' is not Mirror, Primary or Archive."); $layout = '' }
        }

        $selection = & $cell $record 'selectionfile'
        $selectionKind = ''
        if ($selection) {
            if (-not [IO.Path]::IsPathRooted($selection)) { $selection = [IO.Path]::Combine($csvFolder, $selection) }
            if (-not [IO.File]::Exists($selection)) { $problems.Add("line ${line}: the SelectionFile was not found: $selection"); $selection = '' }
            else {
                $selection = [IO.Path]::GetFullPath($selection)
                try {
                    $sets = @(Read-MbtBrowseSelectionSet -Path $selection)
                    $setKinds = @($sets | ForEach-Object { [string]$_.Kind } | Sort-Object -Unique)
                    if ($sets.Count -gt 2 -or ($sets.Count -eq 2 -and ($setKinds -join ',') -ne 'Archive,Primary')) { $problems.Add("line ${line}: the SelectionFile names $($sets.Count) mailbox locations ($($setKinds -join ', ')): a copy takes the Primary, the Archive, or both. Make one file for each place, on its own line."); $selection = '' }
                    elseif ($sets.Count -eq 1 -and [string]$sets[0].Kind -notin 'Primary', 'Archive') { $problems.Add("line ${line}: the SelectionFile is of the $($sets[0].Kind) mailbox: a copy reads the Primary or the Archive."); $selection = '' }
                    else {
                        $selectionKind = if ($sets.Count -eq 2) { 'All' } else { [string]$sets[0].Kind }
                        if ($given = & $cell $record 'sourcelocation') { if ($given -ine $selectionKind) { $problems.Add("line ${line}: SourceLocation $given does not match the SelectionFile, which is of the $(if ($selectionKind -eq 'All') { 'Primary and the Archive' } else { "$selectionKind" }) mailbox: leave SourceLocation empty."); $selection = '' } }
                        $sourceLocation = $selectionKind
                    }
                } catch { $problems.Add("line ${line}: the SelectionFile cannot be read: $($_.Exception.Message)"); $selection = '' }
            }
        }
        if ($source -and $sourceLocation -in 'Archive', 'All' -and $SignedInUser -and -not $AppOnly -and $source -ine $SignedInUser) {
            $problems.Add("line ${line}: $source is not the signed-in account ($SignedInUser). With a delegated sign-in only its primary mailbox can be read: use SourceLocation Primary, or sign in as the application (-AppOnly).")
        }

        $rootText = & $cell $record 'targetroot'
        if (-not $rootText) { $rootText = $DefaultTargetRoot }
        $root = ConvertTo-MbtBatchTargetRoot $rootText
        if ($rootText -and -not $root) { $problems.Add("line ${line}: TargetRoot '$rootText' is not a folder path.") }
        elseif ($root -and @($root -split '/' | Where-Object { -not $_.Trim() }).Count) { $problems.Add("line ${line}: TargetRoot '$rootText' has an empty folder name in it.") }
        $archiveRootText = & $cell $record 'targetarchiveroot'
        if (-not $archiveRootText) { $archiveRootText = $DefaultTargetArchiveRoot }
        $archiveRoot = ConvertTo-MbtBatchTargetRoot $archiveRootText
        if ($archiveRootText -and -not $archiveRoot) { $problems.Add("line ${line}: TargetArchiveRoot '$archiveRootText' is not a folder path.") }
        elseif ($archiveRoot -and @($archiveRoot -split '/' | Where-Object { -not $_.Trim() }).Count) { $problems.Add("line ${line}: TargetArchiveRoot '$archiveRootText' has an empty folder name in it.") }

        # a line that copies both places has a layout, not a target location; a line that copies one place has no layout and no archive folder
        $both = ($sourceLocation -eq 'All')
        if ($both) {
            if ($givenTarget) { $problems.Add("line ${line}: TargetLocation '$givenTarget' is for a line that copies one place: for the Primary and the Archive together leave it empty and use TargetLayout (Mirror, Primary or Archive).") }
            if (-not $layout) { $layout = if ($DefaultTargetLayout) { $DefaultTargetLayout } else { 'Mirror' } }
            if (-not $archiveRoot) { $archiveRoot = $root }
            $targetLocation = switch ($layout) { 'Mirror' { 'Primary + Archive' } default { $layout } }
        }
        else {
            if ($givenLayout) { $problems.Add("line ${line}: TargetLayout is for a line that copies both places (SourceLocation All): a copy of one place goes into the TargetLocation.") }
            if (& $cell $record 'targetarchiveroot') { $problems.Add("line ${line}: TargetArchiveRoot is for a line that copies both places (SourceLocation All): a copy of one place goes below TargetRoot.") }
            $layout = ''; $archiveRoot = ''
        }

        if ($source -and $target -and $sourceLocation -and $targetLocation -and $source -ieq $target) {
            $partPairs = [Collections.Generic.List[object]]::new()
            if ($both) {
                $partPairs.Add(@('Primary', $(if ($layout -eq 'Archive') { 'Archive' } else { 'Primary' })))
                $partPairs.Add(@('Archive', $(if ($layout -eq 'Primary') { 'Primary' } else { 'Archive' })))
            } else { $partPairs.Add(@($sourceLocation, $targetLocation)) }
            foreach ($pair in $partPairs) {
                if ($pair[0] -eq $pair[1]) { $problems.Add("line ${line}: the source and the target are the same mailbox location (the $($pair[0]) mailbox of $source): a copy writes into another location."); break }
            }
        }
        if ($problems.Count -gt $before) { continue }

        $twice = (@($source, $sourceLocation, $target, $targetLocation, $selection, $layout) | ForEach-Object { $_.ToLowerInvariant() }) -join '|'
        if ($seen.ContainsKey($twice)) { $problems.Add("line ${line}: $source ($sourceLocation) is already copied into $target ($targetLocation) on line $($seen[$twice]) (the same source, target, Locations and SelectionFile; upper and lower case are the same)."); continue }
        $seen[$twice] = $line
        $rows.Add([pscustomobject]@{ Number = $number; Line = $line; SourceAddress = $source; SourceLocation = $sourceLocation; TargetAddress = $target; TargetLocation = $targetLocation; TargetLayout = $layout; TargetRoot = $root; TargetArchiveRoot = $archiveRoot; SelectionFile = $selection; Store = '' })
    }
    if (-not $rows.Count -and -not $problems.Count) { $problems.Add('The CSV file has no copy: the first line names the columns, then one copy on each line.') }

    # the store of a line: named after the source location and the choice, so that a source copied into several targets is read once, and another choice never lands in it
    if ($StoreBase) {
        foreach ($r in $rows) {
            $key = Get-MbtCopyChoiceKey -Kind $r.SourceLocation -SelectionFile $r.SelectionFile -Folder $(if ($Choice.ContainsKey('Folder')) { $Choice.Folder }) -Period $(if ($Choice.ContainsKey('Period')) { $Choice.Period }) `
                -IncludeRecoverableItems:($Choice.ContainsKey('IncludeRecoverableItems') -and [bool]$Choice.IncludeRecoverableItems) -RecoverableItemsFolders $(if ($Choice.ContainsKey('RecoverableItemsFolders')) { $Choice.RecoverableItemsFolders })
            $r.Store = [IO.Path]::GetFullPath((Get-MbtCopyStorePath -StoreBase $StoreBase -Source $r.SourceAddress -Kind $r.SourceLocation -Key $key))
        }
    }
    return (& $result)
}

function Compare-MbtCopyBatchRow {
    <# The differences between the rows a batch of copies was made with and the rows of the CSV file given (the same sources, targets, locations, folders and selections), as sentences; none when they agree. #>
    param([Parameter(Mandatory)]$Batch, [Parameter(Mandatory)][object[]]$Row)
    $name = { param($source, $sourceLocation, $target, $targetLocation, $root, $selection, $archiveRoot) "$source ($sourceLocation) into $target ($targetLocation)" + $(if ($root) { " folder $root" }) + $(if ($archiveRoot) { ", archive folder $archiveRoot" }) + $(if ($selection) { ", selection $([IO.Path]::GetFileName($selection))" }) }
    $key = { param($source, $sourceLocation, $target, $targetLocation, $root, $selection, $archiveRoot) (@($source, $sourceLocation, $target, $targetLocation, $root, $selection, $archiveRoot) | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() }) -join '|' }
    $field = { param($row, [string]$name) if ($row -is [System.Collections.IDictionary]) { if ($row.Contains($name)) { [string]$row[$name] } else { '' } } elseif ($row.PSObject.Properties[$name]) { [string]$row.$name } else { '' } }
    $have = @{}; foreach ($r in @($Batch.rows)) { $archiveRoot = & $field $r 'archiveRoot'; $have[(& $key $r.sourceUpn $r.sourceLocation $r.upn $r.location $r.targetRoot $r.selection $archiveRoot)] = (& $name $r.sourceUpn $r.sourceLocation $r.upn $r.location $r.targetRoot $r.selection $archiveRoot) }
    $now = @{}; foreach ($r in $Row) { $archiveRoot = & $field $r 'TargetArchiveRoot'; $now[(& $key $r.SourceAddress $r.SourceLocation $r.TargetAddress $r.TargetLocation $r.TargetRoot $r.SelectionFile $archiveRoot)] = (& $name $r.SourceAddress $r.SourceLocation $r.TargetAddress $r.TargetLocation $r.TargetRoot $r.SelectionFile $archiveRoot) }
    $out = [Collections.Generic.List[string]]::new()
    foreach ($k in $now.Keys) { if (-not $have.ContainsKey($k)) { $out.Add("$($now[$k]) is in the CSV file and not in the batch") } }
    foreach ($k in $have.Keys) { if (-not $now.ContainsKey($k)) { $out.Add("$($have[$k]) is in the batch and not in the CSV file") } }
    return $out.ToArray()
}

#endregion

#region The state of a batch --------------------------------------------------------------------

function Read-MbtBatch {
    <# batch.json of a batch folder as a hashtable, or $null when there is none or it cannot be read. #>
    param([Parameter(Mandatory)][string]$Folder)
    $file = [IO.Path]::Combine($Folder, 'batch.json')
    if (-not [IO.File]::Exists($file)) { return $null }
    try { $b = Read-MbtJsonFile -Path $file -AsHashtable } catch { return $null }
    if ($b -isnot [hashtable] -and $b -isnot [Collections.Specialized.OrderedDictionary]) { return $null }
    if (-not $b.Contains('rows')) { return $null }
    $b.rows = @($b.rows | ForEach-Object { $_ })
    return $b
}

function Get-MbtBatchCount {
    <# What a batch holds, as numbers: rows by state, items done and failed, data. #>
    param([Parameter(Mandatory)]$Batch)
    $c = [ordered]@{ Total = 0; Pending = 0; Running = 0; Complete = 0; Partial = 0; Failed = 0; Done = [long]0; FailedItems = [long]0; Bytes = [long]0; OutsidePeriod = [long]0; Unknown = [long]0; Refused = [long]0; Exported = [long]0 }
    foreach ($r in @($Batch.rows)) {
        $c.Total++
        $state = [string]$r['state']
        switch ($state) { 'complete' { $c.Complete++ } 'partial' { $c.Partial++ } 'failed' { $c.Failed++ } 'running' { $c.Running++ } default { $c.Pending++ } }
        foreach ($pair in @(@('done', 'Done'), @('failedItems', 'FailedItems'), @('bytes', 'Bytes'), @('outsidePeriod', 'OutsidePeriod'), @('unknown', 'Unknown'), @('refused', 'Refused'), @('exported', 'Exported'))) { if ($r.Contains($pair[0]) -and $null -ne $r[$pair[0]]) { $c[$pair[1]] += [long]$r[$pair[0]] } }
    }
    return $c
}

function Get-MbtBatchState {
    <# complete (every mailbox), failed (one failed at least), stopped / partial (the others), pending (nothing started). #>
    param([Parameter(Mandatory)]$Batch)
    $c = Get-MbtBatchCount $Batch
    if ($c.Total -gt 0 -and $c.Complete -eq $c.Total) { return 'complete' }
    if ($c.Failed -gt 0) { return 'failed' }
    if ($c.Complete + $c.Partial + $c.Running -gt 0) { return 'partial' }
    return 'pending'
}

function Save-MbtBatch {
    <# Writes batch.json beside a temporary name and moves it into place: a reader never sees half a file. #>
    param([Parameter(Mandatory)]$Batch, [Parameter(Mandatory)][string]$Folder)
    $Batch.updated = [DateTime]::UtcNow.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z')
    $Batch.state = Get-MbtBatchState $Batch
    $Batch.counts = Get-MbtBatchCount $Batch
    [void][IO.Directory]::CreateDirectory($Folder)
    $json = $Batch | ConvertTo-Json -Depth 8
    $file = [IO.Path]::Combine($Folder, 'batch.json')
    $tmp = '{0}.{1}.tmp' -f $file, $PID
    [IO.File]::WriteAllText($tmp, $json, [Text.UTF8Encoding]::new($false))
    $attempt = 0
    while ($true) {
        $attempt++
        try { [IO.File]::Move($tmp, $file, $true); return }
        catch [IO.IOException], [UnauthorizedAccessException] { if ($attempt -ge 8) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue; throw }; Start-Sleep -Milliseconds (100 * $attempt) }
    }
}

function Write-MbtBatchLog {
    <# One line of batch.log in the batch folder (a time, then the text). Never stops a batch. #>
    param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Text)
    try { [IO.File]::AppendAllText([IO.Path]::Combine($Folder, 'batch.log'), ('{0} {1}{2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), ($Text -replace '[\r\n]+', ' | '), "`r`n"), [Text.UTF8Encoding]::new($false)) } catch { Write-Verbose "batch.log: $($_.Exception.Message)" }
}

function New-MbtBatch {
    <# A batch that has not started: every row pending. The rows are the ones of Read-MbtBatchCsv (-Mode Export), of Read-MbtImportBatchCsv (an import mode) or of Read-MbtCopyBatchCsv (-Mode Copy). #>
    param(
        [Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)][string]$Csv, [Parameter(Mandatory)][object[]]$Row, [Parameter(Mandatory)][hashtable]$Option, [int]$Parallel = 1,
        [ValidateSet('Export', 'Import', 'ImportPst', 'ImportEml', 'Copy')][string]$Mode = 'Export'
    )
    $now = [DateTime]::UtcNow.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z')
    $rows = @(foreach ($r in $Row) {
            if ($Mode -eq 'Copy') {
                [ordered]@{
                    n = [int]$r.Number; line = [int]$r.Line; upn = [string]$r.TargetAddress; location = [string]$r.TargetLocation; sourceUpn = [string]$r.SourceAddress; sourceLocation = [string]$r.SourceLocation
                    selection = [string]$r.SelectionFile; targetRoot = [string]$r.TargetRoot; layout = $(if ($r.PSObject.Properties['TargetLayout']) { [string]$r.TargetLayout } else { '' }); archiveRoot = $(if ($r.PSObject.Properties['TargetArchiveRoot']) { [string]$r.TargetArchiveRoot } else { '' }); store = [string]$r.Store; output = ''
                    state = 'pending'; job = ''; attempts = 0; exit = $null; reason = ''; message = ''; done = $null; failedItems = $null; bytes = $null; outsidePeriod = $null; unknown = $null; refused = $null; exported = $null; started = ''; finished = ''
                }
                continue
            }
            if ($Mode -ne 'Export') {
                [ordered]@{
                    n = [int]$r.Number; line = [int]$r.Line; upn = [string]$r.Mailbox; location = [string]$r.Location; selection = ''; output = ''; store = ''
                    source = [string]$r.Source; targetRoot = [string]$r.TargetRoot; sourceMailbox = [string]$r.SourceMailbox; key = [string]$r.Key; journal = [string]$r.Journal
                    state = 'pending'; job = ''; attempts = 0; exit = $null; reason = ''; message = ''; done = $null; failedItems = $null; bytes = $null; outsidePeriod = $null; unknown = $null; refused = $null; started = ''; finished = ''
                }
                continue
            }
            [ordered]@{
                n = [int]$r.Number; line = [int]$r.Line; upn = [string]$r.UserPrincipalName; location = [string]$r.Location; selection = [string]$r.SelectionFile; output = [string]$r.Output; store = [string]$r.Store
                state = 'pending'; job = ''; attempts = 0; exit = $null; reason = ''; message = ''; done = $null; failedItems = $null; bytes = $null; outsidePeriod = $null; started = ''; finished = ''
            }
        })
    $b = [ordered]@{ v = 1; id = $Id; name = $Id; created = $now; updated = $now; state = 'pending'; mode = $Mode; csv = $Csv; parallel = $Parallel; options = $Option; counts = $null; rows = $rows }
    $b.counts = Get-MbtBatchCount $b
    return $b
}

function Compare-MbtBatchRow {
    <# The differences between the rows a batch was made with and the rows of the CSV now (the same mailboxes, the same locations and selections), as sentences; none when they agree. #>
    param([Parameter(Mandatory)]$Batch, [Parameter(Mandatory)][object[]]$Row)
    $name = { param($upn, $location, $selection) if ($selection) { "$upn (selection $([IO.Path]::GetFileName($selection)))" } else { "$upn ($location)" } }
    $key = { param($upn, $location, $selection) ([string]$upn).ToLowerInvariant() + '|' + $(if ($selection) { 'selection:' + ([string]$selection).ToLowerInvariant() } else { [string]$location }) }
    $have = @{}; foreach ($r in @($Batch.rows)) { $have[(& $key $r.upn $r.location $r.selection)] = (& $name $r.upn $r.location $r.selection) }
    $now = @{}; foreach ($r in $Row) { $now[(& $key $r.UserPrincipalName $r.Location $r.SelectionFile)] = (& $name $r.UserPrincipalName $r.Location $r.SelectionFile) }
    $out = [Collections.Generic.List[string]]::new()
    foreach ($k in $now.Keys) { if (-not $have.ContainsKey($k)) { $out.Add("$($now[$k]) is in the CSV file and not in the batch") } }
    foreach ($k in $have.Keys) { if (-not $now.ContainsKey($k)) { $out.Add("$($have[$k]) is in the batch and not in the CSV file") } }
    return $out.ToArray()
}

function Compare-MbtImportBatchRow {
    <# The differences between the rows an import batch was made with and the rows of the CSV file given (the same targets, locations, sources and folders), as sentences; none when they agree. #>
    param([Parameter(Mandatory)]$Batch, [Parameter(Mandatory)][object[]]$Row)
    $name = { param($mailbox, $location, $source, $root) "$mailbox ($location) from $source" + $(if ($root) { " into $root" }) }
    $key = { param($mailbox, $location, $source, $sourceMailbox, $root) (@($mailbox, $location, $source, $sourceMailbox, $root) | ForEach-Object { ([string]$_).Trim().ToLowerInvariant() }) -join '|' }
    $have = @{}; foreach ($r in @($Batch.rows)) { $have[(& $key $r.upn $r.location $r.source $r.sourceMailbox $r.targetRoot)] = (& $name $r.upn $r.location $r.source $r.targetRoot) }
    $now = @{}; foreach ($r in $Row) { $now[(& $key $r.Mailbox $r.Location $r.Source $r.SourceMailbox $r.TargetRoot)] = (& $name $r.Mailbox $r.Location $r.Source $r.TargetRoot) }
    $out = [Collections.Generic.List[string]]::new()
    foreach ($k in $now.Keys) { if (-not $have.ContainsKey($k)) { $out.Add("$($now[$k]) is in the CSV file and not in the batch") } }
    foreach ($k in $have.Keys) { if (-not $now.ContainsKey($k)) { $out.Add("$($have[$k]) is in the batch and not in the CSV file") } }
    return $out.ToArray()
}

function Get-MbtBatchList {
    <# The batches of a work folder, newest first, with their counts (the console lists them; they are only read). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkFolder, [int]$Limit = 20)
    $root = [IO.Path]::Combine($WorkFolder, 'batches')
    if (-not [IO.Directory]::Exists($root)) { return @() }
    $list = foreach ($d in [IO.Directory]::EnumerateDirectories($root)) {
        $b = Read-MbtBatch -Folder $d
        if ($null -eq $b) { continue }
        $counts = Get-MbtBatchCount $b
        $report = [IO.Path]::Combine($d, 'report.html')
        [pscustomobject]@{
            Id = [string]$b.id; Mode = (Get-MbtBatchMode $b); State = (Get-MbtBatchState $b); Created = [string]$b.created; Updated = [string]$b.updated; Folder = $d
            Total = $counts.Total; Complete = $counts.Complete; Partial = $counts.Partial; Failed = $counts.Failed; Pending = $counts.Pending; Running = $counts.Running
            Report = $(if ([IO.File]::Exists($report)) { $report } else { '' }); Jobs = @($b.rows | ForEach-Object { [string]$_.job } | Where-Object { $_ })
        }
    }
    return @($list | Sort-Object Updated -Descending | Select-Object -First $Limit)
}

function Get-MbtBatchProtection {
    <#
        What the retention of the job folders must keep: the names of the batches that are not finished (a mailbox is not complete, or batch.json cannot be read) and the names of their job folders.
        Returns Batches and Jobs, two sets (not case sensitive).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkFolder)
    $batches = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $jobs = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $root = [IO.Path]::Combine($WorkFolder, 'batches')
    if ([IO.Directory]::Exists($root)) {
        foreach ($d in [IO.Directory]::EnumerateDirectories($root)) {
            $name = [IO.Path]::GetFileName($d)
            $b = Read-MbtBatch -Folder $d
            if ($null -eq $b) { [void]$batches.Add($name); continue }
            if ((Get-MbtBatchState $b) -eq 'complete') { continue }
            [void]$batches.Add($name)
            if ($b.Contains('id') -and $b.id) { [void]$batches.Add([string]$b.id) }
            foreach ($r in @($b.rows)) { if ($r.Contains('job') -and $r.job) { [void]$jobs.Add([string]$r.job) } }
        }
    }
    return [pscustomobject]@{ Batches = $batches; Jobs = $jobs }
}

#endregion

#region One mailbox -----------------------------------------------------------------------------

function Get-MbtBatchRowLock {
    <#
        What a row holds while it runs, as names: an import holds its target mailbox location and its source. The service throttles each mailbox, and an export store or a folder of .eml files is
        opened by one process at a time, so two rows that share a name never run together. A copy holds its target mailbox location and its source mailbox location (the store is read and written by one
        process at a time, and the service throttles the source too). An export row holds nothing.
    #>
    param([Parameter(Mandatory)]$Row)
    if ($Row -is [System.Collections.IDictionary] -and $Row.Contains('sourceUpn') -and [string]$Row['sourceUpn']) {
        # a line that copies both places holds both source places and the target places its layout writes into
        $sourcePlaces = if ([string]$Row['sourceLocation'] -eq 'All') { @('primary', 'archive') } else { @(([string]$Row['sourceLocation']).ToLowerInvariant()) }
        $layout = if ($Row.Contains('layout')) { [string]$Row['layout'] } else { '' }
        $targetPlaces = if ([string]$Row['sourceLocation'] -ne 'All') { @(([string]$Row['location']).ToLowerInvariant()) } elseif ($layout -eq 'Primary') { @('primary') } elseif ($layout -eq 'Archive') { @('archive') } else { @('primary', 'archive') }
        return @(@($targetPlaces | ForEach-Object { 'target:' + ([string]$Row['upn']).ToLowerInvariant() + '|' + $_ }) + @($sourcePlaces | ForEach-Object { 'source:' + ([string]$Row['sourceUpn']).ToLowerInvariant() + '|' + $_ }))
    }
    if (-not ($Row -is [System.Collections.IDictionary]) -or -not $Row.Contains('source') -or -not [string]$Row['source']) { return @() }
    return @(('target:' + ([string]$Row['upn']).ToLowerInvariant() + '|' + ([string]$Row['location']).ToLowerInvariant()), ('source:' + ([string]$Row['source']).ToLowerInvariant()))
}

function Get-MbtBatchChildArgument {
    <#
        The arguments of Invoke-MailboxTransfer.ps1 for one mailbox of a batch: an ordinary export of that mailbox, silent sign-in (nothing is ever asked in the middle of a run), the files of its job.
        The Continue of the console runs the same arguments again (it adds the log folder and the development switch itself).
    #>
    param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)][hashtable]$Option, [Parameter(Mandatory)]$Job, [Parameter(Mandatory)][string]$BatchId, [string]$Mode = 'Export')
    if ($Mode -in $script:MbtBatchImportModes) { return (Get-MbtImportChildArgument -Row $Row -Option $Option -Job $Job -BatchId $BatchId -Mode $Mode) }
    if ($Mode -eq 'Copy') { return (Get-MbtCopyChildArgument -Row $Row -Option $Option -Job $Job -BatchId $BatchId) }
    $a = [Collections.Generic.List[string]]::new()
    $a.AddRange([string[]]@('-Mode', 'Export'))
    $kind = switch ([string]$Row.location) { 'Archive' { 'Archive' } 'Both' { 'All' } default { 'Primary' } }
    $signer = [string]$Option.SignInUser
    $other = (-not $Option.AppOnly) -and $signer -and ([string]$Row.upn -ine $signer)
    if ($Option.AppOnly -or -not $signer) { $a.AddRange([string[]]@('-UserPrincipalName', [string]$Row.upn)) }
    elseif ($other) { $a.AddRange([string[]]@('-UserPrincipalName', $signer, '-MailboxId', [string]$Row.upn)) }
    else { $a.AddRange([string[]]@('-UserPrincipalName', $signer)) }
    if ([string]$Row.selection) { $a.AddRange([string[]]@('-SelectionFile', [string]$Row.selection)) }
    elseif (-not $other) { $a.AddRange([string[]]@('-Mailbox', $kind)) }
    $a.AddRange([string[]]@('-StorePath', [string]$Row.store, '-TempPath', [IO.Path]::Combine($Job.Folder, 'temp')))
    if ([string]$Row.output) { $a.AddRange([string[]]@('-OutputPst', [string]$Row.output)) }
    foreach ($name in @($Option.Pass.Keys | Sort-Object)) {
        $value = $Option.Pass[$name]
        if ($value -is [bool]) { if ($value) { $a.Add("-$name") } }
        elseif ($null -ne $value -and [string]$value -ne '') { $a.AddRange([string[]]@("-$name", [string]$value)) }
    }
    $a.AddRange([string[]]@('-NoPrompt', '-EventsFile', $Job.EventsPath, '-StopFile', $Job.StopPath, '-JobId', $Job.Id, '-BatchId', $BatchId))
    return $a.ToArray()
}

function Get-MbtImportChildArgument {
    <#
        The arguments of Invoke-MailboxTransfer.ps1 for one import of a batch: an ordinary import of that source into that mailbox, silent sign-in, the files of its job. A store import gets a job name made
        of the key of the row, a PST import a journal folder made of it: the same row again, in this batch or another, finds what it did and creates nothing twice.
        A delegated sign-in writes into the mailbox of the signed-in account only (the CSV file was checked for it); the application sign-in names the target mailbox itself.
    #>
    param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)][hashtable]$Option, [Parameter(Mandatory)]$Job, [Parameter(Mandatory)][string]$BatchId, [Parameter(Mandatory)][string]$Mode)
    $a = [Collections.Generic.List[string]]::new()
    $a.AddRange([string[]]@('-Mode', $Mode))
    $signer = [string]$Option.SignInUser
    if (-not $Option.AppOnly -and $signer -and [string]$Row.upn -ine $signer) { throw "A delegated sign-in writes into the mailbox of the signed-in account ($signer) only: $($Row.upn) needs the application sign-in (-AppOnly). Nothing was written." }
    $a.AddRange([string[]]@('-UserPrincipalName', $(if ($Option.AppOnly -or -not $signer) { [string]$Row.upn } else { $signer }), '-Mailbox', [string]$Row.location))
    $temp = [IO.Path]::Combine($Job.Folder, 'temp')
    switch ($Mode) {
        'Import' {
            $a.AddRange([string[]]@('-StorePath', [string]$Row.source, '-JobName', ('import-' + [string]$Row.key), '-TempPath', $temp))
            if ([string]$Row.sourceMailbox) { $a.AddRange([string[]]@('-SourceMailbox', [string]$Row.sourceMailbox)) }
        }
        'ImportPst' { $a.AddRange([string[]]@('-PstPath', [string]$Row.source, '-TempPath', $(if ([string]$Row.journal) { [string]$Row.journal } else { $temp }))) }
        'ImportEml' { $a.AddRange([string[]]@('-EmlPath', [string]$Row.source, '-TempPath', $temp)) }
    }
    if ([string]$Row.targetRoot) { $a.AddRange([string[]]@('-TargetRoot', [string]$Row.targetRoot)) }
    foreach ($name in @($Option.Pass.Keys | Sort-Object)) {
        $value = $Option.Pass[$name]
        if ($value -is [bool]) { if ($value) { $a.Add("-$name") } }
        elseif ($null -ne $value -and [string]$value -ne '') { $a.AddRange([string[]]@("-$name", [string]$value)) }
    }
    $a.AddRange([string[]]@('-NoPrompt', '-EventsFile', $Job.EventsPath, '-StopFile', $Job.StopPath, '-JobId', $Job.Id, '-BatchId', $BatchId))
    return $a.ToArray()
}

function Get-MbtCopyChildArgument {
    <#
        The arguments of Invoke-MailboxTransfer.ps1 for one copy of a batch: an ordinary copy of that source into that target, silent sign-in, the store and the files of its job. The import job of the copy is named
        by the command itself after the target, the store and the folder, so the same line again, in this batch or another, creates nothing twice.
        A delegated sign-in writes into the mailbox of the signed-in account only (the CSV file was checked for it) and reads another mailbox by its address; the application sign-in names the source (the mailbox it
        works on) and the target (-TargetMailboxId) itself. A selection names its own mailbox and location.
    #>
    param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)][hashtable]$Option, [Parameter(Mandatory)]$Job, [Parameter(Mandatory)][string]$BatchId)
    $signer = [string]$Option.SignInUser
    if (-not $Option.AppOnly -and $signer -and [string]$Row.upn -ine $signer) { throw "A delegated sign-in writes into the mailbox of the signed-in account ($signer) only: $($Row.upn) needs the application sign-in (-AppOnly). Nothing was read or written." }
    $a = [Collections.Generic.List[string]]::new()
    $a.AddRange([string[]]@('-Mode', 'Copy'))
    $other = (-not $Option.AppOnly) -and $signer -and ([string]$Row.sourceUpn -ine $signer)
    if ($Option.AppOnly -or -not $signer) { $a.AddRange([string[]]@('-UserPrincipalName', [string]$Row.sourceUpn, '-TargetMailboxId', [string]$Row.upn)) }
    elseif ($other -and -not [string]$Row.selection) { $a.AddRange([string[]]@('-UserPrincipalName', $signer, '-MailboxId', [string]$Row.sourceUpn)) }
    else { $a.AddRange([string[]]@('-UserPrincipalName', $signer)) }
    $both = ([string]$Row.sourceLocation -eq 'All')
    if ([string]$Row.selection) { $a.AddRange([string[]]@('-SelectionFile', [string]$Row.selection)) }
    else { $a.AddRange([string[]]@('-Mailbox', [string]$Row.sourceLocation)) }
    if ($both) { $a.AddRange([string[]]@('-TargetLayout', $(if ($Row.Contains('layout') -and [string]$Row.layout) { [string]$Row.layout } else { 'Mirror' }))) }
    else { $a.AddRange([string[]]@('-TargetMailbox', [string]$Row.location)) }
    if ([string]$Row.targetRoot) { $a.AddRange([string[]]@('-TargetRoot', [string]$Row.targetRoot)) }
    if ($both -and $Row.Contains('archiveRoot') -and [string]$Row.archiveRoot) { $a.AddRange([string[]]@('-TargetArchiveRoot', [string]$Row.archiveRoot)) }
    $a.AddRange([string[]]@('-StorePath', [string]$Row.store, '-TempPath', [IO.Path]::Combine($Job.Folder, 'temp')))
    foreach ($name in @($Option.Pass.Keys | Sort-Object)) {
        $value = $Option.Pass[$name]
        if ($value -is [bool]) { if ($value) { $a.Add("-$name") } }
        elseif ($null -ne $value -and [string]$value -ne '') { $a.AddRange([string[]]@("-$name", [string]$value)) }
    }
    $a.AddRange([string[]]@('-NoPrompt', '-EventsFile', $Job.EventsPath, '-StopFile', $Job.StopPath, '-JobId', $Job.Id, '-BatchId', $BatchId))
    return $a.ToArray()
}

function New-MbtBatchJobFolder {
    <# A job folder like the console makes (<work folder>\jobs\<yyyyMMdd-HHmm>-<export, import or mailcopy>-<4 hex>), for a mailbox of a batch. #>
    param([Parameter(Mandatory)][string]$WorkFolder, [string]$Name, [ValidateSet('export', 'import', 'mailcopy')][string]$Kind = 'export')
    $jobs = [IO.Path]::Combine($WorkFolder, 'jobs')
    if ($Name) { $dir = [IO.Path]::Combine($jobs, $Name); $id = ($Name -replace '^.*-(export|import|mailcopy)-', '$1-') }
    else {
        do {
            $hex = -join ((1..4) | ForEach-Object { '0123456789abcdef'[(Get-Random -Maximum 16)] })
            $Name = '{0}-{1}-{2}' -f (Get-Date).ToString('yyyyMMdd-HHmm'), $Kind, $hex
            $dir = [IO.Path]::Combine($jobs, $Name)
        } while ([IO.Directory]::Exists($dir))
        $id = "$Kind-$hex"
    }
    [void][IO.Directory]::CreateDirectory($dir)
    return [pscustomobject]@{
        Id = $id; Name = $Name; Folder = $dir; EventsPath = [IO.Path]::Combine($dir, 'events.jsonl'); StopPath = [IO.Path]::Combine($dir, 'stop.request'); StatusPath = [IO.Path]::Combine($dir, 'status.json')
        StdOut = [IO.Path]::Combine($dir, 'stdout.txt'); StdErr = [IO.Path]::Combine($dir, 'stderr.txt')
    }
}

function Write-MbtBatchJobRequest {
    <# request.json of the job, as the console writes it: the Continue list of the console reads it, and runs the same arguments again in the same folder. No secret in it. #>
    param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)]$Job, [Parameter(Mandatory)][string[]]$Argument, [Parameter(Mandatory)][hashtable]$Option, [Parameter(Mandatory)][string]$BatchId, [string]$Mode = 'Export')
    $pass = $Option.Pass
    $get = { param($n, $d) if ($pass.Contains($n) -and $null -ne $pass[$n]) { $pass[$n] } else { $d } }
    if ($Mode -eq 'Copy') {
        $signer = [string]$Option.SignInUser
        $app = [bool]$Option.AppOnly -or -not $signer
        $request = [ordered]@{
            Title = ('Copy {0} {1} {2} {3} ({4})' -f [char]0x00B7, $Row.sourceUpn, [char]0x2192, $Row.upn, $Row.location); Created = [DateTime]::UtcNow.ToString('o'); JobId = $Job.Id; JobName = $Job.Id; Batch = $BatchId
            StorePath = [string]$Row.store; SelectionFiles = @(if ([string]$Row.selection) { [string]$Row.selection }); Commands = @(, $Argument)
            Task = 'mailcopy'; Mailbox = $(if ($app) { [string]$Row.sourceUpn } else { $signer }); Shared = $(if (-not $app -and [string]$Row.sourceUpn -ine $signer) { [string]$Row.sourceUpn } else { '' }); McTarget = $(if ($app) { [string]$Row.upn } else { '' })
            TgtLoc = $(if ([string]$Row.sourceLocation -eq 'All') { $(if ($Row.Contains('layout') -and [string]$Row.layout -eq 'Archive') { 'Archive' } else { 'Primary' }) } else { [string]$Row.location }); TgtMode = $(if ([string]$Row.targetRoot) { 'new' } else { 'same' }); TgtName = [string]$Row.targetRoot
            Conc = [int](& $get 'Concurrency' 4); Tries = [int](& $get 'MaxAttempts' 3); Awake = $false
        }
        if ([string]$Row.sourceLocation -eq 'All') { $request.McFromKind = 'All'; $request.McLayout = $(if ($Row.Contains('layout') -and [string]$Row.layout) { [string]$Row.layout } else { 'Mirror' }); $request.TgtArchiveName = $(if ($Row.Contains('archiveRoot')) { [string]$Row.archiveRoot } else { '' }) }
        if ($pass.Contains('RetryWaitSeconds')) { $request.Retry = [int][math]::Round([double]$pass.RetryWaitSeconds) }
        [IO.File]::WriteAllText([IO.Path]::Combine($Job.Folder, 'request.json'), ($request | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
        return
    }
    if ($Mode -in $script:MbtBatchImportModes) {
        $request = [ordered]@{
            Title = ('Import {0} {1} ({2})' -f [char]0x00B7, $Row.upn, $Row.location); Created = [DateTime]::UtcNow.ToString('o'); JobId = $Job.Id; JobName = $Job.Id; Batch = $BatchId
            StorePath = $(if ($Mode -eq 'Import') { [string]$Row.source } else { '' }); SelectionFiles = @(); Commands = @(, $Argument)
            Task = 'import'; Mailbox = [string]$Row.upn; ImSrc = $(switch ($Mode) { 'ImportPst' { 'pst' } 'ImportEml' { 'eml' } default { 'store' } }); ImPath = [string]$Row.source
            TgtLoc = [string]$Row.location; TgtMode = $(if ([string]$Row.targetRoot) { 'new' } else { 'same' }); TgtName = [string]$Row.targetRoot
            Conc = [int](& $get 'Concurrency' 4); Tries = [int](& $get 'MaxAttempts' 3); Awake = $false
        }
        if ($pass.Contains('RetryWaitSeconds')) { $request.Retry = [int][math]::Round([double]$pass.RetryWaitSeconds) }
        [IO.File]::WriteAllText([IO.Path]::Combine($Job.Folder, 'request.json'), ($request | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
        return
    }
    $output = [string]$Row.output
    $req = [ordered]@{
        Title = ('Export {0} {1}' -f [char]0x00B7, $Row.upn); Created = [DateTime]::UtcNow.ToString('o'); JobId = $Job.Id; JobName = $Job.Id; Batch = $BatchId
        StorePath = [string]$Row.store; SelectionFiles = @(if ([string]$Row.selection) { [string]$Row.selection })
        Commands = @(, $Argument)
        Task = 'export'; Mailbox = [string]$Row.upn; Dest = $(if ($output) { 'pst' } else { 'store' }); KeepStore = $true
        OutFolder = $(if ($output) { Split-Path $output -Parent } else { '' }); OutName = $(if ($output) { Split-Path $output -Leaf } else { '' })
        SplitGB = [int][math]::Ceiling([double](& $get 'PstMaxSizeGB' 50)); CheckEach = (-not [bool](& $get 'SkipCheck' $false)); SepAux = [bool](& $get 'SeparateAuxArchives' $false)
        Conc = [int](& $get 'Concurrency' 4); Tries = [int](& $get 'MaxAttempts' 3); Awake = $false
    }
    if ($pass.Contains('RetryWaitSeconds')) { $req.Retry = [int][math]::Round([double]$pass.RetryWaitSeconds) }
    [IO.File]::WriteAllText([IO.Path]::Combine($Job.Folder, 'request.json'), ($req | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
}

function Start-MbtBatchProcess {
    <# Starts the child pwsh of a mailbox: no window, standard input closed, its output copied to stdout.txt and stderr.txt of the job folder. Returns the handle (Process, Pipes). #>
    param([Parameter(Mandatory)][string]$EntryScript, [Parameter(Mandatory)][string[]]$Argument, [Parameter(Mandatory)]$Job, [string]$LogFolder, [switch]$DevelopmentMode)
    $all = [Collections.Generic.List[string]]::new()
    $all.AddRange([string[]]@('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $EntryScript))
    if ($DevelopmentMode) { $all.Add('-DevelopmentMode') }
    if ($LogFolder) { $all.AddRange([string[]]@('-LogFolder', $LogFolder)) }
    $all.AddRange($Argument)
    $psi = [Diagnostics.ProcessStartInfo]::new((Get-Process -Id $PID).Path)
    foreach ($x in $all) { $psi.ArgumentList.Add($x) }
    $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true
    $psi.WorkingDirectory = Split-Path $EntryScript -Parent
    $psi.RedirectStandardInput = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
    $process = [Diagnostics.Process]::Start($psi)
    $process.StandardInput.Close()
    $pipes = foreach ($pair in @(@($process.StandardOutput.BaseStream, $Job.StdOut), @($process.StandardError.BaseStream, $Job.StdErr))) {
        $file = [IO.FileStream]::new($pair[1], [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
        @{ File = $file; Task = $pair[0].CopyToAsync($file) }
    }
    return [pscustomobject]@{ Process = $process; Pipes = @($pipes) }
}

function Get-MbtBatchExitCode {
    <# The exit code of a child that has ended, read before the process is let go; 1 when it cannot be read. #>
    param([Parameter(Mandatory)]$Handle)
    try { return [int]$Handle.Process.ExitCode } catch { return 1 }
}

function Close-MbtBatchHandle {
    <# Ends the copy of the output of a child to its files (waits a few seconds for the end of the data) and lets go of the process. #>
    param([Parameter(Mandatory)]$Handle)
    foreach ($p in @($Handle.Pipes)) {
        if (-not $p) { continue }
        try { $null = $p.Task.Wait(3000) } catch { Write-Verbose "pipe: $_" }
        try { $p.File.Dispose() } catch { Write-Verbose "pipe: $_" }
    }
    $Handle.Pipes = @()
    try { $Handle.Process.Dispose() } catch { Write-Verbose "process: $_" }
}

function Get-MbtBatchFailureText {
    <# Why a mailbox failed, in a sentence: the last error event of its job, else the last line of its error output, else the exit code. #>
    param([Parameter(Mandatory)]$Job, [int]$ExitCode)
    try {
        if ([IO.File]::Exists($Job.EventsPath)) {
            $fs = [IO.FileStream]::new($Job.EventsPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
            try {
                $take = [int][math]::Min($fs.Length, 131072)
                [void]$fs.Seek(-$take, [IO.SeekOrigin]::End)
                $buffer = [byte[]]::new($take); $got = $fs.Read($buffer, 0, $take)
                $lines = [Text.Encoding]::UTF8.GetString($buffer, 0, $got) -split "`n"
            } finally { $fs.Dispose() }
            for ($i = $lines.Count - 1; $i -ge 0; $i--) {
                if ($lines[$i] -notmatch '"e"\s*:\s*"error"') { continue }
                try { $e = $lines[$i] | ConvertFrom-Json -AsHashtable } catch { continue }
                $msg = [string]$e['msg']; $tech = [string]$e['tech']
                if ($tech -and $tech.Length -gt 220) { $tech = $tech.Substring(0, 220) + '...' }
                if ($msg -and $tech -and $tech -ne $msg) { return "$msg ($tech)" }
                if ($msg) { return $msg }
            }
        }
    } catch { Write-Verbose "events: $($_.Exception.Message)" }
    try {
        if ([IO.File]::Exists($Job.StdErr)) {
            $last = @([IO.File]::ReadAllLines($Job.StdErr) | Where-Object { $_.Trim() } | Select-Object -Last 1)
            if ($last.Count) { $text = ($last[0] -replace '\x1b\[[0-9;]*m', '').Trim(); if ($text.Length -gt 240) { $text = $text.Substring(0, 240) + '...' }; return $text }
        }
    } catch { Write-Verbose "stderr: $($_.Exception.Message)" }
    return "the process ended with exit code $ExitCode and no message"
}

#endregion

#region The report ------------------------------------------------------------------------------

function Write-MbtBatchReport {
    <#
    .SYNOPSIS
        <batch folder>\report.html: the overview of the jobs of the batch (the page of the overview of all the jobs, filtered), with the mailboxes that have not started said above the table.
        Stopped after -TimeBudgetSeconds (the page that is there stays; $null is returned). Returns the path.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)]$Batch, [Parameter(Mandatory)][string]$Folder, [int]$TimeBudgetSeconds = 0, [hashtable]$Info)
    $counts = Get-MbtBatchCount $Batch
    $notStarted = @($Batch.rows | Where-Object { -not $_.job }).Count
    $caption = '{0} {1} in the CSV file ({2} complete, {3}{4} not complete, {5} failed)' -f (Format-MbtNumber $counts.Total), $(if ($counts.Total -eq 1) { 'mailbox' } else { 'mailboxes' }), (Format-MbtNumber $counts.Complete), $(if ($counts.Running) { (Format-MbtNumber $counts.Running) + ' running, ' } else { '' }), (Format-MbtNumber $counts.Partial), (Format-MbtNumber $counts.Failed)
    if ($notStarted) { $caption += '; ' + (Format-MbtNumber $notStarted) + $(if ($notStarted -eq 1) { ' has' } else { ' have' }) + ' not started yet' }
    if ((Get-MbtBatchMode $Batch) -eq 'Copy' -and ($counts.Done -gt 0 -or $counts.Exported -gt 0 -or $counts.FailedItems -gt 0 -or $counts.Unknown -gt 0)) {
        $caption += '; ' + (Format-MbtNumber $counts.Exported) + ' items read, ' + (Format-MbtNumber $counts.Done) + ' items copied, ' + (Format-MbtNumber $counts.FailedItems) + ' failed, ' + (Format-MbtNumber $counts.Unknown) + ' uncertain'
    }
    elseif ((Get-MbtBatchMode $Batch) -ne 'Export' -and ($counts.Done -gt 0 -or $counts.FailedItems -gt 0 -or $counts.Refused -gt 0 -or $counts.Unknown -gt 0)) {
        $caption += '; ' + (Format-MbtNumber $counts.Done) + ' items imported, ' + (Format-MbtNumber ($counts.FailedItems + $counts.Refused)) + ' refused, ' + (Format-MbtNumber $counts.Unknown) + ' uncertain'
    }
    $caption += '.'
    $names = @($Batch.rows | ForEach-Object { [string]$_.job } | Where-Object { $_ })
    $splat = @{ WorkFolder = $WorkFolder; Batch = [string]$Batch.id; JobFolderName = $names; OutputPath = [IO.Path]::Combine($Folder, 'report.html'); LinkBase = '../../jobs/'; Caption = $caption; TimeBudgetSeconds = $TimeBudgetSeconds }
    if ($Info) { $splat.Info = $Info }
    return (Write-MbtJobOverview @splat)
}

#endregion

#region Running the rows ------------------------------------------------------------------------

function Complete-MbtBatchRow {
    <# A child has ended: the row takes the result of its job (state, counts, reason) and says it on one line. #>
    param([Parameter(Mandatory)]$Batch, [Parameter(Mandatory)]$Item, [Parameter(Mandatory)][string]$Folder, [bool]$StopAsked, [Parameter(Mandatory)][string]$BatchId, [int]$Position, [int]$Of)
    $row = $Item.Row; $job = $Item.Job
    $mode = Get-MbtBatchMode $Batch
    $task = if ($mode -eq 'Export') { 'export' } elseif ($mode -eq 'Copy') { 'mailcopy' } else { 'import' }
    $exit = if ($null -ne $Item.ExitCode) { [int]$Item.ExitCode } else { 1 }
    $status = $null
    if ([IO.File]::Exists($job.StatusPath)) { try { $status = Read-MbtJsonFile -Path $job.StatusPath -AsHashtable } catch { $status = $null } }
    $reason = if ($null -ne $status -and $status.Contains('reason') -and $status.reason) { [string]$status.reason } else { '' }
    $counts = if ($null -ne $status -and $status.Contains('counts') -and $status.counts -is [System.Collections.IDictionary]) { $status.counts } else { @{} }
    $hasResult = $null -ne $status -and $status.Contains('exitCode') -and $null -ne $status.exitCode
    $message = ''
    if ($exit -eq 0) { $state = 'complete' }
    elseif ($exit -eq 2 -and $reason -eq 'signin-required') { $state = 'partial'; $message = 'sign-in required: no valid sign-in is cached, so nothing could be asked in the middle of a run. Sign in once (run the command again without -NoPrompt) and run it again.' }
    elseif ($exit -eq 2) { $state = 'partial'; $message = if ($reason -eq 'stopped' -or $StopAsked) { 'stopped: run the batch again to continue' } else { 'not complete: run the batch again to continue' } }
    elseif ($exit -eq 3 -and $task -in 'import', 'mailcopy') { $state = 'partial'; $message = 'uncertain items: the answer for some items was lost, they may exist. Run the batch again: they are looked up in the mailbox first (or settle them with -Mode Status -Unknown -Check)' }
    elseif ($exit -eq 1 -and $reason -eq 'refused-items') { $state = 'failed'; $message = 'finished, but some items are refused for good: the report of its job lists them' }
    else { $state = 'failed' }
    if ($state -eq 'failed' -and -not $message) { $message = Get-MbtBatchFailureText -Job $job -ExitCode $exit }
    # a child that ended before it could write its own end: the job still says so in the overview
    if (-not $hasResult -and $state -eq 'failed') {
        try { Write-MbtJobStatus -JobFolder $job.Folder -State error -JobId $job.Id -Task $task -Mode $mode -Source $(if ($row.Contains('sourceUpn')) { [string]$row.sourceUpn } else { [string]$row.upn }) -Batch $BatchId -ExitCode $exit -Reason 'start-failed' -Next retry }
        catch { Write-Verbose "status: $($_.Exception.Message)" }
    }
    $row.state = $state; $row.exit = $exit; $row.reason = $reason; $row.message = $message
    $row.finished = [DateTime]::UtcNow.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z')
    foreach ($pair in @(@('done', 'done'), @('failed', 'failedItems'), @('bytes', 'bytes'), @('outsidePeriod', 'outsidePeriod'), @('unknown', 'unknown'), @('refused', 'refused'), @('exported', 'exported'))) { $v = $counts[$pair[0]]; if ($null -ne $v) { $row[$pair[1]] = [long]$v } }
    $seconds = ([DateTime]::UtcNow - $Item.Started).TotalSeconds
    $detail = @()
    if ($null -ne $row.done) { $detail += (Format-MbtNumber $row.done) + $(if ($mode -eq 'Copy') { ' copied' } elseif ($mode -ne 'Export') { ' imported' } elseif ([long]$row.done -eq 1) { ' item' } else { ' items' }) }
    if ($null -ne $row.failedItems -and [long]$row.failedItems -gt 0) { $detail += (Format-MbtNumber $row.failedItems) + ' failed' }
    if ($row.Contains('refused') -and $null -ne $row.refused -and [long]$row.refused -gt 0) { $detail += (Format-MbtNumber $row.refused) + ' refused for good' }
    if ($row.Contains('unknown') -and $null -ne $row.unknown -and [long]$row.unknown -gt 0) { $detail += (Format-MbtNumber $row.unknown) + ' uncertain' }
    if ($null -ne $row.bytes -and [long]$row.bytes -gt 0) { $detail += (Format-MbtBytes ([double]$row.bytes)) }
    if ($row.Contains('outsidePeriod') -and $null -ne $row.outsidePeriod) { $detail += (Format-MbtNumber $row.outsidePeriod) + ' outside the period' }
    $detail += (Format-MbtDuration $seconds)
    $head = if ($mode -eq 'Copy') { '[{0}/{1}] {2} ({3}) into {4} ({5})' -f $Position, $Of, $row.sourceUpn, $row.sourceLocation, $row.upn, $row.location } else { '[{0}/{1}] {2} ({3})' -f $Position, $Of, $row.upn, $row.location }
    switch ($state) {
        'complete' { Write-MbtItem OK ("$head complete: " + ($detail -join ', ')) }
        'partial' { Write-MbtItem WARN ("$head $message" + $(if ($detail.Count) { ' - ' + ($detail -join ', ') } else { '' })) }
        default { Write-MbtItem FAIL ("$head failed: $message") }
    }
    Write-MbtBatchLog -Folder $Folder -Text ("{0} {1} job {2} exit {3}{4}" -f $row.upn, $state, $job.Id, $exit, $(if ($message) { ": $message" } else { '' }))
}

function Invoke-MbtBatch {
    <#
    .SYNOPSIS
        Runs the rows of a batch that are not complete, -Parallel mailboxes at a time (1 to 4), each as an ordinary export job in a child process.
    .DESCRIPTION
        Rows run in the order of the CSV file: pending ones, ones that were interrupted (running), partial ones (the same job folder: only what is left) and failed ones (tried again). batch.json is
        saved after every change and report.html written after every mailbox. A stop request (a file named stop.request in the batch folder, or -StopFile) asks the running jobs to stop the way the console
        does (their own stop.request) and starts no other: the batch can be run again. The same happens when the command is interrupted (Ctrl+C). Sign-in is never asked for in a child: a mailbox that needs it fails
        with that reason and the batch goes on; after three in a row the others are left for the next run, since they would fail the same way.
        Returns Batch, ExitCode (0 every mailbox complete, 1 a mailbox failed, 2 not complete but nothing lost), Stopped, Counts, Report.
    .PARAMETER Launcher
        Starts a child: a script block that gets the arguments, the job and the batch, and returns an object with Process (HasExited, ExitCode) and Pipes. The default starts a pwsh (the tests give their own).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkFolder,
        [Parameter(Mandatory)]$Batch,
        [Parameter(Mandatory)][string]$EntryScript,
        [Parameter(Mandatory)][hashtable]$Option,
        [ValidateRange(1, 4)][int]$Parallel = 1,
        [string]$StopFile, [string]$LogFolder, [switch]$DevelopmentMode,
        [scriptblock]$Launcher,
        [int]$PollMilliseconds = 1000, [int]$StartGapMilliseconds = 1500, [int]$ReportBudgetSeconds = 20, [int]$StopGraceSeconds = 90
    )
    Set-StrictMode -Version Latest
    $id = [string]$Batch.id
    $folder = Get-MbtBatchFolder -WorkFolder $WorkFolder -Name $id
    [void][IO.Directory]::CreateDirectory($folder)
    $lock = $null
    try { $lock = [IO.FileStream]::new([IO.Path]::Combine($folder, 'batch.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
    catch [IO.IOException] { throw "The batch $id is already running (or stopping): wait for it, or stop it by writing a file named stop.request in $folder." }
    $ownStop = [IO.Path]::Combine($folder, 'stop.request')
    $running = [Collections.Generic.List[object]]::new()
    $stopping = $false; $signinStreak = 0; $abandoned = $false
    $of = @($Batch.rows).Count
    $mode = Get-MbtBatchMode $Batch
    $kind = if ($mode -eq 'Export') { 'export' } elseif ($mode -eq 'Copy') { 'mailcopy' } else { 'import' }
    try {
        if ([IO.File]::Exists($ownStop)) { Remove-Item -LiteralPath $ownStop -Force -ErrorAction SilentlyContinue }
        $queue = [Collections.Generic.List[object]]::new()
        foreach ($r in @($Batch.rows)) { if ([string]$r.state -ne 'complete') { $queue.Add($r) } }
        # the first row of the queue that holds nothing a running row holds (the target mailbox location, the source of an import)
        $nextRow = {
            $busy = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($item in $running) { foreach ($k in (Get-MbtBatchRowLock $item.Row)) { [void]$busy.Add($k) } }
            foreach ($candidate in $queue) {
                $clash = $false
                foreach ($k in (Get-MbtBatchRowLock $candidate)) { if ($busy.Contains($k)) { $clash = $true; break } }
                if (-not $clash) { return $candidate }
            }
            return $null
        }
        $Batch.parallel = $Parallel
        Save-MbtBatch -Batch $Batch -Folder $folder
        Write-MbtBatchLog -Folder $folder -Text ("run started: {0} of {1} mailboxes to do, {2} at a time" -f $queue.Count, $of, $Parallel)
        $lastStart = [DateTime]::MinValue
        $finishedNow = 0
        $startRow = {
            param($row)
            $job = New-MbtBatchJobFolder -WorkFolder $WorkFolder -Name ([string]$row.job) -Kind $kind
            if (Test-Path -LiteralPath $job.StopPath) { Remove-Item -LiteralPath $job.StopPath -Force }
            $argument = Get-MbtBatchChildArgument -Row $row -Option $Option -Job $job -BatchId $id -Mode $mode
            Write-MbtBatchJobRequest -Row $row -Job $job -Argument $argument -Option $Option -BatchId $id -Mode $mode
            if ($row.Contains('journal') -and [string]$row['journal']) { [void][IO.Directory]::CreateDirectory([string]$row['journal']) }
            $title = if ($mode -eq 'Export') { 'Export - ' + $row.upn } elseif ($mode -eq 'Copy') { 'Copy - ' + $row.sourceUpn + ' to ' + $row.upn } else { 'Import - ' + $row.upn }
            $sourceText = if ($mode -eq 'Export') { [string]$row.upn } elseif ($mode -eq 'Copy') { [string]$row.sourceUpn + ' (' + [string]$row.sourceLocation + ')' } else { [string]$row.source }
            $destinationText = if ($mode -ne 'Export') { [string]$row.upn + ' (' + [string]$row.location + ')' + $(if ([string]$row.targetRoot) { ', folder ' + [string]$row.targetRoot }) } elseif ([string]$row.output) { [string]$row.output } else { [string]$row.store }
            try { Write-MbtJobStatus -JobFolder $job.Folder -State running -JobId $job.Id -Task $kind -Title $title -Mode $mode -Source $sourceText -Destination $destinationText -Batch $id } catch { Write-Verbose "status: $($_.Exception.Message)" }
            $row.job = $job.Name; $row.attempts = [int]$row.attempts + 1; $row.state = 'running'; $row.started = [DateTime]::UtcNow.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z'); $row.finished = ''; $row.message = ''
            Save-MbtBatch -Batch $Batch -Folder $folder
            $handle = if ($Launcher) { & $Launcher $argument $job $Batch } else { Start-MbtBatchProcess -EntryScript $EntryScript -Argument $argument -Job $job -LogFolder $LogFolder -DevelopmentMode:$DevelopmentMode }
            Write-MbtBatchLog -Folder $folder -Text "$($row.upn) started: job $($job.Id)"
            Write-MbtItem INFO ('[{0}/{1}] {2} ({3}){4} started{5}' -f ([array]::IndexOf(@($Batch.rows), $row) + 1), $of, $(if ($mode -eq 'Copy') { $row.sourceUpn } else { $row.upn }), $(if ($mode -eq 'Copy') { $row.sourceLocation } else { $row.location }), $(if ($mode -eq 'Copy') { ' into ' + $row.upn + ' (' + $row.location + ')' } else { '' }), $(if ([int]$row.attempts -gt 1) { ' again (try ' + $row.attempts + ')' } else { '' }))
            [pscustomobject]@{ Row = $row; Job = $job; Handle = $handle; Started = [DateTime]::UtcNow; ExitCode = $null }
        }
        try {
            while ($true) {
                if (-not $stopping -and ([IO.File]::Exists($ownStop) -or ($StopFile -and [IO.File]::Exists($StopFile)))) {
                    $stopping = $true
                    Write-MbtItem WARN "stop requested: the $($running.Count) running $(if ($running.Count -eq 1) { 'mailbox is' } else { 'mailboxes are' }) asked to stop, no other starts; run the same command again to continue"
                    Write-MbtBatchLog -Folder $folder -Text 'stop requested'
                    foreach ($item in $running) { try { [IO.File]::WriteAllText($item.Job.StopPath, [DateTime]::UtcNow.ToString('o')) } catch { Write-Verbose "stop: $($_.Exception.Message)" } }
                }
                foreach ($item in @($running)) {
                    if (-not $item.Handle.Process.HasExited) { continue }
                    [void]$running.Remove($item)
                    $item.ExitCode = Get-MbtBatchExitCode -Handle $item.Handle
                    Close-MbtBatchHandle -Handle $item.Handle
                    $finishedNow++
                    Complete-MbtBatchRow -Batch $Batch -Item $item -Folder $folder -StopAsked $stopping -BatchId $id -Position ([array]::IndexOf(@($Batch.rows), $item.Row) + 1) -Of $of
                    if ([string]$item.Row.reason -eq 'signin-required') { $signinStreak++ } else { $signinStreak = 0 }
                    Save-MbtBatch -Batch $Batch -Folder $folder
                    try { $info = @{}; [void](Write-MbtBatchReport -WorkFolder $WorkFolder -Batch $Batch -Folder $folder -TimeBudgetSeconds $ReportBudgetSeconds -Info $info) } catch { Write-MbtItem WARN "the report of the batch could not be written: $($_.Exception.Message)" }
                }
                if ($signinStreak -ge 3 -and -not $abandoned -and -not $Option.AppOnly) {
                    $abandoned = $true
                    Write-MbtItem WARN 'the sign-in is gone: three mailboxes in a row needed it. No other mailbox is started; sign in again (run the command again without -NoPrompt, once) and run the batch again'
                    Write-MbtBatchLog -Folder $folder -Text 'sign-in is gone: the batch stops starting mailboxes'
                }
                if ($stopping -or $abandoned) { if ($running.Count -eq 0) { break } }
                else {
                    while ($running.Count -lt $Parallel -and $queue.Count -gt 0 -and (([DateTime]::UtcNow - $lastStart).TotalMilliseconds -ge $StartGapMilliseconds -or $running.Count -eq 0)) {
                        $row = & $nextRow
                        if ($null -eq $row) { break }
                        [void]$queue.Remove($row)
                        $lastStart = [DateTime]::UtcNow
                        try { $running.Add((& $startRow $row)) }
                        catch {
                            $row.state = 'failed'; $row.message = "the mailbox could not be started: $($_.Exception.Message)"; $row.finished = [DateTime]::UtcNow.ToString('yyyy-MM-dd\THH:mm:ss.fff\Z')
                            $startedFolder = if ($row.job) { [IO.Path]::Combine($WorkFolder, 'jobs', [string]$row.job) } else { '' }
                            if ($startedFolder -and [IO.Directory]::Exists($startedFolder)) {
                                try { Write-MbtJobStatus -JobFolder $startedFolder -State error -Task $kind -Mode $mode -Source ([string]$row.upn) -Batch $id -ExitCode 1 -Reason 'start-failed' -Next retry }
                                catch { Write-Verbose "status: $($_.Exception.Message)" }
                            }
                            Write-MbtItem FAIL ('[{0}/{1}] {2} ({3}) failed: {4}' -f ([array]::IndexOf(@($Batch.rows), $row) + 1), $of, $row.upn, $row.location, $row.message)
                            Write-MbtBatchLog -Folder $folder -Text "$($row.upn) not started: $($_.Exception.Message)"
                            Save-MbtBatch -Batch $Batch -Folder $folder
                        }
                    }
                    if ($running.Count -eq 0 -and $queue.Count -eq 0) { break }
                }
                Start-Sleep -Milliseconds $PollMilliseconds
            }
        } finally {
            # Ctrl+C or an error: the running mailboxes are asked to stop the way the console does, and given time to save what they did
            if ($running.Count) {
                foreach ($item in $running) { try { [IO.File]::WriteAllText($item.Job.StopPath, [DateTime]::UtcNow.ToString('o')) } catch { Write-Verbose "stop: $($_.Exception.Message)" } }
                $until = [DateTime]::UtcNow.AddSeconds($StopGraceSeconds)
                foreach ($item in @($running)) {
                    while (-not $item.Handle.Process.HasExited -and [DateTime]::UtcNow -lt $until) { [Threading.Thread]::Sleep(250) }
                    if (-not $item.Handle.Process.HasExited) { try { $item.Handle.Process.Kill($true) } catch { Write-Verbose "kill: $_" } }
                    try { $item.Handle.Process.WaitForExit(5000) | Out-Null } catch { Write-Verbose "wait: $_" }
                    [void]$running.Remove($item)
                    $item.ExitCode = Get-MbtBatchExitCode -Handle $item.Handle
                    Close-MbtBatchHandle -Handle $item.Handle
                    Complete-MbtBatchRow -Batch $Batch -Item $item -Folder $folder -StopAsked $true -BatchId $id -Position ([array]::IndexOf(@($Batch.rows), $item.Row) + 1) -Of $of
                }
                Save-MbtBatch -Batch $Batch -Folder $folder
            }
        }
        Save-MbtBatch -Batch $Batch -Folder $folder
        $reportPath = $null
        try { $reportPath = Write-MbtBatchReport -WorkFolder $WorkFolder -Batch $Batch -Folder $folder -TimeBudgetSeconds ($ReportBudgetSeconds * 3) } catch { Write-MbtItem WARN "the report of the batch could not be written: $($_.Exception.Message)" }
        $counts = Get-MbtBatchCount $Batch
        $exit = if ($counts.Failed -gt 0) { 1 } elseif ($counts.Complete -eq $counts.Total) { 0 } else { 2 }
        Write-MbtBatchLog -Folder $folder -Text ("run ended: {0} complete, {1} not complete, {2} failed, {3} not started; exit {4}" -f $counts.Complete, ($counts.Partial + $counts.Running), $counts.Failed, $counts.Pending, $exit)
        return [pscustomobject]@{ Batch = $Batch; ExitCode = $exit; Stopped = $stopping; Abandoned = $abandoned; Counts = $counts; Report = $reportPath; Folder = $folder; Finished = $finishedNow }
    } finally {
        if ($lock) { $lock.Dispose() }
    }
}

#endregion

#region What the command line says -------------------------------------------------------------

function Format-MbtBatchToken {
    <# A value of a command line the operator can paste: quoted when it holds a space or a character PowerShell reads. #>
    param([Parameter(Mandatory)][string]$Text)
    if ($Text -match '[\s;&()$''`"]') { return ("'" + $Text.Replace("'", "''") + "'") }
    return $Text
}

function Show-MbtBatchPlan {
    <# The mailboxes of a batch before anything starts (-WhatIf, and the start of a run): one line each with where its files go, or for an import where it comes from and goes to; the first lines only when there are many. #>
    param([Parameter(Mandatory)]$Batch, [int]$Limit = 25)
    $rows = @($Batch.rows)
    $shown = 0
    foreach ($r in $rows) {
        if ($shown -ge $Limit) { break }
        $shown++
        $isCopy = $r.Contains('sourceUpn') -and [string]$r['sourceUpn']
        $isImport = $r.Contains('source') -and [string]$r['source']
        if ($isCopy) {
            $state = if ([string]$r.state -in '', 'pending') { '' } else { " [$($r.state)]" }
            Write-MbtDetail ('{0,4}  {1} ({2}) into {3} ({4}) - {5}{6} - store {7}{8}' -f $r.n, $r.sourceUpn, $r.sourceLocation, $r.upn, $r.location, $(if ([string]$r.selection) { 'selection ' + $r.selection } else { 'the whole mailbox' }), $(if ([string]$r.targetRoot) { ', folder ' + [string]$r.targetRoot } else { ', the folder names of the source' }), $r.store, $state)
            continue
        }
        $where = if ($isImport) { 'from ' + [string]$r['source'] + $(if ([string]$r['sourceMailbox']) { ' (' + [string]$r['sourceMailbox'] + ')' }) + $(if ([string]$r['targetRoot']) { ' into the folder ' + [string]$r['targetRoot'] } else { ' with the folder names of the source' }) }
        elseif ([string]$r.output) { "PST $($r.output)" } else { "store $($r.store)" }
        $what = if ($isImport) { [string]$r.location } elseif ([string]$r.selection) { "selection $($r.selection)" } else { [string]$r.location }
        $state = if ([string]$r.state -in '', 'pending') { '' } else { " [$($r.state)]" }
        Write-MbtDetail ('{0,4}  {1} - {2} - {3}{4}' -f $r.n, $r.upn, $what, $where, $state)
    }
    if ($rows.Count -gt $shown) { Write-MbtDetail ('and {0} more mailbox{1}: the list is in batch.json of the batch folder' -f (Format-MbtNumber ($rows.Count - $shown)), $(if ($rows.Count - $shown -eq 1) { '' } else { 'es' })) }
}

function Write-MbtBatchSummary {
    <#
    .SYNOPSIS
        The card that ends a batch run, in the house format: result and exit code, the numbers, the batch folder, the report, the log, what to do next.
    .PARAMETER Failure
        Sentences for a batch that did not start (the CSV file, the name of the batch): they go in the card as the error.
    .PARAMETER Command
        The command line that runs the batch again.
    #>
    param(
        [Parameter(Mandatory)][int]$ExitCode, $Batch, [string]$Folder, [double]$Seconds, [string]$Report, [string]$AllJobs, [string]$Log,
        [string[]]$Failure = @(), [string]$Command, [switch]$Stopped, [switch]$WhatIf, [string]$Mode = 'Export'
    )
    $word = if ($Mode -eq 'Export') { 'Export' } elseif ($Mode -eq 'Copy') { 'Copy' } else { 'Import' }
    $meaning = switch ($ExitCode) {
        0 { 'complete: every mailbox is complete' }
        1 { 'failed: the CSV file or a mailbox needs a look, see above' }
        2 { 'not complete: nothing was lost, run the same command again to continue' }
        default { 'see the messages above' }
    }
    $status = switch ($ExitCode) { 0 { 'Ok' } 1 { 'Fail' } default { 'Warn' } }
    $title = if ($WhatIf) { "$word batch planned (WhatIf, nothing was started)" }
        elseif ($Failure.Count) { "$word batch not started" }
        elseif ($ExitCode -eq 0) { "$word batch finished" }
        elseif ($Stopped) { "$word batch stopped before the end" }
        elseif ($ExitCode -eq 1) { "$word batch finished with failures" }
        else { "$word batch stopped before the end" }
    if ($WhatIf) { $meaning = 'planned: the file is valid and nothing was started' }
    $values = [ordered]@{ Result = @($(if ($ExitCode -eq 0) { 'Ok' } elseif ($ExitCode -eq 1) { 'Fail' } else { 'Warn' }), "exit code $ExitCode - $meaning") }
    if ($Failure.Count) { $values.Error = @('Fail', $Failure) }
    if ($Batch) {
        $c = Get-MbtBatchCount $Batch
        $values.Mailboxes = @('Store', ('{0} in {1}: {2} complete, {3} not complete, {4} failed, {5} not started' -f (Format-MbtNumber $c.Total), [string]$Batch.id, (Format-MbtNumber $c.Complete), (Format-MbtNumber ($c.Partial + $c.Running)), (Format-MbtNumber $c.Failed), (Format-MbtNumber $c.Pending)))
        if ($Mode -eq 'Copy') {
            if ($c.Done -gt 0 -or $c.Exported -gt 0 -or $c.FailedItems -gt 0 -or $c.Unknown -gt 0) {
                $parts = @('{0} read, {1} copied' -f (Format-MbtNumber $c.Exported), (Format-MbtNumber $c.Done))
                if ($c.FailedItems -gt 0) { $parts += '{0} failed (tried again at the next run)' -f (Format-MbtNumber $c.FailedItems) }
                $parts += '{0} uncertain' -f (Format-MbtNumber $c.Unknown)
                $values.Items = @('Store', ($parts -join ', '))
            }
        }
        elseif ($Mode -ne 'Export') {
            if ($c.Done -gt 0 -or $c.FailedItems -gt 0 -or $c.Refused -gt 0 -or $c.Unknown -gt 0) {
                $parts = @('{0} imported' -f (Format-MbtNumber $c.Done))
                if ($c.Refused -gt 0) { $parts += '{0} refused for good' -f (Format-MbtNumber $c.Refused) }
                if ($c.FailedItems -gt 0) { $parts += '{0} failed (tried again at the next run)' -f (Format-MbtNumber $c.FailedItems) }
                $parts += '{0} uncertain' -f (Format-MbtNumber $c.Unknown)
                $values.Items = @('Store', ($parts -join ', '))
            }
        }
        elseif ($c.Done -gt 0 -or $c.FailedItems -gt 0) {
            $outside = if ($c.OutsidePeriod -gt 0) { ", $(Format-MbtNumber $c.OutsidePeriod) outside the period (skipped)" } else { '' }
            $values.Items = @('Store', ('{0} exported, {1} failed, {2}{3}' -f (Format-MbtNumber $c.Done), (Format-MbtNumber $c.FailedItems), (Format-MbtBytes ([double]$c.Bytes)), $outside))
        }
    }
    if ($Seconds -gt 0) { $values.Duration = @('Clock', (Format-MbtDuration $Seconds)) }
    if ($Folder) { $values.Batch = @('Folder', $Folder) }
    if ($Report) { $values.Report = @('Report', $Report) }
    if ($AllJobs) { $values['All jobs'] = @('Report', $AllJobs) }
    if ($Log) { $values.Log = @('Log', $Log) }
    $next = [Collections.Generic.List[string]]::new()
    if ($WhatIf) { $next.Add('Run the same command without -WhatIf to start the batch.') }
    elseif ($Failure.Count) { $next.Add('Correct what the lines above say, then run the same command again. Nothing was started.') }
    elseif ($ExitCode -eq 0 -and $Mode -eq 'Copy') {
        $next.Add('Open the report: one line for each copy, with its items read and copied and the report of its job. The export store of each source stays in the batch store folder: it is the proof of what was read.')
        $next.Add('Running this batch again adds nothing: a source that was read is not read again, and an item that was copied is never created twice.')
    }
    elseif ($ExitCode -eq 0 -and $Mode -ne 'Export') {
        $next.Add('Open the report: one line for each mailbox, with its items imported and the report of its job.')
        $next.Add('Running this batch again adds nothing: an item that was imported is never created twice.')
    }
    elseif ($ExitCode -eq 0) {
        $next.Add('Open the report: one line for each mailbox, with its items, its data and the report of its job.')
        $next.Add('A mailbox that must be exported again needs a new batch name: this one is complete.')
    } else {
        $next.Add('Run the same command again to continue: complete mailboxes are skipped, the others do only what is left, failed ones are tried again.')
        if ($Command) { $next.Add($Command) }
        $next.Add('Open the report to see each mailbox; the reason of a failed one is on its line above and in batch.log.')
    }
    $values['Next steps'] = @('Plan', $next.ToArray())
    Write-MbtSummary -Title $title -Values $values -Status $status
}

#endregion

# SIG # Begin signature block
# MIImKwYJKoZIhvcNAQcCoIImHDCCJhgCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDJytpfKAeC3Ux9
# 65/M76853OFpuAOG0apKXH7X1GUg96CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYEjCCGA4CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBfrt/KeV4b
# ANiQPOScf4KoP3QnHPKSZ/p/z/XVAjYtnzANBgkqhkiG9w0BAQEFAASCAgBjv8RP
# agATS1YDHjbBDbAqrZ8S1wyyVJCRsV9CucfyTOjjvX8imOWgs/dGkbIQFI+exNo2
# G2ga0r6cHFXPsKcl8+dPhNIXySts1Eb0r0M3FQ94lw1juobYuRMZDgrfGXWf7Aup
# irpLS566OXEqaNZ0dsLxczA6dDRlAerXVx4xPJP0CMZyMmd90iE3Q67Dr2KeDiek
# NbjQn3opedWsgJzE2+937KFLeEgdehBC0OQCTnaeu65M/k+pMK8XgvR8PqCPhtJX
# Kc1Pj6MDkY8uNcJcQBBXhfnFkS8xvHdP1FoH0gCDY3/1KSvZtZBaPoBNysWpxif9
# u+RMZrLGbo63dmknvrPt+gkL1MdDuh49L8+iK0BkQ7Rjp1thTmKAi1fqDMLJ6pDB
# LF6D42SHlk3nwBCDU3bGix6XONqfss5Xo4KKu4yGwSHLH2b43DEO6DbInwVH1kJ1
# EbucMI8DVe707WefMlzqoWC/BBNmY8ClX+qy+JMF88mrxkckhxCwcAxmpQF7rYdc
# Eo+D5LFimODPC/+Y964I0FpjtoWPBc6+Z9HkZZeVtJSPR5DVD3k/Rb8OPKZqX+4m
# 33kDjFCbBkh5t/TY8Kjt4V7prb+RrpuOG8Y++cYyr8TVb0koGFL5nQIQnGoqa3xm
# 0kuI2A4CW91EpvWN0h/tT5ykpuATwY2+3naTTaGCFM8wghTLBgorBgEEAYI3AwMB
# MYIUuzCCFLcGCSqGSIb3DQEHAqCCFKgwghSkAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIDXCqoZMpo5l4lSDEcEtaX6xGUzx0/0PhHuCBW/88/j7AghK4dBruGwi
# lhgPMjAyNjExMjcwNTI3MDhaMAMCAQECEFtfIa6UR4zQaZVE/+x0o76gghGnMIIE
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
# wzSPOzGCAlcwggJTAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTEyNzA1MjcwOFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIFVo
# Gr+sL+H0tYO8ykO8gWWcGSKLBw/GdpiOlRV119gPMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIESDBGAiEAtKSJnsIqKB796krAnRT55J4wHYrnbdTehd1uD8x4IccCIQDW
# raK4Q8YSFXfALUMUVGVIyqTYKn9ca1LxMpqC5JWlaA==
# SIG # End signature block
