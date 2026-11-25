<#
.SYNOPSIS
    Mailbox Transfer - the job folders of a work folder: what a job says about itself in plain words, finding, deleting and resuming a job, the job folder of a command-line run
    (dot-sourced by MailboxTransfer.psm1).

.DESCRIPTION
        Get-MbtDefaultWorkFolder                       the work folder when none is given (%LOCALAPPDATA%\MailboxTransfer, or MAILBOXTRANSFER_WORK_FOLDER)
        Get-MbtJobDescription                          why a job stopped, when, and the step it reached, as two plain sentences
        Find-MbtJob                                    a job by the name of its folder or its job id, whatever the case
        Get-MbtJobRemovalPlan / Remove-MbtJob          what deleting a job deletes and refuses; the deletion itself (never a link followed, never the data of the user)
        Get-MbtJobResumeText                           what a job that is continued does again and does not do again, for its mode
        New-MbtCliJob / Read-MbtJobRequest             the job folder of a single job started from the command line (events, status, request.json with the command and its folder)
        Get-MbtCliJobParameter                         the recorded command of such a job, ready to run again

.NOTES
    Author  : Nicolas Fabert
#>

#region Work folder and small helpers ----------------------------------------------------------------------------

function Get-MbtDefaultWorkFolder {
    <# The folder that holds jobs, batches, stores and logs when none is given: the environment variable MAILBOXTRANSFER_WORK_FOLDER, else %LOCALAPPDATA%\MailboxTransfer. #>
    [CmdletBinding()]
    param()
    if ($env:MAILBOXTRANSFER_WORK_FOLDER) { return [IO.Path]::GetFullPath($env:MAILBOXTRANSFER_WORK_FOLDER) }
    return (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'MailboxTransfer')
}

function Get-MbtJobField {
    <# A field of a status object, a request object or a hashtable, $Default when it is not there (strict mode safe). #>
    param($Object, [string]$Name, $Default = $null)
    if ($null -eq $Object) { return $Default }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name) -and $null -ne $Object[$Name]) { return $Object[$Name] } else { return $Default } }
    $p = $Object.PSObject.Properties[$Name]
    if ($p -and $null -ne $p.Value) { return $p.Value }
    return $Default
}

function Format-MbtJobTime {
    <# A time as the local date and hour, the same whatever the language of Windows (2026-11-24 22:41). #>
    param($Value)
    if ($null -eq $Value) { return '' }
    $d = if ($Value -is [datetime]) { $Value } else { try { [datetime]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal) } catch { return '' } }
    return $d.ToLocalTime().ToString('yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture)
}

function Format-MbtJobCount {
    param($Value)
    return ([long]$Value).ToString('N0', [Globalization.CultureInfo]::InvariantCulture)
}

#endregion

#region What a job says about itself ------------------------------------------------------------------------------

function Get-MbtJobDescription {
    <#
    .SYNOPSIS
        Two plain sentences about a job read by Get-MbtJob: why it is not running (the PC was shut down or the tool closed, you stopped it, an error, a decision is waited for) and when, and
        the step it reached with the counts. A status.json written by an older version has fewer fields: the sentence then says less.
    .DESCRIPTION
        Returns Kind (running, interrupted, stopped, signin, partial, error, decision, complete, other), Why and Progress (empty when nothing is known).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Job)
    $state = [string](Get-MbtJobField $Job 'State' '')
    $effective = [string](Get-MbtJobField $Job 'EffectiveState' $state)
    $reason = [string](Get-MbtJobField $Job 'Reason' '')
    $next = [string](Get-MbtJobField $Job 'Next' '')
    $when = Format-MbtJobTime (Get-MbtJobField $Job 'Updated')
    $at = if ($when) { " on $when" } else { '' }
    $counts = Get-MbtJobField $Job 'Counts'
    $step = Get-MbtJobField $Job 'Step'
    $kind = 'other'
    $why = ''
    if ([bool](Get-MbtJobField $Job 'Alive' $false)) { $kind = 'running'; $why = 'Running now, in another window or process.' }
    elseif ($state -eq 'running') { $kind = 'interrupted'; $why = "Interrupted${at}: the PC was shut down or restarted, or the tool was closed, while it was running. What was done is kept." }
    elseif ($state -eq 'stopped' -or ($effective -eq 'resumable' -and $reason -eq 'stopped')) { $kind = 'stopped'; $why = "Stopped by you$at (Stop and keep progress). Nothing is lost." }
    elseif ($effective -eq 'resumable' -and ($reason -eq 'signin-required' -or $next -eq 'signin')) { $kind = 'signin'; $why = "Stopped$at because the sign-in expired. Sign in again, then continue. Nothing is lost." }
    elseif ($effective -eq 'resumable') { $kind = 'partial'; $why = "Not finished${at}: what is left is done when you continue. Nothing is lost." }
    elseif ($effective -eq 'error') {
        $kind = 'error'
        $failure = Get-MbtJobField $Job 'Failure'
        $code = [string](Get-MbtJobField $failure 'code' '')
        $message = [string](Get-MbtJobField $failure 'message' '')
        $why = if ($message -and $code -and $code -ne 'error') { "Stopped by an error${at}: $message" } else { "Stopped by an error$at. What was done is kept." }
    }
    elseif ($effective -eq 'needs-decision') {
        $kind = 'decision'
        $n = if ($counts) { [long](Get-MbtJobField $counts 'unknown' 0) } else { 0 }
        $why = if ($n -gt 1) { "Almost done${at}: $(Format-MbtJobCount $n) items need your decision." } else { "Almost done${at}: an item needs your decision." }
    }
    elseif ($effective -eq 'complete') {
        $kind = 'complete'
        $why = if ($reason -eq 'refused-items') { "Finished$at. Some items were refused for good: the report lists them." } else { "Finished$at." }
    }
    elseif ($effective) { $why = "State: $effective$at." }

    $progress = ''
    $id = [string](Get-MbtJobField $step 'id' '')
    $title = [string](Get-MbtJobField $step 'title' '')
    if ($title) {
        $index = [int](Get-MbtJobField $step 'index' 0); $of = [int](Get-MbtJobField $step 'of' 0)
        $progress = if ($index -gt 0 -and $of -gt 0) { "Step $index of $of, $title" } else { "Step: $title" }
        $isCopy = ([string](Get-MbtJobField $Job 'Mode' '') -eq 'Copy') -or ([string](Get-MbtJobField $Job 'Task' '') -eq 'mailcopy')
        if ($isCopy -and $id) {
            if ($id -in 'tree', 'list', 'download', 'verify') { $progress += ' (reading the source mailbox)' }
            elseif ($id -in 'check', 'folders', 'send', 'reconcile', 'final') { $progress += ' (adding to the target mailbox)' }
        }
        $done = Get-MbtJobField $step 'done'; $total = Get-MbtJobField $step 'total'
        if ($null -ne $done -and $null -ne $total -and [long]$total -gt 0) { $progress += ", $(Format-MbtJobCount $done) of $(Format-MbtJobCount $total) done" }
        elseif ($null -ne $done -and [long]$done -gt 0) { $progress += ", $(Format-MbtJobCount $done) done" }
    }
    elseif ($counts) {
        $done = Get-MbtJobField $counts 'done'; $total = Get-MbtJobField $counts 'total'
        if ($null -ne $done -and $null -ne $total -and [long]$total -gt 0) { $progress = "$(Format-MbtJobCount $done) of $(Format-MbtJobCount $total) items done" }
        elseif ($null -ne $done -and [long]$done -gt 0) { $progress = "$(Format-MbtJobCount $done) items done" }
    }
    [pscustomobject]@{ Kind = $kind; Why = $why; Progress = $progress }
}

function Get-MbtJobListing {
    <#
    .SYNOPSIS
        The jobs of a work folder for a list: the ones that are not finished first (newest first), then the finished ones (newest first, -FinishedLimit of them), each with its plain words.
    .DESCRIPTION
        Returns Unfinished and Finished (the objects of Get-MbtJob with Why, Progress and Kind added) and FinishedTotal. -ExcludeFolder: names of job folders that are left out (the jobs of a
        batch that is complete). Every job that is not finished is listed; only the finished ones are cut.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkFolder, [int]$FinishedLimit = 50, [string[]]$ExcludeFolder)
    $found = @(Get-MbtJob -WorkFolder $WorkFolder -Limit 400 -ExcludeFolder $ExcludeFolder)
    $unfinished = [Collections.Generic.List[object]]::new(); $finished = [Collections.Generic.List[object]]::new()
    foreach ($j in $found) {
        # the check of a plan (a read-only run of the console) is not a job to continue: its folder is removed after a day
        if ($j.Name -match '-plan-[0-9a-fA-F]{4}$') { continue }
        $d = Get-MbtJobDescription -Job $j
        $line = $j | Select-Object *, @{ n = 'Why'; e = { $d.Why } }, @{ n = 'Progress'; e = { $d.Progress } }, @{ n = 'Kind'; e = { $d.Kind } }
        if ([string]$j.EffectiveState -eq 'complete') { $finished.Add($line) } else { $unfinished.Add($line) }
    }
    [pscustomobject]@{ Unfinished = @($unfinished); Finished = @($finished | Select-Object -First $FinishedLimit); FinishedTotal = $finished.Count }
}

#endregion

#region Find and delete a job -------------------------------------------------------------------------------------

function Find-MbtJob {
    <# A job of a work folder by the name of its folder or by its job id, whatever the case (never a path). Returns the object of Get-MbtJob, or $null. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)][AllowEmptyString()][string]$Name)
    $n = $Name.Trim()
    if (-not $n -or $n -eq '.' -or $n -eq '..' -or $n.IndexOfAny([char[]]'\/:*?"<>|') -ge 0) { return $null }
    $byFolder = @(Get-MbtJob -WorkFolder $WorkFolder -Name $n)
    if ($byFolder.Count) { return $byFolder[0] }
    foreach ($j in @(Get-MbtJob -WorkFolder $WorkFolder -Limit 1000)) { if ($j.JobId -and ([string]$j.JobId).Equals($n, [StringComparison]::OrdinalIgnoreCase)) { return $j } }
    return $null
}

function Get-MbtFolderBytes {
    <# The size of the files below a folder; a link (junction or symbolic link) is never followed. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)
    $total = 0L
    $pending = [Collections.Generic.Stack[IO.DirectoryInfo]]::new()
    $pending.Push([IO.DirectoryInfo]::new($Path))
    while ($pending.Count) {
        $dir = $pending.Pop()
        try {
            foreach ($entry in $dir.EnumerateFileSystemInfos()) {
                if (Test-MbtFolderLink $entry) { continue }
                if ($entry -is [IO.DirectoryInfo]) { $pending.Push($entry) } else { $total += ([IO.FileInfo]$entry).Length }
            }
        } catch { Write-Verbose "size of $($dir.FullName): $($_.Exception.Message)" }
    }
    return $total
}

function Test-MbtPathInside {
    <# True when -Path is below -Folder (the full paths, not case sensitive). #>
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Folder)
    $root = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($Folder)) + [IO.Path]::DirectorySeparatorChar
    return ([IO.Path]::GetFullPath($Path)).StartsWith($root, [StringComparison]::OrdinalIgnoreCase)
}

function Get-MbtJobRemovalPlan {
    <#
    .SYNOPSIS
        What deleting a job would delete, and why it would be refused.
    .DESCRIPTION
        Deleted: the folder of the job (its events, status, request, log, report, and what the job kept in its own folder). Never deleted: what the user chose (an export store, PST files, source
        files, the target mailbox). Refused: a job that is running (Alive), a job of a batch that is not complete, a job whose folder holds the store of the user or is named by another job,
        a folder that is a link. A copy whose store is in the stores folder of the work folder can delete it as well (-IncludeStore), when the copy is complete and no other job that is not
        finished uses the store. Returns Allowed, Reason, Folder, Name, Title, Source, Destination, Updated, Unfinished, Warnings and Store (for a copy: Path, Offered, Allowed, Reason, Bytes).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)]$Job, [switch]$WithSize)
    $name = [string]$Job.Name
    $folder = [string]$Job.Folder
    $unfinished = [string]$Job.EffectiveState -ne 'complete'
    $plan = [pscustomobject]@{
        Name = $name; Folder = $folder; Title = $(if ($Job.Title) { [string]$Job.Title } else { [string]$Job.JobId }); Source = [string]$Job.Source; Destination = [string]$Job.Destination
        Started = $Job.Started; Updated = $Job.Updated; State = [string]$Job.EffectiveState; Unfinished = $unfinished
        Allowed = $true; Reason = ''; Warnings = @(); Store = $null
    }
    $refuse = { param([string]$why) $plan.Allowed = $false; $plan.Reason = $why; return $plan }

    $jobsFolder = [IO.Path]::GetFullPath((Join-Path $WorkFolder 'jobs'))
    $fullFolder = [IO.Path]::GetFullPath($folder)
    if (-not [string]::Equals([IO.Path]::GetDirectoryName($fullFolder), $jobsFolder.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) { return (& $refuse "The folder of the job '$name' is not directly below the jobs folder of the work folder: it is left alone.") }
    if (Test-MbtFolderLink ([IO.DirectoryInfo]::new($fullFolder))) { return (& $refuse "The folder of the job '$name' is a link to another folder: it is left alone. Remove the link by hand.") }
    if ([bool]$Job.Alive) { return (& $refuse "The job '$name' is running now, in another window or process. Stop it first (Stop and keep progress), then delete it.") }
    $held = Get-MbtBatchProtection -WorkFolder $WorkFolder
    $batch = [string]$Job.Batch
    if ($held.Jobs.Contains($name) -or ($batch -and $held.Batches.Contains($batch))) {
        $batchName = if ($batch) { $batch } else { 'a batch' }
        return (& $refuse "The job '$name' belongs to the batch $batchName, which is not complete. It is kept so that the batch can be continued (-ResumeBatch). Finish the batch first.")
    }
    # a job that names a folder of this job (a store, an output) needs it: the other jobs are read for their request
    $path = $fullFolder.TrimEnd('\')
    $escaped = $path.Replace('\', '\\')
    foreach ($other in [IO.Directory]::EnumerateDirectories($jobsFolder)) {
        if ([string]::Equals($other, $fullFolder.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) { continue }
        $request = Join-Path $other 'request.json'
        if (-not [IO.File]::Exists($request)) { continue }
        $text = $null; try { $text = [IO.File]::ReadAllText($request) } catch { continue }
        if ($text.IndexOf($escaped + '\\', [StringComparison]::OrdinalIgnoreCase) -ge 0 -or $text.IndexOf($path + '\', [StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return (& $refuse "The job '$name' keeps files that the job '$([IO.Path]::GetFileName($other))' uses: it is not deleted.")
        }
    }
    # the store of the user inside the job folder would be deleted with it
    $request = Read-MbtJobRequest -Folder $fullFolder
    $storeOfJob = [string](Get-MbtJobField $request 'StorePath' ([string]$Job.Store))
    if (-not $storeOfJob) { $storeOfJob = [string]$Job.Store }
    $isCopy = ([string]$Job.Mode -eq 'Copy') -or ([string]$Job.Task -eq 'mailcopy')
    if ($storeOfJob -and -not $isCopy -and (Test-MbtPathInside -Path $storeOfJob -Folder $fullFolder)) {
        return (& $refuse "The export store of the job '$name' ($storeOfJob) is inside its folder: deleting the job would delete the store. Move the store out of it first.")
    }
    $warnings = [Collections.Generic.List[string]]::new()
    if ($unfinished) { $warnings.Add('This job is not finished. After it is deleted you can no longer continue it from the console or with -ResumeJob.') }
    $journal = $null
    foreach ($candidate in @($fullFolder, (Join-Path $fullFolder 'work'), (Join-Path $fullFolder 'temp'))) { $f = Join-Path $candidate 'pstimport.db'; if ([IO.File]::Exists($f)) { $journal = $f; break } }
    if ($journal) { $warnings.Add('The record of what was already imported from the PST file is in this folder and is deleted with it. Importing the same PST file again later would add its messages a second time.') }
    $plan.Warnings = @($warnings)

    if ($isCopy -and $storeOfJob) {
        $store = [pscustomobject]@{ Path = $storeOfJob; Offered = $false; Allowed = $false; Reason = ''; Bytes = $null }
        $storesFolder = Join-Path $WorkFolder 'stores'
        if (-not (Test-MbtPathInside -Path $storeOfJob -Folder $storesFolder)) { $store.Reason = 'The copy store is outside the stores folder of the work folder: this tool never deletes it.' }
        elseif (-not [IO.Directory]::Exists($storeOfJob)) { $store.Reason = 'The copy store is not there any more.' }
        elseif (Test-MbtFolderLink ([IO.DirectoryInfo]::new($storeOfJob))) { $store.Reason = 'The copy store is a link to another folder: it is left alone.' }
        else {
            $store.Offered = $true
            if ($unfinished) { $store.Reason = 'The copy is not complete: its second part, adding the items to the target mailbox, is not finished, and it reads this store. Finish the copy first.' }
            else {
                $user = $null
                foreach ($j in @(Get-MbtJob -WorkFolder $WorkFolder -Limit 1000)) {
                    if ($j.Name -ieq $name -or [string]$j.EffectiveState -eq 'complete' -or -not $j.Store) { continue }
                    if (([string]$j.Store).TrimEnd('\').Equals($storeOfJob.TrimEnd('\'), [StringComparison]::OrdinalIgnoreCase)) { $user = $j; break }
                }
                if ($user) { $store.Reason = "The copy store is used by the job '$($user.Name)', which is not finished." } else { $store.Allowed = $true }
            }
            if ($WithSize) { $store.Bytes = Get-MbtFolderBytes -Path $storeOfJob }
        }
        $plan.Store = $store
    }
    return $plan
}

function Remove-MbtJob {
    <#
    .SYNOPSIS
        Deletes the folder of a job (and, with -IncludeStore, the copy store of a complete copy that lives in the stores folder of the work folder). Asks first; -WhatIf shows what would go.
    .DESCRIPTION
        Only the folder of the job is deleted: the export store the user chose, PST files, source files and mailboxes are never touched. A link inside the folder is removed and never followed.
        The overview of all the jobs is written again. Returns Name, Folder, Removed, StoreRemoved and StoreBytes.
    #>
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param([Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)][string]$Name, [switch]$IncludeStore)
    $job = Find-MbtJob -WorkFolder $WorkFolder -Name $Name
    if (-not $job) { throw "No job is named '$Name' in $(Join-Path $WorkFolder 'jobs'). -Mode Status lists the jobs." }
    $plan = Get-MbtJobRemovalPlan -WorkFolder $WorkFolder -Job $job -WithSize:$IncludeStore
    if (-not $plan.Allowed) { throw $plan.Reason }
    if ($IncludeStore) {
        if (-not $plan.Store) { throw "-RemoveCopyStore: the job '$($plan.Name)' is not a copy." }
        if (-not $plan.Store.Allowed) { throw "-RemoveCopyStore: $($plan.Store.Reason)" }
    }
    $result = [pscustomobject]@{ Name = $plan.Name; Folder = $plan.Folder; Removed = $false; StoreRemoved = $false; StoreBytes = $null }
    if ($IncludeStore -and $PSCmdlet.ShouldProcess($plan.Store.Path, 'Delete the copy store')) {
        $store = $plan.Store.Path
        if (-not (Test-MbtFolderLink ([IO.DirectoryInfo]::new($store)))) { Remove-MbtFolderTree -Path $store }
        [IO.Directory]::Delete($store)
        $result.StoreRemoved = $true; $result.StoreBytes = $plan.Store.Bytes
        Write-MbtLog INFO "Copy store removed with its job: $store"
    }
    if ($PSCmdlet.ShouldProcess($plan.Folder, 'Delete the job folder')) {
        # status.json goes last: a folder that could not be removed whole still says what the job was
        Remove-MbtFolderTree -Path $plan.Folder -KeepFirst 'status.json'
        [IO.File]::Delete([IO.Path]::Combine($plan.Folder, 'status.json'))
        [IO.Directory]::Delete($plan.Folder)
        $result.Removed = $true
        Write-MbtLog INFO "Job folder removed by request: $($plan.Folder)"
        if (-not $WhatIfPreference) {
            try { $null = Write-MbtJobOverview -WorkFolder $WorkFolder -TimeBudgetSeconds 10 } catch { Write-MbtLog WARN "The overview of all the jobs could not be written: $($_.Exception.Message)" }
        }
    }
    return $result
}

#endregion

#region What a continued job does ----------------------------------------------------------------------------------

function Get-MbtJobResumeText {
    <#
    .SYNOPSIS
        What happens when a job is continued, for its mode, in plain words: what is already done and is not done again, what is checked again, and where it restarts.
    .DESCRIPTION
        Each sentence says what the engine does when the same command runs again (the catalog of the export store, the manifest of the PST files, the import record of the store or the journal of a
        PST, import-state.jsonl of a folder of .eml files, the copy that goes straight to its second part). -Request is the request.json of the job (it adds the PST output of an export and the
        number of attempts). Returns Mode, Heading, Already, Again, From and Notes (lists of sentences).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Job, $Request)
    $mode = [string](Get-MbtJobField $Job 'Mode' '')
    $task = [string](Get-MbtJobField $Job 'Task' '')
    if (-not $mode) { $mode = [string](Get-MbtJobField $Request 'Mode' '') }
    if (-not $mode) {
        $mode = switch ($task) { 'mailcopy' { 'Copy' } 'export' { 'Export' } 'exportpst' { 'ExportPst' } 'import' { 'Import' } default { '' } }
        if ($mode -eq 'Import') { $src = [string](Get-MbtJobField $Request 'ImSrc' 'store'); $mode = switch ($src) { 'pst' { 'ImportPst' } 'eml' { 'ImportEml' } default { 'Import' } } }
    }
    $step = Get-MbtJobField $Job 'Step'
    $stepId = [string](Get-MbtJobField $step 'id' '')
    $stepTitle = [string](Get-MbtJobField $step 'title' '')
    $stepIndex = [int](Get-MbtJobField $step 'index' 0); $stepOf = [int](Get-MbtJobField $step 'of' 0)
    $reached = if ($stepTitle -and $stepIndex -gt 0 -and $stepOf -gt 0) { " Before it stopped it had reached step $stepIndex of $stepOf, $stepTitle." } else { '' }
    $tries = [int](Get-MbtJobField $Request 'Tries' (Get-MbtJobField (Get-MbtJobField $Request 'Parameters') 'MaxAttempts' 3))
    if ($tries -lt 1) { $tries = 3 }
    $kind = (Get-MbtJobDescription -Job $Job).Kind
    $already = [Collections.Generic.List[string]]::new(); $again = [Collections.Generic.List[string]]::new(); $notes = [Collections.Generic.List[string]]::new()
    $heading = ''; $from = ''
    $uncertain = 'Items that were being sent when it stopped may or may not have reached the mailbox. They are looked up in the target folder first, after a short wait so that Microsoft can finish creating them: an item found there counts as done, an item that is not found is sent again, and when several items could be it you are asked.'
    $sendAgain = "Items that failed are tried again, up to $tries tries in all."
    switch ($mode) {
        'Export' {
            $heading = 'Export to an export store'
            $already.Add('The items already downloaded and checked in the export store are not downloaded again.')
            $again.Add('The folder tree and the list of the items are read again from the mailbox, so that what arrived or changed since is taken. An item that changed since is downloaded again.')
            $again.Add("Items that failed are tried again, up to $tries tries in all.")
            $pst = ([string](Get-MbtJobField $Request 'Dest' '') -eq 'pst') -or [bool](Get-MbtJobField $Request 'OutputPst')
            if ($pst) { $again.Add('Then the PST files: the parts already written and checked are kept; the part that was being written when it stopped is written again from its start.') }
            $from = "It restarts from step 1, Read the folder tree, and downloads only what is left.$reached"
        }
        'ExportPst' {
            $heading = 'Write PST files'
            $already.Add('The PST parts that were completed are kept: each one is checked again (file present, the size the manifest says, valid PST header) and is not written again.')
            $again.Add('The part that was being written when it stopped is deleted and written again from its start.')
            $again.Add('The source (the export store or the PST file) is only read.')
            $from = "It restarts at the first step, Write the PST files, after the last part that is complete.$reached"
        }
        'Import' {
            $heading = 'Import from an export store'
            $already.Add('The items already added to the mailbox (recorded as done for this job in the export store) are not sent again.')
            $again.Add('The export store is checked again, and so are the target folders: a missing folder is created, an existing one is used as it is.')
            $again.Add($uncertain)
            $again.Add($sendAgain)
            $from = "It restarts from step 1, Verify the export store, then creates the missing folders and sends only the items that are left.$reached"
        }
        'ImportPst' {
            $heading = 'Import from a PST file'
            $journal = [string](Get-MbtJobField $Request 'JournalPath' '')
            $already.Add($(if ($journal) { "The messages already added to the mailbox, as the journal of the import records them ($journal), are not added again, even when the same PST file is given again." } else { 'The messages already added to the mailbox, as the journal of the import records them, are not added again, even when the same PST file is given again.' }))
            $again.Add('The PST files are checked again.')
            $again.Add($uncertain.Replace('Items', 'Messages').Replace('an item', 'a message').Replace('items', 'messages'))
            $again.Add("Messages that failed are tried again, up to $tries tries in all. A message that Microsoft refused for good is not tried again.")
            $from = "It restarts from step 1, Check the PST files, then sends only the messages that are left.$reached"
        }
        'ImportEml' {
            $heading = 'Import .eml files'
            $already.Add('The files recorded as done in import-state.jsonl, in the folder of the .eml files, are not imported again.')
            $again.Add('The folder is read again to list the files.')
            $again.Add('A file that was being sent when it stopped, or whose answer was lost, is not sent again by itself: it may exist in the mailbox. It is listed as uncertain and you settle it after a look in the mailbox (-ResolveEml Imported or NotImported).')
            $again.Add('A file that failed is tried again.')
            $from = "It restarts at its only step, Import the messages: the files already done are skipped.$reached"
        }
        'ExportEml' {
            $heading = 'Write .eml files'
            $already.Add('The files already written, listed as exported in the manifest of the folder and still in their place, are not written again.')
            $again.Add('The items are listed again. An item whose file is missing, or that the manifest does not list, is written.')
            $from = "It restarts at its only step, Write the .eml files.$reached"
        }
        'Copy' {
            $heading = 'Copy to another mailbox'
            $notes.Add('A copy has two parts: reading the source mailbox into the store on this PC, then adding the items to the target mailbox. The source mailbox is only read.')
            $reading = $stepId -in 'tree', 'list', 'download', 'verify'
            $adding = $stepId -in 'check', 'folders', 'send', 'reconcile', 'final'
            if ($adding) {
                $already.Add('The first part is complete: the source mailbox is not read again, and the store is used as it is.')
                $already.Add('The items already added to the target mailbox are not added again.')
                $again.Add('The store is checked again, and so are the target folders: a missing folder is created, an existing one is used as it is.')
                $again.Add($uncertain)
                $again.Add($sendAgain)
                $from = "It resumes the second part, adding the items to the target mailbox: from step 5 of 9, Verify the export store.$reached"
            }
            elseif ($reading) {
                $already.Add('The items already read into the store are not read again.')
                $again.Add('The folder tree and the list of the items are read again from the source mailbox, so that what arrived or changed since is taken. An item that changed since is read again.')
                $again.Add("Items that failed are tried again, up to $tries tries in all.")
                $again.Add('When the reading is complete, the second part follows: adding the items to the target mailbox.')
                $from = "It resumes the first part, reading the source mailbox: from step 1 of 9, Read the folder tree.$reached"
            }
            else {
                $already.Add('A part that is complete is not done again: when the reading of the source mailbox is complete, the tool goes straight to adding the items to the target mailbox.')
                $already.Add('The items already read into the store are not read again, and the items already added to the target mailbox are not added again.')
                $again.Add('When the reading is not complete, the tool reads what is left first. Then, in the target mailbox, items that were being sent when it stopped are looked up in the target folder first, and items that failed are tried again.')
                $from = 'It resumes the part where it stopped: the reading of the source mailbox, or, when that is complete, the adding of the items to the target mailbox.'
            }
        }
        default {
            $heading = 'Continue the job'
            $already.Add('What the job already did and recorded is not done again.')
            $again.Add('The rest of the job runs again from its first step.')
            $from = "It restarts from the first step.$reached"
        }
    }
    if ($kind -eq 'interrupted') { $notes.Add('The PC was shut down, or the tool was closed, while it ran: it could not finish the item it was working on, and an item that was being sent is uncertain until it is looked up.') }
    elseif ($kind -eq 'signin') { $notes.Add('Sign in again before you continue: the sign-in expired.') }
    $problem = ''
    if ([bool](Get-MbtJobField $Job 'Alive' $false)) { $problem = 'This job is running now, in another window or process. Stop it there first.' }
    elseif ([string](Get-MbtJobField $Job 'Batch' '')) { $problem = "This job belongs to the batch $(Get-MbtJobField $Job 'Batch' ''). A batch is continued from the command line: run the same command with -ResumeBatch $(Get-MbtJobField $Job 'Batch' '')." }
    elseif ($null -eq $Request) { $problem = 'This job has no request.json, so nothing says what it ran. Run its command again from the command line.' }
    elseif ([string](Get-MbtJobField $Request 'Source' '') -eq 'cli') {
        if ([bool](Get-MbtJobField $Request 'NeedsSecret' $false)) { $problem = 'This job was started with a certificate password, which is never saved. Run its command again with -CertificatePassword.' }
        else {
            $folder = [string](Get-MbtJobField $Request 'WorkingDirectory' '')
            if (-not $folder -or -not [IO.Directory]::Exists($folder)) { $problem = "This job ran in the folder '$folder', which is not there any more. Its journals and its relative paths are relative to that folder: create it again, or run its command again from the command line." }
        }
    }
    [pscustomobject]@{ Mode = $mode; Heading = $heading; Already = @($already); Again = @($again); From = $from; Notes = @($notes); Problem = $problem }
}

#endregion

#region The job folder of a command-line run ------------------------------------------------------------------------

# What a run does not record: the job folder gives these (events, stop, id), and they say nothing about what is asked.
$script:MbtCliSkipParameter = @('ResumeJob', 'RemoveJob', 'RemoveCopyStore', 'EventsFile', 'StopFile', 'JobId', 'LogFolder', 'Gui', 'DevelopmentMode', 'Verbose', 'Debug', 'WhatIf', 'Confirm', 'ErrorAction', 'WarningAction', 'InformationAction', 'ErrorVariable', 'WarningVariable', 'InformationVariable', 'OutVariable', 'OutBuffer', 'PipelineVariable', 'ProgressAction', 'WarningVariable')
# What the key of a command leaves out: how it signs in and how it is shown do not make it another command.
$script:MbtCliKeySkipParameter = @('NoPrompt', 'DeviceCode', 'KeepAwake', 'ConfigPath')

function ConvertTo-MbtCliValue {
    <# A parameter value as it is written in request.json: a switch is $true, a date is ISO text, a list is a list, a secure string is never written. #>
    param($Value)
    if ($Value -is [System.Management.Automation.SwitchParameter]) { return [bool]$Value.IsPresent }
    if ($Value -is [datetime]) { return $Value.ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [securestring]) { return $null }
    if ($Value -is [System.Collections.IEnumerable] -and $Value -isnot [string]) { return @($Value | ForEach-Object { ConvertTo-MbtCliValue $_ }) }
    return $Value
}

function Get-MbtCliJobParameter {
    <# The parameters of the command that was recorded, as a hashtable ready to splat; the folder it ran in is in the request (WorkingDirectory). #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Request)
    $recorded = Get-MbtJobField $Request 'Parameters'
    $out = [ordered]@{}
    if ($null -eq $recorded) { return $out }
    foreach ($p in $recorded.PSObject.Properties) {
        $v = $p.Value
        if ($v -is [bool]) { if ($v) { $out[$p.Name] = [System.Management.Automation.SwitchParameter]::new($true) } else { $out[$p.Name] = $false } }
        elseif ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $out[$p.Name] = [string[]]@($v | ForEach-Object { [string]$_ }) }
        else { $out[$p.Name] = $v }
    }
    return $out
}

function Read-MbtJobRequest {
    <# request.json of a job folder as an object, or $null when there is none or it cannot be read. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Folder)
    $file = Join-Path $Folder 'request.json'
    if (-not [IO.File]::Exists($file)) { return $null }
    try { return (Read-MbtJsonFile -Path $file) } catch { return $null }
}

function Get-MbtCliConsoleState {
    <# The values of the console's own request that match a command line, so that the console can continue a job of the command line and show its end page (Mailbox, Dest, OutFolder...). #>
    param([Parameter(Mandatory)][string]$Mode, [Parameter(Mandatory)][hashtable]$Parameter)
    $p = $Parameter
    $get = { param($n, $d) if ($p.ContainsKey($n) -and $null -ne $p[$n]) { $p[$n] } else { $d } }
    $first = { param($v) @($v | Where-Object { $_ })[0] }
    $state = [ordered]@{}
    $state.Mailbox = [string](& $get 'UserPrincipalName' '')
    $state.AuthMode = $(if ($p.ContainsKey('AppOnly') -and $p['AppOnly']) { 'app' } else { 'user' })
    $state.Awake = [bool](& $get 'KeepAwake' $false)
    $outPst = [string](& $get 'OutputPst' '')
    $setOutput = {
        param([string]$Pst)
        if ($Pst) { $state.OutFolder = [string](Split-Path $Pst -Parent); $state.OutName = [string](Split-Path $Pst -Leaf) }
    }
    switch ($Mode) {
        'Export' {
            $state.Task = 'export'; $state.Dest = $(if ($outPst) { 'pst' } else { 'store' }); & $setOutput $outPst
            $state.SplitGB = [int](& $get 'PstMaxSizeGB' 50); $state.CheckEach = -not [bool](& $get 'SkipCheck' $false); $state.SepAux = [bool](& $get 'SeparateAuxArchives' $false); $state.KeepStore = $true
        }
        'ExportPst' {
            if ($p.ContainsKey('PstPath')) { $state.Task = 'copy'; $state.OpSrc = 'pst'; $state.OpPath = [string](& $first $p['PstPath']) }
            else { $state.Task = 'export'; $state.Dest = 'pst' }
            & $setOutput $outPst
            $state.SplitGB = [int](& $get 'PstMaxSizeGB' 50); $state.CheckEach = -not [bool](& $get 'SkipCheck' $false); $state.SepAux = [bool](& $get 'SeparateAuxArchives' $false); $state.KeepStore = $true
        }
        'ExportEml' { $state.Task = 'export'; $state.Dest = 'eml'; $state.OutFolder = [string](& $get 'EmlPath' '') }
        { $_ -in 'Import', 'ImportPst', 'ImportEml' } {
            $state.Task = 'import'
            $state.ImSrc = switch ($Mode) { 'ImportPst' { 'pst' } 'ImportEml' { 'eml' } default { 'store' } }
            $state.ImPath = switch ($Mode) { 'ImportPst' { [string](& $first $p['PstPath']) } 'ImportEml' { [string](& $get 'EmlPath' '') } default { [string](& $get 'StorePath' '') } }
            $state.TgtLoc = [string](& $get 'Mailbox' 'Primary')
            $state.TgtName = [string](& $get 'TargetRoot' ''); $state.TgtMode = $(if ($state.TgtName) { 'new' } else { 'same' })
        }
        'Copy' {
            $state.Task = 'mailcopy'
            $state.McFromKind = [string](& $get 'Mailbox' 'Primary')
            $state.TgtLoc = [string](& $get 'TargetMailbox' 'Primary')
            $state.TgtName = [string](& $get 'TargetRoot' ''); $state.TgtMode = $(if ($state.TgtName) { 'new' } else { 'same' })
            $state.McTarget = [string](& $get 'TargetMailboxId' '')
        }
    }
    return $state
}

function Format-MbtCliCommandLine {
    <# The command of a run as one line an operator can type: .\Invoke-MailboxTransfer.ps1 -Mode Export -Name value. #>
    param([Parameter(Mandatory)][string]$Mode, [Parameter(Mandatory)][hashtable]$Parameter)
    $quote = { param([string]$v) if ($v -match '^[\w\.\-@:\\/]+$') { $v } else { "'" + $v.Replace("'", "''") + "'" } }
    $tokens = [Collections.Generic.List[string]]::new()
    $tokens.Add('.\Invoke-MailboxTransfer.ps1'); $tokens.Add('-Mode'); $tokens.Add($Mode)
    foreach ($name in @($Parameter.Keys | Where-Object { $_ -ne 'Mode' -and $_ -notin $script:MbtCliSkipParameter } | Sort-Object)) {
        $v = ConvertTo-MbtCliValue $Parameter[$name]
        if ($null -eq $v -or $v -eq $false) { continue }
        if ($v -eq $true) { $tokens.Add('-' + $name); continue }
        $tokens.Add('-' + $name); $tokens.Add((@($v | ForEach-Object { & $quote ([string]$_) }) -join ','))
    }
    return ($tokens -join ' ')
}

function Get-MbtCliJobKey {
    <# The key of a command: the same command in the same folder has the same key, so that running it again continues its job. #>
    param([Parameter(Mandatory)][string]$Mode, [Parameter(Mandatory)][hashtable]$Parameter, [Parameter(Mandatory)][string]$WorkingDirectory)
    $parts = [Collections.Generic.List[string]]::new()
    $parts.Add($Mode.ToLowerInvariant()); $parts.Add($WorkingDirectory.TrimEnd('\').ToLowerInvariant())
    foreach ($name in @($Parameter.Keys | Where-Object { $_ -notin $script:MbtCliSkipParameter -and $_ -notin $script:MbtCliKeySkipParameter } | Sort-Object { $_.ToLowerInvariant() })) {
        $v = ConvertTo-MbtCliValue $Parameter[$name]
        if ($null -eq $v) { continue }
        $parts.Add($name.ToLowerInvariant() + '=' + (@($v | ForEach-Object { ([string]$_).ToLowerInvariant() }) -join '|'))
    }
    return [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes(($parts -join "`n")))).Substring(0, 16).ToLowerInvariant()
}

function New-MbtCliJob {
    <#
    .SYNOPSIS
        The job folder of a single job started from the command line: <work folder>\jobs\<date>-<task>-<id>, with request.json (the command, the folder it ran in, the absolute paths of the
        store and the journal), so that the console lists the job and continues it, and -ResumeJob runs it again.
    .DESCRIPTION
        The same command in the same folder again, while the job of that command is not complete and not running, continues that job folder instead of making a new one. Returns Folder, Name,
        EventsFile, JobId and Reused. -Parameter is $PSBoundParameters of the script; -Resolved the absolute paths the script settled (StorePath, TempPath, JournalPath).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$WorkFolder, [Parameter(Mandatory)][string]$Mode, [Parameter(Mandatory)][string]$Task, [Parameter(Mandatory)][hashtable]$Parameter,
        [Parameter(Mandatory)][string]$WorkingDirectory, [Parameter(Mandatory)][string]$ScriptPath, [hashtable]$Resolved = @{}, [string]$Title
    )
    $key = Get-MbtCliJobKey -Mode $Mode -Parameter $Parameter -WorkingDirectory $WorkingDirectory
    $jobs = Join-Path $WorkFolder 'jobs'
    $folder = $null; $reused = $false
    if ([IO.Directory]::Exists($jobs)) {
        $needle = '"Key": "' + $key + '"'
        foreach ($dir in @([IO.DirectoryInfo]::new($jobs).EnumerateDirectories() | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 300)) {
            $requestFile = Join-Path $dir.FullName 'request.json'
            if (-not [IO.File]::Exists($requestFile)) { continue }
            $text = $null; try { $text = [IO.File]::ReadAllText($requestFile) } catch { continue }
            if ($text.IndexOf($needle, [StringComparison]::Ordinal) -lt 0) { continue }
            $job = @(Get-MbtJob -WorkFolder $WorkFolder -Name $dir.Name)
            if ($job.Count -and [string]$job[0].EffectiveState -notin 'complete' -and -not $job[0].Alive) { $folder = $dir.FullName; $reused = $true; break }
        }
    }
    if (-not $folder) {
        $id = -join ((1..4) | ForEach-Object { '0123456789abcdef'[(Get-Random -Maximum 16)] })
        $folder = Join-Path $jobs ('{0}-{1}-{2}' -f (Get-Date).ToString('yyyyMMdd-HHmm', [Globalization.CultureInfo]::InvariantCulture), $Task, $id)
        $null = New-Item -ItemType Directory -Force -Path $folder
    }
    $name = Split-Path $folder -Leaf
    $recorded = [ordered]@{}
    $secret = $false
    foreach ($k in $Parameter.Keys) {
        if ($k -in $script:MbtCliSkipParameter) { continue }
        $v = ConvertTo-MbtCliValue $Parameter[$k]
        if ($null -eq $v) { if ($Parameter[$k] -is [securestring]) { $secret = $true }; continue }
        $recorded[$k] = $v
    }
    $request = [ordered]@{
        Version = 1; Source = 'cli'; Key = $key; Mode = $Mode; Title = $Title; Created = [datetime]::UtcNow.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        JobId = $name; JobName = [string](Get-MbtJobField $Parameter 'JobName' ''); WorkRoot = $WorkFolder; ScriptPath = $ScriptPath; WorkingDirectory = $WorkingDirectory
        CommandLine = (Format-MbtCliCommandLine -Mode $Mode -Parameter $Parameter); NeedsSecret = $secret; Parameters = $recorded
    }
    foreach ($k in $Resolved.Keys) { $request[$k] = $Resolved[$k] }
    foreach ($kv in (Get-MbtCliConsoleState -Mode $Mode -Parameter $Parameter).GetEnumerator()) { if (-not $request.Contains($kv.Key) -and $null -ne $kv.Value -and [string]$kv.Value -ne '') { $request[$kv.Key] = $kv.Value } }
    $tmp = Join-Path $folder "request.json.$PID.tmp"
    [IO.File]::WriteAllText($tmp, ($request | ConvertTo-Json -Depth 6), [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $tmp -Destination (Join-Path $folder 'request.json') -Force
    [pscustomobject]@{ Folder = $folder; Name = $name; EventsFile = (Join-Path $folder 'events.jsonl'); JobId = $name; Reused = $reused }
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB4mj43qQ+d7ePk
# KdlsX3Cbs3xRhhqottc2K1HtxTe396CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAIbMvKIjn4
# oJ+DaIKYE991iqxsyB+8/Y120Zy4TvypBTANBgkqhkiG9w0BAQEFAASCAgBkAhhS
# H0t9Lkim1vgsWBjePB7G+CSocUTpyXpV7fjj6vsogVz7o22PR46z2PaR6vdDE8CU
# q6nDofvpp7UFvH8fDlLfiFXpB1kUvrOT4IGTK27+2VzpeUTSOUKKb9YoKfqlz0X0
# X34eSWPZjvtzwbzokE4D7EV+Z4ABfFemsroMRPOw53rDQ5GzV+GuYutafNwAEALU
# vh96sRLvJc+ZDr50DRFpyU8Dgc2+M6ajZ8k38GyU8qUlnERx2msaDdfMi8CjAtVY
# NyipxBi43CMXQJ1PWT7lS81yE7WQzXunXIxE7RehurVGVqqJbrbBlKoikunjvd02
# gK+RIMzizZcJ5cRhbDPztMptGGuLV/SNeAXvCLi7QmcqwwV72L2kWEgl9UMHlHru
# /1XEENms3zREFellxVxLnk99CRJyJ5Ve5674JbTGpsbq/MZH2tM2BultHmjiKmHL
# RfbSsUj1ZGGaFsUHQDprdpyy7dc/rN81Fjg/VCFH4Qj/zA452dv2ccSye+mzvfxk
# OKySKr6uGf0prrZnzpDjY6Cj3CBQ5oyQqwdcT8rB+gCg2teK8yNIad1kB26nVghz
# DP6ohcWrsTXiXr6O7qJqzIeS2g+QkYe3SBi53irVyhTGezi+r7YP55PjdCoQvHRr
# j6gRgO3C2qivg8WBdM08IWp87RnqQWyb+yHqg6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIDfoohIVHOHiTL8oxNOHE+5ePdHESF6qbgPH92l7y4xJAghe42siVrOk
# qBgPMjAyNjExMjUxNTM3MDNaMAMCAQECEG1ZqP4HjIaMoLdhByef6NOgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTEyNTE1MzcwM1owKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIM1m
# lS5QJgFNHVaOHBPK5gbUxq9IHa5PGMwLSJntXXl+MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiBYRW8hebcpsWOHOCFBesnxX9uh3QWnbmd6hZdfRNyDPAIhAMlf
# 6WWKXUKFb2k8ndq6ykDcd44e2X2Znl8zDDUTQ+a6
# SIG # End signature block
