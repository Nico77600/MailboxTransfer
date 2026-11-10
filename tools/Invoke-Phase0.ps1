#Requires -Version 7.4
<#
.SYNOPSIS
    Phase 0 (lab): proves on the real tenant that the mailbox import and export API of Microsoft Graph does what the
    product needs. Reads by default; exports and imports only when asked.

.DESCRIPTION
    Steps, in this order (each one only runs when its parameters are given):

      1. SIGN IN        Delegated (MSAL.NET, refresh token cached with DPAPI): browser, or with -DeviceCode a code to enter
                        at https://microsoft.com/devicelogin. Shows the account and the permissions really granted in the token.
      2. TREE           Checks the physical mailbox ids (primary and archive), lists the complete folder tree of both
                        with item counts (-Sizes adds the size of each folder) and records them in the export store. Read only.
                        Skipped with -SkipTree. The primary id is read from Graph when not given (-PrimaryMailboxId); the
                        archive id has no API and must be given (-ArchiveMailboxId).
      3. ITEMS          -SourceFolder (-ListItems, or -Export): lists the items of one folder (id, type, size, dates) and
                        records them in the store (Pending).
      4. EXPORT         -Export: exports the chosen items (-ItemId, -First, -Largest or -SubjectContains) into the export
                        store: each answer is streamed to a temporary file, copied into a pack with its SHA-256 (compared with
                        the one computed on arrival), and its preview (subject, sender, recipients, dates, attachments, body
                        excerpt) is decoded with the MS-OXCFXICS reader and recorded; shows time and throughput; then decodes
                        every stored item from its pack (properties, subject, sender, recipients, dates, body, attachments,
                        folder information, envelope).
      5. IMPORT         -Import: creates ONE item read from the store (-ItemId, else the most recently exported one) in
                        -TargetFolder of -TargetMailbox (import job of the store, createImportSession, importItem in "create"
                        mode streamed from the pack), lists the target folder again and shows the new
                        item; -VerifyRoundTrip exports it back and compares it with the source. With -WhatIf only the plan.
                        -CreateTargetFolder creates the target folder (create only) when it does not exist.

    Nothing is ever deleted or modified in a mailbox: the only writes are the creation of the imported item and, with
    -CreateTargetFolder, of the target folder.

    Every default (tenant, application, user, mailbox ids, folders) comes from the Lab section of
    package\config\MailboxTransfer.config.psd1 (the target of the lab is adele.vance@contoso.com) and can be given as a parameter.

    The export store (-ExportPath, default <OutputPath>\export: catalog.db + packs\) replaces separate files: folders, items and
    their state, the streams (in packs), the previews and the import jobs are in it. Only one process may hold it open.

    Everything shown is also written to a transcript in the output folder (it contains subjects and addresses of the
    lab mailbox: keep it private).

.PARAMETER SourceFolder
    Folder of the source mailbox: id, full path, or end of the path ("MBT-Fill-20261028", "Lab/Attachments").
    An ambiguous name is refused with the list of the candidates. Default (with -ListItems or -Export): Lab.SourceFolder.

.PARAMETER TargetFolder
    Folder of the target mailbox, same notation. Default (with -Import): Lab.TargetFolder. When it does not exist,
    -CreateTargetFolder creates it (and its missing parents) with POST .../folders, create only.

.PARAMETER PrimaryMailboxId
    MBX:<mailbox guid>@<tenant id> (or the bare guid). Default: Lab.PrimaryMailboxId of the configuration (adele.vance@contoso.com), else read
    with GET /users/{upn}/settings/exchange (User.Read.All).

.PARAMETER ArchiveMailboxId
    Same for the archive mailbox (default Lab.ArchiveMailboxId). No API gives it: Exchange Online PowerShell, Get-MailboxLocation -User <upn>
    (MailboxLocationType MainArchive).

.PARAMETER DeviceCode
    Sign in with a device code instead of the browser (the browser and WAM fail when no window can be shown). The line
    "DEVICE CODE: <code> URL: <url>" is written to the console at once and to out\device-code.txt (overwritten); the script
    waits until the code expires (about 15 minutes). The token cache is the same: a later run needs no new code.

.PARAMETER OutputPath
    Folder of the results (default: out\phase0 next to the tools folder): report, transcript and, by default, the export store.
    Use the same folder for the export and the import.

.PARAMETER ExportPath
    Folder of the export store (default: <OutputPath>\export). Created on the first run, reopened (with its crash recovery) later.

.EXAMPLE
    # read only: sign in, check the ids, folder trees of the primary and the archive
    .\tools\Invoke-Phase0.ps1 -DeviceCode

.EXAMPLE
    # list the items of the default source folder of the primary mailbox
    .\tools\Invoke-Phase0.ps1 -SkipTree -ListItems

.EXAMPLE
    # export the largest item of a folder and decode it
    .\tools\Invoke-Phase0.ps1 -SkipTree -SourceFolder 'MBT-Fill-20261028' -Export -Largest 1

.EXAMPLE
    # plan of the import, then the import (creation only) with a round-trip check
    .\tools\Invoke-Phase0.ps1 -SkipTree -Import -CreateTargetFolder -WhatIf
    .\tools\Invoke-Phase0.ps1 -SkipTree -Import -CreateTargetFolder -VerifyRoundTrip

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0 (phase 0)
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$UserPrincipalName,
    [string]$PrimaryMailboxId,
    [string]$ArchiveMailboxId,
    [string]$OutputPath,
    [string]$ExportPath,
    [switch]$DeviceCode,

    [switch]$SkipTree,
    [switch]$Sizes,

    [ValidateSet('Primary', 'Archive')][string]$SourceMailbox = 'Primary',
    [string]$SourceFolder,
    [switch]$ListItems,
    [int]$ListLimit = 25,
    [int]$PageSize = 0,

    [switch]$Export,
    [string[]]$ItemId,
    [int]$First,
    [int]$Largest,
    [string]$SubjectContains,
    [int]$ScanLimit = 100,
    [switch]$AllProperties,

    [switch]$Import,
    [ValidateSet('Primary', 'Archive')][string]$TargetMailbox = 'Archive',
    [string]$TargetFolder,
    [switch]$CreateTargetFolder,
    [switch]$VerifyRoundTrip
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# -WhatIf only concerns the creation of the imported item. The result files of the run (reports, streams, transcript) are always written.
$dryRun = [bool]$WhatIfPreference
$WhatIfPreference = $false
$repoRoot = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $repoRoot 'package\MailboxTransfer.psd1') -Force
$config = Import-MbtConfiguration
if (-not $TenantId) { $TenantId = $config.Lab.TenantId }
if (-not $ClientId) { $ClientId = $config.Lab.ClientId }
if (-not $UserPrincipalName) { $UserPrincipalName = $config.Lab.UserPrincipalName }
if (-not $PrimaryMailboxId) { $PrimaryMailboxId = $config.Lab.PrimaryMailboxId }
if (-not $ArchiveMailboxId) { $ArchiveMailboxId = $config.Lab.ArchiveMailboxId }
if (-not $SourceFolder -and ($ListItems -or $Export)) { $SourceFolder = $config.Lab.SourceFolder }
if (-not $TargetFolder -and $Import) { $TargetFolder = $config.Lab.TargetFolder }

if (-not $OutputPath) { $OutputPath = Join-Path $repoRoot 'out\phase0' }
$null = New-Item -ItemType Directory -Force -Path $OutputPath
$OutputPath = (Resolve-Path -LiteralPath $OutputPath).Path
$mailboxes = @{
    Primary = if ($PrimaryMailboxId) { ConvertTo-MbtMailboxId -Value $PrimaryMailboxId -TenantId $TenantId } else { '' }
    Archive = if ($ArchiveMailboxId) { ConvertTo-MbtMailboxId -Value $ArchiveMailboxId -TenantId $TenantId } else { '' }
}
$report = [ordered]@{ Tool = 'MailboxTransfer'; Version = '0.1.0'; Phase = 0; StartedUtc = [DateTime]::UtcNow.ToString('o'); Steps = [ordered]@{} }
$trees = @{}
if (-not $ExportPath) { $ExportPath = Join-Path $OutputPath 'export' }
$store = Open-MbtExportStore -Path $ExportPath -Name 'phase0' -Source "$UserPrincipalName ($TenantId)"

function Get-MailboxId {
    <# The id of the primary or archive mailbox: given, else (primary) read from Graph, else an error that says how to get it. #>
    param([Parameter(Mandatory)][ValidateSet('Primary', 'Archive')][string]$Which)
    if ($mailboxes[$Which]) { return $mailboxes[$Which] }
    if ($Which -eq 'Primary') {
        $mailboxes.Primary = Get-MbtPrimaryMailboxId -UserPrincipalName $UserPrincipalName -TenantId $TenantId
        Write-MbtItem OK "primary mailbox id of $UserPrincipalName read from Graph: $($mailboxes.Primary)"
        return $mailboxes.Primary
    }
    throw ("The archive mailbox id of $UserPrincipalName is not known and no API gives it: pass -ArchiveMailboxId MBX:<mailbox guid>@$TenantId " +
        "(or set Lab.ArchiveMailboxId in the configuration). The guid: Exchange Online PowerShell, Get-MailboxLocation -User $UserPrincipalName " +
        "(MailboxLocationType MainArchive).")
}

function Get-Tree {
    param([Parameter(Mandatory)][ValidateSet('Primary', 'Archive')][string]$Which, [switch]$WithSizes)
    if ($trees.ContainsKey($Which) -and -not $WithSizes) { return $trees[$Which] }
    $trees[$Which] = @(Get-MbtMailboxTree -MailboxId (Get-MailboxId $Which) -Sizes:$WithSizes -PageSize $PageSize)
    return $trees[$Which]
}

function Show-Tree {
    param([Parameter(Mandatory)][object[]]$Tree)
    $Tree | Format-Table -AutoSize `
        @{ n = 'Folder'; e = { ('  ' * $_.Depth) + $_.DisplayName } },
        @{ n = 'Items'; e = { $_.TotalItemCount }; a = 'right' },
        @{ n = 'Sub'; e = { $_.ChildFolderCount }; a = 'right' },
        @{ n = 'Size'; e = { if ($_.TotalSize -ge 0) { Format-MbtBytes $_.TotalSize } else { '-' } }; a = 'right' },
        Type, WellKnownName | Out-String -Width 220 | ForEach-Object { $_.TrimEnd() } | Write-Host
}

function Short([string]$Id) { if ($Id.Length -gt 22) { '...' + $Id.Substring($Id.Length - 22) } else { $Id } }

$transcript = Join-Path $OutputPath ("phase0-{0:yyyyMMdd-HHmmss}.log" -f (Get-Date))
Start-Transcript -LiteralPath $transcript | Out-Null
try {
    # ==== 1. sign in ===============================================================================================
    Write-MbtStep ('1. Sign in (delegated, {0})' -f $(if ($DeviceCode) { 'device code' } else { 'browser' }))
    if ($DeviceCode) { Write-MbtItem INFO 'a line "DEVICE CODE: ..." follows if no valid token is cached; it is also written to out\device-code.txt' }
    $session = Connect-MbtGraph -TenantId $TenantId -ClientId $ClientId -UserPrincipalName $UserPrincipalName -DeviceCode:$DeviceCode -DeviceCodeFile (Join-Path $repoRoot 'out\device-code.txt')
    $t = $session.Token
    Write-MbtItem OK "signed in as $($t.Account) ($($t.Name)), tenant $($t.TenantId), application $($t.AppId)"
    if ($t.Account -and $t.Account -ine $UserPrincipalName) { Write-MbtItem WARN "signed in as $($t.Account), not as $UserPrincipalName (the target of the lab)" }
    Write-MbtDetail "token for $($t.Audience), valid until $($t.ExpiresOn.ToLocalTime().ToString('HH:mm:ss')) (renewed silently with the refresh token)"
    Write-MbtDetail ("permissions in the token: " + ($t.Scopes -join ' '))
    $has = { param($n) $n -in $t.Scopes }
    $problems = @()
    if (-not ((& $has 'MailboxFolder.Read') -or (& $has 'MailboxFolder.ReadWrite'))) { $problems += 'MailboxFolder.Read[Write]' }
    if (-not ((& $has 'MailboxItem.Read') -or (& $has 'MailboxItem.ReadWrite'))) { $problems += 'MailboxItem.Read[Write]' }
    if (-not ((& $has 'MailboxItem.ImportExport') -or ((& $has 'MailboxItem.Export') -and -not $Import))) { $problems += 'MailboxItem.ImportExport' }
    if ($problems) {
        Write-MbtItem WARN ("missing in the token: " + ($problems -join ', ') + ". Run tools\Grant-LabConsent.ps1 (as an administrator) then sign in again.")
    }
    $report.Steps.SignIn = [ordered]@{ Account = $t.Account; Scopes = $t.Scopes; ExpiresOn = $t.ExpiresOn.ToString('o'); MissingPermissions = $problems }

    # ==== 2. mailbox ids and trees =================================================================================
    if (-not $SkipTree) {
        Write-MbtStep '2. Mailbox ids and folder trees (read only)'
        $treeReport = [ordered]@{}
        foreach ($which in 'Primary', 'Archive') {
            if ($which -eq 'Archive' -and -not $mailboxes.Archive) {
                Write-Host ''
                Write-Host '-- Archive' -ForegroundColor White
                Write-MbtItem WARN 'archive mailbox id not given: no API reads it. Pass -ArchiveMailboxId MBX:<mailbox guid>@<tenant id>.'
                Write-MbtDetail "Exchange Online PowerShell: Get-MailboxLocation -User $UserPrincipalName | Format-Table MailboxLocationType, MailboxGuid   (MainArchive)"
                $treeReport[$which] = [ordered]@{ MailboxId = $null; Reachable = $false; Error = 'archive mailbox id not given' }
                continue
            }
            $id = Get-MailboxId $which
            Write-Host ''
            Write-Host "-- $which $id" -ForegroundColor White
            $check = Test-MbtMailboxId -MailboxId $id -TenantId $TenantId -UserPrincipalName $UserPrincipalName
            Write-MbtItem $(if ($check.Syntax -and $check.TenantMatches) { 'OK' } else { 'FAIL' }) "id syntax MBX:<guid>@<tenant>: $($check.Syntax); tenant is the lab tenant: $($check.TenantMatches)"
            if ($check.GraphLookup -eq 'ok') {
                $level = if ($which -eq 'Primary') { if ($check.IsPrimaryPerGraph) { 'OK' } else { 'FAIL' } } else { if ($check.IsPrimaryPerGraph) { 'FAIL' } else { 'OK' } }
                Write-MbtItem $level "Graph says the primary mailbox of $UserPrincipalName is $($check.PrimaryMailboxIdFromGraph)"
                if ($which -eq 'Archive') { Write-MbtDetail 'the archive id cannot be resolved by this API (only primary and shared mailboxes): it is only proven usable below' }
            } else {
                Write-MbtItem INFO "id lookup through GET /users/{upn}/settings/exchange: $($check.GraphLookup) (needs User.Read.All, not part of the lab consent)"
            }
            if (-not $check.Reachable) {
                Write-MbtItem FAIL "the API does not answer for this id: $($check.Error)"
                $treeReport[$which] = [ordered]@{ MailboxId = $id; Reachable = $false; Error = $check.Error }
                continue
            }
            Write-MbtItem OK 'the folders endpoint answers for this id'
            $watch = [Diagnostics.Stopwatch]::StartNew()
            $requestsBefore = $session.Graph.RequestCount
            $tree = @(Get-Tree -Which $which -WithSizes:$Sizes)
            $watch.Stop()
            $items = ($tree | Measure-Object TotalItemCount -Sum).Sum
            Write-MbtItem OK ("{0} folders, {1} items, listed in {2} with {3} requests" -f $tree.Count, $items, (Format-MbtDuration $watch.Elapsed.TotalSeconds), ($session.Graph.RequestCount - $requestsBefore))
            if ($Sizes) { Write-MbtDetail ("total size of the items: " + (Format-MbtBytes (($tree | Where-Object { $_.TotalSize -ge 0 } | Measure-Object TotalSize -Sum).Sum))) }
            Show-Tree $tree
            Save-MbtFolderTree -Store $store -MailboxId $id -Kind $which -UserPrincipalName $UserPrincipalName -Tree $tree
            Write-MbtDetail "folders recorded in the export store ($ExportPath)"
            $treeReport[$which] = [ordered]@{
                MailboxId = $id; Reachable = $true; Syntax = $check.Syntax; TenantMatches = $check.TenantMatches; GraphLookup = $check.GraphLookup
                IsPrimaryPerGraph = $check.IsPrimaryPerGraph; Folders = $tree.Count; Items = $items; Seconds = [math]::Round($watch.Elapsed.TotalSeconds, 2)
            }
        }
        $report.Steps.Trees = $treeReport
    }

    # ==== 3. items of a folder =====================================================================================
    $sourceId = if ($SourceFolder -or $Export) { Get-MailboxId $SourceMailbox }
    $folder = $null
    $listing = @()
    if ($SourceFolder) {
        Write-MbtStep "3. Items of '$SourceFolder' in the $SourceMailbox mailbox"
        $folder = Resolve-MbtFolder -Tree (Get-Tree -Which $SourceMailbox) -Name $SourceFolder
        Write-MbtItem OK "folder: $($folder.Path)  [$($folder.Id)]  ($($folder.TotalItemCount) items announced)"
        $pagesBefore = $session.Graph.PagesRead
        $watch = [Diagnostics.Stopwatch]::StartNew()
        $listing = @(Get-MbtFolderItems -MailboxId $sourceId -FolderId $folder.Id -PageSize $PageSize)
        $watch.Stop()
        $total = ($listing | Measure-Object Size -Sum).Sum
        Write-MbtItem OK ("{0} items, {1} in total, listed in {2} ({3} pages, about {4} items per page{5})" -f $listing.Count, (Format-MbtBytes $(if ($total) { $total } else { 0 })), (Format-MbtDuration $watch.Elapsed.TotalSeconds),
            ($session.Graph.PagesRead - $pagesBefore), $(if ($session.Graph.PagesRead -gt $pagesBefore) { [math]::Round($listing.Count / ($session.Graph.PagesRead - $pagesBefore)) } else { 0 }), $(if ($PageSize) { ", `$top=$PageSize" } else { ', service default page size' }))
        $listing | Group-Object Type | ForEach-Object { Write-MbtDetail ("{0,-22} {1} items" -f $_.Name, $_.Count) }
        $listing | Select-Object -First $ListLimit | Format-Table -AutoSize `
            @{ n = 'Id'; e = { Short $_.Id } }, Type, @{ n = 'Size'; e = { Format-MbtBytes $_.Size }; a = 'right' }, Created, Modified |
            Out-String -Width 200 | ForEach-Object { $_.TrimEnd() } | Write-Host
        if ($listing.Count -gt $ListLimit) { Write-MbtDetail "... $($listing.Count - $ListLimit) more (-ListLimit)" }
        Save-MbtFolderTree -Store $store -MailboxId $sourceId -Kind $SourceMailbox -UserPrincipalName $UserPrincipalName -Tree (Get-Tree -Which $SourceMailbox)
        Save-MbtItemList -Store $store -MailboxId $sourceId -FolderId $folder.Id -Items $listing
        Write-MbtDetail "$($listing.Count) items recorded in the export store (Pending)"
        $report.Steps.Items = [ordered]@{ Folder = $folder.Path; FolderId = $folder.Id; Count = $listing.Count; Bytes = $total; Pages = $session.Graph.PagesRead - $pagesBefore }
    }

    # ==== 4. export and decode =====================================================================================
    $exported = @()
    if ($Export) {
        Write-MbtStep '4. Export (exportItems, streamed into the export store) and decode'
        if (-not $folder -and -not $ItemId) { throw '-Export needs -SourceFolder (or explicit -ItemId).' }
        $chosen = @()
        if ($ItemId) { $chosen = @($ItemId) }
        elseif ($SubjectContains) {
            Write-MbtItem INFO "looking for a subject containing '$SubjectContains' in the first $ScanLimit items (each is exported and decoded, the listing carries no subject)"
            $scanDir = Join-Path $OutputPath 'scan-tmp'
            try {
                $candidates = @($listing | Select-Object -First $ScanLimit)
                foreach ($group in [MailboxTransfer.MailboxApi]::Chunk([string[]]($candidates | ForEach-Object Id), 20)) {
                    $scan = Export-MbtItems -MailboxId $sourceId -ItemId $group -OutputPath $scanDir
                    foreach ($e in $scan.Entries | Where-Object { $_.File }) {
                        $d = Read-MbtFxStream -Path (Join-Path $scanDir $e.File)
                        if ($d.Summary.Subject -and $d.Summary.Subject.Contains($SubjectContains, [StringComparison]::OrdinalIgnoreCase)) { $chosen += $e.ItemId; Write-MbtItem OK "match: '$($d.Summary.Subject)'  [$(Short $e.ItemId)]" }
                    }
                    if ($First -and $chosen.Count -ge $First) { break }
                }
            } finally { if (Test-Path -LiteralPath $scanDir) { Remove-Item -LiteralPath $scanDir -Recurse -Force } }
            if (-not $chosen) { throw "No item with a subject containing '$SubjectContains' among the first $ScanLimit items." }
            if ($First) { $chosen = @($chosen | Select-Object -First $First) }
        }
        elseif ($Largest) { $chosen = @($listing | Sort-Object Size -Descending | Select-Object -First $Largest | ForEach-Object Id) }
        elseif ($First) { $chosen = @($listing | Select-Object -First $First | ForEach-Object Id) }
        else { $chosen = @($listing | Select-Object -First 1 | ForEach-Object Id); Write-MbtItem INFO 'no selection given: the first item of the folder is exported (-First, -Largest, -ItemId, -SubjectContains)' }
        if (-not $chosen) { throw 'Nothing to export: the folder is empty.' }

        Write-MbtItem INFO "exporting $($chosen.Count) item(s), at most 20 per request"
        $result = Export-MbtItemsToStore -Store $store -MailboxId $sourceId -ItemId $chosen -TempPath $OutputPath -FolderId $(if ($folder) { $folder.Id }) -ItemInfo $listing
        foreach ($b in $result.Batches) {
            Write-MbtItem OK ("request {0}: {1} item(s), {2} in {3} = {4}" -f $b.Batch, $b.Items, (Format-MbtBytes $b.Bytes), (Format-MbtDuration $b.Seconds), (Format-MbtRate $b.Bytes $b.Seconds))
        }
        foreach ($e in $result.Entries) {
            if ($e.ErrorCode) { Write-MbtItem FAIL "item $(Short $e.ItemId): $($e.ErrorCode) $($e.ErrorMessage)"; continue }
            $ratio = if ($e.ListedSize) { ' (listed size {0}, stream/listed = {1:0.00})' -f (Format-MbtBytes $e.ListedSize), ($e.Bytes / [math]::Max(1, $e.ListedSize)) } else { '' }
            Write-MbtItem OK ("item {0} stored (key {1}, pack {2}, offset {3}): {4}, sha256 {5}{6}" -f (Short $e.ItemId), $e.ItemKey, $e.PackId, $e.Offset, (Format-MbtBytes $e.Bytes), $e.Sha256.Substring(0, 16), $ratio)
            if (-not $e.Sha256Matches) { Write-MbtItem WARN "the SHA-256 of the stored bytes differs from the one computed on arrival for item $(Short $e.ItemId)" }
            $sum = $e.Summary
            Write-MbtDetail ("preview: '{0}' from {1}, {2} attachment(s)" -f $sum.Subject, $(if ($sum.SenderSmtp) { $sum.SenderSmtp } else { $sum.SenderName }), $sum.Attachments.Count)
            $exported += $e
        }
        $st = $store.GetStats()
        Write-MbtDetail ("store: {0} items ({1}), {2} in {3} pack(s), {4} previews" -f $st.Items, (($st.ItemsByState | ForEach-Object { "$($_.State) $($_.Count)" }) -join ', '), (Format-MbtBytes $st.StoredBytes), $st.Packs, $st.Previews)

        Write-Host ''
        Write-Host '-- Decoding (MS-OXCFXICS reader)' -ForegroundColor White
        $decodedReport = @()
        foreach ($e in $exported) {
            $d = Read-MbtFxStream -Store $store -ItemKey $e.ItemKey
            Show-MbtFxReport -Decoded $d -AllProperties:$AllProperties
            $decodedReport += [ordered]@{
                ItemKey = $e.ItemKey; Envelope = $d.Read.Envelope; Properties = $d.Read.PropertyCount; Markers = $d.Read.MarkerCount
                ReaderError = $d.Read.Error; Truncated = $d.Read.Truncated; FolderInfoInStream = $d.Summary.HasFolderInfo
                Subject = $d.Summary.Subject; MessageClass = $d.Summary.MessageClass; Recipients = $d.Summary.Recipients.Count; Attachments = $d.Summary.Attachments.Count
                RtfCompressed = $d.Summary.HasCompressedRtf
            }
        }
        $report.Steps.Export = [ordered]@{
            Requests = $result.Batches; Decoded = $decodedReport
            Items = @($result.Entries | Select-Object ItemKey, ItemId, ChangeKey, ListedSize, Bytes, Sha256, PackId, Offset, Sha256Matches, ErrorCode, ErrorMessage)
            Store = [ordered]@{ Path = $ExportPath; Items = $st.Items; Packs = $st.Packs; StoredBytes = $st.StoredBytes; Previews = $st.Previews }
        }
    }

    # ==== 5. import ================================================================================================
    if ($Import) {
        Write-MbtStep "5. Import into the $TargetMailbox mailbox (creation only)"
        if (-not $TargetFolder) { throw '-Import needs -TargetFolder.' }
        # The item to import: -ItemId (first), else the one exported by this run, else the most recently exported one of the store.
        $candidates = @(Get-MbtStoreItems -Store $store)
        if ($ItemId) { $candidates = @($candidates | Where-Object { $_.ItemId -in $ItemId }) }
        elseif ($exported.Count) { $keys = @($exported | ForEach-Object ItemKey); $candidates = @($candidates | Where-Object { $_.Key -in $keys }) }
        $chosenItem = @($candidates | Sort-Object ExportedUtc, Key | Select-Object -Last 1)[0]
        if (-not $chosenItem) { throw "No exported item in the store ($ExportPath) to import: run -Export first (same -OutputPath / -ExportPath), or check -ItemId." }
        $targetId = Get-MailboxId $TargetMailbox
        $targetTree = Get-Tree -Which $TargetMailbox
        $target = Resolve-MbtFolder -Tree $targetTree -Name $TargetFolder -NullIfMissing
        if (-not $target) {
            if (-not $CreateTargetFolder) { throw "Folder '$TargetFolder' does not exist in the $TargetMailbox mailbox. Add -CreateTargetFolder to create it (create only), or give an existing -TargetFolder." }
            if ($dryRun) {
                Write-MbtItem WARN "WhatIf: the folder '$TargetFolder' does not exist in the $TargetMailbox mailbox, it would be created (create only), then the item imported in it."
            } else {
                $target = New-MbtFolderPath -MailboxId $targetId -Tree $targetTree -Path $TargetFolder -Confirm:$false
                Write-MbtItem OK "folder created in the $TargetMailbox mailbox: $($target.Path)  [$($target.Id)]"
                $trees.Remove($TargetMailbox)
            }
        }
        $source = Read-MbtFxStream -Store $store -ItemKey $chosenItem.Key
        Write-MbtItem INFO ("store   : item {0} (key {1}, pack {2}, {3}) exported {4:u} from {5}" -f (Short $chosenItem.ItemId), $chosenItem.Key, $chosenItem.PackId, (Format-MbtBytes $chosenItem.Length), $chosenItem.ExportedUtc, $chosenItem.FolderPath)
        Write-MbtItem INFO ("content : '{0}' class {1}, {2} attachment(s)" -f $source.Summary.Subject, $source.Summary.MessageClass, $source.Summary.Attachments.Count)
        $created = $null
        if ($dryRun) {
            if ($target) { Write-MbtItem INFO ("target  : {0} mailbox, folder {1}  [{2}]  ({3} items now)" -f $TargetMailbox, $target.Path, $target.Id, $target.TotalItemCount) }
            Write-MbtItem WARN ("WhatIf: nothing is created. Without -WhatIf the item would be created in {0} from {1}." -f $(if ($target) { $target.Path } else { $TargetFolder }), "store item $($chosenItem.Key)")
        } else {
            Write-MbtItem INFO ("target  : {0} mailbox, folder {1}  [{2}]  ({3} items now)" -f $TargetMailbox, $target.Path, $target.Id, $target.TotalItemCount)
            $beforeItems = @(Get-MbtFolderItems -MailboxId $targetId -FolderId $target.Id -PageSize $PageSize)
            Write-MbtDetail "listed $($beforeItems.Count) items in the target folder before the import"
            $result = Import-MbtStoreItem -Store $store -ItemKey $chosenItem.Key -MailboxId $targetId -FolderId $target.Id -JobName 'phase0' -Confirm:$false
            Write-MbtDetail 'import job of the store: Pending -> InFlight -> Done'
            Write-MbtItem OK ("imported in {0} ({1}): new item id {2}" -f (Format-MbtDuration $result.Seconds), (Format-MbtRate $result.Bytes $result.Seconds), (Short $result.ItemId))
            $afterItems = @(Get-MbtFolderItems -MailboxId $targetId -FolderId $target.Id -PageSize $PageSize)
            $knownBefore = [System.Collections.Generic.HashSet[string]]::new([string[]]@($beforeItems | ForEach-Object Id))
            $new = @($afterItems | Where-Object { -not $knownBefore.Contains($_.Id) })
            $created = @($new | Where-Object Id -EQ $result.ItemId | Select-Object -First 1)
            if (-not $created -and $new.Count -eq 1) { $created = @($new[0]); Write-MbtDetail 'the returned id differs in form from the listed one; the new item is the only one that appeared' }
            Write-MbtItem $(if ($afterItems.Count -eq $beforeItems.Count + 1) { 'OK' } else { 'WARN' }) "target folder: $($beforeItems.Count) -> $($afterItems.Count) items"
            if ($created) {
                $c = $created[0]
                Write-MbtItem OK ("item found in the target folder: type {0}, size {1}, created {2:u}, modified {3:u}" -f $c.Type, (Format-MbtBytes $c.Size), $c.Created, $c.Modified)
            } else { Write-MbtItem WARN "the created item was not found in the listing ($($new.Count) new items appeared)." }
            $roundTrip = $null
            if ($VerifyRoundTrip -and $created) {
                Write-Host ''
                Write-Host '-- Round trip: the created item is exported back and compared with the source' -ForegroundColor White
                $rtDir = Join-Path $OutputPath 'roundtrip-tmp'
                try {
                    $rt = Export-MbtItems -MailboxId $targetId -ItemId @($created[0].Id) -OutputPath $rtDir -FolderId $target.Id -FolderPath $target.Path
                    $back = Read-MbtFxStream -Path (Join-Path $rtDir $rt.Entries[0].File)
                } finally { if (Test-Path -LiteralPath $rtDir) { Remove-Item -LiteralPath $rtDir -Recurse -Force } }
                $s = $source.Summary; $b = $back.Summary
                $checks = [ordered]@{
                    Subject = $s.Subject -ceq $b.Subject; MessageClass = $s.MessageClass -ceq $b.MessageClass
                    Recipients = $s.Recipients.Count -eq $b.Recipients.Count; Attachments = $s.Attachments.Count -eq $b.Attachments.Count
                    AttachmentSizes = (@($s.Attachments | ForEach-Object DataBytes) -join ',') -eq (@($b.Attachments | ForEach-Object DataBytes) -join ',')
                    BodyTextBytes = $s.BodyTextBytes -eq $b.BodyTextBytes; BodyHtmlBytes = $s.BodyHtmlBytes -eq $b.BodyHtmlBytes
                    CreationTime = $s.CreationTime -eq $b.CreationTime; ClientSubmitTime = $s.ClientSubmitTime -eq $b.ClientSubmitTime
                }
                foreach ($k in $checks.Keys) { Write-MbtItem $(if ($checks[$k]) { 'OK' } else { 'WARN' }) ("{0,-17} source / re-exported: {1}" -f $k, $(if ($checks[$k]) { 'same' } else { 'DIFFERENT' })) }
                Write-MbtDetail ("stream size source {0}, re-exported {1}" -f (Format-MbtBytes $source.Read.BytesRead), (Format-MbtBytes $back.Read.BytesRead))
                $roundTrip = $checks
            }
            $report.Steps.Import = [ordered]@{
                StoreItemKey = $chosenItem.Key; SourceItemId = $chosenItem.ItemId; TargetMailbox = $TargetMailbox; TargetFolder = $target.Path; NewItemId = $result.ItemId; ChangeKey = $result.ChangeKey
                Bytes = $result.Bytes; Seconds = [math]::Round($result.Seconds, 3); MiBPerSecond = [math]::Round($result.Bytes / 1MB / [math]::Max(0.001, $result.Seconds), 3)
                ItemsBefore = $beforeItems.Count; ItemsAfter = $afterItems.Count; FoundInListing = [bool]$created; RoundTrip = $roundTrip
            }
        }
    } elseif ($Export -or $SourceFolder) {
        Write-Host ''
        Write-MbtItem INFO 'import not requested (read only). Add -Import -TargetMailbox ... -TargetFolder ... to create the item.'
    }

    # ==== summary ==================================================================================================
    $report.FinishedUtc = [DateTime]::UtcNow.ToString('o')
    $report.GraphRequests = $session.Graph.RequestCount
    $report.Retries = $session.Graph.RetryCount
    $report | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $OutputPath 'phase0-report.json') -Encoding utf8
    Write-MbtStep 'Done'
    Write-MbtItem INFO ("{0} Graph requests, {1} retries (429/503/504)" -f $session.Graph.RequestCount, $session.Graph.RetryCount)
    $st = $store.GetStats()
    Write-MbtDetail ("export store {0}: {1} mailbox(es), {2} folders, {3} items ({4}), {5} in {6} pack(s), {7} previews, {8} import job item(s)" -f $ExportPath, $st.Mailboxes, $st.Folders, $st.Items,
        (($st.ItemsByState | ForEach-Object { "$($_.State) $($_.Count)" }) -join ', '), (Format-MbtBytes $st.StoredBytes), $st.Packs, $st.Previews, $st.ImportJobs)
    Write-MbtDetail "results: $OutputPath  (phase0-report.json, transcript, export\ = the store)"
}
finally {
    Stop-Transcript | Out-Null
    $store.Dispose()
    Disconnect-MbtGraph
}
