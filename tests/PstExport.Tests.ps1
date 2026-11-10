#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the export of a store to PST files (Engine.PstExport.cs, Export-MbtPst, Invoke-MailboxTransfer.ps1 -Mode ExportPst /
    -OutputPst, and -WhatIf in every mode of the entry script).
    Covers: the real folder tree (names, hierarchy, container classes, standard folders), one top folder per location, the merge of the auxiliary
    archives into the Archive tree (or not), the split by size (5 MiB parts, an item larger than a part), the manifest, the resume after a cancelled
    run, a killed process and a damaged part, the selection, -WhatIf, and the memory used. Fake stores and fake services: nothing reaches the network.

    Run:  Invoke-Pester -Path .\tests\PstExport.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-pstexport-tests-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
    $script:Child = Join-Path $PSScriptRoot 'support\pstexport-child.ps1'
    $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:PG = '11111111-2222-3333-4444-555555555555'
    $script:AG = 'aaaaaaaa-0000-0000-0000-00000000000a'
    $script:XG = 'bbbbbbbb-0000-0000-0000-00000000000b'
    $script:P = "MBX:$($script:PG)@$($script:Tenant)"
    $script:A = "MBX:$($script:AG)@$($script:Tenant)"
    $script:X = "MBX:$($script:XG)@$($script:Tenant)"
    $script:Ct = [Threading.CancellationToken]::None

    function Fid([string]$Guid, [long]$N) { [MailboxTransfer.Tests.FakeIds]::Folder($Guid, $N) }

    # A message as the store keeps it: a FastTransfer payload. -Pad bytes of incompressible attachment.
    function New-MsgBytes {
        param([string]$Subject, [string]$Class = 'IPM.Note', [long]$Pad = 0, [int]$Seed = 1)
        $P = [MailboxTransfer.FxProperty]
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($P::FromString(0x001A, $Class))
        $m.Properties.Add($P::FromString(0x0037, $Subject))
        $m.Properties.Add($P::FromString(0x1000, "Body of $Subject"))
        $m.Properties.Add($P::FromInt32(0x0E07, 1))
        $m.Properties.Add($P::FromTime(0x0039, [datetime]::new(2026, 10, 1 + ($Seed % 20), 8, 0, 0, [DateTimeKind]::Utc)))
        if ($Pad -gt 0) {
            $a = [MailboxTransfer.FxAttachment]::new()
            $b = [byte[]]::new($Pad); [Random]::new($Seed).NextBytes($b)
            $a.Properties.Add($P::FromInt32(0x3705, 1)); $a.Properties.Add($P::FromString(0x3707, 'data.bin')); $a.Properties.Add($P::FromBinary(0x3701, $b))
            $m.Attachments.Add($a)
        }
        return , [MailboxTransfer.FxCodec]::EncodePayload($m)
    }

    function Add-Folder {
        param($Store, [string]$Mailbox, [string]$Guid, [long]$N, [string]$Name, [long]$Parent = 0, [string]$Type = 'IPF.Note')
        $f = [MailboxTransfer.StoreFolder]@{ MailboxId = $Mailbox; FolderId = (Fid $Guid $N); ParentFolderId = $(if ($Parent) { Fid $Guid $Parent }); DisplayName = $Name; FolderType = $Type; SourceItemCount = 0 }
        [void]$Store.UpsertFolders([MailboxTransfer.StoreFolder[]]@($f))
    }

    # An exported item of the store; -FolderMailbox: the mailbox whose folder record holds it (an item of an auxiliary archive sits under the folder of the main archive).
    function Add-Item {
        param($Store, [string]$Mailbox, [string]$FolderId, [string]$Subject, [string]$Class = 'IPM.Note', [long]$Pad = 0, [string]$FolderMailbox, [int]$Seed = 1)
        $bytes = New-MsgBytes -Subject $Subject -Class $Class -Pad $Pad -Seed $Seed
        $id = 'I-' + [guid]::NewGuid().ToString('N')
        $e = [MailboxTransfer.StoreItemEntry]@{ MailboxId = $Mailbox; ItemId = $id; FolderId = $FolderId; Size = $bytes.Length; ItemType = $Class }
        if ($FolderMailbox) { $e.FolderMailboxId = $FolderMailbox }
        [void]$Store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@($e))
        $key = $Store.FindItemKey($Mailbox, $id)
        [void]$Store.WriteItem($key, [IO.MemoryStream]::new($bytes), $script:Ct)
        $Store.SetPreview($key, [MailboxTransfer.StorePreview]@{ ItemKey = $key; Subject = $Subject })
        return $key
    }

    <#
        Primary  Inbox (3) > Reports (2) ; Sent Items (1) ; Deleted Items ; Calendar (2, IPF.Appointment) ; Contacts (1, IPF.Contact) ; Projects (no type) > Alpha (2) ; Empty
        Archive  Inbox (2) ; Old mail (2) ; Moved mail (placeholder; its 3 items are in the auxiliary archive) ; Both (1)
        Aux      Old mail (skeleton) ; Moved mail (skeleton: the 3 items sit under the Archive folder) ; Both (skeleton) ; Aux only (its own folder, 1 item)
        Same folder entry key on both sides of the archive (FakeIds: the mailbox part differs).
    #>
    function New-World {
        param([Parameter(Mandatory)][string]$Name, [switch]$NoArchive)
        $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests')
        [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:P; Kind = 'Primary'; Upn = 'adele@contoso.test'; DisplayName = 'Primary' })
        Add-Folder $store $script:P $script:PG 1 'Inbox'
        Add-Folder $store $script:P $script:PG 2 'Reports' 1
        Add-Folder $store $script:P $script:PG 3 'Sent Items'
        Add-Folder $store $script:P $script:PG 4 'Deleted Items'
        Add-Folder $store $script:P $script:PG 5 'Calendar' 0 'IPF.Appointment'
        Add-Folder $store $script:P $script:PG 6 'Contacts' 0 'IPF.Contact'
        Add-Folder $store $script:P $script:PG 7 'Projects' 0 ''
        Add-Folder $store $script:P $script:PG 8 'Alpha' 7
        Add-Folder $store $script:P $script:PG 9 'Empty'
        foreach ($n in 1..3) { [void](Add-Item $store $script:P (Fid $script:PG 1) "Inbox $n" -Seed $n) }
        foreach ($n in 1..2) { [void](Add-Item $store $script:P (Fid $script:PG 2) "Report $n" -Seed $n) }
        [void](Add-Item $store $script:P (Fid $script:PG 3) 'Sent one')
        foreach ($n in 1..2) { [void](Add-Item $store $script:P (Fid $script:PG 5) "Meeting $n" 'IPM.Appointment' -Seed $n) }
        [void](Add-Item $store $script:P (Fid $script:PG 6) 'Carol' 'IPM.Contact')
        foreach ($n in 1..2) { [void](Add-Item $store $script:P (Fid $script:PG 8) "Alpha $n" -Seed $n) }
        if ($NoArchive) { return $store }
        [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:A; Kind = 'Archive'; Upn = 'adele@contoso.test'; DisplayName = 'Archive' })
        [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:X; Kind = 'AuxArchive'; Upn = 'adele@contoso.test'; DisplayName = 'AuxArchive' })
        Add-Folder $store $script:A $script:AG 100 'Inbox'
        Add-Folder $store $script:A $script:AG 101 'Old mail'
        Add-Folder $store $script:A $script:AG 102 'Moved mail'
        Add-Folder $store $script:A $script:AG 103 'Both'
        Add-Folder $store $script:X $script:XG 101 'Old mail'
        Add-Folder $store $script:X $script:XG 102 'Moved mail'
        Add-Folder $store $script:X $script:XG 103 'Both'
        Add-Folder $store $script:X $script:XG 900 'Aux only'
        foreach ($n in 1..2) { [void](Add-Item $store $script:A (Fid $script:AG 100) "Arch inbox $n" -Seed $n) }
        foreach ($n in 1..2) { [void](Add-Item $store $script:A (Fid $script:AG 101) "Old $n" -Seed $n) }
        [void](Add-Item $store $script:A (Fid $script:AG 103) 'Both one')
        foreach ($n in 1..3) { [void](Add-Item $store $script:X (Fid $script:AG 102) "Moved $n" -FolderMailbox $script:A -Seed $n) }
        [void](Add-Item $store $script:X (Fid $script:XG 900) 'Aux own')
        return $store
    }

    # The folders of a PST as path (below the IPM subtree) -> class and subjects of the messages.
    function Get-PstTree {
        param([Parameter(Mandatory)][string]$Path)
        $tree = [ordered]@{}
        $r = [MailboxTransfer.PstReader]::Open($Path)
        try {
            foreach ($f in $r.AllFolders()) {
                $segments = @($f.Path -split '/')
                if ($segments.Count -lt 2 -or $segments[1] -in 'Search Root') { continue }
                $rel = ($segments | Select-Object -Skip 1) -join '/'
                if ($rel -eq 'Deleted Items') { continue }      # the one every PST has
                $subjects = @($r.MessageNids($f, $false) | ForEach-Object { $r.ReadMessage([uint32]$_).GetString(0x0037) } | Sort-Object)
                $tree[$rel] = [pscustomobject]@{ Class = $f.ContainerClass; Subjects = $subjects; Count = $f.ContentCount }
            }
        } finally { $r.Dispose() }
        return $tree
    }

    function New-Options {
        param([Parameter(Mandatory)][string]$Dest, [long]$MiB = 0, [switch]$NoCheck)
        $o = [MailboxTransfer.PstExportOptions]::new()
        $o.Destination = $Dest
        if ($MiB) { $o.MaxPartBytes = $MiB * 1MB }
        $o.Check = -not $NoCheck
        return $o
    }
}

AfterAll {
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'export to PST: the real folder tree' {
    It 'rebuilds names, hierarchy, container classes and items of the primary mailbox, with the standard folders first' {
        $store = New-World -Name 'tree1' -NoArchive
        try {
            $dest = Join-Path $script:Tmp 'tree1\out\mail.pst'
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest), $null, $script:Ct)
            $r.Ok | Should -BeTrue
            $r.Parts.Count | Should -Be 1
            $r.Parts[0].CheckErrors | Should -Be 0
            $r.Parts[0].Checked | Should -BeTrue
            $t = Get-PstTree $dest
            $t['Primary/Inbox'].Subjects | Should -Be @('Inbox 1', 'Inbox 2', 'Inbox 3')
            $t['Primary/Inbox/Reports'].Subjects | Should -Be @('Report 1', 'Report 2')
            $t['Primary/Sent Items'].Count | Should -Be 1
            $t['Primary/Calendar'].Class | Should -Be 'IPF.Appointment'
            $t['Primary/Calendar'].Count | Should -Be 2
            $t['Primary/Contacts'].Class | Should -Be 'IPF.Contact'
            $t['Primary/Projects'].Class | Should -Be 'IPF.Note'
            $t['Primary/Projects/Alpha'].Count | Should -Be 2
            $t['Primary/Empty'].Count | Should -Be 0
            $t['Primary/Deleted Items'].Class | Should -Be 'IPF.Note'
            # the standard folders come first, in the order of Outlook, then the others by name
            $names = @($t.Keys | Where-Object { ($_ -split '/').Count -eq 2 } | ForEach-Object { ($_ -split '/')[1] })
            ($names -join '|') | Should -Be 'Inbox|Sent Items|Deleted Items|Calendar|Contacts|Empty|Projects'
            ($r.Plan.Folders | Where-Object Path -EQ 'Primary/Inbox').WellKnown | Should -Be 'Inbox'
            ($r.Plan.Folders | Where-Object Path -EQ 'Primary/Projects/Alpha').WellKnown | Should -BeNullOrEmpty
        } finally { $store.Dispose() }
    }

    It 'Recoverable Items stay under their own top folder in the PST ("Primary/Recoverable Items/Deletions"), never mixed with the normal folders, and every item is in its folder' {
        # the store as ExportJob.IncludeRecoverableItems writes it: a root without a parent, the exported folders below it
        $store = New-World -Name 'recov1' -NoArchive
        try {
            Add-Folder $store $script:P $script:PG 50 'Recoverable Items'
            Add-Folder $store $script:P $script:PG 51 'Deletions' 50
            Add-Folder $store $script:P $script:PG 52 'Purges' 50
            Add-Folder $store $script:P $script:PG 53 'Versions' 50
            Add-Folder $store $script:P $script:PG 54 'Calendar Logging' 50 'IPF.Appointment'
            foreach ($n in 1..2) { [void](Add-Item $store $script:P (Fid $script:PG 51) "Deleted $n" -Seed $n) }
            [void](Add-Item $store $script:P (Fid $script:PG 52) 'Purged one' -Pad 3000 -Seed 7)
            [void](Add-Item $store $script:P (Fid $script:PG 54) 'Logged meeting' 'IPM.Appointment')
            $dest = Join-Path $script:Tmp 'recov1\out\mail.pst'
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest), $null, $script:Ct)
            $r.Ok | Should -BeTrue
            $r.Parts[0].CheckErrors | Should -Be 0
            $t = Get-PstTree $dest
            $t['Primary/Recoverable Items'].Count | Should -Be 0
            $t['Primary/Recoverable Items/Deletions'].Subjects | Should -Be @('Deleted 1', 'Deleted 2')
            $t['Primary/Recoverable Items/Purges'].Subjects | Should -Be @('Purged one')
            $t['Primary/Recoverable Items/Versions'].Count | Should -Be 0
            $t['Primary/Recoverable Items/Calendar Logging'].Class | Should -Be 'IPF.Appointment'
            $t['Primary/Recoverable Items/Calendar Logging'].Count | Should -Be 1
            $t.Keys | Should -Not -Contain 'Primary/Deletions'
            $t['Primary/Inbox'].Subjects | Should -Be @('Inbox 1', 'Inbox 2', 'Inbox 3') -Because 'the normal folders are untouched'
            $names = @($t.Keys | Where-Object { ($_ -split '/').Count -eq 2 } | ForEach-Object { ($_ -split '/')[1] })
            ($names -join '|') | Should -Be 'Inbox|Sent Items|Deleted Items|Calendar|Contacts|Empty|Projects|Recoverable Items'
            # a selection of one folder of Recoverable Items converts that folder, under its top folder
            $only = New-Options (Join-Path $script:Tmp 'recov1\only\mail.pst')
            $folderKey = ($store.GetFolders($script:P) | Where-Object DisplayName -EQ 'Deletions').Key
            $only.FolderKeys = [System.Collections.Generic.HashSet[long]]::new([long[]]@($folderKey))
            $r2 = [MailboxTransfer.PstExporter]::Run($store, $only, $null, $script:Ct)
            $r2.Ok | Should -BeTrue
            $t2 = Get-PstTree (Join-Path $script:Tmp 'recov1\only\mail.pst')
            $t2['Primary/Recoverable Items/Deletions'].Count | Should -Be 2
            $t2.Keys | Should -Not -Contain 'Primary/Inbox'
        } finally { $store.Dispose() }
    }

    It 'recognises the standard folders by their English, French and German names, and gives the class of the standard folder when the source has none' {
        $rank = 0; $cls = $null
        [MailboxTransfer.PstWellKnown]::Detect("Bo$([char]0xEE)te de r$([char]0xE9)ception", [ref]$rank, [ref]$cls) | Should -Be 'Inbox'
        [MailboxTransfer.PstWellKnown]::Detect("$([char]0xC9)l$([char]0xE9)ments envoy$([char]0xE9)s", [ref]$rank, [ref]$cls) | Should -Be 'Sent Items'
        [MailboxTransfer.PstWellKnown]::Detect('Kalender', [ref]$rank, [ref]$cls) | Should -Be 'Calendar'
        $cls | Should -Be 'IPF.Appointment'
        [MailboxTransfer.PstWellKnown]::Detect('Clients', [ref]$rank, [ref]$cls) | Should -BeNullOrEmpty
        [MailboxTransfer.PstWellKnown]::ContainerClass('', 'IPF.Contact') | Should -Be 'IPF.Contact'
        [MailboxTransfer.PstWellKnown]::ContainerClass('IPF.Task', 'IPF.Contact') | Should -Be 'IPF.Task'
        [MailboxTransfer.PstWellKnown]::ContainerClass($null, $null) | Should -Be 'IPF.Note'
        # a folder with no type whose name is a standard folder (Contacts of a source that gave no class)
        $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp 'wk'), 'wk', 'tests')
        try {
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:P; Kind = 'Primary'; Upn = 'u@x.test'; DisplayName = 'Primary' })
            Add-Folder $store $script:P $script:PG 1 'Contacts' 0 ''
            Add-Folder $store $script:P $script:PG 2 "Bo$([char]0xEE)te de r$([char]0xE9)ception" 0 ''
            [void](Add-Item $store $script:P (Fid $script:PG 1) 'c1' 'IPM.Contact')
            [void](Add-Item $store $script:P (Fid $script:PG 2) 'm1')
            $plan = [MailboxTransfer.PstExporter]::Plan($store, (New-Options (Join-Path $script:Tmp 'wk\o.pst')))
            ($plan.Folders | Where-Object Path -EQ 'Primary/Contacts').ContainerClass | Should -Be 'IPF.Contact'
            ($plan.Folders | Where-Object { $_.Name -like 'Bo*' }).WellKnown | Should -Be 'Inbox'
            $plan.Folders[1].Name | Should -BeLike 'Bo*'     # Inbox first, whatever the language
        } finally { $store.Dispose() }
    }

    It 'makes one top folder per location and merges the auxiliary archive into the Archive tree by their folder entry key (the default)' {
        $store = New-World -Name 'merge'
        try {
            $dest = Join-Path $script:Tmp 'merge\out\mail.pst'
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest), $null, $script:Ct)
            $r.Ok | Should -BeTrue
            @($r.Plan.Locations | ForEach-Object Name) | Should -Be @('Primary', 'Archive')
            $t = Get-PstTree $dest
            @($t.Keys | Where-Object { ($_ -split '/').Count -eq 1 }) | Should -Be @('Primary', 'Archive')
            @($t.Keys | Where-Object { $_ -like 'Archive/*' }) | Should -Be @('Archive/Inbox', 'Archive/Aux only', 'Archive/Both', 'Archive/Moved mail', 'Archive/Old mail')
            $t['Archive/Moved mail'].Subjects | Should -Be @('Moved 1', 'Moved 2', 'Moved 3')
            $t['Archive/Old mail'].Subjects | Should -Be @('Old 1', 'Old 2')
            $t['Archive/Aux only'].Subjects | Should -Be @('Aux own')
            $t['Archive/Inbox'].Subjects | Should -Be @('Arch inbox 1', 'Arch inbox 2')
            (@($t.Keys | Where-Object { $_ -like 'Archive/*Old mail*' })).Count | Should -Be 1      # the skeleton of the auxiliary archive is the same folder
            $r.ItemsWritten | Should -Be 20
        } finally { $store.Dispose() }
    }

    It 'gives each auxiliary archive its own top folder with -SeparateAuxArchives, and the items of an auxiliary archive stay in their own location' {
        $store = New-World -Name 'sep'
        try {
            $dest = Join-Path $script:Tmp 'sep\out\mail.pst'
            $o = New-Options $dest
            $o.MergeAuxIntoArchive = $false
            $r = [MailboxTransfer.PstExporter]::Run($store, $o, $null, $script:Ct)
            $r.Ok | Should -BeTrue
            @($r.Plan.Locations | ForEach-Object Name) | Should -Be @('Primary', 'Archive', 'Auxiliary archive 1 (bbbbbbbb)')
            $t = Get-PstTree $dest
            $t['Archive/Moved mail'].Count | Should -Be 0
            $t['Archive/Old mail'].Subjects | Should -Be @('Old 1', 'Old 2')
            $t['Auxiliary archive 1 (bbbbbbbb)/Moved mail'].Subjects | Should -Be @('Moved 1', 'Moved 2', 'Moved 3')
            $t['Auxiliary archive 1 (bbbbbbbb)/Aux only'].Subjects | Should -Be @('Aux own')
            $t['Auxiliary archive 1 (bbbbbbbb)/Old mail'].Count | Should -Be 0
            $r.ItemsWritten | Should -Be 20
        } finally { $store.Dispose() }
    }

    It 'converts only the mailboxes asked for (kind or id), and the auxiliary archive follows the Archive' {
        $store = New-World -Name 'kinds'
        try {
            $o = New-Options (Join-Path $script:Tmp 'kinds\a.pst') -NoCheck
            $o.Kinds = [MailboxTransfer.StoreMailboxKind[]]@('Primary')
            $p = [MailboxTransfer.PstExporter]::Plan($store, $o)
            @($p.Locations | ForEach-Object Name) | Should -Be @('Primary')
            $p.Items | Should -Be 11
            $o.Kinds = [MailboxTransfer.StoreMailboxKind[]]@('Archive')
            $p = [MailboxTransfer.PstExporter]::Plan($store, $o)
            @($p.Locations | ForEach-Object Name) | Should -Be @('Archive')
            $p.Items | Should -Be 9
            $o.Kinds = [MailboxTransfer.StoreMailboxKind[]]@('AuxArchive')
            $p = [MailboxTransfer.PstExporter]::Plan($store, $o)
            $p.Items | Should -Be 4
            $o.Kinds = $null; $o.MailboxIds = [string[]]@($script:P)
            [MailboxTransfer.PstExporter]::Plan($store, $o).Items | Should -Be 11
        } finally { $store.Dispose() }
    }

    It 'converts a selection of the browse view: a folder with its subfolders, or single items' {
        $store = New-World -Name 'sel'
        try {
            $sel = [MailboxTransfer.BrowseSelection]@{ Kind = 'Primary'; MailboxId = $script:P }
            $sel.Entries.Add([MailboxTransfer.SelectionEntry]@{ FolderId = (Fid $script:PG 1); FolderPath = 'Inbox'; IncludeSubfolders = $true })
            $r = [MailboxTransfer.PstExportSelection]::Resolve($store, $sel)
            $r.EntriesMatched | Should -Be 1
            $dest = Join-Path $script:Tmp 'sel\o.pst'
            $o = New-Options $dest
            $o.FolderKeys = $r.FolderKeys
            $res = [MailboxTransfer.PstExporter]::Run($store, $o, $null, $script:Ct)
            $res.Ok | Should -BeTrue
            $res.ItemsWritten | Should -Be 5
            @((Get-PstTree $dest).Keys) | Should -Be @('Primary', 'Primary/Inbox', 'Primary/Inbox/Reports')
            # items only
            $one = $store.GetItem($script:P, 'x') ; $one | Should -BeNullOrEmpty
            $keys = @($store.GetFolderTree($script:P) | ForEach-Object { $_ })
            $itemKey = (Get-MbtStoreItems -Store $store -MailboxId $script:P | Where-Object { $_.Subject -eq 'Sent one' }).Key
            $sel2 = [MailboxTransfer.BrowseSelection]@{ Kind = 'Primary'; MailboxId = $script:P }
            $itemId = $store.GetItem($itemKey).ItemId
            $sel2.Entries.Add([MailboxTransfer.SelectionEntry]@{ FolderId = (Fid $script:PG 3); FolderPath = 'Sent Items'; Items = [Collections.Generic.List[MailboxTransfer.SelectionItem]]@([MailboxTransfer.SelectionItem]@{ Id = $itemId }) })
            $sel2.Entries.Add([MailboxTransfer.SelectionEntry]@{ FolderId = 'not-there'; FolderPath = 'Nowhere' })
            $r2 = [MailboxTransfer.PstExportSelection]::Resolve($store, $sel2)
            $r2.ItemKeys.Count | Should -Be 1
            $r2.Problems.Count | Should -Be 1
            $o2 = New-Options (Join-Path $script:Tmp 'sel\o2.pst') -NoCheck
            $o2.FolderKeys = $r2.FolderKeys; $o2.ItemKeys = $r2.ItemKeys
            $res2 = [MailboxTransfer.PstExporter]::Run($store, $o2, $null, $script:Ct)
            $res2.ItemsWritten | Should -Be 1
            @((Get-PstTree (Join-Path $script:Tmp 'sel\o2.pst')).Keys) | Should -Be @('Primary', 'Primary/Sent Items')
        } finally { $store.Dispose() }
    }
}

Describe 'export to PST: a store that is not fully exported' {
    It 'exits 2 and says Complete = false in the manifest when items of the store are still Pending or Failed: they are not in the PST' {
        $store = New-World -Name 'partial' -NoArchive
        $pending = [MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:P; ItemId = 'I-pending'; FolderId = (Fid $script:PG 1); Size = 100; ItemType = 'IPM.Note' }
        [void]$store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@($pending))      # listed, never exported
        $store.Dispose()
        $dest = Join-Path $script:Tmp 'partial-out\p.pst'
        $res = Export-MbtPst -Store (Join-Path $script:Tmp 'partial') -Destination $dest 6>$null
        $res.ExitCode | Should -Be 2
        $res.Plan.ItemsNotExported | Should -Be 1
        $res.Result.Complete | Should -BeFalse
        ([MailboxTransfer.PstExportManifest]::Load([MailboxTransfer.PstExportManifest]::PathFor($dest))).Complete | Should -BeFalse
        ($res.Result.Notes -join ' ') | Should -BeLike '*not exported yet*'
    }
}

Describe 'export to PST: split by size' {
    BeforeAll {
        # 4 folders, 14 items of 1.2 MiB each (incompressible), parts of 5 MiB: 3 or 4 items a part
        function New-BigWorld {
            param([string]$Name, [switch]$WithHuge)
            $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests')
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:P; Kind = 'Primary'; Upn = 'adele@contoso.test'; DisplayName = 'Primary' })
            Add-Folder $store $script:P $script:PG 1 'Inbox'
            Add-Folder $store $script:P $script:PG 2 'Reports' 1
            Add-Folder $store $script:P $script:PG 3 'Clients'
            Add-Folder $store $script:P $script:PG 4 'Empty'
            $n = 0
            foreach ($f in 1, 1, 2, 2, 3) { foreach ($i in 1..3) { $n++; [void](Add-Item $store $script:P (Fid $script:PG $f) "Item $n" -Pad 1200000 -Seed $n) } }
            if ($WithHuge) { [void](Add-Item $store $script:P (Fid $script:PG 2) 'HUGE' -Pad 6800000 -Seed 99) }
            return $store
        }
    }

    It 'rolls to a new PST on item boundaries, recreates the folder path in each part, names the files name.pst, name-002.pst... and loses nothing' {
        $store = New-BigWorld -Name 'split1'
        try {
            $dest = Join-Path $script:Tmp 'split1\out\box.pst'
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5), $null, $script:Ct)
            $r.Ok | Should -BeTrue
            $r.Parts.Count | Should -BeGreaterThan 2
            $r.Parts.Count | Should -Be $r.Plan.EstimatedParts
            ($r.Parts | ForEach-Object File) | Should -Be (@('box.pst') + @(2..$r.Parts.Count | ForEach-Object { 'box-{0:D3}.pst' -f $_ }))
            $total = 0
            $seen = @{}
            foreach ($p in $r.Parts) {
                (Get-Item -LiteralPath (Join-Path (Split-Path $dest) $p.File)).Length | Should -Be $p.Bytes
                $p.Bytes | Should -BeLessThan (5MB)
                $p.CheckErrors | Should -Be 0
                $t = Get-PstTree (Join-Path (Split-Path $dest) $p.File)
                foreach ($k in $t.Keys) {
                    foreach ($s in $t[$k].Subjects) { $seen["$s"] = $true; $total++ }
                    # the path of every item is in the part: the ancestors are there too
                    $parent = ($k -split '/' | Select-Object -SkipLast 1) -join '/'
                    if ($parent) { $t.Contains($parent) | Should -BeTrue }
                }
                @($t.Keys | Where-Object { $t[$_].Subjects.Count }).Count | Should -BeGreaterThan 0
            }
            $total | Should -Be 15
            $seen.Count | Should -Be 15
            # the first part holds the whole tree, empty folders included; the next ones only what their items need
            (Get-PstTree $dest).Contains('Primary/Empty') | Should -BeTrue
            (Get-PstTree (Join-Path (Split-Path $dest) $r.Parts[-1].File)).Contains('Primary/Empty') | Should -BeFalse
            $r.ItemsWritten | Should -Be 15
        } finally { $store.Dispose() }
    }

    It 'writes an item larger than the part alone in its own part, and reports it' {
        $store = New-BigWorld -Name 'split2' -WithHuge
        try {
            $dest = Join-Path $script:Tmp 'split2\out\box.pst'
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5), $null, $script:Ct)
            $r.Ok | Should -BeTrue
            $r.Oversize.Count | Should -Be 1
            $r.Oversize[0].Subject | Should -Be 'HUGE'
            $big = $r.Parts | Where-Object Oversize
            @($big).Count | Should -Be 1
            $big.ItemCount | Should -Be 1
            $big.Bytes | Should -BeGreaterThan 5MB
            $big.CheckErrors | Should -Be 0
            $r.Oversize[0].Part | Should -Be $big.Index
            $t = Get-PstTree (Join-Path (Split-Path $dest) $big.File)
            @($t.Keys | Where-Object { $t[$_].Subjects.Count }) | Should -Be @('Primary/Inbox/Reports')
            $t['Primary/Inbox/Reports'].Subjects | Should -Be @('HUGE')
            $r.Plan.OversizeItems.Count | Should -Be 1
            $r.Plan.Notes -join ' ' | Should -Match 'larger than the PST limit'
            ($r.Parts | Measure-Object ItemCount -Sum).Sum | Should -Be 16
            ([MailboxTransfer.PstExportManifest]::Load($r.ManifestPath)).Oversize.Count | Should -Be 1
        } finally { $store.Dispose() }
    }

    It 'refuses a limit under 1 MiB' {
        $store = New-World -Name 'lim' -NoArchive
        try { { [MailboxTransfer.PstExporter]::Run($store, (New-Options (Join-Path $script:Tmp 'lim\o.pst') -MiB 0 | ForEach-Object { $_.MaxPartBytes = 1000; $_ }), $null, $script:Ct) } | Should -Throw '*at least 1 MiB*' }
        finally { $store.Dispose() }
    }
}

Describe 'export to PST: manifest' {
    It 'lists the parts, the folders, the items with the SHA-256 of their source stream, the sizes and the options' {
        $store = New-World -Name 'man1'
        try {
            $dest = Join-Path $script:Tmp 'man1\out\mail.pst'
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -NoCheck), $null, $script:Ct)
            $m = [MailboxTransfer.PstExportManifest]::Load((Join-Path (Split-Path $dest) 'mail.manifest.json'))
            $m.Complete | Should -BeTrue
            $m.MaxPartBytes | Should -Be (50GB)
            $m.MergeAuxIntoArchive | Should -BeTrue
            $m.PlannedItems | Should -Be 20
            @($m.Locations | ForEach-Object Name) | Should -Be @('Primary', 'Archive')
            $m.Parts.Count | Should -Be 1
            $m.Parts[0].File | Should -Be 'mail.pst'
            $m.Parts[0].Bytes | Should -Be (Get-Item -LiteralPath $dest).Length
            $m.Parts[0].Items.Count | Should -Be 20
            ($m.Folders | Where-Object Path -EQ 'Archive/Moved mail').Items | Should -Be 3
            $folderOf = @{}; $m.Folders | ForEach-Object { $folderOf[$_.Index] = $_.Path }
            foreach ($rec in $m.Parts[0].Items) {
                $item = $store.GetItem($rec.Key)
                $rec.Sha256 | Should -Be $item.Sha256
                $rec.Length | Should -Be $item.Length
                $folderOf.ContainsKey($rec.Folder) | Should -BeTrue
            }
            # the SHA of the source stream really is the SHA of the bytes in the pack
            $any = $m.Parts[0].Items[0]
            $s = $store.OpenItem($any.Key)
            try { [BitConverter]::ToString([Security.Cryptography.SHA256]::HashData($s)).Replace('-', '').ToLowerInvariant() | Should -Be $any.Sha256.ToLowerInvariant() } finally { $s.Dispose() }
            Test-Json -Json (Get-Content -LiteralPath $r.ManifestPath -Raw) | Should -BeTrue
        } finally { $store.Dispose() }
    }

    It 'is rewritten after each part and says Complete only at the end' {
        $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp 'man2'), 'man2', 'tests')
        try {
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:P; Kind = 'Primary'; Upn = 'u@x.test'; DisplayName = 'Primary' })
            Add-Folder $store $script:P $script:PG 1 'Inbox'
            foreach ($n in 1..9) { [void](Add-Item $store $script:P (Fid $script:PG 1) "Big $n" -Pad 1200000 -Seed $n) }
            $dest = Join-Path $script:Tmp 'man2\out\mail.pst'
            $seen = [Collections.Generic.List[string]]::new()
            $sink = [MailboxTransfer.PstExportProgressSink]::new([Action[MailboxTransfer.PstExportProgress]]{
                    param($e)
                    if ($e.Phase -eq 'manifest') { $m = [MailboxTransfer.PstExportManifest]::Load((Join-Path (Split-Path $dest) 'mail.manifest.json')); $seen.Add("$($m.Parts.Count):$($m.Complete)") }
                }.GetNewClosure())
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5 -NoCheck), $sink, $script:Ct)
            $r.Ok | Should -BeTrue
            $seen.Count | Should -Be $r.Parts.Count
            $seen[0] | Should -Be '1:False'
            ([MailboxTransfer.PstExportManifest]::Load($r.ManifestPath)).Complete | Should -BeTrue
        } finally { $store.Dispose() }
    }
}

Describe 'export to PST: resume' {
    BeforeAll {
        function New-ResumeWorld {
            param([string]$Name)
            $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests')
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:P; Kind = 'Primary'; Upn = 'u@x.test'; DisplayName = 'Primary' })
            Add-Folder $store $script:P $script:PG 1 'Inbox'
            Add-Folder $store $script:P $script:PG 2 'Archive2025'
            foreach ($n in 1..16) { [void](Add-Item $store $script:P (Fid $script:PG (1 + ($n % 2))) "Item $n" -Pad 1200000 -Seed $n) }
            return $store
        }
        function Get-AllSubjects($dest, $parts) {
            $all = foreach ($p in $parts) { $t = Get-PstTree (Join-Path (Split-Path $dest) $p.File); foreach ($k in $t.Keys) { $t[$k].Subjects } }
            return @($all | Sort-Object)
        }
    }

    It 'continues after a cancelled run: finished parts are kept, the unfinished part is deleted and written again, nothing is written twice' {
        $store = New-ResumeWorld -Name 'res1'
        try {
            $dest = Join-Path $script:Tmp 'res1\out\mail.pst'
            $cts = [Threading.CancellationTokenSource]::new()
            $state = @{ Done = $false }
            $sink = [MailboxTransfer.PstExportProgressSink]::new([Action[MailboxTransfer.PstExportProgress]]{
                    param($e)
                    if ($e.Phase -eq 'write' -and $e.PartIndex -eq 2 -and $e.ItemsDone -ge 5 -and -not $state.Done) { $state.Done = $true; $cts.Cancel() }
                }.GetNewClosure())
            $first = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5), $sink, $cts.Token)
            $first.Cancelled | Should -BeTrue
            $first.Ok | Should -BeFalse
            $first.Parts.Count | Should -Be 1
            Test-Path -LiteralPath (Join-Path (Split-Path $dest) 'mail-002.pst') | Should -BeFalse
            $m = [MailboxTransfer.PstExportManifest]::Load($first.ManifestPath)
            $m.Complete | Should -BeFalse
            $m.Parts.Count | Should -Be 1
            $kept = (Get-Item -LiteralPath $dest).LastWriteTimeUtc
            $second = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5), $null, $script:Ct)
            $second.Ok | Should -BeTrue
            $second.PartsReused | Should -Be 1
            $second.ItemsReused | Should -Be $m.Parts[0].ItemCount
            $second.ItemsWritten + $second.ItemsReused | Should -Be 16
            (Get-Item -LiteralPath $dest).LastWriteTimeUtc | Should -Be $kept
            $all = Get-AllSubjects $dest $second.Parts
            $all.Count | Should -Be 16
            ($all | Select-Object -Unique).Count | Should -Be 16
            ([MailboxTransfer.PstExportManifest]::Load($second.ManifestPath)).Complete | Should -BeTrue
            # a third run has nothing left to do
            $third = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5), $null, $script:Ct)
            $third.Ok | Should -BeTrue
            $third.ItemsWritten | Should -Be 0
            $third.PartsWritten | Should -Be 0
            $third.PartsReused | Should -Be $second.Parts.Count
        } finally { $store.Dispose() }
    }

    It 'continues after the process was killed in the middle of a part' {
        $store = New-ResumeWorld -Name 'res2'
        $store.Dispose()
        $dest = Join-Path $script:Tmp 'res2\out\mail.pst'
        $dir = Split-Path $dest
        $log = Join-Path $script:Tmp 'res2-child.log'
        $q = { param($s) '"' + $s + '"' }
        $proc = Start-Process -FilePath (Get-Process -Id $PID).Path -PassThru -WindowStyle Hidden -RedirectStandardOutput $log -RedirectStandardError "$log.err" -ArgumentList @('-NoProfile', '-File', (& $q $script:Child), '-StorePath', (& $q (Join-Path $script:Tmp 'res2')), '-Dest', (& $q $dest), '-MaxMiB', 5, '-DelayMs', 400, '-SkipCheck')
        try {
            $deadline = [DateTime]::UtcNow.AddSeconds(90)
            # part 1 is in the manifest and part 2 has begun (its file exists): kill
            while ([DateTime]::UtcNow -lt $deadline -and -not $proc.HasExited -and -not ((Test-Path -LiteralPath (Join-Path $dir 'mail.manifest.json')) -and (Test-Path -LiteralPath (Join-Path $dir 'mail-002.pst')))) { Start-Sleep -Milliseconds 50 }
            $proc.HasExited | Should -BeFalse -Because ('the child must still be writing when it is killed: ' + ((Get-Content -LiteralPath $log, "$log.err" -ErrorAction SilentlyContinue) -join ' | '))
            Stop-Process -Id $proc.Id -Force
            $proc.WaitForExit(10000) | Out-Null
        } finally { if (-not $proc.HasExited) { Stop-Process -Id $proc.Id -Force } }
        Test-Path -LiteralPath (Join-Path $dir 'mail-002.pst') | Should -BeTrue      # the part the killed process left
        $store = Open-MbtExportStore -Path (Join-Path $script:Tmp 'res2') -ReadOnly
        try {
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5), $null, $script:Ct)
            $r.Ok | Should -BeTrue
            $r.PartsReused | Should -BeGreaterThan 0
            $all = Get-AllSubjects $dest $r.Parts
            $all.Count | Should -Be 16
            ($all | Select-Object -Unique).Count | Should -Be 16
            foreach ($p in $r.Parts) { $p.CheckErrors | Should -Be 0; (Get-Item -LiteralPath (Join-Path $dir $p.File)).Length | Should -Be $p.Bytes }
            @(Get-ChildItem -LiteralPath $dir -Filter '*.pst').Count | Should -Be $r.Parts.Count
        } finally { $store.Dispose() }
    }

    It 'writes again a finished part that was damaged or lost, and everything after it' {
        $store = New-ResumeWorld -Name 'res3'
        try {
            $dest = Join-Path $script:Tmp 'res3\out\mail.pst'
            $dir = Split-Path $dest
            $first = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5 -NoCheck), $null, $script:Ct)
            $first.Parts.Count | Should -BeGreaterThan 3
            # part 3 cut short, part 2 intact
            $p3 = Join-Path $dir 'mail-003.pst'
            $fs = [IO.File]::Open($p3, 'Open', 'ReadWrite'); $fs.SetLength($fs.Length - 4096); $fs.Dispose()
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5 -NoCheck), $null, $script:Ct)
            $r.Ok | Should -BeTrue
            $r.PartsReused | Should -Be 2
            $r.PartsRebuilt | Should -Be ($first.Parts.Count - 2)
            ($r.Notes -join ' ') | Should -Match 'Part 3 .* is not usable'
            $all = Get-AllSubjects $dest $r.Parts
            ($all | Select-Object -Unique).Count | Should -Be 16
            # a part whose header is garbage
            [IO.File]::WriteAllBytes((Join-Path $dir 'mail-002.pst'), [byte[]](1..200))
            $r2 = [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5 -NoCheck), $null, $script:Ct)
            $r2.PartsReused | Should -Be 1
            (Get-AllSubjects $dest $r2.Parts | Select-Object -Unique).Count | Should -Be 16
        } finally { $store.Dispose() }
    }

    It 'refuses to mix options, to overwrite files it does not know, and starts again with Force' {
        $store = New-ResumeWorld -Name 'res4'
        try {
            $dest = Join-Path $script:Tmp 'res4\out\mail.pst'
            [void][MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 5 -NoCheck), $null, $script:Ct)
            { [MailboxTransfer.PstExporter]::Run($store, (New-Options $dest -MiB 6 -NoCheck), $null, $script:Ct) } | Should -Throw '*other options*'
            $o = New-Options $dest -MiB 6 -NoCheck; $o.Force = $true
            $r = [MailboxTransfer.PstExporter]::Run($store, $o, $null, $script:Ct)
            $r.Ok | Should -BeTrue
            $r.PartsReused | Should -Be 0
            $r.ItemsWritten | Should -Be 16
            @(Get-ChildItem -LiteralPath (Split-Path $dest) -Filter '*.pst').Count | Should -Be $r.Parts.Count      # no part of the first run is left over
            $other = Join-Path $script:Tmp 'res4\other\mail.pst'
            $null = New-Item -ItemType Directory -Path (Split-Path $other)
            [IO.File]::WriteAllBytes($other, [byte[]](1..10))
            { [MailboxTransfer.PstExporter]::Run($store, (New-Options $other -MiB 5 -NoCheck), $null, $script:Ct) } | Should -Throw '*without a manifest*'
        } finally { $store.Dispose() }
    }
}

Describe 'export to PST: the folder of an auxiliary archive (English name, 1.0.x name recognised when an export is continued)' {
    BeforeAll {
        $script:LegacyName = 'Archive auxiliaire 1 (bbbbbbbb)'     # what 1.0.x wrote: the one French literal of the tests, it guards the alias
        # Primary: 4 items; auxiliary archive: 6 items of its own. Parts of 5 MiB hold 4 items: part 1 is the primary, the auxiliary archive follows in part 2 and 3.
        function New-AuxWorld {
            param([string]$Name)
            $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests')
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:P; Kind = 'Primary'; Upn = 'u@x.test'; DisplayName = 'Primary' })
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:X; Kind = 'AuxArchive'; Upn = 'u@x.test'; DisplayName = 'AuxArchive' })
            Add-Folder $store $script:P $script:PG 1 'Inbox'
            Add-Folder $store $script:X $script:XG 900 'Aux only'
            foreach ($n in 1..4) { [void](Add-Item $store $script:P (Fid $script:PG 1) "Primary $n" -Pad 1200000 -Seed $n) }
            foreach ($n in 1..6) { [void](Add-Item $store $script:X (Fid $script:XG 900) "Aux $n" -Pad 1200000 -Seed (10 + $n)) }
            return $store
        }
        function New-AuxOptions([string]$Dest, [switch]$Legacy) {
            $o = New-Options $Dest -MiB 5
            $o.MergeAuxIntoArchive = $false
            if ($Legacy) { $o.AuxArchiveNameFormat = [MailboxTransfer.PstExportOptions]::LegacyAuxArchiveNameFormat }     # what 1.0.x wrote
            return $o
        }
        function Get-PartTops($Dest, $Parts) {
            $tops = foreach ($p in $Parts) { $t = Get-PstTree (Join-Path (Split-Path $Dest) $p.File); foreach ($k in $t.Keys) { ($k -split '/')[0] } }
            return @($tops | Sort-Object -Unique)
        }
        # Runs an export that is cancelled in part 2 (the auxiliary archive is half written), as a 1.0.x or a 1.1 run.
        function Start-CancelledAuxExport($Store, [string]$Dest, [switch]$Legacy) {
            $cts = [Threading.CancellationTokenSource]::new()
            $state = @{ Done = $false }
            $sink = [MailboxTransfer.PstExportProgressSink]::new([Action[MailboxTransfer.PstExportProgress]]{
                    param($e)
                    if ($e.Phase -eq 'write' -and $e.PartIndex -eq 2 -and $e.ItemsDone -ge 6 -and -not $state.Done) { $state.Done = $true; $cts.Cancel() }
                }.GetNewClosure())
            $r = [MailboxTransfer.PstExporter]::Run($Store, (New-AuxOptions $Dest -Legacy:$Legacy), $sink, $cts.Token)
            $r.Cancelled | Should -BeTrue
            $r.Parts.Count | Should -Be 1
            return $r
        }
    }

    It 'names the folder in English with the number and the first 8 characters of the mailbox id, in the plan, the PST and the manifest' {
        $store = New-AuxWorld -Name 'auxname'
        try {
            $dest = Join-Path $script:Tmp 'auxname\out\mail.pst'
            [MailboxTransfer.PstExportOptions]::new().AuxArchiveNameFormat | Should -Be 'Auxiliary archive {0} ({1})'
            @([MailboxTransfer.PstExporter]::Plan($store, (New-AuxOptions $dest)).Locations | ForEach-Object Name) | Should -Be @('Primary', 'Auxiliary archive 1 (bbbbbbbb)')
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-AuxOptions $dest), $null, $script:Ct)
            $r.Ok | Should -BeTrue
            $r.ItemsWritten | Should -Be 10
            @([MailboxTransfer.PstExportManifest]::Load($r.ManifestPath).Locations | ForEach-Object Name) | Should -Be @('Primary', 'Auxiliary archive 1 (bbbbbbbb)')
            Get-PartTops $dest $r.Parts | Should -Be @('Auxiliary archive 1 (bbbbbbbb)', 'Primary')
            $all = foreach ($p in $r.Parts) { $t = Get-PstTree (Join-Path (Split-Path $dest) $p.File); foreach ($k in $t.Keys | Where-Object { $_ -like 'Auxiliary archive 1 (bbbbbbbb)/Aux only' }) { $t[$k].Subjects } }
            @($all).Count | Should -Be 6
        } finally { $store.Dispose() }
    }

    It 'continues an export of 1.0.x under the name its first part already has: no second folder, no duplicate item' {
        $store = New-AuxWorld -Name 'auxold'
        try {
            $dest = Join-Path $script:Tmp 'auxold\out\mail.pst'
            $legacy = $script:LegacyName
            $first = Start-CancelledAuxExport $store $dest -Legacy
            @([MailboxTransfer.PstExportManifest]::Load($first.ManifestPath).Locations | ForEach-Object Name) | Should -Be @('Primary', $legacy)
            # the plan of the continuation (what the console shows) already says the name that will be written
            @([MailboxTransfer.PstExporter]::Plan($store, (New-AuxOptions $dest)).Locations | ForEach-Object Name) | Should -Be @('Primary', $legacy)
            $second = [MailboxTransfer.PstExporter]::Run($store, (New-AuxOptions $dest), $null, $script:Ct)      # 1.1 options: the new name
            $second.Ok | Should -BeTrue
            $second.PartsReused | Should -Be 1
            $second.ItemsReused + $second.ItemsWritten | Should -Be 10
            @([MailboxTransfer.PstExportManifest]::Load($second.ManifestPath).Locations | ForEach-Object Name) | Should -Be @('Primary', $legacy)
            Get-PartTops $dest $second.Parts | Should -Be @($legacy, 'Primary')
            $subjects = foreach ($p in $second.Parts) { $t = Get-PstTree (Join-Path (Split-Path $dest) $p.File); foreach ($k in $t.Keys) { $t[$k].Subjects } }
            @($subjects).Count | Should -Be 10
            @($subjects | Select-Object -Unique).Count | Should -Be 10
            # a third run has nothing left and does not change the name either
            $third = [MailboxTransfer.PstExporter]::Run($store, (New-AuxOptions $dest), $null, $script:Ct)
            $third.ItemsWritten | Should -Be 0
            @([MailboxTransfer.PstExportManifest]::Load($third.ManifestPath).Locations | ForEach-Object Name) | Should -Be @('Primary', $legacy)
        } finally { $store.Dispose() }
    }

    It 'keeps the new name when an export of 1.1 is continued, and uses it for a new export over an old one with Force' {
        $store = New-AuxWorld -Name 'auxnew'
        try {
            $dest = Join-Path $script:Tmp 'auxnew\out\mail.pst'
            $first = Start-CancelledAuxExport $store $dest
            $second = [MailboxTransfer.PstExporter]::Run($store, (New-AuxOptions $dest), $null, $script:Ct)
            $second.Ok | Should -BeTrue
            Get-PartTops $dest $second.Parts | Should -Be @('Auxiliary archive 1 (bbbbbbbb)', 'Primary')
            # an old export started again from zero is a new export: the new name
            $oldDest = Join-Path $script:Tmp 'auxnew\old\mail.pst'
            $null = Start-CancelledAuxExport $store $oldDest -Legacy
            $o = New-AuxOptions $oldDest; $o.Force = $true
            $forced = [MailboxTransfer.PstExporter]::Run($store, $o, $null, $script:Ct)
            $forced.Ok | Should -BeTrue
            $forced.PartsReused | Should -Be 0
            Get-PartTops $oldDest $forced.Parts | Should -Be @('Auxiliary archive 1 (bbbbbbbb)', 'Primary')
        } finally { $store.Dispose() }
    }
}

Describe 'export to PST: the cmdlet and the entry script' {
    It 'Export-MbtPst converts a store given by its path, names the file after the user when given a folder, and skips the check on request' {
        $null = (New-World -Name 'cmd1').Dispose()
        $out = Join-Path $script:Tmp 'cmd1-out'
        $text = Export-MbtPst -Store (Join-Path $script:Tmp 'cmd1') -Destination $out -SkipCheck *>&1 | Out-String
        Test-Path -LiteralPath (Join-Path $out 'adele.pst') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $out 'adele.manifest.json') | Should -BeTrue
        $text | Should -Match 'export to PST finished: 1 part'
        $text | Should -Match 'not checked'
    }

    It 'Export-MbtPst checks each part by default, takes -MaxSizeMiB and -MailboxId, and honours -WhatIf without writing anything' {
        $null = (New-World -Name 'cmd2').Dispose()
        $out = Join-Path $script:Tmp 'cmd2-out'
        $plan = Export-MbtPst -Store (Join-Path $script:Tmp 'cmd2') -Destination (Join-Path $out 'm.pst') -WhatIf 6>&1 | Out-String
        $plan | Should -Match 'plan: 20 item\(s\)'
        $plan | Should -Match 'estimated 1 PST file'
        $plan | Should -Match 'Primary/Inbox'
        $plan | Should -Match 'Archive/Moved mail'
        Test-Path -LiteralPath $out | Should -BeFalse
        $r = Export-MbtPst -Store (Join-Path $script:Tmp 'cmd2') -Destination (Join-Path $out 'm.pst') -Mailbox Archive -MaxSizeMiB 100 6>$null
        $r.ExitCode | Should -Be 0
        $r.Result.Parts[0].Checked | Should -BeTrue
        $r.Result.Parts[0].CheckErrors | Should -Be 0
        $r.Result.ItemsWritten | Should -Be 9
        $r.Result.Plan.MaxPartBytes | Should -Be (100MB)
        @((Get-PstTree (Join-Path $out 'm.pst')).Keys | Where-Object { ($_ -split '/').Count -eq 1 }) | Should -Be @('Archive')
    }

    It 'Export-MbtPst -SelectionFile converts what a selection file names' {
        $s = New-World -Name 'cmd3'
        $sel = [MailboxTransfer.BrowseSelection]@{ Kind = 'Primary'; MailboxId = $script:P }
        $sel.Entries.Add([MailboxTransfer.SelectionEntry]@{ FolderId = (Fid $script:PG 5); FolderPath = 'Calendar' })
        $file = Join-Path $script:Tmp 'cmd3-sel.json'
        $sel.Save($file)
        $s.Dispose()
        $out = Join-Path $script:Tmp 'cmd3-out\sel.pst'
        $r = Export-MbtPst -Store (Join-Path $script:Tmp 'cmd3') -Destination $out -SelectionFile $file -SkipCheck 6>$null
        $r.Result.ItemsWritten | Should -Be 2
        @((Get-PstTree $out).Keys) | Should -Be @('Primary', 'Primary/Calendar')
    }

    It 'the entry script converts a store with -Mode ExportPst, resumes, and -WhatIf prints the plan and writes nothing' {
        $null = (New-World -Name 'ent1').Dispose()
        $store = Join-Path $script:Tmp 'ent1'
        $out = Join-Path $script:Tmp 'ent1-out\mail.pst'
        $global:LASTEXITCODE = $null
        $plan = & $script:Entry -Mode ExportPst -StorePath $store -OutputPst $out -PstMaxSizeMiB 5 -WhatIf *>&1 | Out-String
        $global:LASTEXITCODE | Should -Be 0
        $plan | Should -Match 'plan: 20 item\(s\)'
        $plan | Should -Match 'WhatIf: no PST and no manifest was written'
        Test-Path -LiteralPath (Split-Path $out) | Should -BeFalse
        $text = & $script:Entry -Mode ExportPst -StorePath $store -OutputPst $out *>&1 | Out-String
        $global:LASTEXITCODE | Should -Be 0
        $text | Should -Match 'export to PST finished'
        $text | Should -Match 'check clean'
        Test-Path -LiteralPath $out | Should -BeTrue
        $again = & $script:Entry -Mode ExportPst -StorePath $store -OutputPst $out *>&1 | Out-String
        $global:LASTEXITCODE | Should -Be 0
        $again | Should -Match '1 part\(s\) of an earlier run verified and kept'
        { & $script:Entry -Mode ExportPst -StorePath (Join-Path $script:Tmp 'nostore') -OutputPst $out *>&1 | Out-Null } | Should -Throw '*No store in*'
        { & $script:Entry -Mode Status -StorePath $store -OutputPst $out *>&1 | Out-Null } | Should -Throw '*-OutputPst belongs to*'
    }
}

Describe 'export to PST: file names (Windows device names, Outlook crashes on aux.pst)' {
    BeforeAll {
        # a store with the given kinds of mailbox (one folder with -Items items each)
        function New-NamedStore {
            param([string]$Name, [string]$Upn, [string[]]$Kinds, [long]$Pad = 0, [int]$Items = 2)
            $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests')
            $n = 0
            foreach ($k in $Kinds) {
                $n++
                $guid = '{0:x8}-0000-0000-0000-0000000000{1:x2}' -f (0xcc000000 + $n), $n
                $id = "MBX:$guid@$($script:Tenant)"
                [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $id; Kind = $k; Upn = $Upn; DisplayName = $k })
                Add-Folder $store $id $guid 1 'Inbox'
                foreach ($i in 1..$Items) { [void](Add-Item $store $id (Fid $guid 1) "$k $i" -Pad $Pad -Seed $i) }
            }
            return $store
        }
        function Get-Default {
            param($Store, [string]$Dir = 'C:\out', $Options)
            $a = @{ Destination = $Dir; Store = $Store }
            if ($Options) { $a.Options = $Options }
            Split-Path (Resolve-MbtPstDestination @a 6>$null) -Leaf
        }
    }

    It 'recognises the Windows device names, with any extension and case, and only them' {
        foreach ($n in 'aux', 'AUX.pst', 'Aux.v2.pst', 'con', 'PRN.txt', 'nul.tar.gz', 'COM1', 'com9.pst', 'LPT0', 'lpt9.pst', "COM$([char]0xB9).pst", "LPT$([char]0xB2)", "COM$([char]0xB3)", 'CONIN$', 'CONOUT$.pst', 'aux ', 'aux.', 'aux .pst') {
            [MailboxTransfer.FileNaming]::IsReservedDeviceName($n) | Should -BeTrue -Because $n
        }
        foreach ($n in 'auxiliary', 'aux-archive.pst', 'aux_.pst', 'my.aux.pst', 'com10.pst', 'COM.pst', 'console.pst', 'lpt.pst', 'adele.pst', 'xaux') {
            [MailboxTransfer.FileNaming]::IsReservedDeviceName($n) | Should -BeFalse -Because $n
        }
    }

    It 'makes a name safe: device name, invalid characters, trailing dots and spaces' {
        [MailboxTransfer.FileNaming]::SafeFileName('aux.pst') | Should -Be 'aux_.pst'
        [MailboxTransfer.FileNaming]::SafeFileName('CON.v2.pst') | Should -Be 'CON_.v2.pst'
        [MailboxTransfer.FileNaming]::SafeFileName('nul') | Should -Be 'nul_'
        [MailboxTransfer.FileNaming]::SafeFileName('a|b?c.pst') | Should -Be 'a_b_c.pst'
        [MailboxTransfer.FileNaming]::SafeFileName('mail. . ') | Should -Be 'mail'
        [MailboxTransfer.FileNaming]::SafeFileName('') | Should -Be 'mailbox'
        [MailboxTransfer.FileNaming]::SafeFileName('adele.pst') | Should -Be 'adele.pst'
        [MailboxTransfer.FileNaming]::IsReservedDeviceName([MailboxTransfer.FileNaming]::SafeFileName('aux.pst')) | Should -BeFalse
    }

    It 'refuses a reserved name given by the operator before any work, with the name to use instead' {
        $null = (New-NamedStore 'rn1' 'adele@contoso.test' @('Primary')).Dispose()
        $store = Join-Path $script:Tmp 'rn1'
        $out = Join-Path $script:Tmp 'rn1-out'
        { Export-MbtPst -Store $store -Destination (Join-Path $out 'aux.pst') *>&1 | Out-Null } | Should -Throw "*'aux.pst' cannot be used*AUX*'aux_.pst'*"
        { Export-MbtPst -Store $store -Destination (Join-Path $out 'AUX.PST') -WhatIf *>&1 | Out-Null } | Should -Throw "*'AUX.PST' cannot be used*"
        { Export-MbtPst -Store $store -Destination (Join-Path $out 'con.v2.pst') *>&1 | Out-Null } | Should -Throw "*'con_.v2.pst'*"
        { Export-MbtPst -Store $store -Destination (Join-Path $out 'lpt1.pst') *>&1 | Out-Null } | Should -Throw '*LPT1*'
        { Export-MbtPst -Store $store -Destination (Join-Path $out 'bad|name.pst') *>&1 | Out-Null } | Should -Throw "*'bad_name.pst'*"
        # a folder that would be created with a reserved name
        { Export-MbtPst -Store $store -Destination (Join-Path $out 'nul') *>&1 | Out-Null } | Should -Throw "*'nul' cannot be used*"
        { Export-MbtPst -Store $store -Destination (Join-Path $out 'prn\mail.pst') *>&1 | Out-Null } | Should -Throw "*'prn'*"
        # nothing was created: not the folder, not a manifest
        Test-Path -LiteralPath $out | Should -BeFalse
        # the engine refuses it too (Plan and Run), whoever calls it
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try {
            $o = [MailboxTransfer.PstExportOptions]::new(); $o.Destination = Join-Path $out 'aux.pst'
            { [void][MailboxTransfer.PstExporter]::Plan($s, $o, $script:Ct) } | Should -Throw '*cannot be used*'
            { [void][MailboxTransfer.PstExporter]::Run($s, $o, $null, $script:Ct) } | Should -Throw '*cannot be used*'
        } finally { $s.Dispose() }
        Test-Path -LiteralPath $out | Should -BeFalse
        # the entry script says it before connecting or reading anything
        { & $script:Entry -Mode ExportPst -StorePath $store -OutputPst (Join-Path $out 'aux.pst') *>&1 | Out-Null } | Should -Throw "*'aux.pst' cannot be used*"
        Test-Path -LiteralPath $out | Should -BeFalse
    }

    It 'accepts a name that only looks like a device name' {
        $null = (New-NamedStore 'rn2' 'adele@contoso.test' @('Primary')).Dispose()
        $out = Join-Path $script:Tmp 'rn2-out'
        foreach ($n in 'aux-archive.pst', 'auxiliary.pst', 'com10.pst') {
            $r = Export-MbtPst -Store (Join-Path $script:Tmp 'rn2') -Destination (Join-Path $out $n) -SkipCheck -Force 6>$null
            $r.ExitCode | Should -Be 0
            Test-Path -LiteralPath (Join-Path $out $n) | Should -BeTrue
        }
    }

    It 'makes a name derived from the user safe, and says so' {
        $null = (New-NamedStore 'rn3' 'aux@contoso.test' @('Primary')).Dispose()
        $out = Join-Path $script:Tmp 'rn3-out'
        $text = Export-MbtPst -Store (Join-Path $script:Tmp 'rn3') -Destination $out -SkipCheck *>&1 | Out-String
        $text | Should -Match "derived from the user, 'aux', is a reserved Windows device name.*'aux_.pst'"
        Test-Path -LiteralPath (Join-Path $out 'aux_.pst') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $out 'aux_.manifest.json') | Should -BeTrue
        @(Get-ChildItem -LiteralPath $out -Name) | Should -Not -Contain 'aux.pst'
        foreach ($u in 'nul@x.test', 'COM3@x.test', 'a.b+c@x.test', '@x.test') {
            $s = New-NamedStore ('rn3u' + [guid]::NewGuid().ToString('N').Substring(0, 6)) $u @('Primary')
            try { $leaf = Get-Default $s } finally { $s.Dispose() }
            [MailboxTransfer.FileNaming]::IsReservedDeviceName($leaf) | Should -BeFalse -Because $u
        }
    }

    It 'names the export of an auxiliary archive user-AuxArchiveN (the number of the auxiliary archive) and not after the user alone' {
        $s = New-NamedStore 'rn4' 'adele@contoso.test' @('Primary', 'Archive', 'AuxArchive', 'AuxArchive')
        try {
            Get-Default $s | Should -Be 'adele.pst'                                           # all: not auxiliary archives alone
            $o = [MailboxTransfer.PstExportOptions]::new(); $o.Kinds = [MailboxTransfer.StoreMailboxKind[]]@('AuxArchive')
            Get-Default $s -Options $o | Should -Be 'adele-AuxArchives.pst'                    # the two of them
            $aux = @($s.GetMailboxes() | Where-Object Kind -EQ 'AuxArchive' | Sort-Object Key)
            $o2 = [MailboxTransfer.PstExportOptions]::new(); $o2.MailboxIds = [string[]]@($aux[1].Id)
            Get-Default $s -Options $o2 | Should -Be 'adele-AuxArchive2.pst'                   # the second one
            $o3 = [MailboxTransfer.PstExportOptions]::new(); $o3.Kinds = [MailboxTransfer.StoreMailboxKind[]]@('Archive')
            Get-Default $s -Options $o3 | Should -Be 'adele.pst'                               # Archive takes the auxiliary archives along
        } finally { $s.Dispose() }
        $s2 = New-NamedStore 'rn4b' 'aux@contoso.test' @('AuxArchive')
        try { Get-Default $s2 | Should -Be 'aux-AuxArchive1.pst' } finally { $s2.Dispose() }
        $out = Join-Path $script:Tmp 'rn4-out'
        $r = Export-MbtPst -Store (Join-Path $script:Tmp 'rn4b') -Destination $out -SkipCheck 6>$null
        $r.ExitCode | Should -Be 0
        Test-Path -LiteralPath (Join-Path $out 'aux-AuxArchive1.pst') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $out 'aux-AuxArchive1.manifest.json') | Should -BeTrue
        $r2 = Export-MbtPst -Store (Join-Path $script:Tmp 'rn4b') -Destination (Join-Path $script:Tmp 'rn4-out2') -Mailbox AuxArchive -SkipCheck 6>$null
        Test-Path -LiteralPath (Join-Path $script:Tmp 'rn4-out2\aux-AuxArchive1.pst') | Should -BeTrue
    }

    It 'never names the first part of a split export after a device name: aux_.pst, aux_-002.pst...' {
        $null = (New-NamedStore 'rn5' 'aux@contoso.test' @('Primary') -Pad 1200000 -Items 8).Dispose()
        $out = Join-Path $script:Tmp 'rn5-out'
        $r = Export-MbtPst -Store (Join-Path $script:Tmp 'rn5') -Destination $out -MaxSizeMiB 5 -SkipCheck 6>$null
        $r.ExitCode | Should -Be 0
        $names = @(Get-ChildItem -LiteralPath $out -Filter '*.pst' -Name | Sort-Object)
        $names.Count | Should -BeGreaterThan 1
        $names | Should -Contain 'aux_.pst'
        $names | Should -Contain 'aux_-002.pst'
        foreach ($n in $names) { [MailboxTransfer.FileNaming]::IsReservedDeviceName($n) | Should -BeFalse }
        (Split-Path ([MailboxTransfer.PstExporter]::PartPath((Join-Path $out 'aux_.pst'), 1)) -Leaf) | Should -Be 'aux_.pst'
        (Split-Path ([MailboxTransfer.PstExporter]::PartPath((Join-Path $out 'aux_.pst'), 2)) -Leaf) | Should -Be 'aux_-002.pst'
        # a destination whose head is a device name stays refused whatever the part
        { [MailboxTransfer.FileNaming]::Require((Join-Path $out 'aux.v2.pst'), $true, 'x') } | Should -Throw '*aux_.v2.pst*'
    }

    It 'refuses a reserved name for a repaired PST, a selection file and a store folder' {
        $src = Join-Path $script:Tmp 'rn6-source.pst'; Set-Content -LiteralPath $src 'x'
        { Repair-MbtPst -Path $src -Destination (Join-Path $script:Tmp 'aux.pst') } | Should -Throw "*Repair destination*'aux.pst' cannot be used*"
        { Repair-MbtPst -Path $src -Destination (Join-Path $script:Tmp 'NUL.pst') -WhatIf } | Should -Throw '*cannot be used*'
        $sel = [MailboxTransfer.BrowseSelection]@{ Kind = 'Primary'; MailboxId = $script:P }
        { Save-MbtBrowseSelection -Selection $sel -Path (Join-Path $script:Tmp 'rn6\con.json') } | Should -Throw "*Selection file*'con.json' cannot be used*"
        Test-Path -LiteralPath (Join-Path $script:Tmp 'rn6') | Should -BeFalse
        (Save-MbtBrowseSelection -Selection $sel -Path (Join-Path $script:Tmp 'rn6\con-sel.json')) | Should -Match 'con-sel\.json$'
        { Open-MbtExportStore -Path (Join-Path $script:Tmp 'rn6-store\aux') } | Should -Throw "*Store folder*'aux' cannot be used*"
        Test-Path -LiteralPath (Join-Path $script:Tmp 'rn6-store') | Should -BeFalse
    }
}

Describe 'WhatIf in every mode of the entry script' {
    BeforeAll {
        function Invoke-Entry {
            param([Parameter(Mandatory)]$Handler, [hashtable]$Parameters = @{})
            $global:MbtEntryHandler = $Handler
            $global:LASTEXITCODE = $null
            $text = & $script:Entry @Parameters *>&1 | Out-String
            $global:MbtEntryExit = $LASTEXITCODE
            return $text
        }
        # Inbox (4) > Projects (2), Empty; in both APIs (Mail API: folder sizes; import-export API: the items)
        function New-Service {
            $ie = [MailboxTransfer.Tests.FakeMailboxService]::new()
            $mail = [MailboxTransfer.Tests.FakeMailApi]::new($ie)
            $null = $ie.AddMailbox($script:P); $mail.AddMailbox($script:P)
            $null = $mail.AddFolder($script:P, 'F-INBOX', 'Inbox', $null, 'inbox')
            $null = $mail.AddFolder($script:P, 'F-PROJ', 'Projects', 'F-INBOX')
            $null = $mail.AddFolder($script:P, 'F-EMPTY', 'Empty')
            $null = $ie.AddFolder($script:P, 'F-INBOX', 'Inbox')
            $null = $ie.AddFolder($script:P, 'F-PROJ', 'Projects', 'F-INBOX')
            $null = $ie.AddFolder($script:P, 'F-EMPTY', 'Empty')
            $add = {
                param($folder, $id, $subject)
                $b = [MailboxTransfer.Tests.FxBuilder]::new(); $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $subject)
                $data = $b.ToArray()
                $day = [datetime]::new(2026, 10, 1, 8, 0, 0, [DateTimeKind]::Utc)
                $null = $mail.AddMessage($script:P, $folder, $id, $subject, $day, $data.Length)
                $null = $ie.AddItem($script:P, $folder, $id, $data, "ck-$id", $day)
            }
            foreach ($i in 1..4) { & $add 'F-INBOX' "I$i" "Inbox $i" }
            foreach ($i in 1..2) { & $add 'F-PROJ' "P$i" "Project $i" }
            return [pscustomobject]@{ Mail = $mail; Ie = $ie; Handler = $mail }
        }
    }
    BeforeEach {
        Mock Connect-MbtGraph {
            $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
            $session = Set-MbtSession -Graph ([MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $global:MbtEntryHandler))
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('Mail.Read', 'MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtEntryHandler, MbtEntryExit -Scope Global -ErrorAction SilentlyContinue }

    It 'Export without a selection prints the plan (folders, item counts, announced sizes, PST files) and creates nothing: no store, no export request' {
        $svc = New-Service
        $store = Join-Path $script:Tmp 'wi-store'
        $text = Invoke-Entry $svc.Handler @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:P; WhatIf = $true; PstMaxSizeMiB = 5; OutputPst = (Join-Path $script:Tmp 'wi-out\m.pst') }
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'plan: 1 location\(s\)'
        $text | Should -Match 'Inbox\s+4 item\(s\)'
        $text | Should -Match 'Inbox/Projects\s+2 item\(s\)'
        $text | Should -Match '6 item\(s\)'
        $text | Should -Match 'PST file\(s\) of at most 5[,.]0 MiB'
        $text | Should -Match 'WhatIf: nothing was exported and no store was created'
        $text | Should -Not -Match 'export finished'
        Test-Path -LiteralPath $store | Should -BeFalse
        Test-Path -LiteralPath (Join-Path $script:Tmp 'wi-out') | Should -BeFalse
        @($svc.Ie.RequestsOf('POST') | Where-Object { $_ -match 'exportItems|importItem|createImportSession' }).Count | Should -Be 0 -Because ($svc.Ie.RequestsOf('POST') -join '; ')
        # -Folder narrows the plan
        $one = Invoke-Entry $svc.Handler @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:P; WhatIf = $true; Folder = 'Inbox/Projects' }
        $one | Should -Match 'Inbox/Projects\s+2 item\(s\)'
        $one | Should -Not -Match 'Inbox\s+4 item\(s\)'
        Test-Path -LiteralPath $store | Should -BeFalse
        # the same without -WhatIf really exports and converts (-OutputPst)
        $real = Invoke-Entry $svc.Handler @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:P; TempPath = (Join-Path $script:Tmp 'wi-tmp'); OutputPst = (Join-Path $script:Tmp 'wi-out\m.pst') }
        $global:MbtEntryExit | Should -Be 0
        $real | Should -Match 'export finished: 6 items exported'
        $real | Should -Match 'export to PST finished: 1 part'
        $t = Get-PstTree (Join-Path $script:Tmp 'wi-out\m.pst')
        $t['Primary/Inbox'].Count | Should -Be 4
        $t['Primary/Inbox/Projects'].Count | Should -Be 2
    }

    It 'Export of one folder to a PST writes that folder and the path to it, not the empty folders of the rest of the mailbox' {
        $svc = New-Service
        $store = Join-Path $script:Tmp 'one-store'
        $pst = Join-Path $script:Tmp 'one-out\m.pst'
        $null = Invoke-Entry $svc.Handler @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:P; TempPath = (Join-Path $script:Tmp 'one-tmp'); Folder = 'Inbox/Projects'; OutputPst = $pst }
        $global:MbtEntryExit | Should -Be 0
        $t = Get-PstTree $pst
        $t['Primary/Inbox/Projects'].Count | Should -Be 2
        $t.Contains('Primary/Inbox') | Should -BeTrue -Because 'the path to the folder is kept'
        $t.Contains('Primary/Empty') | Should -BeFalse -Because 'a folder of the mailbox that was not exported is not written'
        # the same store converted again with no selection gives the same tree
        $again = Join-Path $script:Tmp 'one-out2\m.pst'
        $null = Invoke-Entry $svc.Handler @{ Mode = 'ExportPst'; StorePath = $store; OutputPst = $again }
        (Get-PstTree $again).Contains('Primary/Empty') | Should -BeFalse
    }
    It 'Import -WhatIf, Status -ResolveUnknown -WhatIf and Browse -SaveSelection -WhatIf change nothing' {
        $svc = New-Service
        $store = Join-Path $script:Tmp 'wi-store2'
        $null = Invoke-Entry $svc.Handler @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:P; TempPath = (Join-Path $script:Tmp 'wi-tmp2') }
        $global:MbtEntryExit | Should -Be 0
        # an import row left Unknown
        $s = Open-MbtExportStore -Path $store
        try {
            $key = (Get-MbtStoreItems -Store $s | Select-Object -First 1).Key
            $job = $s.CreateImportJob('j1', $script:P)
            [void]$s.AddImportItems($job, [MailboxTransfer.StoreImportPlan[]]@([MailboxTransfer.StoreImportPlan]@{ ItemKey = $key; TargetMailbox = $script:P; TargetFolder = 'F-INBOX' }))
            $row = @($s.NextImportItems($job, 1))[0]
            $s.MarkImportUnknown($row.ImportKey, 'timeout')
        } finally { $s.Dispose() }
        $text = Invoke-Entry $svc.Handler @{ Mode = 'Status'; StorePath = $store; ResolveUnknown = $row.ImportKey; As = 'Imported'; WhatIf = $true }
        $text | Should -Match 'the row was not settled'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try { @($s.GetImportItems($job, 'Unknown', 10)).Count | Should -Be 1 } finally { $s.Dispose() }
        $null = Invoke-Entry $svc.Handler @{ Mode = 'Status'; StorePath = $store; ResolveUnknown = $row.ImportKey; As = 'Imported' }
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try { @($s.GetImportItems($job, 'Unknown', 10)).Count | Should -Be 0 } finally { $s.Dispose() }
        # import: the plan only
        $posts = @($svc.Ie.RequestsOf('POST')).Count
        $imp = Invoke-Entry $svc.Handler @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:P; TargetRoot = 'MBT-WI'; WhatIf = $true }
        $imp | Should -Match 'WhatIf: nothing was created'
        @($svc.Ie.RequestsOf('POST')).Count | Should -Be $posts
    }

    It 'Browse -SaveSelection -WhatIf prints the plan and writes no file' {
        $ie = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $mail = [MailboxTransfer.Tests.FakeMailApi]::new($ie)
        $null = $ie.AddMailbox($script:P); $mail.AddMailbox($script:P)
        $null = $mail.AddFolder($script:P, 'F-INBOX', 'Inbox', $null, 'inbox')
        $null = $ie.AddFolder($script:P, 'F-INBOX', 'Inbox')
        $null = $mail.AddMessage($script:P, 'F-INBOX', 'M1', 'Alpha', [datetime]::new(2026, 10, 1, 8, 0, 0, [DateTimeKind]::Utc), 100)
        $b = [MailboxTransfer.Tests.FxBuilder]::new(); $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, 'Alpha')
        $null = $ie.AddItem($script:P, 'F-INBOX', 'M1', $b.ToArray(), 'ck-M1', [datetime]::new(2026, 10, 1, 8, 0, 0, [DateTimeKind]::Utc))
        $file = Join-Path $script:Tmp 'wi-sel\sel.json'
        $text = Invoke-Entry $mail @{ Mode = 'Browse'; PrimaryMailboxId = $script:P; Folder = 'Inbox'; SaveSelection = $file; WhatIf = $true; StorePath = (Join-Path $script:Tmp 'wi-nostore') }
        $text | Should -Match 'the selection file was not written'
        Test-Path -LiteralPath $file | Should -BeFalse
        $real = Invoke-Entry $mail @{ Mode = 'Browse'; PrimaryMailboxId = $script:P; Folder = 'Inbox'; SaveSelection = $file; StorePath = (Join-Path $script:Tmp 'wi-nostore') }
        $real | Should -Match 'selection written to'
        Test-Path -LiteralPath $file | Should -BeTrue
    }

    It 'Status and Browse of a folder read only: -WhatIf changes nothing there either' {
        $svc = New-Service
        $text = Invoke-Entry $svc.Handler @{ Mode = 'Status'; StorePath = (Join-Path $script:Tmp 'wi-none'); WhatIf = $true }
        $text | Should -Match 'no store in'
        Test-Path -LiteralPath (Join-Path $script:Tmp 'wi-none') | Should -BeFalse
    }
}

Describe 'export to PST: memory' -Tag 'Memory', 'Slow' {
    It 'does not depend on the size of the items: 8 items of 40 MiB are converted in a child process that grows by far less than the data' {
        $dir = Join-Path $script:Tmp 'mem'
        $store = [MailboxTransfer.ExportStore]::Create($dir, 'mem', 'tests')
        try {
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:P; Kind = 'Primary'; Upn = 'u@x.test'; DisplayName = 'Primary' })
            Add-Folder $store $script:P $script:PG 1 'Inbox'
            $P = [MailboxTransfer.FxProperty]
            foreach ($n in 1..8) {
                $m = [MailboxTransfer.FxMessage]::new()
                $m.Properties.Add($P::FromString(0x001A, 'IPM.Note')); $m.Properties.Add($P::FromString(0x0037, "Large $n"))
                $a = [MailboxTransfer.FxAttachment]::new()
                $a.Properties.Add($P::FromInt32(0x3705, 1)); $a.Properties.Add($P::FromString(0x3707, 'big.bin'))
                $a.Properties.Add($P::CreateLazy(0x3701, 0x0102, [MailboxTransfer.TestSupport.PatternBlob]::new(40MB, [uint64]$n)))
                $m.Attachments.Add($a)
                $tmp = Join-Path $script:Tmp "mem-$n.fx"
                $fs = [IO.File]::Create($tmp)
                try { [MailboxTransfer.FxCodec]::EncodePayload($m, $fs) } finally { $fs.Dispose() }
                $id = "I-$n"
                [void]$store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@([MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:P; ItemId = $id; FolderId = (Fid $script:PG 1); Size = (Get-Item $tmp).Length; ItemType = 'IPM.Note' }))
                $in = [IO.File]::OpenRead($tmp)
                try { [void]$store.WriteItem($store.FindItemKey($script:P, $id), $in, $script:Ct) } finally { $in.Dispose() }
                Remove-Item $tmp
            }
        } finally { $store.Dispose() }
        $dest = Join-Path $script:Tmp 'mem-out\mem.pst'
        $text = & (Get-Process -Id $PID).Path -NoProfile -File $script:Child -StorePath $dir -Dest $dest -MaxMiB 200 -SkipCheck
        $line = @($text | Where-Object { $_ -like 'RESULT *' })[0]
        $line | Should -Not -BeNullOrEmpty -Because ($text -join "`n")
        $r = $line.Substring(7) | ConvertFrom-Json
        $r.Complete | Should -BeTrue
        $r.Parts | Should -BeGreaterThan 1
        ($r.PeakMiB - $r.BaselinePeakMiB) | Should -BeLessThan 150 -Because "the peak working set grew from $($r.BaselinePeakMiB) to $($r.PeakMiB) MiB for 320 MiB of items"
    }
}
