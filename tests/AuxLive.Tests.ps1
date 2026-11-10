#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of what the first live auxiliary archive showed (lab tenant, 2026-10-31): a folder being moved to an auxiliary archive
    exists there as a skeleton whose items cannot be queried yet (500 "Cannot query rows in a table"); the same folder has the same folder entry key
    in the id in the main archive and in the auxiliary archive; the import-export API gives no folder size (-1).
    Covers: the decoding of the ids, the classification and bounded retry of the "not now" answers, the browse state ContentNotAvailable, the
    correlation of a folder across the archives, the export (deferred folders, one record per message, resumable exit code), the plan of a selection
    made before a folder moved, and the display of unknown sizes. Nothing here reaches the network (FakeMailboxService).

    Run:  Invoke-Pester -Path .\tests\AuxLive.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-auxlive-tests-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp
    $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:ArchG = 'aaaaaaaa-0000-0000-0000-00000000000a'
    $script:AuxG = 'bbbbbbbb-0000-0000-0000-00000000000b'
    $script:Arch = "MBX:$($script:ArchG)@$($script:Tenant)"
    $script:Aux1 = "MBX:$($script:AuxG)@$($script:Tenant)"
    $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'

    function New-Message([string]$Subject) {
        $b = [MailboxTransfer.Tests.FxBuilder]::new()
        $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $Subject)
        return $b.ToArray()
    }
    function New-FakeClient {
        param([Parameter(Mandatory)]$Service, [double]$DelayScale = 0)
        $o = [MailboxTransfer.GraphOptions]::new()
        $o.DelayScale = $DelayScale
        $o.UseJitter = $false
        return [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
    }
    function New-TestStoreAt([string]$Name) { return [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests') }
    function Fid([string]$Guid, [long]$N) { [MailboxTransfer.Tests.FakeIds]::Folder($Guid, $N) }
    function Mid([string]$Guid, [long]$Folder, [long]$N) { [MailboxTransfer.Tests.FakeIds]::Message($Guid, $Folder, $N) }

    <#
        The archive and its first auxiliary archive in the middle of auto-expansion, with ids in the real layout (same folder entry key on both sides, other
        mailbox part):
          Archive  Inbox (100, 3 items)   Old mail (101, 4 items: being moved)   Moved mail (102, empty placeholder)   Both (103, 2 items)
          Aux      Old mail (101, skeleton: 0 items, or -AuxOldItems of them; listing fails)   Moved mail (102, 3 items)   Both (103, the same 2 messages)
                   Inbox (900, 1 item: the auxiliary archive's OWN folder that happens to have the path of the main archive's Inbox)
    #>
    function New-MoveWorld {
        param([int]$AuxOldItems = 0, [switch]$NoFailure)
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Arch); $null = $svc.AddMailbox($script:Aux1)
        $null = $svc.AddFolder($script:Arch, (Fid $script:ArchG 100), 'Inbox')
        $null = $svc.AddFolder($script:Arch, (Fid $script:ArchG 101), 'Old mail')
        $null = $svc.AddFolder($script:Arch, (Fid $script:ArchG 102), 'Moved mail')
        $null = $svc.AddFolder($script:Arch, (Fid $script:ArchG 103), 'Both')
        $null = $svc.AddFolder($script:Aux1, (Fid $script:AuxG 101), 'Old mail')
        $null = $svc.AddFolder($script:Aux1, (Fid $script:AuxG 102), 'Moved mail')
        $null = $svc.AddFolder($script:Aux1, (Fid $script:AuxG 103), 'Both')
        $null = $svc.AddFolder($script:Aux1, (Fid $script:AuxG 900), 'Inbox')
        foreach ($n in 1..3) { $null = $svc.AddItem($script:Arch, (Fid $script:ArchG 100), (Mid $script:ArchG 100 $n), (New-Message "Inbox $n"), "ck-i$n") }
        foreach ($n in 1..4) { $null = $svc.AddItem($script:Arch, (Fid $script:ArchG 101), (Mid $script:ArchG 101 $n), (New-Message "Old $n"), "ck-o$n") }
        foreach ($n in 1..2) { $null = $svc.AddItem($script:Arch, (Fid $script:ArchG 103), (Mid $script:ArchG 103 $n), (New-Message "Both $n"), "ck-b$n") }
        foreach ($n in 1..3) { $null = $svc.AddItem($script:Aux1, (Fid $script:AuxG 102), (Mid $script:AuxG 102 $n), (New-Message "Moved $n"), "ck-m$n") }
        foreach ($n in 1..2) { $null = $svc.AddItem($script:Aux1, (Fid $script:AuxG 103), (Mid $script:AuxG 103 $n), (New-Message "Both $n"), "ck-b$n") }
        $null = $svc.AddItem($script:Aux1, (Fid $script:AuxG 900), (Mid $script:AuxG 900 1), (New-Message 'Aux own inbox'), 'ck-x1')
        if ($AuxOldItems -gt 0) { foreach ($n in 1..$AuxOldItems) { $null = $svc.AddItem($script:Aux1, (Fid $script:AuxG 101), (Mid $script:AuxG 101 $n), (New-Message "Old $n"), "ck-o$n") } }
        if (-not $NoFailure) { $svc.ListingFailures["$($script:Aux1)|$(Fid $script:AuxG 101)"] = [MailboxTransfer.Tests.FakeListingFailure]::new() }
        return $svc
    }

    function Invoke-ExportJob {
        param([Parameter(Mandatory)]$Service, [Parameter(Mandatory)]$Store, [string]$Mailbox, [string]$Kind, $Plan, [int]$TimeoutSeconds = 60, [int]$Concurrency = 2, [scriptblock]$Options)
        $client = New-FakeClient -Service $Service
        $o = [MailboxTransfer.ExportOptions]::new()
        $o.MailboxId = $Mailbox; $o.Kind = $Kind; $o.Upn = 'adele@contoso.test'; $o.Concurrency = $Concurrency
        $o.TempPath = Join-Path $script:Tmp ('tmp-' + [guid]::NewGuid().ToString('N'))
        if ($Plan) { $o.Plan = $Plan }
        if ($Options) { & $Options $o }
        $job = [MailboxTransfer.ExportJob]::Start([MailboxTransfer.MailboxApi]::new($client), $Store, $o)
        $messages = [System.Collections.Generic.List[string]]::new()
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        while (-not $job.Wait(100)) {
            foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
            if ([DateTime]::UtcNow -gt $deadline) { $job.Cancel(); throw 'the export job did not finish in time' }
        }
        foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
        return [pscustomobject]@{ Job = $job; Messages = $messages; Client = $client; Result = $job.Result }
    }
    function Get-ExportedCount($store, [string]$Mailbox) { (@($store.GetItemStats($Mailbox) | Where-Object State -EQ 'Exported')[0]).Count }
    function Get-Browser($Service) { [MailboxTransfer.MailboxBrowser]::new([MailboxTransfer.MailboxApi]::new((New-FakeClient -Service $Service))) }
    function Get-Loc([string]$Kind) { [MailboxTransfer.MailboxBrowser]::Location($Kind, $(if ($Kind -eq 'Archive') { $script:Arch } else { $script:Aux1 })) }
}

AfterAll {
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Exchange ids: the folder entry key and the message entry key' {
    BeforeAll {
        # ids read on the lab tenant 2026-10-31: the same folders in the main archive (1a2b3c04...) and in its first auxiliary archive (1a2b3c06...)
        $script:LiveArchiveCopy = 'AAMkADFhMmIzYzA0LTVlN2YtNGExYi05YzJkLTNlNGY1YTZiN2M4ZAAuAAAAAAAIunku9_YeRaf6Vg0UW2VeAQAaKzwHXn9KG5wtPk9aa3yNAACZfyuDAAA='
        $script:LiveAuxCopy = 'AAMkADFhMmIzYzA2LTVlN2YtNGExYi05YzJkLTNlNGY1YTZiN2M4ZAAuAAAAAAACUb1kK9FBTLKOOG_77-D5AQAaKzwHXn9KG5wtPk9aa3yNAACZfyuDAAA='
        $script:LiveArchiveClients = 'AAMkADFhMmIzYzA0LTVlN2YtNGExYi05YzJkLTNlNGY1YTZiN2M4ZAAuAAAAAAAIunku9_YeRaf6Vg0UW2VeAQAaKzwHXn9KG5wtPk9aa3yNAACZfyuXAAA='
        $script:LiveAuxClients = 'AAMkADFhMmIzYzA2LTVlN2YtNGExYi05YzJkLTNlNGY1YTZiN2M4ZAAuAAAAAAACUb1kK9FBTLKOOG_77-D5AQAaKzwHXn9KG5wtPk9aa3yNAACZfyuXAAA='
        $script:LiveAuxOwnArchive = 'AAMkADFhMmIzYzA2LTVlN2YtNGExYi05YzJkLTNlNGY1YTZiN2M4ZAAuAAAAAAACUb1kK9FBTLKOOG_77-D5AQCXcroZTqGBQ5sW06XgraM-AAAAAAEaAAA='
    }
    It 'takes apart the live ids: mailbox guid in text, EntryID of a folder; only the mailbox part and the provider uid differ between the two archives' {
        $a = [MailboxTransfer.ExchangeEntryId]::Parse($script:LiveArchiveCopy)
        $x = [MailboxTransfer.ExchangeEntryId]::Parse($script:LiveAuxCopy)
        $a.MailboxGuid | Should -Be '1a2b3c04-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $x.MailboxGuid | Should -Be '1a2b3c06-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $a.IsFolder | Should -BeTrue
        $a.EntryType | Should -Be 1
        $a.ProviderUid | Should -Be '08ba792ef7f61e45a7fa560d145b655e'
        $x.ProviderUid | Should -Be '0251bd642bd1414cb28e386ffbefe0f9'
        $a.FolderKey | Should -Be '1a2b3c075e7f4a1b9c2d3e4f5a6b7c8d-0000997f2b83'
        $x.FolderKey | Should -Be $a.FolderKey
        [MailboxTransfer.ExchangeEntryId]::FolderKeyOf($script:LiveArchiveClients) | Should -Be ([MailboxTransfer.ExchangeEntryId]::FolderKeyOf($script:LiveAuxClients))
        [MailboxTransfer.ExchangeEntryId]::FolderKeyOf($script:LiveArchiveClients) | Should -Not -Be $a.FolderKey
    }
    It 'the auxiliary archive''s own folder (its Archive folder) has another key: same name does not make the same folder' {
        $own = [MailboxTransfer.ExchangeEntryId]::Parse($script:LiveAuxOwnArchive)
        $own.FolderKey | Should -Be '9772ba194ea181439b16d3a5e0ada33e-00000000011a'
        $own.FolderKey | Should -Not -Be ([MailboxTransfer.ExchangeEntryId]::FolderKeyOf($script:LiveArchiveCopy))
    }
    It 'builds and reads back folder and message ids; the folder key of a message id is the key of its folder' {
        $f = Fid $script:ArchG 4242
        $m = Mid $script:ArchG 4242 7
        $pf = [MailboxTransfer.ExchangeEntryId]::Parse($f)
        $pm = [MailboxTransfer.ExchangeEntryId]::Parse($m)
        $pf.MailboxGuid | Should -Be $script:ArchG
        $pf.MessageKey | Should -BeNullOrEmpty
        $pm.EntryType | Should -Be 7
        $pm.FolderKey | Should -Be $pf.FolderKey
        $pm.MessageKey | Should -Not -BeNullOrEmpty
        # the same folder / message in the other mailbox: same keys
        [MailboxTransfer.ExchangeEntryId]::FolderKeyOf((Fid $script:AuxG 4242)) | Should -Be $pf.FolderKey
        [MailboxTransfer.ExchangeEntryId]::MessageKeyOf((Mid $script:AuxG 4242 7)) | Should -Be ([MailboxTransfer.ExchangeEntryId]::MessageKeyOf($m))
        [MailboxTransfer.ExchangeEntryId]::MessageKeyOf((Mid $script:AuxG 4242 8)) | Should -Not -Be ([MailboxTransfer.ExchangeEntryId]::MessageKeyOf($m))
        (Fid $script:ArchG 4242) | Should -Not -Be (Fid $script:AuxG 4242)
    }
    It 'live (2026-10-31): a message keeps its message entry key through the move to the auxiliary archive, only the mailbox part of the id changes (332 of 332 items compared)' {
        # the same message listed in the main archive before the move (2026-10-30) and in the auxiliary archive after it
        $pre = 'AAMkADFhMmIzYzA0LTVlN2YtNGExYi05YzJkLTNlNGY1YTZiN2M4ZABGAAAAAAAIunku9_YeRaf6Vg0UW2VeBwAaKzwHXn9KG5wtPk9aa3yNAACZfyumAAAaKzwHXn9KG5wtPk9aa3yNAACZfzkxAAA='
        $post = 'AAMkADFhMmIzYzA2LTVlN2YtNGExYi05YzJkLTNlNGY1YTZiN2M4ZABGAAAAAAACUb1kK9FBTLKOOG_77-D5BwAaKzwHXn9KG5wtPk9aa3yNAACZfyumAAAaKzwHXn9KG5wtPk9aa3yNAACZfzkxAAA='
        $a = [MailboxTransfer.ExchangeEntryId]::Parse($pre)
        $b = [MailboxTransfer.ExchangeEntryId]::Parse($post)
        $a.IsFolder | Should -BeFalse
        $a.EntryType | Should -Be 7
        $a.MailboxGuid | Should -Be '1a2b3c04-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $b.MailboxGuid | Should -Be '1a2b3c06-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $pre | Should -Not -Be $post
        $a.MessageKey | Should -Be '1a2b3c075e7f4a1b9c2d3e4f5a6b7c8d-0000997f3931'
        $b.MessageKey | Should -Be $a.MessageKey
        $b.FolderKey | Should -Be $a.FolderKey
        [MailboxTransfer.ExchangeEntryId]::MessageKeyOf($post) | Should -Be ([MailboxTransfer.ExchangeEntryId]::MessageKeyOf($pre))
    }
    It 'an id that is not in the layout is not decoded: its key is the id itself (the path decides elsewhere)' {
        foreach ($bad in 'F-INBOX', 'AAMk', '', ('A' * 120), ($script:LiveArchiveCopy.Substring(0, 100))) {
            [MailboxTransfer.ExchangeEntryId]::Parse($bad) | Should -BeNullOrEmpty
            [MailboxTransfer.ExchangeEntryId]::IsDecodable($bad) | Should -BeFalse
            [MailboxTransfer.ExchangeEntryId]::FolderKeyOf($bad) | Should -Be $bad
            [MailboxTransfer.ExchangeEntryId]::MessageKeyOf($bad) | Should -Be $bad
        }
    }
}

Describe 'Service state: "not now" answers of the item listing' {
    It 'classifies the 500 "Cannot query rows in a table" and the documented transient mailbox states, not the other errors' {
        $e500 = [MailboxTransfer.GraphException]::new(500, 'ErrorInternalServerError', 'GET https://x/items -> 500 ErrorInternalServerError: An internal server error occurred. The operation failed., Cannot query rows in a table.', 'https://x/items')
        [MailboxTransfer.MailboxApi]::IsContentNotReady($e500) | Should -BeTrue
        [MailboxTransfer.MailboxApi]::ServiceMessageOf($e500) | Should -Be 'An internal server error occurred. The operation failed., Cannot query rows in a table.'
        foreach ($code in 'ErrorMailboxMoveInProgress', 'ErrorMailboxStoreUnavailable', 'ErrorInternalServerTransientError') {
            [MailboxTransfer.MailboxApi]::IsContentNotReady([MailboxTransfer.GraphException]::new(503, $code, 'GET u -> 503 ' + $code + ': busy', 'u')) | Should -BeTrue -Because $code
        }
        [MailboxTransfer.MailboxApi]::IsContentNotReady([MailboxTransfer.GraphException]::new(500, 'ErrorInternalServerError', 'GET u -> 500 ErrorInternalServerError: something else', 'u')) | Should -BeFalse
        [MailboxTransfer.MailboxApi]::IsContentNotReady([MailboxTransfer.GraphException]::new(404, 'ErrorItemNotFound', 'GET u -> 404 ErrorItemNotFound: not found', 'u')) | Should -BeFalse
        [MailboxTransfer.MailboxApi]::IsContentNotReady([MailboxTransfer.GraphException]::new(403, 'ErrorAccessDenied', 'denied', 'u')) | Should -BeFalse
        [MailboxTransfer.MailboxApi]::IsContentNotReady($null) | Should -BeFalse
    }
    It 'EnumerateItems asks again a bounded number of times: a state that passes is invisible' {
        $svc = New-MoveWorld -AuxOldItems 4
        $failure = $svc.ListingFailures["$($script:Aux1)|$(Fid $script:AuxG 101)"]
        $failure.Left = 2
        $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Service $svc))
        $items = @($api.EnumerateItems($script:Aux1, (Fid $script:AuxG 101), 0))
        $items.Count | Should -Be 4
        $failure.Hits | Should -Be 2
        $svc.RequestsOf("folders/$([uri]::EscapeDataString((Fid $script:AuxG 101)))/items") | Should -Be (2 + 2) -Because 'two refused answers, then two pages of three and one'
    }
    It 'a state that stays: ContentNotAvailableException after ContentAttempts tries, with the message of the service' {
        $svc = New-MoveWorld -AuxOldItems 4
        $failure = $svc.ListingFailures["$($script:Aux1)|$(Fid $script:AuxG 101)"]
        $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Service $svc))
        $err = $null
        try { $null = @($api.EnumerateItems($script:Aux1, (Fid $script:AuxG 101), 0)) } catch { $err = $_.Exception; while ($err -and $err -isnot [MailboxTransfer.ContentNotAvailableException]) { $err = $err.InnerException } }
        $err | Should -Not -BeNullOrEmpty
        $err.Attempts | Should -Be 3
        $err.StatusCode | Should -Be 500
        $err.ServiceMessage | Should -BeLike '*Cannot query rows in a table*'
        $err.MailboxId | Should -Be $script:Aux1
        $failure.Hits | Should -Be 3
        $api.ContentAttempts = 1
        $failure.Hits = 0
        try { $null = @($api.EnumerateItems($script:Aux1, (Fid $script:AuxG 101), 0)) } catch { $err = $_.Exception; while ($err -and $err -isnot [MailboxTransfer.ContentNotAvailableException]) { $err = $err.InnerException } }
        $failure.Hits | Should -Be 1 -Because 'ContentAttempts 1 asks once'
    }
    It 'any other 500 is not retried and stays a GraphException' {
        $svc = New-MoveWorld -AuxOldItems 4
        $failure = $svc.ListingFailures["$($script:Aux1)|$(Fid $script:AuxG 101)"]
        $failure.Message = 'something else broke'
        $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Service $svc))
        { @($api.EnumerateItems($script:Aux1, (Fid $script:AuxG 101), 0)) } | Should -Throw '*something else broke*'
        $failure.Hits | Should -Be 1
    }
}

Describe 'Browse: a folder whose content is not available yet' {
    It 'ListItems answers a state, not an exception; the tree and the other folders go on working' {
        $svc = New-MoveWorld
        $b = Get-Browser $svc
        $loc = Get-Loc AuxArchive
        $tree = $b.GetFolderTree($loc)
        $tree.Folders.Count | Should -Be 4
        $old = $tree.Folders | Where-Object { $_.Path -eq 'Old mail' }
        $page = $b.ListItems($loc, $old)
        $page.State | Should -Be 'ContentNotAvailable'
        $page.Items.Count | Should -Be 0
        $page.StateMessage | Should -BeLike '*Cannot query rows in a table*'
        $page.Notes[0] | Should -BeLike '*not available yet*'
        $page.Notes[0] | Should -BeLike '*still in the main archive*'
        $old.State | Should -Be 'ContentNotAvailable'
        $old.StateMessage | Should -BeLike '*Cannot query rows*'
        # the other folders of the same mailbox
        $moved = $tree.Folders | Where-Object { $_.Path -eq 'Moved mail' }
        $ok = $b.ListItems($loc, $moved)
        $ok.State | Should -Be 'Available'
        $ok.Items.Count | Should -Be 3
        $moved.State | Should -Be 'Available'
        # the failure is only a state of that moment: it passes, the folder reads again
        $svc.ListingFailures.Clear()
        $null = $svc.AddItem($script:Aux1, (Fid $script:AuxG 101), (Mid $script:AuxG 101 1), (New-Message 'Old 1'), 'ck-o1')
        $again = $b.ListItems($loc, $old)
        $again.State | Should -Be 'Available'
        $again.Items.Count | Should -Be 1
        $old.State | Should -Be 'Available'
        $old.StateMessage | Should -BeNullOrEmpty
    }
    It 'Get-MbtBrowseItem warns with the note instead of throwing' {
        $svc = New-MoveWorld
        $global:MbtAuxClient = New-FakeClient -Service $svc
        $null = Set-MbtSession -Graph $global:MbtAuxClient
        try {
            $loc = New-MbtBrowseLocation -Kind AuxArchive -MailboxId $script:Aux1
            $folder = (Get-MbtBrowseFolder -Location $loc -AsTree).Folders | Where-Object Path -EQ 'Old mail'
            $rows = @(Get-MbtBrowseItem -Location $loc -Folder $folder -WarningVariable w -WarningAction SilentlyContinue)
            $rows.Count | Should -Be 0
            ($w -join ' ') | Should -BeLike '*not available yet*'
            $folder.State | Should -Be 'ContentNotAvailable'
        } finally { Remove-Variable -Name MbtAuxClient -Scope Global -ErrorAction SilentlyContinue }
    }
}

Describe 'Browse: the Mail API of the main archive on a folder whose content moved' {
    It 'answers 500 "Cannot query rows in a table" (live): the view falls back to the import-export listing, which follows the redirect to the auxiliary archive' {
        $ie = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $mail = [MailboxTransfer.Tests.FakeMailApi]::new($ie)
        $null = $ie.AddMailbox($script:Arch); $null = $ie.AddMailbox($script:Aux1); $mail.AddMailbox($script:Arch)
        $null = $mail.AddFolder($script:Arch, 'F-MOVED', 'Old mail')
        $null = $ie.AddFolder($script:Arch, 'F-MOVED', 'Old mail'); $null = $ie.AddFolder($script:Aux1, 'F-MOVED', 'Old mail')
        foreach ($i in 1..2) { $null = $ie.AddItem($script:Aux1, 'F-MOVED', "M$i", (New-Message "Old $i"), "ck$i") }
        $ie.MovedFolderToMailbox['F-MOVED'] = $script:Aux1; $ie.RedirectFolders = $true
        $mail.MessagesNotReady.Add("$($script:Arch)|F-MOVED") | Out-Null
        $b = [MailboxTransfer.MailboxBrowser]::new([MailboxTransfer.MailboxApi]::new((New-FakeClient -Service $mail)))
        $loc = Get-Loc Archive
        $folder = $b.GetFolderTree($loc).Folders | Where-Object Path -EQ 'Old mail'
        $page = $b.ListItems($loc, $folder)
        $page.Source | Should -Be 'ImportExportApi'
        $page.Items.Count | Should -Be 2
        $page.State | Should -Be 'Available'
        $page.Notes[0] | Should -BeLike '*Mail API cannot read this folder now*Cannot query rows*moved to an auxiliary archive*'
    }
}

Describe 'Browse: an unknown size is unknown, never summed' {
    It 'the import-export API gives no folder size (-1): the totals count those folders apart' {
        $svc = New-MoveWorld
        $tree = (Get-Browser $svc).GetFolderTree((Get-Loc AuxArchive))
        ($tree.Folders | Where-Object SizeKnown).Count | Should -Be 0
        $t = $tree.Totals()
        $t.Folders | Should -Be 4
        $t.Items | Should -Be 6
        $t.KnownSizeBytes | Should -Be 0
        $t.FoldersWithUnknownSize | Should -Be 4
        $t.SizeUnknown | Should -BeTrue
        $t.ToString() | Should -Be '4 folder(s), 6 item(s), size unknown'
    }
    It 'a mix: the known sizes are added, the unknown ones are counted apart' {
        $folders = @(
            [MailboxTransfer.BrowseFolder]@{ TotalItemCount = 3; SizeBytes = 1024 }
            [MailboxTransfer.BrowseFolder]@{ TotalItemCount = 2; SizeBytes = 2048 }
            [MailboxTransfer.BrowseFolder]@{ TotalItemCount = 5; SizeBytes = -1 }
            [MailboxTransfer.BrowseFolder]@{ TotalItemCount = 0; SizeBytes = 0 })
        $t = [MailboxTransfer.BrowseTotals]::Of([MailboxTransfer.BrowseFolder[]]$folders)
        $t.KnownSizeBytes | Should -Be 3072
        $t.FoldersWithUnknownSize | Should -Be 1
        $t.SizeComplete | Should -BeFalse
        $t.SizeUnknown | Should -BeFalse
        $t.ToString() | Should -Be '4 folder(s), 10 item(s), 3.0 KiB (1 folder with unknown size, not counted)'
        [MailboxTransfer.BrowseTotals]::Of([MailboxTransfer.BrowseFolder[]]@($folders[0], $folders[1])).SizeComplete | Should -BeTrue
    }
    It 'Show-MbtBrowseTree prints "unknown" for the size and ends with the total line; a selection''s total does not add -1 either' {
        $svc = New-MoveWorld
        $tree = (Get-Browser $svc).GetFolderTree((Get-Loc AuxArchive))
        $text = Show-MbtBrowseTree -Tree $tree *>&1 | Out-String
        $text | Should -Match 'Moved mail\s+3 item\(s\)\s+unknown'
        $text | Should -Not -Match '-0|-1 B'
        $text | Should -Match 'total: 4 folder\(s\), 6 item\(s\), size unknown'
        (Format-MbtBrowseTotals -Totals $tree.Totals()) | Should -Be '4 folder(s), 6 item(s), size unknown'
        $sel = [MailboxTransfer.BrowseSelection]::For((Get-Loc AuxArchive))
        foreach ($f in $tree.Folders) { $null = $sel.AddFolder($f) }
        $st = [MailboxTransfer.BrowseTotals]::Of($sel)
        $st.KnownSizeBytes | Should -Be 0
        $st.SizeUnknown | Should -BeTrue
    }
}

Describe 'Browse: the same folder in the main archive and in an auxiliary archive' {
    BeforeAll {
        function Get-Trees($svc) {
            $b = Get-Browser $svc
            $client = $b.Graph
            $archive = $b.GetFolderTree((Get-Loc Archive))     # the Mail API of the fake refuses nothing here: the fake service is the only API, so this is the import-export tree
            $aux = $b.GetFolderTree((Get-Loc AuxArchive))
            return @($archive, $aux)
        }
    }
    It 'finds it by the folder entry key and says where the items are' {
        $svc = New-MoveWorld
        # the Mail API is not part of this fake: an Archive location falls back to the import-export tree on the refusal it gets (404)
        $trees = Get-Trees $svc
        $groups = @([MailboxTransfer.FolderCorrelation]::Annotate([MailboxTransfer.BrowseTree[]]$trees))
        $groups.Count | Should -Be 3 -Because 'Old mail, Moved mail and Both exist on both sides; Inbox (100) and the auxiliary archive''s own Inbox (900) are different folders'
        $by = @{}; foreach ($g in $groups) { $by[$g.Path] = $g }
        $by['Old mail'].MatchedBy | Should -Be 'EntryId'
        $by['Old mail'].State | Should -Be 'BeingMoved'
        $by['Old mail'].Text | Should -Be 'being moved to aux archive 1 (bbbbbbbb) (the items are still in the main archive)'
        $by['Moved mail'].State | Should -Be 'Moved'
        $by['Moved mail'].Text | Should -BeLike 'moved to aux archive 1 (bbbbbbbb)*'
        $by['Both'].State | Should -Be 'BothSides'
        $by['Both'].Text | Should -BeLike 'items on both sides: main archive 2, aux archive 1 (bbbbbbbb) 2*'
        $oldArch = $trees[0].Folders | Where-Object Path -EQ 'Old mail'
        $oldArch.MoveState | Should -Be 'BeingMoved'
        $oldArch.MoveText | Should -Be 'being moved to aux archive 1 (bbbbbbbb)'
        ($trees[1].Folders | Where-Object Path -EQ 'Old mail').MoveText | Should -BeLike 'being moved here from the main archive*'
        ($trees[1].Folders | Where-Object Path -EQ 'Moved mail').MoveText | Should -Be 'moved here from the main archive'
        # same path, different folder entry key: not the same folder
        ($trees[0].Folders | Where-Object Path -EQ 'Inbox').MoveState | Should -Be 'None'
        ($trees[1].Folders | Where-Object Path -EQ 'Inbox').MoveState | Should -Be 'None'
    }
    It 'Show-MbtBrowseTree shows the move next to the folder (Add-MbtBrowseMoveInfo is the module function)' {
        $svc = New-MoveWorld
        $trees = Get-Trees $svc
        $groups = @(Add-MbtBrowseMoveInfo -Tree $trees)
        $groups.Count | Should -Be 3
        $text = Show-MbtBrowseTree -Tree $trees[0] *>&1 | Out-String
        $text | Should -Match 'Old mail\s+4 item\(s\).*\[being moved to aux archive 1 \(bbbbbbbb\)\]'
        $text | Should -Match 'Moved mail\s+0 item\(s\).*\[moved to aux archive 1'
        $aux = Show-MbtBrowseTree -Tree $trees[1] *>&1 | Out-String
        $aux | Should -Match 'Moved mail\s+3 item\(s\).*\[moved here from the main archive\]'
    }
    It 'falls back to the path only for ids that are not in the Exchange layout, and never joins two folders of one mailbox' {
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Arch); $null = $svc.AddMailbox($script:Aux1)
        $null = $svc.AddFolder($script:Arch, 'F-MOVED', 'Old mail')
        $null = $svc.AddFolder($script:Aux1, 'X-MOVED', 'Old mail')
        $null = $svc.AddFolder($script:Arch, 'F-TWIN-1', 'Twin'); $null = $svc.AddFolder($script:Arch, 'F-TWIN-2', 'Twin')
        $null = $svc.AddFolder($script:Aux1, 'X-TWIN', 'Twin')
        $null = $svc.AddItem($script:Aux1, 'X-MOVED', 'M1', (New-Message 'm1'))
        $trees = Get-Trees $svc
        $groups = @([MailboxTransfer.FolderCorrelation]::Annotate([MailboxTransfer.BrowseTree[]]$trees))
        $groups.Count | Should -Be 1 -Because 'Twin has two folders of that path in the main archive: ambiguous, not joined'
        $groups[0].Path | Should -Be 'Old mail'
        $groups[0].MatchedBy | Should -Be 'Path'
        $groups[0].State | Should -Be 'Moved'
    }
}

Describe 'Export: a folder whose content is not available yet is deferred, not empty, and the others go on' {
    It 'an auxiliary archive export defers the folder with the message of the service, exports the rest, and the second run finishes it' {
        $svc = New-MoveWorld -AuxOldItems 4
        $store = New-TestStoreAt 'dx1'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Completed | Should -BeTrue
            $r.Result.Exported | Should -Be (3 + 2 + 1) -Because 'Moved mail, Both and the own Inbox; Old mail (4 items) is deferred'
            $r.Result.FoldersDeferred.Count | Should -Be 1
            $d = $r.Result.FoldersDeferred[0]
            $d.Path | Should -Be 'Old mail'
            $d.MailboxId | Should -Be $script:Aux1
            $d.Message | Should -BeLike '*Cannot query rows in a table*'
            $d.StatusCode | Should -Be 500
            $d.Attempts | Should -Be 3
            $r.Result.NeedsAnotherRun | Should -BeTrue
            ($r.Result.Problems -join "`n") | Should -BeLike "*folder 'Old mail' in $($script:Aux1): content not available yet*"
            $r.Result.FoldersNotListed | Should -Be 0 -Because 'a deferred folder is not a failed one'
            # nothing of that folder is recorded as exported or empty
            $svc.ListingFailures["$($script:Aux1)|$(Fid $script:AuxG 101)"].Hits | Should -Be 3
            # the state passes: the same command does what is left
            $svc.ListingFailures.Clear()
            $again = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $again.Result.FoldersDeferred.Count | Should -Be 0
            $again.Result.Exported | Should -Be 4
            $again.Result.NeedsAnotherRun | Should -BeFalse
            Get-ExportedCount $store $script:Aux1 | Should -Be 10
        } finally { $store.Dispose() }
    }
    It 'a state that passes during the retries is invisible to the export' {
        $svc = New-MoveWorld -AuxOldItems 4
        $svc.ListingFailures["$($script:Aux1)|$(Fid $script:AuxG 101)"].Left = 2
        $store = New-TestStoreAt 'dx2'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $r.Result.FoldersDeferred.Count | Should -Be 0
            $r.Result.Exported | Should -Be 10
            $r.Result.NeedsAnotherRun | Should -BeFalse
        } finally { $store.Dispose() }
    }
    It 'the main archive export follows an expanded folder to the auxiliary archive: a listing that refuses there defers it, the other folders are exported' {
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Arch); $null = $svc.AddMailbox($script:Aux1)
        $null = $svc.AddFolder($script:Arch, 'F-INBOX', 'Inbox')
        $null = $svc.AddFolder($script:Arch, 'F-MOVED', 'Old mail')
        $null = $svc.AddFolder($script:Aux1, 'F-MOVED', 'Old mail')
        foreach ($i in 1..3) { $null = $svc.AddItem($script:Arch, 'F-INBOX', "A$i", (New-Message "Archive $i"), "ck-A$i") }
        foreach ($i in 1..4) { $null = $svc.AddItem($script:Aux1, 'F-MOVED', "M$i", (New-Message "Old $i"), "ck-M$i") }
        $svc.MovedFolderToMailbox['F-MOVED'] = $script:Aux1
        $svc.RedirectFolders = $true
        $svc.ListingFailures["$($script:Aux1)|F-MOVED"] = [MailboxTransfer.Tests.FakeListingFailure]::new()
        $store = New-TestStoreAt 'dx3'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $r.Result.Completed | Should -BeTrue
            $r.Result.Exported | Should -Be 3
            $r.Result.FoldersDeferred.Count | Should -Be 1
            $r.Result.FoldersDeferred[0].Path | Should -Be 'Old mail'
            $r.Result.NeedsAnotherRun | Should -BeTrue
            $svc.ListingFailures.Clear()
            $again = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $again.Result.Exported | Should -Be 4
            $again.Result.NeedsAnotherRun | Should -BeFalse
            $store.GetStats().Items | Should -Be 7
        } finally { $store.Dispose() }
    }
}

Describe 'Export: a location that keeps refusing is asked once per folder' {
    BeforeAll {
        # F01..F09, one item each; the listing of F01-F04 and F06-F09 refuses, F05 lists
        function New-RefusingService {
            $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
            $null = $svc.AddMailbox($script:Aux1)
            foreach ($n in 1..9) {
                $id = 'F{0:00}' -f $n
                $null = $svc.AddFolder($script:Aux1, $id, "Folder $id")
                $null = $svc.AddItem($script:Aux1, $id, "I$id", (New-Message "Item $id"), "ck$n")
                if ($n -ne 5) { $svc.ListingFailures["$($script:Aux1)|$id"] = [MailboxTransfer.Tests.FakeListingFailure]::new() }
            }
            return $svc
        }
        function Get-Hits($svc) { 1..9 | ForEach-Object { $k = 'F{0:00}' -f $_; if ($svc.ListingFailures.ContainsKey("$($script:Aux1)|$k")) { $svc.ListingFailures["$($script:Aux1)|$k"].Hits } else { 0 } } }
    }
    It 'after NotReadyTrip folders in a row the next ones are asked once; a folder that lists puts the full asks back' {
        $svc = New-RefusingService
        $store = New-TestStoreAt 'fd1'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive' -Concurrency 1
            @(Get-Hits $svc) | Should -Be @(3, 3, 3, 1, 0, 3, 3, 3, 1)
            $r.Result.FoldersDeferred.Count | Should -Be 8
            $r.Result.Exported | Should -Be 1
            $r.Result.FoldersDeferredFast | Should -Be 2
            @($r.Result.FoldersDeferred | Where-Object Attempts -EQ 1).Count | Should -Be 2
            @($r.Result.FoldersDeferred | Where-Object Attempts -EQ 3).Count | Should -Be 6
            @($r.Messages | Where-Object { $_ -like 'the last 3 folders of this location were not readable yet*asked once*' }).Count | Should -Be 2
            $r.Result.NeedsAnotherRun | Should -BeTrue
        } finally { $store.Dispose() }
    }
    It 'NotReadyTrip is configurable: 0 always asks in full, 1 asks once after the first refusal' {
        $svc = New-RefusingService
        $store = New-TestStoreAt 'fd2'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive' -Concurrency 1 -Options { param($o) $o.NotReadyTrip = 0 }
            @(Get-Hits $svc) | Should -Be @(3, 3, 3, 3, 0, 3, 3, 3, 3)
            $r.Result.FoldersDeferredFast | Should -Be 0
        } finally { $store.Dispose() }
        $svc = New-RefusingService
        $store = New-TestStoreAt 'fd3'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive' -Concurrency 1 -Options { param($o) $o.NotReadyTrip = 1 }
            @(Get-Hits $svc) | Should -Be @(3, 1, 1, 1, 0, 3, 1, 1, 1)
        } finally { $store.Dispose() }
    }
    It 'the deferred folders are exported by the next run once they list' {
        $svc = New-RefusingService
        $store = New-TestStoreAt 'fd4'
        try {
            $null = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive' -Concurrency 1
            $svc.ListingFailures.Clear()
            $again = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive' -Concurrency 1
            $again.Result.FoldersDeferred.Count | Should -Be 0
            $again.Result.Exported | Should -Be 8
            Get-ExportedCount $store $script:Aux1 | Should -Be 9
        } finally { $store.Dispose() }
    }
    It 'the browse is not affected: a folder is still asked in full' {
        $svc = New-RefusingService
        $b = Get-Browser $svc
        $loc = Get-Loc AuxArchive
        $f = $b.GetFolderTree($loc).Folders | Where-Object Path -EQ 'Folder F01'
        $page = $b.ListItems($loc, $f)
        $page.State | Should -Be 'ContentNotAvailable'
        $svc.ListingFailures["$($script:Aux1)|F01"].Hits | Should -Be 3
    }
}

Describe 'Export: one record per message, wherever the move has got to' {
    It 'archive first, then the auxiliary archive: the folders are recognised by their key, the messages seen on both sides are recorded once' {
        $svc = New-MoveWorld
        $store = New-TestStoreAt 'ox1'
        try {
            $a = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $a.Result.Exported | Should -Be 9
            $x = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $x.Result.FoldersDeferred.Count | Should -Be 0 -Because 'the skeleton of Old mail announces no item: nothing to list, nothing to defer'
            $x.Result.Exported | Should -Be 4 -Because 'Moved mail 3 and the auxiliary archive''s own Inbox 1; the two messages of Both were exported from the main archive'
            $x.Result.ItemsDeduplicated | Should -Be 2
            @($x.Result.MovesInProgress | Where-Object { $_ -like "*'Old mail' is being moved to this auxiliary archive: 4 item(s)*" }).Count | Should -Be 1
            @($x.Result.MovesInProgress | Where-Object { $_ -like "*'Both' has items on both sides*" }).Count | Should -Be 1
            $store.GetStats().Items | Should -Be 13
            # no item of the main archive was asked of the auxiliary archive, nor the other way
            (@($svc.ExportCalls | Where-Object Mailbox -EQ $script:Aux1 | ForEach-Object Items) | Measure-Object -Sum).Sum | Should -Be 4
            (@($svc.ExportCalls | Where-Object Mailbox -EQ $script:Arch | ForEach-Object Items) | Measure-Object -Sum).Sum | Should -Be 9
            # the items of Moved mail are recorded under the folder row of the main archive tree
            $moved = $store.GetFolders($script:Arch) | Where-Object DisplayName -EQ 'Moved mail'
            $moved | Should -Not -BeNullOrEmpty
            $store.GetItem($script:Aux1, (Mid $script:AuxG 102 1)).FolderKey | Should -Be $moved.Key
            @($store.GetFolders($script:Aux1) | ForEach-Object DisplayName) | Should -Be @('Inbox') -Because 'only the folder that is not in the main archive tree has a row of its own'
            # a third run: nothing changes
            $again = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $again.Result.Exported | Should -Be 0
            $store.GetStats().Items | Should -Be 13
        } finally { $store.Dispose() }
    }
    It 'the auxiliary archive first, then the archive: still one record per message' {
        $svc = New-MoveWorld
        $store = New-TestStoreAt 'ox2'
        try {
            $x = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $x.Result.Exported | Should -Be 6
            $a = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $a.Result.ItemsDeduplicated | Should -Be 2
            $a.Result.Exported | Should -Be 7 -Because 'Inbox 3 + Old mail 4; the two messages of Both are already exported from the auxiliary archive'
            $store.GetStats().Items | Should -Be 13
        } finally { $store.Dispose() }
    }
    It 'a record of an item not exported yet follows the message to the mailbox that lists it now (new id), and is exported from there' {
        $svc = New-MoveWorld
        $store = New-TestStoreAt 'ox3'
        try {
            # the main archive was inventoried (items Pending), the export did not run; then Both moved
            $null = $store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:Arch; Kind = 'Archive' })
            $null = $store.UpsertFolders([MailboxTransfer.StoreFolder[]]@([MailboxTransfer.StoreFolder]@{ MailboxId = $script:Arch; FolderId = (Fid $script:ArchG 103); DisplayName = 'Both' }))
            $null = $store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@(
                [MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:Arch; ItemId = (Mid $script:ArchG 103 1); FolderId = (Fid $script:ArchG 103); ChangeKey = 'ck-b1'; Size = 10 }
                [MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:Arch; ItemId = (Mid $script:ArchG 103 2); FolderId = (Fid $script:ArchG 103); ChangeKey = 'ck-b2'; Size = 10 }))
            $x = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $x.Result.ItemsRehomed | Should -Be 2
            $x.Result.ItemsDeduplicated | Should -Be 0
            $store.FindItemKey($script:Arch, (Mid $script:ArchG 103 1)) | Should -BeNullOrEmpty
            $store.GetItem($script:Aux1, (Mid $script:AuxG 103 1)).State | Should -Be 'Exported'
            $store.GetStats().Items | Should -Be (2 + 3 + 1)
        } finally { $store.Dispose() }
    }
    It 'the same id in two mailboxes (a store written before the move was known) is still one record' {
        $svc = New-MoveWorld
        $store = New-TestStoreAt 'ox4'
        try {
            $null = $store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:Arch; Kind = 'Archive' })
            $null = $store.UpsertFolders([MailboxTransfer.StoreFolder[]]@([MailboxTransfer.StoreFolder]@{ MailboxId = $script:Arch; FolderId = (Fid $script:AuxG 103); DisplayName = 'Both' }))
            $null = $store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@([MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:Arch; ItemId = (Mid $script:AuxG 103 1); FolderId = (Fid $script:AuxG 103); Size = 10 }))
            $x = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $store.GetStats().Items | Should -Be (1 + 1 + 3 + 1) -Because 'the two messages of Both (one already recorded), Moved mail 3, the own Inbox 1'
        } finally { $store.Dispose() }
    }
}

Describe 'A selection made on the main archive resolves after the folder moved to an auxiliary archive' {
    BeforeAll {
        function New-Entry([string]$FolderId, [string]$Path, $Items) {
            $e = [MailboxTransfer.SelectionEntry]@{ FolderId = $FolderId; FolderPath = $Path; Exportable = $true }
            if ($Items) { $e.Items = [System.Collections.Generic.List[MailboxTransfer.SelectionItem]]::new(); foreach ($i in $Items) { $e.Items.Add($i) } }
            return $e
        }
        function New-Selection([MailboxTransfer.SelectionEntry[]]$Entries) {
            $sel = [MailboxTransfer.BrowseSelection]::For((Get-Loc Archive))
            foreach ($e in $Entries) { $sel.Entries.Add($e) }
            return $sel
        }
        function New-GoneWorld {
            # 'Gone' (777) was in the main archive when the selection was saved; now only the auxiliary archive has it (4 items)
            $svc = New-MoveWorld -NoFailure
            $null = $svc.AddFolder($script:Aux1, (Fid $script:AuxG 777), 'Gone')
            foreach ($n in 1..4) { $null = $svc.AddItem($script:Aux1, (Fid $script:AuxG 777), (Mid $script:AuxG 777 $n), (New-Message "Gone $n"), "ck-g$n") }
            return $svc
        }
        function Get-Others { @((Get-Loc AuxArchive)) }
    }
    It 'a folder gone from the main archive is found in the auxiliary archive by its folder entry key: Match Moved, with the mailbox' {
        $svc = New-GoneWorld
        $sel = New-Selection @(New-Entry (Fid $script:ArchG 777) 'Gone' $null)
        $b = Get-Browser $svc
        $plan = $b.ResolvePlan($sel, $true, [Threading.CancellationToken]::None, $null, [MailboxTransfer.BrowseLocation[]]@((Get-Others)))
        $plan.IsComplete | Should -BeTrue
        $plan.Folders.Count | Should -Be 1
        $plan.Folders[0].Match | Should -Be 'Moved'
        $plan.Folders[0].MailboxId | Should -Be $script:Aux1
        $plan.Folders[0].ExportFolderId | Should -Be (Fid $script:AuxG 777)
        $plan.Folders[0].Note | Should -BeLike '*moved to AuxArchive*since the selection was made*'
        @($plan.Mailboxes) | Should -Be @($script:Aux1)
        # without the other archives to search the same selection is what it was: unresolved
        $old = $b.ResolvePlan($sel, $true)
        $old.IsComplete | Should -BeFalse
        $old.Folders[0].Match | Should -Be 'Unresolved'
    }
    It 'a placeholder left in the main archive: the plan also names the auxiliary archive that holds the items' {
        $svc = New-GoneWorld
        $moved = (Get-Browser $svc).GetFolderTree((Get-Loc Archive)).Folders | Where-Object Path -EQ 'Moved mail'
        $sel = New-Selection @(New-Entry $moved.Id 'Moved mail' $null)
        $plan = (Get-Browser $svc).ResolvePlan($sel, $true, [Threading.CancellationToken]::None, $null, [MailboxTransfer.BrowseLocation[]]@((Get-Others)))
        $plan.IsComplete | Should -BeTrue
        $plan.Folders.Count | Should -Be 2
        $plan.Folders[0].Match | Should -Be 'SameId'
        $plan.Folders[0].MailboxId | Should -BeNullOrEmpty
        $plan.Folders[1].Match | Should -Be 'Moved'
        $plan.Folders[1].MailboxId | Should -Be $script:Aux1
        $plan.Folders[1].ExportFolderId | Should -Be (Fid $script:AuxG 102)
        @($plan.Mailboxes) | Should -Be @($script:Arch, $script:Aux1)
        $plan.ForMailbox($script:Aux1).Folders.Count | Should -Be 1
        $plan.ForMailbox($script:Arch).Folders.Count | Should -Be 1
    }
    It 'a folder that did not move is resolved as before, and the other archives are not read' {
        $svc = New-GoneWorld
        $inbox = (Get-Browser $svc).GetFolderTree((Get-Loc Archive)).Folders | Where-Object Path -EQ 'Inbox'
        $sel = New-Selection @(New-Entry $inbox.Id 'Inbox' $null)
        $before = $svc.RequestsOf("mailboxes/$($script:Aux1)")
        $plan = (Get-Browser $svc).ResolvePlan($sel, $true, [Threading.CancellationToken]::None, $null, [MailboxTransfer.BrowseLocation[]]@((Get-Others)))
        $plan.Folders.Count | Should -Be 1
        $plan.Folders[0].Match | Should -Be 'SameId'
        $svc.RequestsOf("mailboxes/$($script:Aux1)") | Should -Be $before
    }
    It 'items ticked on the main archive are found in the auxiliary archive by their message entry key' {
        $svc = New-GoneWorld
        $items = @(
            [MailboxTransfer.SelectionItem]@{ Id = (Mid $script:ArchG 777 2); Size = 10 }
            [MailboxTransfer.SelectionItem]@{ Id = (Mid $script:ArchG 777 4); Size = 10 }
            [MailboxTransfer.SelectionItem]@{ Id = (Mid $script:ArchG 777 99); Size = 10 })
        $sel = New-Selection @(New-Entry (Fid $script:ArchG 777) 'Gone' $items)
        $plan = (Get-Browser $svc).ResolvePlan($sel, $true, [Threading.CancellationToken]::None, $null, [MailboxTransfer.BrowseLocation[]]@((Get-Others)))
        $plan.IsComplete | Should -BeFalse -Because 'message 99 exists nowhere'
        $plan.Problems.Count | Should -Be 1
        $plan.Problems[0] | Should -BeLike "*item $(Mid $script:ArchG 777 99) is not in the import-export listing*"
        $plan.Folders[0].Match | Should -Be 'Moved'
        @($plan.Folders[0].ItemIds) | Should -Be @((Mid $script:AuxG 777 2), (Mid $script:AuxG 777 4))
        $plan.Folders[0].Note | Should -BeLike '*2 by entry key*1 not found*'
    }
    It 'items of a folder that did not move, some of which moved with it elsewhere, are taken from both places' {
        $svc = New-GoneWorld
        # 'Both' (103) exists on both sides with the same two messages: ticking them on the main archive takes each message once
        $inboxBoth = (Get-Browser $svc).GetFolderTree((Get-Loc Archive)).Folders | Where-Object Path -EQ 'Both'
        $items = @([MailboxTransfer.SelectionItem]@{ Id = (Mid $script:ArchG 103 1); Size = 10 }, [MailboxTransfer.SelectionItem]@{ Id = (Mid $script:ArchG 103 2); Size = 10 })
        $sel = New-Selection @(New-Entry $inboxBoth.Id 'Both' $items)
        $plan = (Get-Browser $svc).ResolvePlan($sel, $true, [Threading.CancellationToken]::None, $null, [MailboxTransfer.BrowseLocation[]]@((Get-Others)))
        $plan.IsComplete | Should -BeTrue
        $plan.Folders.Count | Should -Be 1
        @($plan.Folders[0].ItemIds) | Should -Be @((Mid $script:ArchG 103 1), (Mid $script:ArchG 103 2))
    }
    It 'a folder whose content cannot be read yet is a problem of the plan, never an empty selection' {
        $svc = New-GoneWorld
        $svc.ListingFailures["$($script:Aux1)|$(Fid $script:AuxG 777)"] = [MailboxTransfer.Tests.FakeListingFailure]::new()
        $sel = New-Selection @(New-Entry (Fid $script:ArchG 777) 'Gone' @([MailboxTransfer.SelectionItem]@{ Id = (Mid $script:ArchG 777 1); Size = 10 }))
        $plan = (Get-Browser $svc).ResolvePlan($sel, $true, [Threading.CancellationToken]::None, $null, [MailboxTransfer.BrowseLocation[]]@((Get-Others)))
        $plan.IsComplete | Should -BeFalse
        $plan.Problems[0] | Should -BeLike "*'Gone': the items cannot be checked, the content of the folder is not available yet*Cannot query rows*"
    }
    It 'live (2026-10-31): the service redirects the folder request of a moved folder to the auxiliary archive, which answers with its own folder id: the plan exports it from there, items included' {
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Arch); $null = $svc.AddMailbox($script:Aux1)
        $null = $svc.AddFolder($script:Arch, 'F-MOVED', 'Old mail'); $null = $svc.AddFolder($script:Aux1, 'F-MOVED', 'Old mail')
        foreach ($i in 1..3) { $null = $svc.AddItem($script:Aux1, 'F-MOVED', "M$i", (New-Message "Old $i"), "ck$i") }
        $svc.MovedFolderToMailbox['F-MOVED'] = $script:Aux1; $svc.RedirectFolders = $true
        $whole = New-Selection @(New-Entry 'F-MOVED' 'Old mail' $null)
        $plan = (Get-Browser $svc).ResolvePlan($whole, $true)
        $plan.IsComplete | Should -BeTrue
        $plan.Folders[0].Match | Should -Be 'Moved'
        $plan.Folders[0].MailboxId | Should -Be $script:Aux1
        $plan.Folders[0].Note | Should -BeLike '*redirected the folder to MBX:bbbbbbbb*'
        @($plan.Mailboxes) | Should -Be @($script:Aux1)
        $items = New-Selection @(New-Entry 'F-MOVED' 'Old mail' @([MailboxTransfer.SelectionItem]@{ Id = 'M2'; Size = 10 }))
        $p2 = (Get-Browser $svc).ResolvePlan($items, $true)
        $p2.IsComplete | Should -BeTrue
        @($p2.Folders[0].ItemIds) | Should -Be @('M2')
        $p2.Folders[0].MailboxId | Should -Be $script:Aux1
        $store = New-TestStoreAt 'pl2'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive' -Plan $p2.ForMailbox($script:Aux1)
            $r.Result.Exported | Should -Be 1
        } finally { $store.Dispose() }
    }
    It 'each mailbox exports its part of the plan; a job refuses a plan that names another mailbox' {
        $svc = New-GoneWorld
        $sel = New-Selection @(New-Entry (Fid $script:ArchG 777) 'Gone' $null)
        $plan = (Get-Browser $svc).ResolvePlan($sel, $true, [Threading.CancellationToken]::None, $null, [MailboxTransfer.BrowseLocation[]]@((Get-Others)))
        $store = New-TestStoreAt 'pl1'
        try {
            { Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive' -Plan $plan } | Should -Throw '*other mailboxes*ForMailbox*'
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive' -Plan $plan.ForMailbox($script:Aux1)
            $r.Result.Exported | Should -Be 4 -Because 'only the planned folder, not the other 8 items of the auxiliary archive'
            $r.Result.Failed | Should -Be 0
            $store.GetStats().Items | Should -Be 4
        } finally { $store.Dispose() }
    }
}

Describe 'Invoke-MailboxTransfer.ps1: the exit code of an export that is not complete' {
    BeforeAll {
        function Invoke-Entry {
            param([Parameter(Mandatory)]$Service, [hashtable]$Parameters = @{})
            $global:MbtEntryClient = New-FakeClient -Service $Service
            $global:LASTEXITCODE = $null
            $text = & $script:Entry @Parameters *>&1 | Out-String
            $global:MbtEntryExit = $LASTEXITCODE
            return $text
        }
    }
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph $global:MbtEntryClient
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('Mail.Read', 'MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtEntryClient, MbtEntryExit -Scope Global -ErrorAction SilentlyContinue }
    It 'a deferred folder: the summary names it with the message of the service, the exit code is 2, the same command later finishes with 0' {
        $svc = New-MoveWorld -AuxOldItems 4
        $store = Join-Path $script:Tmp 'ex-aux'
        $text = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'AuxArchive'; MailboxId = $script:Aux1; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ex-tmp1') }
        $global:MbtEntryExit | Should -Be 2
        $text | Should -Match 'export finished: 6 items exported'
        $text | Should -Match '1 folder\(s\) DEFERRED, not exported'
        $text | Should -Match "'Old mail' in $([regex]::Escape($script:Aux1)): content not available yet \(HTTP 500 ErrorInternalServerError: .*Cannot query rows in a table"
        $text | Should -Match 'exit code 2: the export is not complete'
        $svc.ListingFailures.Clear()
        $again = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'AuxArchive'; MailboxId = $script:Aux1; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ex-tmp1') }
        $global:MbtEntryExit | Should -Be 0
        $again | Should -Match 'export finished: 4 items exported'
        $again | Should -Not -Match 'folder\(s\) DEFERRED'
    }
    It 'a complete export keeps exit code 0' {
        $svc = New-MoveWorld -NoFailure
        $text = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'AuxArchive'; MailboxId = $script:Aux1; StorePath = (Join-Path $script:Tmp 'ex-ok'); TempPath = (Join-Path $script:Tmp 'ex-tmp2') }
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'export finished: 6 items exported'
    }
    It '-Mailbox All: primary, archive, then the auxiliary archive given by -LocationsFile: one record per message, exit code 0' {
        $svc = New-MoveWorld -NoFailure
        $primary = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
        $null = $svc.AddMailbox($primary); $null = $svc.AddFolder($primary, 'P-INBOX', 'Inbox')
        foreach ($i in 1..2) { $null = $svc.AddItem($primary, 'P-INBOX', "P$i", (New-Message "Primary $i"), "ck-P$i") }
        $file = Join-Path $script:Tmp 'locations-all.txt'
        Set-Content -LiteralPath $file -Value "MainArchive $($script:ArchG)`nAuxArchive $($script:AuxG)"
        $store = Join-Path $script:Tmp 'ex-all'
        $text = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'All'; PrimaryMailboxId = $primary; ArchiveMailboxId = $script:Arch; LocationsFile = $file; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ex-tmp3') }
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'export finished: 2 items exported'
        $text | Should -Match 'export finished: 9 items exported'
        $text | Should -Match 'export finished: 4 items exported'
        $text | Should -Match 'recorded once'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try { $s.GetStats().Items | Should -Be (2 + 9 + 4) } finally { $s.Dispose() }
    }
    It 'a selection saved on the main archive exports from the auxiliary archive after the folder moved (the other archives come from -LocationsFile)' {
        $svc = New-MoveWorld -NoFailure
        $null = $svc.AddFolder($script:Aux1, (Fid $script:AuxG 777), 'Gone')
        foreach ($n in 1..4) { $null = $svc.AddItem($script:Aux1, (Fid $script:AuxG 777), (Mid $script:AuxG 777 $n), (New-Message "Gone $n"), "ck-g$n") }
        $sel = [MailboxTransfer.BrowseSelection]::For((Get-Loc Archive))
        $sel.Entries.Add([MailboxTransfer.SelectionEntry]@{ FolderId = (Fid $script:ArchG 777); FolderPath = 'Gone'; Exportable = $true })
        $selFile = Join-Path $script:Tmp 'moved-selection.json'
        $sel.Save($selFile)
        $file = Join-Path $script:Tmp 'locations-moved.txt'
        Set-Content -LiteralPath $file -Value "MainArchive $($script:ArchG)`nAuxArchive $($script:AuxG)"
        $store = Join-Path $script:Tmp 'ex-moved'
        $text = Invoke-Entry $svc @{ Mode = 'Export'; SelectionFile = $selFile; LocationsFile = $file; ArchiveMailboxId = $script:Arch; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ex-tmp4') }
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'Gone: folder in MBX:bbbbbbbb.*\[Moved\]'
        $text | Should -Match "nothing of the selection is in the Archive mailbox"
        $text | Should -Match 'export finished: 4 items exported'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try { $s.GetStats().Items | Should -Be 4 } finally { $s.Dispose() }
    }
    It '-Mode Browse -Mailbox AuxArchive on a folder being moved says so and exits 0' {
        $svc = New-MoveWorld
        $text = Invoke-Entry $svc @{ Mode = 'Browse'; Mailbox = 'AuxArchive'; MailboxId = $script:Aux1; Folder = 'Old mail'; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'content not available yet \(.*Cannot query rows in a table'
    }
}
