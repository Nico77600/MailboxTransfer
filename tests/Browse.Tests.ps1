#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the live browse (Engine.Browse.cs, the Browse functions of the module, -Mode Browse and
    -Mode Export -SelectionFile of the entry script). Nothing here reaches the network: the engine talks to FakeMailApi
    (tests\FakeMailApi.cs), a fake Microsoft Graph Mail API that hands everything under /admin/exchange/ to the fake
    import and export service (tests\FakeMailboxService.cs), so one client talks to both APIs as it does against Graph.

    Run:  Invoke-Pester -Path .\tests\Browse.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-browse-tests-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp
    $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:Src = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
    $script:Arch = "MBX:aaaaaaaa-0000-0000-0000-00000000000a@$($script:Tenant)"
    $script:Aux1 = "MBX:bbbbbbbb-0000-0000-0000-00000000000b@$($script:Tenant)"

    function New-FakeClient {
        param([Parameter(Mandatory)]$Handler, [double]$DelayScale = 0)
        $o = [MailboxTransfer.GraphOptions]::new()
        $o.DelayScale = $DelayScale
        $o.UseJitter = $false
        return [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Handler)
    }

    function New-Browser {
        param([Parameter(Mandatory)]$World, [double]$DelayScale = 0)
        $client = New-FakeClient -Handler $World.Mail -DelayScale $DelayScale
        $World.Client = $client
        return [MailboxTransfer.MailboxBrowser]::new([MailboxTransfer.MailboxApi]::new($client))
    }

    function New-Message([string]$Subject, [int]$Padding = 0) {
        $b = [MailboxTransfer.Tests.FxBuilder]::new()
        $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $Subject)
        if ($Padding -gt 0) { $null = $b.Bin(0x7000, [byte[]]::new($Padding)) }
        return $b.ToArray()
    }

    function Get-Day([int]$Day) { [datetime]::new(2026, 10, $Day, 8, 0, 0, [DateTimeKind]::Utc) }

    <#
        The lab mailbox in miniature, in both APIs. Divergent: the import-export API uses other ids than the Mail API (IE- prefix on the
        folders, I- on the items), to exercise the fallbacks of the correlation (the lab tenant gives the same ids).
          Primary  Inbox (5 messages; announces one more child, "Inbox Notes", a non-mail folder only the import-export API lists)
                     Inbox/Projects (3), Sent Items (3), Empty, Calendar (non-mail, 2 items), Recoverable Items/Deletions (Mail API only)
          Archive  Old Mail (4), Deleted Items
          Aux      Expanded (3) with Expanded/Sub (1), import-export API only
    #>
    function New-World {
        param([switch]$Divergent)
        $ie = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $ie.RootOnlyFolders = $true
        $mail = [MailboxTransfer.Tests.FakeMailApi]::new($ie)
        $mail.DeltaPageSize = 2
        $w = [pscustomobject]@{ Ie = $ie; Mail = $mail; Client = $null; Divergent = [bool]$Divergent }
        $fid = { param($id) if ($Divergent) { "IE-$id" } else { $id } }
        $iid = { param($id) if ($Divergent) { "I-$id" } else { $id } }
        foreach ($mbx in $script:Src, $script:Arch, $script:Aux1) { $null = $ie.AddMailbox($mbx) }
        foreach ($mbx in $script:Src, $script:Arch) { $mail.AddMailbox($mbx) }
        # one message in both APIs: the Mail API row and the item of the import-export listing (same size, same date)
        $addMsg = {
            param($mbx, $folder, $id, $subject, $day, $pad)
            $data = New-Message $subject $pad
            $m = $mail.AddMessage($mbx, $folder, $id, $subject, (Get-Day $day), $data.Length)
            $null = $ie.AddItem($mbx, (& $fid $folder), (& $iid $id), $data, "ck-$id", (Get-Day $day))
            $m
        }
        # primary
        $null = $mail.AddFolder($script:Src, 'F-INBOX', 'Inbox', $null, 'inbox', $false, 1)
        $null = $mail.AddFolder($script:Src, 'F-PROJ', 'Projects', 'F-INBOX')
        $null = $mail.AddFolder($script:Src, 'F-SENT', 'Sent Items', $null, 'sentitems')
        $null = $mail.AddFolder($script:Src, 'F-EMPTY', 'Empty')
        $null = $mail.AddFolder($script:Src, 'R-ROOT', 'Recoverable Items', $null, $null, $true)
        $null = $mail.AddFolder($script:Src, 'R-DEL', 'Deletions', 'R-ROOT', $null, $true)
        $null = $ie.AddFolder($script:Src, (& $fid 'F-INBOX'), 'Inbox')
        $null = $ie.AddFolder($script:Src, (& $fid 'F-PROJ'), 'Projects', (& $fid 'F-INBOX'))
        $null = $ie.AddFolder($script:Src, (& $fid 'F-SENT'), 'Sent Items')
        $null = $ie.AddFolder($script:Src, (& $fid 'F-EMPTY'), 'Empty')
        $null = $ie.AddFolder($script:Src, (& $fid 'F-CAL'), 'Calendar', 'ROOT', 'IPF.Appointment')
        $null = $ie.AddFolder($script:Src, (& $fid 'F-NOTES'), 'Inbox Notes', (& $fid 'F-INBOX'), 'IPF.StickyNote')
        $a = & $addMsg $script:Src 'F-INBOX' 'M-I1' 'Alpha report' 1 100
        $a.To = @('Bob <bob@contoso.test>', 'Carol <carol@contoso.test>'); $a.Cc = @('Dan <dan@contoso.test>')
        $a.Attachments.Add([MailboxTransfer.Tests.FakeMailAttachment]@{ Id = 'A-I1-1'; Name = 'report.pdf'; ContentType = 'application/pdf'; Size = 5000 })
        $b = & $addMsg $script:Src 'F-INBOX' 'M-I2' 'Bravo' 2 200
        $b.Importance = 'high'; $b.Categories = @('Red', 'Blue'); $b.Sensitivity = 2; $b.Flag = 'flagged'; $b.IsRead = $false
        $null = & $addMsg $script:Src 'F-INBOX' 'M-I3' 'Charlie' 3 300
        $d = & $addMsg $script:Src 'F-INBOX' 'M-I4' 'delta' 4 400
        $d.Body = ('lorem ipsum ' * 250)
        $inner = [MailboxTransfer.Tests.FakeMailMessage]@{ Id = 'inner'; Subject = 'Inner'; FromName = 'Zed'; FromAddress = 'zed@contoso.test'; To = @('Yan <yan@contoso.test>'); Body = 'the attached message' }
        $inner.Attachments.Add([MailboxTransfer.Tests.FakeMailAttachment]@{ Id = 'A-IN-1'; Name = 'inner.txt'; Size = 12 })
        $d.Attachments.Add([MailboxTransfer.Tests.FakeMailAttachment]@{ Id = 'A-I4-1'; Name = 'Attached.msg'; ContentType = 'message/rfc822'; Size = 900; Item = $inner })
        $null = & $addMsg $script:Src 'F-INBOX' 'M-I5' 'Echo' 5 500
        $null = & $addMsg $script:Src 'F-PROJ' 'M-P1' 'Project one' 6 150
        $null = & $addMsg $script:Src 'F-PROJ' 'M-P2' 'Project two' 7 250
        $null = & $addMsg $script:Src 'F-PROJ' 'M-P3' 'Project three' 8 350
        $null = & $addMsg $script:Src 'F-SENT' 'M-S1' 'Sent one' 9 40
        $null = & $addMsg $script:Src 'F-SENT' 'M-S2' 'Sent two' 10 80
        $null = & $addMsg $script:Src 'F-SENT' 'M-S3' 'Sent three' 11 120
        $null = $ie.AddItem($script:Src, (& $fid 'F-CAL'), 'C1', [byte[]]::new(3000), 'ck-C1', (Get-Day 12))
        $null = $ie.AddItem($script:Src, (& $fid 'F-CAL'), 'C2', [byte[]]::new(3100), 'ck-C2', (Get-Day 13))
        # archive
        $null = $mail.AddFolder($script:Arch, 'F-OLD', 'Old Mail')
        $adel = $mail.AddFolder($script:Arch, 'F-ADEL', 'Deleted Items', $null, 'archivedeleteditems')
        $adel.NonMessageItems = 2    # two appointments: counted by the folder, not messages
        $mix = $mail.AddFolder($script:Arch, 'F-MIX', 'Mixed')
        $mix.NonMessageItems = 1
        $null = $ie.AddFolder($script:Arch, (& $fid 'F-OLD'), 'Old Mail')
        $null = $ie.AddFolder($script:Arch, (& $fid 'F-ADEL'), 'Deleted Items')
        $null = $ie.AddFolder($script:Arch, (& $fid 'F-MIX'), 'Mixed')
        $null = $ie.AddItem($script:Arch, (& $fid 'F-ADEL'), 'D1', [byte[]]::new(31167), 'ck-D1', (Get-Day 16))
        $null = $ie.AddItem($script:Arch, (& $fid 'F-ADEL'), 'D2', [byte[]]::new(31475), 'ck-D2', (Get-Day 17))
        $null = & $addMsg $script:Arch 'F-MIX' 'M-X1' 'Mixed one' 18 210
        $null = & $addMsg $script:Arch 'F-MIX' 'M-X2' 'Mixed two' 19 220
        $null = $ie.AddItem($script:Arch, (& $fid 'F-MIX'), 'D3', [byte[]]::new(5000), 'ck-D3', (Get-Day 20))
        foreach ($i in 1..4) { $null = & $addMsg $script:Arch 'F-OLD' "M-A$i" "Old $i" (14 + $i) (600 + 100 * $i) }
        # auxiliary archive: import-export API only
        $null = $ie.AddFolder($script:Aux1, 'F-AUX', 'Expanded')
        $null = $ie.AddFolder($script:Aux1, 'F-AUX-SUB', 'Sub', 'F-AUX')
        foreach ($i in 1..3) { $null = $ie.AddItem($script:Aux1, 'F-AUX', "X$i", [byte[]]::new(300 + 100 * $i), "ck-X$i", (Get-Day (20 + $i))) }
        $null = $ie.AddItem($script:Aux1, 'F-AUX-SUB', 'X4', [byte[]]::new(900), 'ck-X4', (Get-Day 25))
        return $w
    }

    <#
        Recoverable Items in the import-export API of a mailbox, as the service has them (measured live 2026-11-09): not in the top-level listing (Hidden), reachable by id and by the well-known
        name of the root. Primary: Deletions (2 messages), Purges (1 message), Calendar Logging (3 appointments), SubstrateHolds (1 item), DiscoveryHolds and Versions empty; the messages of
        Deletions and Purges are in the Mail API too. Another mailbox (-Tag): the same folders, 1 item in Deletions, nothing in the Mail API. -Skip leaves folders out.
    #>
    function Add-Recoverable {
        param([Parameter(Mandatory)]$World, [string]$Mailbox = $script:Src, [string]$Tag = '', [string[]]$Skip = @())
        $ie = $World.Ie; $mail = $World.Mail
        $fid = { param($id) if ($World.Divergent) { "IE-$Tag$id" } else { "$Tag$id" } }
        $iid = { param($id) if ($World.Divergent) { "I-$Tag$id" } else { "$Tag$id" } }
        $root = $ie.AddFolder($Mailbox, (& $fid 'R-ROOT'), 'Recoverable Items'); $root.Hidden = $true; $root.WellKnown = 'recoverableitemsroot'
        $kids = [ordered]@{ 'R-CAL' = 'Calendar Logging'; 'R-DEL' = 'Deletions'; 'R-DISC' = 'DiscoveryHolds'; 'R-PURGE' = 'Purges'; 'R-SUBS' = 'SubstrateHolds'; 'R-VERS' = 'Versions' }
        foreach ($id in $kids.Keys) { if ($kids[$id] -in $Skip) { continue }; $f = $ie.AddFolder($Mailbox, (& $fid $id), $kids[$id], (& $fid 'R-ROOT')); $f.Hidden = $true }
        $add = {
            param($folder, $id, $subject, $day, $pad)
            $data = New-Message $subject $pad
            if (-not $Tag) { $null = $mail.AddMessage($Mailbox, $folder, $id, $subject, (Get-Day $day), $data.Length) }
            $null = $ie.AddItem($Mailbox, (& $fid $folder), (& $iid $id), $data, "ck-$id", (Get-Day $day))
        }
        if ($Tag) { & $add 'R-DEL' 'RM-1' 'Deleted in another mailbox' 20 111; return }
        $null = $mail.AddFolder($Mailbox, 'R-PURGE', 'Purges', 'R-ROOT', $null, $true)
        & $add 'R-DEL' 'RM-1' 'Deleted one' 20 111
        & $add 'R-DEL' 'RM-2' 'Deleted two' 21 222
        & $add 'R-PURGE' 'RM-3' 'Purged one' 22 333
        foreach ($i in 1..3) { $null = $ie.AddItem($Mailbox, (& $fid 'R-CAL'), "RL$i", [byte[]]::new(2000 + $i), "ck-RL$i", (Get-Day (22 + $i))) }
        $null = $ie.AddItem($Mailbox, (& $fid 'R-SUBS'), 'RS1', [byte[]]::new(700), 'ck-RS1', (Get-Day 27))
    }

    function Invoke-OptionsExport {
        <# Runs an ExportJob with the options to its end; returns the job, its messages and its result. #>
        param([Parameter(Mandatory)]$World, [Parameter(Mandatory)]$Store, [scriptblock]$Configure, [string]$Mailbox = $script:Src, [string]$Kind = 'Primary')
        $client = New-FakeClient -Handler $World.Mail
        $o = [MailboxTransfer.ExportOptions]::new()
        $o.MailboxId = $Mailbox; $o.Kind = $Kind; $o.Upn = 'adele@contoso.test'; $o.Concurrency = 2
        $o.TempPath = Join-Path $script:Tmp ('tmp-' + [guid]::NewGuid().ToString('N'))
        if ($Configure) { & $Configure $o }
        $job = [MailboxTransfer.ExportJob]::Start([MailboxTransfer.MailboxApi]::new($client), $Store, $o)
        $messages = [System.Collections.Generic.List[string]]::new()
        $deadline = [DateTime]::UtcNow.AddSeconds(60)
        while (-not $job.Wait(100)) {
            foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
            if ([DateTime]::UtcNow -gt $deadline) { $job.Cancel(); throw 'the export job did not finish in time' }
        }
        foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
        return [pscustomobject]@{ Job = $job; Messages = $messages; Result = $job.Result }
    }

    function Get-StorePaths($Store, [string]$Mailbox = $script:Src) {
        $out = [System.Collections.Generic.List[string]]::new()
        $walk = $null
        $walk = { param($nodes) foreach ($n in $nodes) { $out.Add($n.Path); & $walk $n.Children } }
        & $walk @($Store.GetFolderTree($Mailbox))
        return @($out | Sort-Object)
    }

    function Get-Loc([string]$Kind) {
        $id = switch ($Kind) { 'Primary' { $script:Src } 'Archive' { $script:Arch } default { $script:Aux1 } }
        return [MailboxTransfer.MailboxBrowser]::Location($Kind, $id, 'test')
    }

    function Get-Subjects($Page) { @($Page.Items | ForEach-Object { $_.Subject }) -join '|' }

    function New-TestStoreAt([string]$Name) {
        return [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests')
    }

    function Invoke-PlanExport {
        <# Runs an ExportJob with a plan to its end; returns the job, its messages and the client. #>
        param([Parameter(Mandatory)]$World, [Parameter(Mandatory)]$Store, [Parameter(Mandatory)]$Plan, [string]$Mailbox = $script:Src, [string]$Kind = 'Primary')
        $client = New-FakeClient -Handler $World.Mail
        $o = [MailboxTransfer.ExportOptions]::new()
        $o.MailboxId = $Mailbox; $o.Kind = $Kind; $o.Upn = 'adele@contoso.test'; $o.Concurrency = 2; $o.Plan = $Plan
        $o.TempPath = Join-Path $script:Tmp ('tmp-' + [guid]::NewGuid().ToString('N'))
        $job = [MailboxTransfer.ExportJob]::Start([MailboxTransfer.MailboxApi]::new($client), $Store, $o)
        $messages = [System.Collections.Generic.List[string]]::new()
        $deadline = [DateTime]::UtcNow.AddSeconds(60)
        while (-not $job.Wait(100)) {
            foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
            if ([DateTime]::UtcNow -gt $deadline) { $job.Cancel(); throw 'the export job did not finish in time' }
        }
        foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
        return [pscustomobject]@{ Job = $job; Messages = $messages; Result = $job.Result }
    }
}

AfterAll {
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Browse: folders' {
    It 'Primary tree by the Mail API: delta in pages with Prefer, counts and sizes, the folders only the import-export API lists merged in' {
        $w = New-World; $b = New-Browser $w
        $tree = $b.GetFolderTree((Get-Loc Primary))
        $tree.FromMailApi | Should -BeTrue
        ($tree.Folders | ForEach-Object Path) -join '|' | Should -Be 'Calendar|Empty|Inbox|Inbox/Inbox Notes|Inbox/Projects|Sent Items'
        $inbox = $tree.Find('Inbox')
        $inbox.TotalItemCount | Should -Be 5
        $inbox.SizeBytes | Should -Be ((New-Message 'Alpha report' 100).Length + (New-Message 'Bravo' 200).Length + (New-Message 'Charlie' 300).Length + (New-Message 'delta' 400).Length + (New-Message 'Echo' 500).Length)
        $inbox.UnreadItemCount | Should -Be 1
        $inbox.WellKnownName | Should -Be 'inbox'
        $inbox.Source | Should -Be 'MailApi'
        $inbox.ExportFolderId | Should -Be $inbox.Id
        $inbox.InExportTree | Should -BeTrue
        $inbox.ChildFolderCount | Should -Be 2
        $tree.Find('Inbox/Projects').Depth | Should -Be 1
        $cal = $tree.Find('Calendar')
        $cal.Source | Should -Be 'ImportExportApi'
        $cal.FolderClass | Should -Be 'IPF.Appointment'
        $cal.TotalItemCount | Should -Be 2
        $cal.SizeBytes | Should -Be -1
        $tree.Find('Inbox/Inbox Notes').Source | Should -Be 'ImportExportApi'
        $delta = @($w.Mail.Requests() | Where-Object { $_ -like '*mailFolders/delta*' })
        $delta.Count | Should -Be 2 -Because 'four mail folders, two a page'
        $delta | ForEach-Object { $_ | Should -Match '\[Prefer: odata.maxpagesize=200\]' }
        $delta[1] | Should -Match 'skiptoken=2' -Because 'the nextLink of the service is followed as returned'
        $w.Ie.RequestsOf('/childFolders') | Should -Be 1 -Because 'only Inbox announces more children than the Mail API lists'
        $tree.Notes.Count | Should -Be 0
    }

    It 'the archive is addressed as a mailbox of its own (/users/MBX:archive-guid@tenant), never through the user' {
        $w = New-World; $b = New-Browser $w
        $tree = $b.GetFolderTree((Get-Loc Archive))
        ($tree.Folders | ForEach-Object Path) -join '|' | Should -Be 'Deleted Items|Mixed|Old Mail'
        $tree.Find('Old Mail').TotalItemCount | Should -Be 4
        $tree.Find('Deleted Items').WellKnownName | Should -Be 'archivedeleteditems'
        $mailRequests = @($w.Mail.Requests())
        $mailRequests.Count | Should -BeGreaterThan 0
        foreach ($r in $mailRequests) { $r | Should -Match "^GET https://graph\.microsoft\.com/v1\.0/users/$([regex]::Escape($script:Arch))/mailFolders" }
        ($mailRequests -join ' ') | Should -Not -Match '/me/|/users/[^M]'
        $items = $b.ListItems((Get-Loc Archive), $tree.Find('Old Mail'))
        $items.Items.Count | Should -Be 4
        @($w.Mail.Requests() | Where-Object { $_ -like '*/messages*' })[0] | Should -Match "/users/$([regex]::Escape($script:Arch))/mailFolders/F-OLD/messages"
    }

    It 'Recoverable Items are added only when asked: the tree and the counts come from the import-export API, the folders are exportable, the messages are read through the Mail API' {
        $w = New-World; Add-Recoverable $w; $b = New-Browser $w
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $tree.Folders.Count | Should -Be 6
        $w.Mail.RequestsOf('recoverableitemsroot') | Should -Be 0
        $w.Ie.RequestsOf('recoverableitemsroot') | Should -Be 0
        $b.IncludeRecoverableItems = $true
        $tree = $b.GetFolderTree($loc)
        $tree.Folders.Count | Should -Be 13 -Because '6 folders + the root of Recoverable Items + its 6 folders'
        ($tree.Folders | Where-Object Recoverable | ForEach-Object Path) -join '|' | Should -Be 'Recoverable Items|Recoverable Items/Calendar Logging|Recoverable Items/Deletions|Recoverable Items/DiscoveryHolds|Recoverable Items/Purges|Recoverable Items/SubstrateHolds|Recoverable Items/Versions'
        $w.Ie.RequestsOf('/folders/recoverableitemsroot') | Should -Be 1 -Because 'the root is asked by its well-known name'
        $r = $tree.Find('Recoverable Items/Deletions')
        $r.Recoverable | Should -BeTrue
        $r.InExportTree | Should -BeTrue
        $r.ExportFolderId | Should -Be $r.Id
        $r.TotalItemCount | Should -Be 2
        $r.Source | Should -Be 'MailApi' -Because 'the Mail API knows the folder: its messages are listed through it, with the size it gives'
        $r.SizeBytes | Should -Be ((New-Message 'Deleted one' 111).Length + (New-Message 'Deleted two' 222).Length)
        $cal = $tree.Find('Recoverable Items/Calendar Logging')
        $cal.TotalItemCount | Should -Be 3 -Because 'the import-export API counts appointments too'
        $cal.Source | Should -Be 'ImportExportApi' -Because 'the Mail API does not know that folder'
        $cal.SizeBytes | Should -Be -1
        $root = $tree.Find('Recoverable Items')
        $root.ParentId | Should -BeNullOrEmpty
        $root.Depth | Should -Be 0
        $r.ParentId | Should -Be $root.Id
        $tree.Notes.Count | Should -Be 0
        # the rows of Deletions are the messages of the Mail API
        $page = $b.ListItems($loc, $r)
        Get-Subjects $page | Should -Be 'Deleted two|Deleted one'
        $page.Source | Should -Be 'MailApi'
        # Calendar Logging is not messages: the rows are the import-export listing
        $calPage = $b.ListItems($loc, $cal)
        $calPage.Source | Should -Be 'ImportExportApi'
        $calPage.Items.Count | Should -Be 3
    }

    It 'RecoverableItemsFolders limits the folders of Recoverable Items in the tree (the root is always there); Recoverable Items work on the Archive and an auxiliary archive too' {
        $w = New-World; Add-Recoverable $w; Add-Recoverable $w -Mailbox $script:Arch -Tag 'A-'; Add-Recoverable $w -Mailbox $script:Aux1 -Tag 'X-'
        $b = New-Browser $w
        $b.IncludeRecoverableItems = $true
        $b.RecoverableItemsFolders = [string[]]@('purges', 'Deletions')
        $tree = $b.GetFolderTree((Get-Loc Primary))
        ($tree.Folders | Where-Object Recoverable | ForEach-Object Path) -join '|' | Should -Be 'Recoverable Items|Recoverable Items/Deletions|Recoverable Items/Purges'
        $b.RecoverableItemsFolders = $null
        $arch = $b.GetFolderTree((Get-Loc Archive))
        $arch.FromMailApi | Should -BeTrue
        ($arch.Folders | Where-Object Recoverable | Measure-Object).Count | Should -Be 7
        $arch.Find('Recoverable Items/Deletions').TotalItemCount | Should -Be 1
        $aux = $b.GetFolderTree((Get-Loc AuxArchive))
        $aux.FromMailApi | Should -BeFalse
        ($aux.Folders | Where-Object Recoverable | Measure-Object).Count | Should -Be 7
        $aux.Find('Recoverable Items/Deletions').Source | Should -Be 'ImportExportApi'
        $w.Mail.RequestsOf("/users/$($script:Aux1)") | Should -Be 0
    }

    It 'lazy listing with Recoverable Items: the root at the first level, then its folders (only the ones asked for), counts from the import-export API, rows from the Mail API' {
        $w = New-World; Add-Recoverable $w; $b = New-Browser $w
        $b.IncludeRecoverableItems = $true
        $b.RecoverableItemsFolders = [string[]]@('Deletions', 'Purges')
        $loc = Get-Loc Primary
        $roots = $b.ListFolders($loc, $null)
        ($roots | ForEach-Object Path) -join '|' | Should -Be 'Calendar|Empty|Inbox|Recoverable Items|Sent Items'
        $root = $roots | Where-Object DisplayName -EQ 'Recoverable Items'
        $root.Recoverable | Should -BeTrue
        $root.ExportFolderId | Should -Be $root.Id
        $root.ChildFolderCount | Should -Be 6
        $kids = $b.ListFolders($loc, $root)
        ($kids | ForEach-Object Path) -join '|' | Should -Be 'Recoverable Items/Deletions|Recoverable Items/Purges'
        ($kids | ForEach-Object TotalItemCount) -join ',' | Should -Be '2,1'
        $kids[0].InExportTree | Should -BeTrue
        $kids[0].Depth | Should -Be 1
        Get-Subjects ($b.ListItems($loc, $kids[0])) | Should -Be 'Deleted two|Deleted one'
        $b.RecoverableItemsFolders = $null
        @($b.ListFolders($loc, $root)).Count | Should -Be 6
        $w.Mail.RequestsOf('mailFolders/delta') | Should -Be 0
    }

    It 'a mailbox without Recoverable Items (404) is a note in the tree, not an error' {
        $w = New-World; $b = New-Browser $w
        $b.IncludeRecoverableItems = $true
        $tree = $b.GetFolderTree((Get-Loc Primary))
        ($tree.Folders | Where-Object Recoverable | Measure-Object).Count | Should -Be 0
        ($tree.Notes -join ' ') | Should -Match 'Recoverable Items were not read'
    }

    It 'IncludeNonMailFolders off: the import-export API is not asked at all' {
        $w = New-World; $b = New-Browser $w
        $b.IncludeNonMailFolders = $false
        $tree = $b.GetFolderTree((Get-Loc Primary))
        ($tree.Folders | ForEach-Object Path) -join '|' | Should -Be 'Empty|Inbox|Inbox/Projects|Sent Items'
        $w.Ie.Log.Count | Should -Be 0
    }

    It 'lazy listing: the first level, then the children of a folder (merged with the non-mail children), at the cost of one request each' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $roots = $b.ListFolders($loc, $null)
        ($roots | ForEach-Object Path) -join '|' | Should -Be 'Calendar|Empty|Inbox|Sent Items'
        $b.IncludeRecoverableItems = $true
        $b.RecoverableItemsFolders = [string[]]@('Deletions', 'Purges')
        $withRec = $b.ListFolders($loc, $null)
        ($withRec | ForEach-Object Path) -join '|' | Should -Be 'Calendar|Empty|Inbox|Sent Items' -Because 'this world has no Recoverable Items in the import-export API: nothing is added'
        $b.IncludeRecoverableItems = $false
        $inbox = $roots | Where-Object DisplayName -EQ 'Inbox'
        $kids = $b.ListFolders($loc, $inbox)
        ($kids | ForEach-Object Path) -join '|' | Should -Be 'Inbox/Inbox Notes|Inbox/Projects'
        ($kids | Where-Object DisplayName -EQ 'Projects').Depth | Should -Be 1
        $cal = $roots | Where-Object DisplayName -EQ 'Calendar'
        @($b.ListFolders($loc, $cal)).Count | Should -Be 0
        $w.Mail.RequestsOf('mailFolders/delta') | Should -Be 0 -Because 'the lazy view never reads the whole tree'
    }

    It 'auxiliary archive: the Mail API does not reach it, the tree comes from the import-export API and says what is missing' {
        $w = New-World; $b = New-Browser $w
        $tree = $b.GetFolderTree((Get-Loc AuxArchive))
        $tree.FromMailApi | Should -BeFalse
        $tree.FallbackReason | Should -Match 'auxiliary archive'
        ($tree.Notes -join ' ') | Should -Match 'after the export'
        ($tree.Folders | ForEach-Object Path) -join '|' | Should -Be 'Expanded|Expanded/Sub'
        $tree.Find('Expanded').TotalItemCount | Should -Be 3
        $tree.Find('Expanded').Source | Should -Be 'ImportExportApi'
        $w.Mail.Requests().Count | Should -Be 0 -Because 'no /users/... request for an auxiliary archive'
    }

    It 'a mailbox the Mail API refuses (403: no Mail.Read) falls back to the import-export listing, with the reason' {
        $w = New-World; $b = New-Browser $w
        $null = $w.Mail.Forbidden.Add($script:Src)
        $tree = $b.GetFolderTree((Get-Loc Primary))
        $tree.FromMailApi | Should -BeFalse
        $tree.FallbackReason | Should -Match 'HTTP 403 ErrorAccessDenied'
        $tree.Folders.Count | Should -Be 6
        $tree.Notes[0] | Should -Match 'import-export listing'
        $page = $b.ListItems((Get-Loc Primary), $tree.Find('Inbox'))
        $page.Source | Should -Be 'ImportExportApi'
        $page.Items.Count | Should -Be 3 -Because 'the fake import-export API pages by 3'
    }

    It 'cancellation stops the tree; progress is reported through the callback, nothing is written to the console' {
        $w = New-World; $b = New-Browser $w
        $seen = [System.Collections.Generic.List[object]]::new()
        $cb = [Action[MailboxTransfer.BrowseProgress]]{ param($p) $seen.Add($p) }
        $null = $b.GetFolderTree((Get-Loc Primary), [Threading.CancellationToken]::None, $cb)
        $seen.Count | Should -BeGreaterThan 1
        $seen[$seen.Count - 1].Stage | Should -Be 'folders'
        $seen[$seen.Count - 1].Message | Should -Be 'done'
        $cts = [Threading.CancellationTokenSource]::new()
        $cts.Cancel()
        $cancelled = $null
        try { $null = $b.GetFolderTree((Get-Loc Primary), $cts.Token) } catch { $cancelled = $_.Exception }
        $cancelled | Should -Not -BeNullOrEmpty
        ($cancelled -is [OperationCanceledException] -or $cancelled.InnerException -is [OperationCanceledException]) | Should -BeTrue
    }
}

Describe 'Browse: items' {
    It 'rows are paged and lazy, newest first, with every header field of the Mail API' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $inbox = $b.GetFolderTree($loc).Find('Inbox')
        $q = [MailboxTransfer.BrowseItemsQuery]::new()
        $q.PageSize = 2
        $p1 = $b.ListItems($loc, $inbox, $q)
        Get-Subjects $p1 | Should -Be 'Echo|delta'
        $p1.HasMore | Should -BeTrue
        $p1.Source | Should -Be 'MailApi'
        $q.Token = $p1.NextToken
        $p2 = $b.ListItems($loc, $inbox, $q)
        Get-Subjects $p2 | Should -Be 'Charlie|Bravo'
        $q.Token = $p2.NextToken
        $p3 = $b.ListItems($loc, $inbox, $q)
        Get-Subjects $p3 | Should -Be 'Alpha report'
        $p3.HasMore | Should -BeFalse
        $p3.NextToken | Should -BeNullOrEmpty
        $bravo = $p2.Items[1]
        $bravo.Id | Should -Be 'M-I2'
        $bravo.FolderId | Should -Be 'F-INBOX'
        $bravo.FromName | Should -Be 'Alice'
        $bravo.FromAddress | Should -Be 'alice@contoso.test'
        $bravo.ToCount | Should -Be 1
        $bravo.CcCount | Should -Be 0
        $bravo.Importance | Should -Be 'high'
        ($bravo.Categories -join ',') | Should -Be 'Red,Blue'
        $bravo.Sensitivity | Should -Be 'private'
        $bravo.FlagStatus | Should -Be 'flagged'
        $bravo.IsRead | Should -BeFalse
        $bravo.ItemClass | Should -Be 'IPM.Note'
        $bravo.Size | Should -Be (New-Message 'Bravo' 200).Length
        $bravo.Received.UtcDateTime | Should -Be (Get-Day 2)
        $bravo.HasAttachments | Should -BeFalse
        $bravo.HasDetails | Should -BeTrue
        $alpha = $p3.Items[0]
        $alpha.HasAttachments | Should -BeTrue
        $alpha.ToCount | Should -Be 2
        $alpha.CcCount | Should -Be 1
        $messageRequests = @($w.Mail.Requests() | Where-Object { $_ -like '*/messages*' })
        $messageRequests.Count | Should -Be 3
        $messageRequests[0] | Should -Match '\$top=2'
        $messageRequests[0] | Should -Match 'orderby=receivedDateTime%20desc'
        $messageRequests[0] | Should -Match 'expand=singleValueExtendedProperties'
        $messageRequests[0] | Should -Match 'Integer%200x0E08'
        $messageRequests[0] | Should -Not -Match 'sensitivity' -Because 'the sensitivity is an extended property (0x0036), not a property of the message'
        $messageRequests[1] | Should -Match 'skip=2'
    }

    It 'sort orders: received, sent, subject, both ways' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $inbox = $b.GetFolderTree($loc).Find('Inbox')
        $expected = [ordered]@{
            ReceivedDescending = 'Echo|delta|Charlie|Bravo|Alpha report'; ReceivedAscending = 'Alpha report|Bravo|Charlie|delta|Echo'
            SentDescending = 'Echo|delta|Charlie|Bravo|Alpha report'; SentAscending = 'Alpha report|Bravo|Charlie|delta|Echo'
            SubjectAscending = 'Alpha report|Bravo|Charlie|delta|Echo'; SubjectDescending = 'Echo|delta|Charlie|Bravo|Alpha report'
        }
        foreach ($k in $expected.Keys) {
            $q = [MailboxTransfer.BrowseItemsQuery]::new()
            $q.Sort = [MailboxTransfer.BrowseSort]$k
            Get-Subjects ($b.ListItems($loc, $inbox, $q)) | Should -Be $expected[$k] -Because $k
        }
    }

    It 'Recipients off leaves To and Cc out of the request (their counts are not read)' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $q = [MailboxTransfer.BrowseItemsQuery]::new()
        $q.Recipients = $false
        $page = $b.ListItems($loc, $b.GetFolderTree($loc).Find('Inbox'), $q)
        @($w.Mail.Requests() | Where-Object { $_ -like '*/messages*' })[0] | Should -Not -Match 'toRecipients'
        $page.Items[0].ToCount | Should -Be -1
        $page.Items[0].Subject | Should -Be 'Echo'
    }

    It 'EnumerateItems pages lazily: stopping early costs no further page' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $inbox = $b.GetFolderTree($loc).Find('Inbox')
        $q = [MailboxTransfer.BrowseItemsQuery]::new()
        $q.PageSize = 2
        $first = @($b.EnumerateItems($loc, $inbox, $q) | Select-Object -First 3)
        ($first | ForEach-Object Subject) -join '|' | Should -Be 'Echo|delta|Charlie'
        $w.Mail.RequestsOf('/messages') | Should -Be 2
        @($b.EnumerateItems($loc, $inbox, $q)).Count | Should -Be 5
    }

    It 'the rows of a folder only the import-export API lists (Calendar) are size and dates, with a note' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $cal = $b.GetFolderTree($loc).Find('Calendar')
        $mailRequests = $w.Mail.Requests().Count
        $page = $b.ListItems($loc, $cal)
        $page.Source | Should -Be 'ImportExportApi'
        $page.Items.Count | Should -Be 2
        $page.Items[0].Subject | Should -BeNullOrEmpty
        $page.Items[0].HasDetails | Should -BeFalse
        $page.Items[0].Size | Should -Be 3000
        $page.Items[0].Created.UtcDateTime | Should -Be (Get-Day 12)
        ($page.Notes -join ' ') | Should -Match 'after the export'
        $w.Mail.RequestsOf('C1') | Should -Be 0
        $w.Mail.RequestsOf('/messages') | Should -Be 0
    }

    It 'auxiliary archive rows: import-export API, paged with its own token; no request to the Mail API' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc AuxArchive
        $tree = $b.GetFolderTree($loc)
        $q = [MailboxTransfer.BrowseItemsQuery]::new()
        $q.PageSize = 2
        $p1 = $b.ListItems($loc, $tree.Find('Expanded'), $q)
        $p1.Items.Count | Should -Be 3 -Because 'the fake service pages by 3 whatever $top says; the engine follows its nextLink'
        $q.PageSize = 1
        $all = @($b.EnumerateItems($loc, $tree.Find('Expanded'), $q))
        $all.Count | Should -Be 3
        ($all | ForEach-Object Id) -join '|' | Should -Be 'X1|X2|X3'
        $all[0].Size | Should -Be 400
        $w.Mail.Requests().Count | Should -Be 0
    }

    It 'a folder the Mail API refuses is listed by the import-export API, with the reason' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $sent = $b.GetFolderTree($loc).Find('Sent Items')
        $null = $w.Mail.Forbidden.Add($script:Src)
        $page = $b.ListItems($loc, $sent)
        $page.Source | Should -Be 'ImportExportApi'
        $page.Notes[0] | Should -Match 'Mail API refused the folder \(HTTP 403 ErrorAccessDenied\)'
        $page.Items.Count | Should -Be 3
    }

    It 'the items of a folder that are not messages (appointments in Deleted Items) are listed by the import-export API, with a note; in a mixed folder the note says how many are missing' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Archive
        $tree = $b.GetFolderTree($loc)
        $del = $tree.Find('Deleted Items')
        $del.TotalItemCount | Should -Be 2
        $page = $b.ListItems($loc, $del)
        $page.Source | Should -Be 'ImportExportApi'
        $page.Items.Count | Should -Be 2
        $page.Items[0].Size | Should -Be 31167
        $page.Notes[0] | Should -Match '2 of the 2 item\(s\) of this folder are not messages'
        $mixed = $tree.Find('Mixed')
        $m = $b.ListItems($loc, $mixed)
        $m.Source | Should -Be 'MailApi'
        $m.Items.Count | Should -Be 2
        $m.Notes[0] | Should -Match '1 of the 3 item\(s\) of this folder are not messages'
        $full = $b.ListItems($loc, $tree.Find('Old Mail'))
        $full.Notes.Count | Should -Be 0 -Because 'every item of Old Mail is a message'
    }

    It 'a 429 with Retry-After on a page is waited for (the delay is the one the service gave) and the page is returned' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $inbox = $b.GetFolderTree($loc).Find('Inbox')
        $w.Mail.Throttle(2, 7, '/messages')
        $page = $b.ListItems($loc, $inbox)
        $page.Items.Count | Should -Be 5
        $w.Mail.Throttled | Should -Be 2
        $w.Mail.RequestsOf('/messages') | Should -Be 3
        $w.Client.Delays.Count | Should -Be 2
        foreach ($d in $w.Client.Delays) { $d.TotalSeconds | Should -Be 7 }
        $w.Client.Throttled429 | Should -Be 2
    }

    It 'a 429 on the tree (delta) is waited for too, then the tree is complete' {
        $w = New-World; $b = New-Browser $w
        $w.Mail.Throttle(1, 3, 'delta')
        $tree = $b.GetFolderTree((Get-Loc Primary))
        $tree.Folders.Count | Should -Be 6
        $w.Client.Delays.Count | Should -Be 1
        $w.Client.Delays[0].TotalSeconds | Should -Be 3
    }
}

Describe 'Browse: preview' {
    It 'header, text body cut at BodyChars, recipients and attachments; an attached item is read one level down' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $inbox = $b.GetFolderTree($loc).Find('Inbox')
        $rows = @($b.ListItems($loc, $inbox).Items)
        $delta = $rows | Where-Object Subject -EQ 'delta'
        $o = [MailboxTransfer.BrowsePreviewOptions]::new()
        $o.BodyChars = 100
        $p = $b.GetItemPreview($loc, $delta, $o)
        $p.Source | Should -Be 'MailApi'
        $p.Header.Subject | Should -Be 'delta'
        $p.Header.Size | Should -Be (New-Message 'delta' 400).Length
        $p.To.Count | Should -Be 1
        $p.To[0].Address | Should -Be 'bob@contoso.test'
        $p.BodyText.Length | Should -Be 100
        $p.BodyTrimmed | Should -BeTrue
        $p.BodyText | Should -Not -Match '<html'
        $p.BodyPreview.Length | Should -Be 255
        @($w.Mail.Requests() | Where-Object { $_ -like '*/messages/M-I4?*' })[0] | Should -Match 'Prefer: outlook.body-content-type="text"'
        $p.Attachments.Count | Should -Be 1
        $att = $p.Attachments[0]
        $att.Kind | Should -Be 'item'
        $att.Name | Should -Be 'Attached.msg'
        $att.IsInline | Should -BeFalse
        $att.Item.Header.Subject | Should -Be 'Inner'
        $att.Item.Header.FromAddress | Should -Be 'zed@contoso.test'
        $att.Item.Header.ItemClass | Should -Be 'microsoft.graph.eventMessage'
        $att.Item.To[0].Address | Should -Be 'yan@contoso.test'
        $att.Item.Attachments[0].Name | Should -Be 'inner.txt'
        $alpha = $rows | Where-Object Subject -EQ 'Alpha report'
        $pa = $b.GetItemPreview($loc, $alpha)
        $pa.Attachments[0].Kind | Should -Be 'file'
        $pa.Attachments[0].Name | Should -Be 'report.pdf'
        $pa.Attachments[0].Size | Should -Be 5000
        $pa.Attachments[0].ContentType | Should -Be 'application/pdf'
        $pa.BodyTrimmed | Should -BeFalse
    }

    It 'BodyChars 0 asks for no body; NestedDepth 0 reads no attached item; a message without attachment costs no attachment request' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $rows = @($b.ListItems($loc, $b.GetFolderTree($loc).Find('Inbox')).Items)
        $o = [MailboxTransfer.BrowsePreviewOptions]::new()
        $o.BodyChars = 0; $o.NestedDepth = 0
        $p = $b.GetItemPreview($loc, ($rows | Where-Object Subject -EQ 'delta'), $o)
        $p.BodyText | Should -BeNullOrEmpty
        $p.BodyPreview | Should -Not -BeNullOrEmpty
        $p.Attachments[0].Item | Should -BeNullOrEmpty
        @($w.Mail.Requests() | Where-Object { $_ -like '*/messages/M-I4?*' })[0] | Should -Not -Match 'Prefer'
        $w.Mail.RequestsOf('attachments/A-I4-1') | Should -Be 0
        $before = $w.Mail.RequestsOf('/attachments')
        $none = $b.GetItemPreview($loc, ($rows | Where-Object Subject -EQ 'Charlie'))
        $none.Attachments.Count | Should -Be 0
        $w.Mail.RequestsOf('/attachments') | Should -Be $before
    }

    It 'a row of the import-export API (auxiliary archive, non-mail folder) previews to itself with a note and no request' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc AuxArchive
        $row = $b.ListItems($loc, $b.GetFolderTree($loc).Find('Expanded')).Items[0]
        $p = $b.GetItemPreview($loc, $row)
        $p.Source | Should -Be 'ImportExportApi'
        $p.Header.Size | Should -Be 400
        $p.Notes[0] | Should -Match 'after the export'
        $p.Attachments.Count | Should -Be 0
        $w.Mail.Requests().Count | Should -Be 0
    }

    It 'a message that is not there is a GraphException 404, not an empty preview' {
        $w = New-World; $b = New-Browser $w
        $row = [MailboxTransfer.BrowseItem]@{ Id = 'M-NOPE'; Source = 'MailApi'; FolderId = 'F-INBOX' }
        $failure = $null
        try { $null = $b.GetItemPreview((Get-Loc Primary), $row) } catch { $failure = $_.Exception }
        $ge = Get-MbtGraphException ([System.Management.Automation.ErrorRecord]::new($failure, 'x', 'NotSpecified', $null))
        $ge | Should -Not -BeNullOrEmpty
        $ge.StatusCode | Should -Be 404
    }
}

Describe 'Browse: correlation of the two APIs and selection -> export plan' {
    It 'the lab case: the same ids in both APIs; each entry is confirmed by one GET of the folder in the import-export API' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddFolder($tree.Find('Inbox'), $true)
        $null = $sel.AddFolder($tree.Find('Calendar'))
        $sent = $tree.Find('Sent Items')
        $rows = @($b.ListItems($loc, $sent).Items)
        $null = $sel.AddItems($sent, [MailboxTransfer.BrowseItem[]]@($rows | Select-Object -First 2))
        $plan = $b.ResolvePlan($sel)
        $plan.IsComplete | Should -BeTrue
        ($plan.Folders | ForEach-Object Match) -join ',' | Should -Be 'SameId,SameId,SameId'
        $plan.Folders[0].IncludeSubfolders | Should -BeTrue
        $plan.Folders[0].ItemIds | Should -BeNullOrEmpty
        $plan.Folders[0].ExportFolderId | Should -Be 'F-INBOX'
        $plan.Folders[1].Path | Should -Be 'Calendar'
        $plan.Folders[2].ItemIds.Count | Should -Be 2
        ($plan.Folders[2].ItemIds | Sort-Object) -join ',' | Should -Be 'M-S2,M-S3' -Because 'the two newest rows'
        $plan.Folders[2].Note | Should -Match '2 item\(s\) by id'
        @($w.Ie.Log | Where-Object { $_ -match '^GET .*/folders/F-INBOX$' }).Count | Should -Be 1 -Because 'one GET of the folder confirms the id'
        $w.Ie.RequestsOf('/folders/F-SENT') | Should -BeGreaterThan 0
    }

    It 'the id of the Mail API is unknown to the import-export API: the folder is found by its path, the items by size and date' {
        $w = New-World -Divergent; $b = New-Browser $w
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        # the Mail tree has the Mail ids and no folder twice: the folders of the same name at the same place are the same folder under another id
        $tree.Find('Inbox').Id | Should -Be 'F-INBOX'
        $tree.Find('Inbox').ExportFolderId | Should -Be 'IE-F-INBOX'
        $tree.Find('Inbox/Projects').ExportFolderId | Should -Be 'IE-F-PROJ'
        ($tree.Folders | ForEach-Object Path) -join '|' | Should -Be 'Calendar|Empty|Inbox|Inbox/Inbox Notes|Inbox/Projects|Sent Items'
        $tree.Find('Inbox/Inbox Notes').ParentId | Should -Be 'F-INBOX'
        ($tree.Notes -join ' ') | Should -Match 'other folder ids than the Mail API'
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddFolder($tree.Find('Inbox'), $true)
        $sent = $tree.Find('Sent Items')
        $rows = @($b.ListItems($loc, $sent).Items)
        $null = $sel.AddItems($sent, [MailboxTransfer.BrowseItem[]]$rows)
        $plan = $b.ResolvePlan($sel)
        $plan.IsComplete | Should -BeTrue -Because ($plan.Problems -join ' ')
        $plan.Folders[0].Match | Should -Be 'ByPath'
        $plan.Folders[0].ExportFolderId | Should -Be 'IE-F-INBOX'
        $plan.Folders[0].Note | Should -Match 'matched by path when the tree was read'
        $plan.Folders[1].Match | Should -Be 'ByPath'
        $plan.Folders[1].ExportFolderId | Should -Be 'IE-F-SENT'
        ($plan.Folders[1].ItemIds | Sort-Object) -join ',' | Should -Be 'I-M-S1,I-M-S2,I-M-S3'
        $plan.Folders[1].Note | Should -Match '3 by size and date'
    }

    It 'a selection written without the hint (by another console): the folder is looked up by id, then by path in the whole import-export tree' {
        $w = New-World -Divergent; $b = New-Browser $w
        $sel = [MailboxTransfer.BrowseSelection]::For((Get-Loc Primary))
        $sel.Entries.Add([MailboxTransfer.SelectionEntry]@{ FolderId = 'F-PROJ'; FolderPath = 'Inbox/Projects' })
        $sel.Entries.Add([MailboxTransfer.SelectionEntry]@{ FolderId = 'F-INBOX'; FolderPath = 'inbox' })
        $plan = $b.ResolvePlan($sel)
        $plan.IsComplete | Should -BeTrue -Because ($plan.Problems -join ' ')
        ($plan.Folders | ForEach-Object Match) -join ',' | Should -Be 'ByPath,ByPath'
        ($plan.Folders | ForEach-Object ExportFolderId) -join ',' | Should -Be 'IE-F-PROJ,IE-F-INBOX'
        $plan.Folders[0].Note | Should -Match 'found by its path'
        $w.Ie.RequestsOf('/folders/F-PROJ') | Should -Be 1
        $w.Ie.RequestsOf('/childFolders') | Should -BeGreaterThan 0 -Because 'the whole import-export tree is read once, for the path lookup'
    }

    It 'an item the import-export listing does not know, or that fits several, is reported and never guessed' {
        $w = New-World -Divergent; $b = New-Browser $w
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $sent = $tree.Find('Sent Items')
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddItems($sent, [MailboxTransfer.BrowseItem[]]@([MailboxTransfer.BrowseItem]@{ Id = 'M-GHOST'; Size = 5; Created = [DateTimeOffset]::new((Get-Day 1)) }))
        # two import-export items of the same size and date: ambiguous
        $null = $w.Ie.AddItem($script:Src, 'IE-F-SENT', 'I-TWIN1', [byte[]]::new(77), 'ck', (Get-Day 3))
        $null = $w.Ie.AddItem($script:Src, 'IE-F-SENT', 'I-TWIN2', [byte[]]::new(77), 'ck', (Get-Day 3))
        $null = $sel.AddItems($sent, [MailboxTransfer.BrowseItem[]]@([MailboxTransfer.BrowseItem]@{ Id = 'M-TWIN'; Size = 77; Created = [DateTimeOffset]::new((Get-Day 3)) }))
        $plan = $b.ResolvePlan($sel)
        $plan.IsComplete | Should -BeFalse
        $plan.Problems.Count | Should -Be 2
        ($plan.Problems -join ' ') | Should -Match 'item M-GHOST is not in the import-export listing'
        ($plan.Problems -join ' ') | Should -Match 'item M-TWIN matches 2 items .* not guessed'
        $plan.Folders[0].ItemIds.Count | Should -Be 0
    }

    It 'a folder of Recoverable Items resolves like any folder; the root with its subfolders counts the folders asked for; a folder that exists nowhere is Unresolved, the plan is not complete and cannot start an export' {
        $w = New-World; Add-Recoverable $w; $b = New-Browser $w
        $b.IncludeRecoverableItems = $true
        $b.RecoverableItemsFolders = [MailboxTransfer.RecoverableItems]::DefaultFolders
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $one = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $one.AddFolder($tree.Find('Recoverable Items/Deletions'))
        $p1 = $b.ResolvePlan($one)
        $p1.IsComplete | Should -BeTrue
        $p1.Folders[0].Match | Should -Be 'SameId'
        $p1.Folders[0].Items | Should -Be 2
        $all = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $all.AddFolder($tree.Find('Recoverable Items'), $true)
        $p2 = $b.ResolvePlan($all)
        $p2.IsComplete | Should -BeTrue
        $p2.Folders[0].Items | Should -Be 3 -Because 'Deletions 2 + Purges 1 (+ Versions and DiscoveryHolds, empty); Calendar Logging and SubstrateHolds are not in the default set'
        $p2.Folders[0].FolderCount | Should -Be 5
        $items = [MailboxTransfer.BrowseSelection]::For($loc)
        $del = $tree.Find('Recoverable Items/Deletions')
        $null = $items.AddItems($del, [MailboxTransfer.BrowseItem[]]@($b.ListItems($loc, $del).Items | Select-Object -First 1))
        $p3 = $b.ResolvePlan($items)
        $p3.IsComplete | Should -BeTrue
        $p3.Folders[0].ItemIds.Count | Should -Be 1
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddFolder($tree.Find('Recoverable Items/Deletions'))
        $sel.Entries.Add([MailboxTransfer.SelectionEntry]@{ FolderId = 'F-GONE'; FolderPath = 'Gone/Folder' })
        $plan = $b.ResolvePlan($sel)
        $plan.IsComplete | Should -BeFalse
        ($plan.Folders | ForEach-Object Match) -join ',' | Should -Be 'SameId,Unresolved'
        $plan.Problems.Count | Should -Be 1
        $plan.Problems[0] | Should -Match "no such folder in the import-export tree"
        $store = New-TestStoreAt 'plan-incomplete'
        try {
            $o = [MailboxTransfer.ExportOptions]::new()
            $o.MailboxId = $script:Src; $o.TempPath = Join-Path $script:Tmp 'x'; $o.Plan = $plan
            { [MailboxTransfer.ExportJob]::Start([MailboxTransfer.MailboxApi]::new($b.Graph), $store, $o) } | Should -Throw '*cannot be exported*'
        } finally { $store.Dispose() }
    }

    It 'NoVerify trusts the ids and makes no request to the import-export API' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddFolder($tree.Find('Inbox'))
        $logged = $w.Ie.Log.Count
        $plan = $b.ResolvePlan($sel, $false)
        $plan.IsComplete | Should -BeTrue
        $plan.Folders[0].Note | Should -Be 'not verified'
        $w.Ie.Log.Count | Should -Be $logged
    }

    It 'the selection file round-trips as JSON and refuses what is not a selection' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Archive
        $tree = $b.GetFolderTree($loc)
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddFolder($tree.Find('Old Mail'), $true)
        $row = $b.ListItems($loc, $tree.Find('Old Mail')).Items[0]
        $null = $sel.AddItems($tree.Find('Deleted Items'), [MailboxTransfer.BrowseItem[]]@($row))
        $file = Join-Path $script:Tmp 'sel.json'
        $sel.Save($file)
        $json = Get-Content -LiteralPath $file -Raw | ConvertFrom-Json
        $json.Version | Should -Be 1
        $json.Kind | Should -Be 'Archive'
        $json.MailboxId | Should -Be $script:Arch
        $json.Entries[0].FolderPath | Should -Be 'Old Mail'
        $json.Entries[0].IncludeSubfolders | Should -BeTrue
        $json.Entries[1].Items[0].Id | Should -Be $row.Id
        $back = [MailboxTransfer.BrowseSelection]::Load($file)
        $back.Entries.Count | Should -Be 2
        $back.Entries[1].Items[0].Size | Should -Be $row.Size
        $back.Entries[1].Items[0].Created.UtcDateTime | Should -Be $row.Created.UtcDateTime
        $back.Remove($back.Entries[0].FolderId) | Should -BeTrue
        $back.Entries.Count | Should -Be 1
        { [MailboxTransfer.BrowseSelection]::FromJson('{"Version":2,"Kind":"Primary","MailboxId":"x"}') } | Should -Throw '*version 2*'
        { [MailboxTransfer.BrowseSelection]::FromJson('{"Version":1}') } | Should -Throw '*names no mailbox*'
        { [MailboxTransfer.BrowseSelection]::FromJson('not json') } | Should -Throw '*not valid JSON*'
    }
}

Describe 'Browse: a selection drives the export (ExportOptions.Plan)' {
    It 'a folder with its subfolders, and named items of another folder: only those are inventoried and exported' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddFolder($tree.Find('Inbox'), $true)
        $sent = $tree.Find('Sent Items')
        $null = $sel.AddItems($sent, [MailboxTransfer.BrowseItem[]]@($b.ListItems($loc, $sent).Items | Select-Object -First 2))
        $plan = $b.ResolvePlan($sel)
        $store = New-TestStoreAt 'plan-export'
        try {
            $r = Invoke-PlanExport -World $w -Store $store -Plan $plan
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Exported | Should -Be 10 -Because '5 in Inbox + 3 in Projects (+ the empty Inbox Notes) + the 2 newest of Sent Items'
            $r.Result.Failed | Should -Be 0
            ($r.Messages -join "`n") | Should -Match 'plan: 2 entries, 4 folder\(s\) in scope, 2 item\(s\) named'
            $stats = $store.GetItemStats($script:Src)
            ($stats | Where-Object State -EQ 'Exported').Count | Should -Be 10
            ($stats | Where-Object State -EQ 'Pending').Count | Should -Be 0
            $w.Ie.ExportCalls.Count | Should -BeGreaterThan 0
            $w.Ie.RequestsOf('/folders/F-EMPTY/items') | Should -Be 0 -Because 'a folder outside the selection is not even listed'
            $w.Ie.RequestsOf('/folders/F-CAL/items') | Should -Be 0
            # the store records the folders of the selection and the path to them, not the rest of the tree (a PST made from it holds no empty folders around them)
            @($store.GetFolders($script:Src)).Count | Should -Be 4
        } finally { $store.Dispose() }
    }

    It 'a second run with another selection exports only what it names, and adds to the store of the first' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $store = New-TestStoreAt 'plan-twice'
        try {
            $s1 = [MailboxTransfer.BrowseSelection]::For($loc)
            $null = $s1.AddFolder($tree.Find('Inbox/Projects'))
            $r1 = Invoke-PlanExport -World $w -Store $store -Plan ($b.ResolvePlan($s1))
            $r1.Result.Exported | Should -Be 3
            $s2 = [MailboxTransfer.BrowseSelection]::For($loc)
            $sent = $tree.Find('Sent Items')
            $null = $s2.AddItems($sent, [MailboxTransfer.BrowseItem[]]@($b.ListItems($loc, $sent).Items | Where-Object Subject -EQ 'Sent one'))
            $r2 = Invoke-PlanExport -World $w -Store $store -Plan ($b.ResolvePlan($s2))
            $r2.Result.Exported | Should -Be 1
            $stats = $store.GetItemStats($script:Src)
            ($stats | Where-Object State -EQ 'Exported').Count | Should -Be 4
            ($stats | Where-Object State -EQ 'Pending').Count | Should -Be 0
        } finally { $store.Dispose() }
    }

    It 'refuses a plan with a FolderPath, an empty plan, and a plan folder the mailbox does not have' {
        $w = New-World; $b = New-Browser $w
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddFolder($tree.Find('Inbox'))
        $plan = $b.ResolvePlan($sel)
        $store = New-TestStoreAt 'plan-refused'
        try {
            $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Handler $w.Mail))
            $o = [MailboxTransfer.ExportOptions]::new()
            $o.MailboxId = $script:Src; $o.TempPath = Join-Path $script:Tmp 'x1'; $o.Plan = $plan; $o.FolderPath = 'Inbox'
            { [MailboxTransfer.ExportJob]::Start($api, $store, $o) } | Should -Throw '*exclusive*'
            $o.FolderPath = $null; $o.Plan = [MailboxTransfer.ExportPlan]::new()
            { [MailboxTransfer.ExportJob]::Start($api, $store, $o) } | Should -Throw '*plan is empty*'
            $bad = [MailboxTransfer.ExportPlan]::new()
            $bad.Folders.Add([MailboxTransfer.ExportPlanFolder]@{ SourceFolderId = 'X'; ExportFolderId = 'F-NOT-THERE'; Path = 'X'; Match = 'SameId' })
            $r = Invoke-PlanExport -World $w -Store $store -Plan $bad
            $r.Job.Error.Message | Should -Match 'not found in this mailbox'
        } finally { $store.Dispose() }
    }
}

Describe 'Browse: module functions and the entry script' {
    BeforeAll {
        $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
        function Invoke-Entry {
            param([Parameter(Mandatory)]$World, [hashtable]$Parameters = @{}, [string[]]$Scopes = @('Mail.Read', 'MailboxFolder.Read', 'MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'))
            # a new client at each sign-in, as the real Connect-MbtGraph makes one (the entry script may disconnect and connect again)
            $handler = $World.Mail
            $global:MbtEntryFactory = {
                $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
                [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $handler)
            }.GetNewClosure()
            $global:MbtEntryScopes = $Scopes
            $global:LASTEXITCODE = $null
            $text = & $script:Entry @Parameters *>&1 | Out-String
            $global:MbtEntryExit = $LASTEXITCODE
            return $text
        }
    }
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph (& $global:MbtEntryFactory)
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = $global:MbtEntryScopes; ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtEntryFactory, MbtEntryExit, MbtEntryScopes -Scope Global -ErrorAction SilentlyContinue }

    It 'Get-MbtBrowseFolder / Get-MbtBrowseItem / Get-MbtBrowsePreview work on a session (the functions the console will call)' {
        $w = New-World
        $null = Set-MbtSession -Graph (New-FakeClient -Handler $w.Mail)
        try {
            $loc = New-MbtBrowseLocation -Kind Primary -MailboxId $script:Src
            $folders = @(Get-MbtBrowseFolder -Location $loc)
            $folders.Count | Should -Be 6
            $inbox = Get-MbtBrowseFolder -Location $loc -Path 'Inbox'
            $inbox.TotalItemCount | Should -Be 5
            (Get-MbtBrowseFolder -Location $loc -Lazy | ForEach-Object Path) -join '|' | Should -Be 'Calendar|Empty|Inbox|Sent Items'
            $rows = @(Get-MbtBrowseItem -Location $loc -Folder $inbox -Top 3 -PageSize 2)
            ($rows | ForEach-Object Subject) -join '|' | Should -Be 'Echo|delta|Charlie'
            $page = Get-MbtBrowseItem -Location $loc -Folder 'Inbox' -PageSize 4 -Page -Sort SubjectAscending
            $page.Items.Count | Should -Be 4
            $page.HasMore | Should -BeTrue
            $next = Get-MbtBrowseItem -Location $loc -Folder $inbox -PageSize 4 -Page -Token $page.NextToken
            Get-Subjects $next | Should -Be 'Echo'
            $p = Get-MbtBrowsePreview -Location $loc -Item $rows[0] -BodyChars 50
            $p.Header.Subject | Should -Be 'Echo'
            (Get-MbtBrowsePreview -Location $loc -ItemId 'M-I1').Header.Subject | Should -Be 'Alpha report'
            $sel = New-MbtBrowseSelection -Location $loc
            $null = Add-MbtBrowseSelection -Selection $sel -Folder $inbox -IncludeSubfolders
            $file = Save-MbtBrowseSelection -Selection $sel -Path (Join-Path $script:Tmp 'fn-sel.json')
            (Read-MbtBrowseSelection -Path $file).Entries[0].FolderPath | Should -Be 'Inbox'
            $plan = Resolve-MbtBrowsePlan -Selection (Read-MbtBrowseSelection -Path $file)
            $plan.IsComplete | Should -BeTrue
            { Get-MbtBrowseFolder -Location $loc -Path 'Nope' } | Should -Throw '*not found*'
        } finally { Disconnect-MbtGraph }
    }

    It 'Find-MbtBrowseFolder walks down the path one level at a time (no whole tree), says when two folders share a name, falls back to the tree for the end of a path' {
        $w = New-World
        $null = $w.Mail.AddFolder($script:Src, 'F-DUP1', 'Twin')
        $null = $w.Mail.AddFolder($script:Src, 'F-DUP2', 'Twin')
        $null = Set-MbtSession -Graph (New-FakeClient -Handler $w.Mail)
        try {
            $loc = New-MbtBrowseLocation -Kind Primary -MailboxId $script:Src
            $f = Find-MbtBrowseFolder -Location $loc -Path 'Inbox/Projects'
            $f.Path | Should -Be 'Inbox/Projects'
            $f.TotalItemCount | Should -Be 3
            $w.Mail.RequestsOf('mailFolders/delta') | Should -Be 0
            $w.Mail.RequestsOf('/childFolders') | Should -Be 2 -Because 'msgfolderroot, then Inbox: one request a level'
            { Find-MbtBrowseFolder -Location $loc -Path 'Twin' } | Should -Throw '*is ambiguous at ''Twin''*F-DUP1*F-DUP2*'
            (Find-MbtBrowseFolder -Location $loc -Path 'Projects').Path | Should -Be 'Inbox/Projects' -Because 'the end of a path is searched in the whole tree'
            $w.Mail.RequestsOf('mailFolders/delta') | Should -BeGreaterThan 0
            { Find-MbtBrowseFolder -Location $loc -Path 'Nope/Nothing' } | Should -Throw '*not found in this mailbox*'
        } finally { Disconnect-MbtGraph }
    }

    It 'Get-MbtBrowseLocation converts Get-MbtMailboxLocation (parameters, operator file) to browse locations' {
        $file = Join-Path $script:Tmp 'loc.txt'
        Set-Content -LiteralPath $file -Value ("Primary $($script:Src.Substring(4, 36))`nMainArchive $($script:Arch.Substring(4, 36))`nAuxArchive $($script:Aux1.Substring(4, 36))")
        $locs = @(Get-MbtBrowseLocation -TenantId $script:Tenant -LocationsFile $file)
        ($locs | ForEach-Object Kind) -join ',' | Should -Be 'Primary,Archive,AuxArchive'
        ($locs | ForEach-Object HasMailApi) -join ',' | Should -Be 'True,True,False'
        $locs[2].MailboxId | Should -Be $script:Aux1
    }

    It '-Mode Browse prints the tree of the primary mailbox: counts, sizes, unread, and the folders of the import-export API marked' {
        $w = New-World
        $text = Invoke-Entry $w @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $text | Should -Match 'signed in as adele.vance@contoso.com'
        $text | Should -Match 'browsing the Primary mailbox MBX:11111111.*\(Microsoft Graph Mail API\)'
        $text | Should -Match '6 folder\(s\), source: Microsoft Graph Mail API'
        $text | Should -Match 'Inbox\s+5 item\(s\)\s+\d+(,\d+)? (B|KiB)\s+unread 1'
        $text | Should -Match '  Projects\s+3 item\(s\)'
        $text | Should -Match 'Calendar\s+2 item\(s\)\s+unknown\s+\[IPF\.Appointment, import-export API\]'
        $text | Should -Match 'total: 6 folder\(s\), 13 item\(s\), .*\(2 folders with unknown size, not counted\)'
        $text | Should -Not -Match 'Recoverable'
        $text | Should -Not -Match '(?m)^ {6}x {2}\S'
        $global:MbtEntryExit | Should -Not -Be 1
        (Test-Path (Join-Path $script:Tmp 'cli-none')) | Should -BeFalse -Because 'Browse writes nothing'
    }

    It '-Mode Browse -Mailbox Archive reads the archive as /users/MBX:archive-guid; -RecoverableItems adds the Recoverable Items folders' {
        $w = New-World
        $text = Invoke-Entry $w @{ Mode = 'Browse'; Mailbox = 'Archive'; ArchiveMailboxId = $script:Arch; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $text | Should -Match 'browsing the Archive mailbox MBX:aaaaaaaa'
        $text | Should -Match 'Old Mail\s+4 item\(s\)'
        $w.Mail.RequestsOf("/users/$($script:Arch)/") | Should -BeGreaterThan 0
        $w.Mail.RequestsOf("/users/$($script:Src)/") | Should -Be 0
        $wr = New-World; Add-Recoverable $wr
        $rec = Invoke-Entry $wr @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; RecoverableItems = $true; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $rec | Should -Match 'Deletions\s+2 item\(s\).*\[Recoverable Items\]'
        $rec | Should -Match 'Calendar Logging\s+3 item\(s\)'
        $rec | Should -Not -Match 'not exportable'
        $few = Invoke-Entry $wr @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; RecoverableItems = $true; RecoverableItemsFolders = @('Purges'); StorePath = (Join-Path $script:Tmp 'cli-none') }
        $few | Should -Match 'Purges\s+1 item\(s\)'
        $few | Should -Not -Match 'Deletions'
    }

    It '-Mode Browse -Folder lists the newest rows (-Top), says how many there are, and -ItemId previews a row' {
        $w = New-World
        $common = @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $list = Invoke-Entry $w ($common + @{ Folder = 'Inbox'; Top = 3 })
        $list | Should -Match 'folder Inbox: 5 item\(s\)'
        $w.Mail.RequestsOf('mailFolders/delta') | Should -Be 0 -Because 'a folder is found by walking down its path, not by reading the whole tree'
        $list | Should -Match '1\s+2026-10-05 08:00\s+Alice\s+Echo\s'
        $list | Should -Match '2\s+2026-10-04 08:00\s+Alice\s+delta\s'
        $list | Should -Match '3\s+2026-10-03 08:00\s+Alice\s+Charlie\s'
        $list | Should -Not -Match 'Bravo'
        $list | Should -Match '3 of 5 item\(s\) shown'
        $sorted = Invoke-Entry $w ($common + @{ Folder = 'Inbox'; Top = 0; Sort = 'SubjectAscending' })
        $sorted | Should -Match '1\s+2026-10-01 08:00\s+Alice\s+Alpha report\s+.*\sA\s+to 2/cc 1'
        $sorted | Should -Match 'Bravo.*\s!U?F?S?\s'
        $preview = Invoke-Entry $w ($common + @{ Folder = 'Inbox'; ItemId = '2'; Top = 5 })
        $preview | Should -Match 'preview of row 2'
        $preview | Should -Match 'Subject    : delta'
        $preview | Should -Match 'From       : Alice <alice@contoso.test>'
        $preview | Should -Match 'Attachment : Attached\.msg \(item'
        $preview | Should -Match 'Subject    : Inner'
        $preview | Should -Match 'Body       : lorem ipsum'
        { Invoke-Entry $w ($common + @{ ItemId = '1' }) } | Should -Throw '*-ItemId needs -Folder*'
        { Invoke-Entry $w ($common + @{ Folder = 'Nope' }) } | Should -Throw "*Folder 'Nope' not found*"
        { Invoke-Entry $w ($common + @{ Folder = 'Inbox'; ItemId = '9' }) } | Should -Throw '*the listing has 5 row(s)*'
    }

    It '-Mode Browse -Mailbox AuxArchive: tree and sizes from the import-export API, a note that the details come after the export, no Mail API request' {
        $w = New-World
        $common = @{ Mode = 'Browse'; Mailbox = 'AuxArchive'; MailboxId = $script:Aux1; StorePath = (Join-Path $script:Tmp 'cli-none'); LocationsFile = $null }
        $common.Remove('LocationsFile')
        # the kind is given with -Mailbox, the id with -MailboxId
        $text = Invoke-Entry $w $common
        $text | Should -Match 'browsing the AuxArchive mailbox MBX:bbbbbbbb.*import-export API only'
        $text | Should -Match 'Auxiliary archive: folders and items come from the import-export API'
        $text | Should -Match 'Expanded\s+3 item\(s\)'
        $text | Should -Match '  Sub\s+1 item\(s\)'
        $list = Invoke-Entry $w ($common + @{ Folder = 'Expanded' })
        $list | Should -Match 'created 2026-10-21 08:00\s+modified 2026-10-21 08:00\s+.*IPM\.Note'
        $list | Should -Match 'after the export'
        $prev = Invoke-Entry $w ($common + @{ Folder = 'Expanded'; ItemId = '1' })
        $prev | Should -Match 'Item       : X1'
        $prev | Should -Match 'only the size and the dates are known before the export'
        $w.Mail.Requests().Count | Should -Be 0
    }

    It 'a token cached before the grant (no Mail.Read) is renewed once with the refresh token; still without it, the browse warns and falls back' {
        $w = New-World
        $text = Invoke-Entry $w @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; StorePath = (Join-Path $script:Tmp 'cli-none') } -Scopes @('MailboxFolder.Read', 'MailboxItem.Read')
        $text | Should -Match 'Mail\.Read is not in the cached token: asking for a new one'
        $text | Should -Match 'missing in the token: Mail\.Read'
        Should -Invoke Connect-MbtGraph -Times 1 -Exactly -ParameterFilter { $ForceRefresh }
        Should -Invoke Connect-MbtGraph -Times 2 -Exactly
        $w2 = New-World
        $ok = Invoke-Entry $w2 @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $ok | Should -Not -Match 'asking for a new one'
        $w3 = New-World
        $null = Invoke-Entry $w3 @{ Mode = 'Browse'; Mailbox = 'AuxArchive'; MailboxId = $script:Aux1; StorePath = (Join-Path $script:Tmp 'cli-none') } -Scopes @('MailboxFolder.Read', 'MailboxItem.Read')
        Should -Invoke Connect-MbtGraph -Times 4 -Exactly -Because 'an auxiliary archive needs no Mail.Read: one sign-in, no renewal'
    }

    It 'Browse -SaveSelection, then Export -SelectionFile (-WhatIf first): only the selection is exported' {
        $w = New-World
        $sel = Join-Path $script:Tmp 'cli-sel.json'
        $text = Invoke-Entry $w @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; Folder = 'Inbox'; IncludeSubfolders = $true; SaveSelection = $sel; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $text | Should -Match 'selection written to'
        $text | Should -Match '(?m)^ {6}\+ {2}Inbox: folder and its subfolders \[SameId\]'
        $text | Should -Match 'export it with: -Mode Export -SelectionFile'
        (Test-Path -LiteralPath $sel) | Should -BeTrue
        $store = Join-Path $script:Tmp 'cli-store'
        $dry = Invoke-Entry $w @{ Mode = 'Export'; SelectionFile = $sel; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-tmp'); WhatIf = $true }
        $dry | Should -Match 'selection .*: 1 entry for the Primary mailbox MBX:11111111'
        $dry | Should -Match 'WhatIf: the selection is valid, nothing was exported'
        (Test-Path (Join-Path $store 'catalog.db')) | Should -BeFalse
        $go = Invoke-Entry $w @{ Mode = 'Export'; SelectionFile = $sel; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-tmp'); Concurrency = 2 }
        $go | Should -Match 'export finished: 8 items exported'
        $go | Should -Not -Match '(?m)^ {6}x {2}\S'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try {
            $s.GetMailboxes()[0].Kind | Should -Be 'Primary'
            ($s.GetItemStats($script:Src) | Where-Object State -EQ 'Exported').Count | Should -Be 8
        } finally { $s.Dispose() }
        $again = Invoke-Entry $w @{ Mode = 'Export'; SelectionFile = $sel; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-tmp'); SkipInventory = $true }
        $again | Should -Match 'export finished: 0 items exported'
    }

    It 'Browse -SaveSelection with -ItemId selects those items (by row number), the plan says how they were matched' {
        $w = New-World
        $sel = Join-Path $script:Tmp 'cli-items.json'
        $text = Invoke-Entry $w @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; Folder = 'Inbox'; ItemId = @('1', '3'); SaveSelection = $sel; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $text | Should -Match 'Inbox: 2 item\(s\) \[SameId\] - 2 item\(s\) in 1 folder\(s\), .* - 2 item\(s\) by id'
        $json = Get-Content -LiteralPath $sel -Raw | ConvertFrom-Json
        (@($json.Entries[0].Items) | ForEach-Object Id) -join ',' | Should -Be 'M-I5,M-I3'
        $store = Join-Path $script:Tmp 'cli-store-items'
        $go = Invoke-Entry $w @{ Mode = 'Export'; SelectionFile = $sel; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-tmp2') }
        $go | Should -Match 'export finished: 2 items exported'
    }

    It 'a selection of Recoverable Items is exported like any other: the items sit under the top folder "Recoverable Items", with the other folders untouched' {
        $w = New-World; Add-Recoverable $w
        $sel = Join-Path $script:Tmp 'cli-rec.json'
        $text = Invoke-Entry $w @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; Folder = 'Recoverable Items/Deletions'; SaveSelection = $sel; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $text | Should -Match '(?m)^ {6}\+ {2}Recoverable Items/Deletions: folder \[SameId\] - 2 item\(s\) in 1 folder\(s\)'
        $global:MbtEntryExit | Should -Not -Be 1
        $store = Join-Path $script:Tmp 'cli-store-rec'
        $go = Invoke-Entry $w @{ Mode = 'Export'; SelectionFile = $sel; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-tmp3'); Concurrency = 2 }
        $go | Should -Match 'export finished: 2 items exported'
        $go | Should -Match "Recoverable Items: \d+ folder\(s\) recorded under 'Recoverable Items'. An import puts them in a normal folder named 'Restored deleted items'"
        $go | Should -Not -Match '(?m)^ {6}x {2}\S'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try {
            $paths = Get-StorePaths $s
            $paths | Should -Contain 'Recoverable Items/Deletions'
            $paths | Should -Not -Contain 'Deletions' -Because 'never mixed with the normal folders'
            ($s.GetItemStats($script:Src) | Where-Object State -EQ 'Exported').Count | Should -Be 2
        } finally { $s.Dispose() }
    }

    It 'Export -SelectionFile of the root of Recoverable Items takes the default folders only; -RecoverableItemsFolders chooses others' {
        $w = New-World; Add-Recoverable $w
        $sel = Join-Path $script:Tmp 'cli-rec-root.json'
        $null = Invoke-Entry $w @{ Mode = 'Browse'; PrimaryMailboxId = $script:Src; Folder = 'Recoverable Items'; IncludeSubfolders = $true; SaveSelection = $sel; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $store = Join-Path $script:Tmp 'cli-store-rec-root'
        $go = Invoke-Entry $w @{ Mode = 'Export'; SelectionFile = $sel; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-tmp3b'); Concurrency = 2 }
        $go | Should -Match 'export finished: 3 items exported' -Because 'Deletions 2 + Purges 1; Calendar Logging and SubstrateHolds are only exported when named'
        $store2 = Join-Path $script:Tmp 'cli-store-rec-root2'
        $go2 = Invoke-Entry $w @{ Mode = 'Export'; SelectionFile = $sel; RecoverableItemsFolders = @('Calendar Logging', 'SubstrateHolds'); StorePath = $store2; TempPath = (Join-Path $script:Tmp 'cli-tmp3c'); Concurrency = 2 }
        $go2 | Should -Match 'export finished: 4 items exported' -Because 'Calendar Logging 3 + SubstrateHolds 1'
        { Invoke-Entry $w @{ Mode = 'Export'; SelectionFile = $sel; IncludeRecoverableItems = $true; StorePath = (Join-Path $script:Tmp 'cli-x1'); TempPath = (Join-Path $script:Tmp 'cli-x2') } } | Should -Throw '*With -SelectionFile, select the folders of Recoverable Items*'
    }

    It 'Export -SelectionFile checks the mailbox against -Mailbox / -MailboxId and refuses -Folder' {
        $w = New-World
        $sel = Join-Path $script:Tmp 'cli-sel2.json'
        $null = Invoke-Entry $w @{ Mode = 'Browse'; Mailbox = 'Archive'; ArchiveMailboxId = $script:Arch; Folder = 'Old Mail'; SaveSelection = $sel; StorePath = (Join-Path $script:Tmp 'cli-none') }
        $base = @{ Mode = 'Export'; SelectionFile = $sel; StorePath = (Join-Path $script:Tmp 'cli-store-chk'); TempPath = (Join-Path $script:Tmp 'cli-tmp4') }
        { Invoke-Entry $w ($base + @{ Mailbox = 'Primary' }) } | Should -Throw '*does not match the selection*'
        { Invoke-Entry $w ($base + @{ MailboxId = $script:Src }) } | Should -Throw '*is not the mailbox of the selection*'
        { Invoke-Entry $w ($base + @{ Folder = 'Old Mail' }) } | Should -Throw '*exclusive*'
        $go = Invoke-Entry $w ($base + @{ Concurrency = 2 })
        $go | Should -Match 'exporting the Archive mailbox MBX:aaaaaaaa'
        $go | Should -Match 'export finished: 4 items exported'
    }
}


Describe 'Recoverable Items: export of a mailbox (ExportOptions.IncludeRecoverableItems)' {
    It 'exports the default set under one top folder; Calendar Logging and SubstrateHolds are left out; without the option nothing of Recoverable Items is asked' {
        $w = New-World; Add-Recoverable $w
        $store = New-TestStoreAt 'rec-default'
        try {
            $r = Invoke-OptionsExport -World $w -Store $store -Configure { param($o) $o.IncludeRecoverableItems = $true }
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Failed | Should -Be 0
            $r.Result.Exported | Should -Be 16 -Because '13 items of the normal folders + Deletions 2 + Purges 1'
            $r.Result.RecoverableFolders | Should -Be 5 -Because 'the root + Deletions, DiscoveryHolds, Purges, Versions'
            ($r.Messages -join "`n") | Should -Match 'Recoverable Items: 5 folder\(s\) read through the import-export API \(Deletions 2, DiscoveryHolds 0, Purges 1, Versions 0\); not asked for: Calendar Logging, SubstrateHolds'
            $paths = Get-StorePaths $store
            ($paths | Where-Object { $_ -like 'Recoverable Items*' }) -join '|' | Should -Be 'Recoverable Items|Recoverable Items/Deletions|Recoverable Items/DiscoveryHolds|Recoverable Items/Purges|Recoverable Items/Versions'
            ($paths | Where-Object { $_ -notlike 'Recoverable Items*' -and $_ -notlike '*/*' }) -join '|' | Should -Be 'Calendar|Empty|Inbox|Sent Items' -Because 'the folders of Recoverable Items are never at the top, among the normal ones'
            $tops = @($store.GetFolderTree($script:Src) | Where-Object DisplayName -EQ 'Recoverable Items')
            $tops.Count | Should -Be 1
            $tops[0].ParentFolderId | Should -BeNullOrEmpty
            $deletions = @($store.GetFolderTree($script:Src) | ForEach-Object { $_; $_.Children } | Where-Object Path -EQ 'Recoverable Items/Deletions')
            $deletions[0].Own.Exported | Should -Be 2
        } finally { $store.Dispose() }
        $w2 = New-World; Add-Recoverable $w2
        $plain = New-TestStoreAt 'rec-none'
        try {
            $r2 = Invoke-OptionsExport -World $w2 -Store $plain
            $r2.Result.Exported | Should -Be 13
            $r2.Result.RecoverableFolders | Should -Be 0
            $w2.Ie.RequestsOf('recoverableitemsroot') | Should -Be 0
            (Get-StorePaths $plain | Where-Object { $_ -like 'Recoverable*' }).Count | Should -Be 0
        } finally { $plain.Dispose() }
    }

    It 'an explicit list names the folders (any case, Calendar Logging and SubstrateHolds included); a name the mailbox does not have is a problem; a name that does not exist is refused' {
        $w = New-World; Add-Recoverable $w
        $store = New-TestStoreAt 'rec-explicit'
        try {
            $r = Invoke-OptionsExport -World $w -Store $store -Configure { param($o) $o.IncludeRecoverableItems = $true; $o.RecoverableItemsFolders = [string[]]@('calendar logging', 'SubstrateHolds') }
            $r.Result.Exported | Should -Be 17 -Because '13 + Calendar Logging 3 + SubstrateHolds 1'
            (Get-StorePaths $store | Where-Object { $_ -like 'Recoverable Items*' }) -join '|' | Should -Be 'Recoverable Items|Recoverable Items/Calendar Logging|Recoverable Items/SubstrateHolds'
        } finally { $store.Dispose() }
        $w2 = New-World; Add-Recoverable $w2 -Skip 'Versions'
        $store2 = New-TestStoreAt 'rec-missing'
        try {
            $r2 = Invoke-OptionsExport -World $w2 -Store $store2 -Configure { param($o) $o.IncludeRecoverableItems = $true; $o.RecoverableItemsFolders = [string[]]@('Deletions', 'Versions') }
            $r2.Result.Problems -join ' ' | Should -Match "Recoverable Items has no folder 'Versions' in this mailbox"
            $r2.Result.NeedsAnotherRun | Should -BeFalse
            $r2.Result.Exported | Should -Be 15
            # the default set with a folder the mailbox does not have is only said
            $store3 = New-TestStoreAt 'rec-missing-default'
            try {
                $r3 = Invoke-OptionsExport -World $w2 -Store $store3 -Configure { param($o) $o.IncludeRecoverableItems = $true }
                ($r3.Result.Problems -join ' ') | Should -Not -Match 'Recoverable Items has no folder'
                ($r3.Messages -join "`n") | Should -Match "no folder 'Versions' in this mailbox \(nothing to export from it\)"
            } finally { $store3.Dispose() }
        } finally { $store2.Dispose() }
        $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Handler $w.Mail))
        $o = [MailboxTransfer.ExportOptions]::new()
        $o.MailboxId = $script:Src; $o.TempPath = Join-Path $script:Tmp 'rx'; $o.IncludeRecoverableItems = $true; $o.RecoverableItemsFolders = [string[]]@('Trash')
        { [MailboxTransfer.ExportJob]::Start($api, (New-TestStoreAt 'rec-bad'), $o) } | Should -Throw "*'Trash' is not a folder of Recoverable Items. The folders are: Calendar Logging, Deletions, DiscoveryHolds, Purges, SubstrateHolds, Versions*"
    }

    It 'refuses what does not make sense: the option with a plan or with another folder, a list without the option' {
        $w = New-World; Add-Recoverable $w
        $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Handler $w.Mail))
        $store = New-TestStoreAt 'rec-refused'
        try {
            $o = [MailboxTransfer.ExportOptions]::new()
            $o.MailboxId = $script:Src; $o.TempPath = Join-Path $script:Tmp 'ry'
            $o.IncludeRecoverableItems = $true; $o.FolderPath = 'Inbox'
            { [MailboxTransfer.ExportJob]::Start($api, $store, $o) } | Should -Throw "*cannot be combined with the folder 'Inbox'*"
            $o.FolderPath = $null; $o.RecoverableItemsFolders = $null
            $o.IncludeRecoverableItems = $false; $o.RecoverableItemsFolders = [string[]]@('Purges')
            { [MailboxTransfer.ExportJob]::Start($api, $store, $o) } | Should -Throw '*RecoverableItemsFolders needs IncludeRecoverableItems*'
            $b = New-Browser $w
            $loc = Get-Loc Primary
            $sel = [MailboxTransfer.BrowseSelection]::For($loc)
            $null = $sel.AddFolder($b.GetFolderTree($loc).Find('Inbox'))
            $o.RecoverableItemsFolders = $null; $o.IncludeRecoverableItems = $true; $o.Plan = $b.ResolvePlan($sel)
            { [MailboxTransfer.ExportJob]::Start($api, $store, $o) } | Should -Throw '*IncludeRecoverableItems and Plan are exclusive*'
        } finally { $store.Dispose() }
    }

    It 'FolderPath below Recoverable Items asks for them by itself: only that folder is inventoried and exported' {
        $w = New-World; Add-Recoverable $w
        $store = New-TestStoreAt 'rec-path'
        try {
            $r = Invoke-OptionsExport -World $w -Store $store -Configure { param($o) $o.FolderPath = 'Recoverable Items/Deletions' }
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Exported | Should -Be 2
            $w.Ie.RequestsOf('/folders/F-INBOX/items') | Should -Be 0
            # the top folder itself takes the default set (not Calendar Logging, not SubstrateHolds), or the explicit list
            $store2 = New-TestStoreAt 'rec-path-root'
            try {
                $root = Invoke-OptionsExport -World $w -Store $store2 -Configure { param($o) $o.FolderPath = 'Recoverable Items' }
                $root.Result.Exported | Should -Be 3
                ($root.Messages -join "`n") | Should -Match 'not asked for: Calendar Logging, SubstrateHolds'
            } finally { $store2.Dispose() }
            $store3 = New-TestStoreAt 'rec-path-root-list'
            try {
                $listed = Invoke-OptionsExport -World $w -Store $store3 -Configure { param($o) $o.FolderPath = 'Recoverable Items'; $o.RecoverableItemsFolders = [string[]]@('Calendar Logging') }
                $listed.Result.Exported | Should -Be 3
                (Get-StorePaths $store3 | Where-Object { $_ -like 'Recoverable Items/*' }) -join '|' | Should -Be 'Recoverable Items/Calendar Logging'
            } finally { $store3.Dispose() }
            $bad = Invoke-OptionsExport -World (New-World) -Store (New-TestStoreAt 'rec-path-none') -Configure { param($o) $o.FolderPath = 'Recoverable Items/Deletions' }
            $bad.Job.Error.Message | Should -Match "Folder 'Recoverable Items/Deletions' not found in this mailbox"
        } finally { $store.Dispose() }
    }

    It 'a plan that names the root with its subfolders exports the default folders; a plan that names Calendar Logging itself gets it' {
        $w = New-World; Add-Recoverable $w; $b = New-Browser $w
        $b.IncludeRecoverableItems = $true
        $loc = Get-Loc Primary
        $tree = $b.GetFolderTree($loc)
        $sel = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel.AddFolder($tree.Find('Recoverable Items'), $true)
        $store = New-TestStoreAt 'rec-plan'
        try {
            $r = Invoke-PlanExport -World $w -Store $store -Plan ($b.ResolvePlan($sel))
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Exported | Should -Be 3 -Because 'Deletions 2 + Purges 1; the plan did not name Calendar Logging or SubstrateHolds'
            $w.Ie.RequestsOf('/folders/R-CAL/items') | Should -Be 0
        } finally { $store.Dispose() }
        $sel2 = [MailboxTransfer.BrowseSelection]::For($loc)
        $null = $sel2.AddFolder($tree.Find('Recoverable Items/Calendar Logging'))
        $store2 = New-TestStoreAt 'rec-plan2'
        try {
            $r2 = Invoke-PlanExport -World $w -Store $store2 -Plan ($b.ResolvePlan($sel2))
            $r2.Result.Exported | Should -Be 3
            (Get-StorePaths $store2 | Where-Object { $_ -like 'Recoverable Items*' }).Count | Should -BeGreaterThan 1
        } finally { $store2.Dispose() }
    }

    It 'the main archive and an auxiliary archive export their Recoverable Items too' {
        $w = New-World
        Add-Recoverable $w -Mailbox $script:Arch -Tag 'A-'; Add-Recoverable $w -Mailbox $script:Aux1 -Tag 'X-'
        $store = New-TestStoreAt 'rec-archives'
        try {
            $r = Invoke-OptionsExport -World $w -Store $store -Mailbox $script:Arch -Kind Archive -Configure { param($o) $o.IncludeRecoverableItems = $true }
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Exported | Should -Be 10 -Because 'Old Mail 4 + Deleted Items 2 + Mixed 3 + Deletions 1'
            (Get-StorePaths $store $script:Arch | Where-Object { $_ -like 'Recoverable Items*' }).Count | Should -Be 5
            $x = Invoke-OptionsExport -World $w -Store $store -Mailbox $script:Aux1 -Kind AuxArchive -Configure { param($o) $o.IncludeRecoverableItems = $true }
            $x.Job.Error | Should -BeNullOrEmpty
            $x.Result.Exported | Should -Be 5 -Because 'Expanded 3 + Sub 1 + Deletions 1'
            (Get-StorePaths $store $script:Aux1 | Where-Object { $_ -like 'Recoverable Items*' }).Count | Should -Be 5
        } finally { $store.Dispose() }
    }

}

Describe 'Recoverable Items: restore, and a mailbox that cannot be reached' {
    It 'an import never goes back into Recoverable Items: its folders are created in "Restored deleted items" (below the target root, or at the top), a mapping wins' {
        $o = [MailboxTransfer.ImportOptions]::new()
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Recoverable Items') | Should -Be 'Restored deleted items'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Recoverable Items/Deletions') | Should -Be 'Restored deleted items/Deletions'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'recoverable items/Purges/Sub') | Should -Be 'Restored deleted items/Purges/Sub'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Inbox/Recoverable Items') | Should -Be 'Inbox/Recoverable Items' -Because 'only the top folder of the export'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Recoverable Items Archive') | Should -Be 'Recoverable Items Archive'
        $o.TargetRoot = 'Imported 2026'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Recoverable Items/Deletions') | Should -Be 'Imported 2026/Restored deleted items/Deletions'
        $o.Mapping = [System.Collections.Generic.Dictionary[string, string]]::new(); $o.Mapping.Add('Recoverable Items', 'Back')
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Recoverable Items/Deletions') | Should -Be 'Imported 2026/Back/Deletions'
        $p = [MailboxTransfer.PstImportOptions]::new()
        [MailboxTransfer.PstImportPlanner]::MapPath($p, 'Imported', 'Recoverable Items/Deletions') | Should -Be 'Imported/Restored deleted items/Deletions'
        [MailboxTransfer.PstImportPlanner]::MapPath($p, '', 'Recoverable Items/Deletions') | Should -Be 'Restored deleted items/Deletions'
        [MailboxTransfer.PstImportPlanner]::MapPath($p, 'Imported', 'Inbox') | Should -Be 'Imported/Inbox'
    }

    It 'importing an export of Recoverable Items creates the normal folder "Restored deleted items" and puts the items there' {
        $w = New-World; Add-Recoverable $w
        $store = New-TestStoreAt 'rec-restore'
        try {
            $r = Invoke-OptionsExport -World $w -Store $store -Configure { param($o) $o.FolderPath = 'Recoverable Items/Deletions' }
            $r.Result.Exported | Should -Be 2
            $client = New-FakeClient -Handler $w.Mail
            $api = [MailboxTransfer.MailboxApi]::new($client)
            $io = [MailboxTransfer.ImportOptions]::new()
            $io.SourceMailboxId = $script:Src; $io.TargetMailboxId = $script:Arch; $io.Concurrency = 1; $io.AutoReconcile = $false; $io.SourceFolderPath = 'Recoverable Items/Deletions'
            $plan = [MailboxTransfer.ImportPlanner]::Build($store, $io, $api.GetFolderTree($script:Arch))
            ($plan.PathsToCreate -join '|') | Should -Be 'Restored deleted items|Restored deleted items/Deletions'
            ($plan.Folders | Where-Object Items -GT 0 | ForEach-Object TargetPath) -join '|' | Should -Be 'Restored deleted items/Deletions'
            $job = [MailboxTransfer.ImportJob]::Start($api, $store, $io)
            $deadline = [DateTime]::UtcNow.AddSeconds(60)
            while (-not $job.Wait(100)) { if ([DateTime]::UtcNow -gt $deadline) { $job.Cancel(); throw 'the import did not finish in time' } }
            $job.Error | Should -BeNullOrEmpty
            $job.Result.Done | Should -Be 2
            $names = @($w.Ie.FoldersOf($script:Arch) | ForEach-Object Name)
            $names | Should -Contain 'Restored deleted items'
            $names | Should -Not -Contain 'Recoverable Items'
            $restored = $w.Ie.FoldersOf($script:Arch) | Where-Object Name -EQ 'Deletions'
            @($w.Ie.ItemsOf($script:Arch, $restored.Id)).Count | Should -Be 2
        } finally { $store.Dispose() }
    }

    It 'a soft-deleted or inactive mailbox (bare 500, or 404 ErrorInvalidUser on the folder listing) is explained: not reachable by the API, and the supported routes' {
        $bare = 'MBX:cccccccc-0000-0000-0000-00000000000c@' + $script:Tenant
        $invalid = 'MBX:dddddddd-0000-0000-0000-00000000000d@' + $script:Tenant
        $w = New-World
        foreach ($m in $bare, $invalid) { $null = $w.Ie.AddMailbox($m) }
        $null = $w.Ie.BareFailureMailboxes.Add($bare); $null = $w.Ie.InvalidUserMailboxes.Add($invalid)
        $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Handler $w.Mail))
        foreach ($m in $bare, $invalid) {
            $err = $null
            try { $null = $api.GetFolderTree($m) } catch { $err = $_.Exception; if ($err.InnerException -is [MailboxTransfer.MailboxNotReachableException]) { $err = $err.InnerException } }
            $err | Should -BeOfType ([MailboxTransfer.MailboxNotReachableException])
            $err.MailboxId | Should -Be $m
            $err.Message | Should -Match 'cannot be reached by the mailbox import-export API'
            $err.Message | Should -Match 'soft-deleted mailbox \(its user was deleted less than 30 days ago\)'
            $err.Message | Should -Match 'inactive mailbox'
            $err.Message | Should -Match 'Undo-SoftDeletedMailbox'
            $err.Message | Should -Match 'New-MailboxRestoreRequest'
            $err.Message | Should -Match 'export it once it is active again|export that active mailbox'
            $err.InnerException | Should -BeOfType ([MailboxTransfer.GraphException])
        }
        $api.GetRootFolders($script:Src).Count | Should -BeGreaterThan 0
        { $api.GetRootFolders($bare) } | Should -Throw '*soft-deleted*'
        # the export says it as its error, the browse as an exception
        $store = New-TestStoreAt 'unreachable'
        try {
            $r = Invoke-OptionsExport -World $w -Store $store -Mailbox $bare
            $r.Job.Error.Message | Should -Match 'Undo-SoftDeletedMailbox'
        } finally { $store.Dispose() }
        $b = New-Browser $w
        { $b.GetFolderTree([MailboxTransfer.MailboxBrowser]::Location('Primary', $invalid, 'test')) } | Should -Throw '*cannot be reached by the mailbox import-export API*'
        # an answer that explains itself, or another refusal, is not this
        [MailboxTransfer.MailboxApi]::IsMailboxNotRoutable([MailboxTransfer.GraphException]::new(500, '', 'GET https://x -> 500 ', 'x')) | Should -BeTrue
        [MailboxTransfer.MailboxApi]::IsMailboxNotRoutable([MailboxTransfer.GraphException]::new(404, 'ErrorInvalidUser', 'GET https://x -> 404 ErrorInvalidUser: nope', 'x')) | Should -BeTrue
        [MailboxTransfer.MailboxApi]::IsMailboxNotRoutable([MailboxTransfer.GraphException]::new(500, 'ErrorInternalServerError', 'GET https://x -> 500 ErrorInternalServerError: boom', 'x')) | Should -BeFalse
        [MailboxTransfer.MailboxApi]::IsMailboxNotRoutable([MailboxTransfer.GraphException]::new(404, 'ErrorItemNotFound', 'GET https://x -> 404 ErrorItemNotFound: gone', 'x')) | Should -BeFalse
        [MailboxTransfer.MailboxApi]::IsMailboxNotRoutable([MailboxTransfer.GraphException]::new(403, 'ErrorAccessDenied', 'GET https://x -> 403 ErrorAccessDenied: no', 'x')) | Should -BeFalse
    }

}

Describe 'Recoverable Items: the entry script' {
    BeforeAll {
        $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
        function Invoke-Entry {
            param([Parameter(Mandatory)]$World, [hashtable]$Parameters = @{}, [string[]]$Scopes = @('Mail.Read', 'MailboxFolder.Read', 'MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'))
            $handler = $World.Mail
            $global:MbtEntryFactory = {
                $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
                [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $handler)
            }.GetNewClosure()
            $global:MbtEntryScopes = $Scopes
            $global:LASTEXITCODE = $null
            $text = & $script:Entry @Parameters *>&1 | Out-String
            $global:MbtEntryExit = $LASTEXITCODE
            return $text
        }
    }
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph (& $global:MbtEntryFactory)
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = $global:MbtEntryScopes; ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtEntryFactory, MbtEntryExit, MbtEntryScopes -Scope Global -ErrorAction SilentlyContinue }
    It '-Mode Export -IncludeRecoverableItems: -WhatIf plans the default folders, the export writes them, the entry script refuses the combinations' {
        $w = New-World; Add-Recoverable $w
        $store = Join-Path $script:Tmp 'cli-rec-all'
        $dry = Invoke-Entry $w @{ Mode = 'Export'; PrimaryMailboxId = $script:Src; IncludeRecoverableItems = $true; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-rec-tmp'); WhatIf = $true }
        $dry | Should -Match 'Deletions\s+2 item\(s\)'
        $dry | Should -Match 'Purges\s+1 item\(s\)'
        $dry | Should -Not -Match 'Calendar Logging'
        $dry | Should -Match 'total: 16 item\(s\)'
        (Test-Path (Join-Path $store 'catalog.db')) | Should -BeFalse
        $plain = Invoke-Entry $w @{ Mode = 'Export'; PrimaryMailboxId = $script:Src; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-rec-tmp'); WhatIf = $true }
        $plain | Should -Not -Match 'Deletions'
        $go = Invoke-Entry $w @{ Mode = 'Export'; PrimaryMailboxId = $script:Src; IncludeRecoverableItems = $true; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-rec-tmp'); Concurrency = 2 }
        $go | Should -Match 'export finished: 16 items exported'
        $go | Should -Not -Match '(?m)^ {6}x {2}\S'
        $explicit = Invoke-Entry $w @{ Mode = 'Export'; PrimaryMailboxId = $script:Src; IncludeRecoverableItems = $true; RecoverableItemsFolders = @('Calendar Logging'); StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-rec-tmp'); WhatIf = $true }
        $explicit | Should -Match 'Calendar Logging\s+3 item\(s\)'
        $explicit | Should -Not -Match 'Deletions'
        { Invoke-Entry $w @{ Mode = 'Export'; PrimaryMailboxId = $script:Src; RecoverableItemsFolders = @('Purges'); StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-rec-tmp') } } | Should -Throw '*-RecoverableItemsFolders needs -IncludeRecoverableItems*'
        { Invoke-Entry $w @{ Mode = 'Export'; PrimaryMailboxId = $script:Src; IncludeRecoverableItems = $true; Folder = 'Inbox'; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-rec-tmp') } } | Should -Throw '*cannot be combined with -Folder*'
        { Invoke-Entry $w @{ Mode = 'Export'; PrimaryMailboxId = $script:Src; RecoverableItemsFolders = @('Trash'); IncludeRecoverableItems = $true; StorePath = $store; TempPath = (Join-Path $script:Tmp 'cli-rec-tmp') } } | Should -Throw "*'Trash' is not a folder of Recoverable Items*"
        { Invoke-Entry $w @{ Mode = 'Status'; IncludeRecoverableItems = $true } } | Should -Throw '*belong to -Mode Export*'
    }

    It 'the entry script stops an export of an unreachable mailbox with the explanation' {
        $bare = 'MBX:cccccccc-0000-0000-0000-00000000000c@' + $script:Tenant
        $w = New-World
        $null = $w.Ie.AddMailbox($bare); $null = $w.Ie.BareFailureMailboxes.Add($bare)
        { $null = Invoke-Entry $w @{ Mode = 'Export'; PrimaryMailboxId = $bare; StorePath = (Join-Path $script:Tmp 'cli-unreach'); TempPath = (Join-Path $script:Tmp 'cli-unreach-tmp') } } | Should -Throw '*Undo-SoftDeletedMailbox*'
    }
}
