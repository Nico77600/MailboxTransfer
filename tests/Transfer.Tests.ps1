#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the export and import engines (Engine.Export.cs, Engine.Import.cs, Engine.Transfer.cs).
    Nothing here reaches the network: the engines talk to FakeMailboxService (tests\FakeMailboxService.cs), an in-memory
    mailbox import and export service that can throttle, fail and measure the concurrency it sees.

    Run:  Invoke-Pester -Path .\tests\Transfer.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-transfer-tests-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp
    $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:Src = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
    $script:Dst = "MBX:99999999-8888-7777-6666-555555555555@$($script:Tenant)"

    function New-Message([string]$Subject, [int]$Padding = 0) {
        $b = [MailboxTransfer.Tests.FxBuilder]::new()
        $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $Subject)
        if ($Padding -gt 0) { $null = $b.Bin(0x7000, [byte[]]::new($Padding)) }
        return $b.ToArray()
    }

    function Get-Sha([byte[]]$Bytes) { [BitConverter]::ToString([Security.Cryptography.SHA256]::HashData($Bytes)).Replace('-', '').ToLowerInvariant() }

    function New-FakeClient {
        param([Parameter(Mandatory)]$Service, [double]$DelayScale = 0)
        $o = [MailboxTransfer.GraphOptions]::new()
        $o.DelayScale = $DelayScale
        $o.UseJitter = $false
        return [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
    }

    # A mailbox: Inbox (6 items), Inbox/Projects (4 items), Empty (no item).
    function New-SourceService {
        param([int]$Padding = 0, [int]$InboxItems = 6, [int]$ProjectItems = 4)
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Src)
        $null = $svc.AddFolder($script:Src, 'F-INBOX', 'Inbox')
        $null = $svc.AddFolder($script:Src, 'F-PROJ', 'Projects', 'F-INBOX')
        $null = $svc.AddFolder($script:Src, 'F-EMPTY', 'Empty')
        foreach ($i in 1..$InboxItems | Where-Object { $InboxItems -gt 0 }) { $null = $svc.AddItem($script:Src, 'F-INBOX', "I$i", (New-Message "Inbox $i" $Padding), "ck-I$i") }
        foreach ($i in 1..$ProjectItems | Where-Object { $ProjectItems -gt 0 }) { $null = $svc.AddItem($script:Src, 'F-PROJ', "P$i", (New-Message "Project $i" $Padding), "ck-P$i") }
        return $svc
    }

    function New-TestStoreAt([string]$Name) {
        return [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests')
    }

    function Invoke-ExportJob {
        <# Runs an ExportJob to its end and returns the job, the messages it wrote and the client. #>
        param(
            [Parameter(Mandatory)]$Service, [Parameter(Mandatory)]$Store, [int]$Concurrency = 2, [double]$DelayScale = 0, [string]$Kind = 'Primary',
            [scriptblock]$Options, [int]$TimeoutSeconds = 60, [string]$Mailbox = $script:Src
        )
        $client = New-FakeClient -Service $Service -DelayScale $DelayScale
        $o = [MailboxTransfer.ExportOptions]::new()
        $o.MailboxId = $Mailbox; $o.Kind = $Kind; $o.Upn = 'adele@contoso.test'; $o.Concurrency = $Concurrency
        $o.TempPath = Join-Path $script:Tmp ('tmp-' + [guid]::NewGuid().ToString('N'))
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

    $script:Arch = "MBX:aaaaaaaa-0000-0000-0000-00000000000a@$($script:Tenant)"
    $script:Aux1 = "MBX:bbbbbbbb-0000-0000-0000-00000000000b@$($script:Tenant)"
    # The main archive lists 'Inbox' (3 items) and the placeholder 'Old mail' with its child 'Y2019' (nothing in them): their 4 + 2 items live in the
    # auxiliary archive, which lists the same folder ids. Folder requests for them answer 308, exportItems answers ErrorArchiveFolderMovedPermanently.
    function Add-ExpandedArchive($svc) {
        $null = $svc.AddMailbox($script:Arch); $null = $svc.AddMailbox($script:Aux1)
        $null = $svc.AddFolder($script:Arch, 'F-INBOX', 'Inbox')
        $null = $svc.AddFolder($script:Arch, 'F-MOVED', 'Old mail')
        $null = $svc.AddFolder($script:Arch, 'F-MOVED-SUB', 'Y2019', 'F-MOVED')
        $null = $svc.AddFolder($script:Aux1, 'F-MOVED', 'Old mail')
        $null = $svc.AddFolder($script:Aux1, 'F-MOVED-SUB', 'Y2019', 'F-MOVED')
        foreach ($i in 1..3) { $null = $svc.AddItem($script:Arch, 'F-INBOX', "A$i", (New-Message "Archive $i"), "ck-A$i") }
        foreach ($i in 1..4) { $null = $svc.AddItem($script:Aux1, 'F-MOVED', "M$i", (New-Message "Old $i"), "ck-M$i") }
        foreach ($i in 1..2) { $null = $svc.AddItem($script:Aux1, 'F-MOVED-SUB', "S$i", (New-Message "Old 2019 $i"), "ck-S$i") }
        $svc.MovedFolderToMailbox['F-MOVED'] = $script:Aux1
        $svc.MovedFolderToMailbox['F-MOVED-SUB'] = $script:Aux1
        $svc.RedirectFolders = $true
    }
    function New-ExpandedArchive { $svc = [MailboxTransfer.Tests.FakeMailboxService]::new(); Add-ExpandedArchive $svc; return $svc }
    function Find-Node($nodes, [string]$Name) {
        foreach ($n in $nodes) { if ($n.DisplayName -eq $Name) { return $n }; $f = Find-Node $n.Children $Name; if ($f) { return $f } }
    }

    $script:Aux = "MBX:aaaaaaaa-1111-2222-3333-444444444444@$($script:Tenant)"
    function Add-Targets($svc) {
        $null = $svc.AddMailbox($script:Dst)
        $null = $svc.AddFolder($script:Dst, 'D-INBOX', 'Inbox')
        $null = $svc.AddMailbox($script:Aux)
        $null = $svc.AddFolder($script:Aux, 'D-INBOX', 'Inbox')
    }
    function Get-FolderPaths($svc, $mailbox) {
        $folders = $svc.FoldersOf($mailbox)
        $byId = @{}; foreach ($f in $folders) { $byId[$f.Id] = $f }
        @($folders | ForEach-Object {
            $p = $_.Name; $c = $_
            while ($byId.ContainsKey($c.ParentId)) { $c = $byId[$c.ParentId]; $p = "$($c.Name)/$p" }
            $p
        }) | Sort-Object
    }
    function New-ExportedStore($svc, [string]$Name) {
        $store = New-TestStoreAt $Name
        $null = Invoke-ExportJob -Service $svc -Store $store
        return $store
    }
    function Invoke-ImportJob {
        # The automatic reconciliation is off unless -Reconcile: most tests look at the raw states (Unknown stays Unknown).
        param([Parameter(Mandatory)]$Service, [Parameter(Mandatory)]$Store, [int]$Concurrency = 2, [double]$DelayScale = 0, [scriptblock]$Options, [string]$Name = 'imp', [int]$TimeoutSeconds = 60, [switch]$Reconcile)
        $client = New-FakeClient -Service $Service -DelayScale $DelayScale
        $o = [MailboxTransfer.ImportOptions]::new()
        $o.SourceMailboxId = $script:Src; $o.TargetMailboxId = $script:Dst; $o.Concurrency = $Concurrency; $o.JobName = $Name
        if ($Reconcile) { $o.ReconcileDelaySeconds = 0 } else { $o.AutoReconcile = $false }
        if ($Options) { & $Options $o }
        $job = [MailboxTransfer.ImportJob]::Start([MailboxTransfer.MailboxApi]::new($client), $Store, $o)
        $messages = [System.Collections.Generic.List[string]]::new()
        $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
        while (-not $job.Wait(100)) {
            foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
            if ([DateTime]::UtcNow -gt $deadline) { $job.Cancel(); throw 'the import job did not finish in time' }
        }
        foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
        return [pscustomobject]@{ Job = $job; Messages = $messages; Client = $client; Result = $job.Result }
    }
}

AfterAll {
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Export engine: a whole mailbox into the store' {
    It 'records the tree, lists every folder page by page, exports everything and fills the previews' {
        $svc = New-SourceService
        $svc.PageSize = 3
        $store = New-TestStoreAt 'whole'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 2
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Completed | Should -BeTrue
            $r.Result.Folders | Should -Be 3
            $r.Result.Inventoried | Should -Be 10
            $r.Result.Exported | Should -Be 10
            $r.Result.Failed | Should -Be 0
            $r.Result.PreviewWarnings | Should -Be 0
            $svc.RequestsOf('skip=') | Should -BeGreaterThan 0 -Because 'the nextLink of the pages is followed'
            $store.GetMailboxes()[0].Kind | Should -Be 'Primary'
            $tree = $store.GetFolderTree($script:Src)
            ($tree | Where-Object DisplayName -EQ 'Inbox').Children[0].Path | Should -Be 'Inbox/Projects'
            ($tree | Where-Object DisplayName -EQ 'Inbox').Total.Exported | Should -Be 10
            $stats = $store.GetItemStats($script:Src)
            ($stats | Where-Object State -EQ 'Exported').Count | Should -Be 10
            ($stats | Where-Object State -EQ 'Pending').Count | Should -Be 0
            foreach ($item in $svc.ItemsOf($script:Src)) {
                $stored = $store.GetItem($script:Src, $item.Id)
                $stored.State | Should -Be 'Exported'
                $stored.Sha256 | Should -Be (Get-Sha $item.Data)
                $stored.ChangeKey | Should -Be $item.ChangeKey
                $store.GetPreview($stored.Key).Subject | Should -Match "Inbox|Project"
            }
            (Get-ChildItem (Join-Path $script:Tmp 'whole\packs') -File).Count | Should -BeGreaterThan 0
            @($store.VerifyAll([Threading.CancellationToken]::None)).Count | Should -Be 0
            $r.Messages | Should -Contain 'tree: 3 folders'
        } finally { $store.Dispose() }
    }
    It 'marks an archive mailbox as Archive and exports only the folder asked for' {
        $svc = New-SourceService
        $store = New-TestStoreAt 'archive'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Kind Archive -Options { param($o) $o.FolderPath = 'Inbox/Projects' }
            $r.Result.Exported | Should -Be 4
            $store.GetMailboxes()[0].Kind | Should -Be 'Archive'
            ($store.GetItemStats($script:Src) | Where-Object State -EQ 'Pending').Count | Should -Be 0
            $store.GetItem($script:Src, 'I1') | Should -BeNullOrEmpty -Because 'the other folders were not listed'
            @($store.GetFolders($script:Src) | ForEach-Object Path | Sort-Object) | Should -Be @('Inbox', 'Inbox/Projects') -Because 'the store keeps the folder asked for and the path to it, not the empty rest of the tree'
        } finally { $store.Dispose() }
    }
    It 'records the empty folders of a folder asked for with its subfolders, and none of the others' {
        $svc = New-SourceService -InboxItems 0 -ProjectItems 0
        $store = New-TestStoreAt 'emptyscope'
        try {
            $null = Invoke-ExportJob -Service $svc -Store $store -Options { param($o) $o.FolderPath = 'Inbox' }
            @($store.GetFolders($script:Src) | ForEach-Object Path | Sort-Object) | Should -Be @('Inbox', 'Inbox/Projects')
        } finally { $store.Dispose() }
    }
}

Describe 'Export engine: batches' {
    It 'asks at most 20 items per exportItems request' {
        $svc = New-SourceService -InboxItems 45 -ProjectItems 0
        $store = New-TestStoreAt 'b20'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Exported | Should -Be 45
            @($svc.ExportCalls | ForEach-Object Items | Sort-Object -Descending) | Should -Be @(20, 20, 5)
        } finally { $store.Dispose() }
    }
    It 'bounds the announced bytes of a request (an item over the bound goes alone)' {
        $svc = New-SourceService -InboxItems 5 -ProjectItems 0 -Padding 100000
        $null = $svc.AddItem($script:Src, 'F-INBOX', 'BIG', (New-Message 'Big' 300000), 'ck-big')
        $store = New-TestStoreAt 'bbytes'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 1 -Options { param($o) $o.BatchBytes = 250000 }
            $r.Result.Exported | Should -Be 6
            foreach ($c in $svc.ExportCalls) { if ($c.Items -gt 1) { $c.Bytes | Should -BeLessOrEqual 250000 } }
            @($svc.ExportCalls | Where-Object { $_.Items -eq 1 -and $_.Bytes -gt 250000 }).Count | Should -Be 1
            $svc.ExportCalls.Count | Should -BeGreaterOrEqual 3
        } finally { $store.Dispose() }
    }
    It 'asks again item by item when a request of several items fails' {
        $svc = New-SourceService -InboxItems 6 -ProjectItems 0
        $svc.FailExportAbove = 1
        $store = New-TestStoreAt 'split'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Exported | Should -Be 6
            $r.Result.Failed | Should -Be 0
            ($r.Messages | Where-Object { $_ -like '*asking again item by item*' }).Count | Should -BeGreaterThan 0
        } finally { $store.Dispose() }
    }
}

Describe 'Export engine: resume' {
    It 'exports only what is left: Failed items are tried again, Exported ones are not asked for again' {
        $svc = New-SourceService
        $null = $svc.BadItemIds.Add('I2'); $null = $svc.BadItemIds.Add('P3')
        $store = New-TestStoreAt 'resume'
        try {
            $first = Invoke-ExportJob -Service $svc -Store $store
            $first.Result.Exported | Should -Be 8
            $first.Result.Failed | Should -Be 2
            $store.GetItem($script:Src, 'I2').State | Should -Be 'Failed'
            $store.GetItem($script:Src, 'I2').LastError | Should -BeLike '*ErrorItemNotFound*'
            $svc.BadItemIds.Clear()
            $svc.ExportCalls.Clear()
            $second = Invoke-ExportJob -Service $svc -Store $store
            $second.Result.RequeuedFailed | Should -Be 2
            $second.Result.PendingAtStart | Should -Be 2
            $second.Result.Exported | Should -Be 2
            ($svc.ExportCalls | Measure-Object Items -Sum).Sum | Should -Be 2
            ($store.GetItemStats($script:Src) | Where-Object State -EQ 'Exported').Count | Should -Be 10
            $third = Invoke-ExportJob -Service $svc -Store $store
            $third.Result.PendingAtStart | Should -Be 0
            $svc.ExportCalls.Count | Should -Be 1 -Because 'the third run asked for nothing'
        } finally { $store.Dispose() }
    }
    It 'leaves an item Failed once it was tried MaxAttempts times, and SkipInventory does not list again' {
        $svc = New-SourceService
        $null = $svc.BadItemIds.Add('I2')
        $store = New-TestStoreAt 'resume2'
        try {
            $null = Invoke-ExportJob -Service $svc -Store $store
            $svc.BadItemIds.Clear()
            $listings = $svc.RequestsOf('/items')
            $r = Invoke-ExportJob -Service $svc -Store $store -Options { param($o) $o.MaxAttempts = 1; $o.SkipInventory = $true }
            $r.Result.RequeuedFailed | Should -Be 0
            $store.GetItem($script:Src, 'I2').State | Should -Be 'Failed'
            $svc.RequestsOf('/items') | Should -Be $listings -Because 'SkipInventory'
            $r = Invoke-ExportJob -Service $svc -Store $store -Options { param($o) $o.MaxAttempts = 2; $o.SkipInventory = $true }
            $r.Result.Exported | Should -Be 1
            $store.GetItem($script:Src, 'I2').State | Should -Be 'Exported'
        } finally { $store.Dispose() }
    }
    It 'exports again an item whose changeKey changed, and only that one' {
        $svc = New-SourceService
        $store = New-TestStoreAt 'ckey'
        try {
            $null = Invoke-ExportJob -Service $svc -Store $store
            $old = $store.GetItem($script:Src, 'P2').Sha256
            $newData = New-Message 'Project 2 edited' 50
            $svc.ChangeItem($script:Src, 'P2', $newData, 'ck-P2-v2')
            $svc.ExportCalls.Clear()
            $r = Invoke-ExportJob -Service $svc -Store $store
            $r.Result.PendingAtStart | Should -Be 1
            ($svc.ExportCalls | Measure-Object Items -Sum).Sum | Should -Be 1
            $now = $store.GetItem($script:Src, 'P2')
            $now.State | Should -Be 'Exported'
            $now.Sha256 | Should -Be (Get-Sha $newData)
            $now.Sha256 | Should -Not -Be $old
            $now.ChangeKey | Should -Be 'ck-P2-v2'
            $store.GetPreview($now.Key).Subject | Should -Be 'Project 2 edited'
        } finally { $store.Dispose() }
    }
}

Describe 'Export engine: stopping' {
    It 'stops the job on 403 (permission missing) with the error of the service' {
        $svc = New-SourceService
        $svc.ForbidExport = $true
        $store = New-TestStoreAt 'forbid'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store
            $r.Job.Error | Should -Not -BeNullOrEmpty
            $r.Job.Error.StatusCode | Should -Be 403
            $r.Result.Completed | Should -BeFalse
        } finally { $store.Dispose() }
    }
    It 'stops after 25 items failed in a row instead of marking everything Failed' {
        $svc = New-SourceService -InboxItems 60 -ProjectItems 0
        foreach ($i in 1..60) { $null = $svc.BadItemIds.Add("I$i") }
        $store = New-TestStoreAt 'breaker'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 1
            $r.Job.Error.Message | Should -BeLike '*25 items failed in a row*'
            ($store.GetItemStats($script:Src) | Where-Object State -EQ 'Pending').Count | Should -BeGreaterThan 0
        } finally { $store.Dispose() }
    }
    It 'refuses a concurrency over 4 and a batch over 20' {
        $o = [MailboxTransfer.ExportOptions]::new()
        $o.MailboxId = $script:Src; $o.TempPath = $script:Tmp; $o.Concurrency = 5
        $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Service (New-SourceService)))
        $store = New-TestStoreAt 'limits'
        try {
            { [MailboxTransfer.ExportJob]::Start($api, $store, $o) } | Should -Throw '*between 1 and 4*'
            $o.Concurrency = 4; $o.BatchItems = 21
            { [MailboxTransfer.ExportJob]::Start($api, $store, $o) } | Should -Throw '*BatchItems*'
        } finally { $store.Dispose() }
    }
}

Describe 'Throttling: compliance, not evasion' {
    It 'honours Retry-After exactly and makes every worker wait for it' {
        $svc = New-SourceService -InboxItems 40 -ProjectItems 0
        $svc.LatencyMs = 20
        $store = New-TestStoreAt 'gate'
        try {
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $svc.Throttle(1, 2, 'exportItems')
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 4 -DelayScale 1 -Options { param($o) $o.BatchItems = 2 }
            $sw.Stop()
            $r.Result.Exported | Should -Be 40
            $r.Client.Throttled429 | Should -Be 1
            $ev = @($r.Job.GetThrottleEvents())
            $ev.Count | Should -Be 1
            $ev[0].Status | Should -Be 429
            $ev[0].HadRetryAfter | Should -BeTrue
            $ev[0].WaitSeconds | Should -Be 2
            $sw.Elapsed.TotalSeconds | Should -BeGreaterOrEqual 1.9
            $at = $svc.ThrottledAt[0]
            $inside = @($svc.Times | Where-Object { $_ -gt $at.AddMilliseconds(100) -and $_ -lt $at.AddSeconds(1.9) })
            $inside.Count | Should -Be 0 -Because 'no request is sent while the Retry-After of one of them is running'
            $r.Messages | Where-Object { $_ -like "throttled: *429*Retry-After 2 s" } | Should -Not -BeNullOrEmpty
        } finally { $store.Dispose() }
    }
    It 'halves the concurrency on 429, once per burst, never goes over 4, and ramps back up' {
        $l = [MailboxTransfer.AdaptiveLimiter]::new(4)
        $l.RampAfter = [TimeSpan]::FromMilliseconds(150)
        $l.ReduceCooldown = [TimeSpan]::FromMilliseconds(100)
        $l.Limit | Should -Be 4
        $l.OnThrottle() | Should -BeLike 'concurrency lowered from 4 to 2*'
        $l.OnThrottle() | Should -BeNullOrEmpty -Because 'several workers refused together count once'
        $l.Limit | Should -Be 2
        Start-Sleep -Milliseconds 150
        $l.OnThrottle() | Should -BeLike 'concurrency lowered from 2 to 1*'
        $l.OnThrottle() | Should -BeNullOrEmpty -Because 'it never goes under 1'
        Start-Sleep -Milliseconds 200
        $slot = $l.Enter([Threading.CancellationToken]::None); $slot.Dispose()
        $l.Limit | Should -Be 2
        Start-Sleep -Milliseconds 200
        $slot = $l.Enter([Threading.CancellationToken]::None); $slot.Dispose()
        $l.Limit | Should -Be 3
        foreach ($i in 1..10) { Start-Sleep -Milliseconds 160; $slot = $l.Enter([Threading.CancellationToken]::None); $slot.Dispose() }
        $l.Limit | Should -Be 4 -Because 'the maximum is the documented 4'
    }
    It 'lowers the concurrency of an export that is throttled, and the service never sees more than 4 requests at once' {
        $svc = New-SourceService -InboxItems 40 -ProjectItems 0
        $svc.LatencyMs = 40
        $store = New-TestStoreAt 'adaptive'
        try {
            $svc.Throttle(2, 0, 'exportItems')
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 4
            $r.Result.Exported | Should -Be 40
            $svc.MaxConcurrent | Should -BeLessOrEqual 4
            ($r.Messages | Where-Object { $_ -like 'concurrency lowered from 4 to 2*' }).Count | Should -Be 1
            $r.Job.Snapshot().Limit | Should -BeLessThan 4
        } finally { $store.Dispose() }
        $svc2 = New-SourceService -InboxItems 20 -ProjectItems 0
        $svc2.LatencyMs = 30
        $store2 = New-TestStoreAt 'one'
        try {
            $null = Invoke-ExportJob -Service $svc2 -Store $store2 -Concurrency 1
            $svc2.MaxConcurrent | Should -Be 1
        } finally { $store2.Dispose() }
    }
}

Describe 'Windows of the sustained rate' {
    It 'closes a window every WindowSeconds with its rate, its items and its 429, then a last partial one' {
        $svc = New-SourceService -InboxItems 40 -ProjectItems 0 -Padding 2000
        $svc.LatencyMs = 100
        $store = New-TestStoreAt 'windows'
        $saved = [MailboxTransfer.TransferJob]::DefaultWindowSeconds
        try {
            [MailboxTransfer.TransferJob]::DefaultWindowSeconds = 1
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 1 -Options { param($o) $o.BatchItems = 2 }
            $w = @($r.Job.GetWindows())
            $w.Count | Should -BeGreaterOrEqual 2
            ($w | Measure-Object Items -Sum).Sum | Should -Be 40
            ($w | Measure-Object Bytes -Sum).Sum | Should -Be $r.Result.Bytes
            $w[0].Index | Should -Be 1
            $w[0].MiBPerSecond | Should -BeGreaterThan 0
            $w[0].Limit | Should -Be 1
            @($r.Messages | Where-Object { $_ -match '^window 1 \(00:00:00-00:00:0\d\): .* MiB/s, .* GiB/h, \d+ items, 429: 0, concurrency 1$' }).Count | Should -Be 1
        } finally { [MailboxTransfer.TransferJob]::DefaultWindowSeconds = $saved; $store.Dispose() }
    }
    It 'a last window shorter than half a second still holds what it saw: the windows always add up to the whole transfer' {
        # The last window was only closed when it was longer than 0.5 s: a short one lost its items (40 expected, 38 seen when the
        # job ended just after a boundary). With a window longer than the job there is only that last window, whatever the timing.
        $svc = New-SourceService -InboxItems 40 -ProjectItems 0 -Padding 2000
        $svc.LatencyMs = 100                  # two requests of 100 ms: the transfer lasts about 0.2 s, less than the 0.5 s the last window needed
        $store = New-TestStoreAt 'windows-short'
        $saved = [MailboxTransfer.TransferJob]::DefaultWindowSeconds
        try {
            [MailboxTransfer.TransferJob]::DefaultWindowSeconds = 3600
            $r = Invoke-ExportJob -Service $svc -Store $store -Concurrency 1 -Options { param($o) $o.BatchItems = 20 }
            $w = @($r.Job.GetWindows())
            $w.Count | Should -Be 1
            $w[0].Items | Should -Be 40
            $w[0].Bytes | Should -Be $r.Result.Bytes
        } finally { [MailboxTransfer.TransferJob]::DefaultWindowSeconds = $saved; $store.Dispose() }
    }
}

Describe 'Import engine: folders' {
    It 'maps a source path to a target path: same path, below a root, explicit mapping (longest prefix wins)' {
        $o = [MailboxTransfer.ImportOptions]::new()
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Inbox/Projects') | Should -Be 'Inbox/Projects'
        $o.TargetRoot = 'MBT-Phase1-A'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Inbox/Projects') | Should -Be 'MBT-Phase1-A/Inbox/Projects'
        $map = [System.Collections.Generic.Dictionary[string, string]]::new()
        $map['Inbox'] = 'Mail'; $map['Inbox/Projects'] = 'Work/Projets'
        $o.Mapping = $map
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Inbox') | Should -Be 'MBT-Phase1-A/Mail'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Inbox/Other') | Should -Be 'MBT-Phase1-A/Mail/Other'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Inbox/Projects/2026') | Should -Be 'MBT-Phase1-A/Work/Projets/2026'
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Inboxes') | Should -Be 'MBT-Phase1-A/Inboxes' -Because 'a prefix is a whole folder name'
        $o.TargetRoot = $null
        [MailboxTransfer.ImportPlanner]::MapPath($o, 'Inbox/Projects') | Should -Be 'Work/Projets'
    }
    It 'plans what would be created without sending anything (parents before children)' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'plan'
        try {
            $o = [MailboxTransfer.ImportOptions]::new()
            $o.SourceMailboxId = $script:Src; $o.TargetMailboxId = $script:Dst; $o.TargetRoot = 'MBT-Phase1-X/Deep'
            $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Service $svc))
            $posts = $svc.RequestsOf('POST')
            $plan = [MailboxTransfer.ImportPlanner]::Build($store, $o, $api.GetFolderTree($script:Dst))
            @($plan.PathsToCreate) | Should -Be @('MBT-Phase1-X', 'MBT-Phase1-X/Deep', 'MBT-Phase1-X/Deep/Empty', 'MBT-Phase1-X/Deep/Inbox', 'MBT-Phase1-X/Deep/Inbox/Projects')
            $plan.Items | Should -Be 10
            $plan.FoldersExisting | Should -Be 0
            $svc.RequestsOf('POST') | Should -Be $posts
            $o.TargetRoot = $null
            $same = [MailboxTransfer.ImportPlanner]::Build($store, $o, $api.GetFolderTree($script:Dst))
            @($same.PathsToCreate) | Should -Be @('Empty', 'Inbox/Projects')
            $same.FoldersExisting | Should -Be 1
        } finally { $store.Dispose() }
    }
    It 'creates the tree below the root, imports every item once with one session, and a second run creates nothing' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'imp1'
        try {
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 4 -Options { param($o) $o.TargetRoot = 'MBT-Phase1-A' }
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Completed | Should -BeTrue
            $r.Result.FoldersCreated | Should -Be 4
            $r.Result.Done | Should -Be 10
            $r.Result.Failed | Should -Be 0
            $r.Result.Unknown | Should -Be 0
            @(Get-FolderPaths $svc $script:Dst) | Should -Be @('Inbox', 'MBT-Phase1-A', 'MBT-Phase1-A/Empty', 'MBT-Phase1-A/Inbox', 'MBT-Phase1-A/Inbox/Projects')
            $svc.MaxConcurrent | Should -BeLessOrEqual 4
            $r.Result.SessionsCreated | Should -Be 1 -Because 'one session serves every item until it is about to expire'
            $r.Result.SessionLifetimeSeconds[0] | Should -BeGreaterThan 3500
            $imported = $svc.ItemsOf($script:Dst)
            $imported.Count | Should -Be 10
            @($imported | ForEach-Object { Get-Sha $_.Data } | Sort-Object) | Should -Be @($svc.ItemsOf($script:Src) | ForEach-Object { Get-Sha $_.Data } | Sort-Object)
            $inboxId = ($svc.FoldersOf($script:Dst) | Where-Object { $_.Name -eq 'Inbox' -and $_.Id -ne 'D-INBOX' }).Id
            @($svc.ItemsOf($script:Dst, $inboxId)).Count | Should -Be 6
            $svc.ItemsOf($script:Dst, 'D-INBOX').Count | Should -Be 0 -Because 'the existing Inbox of the target is not touched'
            ($store.GetImportStats($r.Result.JobId) | Where-Object State -EQ 'Done').Count | Should -Be 10
            $again = Invoke-ImportJob -Service $svc -Store $store -Options { param($o) $o.TargetRoot = 'MBT-Phase1-A' }
            $again.Result.FoldersCreated | Should -Be 0
            $again.Result.Done | Should -Be 0
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
            @(Get-FolderPaths $svc $script:Dst).Count | Should -Be 5
            $svc.RequestsOf('importItem') | Should -Be 10
        } finally { $store.Dispose() }
    }
    It 'at the same path: reuses the folders that exist and creates only the missing ones' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'imp2'
        try {
            $r = Invoke-ImportJob -Service $svc -Store $store
            $r.Result.FoldersCreated | Should -Be 2
            @(Get-FolderPaths $svc $script:Dst) | Should -Be @('Empty', 'Inbox', 'Inbox/Projects')
            $svc.ItemsOf($script:Dst, 'D-INBOX').Count | Should -Be 6
        } finally { $store.Dispose() }
    }
    It 'follows an explicit mapping' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'imp3'
        try {
            $r = Invoke-ImportJob -Service $svc -Store $store -Options { param($o)
                $map = [System.Collections.Generic.Dictionary[string, string]]::new(); $map['Inbox/Projects'] = 'Archive/Work'
                $o.Mapping = $map; $o.TargetRoot = 'R' }
            $r.Result.Done | Should -Be 10
            @(Get-FolderPaths $svc $script:Dst) | Should -Be @('Inbox', 'R', 'R/Archive', 'R/Archive/Work', 'R/Empty', 'R/Inbox')
            $work = ($svc.FoldersOf($script:Dst) | Where-Object Name -EQ 'Work').Id
            $svc.ItemsOf($script:Dst, $work).Count | Should -Be 4
        } finally { $store.Dispose() }
    }
}

Describe 'Import engine: item states, sessions, throttling' {
    It 'marks Failed what the service refused (nothing created), Unknown what may exist, and replays only the Failed ones' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'st1'
        try {
            $svc.ImportFailures.Enqueue(400); $svc.ImportFailures.Enqueue(504)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 8
            $r.Result.Failed | Should -Be 1
            $r.Result.Unknown | Should -Be 1
            $svc.ItemsOf($script:Dst).Count | Should -Be 8
            $failed = @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::Failed, 10))
            $unknown = @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::Unknown, 10))
            $failed.Count | Should -Be 1
            $failed[0].LastError | Should -BeLike '*400*'
            $unknown.Count | Should -Be 1
            $unknown[0].LastError | Should -BeLike '*504*'
            ($r.Messages | Where-Object { $_ -like '*UNKNOWN*may have been created*' }).Count | Should -Be 1
            $second = Invoke-ImportJob -Service $svc -Store $store
            $second.Result.Done | Should -Be 1 -Because 'the Failed one is tried again'
            $svc.ItemsOf($script:Dst).Count | Should -Be 9
            @($store.GetImportItems($second.Result.JobId, [MailboxTransfer.StoreImportState]::Unknown, 10)).Count | Should -Be 1 -Because 'an Unknown item is never replayed by itself'
            $store.ResolveImportUnknown($unknown[0].ImportKey, $false, $null)
            $third = Invoke-ImportJob -Service $svc -Store $store
            $third.Result.Done | Should -Be 1
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
        } finally { $store.Dispose() }
    }
    It 'a 2xx answer of importItem that cannot be read is Unknown (the item exists), never Failed and never sent again; an empty answer is Done' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'ans1'
        try {
            $svc.ImportAnswers.Enqueue('garbage')
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 9
            $r.Result.Unknown | Should -Be 1
            $r.Result.Failed | Should -Be 0
            $r.Result.JobComplete | Should -BeFalse
            @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::Unknown, 10))[0].LastError | Should -BeLike '*answer*'
            $svc.RequestsOf('/importItem?') | Should -Be 10
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
            $second = Invoke-ImportJob -Service $svc -Store $store
            $second.Result.Done | Should -Be 0
            $svc.RequestsOf('/importItem?') | Should -Be 10 -Because 'an Unknown item is never sent again by itself'
        } finally { $store.Dispose() }
        $svc2 = New-SourceService
        Add-Targets $svc2
        $store2 = New-ExportedStore $svc2 'ans2'
        try {
            $svc2.ImportAnswers.Enqueue('')
            $e = Invoke-ImportJob -Service $svc2 -Store $store2 -Concurrency 1
            $e.Result.Done | Should -Be 10
            $e.Result.JobComplete | Should -BeTrue
        } finally { $store2.Dispose() }
    }
    It 'leaves the job incomplete when items were left out: a target folder that cannot be created, items of the store not exported yet' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'left1'
        try {
            $svc.RefuseFolderNames.Add('Projects'); $svc.RefuseFolderStatus = 400
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 6
            $r.Result.NotQueued | Should -Be 4
            $r.Result.JobComplete | Should -BeFalse
            $r.Result.JobSummary() | Should -BeLike '*4 left out of the job*'
            $svc.RefuseFolderNames.Clear()
            $again = Invoke-ImportJob -Service $svc -Store $store
            $again.Result.Done | Should -Be 4
            $again.Result.NotQueued | Should -Be 0
            $again.Result.JobComplete | Should -BeTrue
        } finally { $store.Dispose() }
        $svc2 = New-SourceService
        Add-Targets $svc2
        $null = $svc2.BadItemIds.Add('I2')
        $store2 = New-TestStoreAt 'left2'
        try {
            $null = Invoke-ExportJob -Service $svc2 -Store $store2 -Options { param($o) $o.MaxAttempts = 1 }
            $r2 = Invoke-ImportJob -Service $svc2 -Store $store2
            $r2.Result.Done | Should -Be 9
            $r2.Result.NotQueued | Should -Be 1
            $r2.Result.JobComplete | Should -BeFalse
            ($r2.Result.Problems -join ' ') | Should -BeLike '*not exported yet*'
        } finally { $store2.Dispose() }
    }
    It 'sends again in place an item whose upload was interrupted before its last byte (Done, created once)' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'cut1'
        try {
            $svc.UploadDrops.Enqueue(60)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 10
            $r.Result.Failed | Should -Be 0
            $r.Result.Unknown | Should -Be 0
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
            $r.Client.RetryCount | Should -Be 1
            ($r.Messages | Where-Object { $_ -like 'upload interrupted after * of * bytes, nothing created, retrying*' }).Count | Should -Be 1
        } finally { $store.Dispose() }
    }
    It 'marks Failed (nothing created), not Unknown, an item whose upload is interrupted on every attempt' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'cut2'
        try {
            1..6 | ForEach-Object { $svc.UploadDrops.Enqueue(60) }
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 9
            $r.Result.Failed | Should -Be 1
            $r.Result.Unknown | Should -Be 0
            $failed = @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::Failed, 10))
            $failed.Count | Should -Be 1
            $failed[0].LastError | Should -BeLike '*upload interrupted*nothing created*'
            ($r.Messages | Where-Object { $_ -like '*failed (nothing created)*upload interrupted*' }).Count | Should -Be 1
            $second = Invoke-ImportJob -Service $svc -Store $store
            $second.Result.Done | Should -Be 1 -Because 'the Failed one is tried again by the next run'
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
        } finally { $store.Dispose() }
    }
    It 'sends again in place an item whose request never left (connection refused before any byte), Done and created once' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'ns1'
        try {
            $svc.UploadDrops.Enqueue(-2); $svc.UploadDrops.Enqueue(-2)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 10
            $r.Result.Failed | Should -Be 0
            $r.Result.Unknown | Should -Be 0
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
            $r.Client.RetryCount | Should -Be 2
            ($r.Messages | Where-Object { $_ -like 'request not sent (no byte of the body reached the transport), nothing created, retrying*' }).Count | Should -Be 2
        } finally { $store.Dispose() }
    }
    It 'marks Failed (nothing created), not Unknown, an item whose request is not sent on every attempt, and tries it again at the next run' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'ns2'
        try {
            1..6 | ForEach-Object { $svc.UploadDrops.Enqueue(-2) }
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 9
            $r.Result.Failed | Should -Be 1
            $r.Result.Unknown | Should -Be 0
            $failed = @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::Failed, 10))
            $failed.Count | Should -Be 1
            $failed[0].LastError | Should -BeLike '*request not sent*nothing created*'
            $r.Result.JobComplete | Should -BeFalse
            $r.Result.JobFailed | Should -Be 1
            $second = Invoke-ImportJob -Service $svc -Store $store
            $second.Result.Done | Should -Be 1
            $second.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
        } finally { $store.Dispose() }
    }
    It 'keeps Unknown an item after a 504 (an HTTP status was received after the body was sent)' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'ns3'
        try {
            $svc.ImportFailures.Enqueue(504)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 9
            $r.Result.Failed | Should -Be 0
            $r.Result.Unknown | Should -Be 1
            $r.Client.RetryCount | Should -Be 0
        } finally { $store.Dispose() }
    }
    It 'reports the state of the whole job: an Unknown row left by an earlier run makes the job incomplete although this run did nothing wrong' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'job1'
        try {
            $svc.UploadDrops.Enqueue(-1)
            $first = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'big'
            $first.Result.Unknown | Should -Be 1
            $first.Result.JobTotal | Should -Be 10
            $first.Result.JobDone | Should -Be 9
            $first.Result.JobUnknown | Should -Be 1
            $first.Result.JobUnknownEarlier | Should -Be 0
            $first.Result.JobComplete | Should -BeFalse

            $again = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'big'
            $again.Result.Done | Should -Be 0
            $again.Result.Failed | Should -Be 0
            $again.Result.Unknown | Should -Be 0 -Because 'this run did nothing'
            $again.Result.JobKnown | Should -BeTrue
            $again.Result.JobTotal | Should -Be 10
            $again.Result.JobDone | Should -Be 9
            $again.Result.JobUnknown | Should -Be 1
            $again.Result.JobUnknownEarlier | Should -Be 1
            $again.Result.JobPending | Should -Be 0
            $again.Result.JobComplete | Should -BeFalse
            $again.Result.JobSummary() | Should -BeLike "job 'big' incomplete: 9 of 10 done, 1 unknown (1 from an earlier run, *"
            ($again.Messages | Where-Object { $_ -like "job 'big' resumed: 9 done, 0 failed, 1 unknown*" }).Count | Should -Be 1

            $key = @($store.GetImportItems($again.Result.JobId, [MailboxTransfer.StoreImportState]::Unknown, 10))[0].ImportKey
            $store.ResolveImportUnknown($key, $false, $null)
            $last = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'big'
            $last.Result.Done | Should -Be 1
            $last.Result.JobComplete | Should -BeTrue
            $last.Result.JobSummary() | Should -Be "job 'big' complete: 10 of 10 done"
        } finally { $store.Dispose() }
    }
    It 'says at the start of a resumed job how many rows recovery turned from InFlight into Unknown' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'job2'
        $path = $store.Root
        $job = $store.CreateImportJob('crashed', $script:Dst)
        $keys = @($store.GetFolderTree($script:Src) | ForEach-Object { $_ } | ForEach-Object { $store.GetItems($_.FolderKey, 0, 50, 'ReceivedDate', $false, 'Exported').Items } | ForEach-Object Key)
        $null = $store.AddImportItems($job, @($keys | ForEach-Object { [MailboxTransfer.StoreImportPlan]@{ ItemKey = $_; TargetMailbox = $script:Dst; TargetFolder = 'D-INBOX' } }))
        $null = $store.NextImportItems($job, 3)
        $store.Abandon()
        $again = [MailboxTransfer.ExportStore]::Open($path)
        try {
            $again.LastRecovery.ImportInFlightToUnknownByJob[$job] | Should -Be 3
            $r = Invoke-ImportJob -Service $svc -Store $again -Concurrency 1 -Name 'crashed'
            $r.Result.Done | Should -Be 7
            $r.Result.Unknown | Should -Be 0
            $r.Result.RecoveredToUnknown | Should -Be 3
            $r.Result.JobUnknown | Should -Be 3
            $r.Result.JobUnknownEarlier | Should -Be 3
            $r.Result.JobComplete | Should -BeFalse
            ($r.Messages | Where-Object { $_ -like "recovery: 3 rows of the job 'crashed' were InFlight when an earlier run stopped and are now Unknown*" }).Count | Should -Be 1
            $r.Messages.IndexOf(($r.Messages | Where-Object { $_ -like 'recovery:*' } | Select-Object -First 1)) | Should -BeLessThan $r.Messages.IndexOf(($r.Messages | Where-Object { $_ -like 'plan:*' } | Select-Object -First 1))
        } finally { $again.Dispose() }
    }
    It 'finds the item with -Check when the service created it and the answer was lost, and finds nothing when it was not created' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'chk1'
        try {
            $null = Set-MbtSession -Graph (New-FakeClient -Service $svc)
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'lost'
            $r.Result.Unknown | Should -Be 1
            $svc.ItemsOf($script:Dst).Count | Should -Be 10 -Because 'the service created the item although the answer was lost'
            $done = @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::Done, 20) | ForEach-Object TargetItemId)
            $created = @($svc.ItemsOf($script:Dst) | Where-Object { $_.Id -notin $done })
            $created.Count | Should -Be 1
            $rows = @(Get-MbtImportUnknown -Store $store -Check)
            $rows.Count | Should -Be 1
            @($rows[0].Candidates) | Should -Be @($created[0].Id)
        } finally { $store.Dispose() }
        $svc2 = New-SourceService
        Add-Targets $svc2
        $store2 = New-ExportedStore $svc2 'chk2'
        try {
            $null = Set-MbtSession -Graph (New-FakeClient -Service $svc2)
            $svc2.UploadDrops.Enqueue(-1)
            $r2 = Invoke-ImportJob -Service $svc2 -Store $store2 -Concurrency 1 -Name 'lost'
            $r2.Result.Unknown | Should -Be 1
            $svc2.ItemsOf($script:Dst).Count | Should -Be 9
            @((Get-MbtImportUnknown -Store $store2 -Check)[0].Candidates).Count | Should -Be 0
        } finally { $store2.Dispose() }
    }
    It 'settles at once the Unknown rows with no matching item (-Absent), leaves those with candidates, and honours -WhatIf and -JobName' {
        $svc = New-SourceService
        Add-Targets $svc
        # distinct sizes: the check matches on size, so equal-sized items in the same folder would all be candidates of each other
        foreach ($i in 1..6) { $svc.ChangeItem($script:Src, "I$i", (New-Message "Inbox $i" ($i * 100)), "ck-I$i") }
        foreach ($i in 1..4) { $svc.ChangeItem($script:Src, "P$i", (New-Message "Project $i" ($i * 100 + 50)), "ck-P$i") }
        $store = New-ExportedStore $svc 'abs1'
        try {
            $null = Set-MbtSession -Graph (New-FakeClient -Service $svc)
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1)                                  # row 1: created, answer lost
            $svc.ImportFailures.Enqueue(504); $svc.ImportFailures.Enqueue(504)   # rows 2 and 3: not created
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'mixed'
            $r.Result.Done | Should -Be 7
            $r.Result.Unknown | Should -Be 3
            $count = { param($state) @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::$state, 20)).Count }

            $dry = @(Resolve-MbtImportUnknown -Store $store -Absent -WhatIf)
            $dry.Count | Should -Be 3
            @($dry | Where-Object Settled).Count | Should -Be 0
            & $count 'Unknown' | Should -Be 3 -Because '-WhatIf settles nothing'
            @(Resolve-MbtImportUnknown -Store $store -Absent -JobName 'other' -Confirm:$false).Count | Should -Be 0

            $out = @(Resolve-MbtImportUnknown -Store $store -Absent -Confirm:$false)
            $out.Count | Should -Be 3
            @($out | Where-Object Settled).Count | Should -Be 2
            @($out | Where-Object { -not $_.Settled })[0].Candidates.Count | Should -Be 1
            & $count 'Unknown' | Should -Be 1 -Because 'the row whose item exists is left alone'
            & $count 'Pending' | Should -Be 2

            $next = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'mixed'
            $next.Result.Done | Should -Be 2
            $next.Result.JobUnknown | Should -Be 1
            $next.Result.JobComplete | Should -BeFalse
            $svc.ItemsOf($script:Dst).Count | Should -Be 10 -Because 'the two absent items were created once, the existing one was not replayed'
            $left = @(Get-MbtImportUnknown -Store $store -Check)[0]
            Resolve-MbtImportUnknown -Store $store -ImportKey $left.ImportKey -Imported -TargetItemId $left.Candidates[0] -Confirm:$false
            $final = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'mixed'
            $final.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
        } finally { $store.Dispose() }
    }
    It 'keeps Unknown an item whose body was fully sent and whose answer was lost' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'cut3'
        try {
            $svc.UploadDrops.Enqueue(-1)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 9
            $r.Result.Failed | Should -Be 0
            $r.Result.Unknown | Should -Be 1
            $r.Client.RetryCount | Should -Be 0 -Because 'a request that may have created the item is never sent again'
            @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::Unknown, 10)).Count | Should -Be 1
        } finally { $store.Dispose() }
    }
    It 'turns the items InFlight at a crash into Unknown when the store is opened again' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'st2'
        $path = $store.Root
        $job = $store.CreateImportJob('crash', $script:Dst)
        $keys = @($store.GetFolderTree($script:Src) | ForEach-Object { $_ } | ForEach-Object { $store.GetItems($_.FolderKey, 0, 50, 'ReceivedDate', $false, 'Exported').Items } | ForEach-Object Key)
        $null = $store.AddImportItems($job, @($keys | ForEach-Object { [MailboxTransfer.StoreImportPlan]@{ ItemKey = $_; TargetMailbox = $script:Dst; TargetFolder = 'D-INBOX' } }))
        $null = $store.NextImportItems($job, 3)
        $store.Abandon()
        $again = [MailboxTransfer.ExportStore]::Open($path)
        try {
            $again.LastRecovery.ImportInFlightToUnknown | Should -Be 3
            @($again.GetImportItems($job, [MailboxTransfer.StoreImportState]::Unknown, 10)).Count | Should -Be 3
        } finally { $again.Dispose() }
    }
    It 'creates a session for the auxiliary archive named by a 409, imports there, and reuses the sessions' {
        $svc = New-SourceService -ProjectItems 0
        Add-Targets $svc
        $svc.MovedFolderToMailbox['D-INBOX'] = $script:Aux
        $store = New-ExportedStore $svc 'aux1'
        try {
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 6
            $r.Result.Failed | Should -Be 0
            $svc.ItemsOf($script:Aux, 'D-INBOX').Count | Should -Be 6
            $svc.ItemsOf($script:Dst, 'D-INBOX').Count | Should -Be 0
            ($r.Messages | Where-Object { $_ -like "the folder lives in the auxiliary archive $($script:Aux)*" }).Count | Should -BeGreaterThan 0
            $r.Result.SessionsCreated | Should -Be 2 -Because 'one session for the main archive and one for the auxiliary one, then reused'
        } finally { $store.Dispose() }
    }
    It 'says clearly that the auxiliary archive is refused too, naming the mailbox, and creates nothing' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $store = New-ExportedStore $svc 'aux2'
        try {
            # the service keeps answering 409 for the auxiliary archive: a session for it did not help
            $svc.Force409Mailbox = $script:Aux
            $svc.ImportFailures.Enqueue(409); $svc.ImportFailures.Enqueue(409)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 2
            $r.Result.Failed | Should -Be 1
            $failed = @($store.GetImportItems($r.Result.JobId, [MailboxTransfer.StoreImportState]::Failed, 10))
            $failed[0].LastError | Should -BeLike "*auxiliary archive*$($script:Aux)*Nothing was created*"
            $svc.ItemsOf($script:Dst).Count | Should -Be 2
        } finally { $store.Dispose() }
    }
    It 'renews a session that the service refuses (401) and one that is about to expire' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 4
        Add-Targets $svc
        $store = New-ExportedStore $svc 'sess'
        try {
            $svc.ImportFailures.Enqueue(401)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 4
            $r.Result.SessionsCreated | Should -Be 2
            ($r.Messages | Where-Object { $_ -like '*refused (401)*' }).Count | Should -Be 1
        } finally { $store.Dispose() }
        $svc2 = New-SourceService -ProjectItems 0 -InboxItems 4
        Add-Targets $svc2
        $store2 = New-ExportedStore $svc2 'sess2'
        try {
            $svc2.ImportExpirationOffsetMinutes = '1'
            $r2 = Invoke-ImportJob -Service $svc2 -Store $store2 -Concurrency 1
            $r2.Result.Done | Should -Be 4
            $r2.Result.SessionsCreated | Should -Be 4 -Because 'a session that expires within two minutes is not reused'
        } finally { $store2.Dispose() }
    }
    It 'stops on 403 and honours Retry-After on the import like on the export' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 30
        Add-Targets $svc
        $svc.LatencyMs = 20
        $store = New-ExportedStore $svc 'thr'
        try {
            $svc.Throttle(1, 2, 'importItem')
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 4 -DelayScale 1
            $sw.Stop()
            $r.Result.Done | Should -Be 30
            $r.Client.Throttled429 | Should -Be 1
            $sw.Elapsed.TotalSeconds | Should -BeGreaterOrEqual 1.9
            $at = $svc.ThrottledAt[0]
            @($svc.Times | Where-Object { $_ -gt $at.AddMilliseconds(150) -and $_ -lt $at.AddSeconds(1.9) }).Count | Should -Be 0
            ($r.Messages | Where-Object { $_ -like 'concurrency lowered from 4 to 2*' }).Count | Should -Be 1
            $r.Result.Failed | Should -Be 0 -Because 'a 429 is waited for, the item is not lost'
        } finally { $store.Dispose() }
    }
}

Describe 'Invoke-MailboxTransfer.ps1 (against the fake service)' {
    BeforeAll {
        $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
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
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtEntryClient, MbtEntryExit -Scope Global -ErrorAction SilentlyContinue }
    It 'Status without a store says so, and shows that the engine compiles' {
        $text = Invoke-Entry (New-SourceService) @{ Mode = 'Status'; StorePath = (Join-Path $script:Tmp 'entry-none') }
        $text | Should -Match 'the C# engine compiles'
        $text | Should -Match 'no store in'
    }
    It 'Export of the primary mailbox: tree, items, previews, progress and a report without throttling' {
        $svc = New-SourceService
        $store = Join-Path $script:Tmp 'entry-store'
        $text = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; TempPath = (Join-Path $script:Tmp 'entry-tmp'); PrimaryMailboxId = $script:Src; Concurrency = 2 }
        $text | Should -Match 'signed in as adele.vance@contoso.com'
        $text | Should -Match 'export finished: 10 items exported'
        $text | Should -Match 'no throttling answer'
        $text | Should -Not -Match '(?m)^ {6}x {2}\S'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try {
            $s.GetMailboxes()[0].Kind | Should -Be 'Primary'
            ($s.GetItemStats($script:Src) | Where-Object State -EQ 'Exported').Count | Should -Be 10
        } finally { $s.Dispose() }
        $again = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; TempPath = (Join-Path $script:Tmp 'entry-tmp'); PrimaryMailboxId = $script:Src; SkipInventory = $true }
        $again | Should -Match 'export finished: 0 items exported'
    }
    It 'Export of the archive records the mailbox as Archive' {
        $svc = New-SourceService
        $store = Join-Path $script:Tmp 'entry-arch'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'Archive'; ArchiveMailboxId = $script:Src; StorePath = $store; TempPath = (Join-Path $script:Tmp 'entry-tmp2'); Folder = 'Inbox/Projects' }
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try {
            $s.GetMailboxes()[0].Kind | Should -Be 'Archive'
            ($s.GetItemStats($script:Src) | Where-Object State -EQ 'Exported').Count | Should -Be 4
        } finally { $s.Dispose() }
    }
    It 'Import -WhatIf shows the plan and sends nothing; Import creates the tree under the root and imports; Status shows it' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = Join-Path $script:Tmp 'entry-imp'
        $common = @{ StorePath = $store; PrimaryMailboxId = $script:Src }
        $null = Invoke-Entry $svc ($common + @{ Mode = 'Export'; TempPath = (Join-Path $script:Tmp 'entry-tmp3') })
        $posts = $svc.RequestsOf('POST')
        $plan = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; TargetRoot = 'MBT-Phase1-E'; WhatIf = $true }
        $plan | Should -Match 'plan: 10 items'
        $plan | Should -Match 'would create folder: MBT-Phase1-E/Inbox/Projects'
        $plan | Should -Match 'WhatIf: nothing was created'
        $svc.RequestsOf('POST') | Should -Be $posts
        $text = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; TargetRoot = 'MBT-Phase1-E'; Concurrency = 2 }
        $text | Should -Match 'import finished: 10 items imported'
        $text | Should -Match '1 import session\(s\) created'
        $text | Should -Not -Match '(?m)^ {6}x {2}\S'
        $svc.ItemsOf($script:Dst).Count | Should -Be 10
        $status = Invoke-Entry $svc @{ Mode = 'Status'; StorePath = $store }
        $status | Should -Match 'Primary mailbox MBX:11111111'
        $status | Should -Match 'Exported 10'
        $status | Should -Match "job 1 'import'"
        $status | Should -Match 'Done 10'
        $status | Should -Match 'Projects'
    }
    It 'Import follows a mapping file, and refuses an ambiguous source or a missing store' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = Join-Path $script:Tmp 'entry-map'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:Src; TempPath = (Join-Path $script:Tmp 'entry-tmp4') }
        $null = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'Archive'; ArchiveMailboxId = $script:Dst; StorePath = $store; TempPath = (Join-Path $script:Tmp 'entry-tmp4') }
        $map = Join-Path $script:Tmp 'map.psd1'
        Set-Content -LiteralPath $map -Value "@{ 'Inbox/Projects' = 'Archive/Work' }"
        $text = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; SourceMailbox = 'Primary'; PrimaryMailboxId = $script:Dst; TargetRoot = 'R'; MappingFile = $map; WhatIf = $true }
        $text | Should -Match 'would create folder: R/Archive/Work'
        { Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; WhatIf = $true } } | Should -Throw '*choose with -SourceMailbox*'
        { Invoke-Entry $svc @{ Mode = 'Import'; StorePath = (Join-Path $script:Tmp 'entry-nothing'); PrimaryMailboxId = $script:Dst } } | Should -Throw '*No store in*'
    }
    It 'lists and settles the Unknown import rows' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $store = Join-Path $script:Tmp 'entry-unk'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:Src; TempPath = (Join-Path $script:Tmp 'entry-tmp5') }
        $svc.ImportFailures.Enqueue(504)
        $null = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; Concurrency = 1; NoAutoReconcile = $true }
        $list = Invoke-Entry $svc @{ Mode = 'Status'; StorePath = $store; Unknown = $true; Check = $true; PrimaryMailboxId = $script:Dst }
        $list | Should -Match 'import 1 \(job ''import''\)'
        $list | Should -Match '504'
        $list | Should -Match 'no matching item found in the target folder'
        $null = Invoke-Entry $svc @{ Mode = 'Status'; StorePath = $store; ResolveUnknown = 1; As = 'NotImported' }
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try { @(Get-MbtImportUnknown -Store $s).Count | Should -Be 0 } finally { $s.Dispose() }
        $again = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; Concurrency = 1 }
        $again | Should -Match 'import finished: 1 items imported'
        $svc.ItemsOf($script:Dst).Count | Should -Be 3
    }
    It 'reports the whole job at the end of an import, exits non-zero while it is incomplete, and settles the absent rows in bulk' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $store = Join-Path $script:Tmp 'entry-job'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:Src; TempPath = (Join-Path $script:Tmp 'entry-tmp6') }
        $svc.ImportFailures.Enqueue(504)
        $first = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; Concurrency = 1; NoAutoReconcile = $true }
        $global:MbtEntryExit | Should -Be 3
        $first | Should -Match 'import finished: 2 items imported .*1 unknown \(may exist\).*; job ''import'' INCOMPLETE: 2 of 3 done, 0 failed, 1 unknown, 0 pending'
        $first | Should -Match 'job ''import'' incomplete: 2 of 3 done, 1 unknown'
        $first | Should -Match 'exit code 3'

        $resumed = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; Concurrency = 1; NoAutoReconcile = $true }
        $global:MbtEntryExit | Should -Be 3
        $resumed | Should -Match 'import finished: 0 items imported .*0 unknown \(may exist\).*; job ''import'' INCOMPLETE: 2 of 3 done, 0 failed, 1 unknown'
        $resumed | Should -Match '1 from an earlier run'
        $resumed | Should -Match 'job ''import'' resumed: 2 done, 0 failed, 1 unknown'
        $resumed | Should -Match '-Mode Status -Unknown -Check'

        { Invoke-Entry $svc @{ Mode = 'Status'; StorePath = $store; Unknown = $true; ResolveAbsent = $true } } | Should -Throw '*-ResolveAbsent needs -Unknown -Check*'
        $dry = Invoke-Entry $svc @{ Mode = 'Status'; StorePath = $store; Unknown = $true; Check = $true; ResolveAbsent = $true; WhatIf = $true; PrimaryMailboxId = $script:Dst }
        $dry | Should -Match 'WhatIf: no row was settled'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try { @(Get-MbtImportUnknown -Store $s).Count | Should -Be 1 } finally { $s.Dispose() }

        $bulk = Invoke-Entry $svc @{ Mode = 'Status'; StorePath = $store; Unknown = $true; Check = $true; ResolveAbsent = $true; PrimaryMailboxId = $script:Dst }
        $bulk | Should -Match 'settled: no matching item found in the target folder, back to Pending'
        $bulk | Should -Match '1 Unknown row\(s\) checked: 1 settled as not imported, 0 left Unknown'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try { @(Get-MbtImportUnknown -Store $s).Count | Should -Be 0 } finally { $s.Dispose() }

        $done = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; Concurrency = 1 }
        $global:MbtEntryExit | Should -Be 0
        $done | Should -Match 'import finished: 1 items imported .*; job ''import'' COMPLETE: 3 of 3 done'
        $done | Should -Match 'job ''import'' complete: 3 of 3 done'
        $svc.ItemsOf($script:Dst).Count | Should -Be 3
    }
    It 'finishes by itself when the answer of an item is lost: reconciled, exit code 0, and says so; -WhatIf announces it' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $store = Join-Path $script:Tmp 'entry-auto'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:Src; TempPath = (Join-Path $script:Tmp 'entry-tmp8') }
        $plan = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; WhatIf = $true; ReconcileDelaySeconds = 120 }
        $plan | Should -Match 'automatic reconciliation is on: .* 120 s after the request ended \(at most 3 lookups per item\); 0 Unknown row\(s\)'
        $svc.CreateOnLostAnswer = $true
        $svc.UploadDrops.Enqueue(-1)
        $text = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; Concurrency = 1; ReconcileDelaySeconds = 0 }
        $global:MbtEntryExit | Should -Be 0
        $text | Should -Match 'import finished: 3 items imported .*0 unknown \(may exist\).*job ''import'' COMPLETE: 3 of 3 done'
        $text | Should -Match 'reconciliation: 1 Unknown row\(s\) found created in the target \(marked Done\), 0 found absent and imported again, 0 left Unknown'
        $text | Should -Match 'reconciled: the item was created, id NEW'
        $svc.ItemsOf($script:Dst).Count | Should -Be 3
    }
    It 'exits with 2 when the job only has Failed rows left (run it again)' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $store = Join-Path $script:Tmp 'entry-fail'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:Src; TempPath = (Join-Path $script:Tmp 'entry-tmp7') }
        1..6 | ForEach-Object { $svc.UploadDrops.Enqueue(-2) }
        $text = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; Concurrency = 1 }
        $global:MbtEntryExit | Should -Be 2
        $text | Should -Match '1 failed \(nothing created\), 0 unknown'
        $text | Should -Match 'job ''import'' incomplete: 2 of 3 done, 1 failed'
        $again = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; PrimaryMailboxId = $script:Dst; Concurrency = 1 }
        $global:MbtEntryExit | Should -Be 0
    }
    It 'reads a mapping file in .psd1 or .json and refuses a target that is not a path' {
        $psd1 = Join-Path $script:Tmp 'm1.psd1'; $json = Join-Path $script:Tmp 'm1.json'; $bad = Join-Path $script:Tmp 'm2.psd1'
        Set-Content -LiteralPath $psd1 -Value "@{ 'Inbox/Projects' = 'Archive/Work'; 'Sent Items' = 'Archive/Sent' }"
        Set-Content -LiteralPath $json -Value '{ "Inbox/Projects": "Archive/Work" }'
        Set-Content -LiteralPath $bad -Value "@{ 'Inbox' = 5 }"
        $m = Import-MbtMapping -Path $psd1
        $m['Inbox/Projects'] | Should -Be 'Archive/Work'
        $m.Count | Should -Be 2
        (Import-MbtMapping -Path $json)['Inbox/Projects'] | Should -Be 'Archive/Work'
        { Import-MbtMapping -Path $bad } | Should -Throw '*must be a path*'
        { Import-MbtMapping -Path (Join-Path $script:Tmp 'nothing.psd1') } | Should -Throw '*not found*'
    }
}

Describe 'Import engine: automatic reconciliation of the Unknown rows' {
    BeforeAll {
        # Indexes of the requests that listed the items of the target (the export of the source listed the source, a different mailbox).
        function Get-TargetListings($svc) { @(0..($svc.Log.Count - 1) | Where-Object { $svc.Log[$_] -like '*MBX:99999999*/items*' }) }
        function Get-Row($store, $result, [string]$State) { @($store.GetImportItems($result.Result.JobId, [MailboxTransfer.StoreImportState]::$State, 100)) }
    }
    It 'is on by default, with a settle delay of 5 minutes and 3 rounds' {
        $o = [MailboxTransfer.ImportOptions]::new()
        $o.AutoReconcile | Should -BeTrue
        $o.ReconcileDelaySeconds | Should -Be 300
        $o.ReconcileRounds | Should -Be 3
    }
    It 'marks Done an item created although the answer was lost, with its id, and creates no duplicate' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc1'
        try {
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Done | Should -Be 10
            $r.Result.Failed | Should -Be 0
            $r.Result.Unknown | Should -Be 0
            $r.Result.ReconciledCreated | Should -Be 1
            $r.Result.ReconciledReimported | Should -Be 0
            $r.Result.ReconcileAmbiguous | Should -Be 0
            $r.Result.JobComplete | Should -BeTrue
            $r.Result.ReconcileSummary() | Should -BeLike 'reconciliation: 1 Unknown row(s) found created*, 0 found absent and imported again, 0 left Unknown*'
            $svc.ItemsOf($script:Dst).Count | Should -Be 10 -Because 'the item that exists is not imported a second time'
            $ids = @(Get-Row $store $r 'Done' | ForEach-Object TargetItemId)
            $ids.Count | Should -Be 10
            @($ids | Select-Object -Unique).Count | Should -Be 10
            @($svc.ItemsOf($script:Dst).Id | Where-Object { $_ -notin $ids }).Count | Should -Be 0 -Because 'every item of the target is accounted for by exactly one row'
            ($r.Messages | Where-Object { $_ -like 'import * reconciled: the item was created, id NEW*' }).Count | Should -Be 1
        } finally { $store.Dispose() }
    }
    It 'imports again, in the same run, an item that was not created (lost answer or 504), as one more attempt' {
        foreach ($case in 'lost', '504') {
            $svc = New-SourceService
            Add-Targets $svc
            $store = New-ExportedStore $svc "rc2$case"
            try {
                if ($case -eq 'lost') { $svc.UploadDrops.Enqueue(-1) } else { $svc.ImportFailures.Enqueue(504) }
                $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile
                $r.Result.Done | Should -Be 10 -Because $case
                $r.Result.Unknown | Should -Be 0
                $r.Result.ReconciledCreated | Should -Be 0
                $r.Result.ReconciledReimported | Should -Be 1
                $r.Result.JobComplete | Should -BeTrue
                $svc.ItemsOf($script:Dst).Count | Should -Be 10
                @(Get-Row $store $r 'Done' | Where-Object Attempts -EQ 2).Count | Should -Be 1 -Because 'the second import is an attempt of the same row'
                ($r.Messages | Where-Object { $_ -like 'import * reconciled: no item was created*imported again (attempt 2 of 3)' }).Count | Should -Be 1
            } finally { $store.Dispose() }
        }
    }
    It 'does not take for the lost item an older item of the same size that was already in the folder' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 4
        Add-Targets $svc
        $null = $svc.AddItem($script:Dst, 'D-INBOX', 'OLD1', (New-Message 'Inbox 1'), 'ck-old')   # same size as the items to import, created long before
        $store = New-ExportedStore $svc 'rc3'
        try {
            $svc.UploadDrops.Enqueue(-1)   # not created
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile
            $r.Result.ReconciledCreated | Should -Be 0 -Because 'a match on the size alone would have taken OLD1 and lost an item'
            $r.Result.ReconciledReimported | Should -Be 1
            $r.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst, 'D-INBOX').Count | Should -Be 5
            @(Get-Row $store $r 'Done' | ForEach-Object TargetItemId) | Should -Not -Contain 'OLD1'
        } finally { $store.Dispose() }
    }
    It 'finds the right item among items of the same size: those the store already accounts for are not candidates' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 6
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc4'
        try {
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1)
            $t0 = [DateTime]::UtcNow
            # the delay (1 s) is longer than the queue: the lookup happens at the end, when 5 items of the same size are Done and the lost one exists
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile -Options { param($o) $o.ReconcileDelaySeconds = 1 }
            $r.Result.ReconciledCreated | Should -Be 1
            $r.Result.ReconcileAmbiguous | Should -Be 0
            $r.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst).Count | Should -Be 6
            @(Get-Row $store $r 'Done' | ForEach-Object TargetItemId | Select-Object -Unique).Count | Should -Be 6
            $listing = @(Get-TargetListings $svc)
            $listing.Count | Should -Be 2 -Because 'one listing of the folder (6 items, 3 per page) serves the lookup'
            ($svc.Times[$listing[0]] - $t0).TotalSeconds | Should -BeGreaterOrEqual 0.9 -Because 'the settle delay is respected'
        } finally { $store.Dispose() }
    }
    It 'leaves Unknown, and reports, items of the same size that cannot be told apart (two lost items, both created)' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 6
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc5'
        try {
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1); $svc.UploadDrops.Enqueue(-1)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile -Options { param($o) $o.ReconcileDelaySeconds = 1.5 }
            $r.Result.Done | Should -Be 4
            $r.Result.Unknown | Should -Be 2
            $r.Result.ReconcileAmbiguous | Should -Be 2
            $r.Result.ReconciledCreated | Should -Be 0
            $r.Result.JobUnknown | Should -Be 2
            $r.Result.JobComplete | Should -BeFalse
            $svc.ItemsOf($script:Dst).Count | Should -Be 6 -Because 'nothing was imported again'
            ($r.Messages | Where-Object { $_ -like 'import * left Unknown: 2 items in the target folder could be it*' }).Count | Should -Be 2
            ($r.Result.Problems | Where-Object { $_ -like '*left Unknown*' }).Count | Should -Be 2
        } finally { $store.Dispose() }
    }
    It 'reconciles at the start of a run the Unknown rows left by an earlier run, once their delay is over' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc6'
        try {
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1)
            $first = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc'      # without reconciliation: the row stays Unknown
            $first.Result.Unknown | Should -Be 1
            $first.Result.JobComplete | Should -BeFalse
            $failedAt = (Get-Row $store $first 'Unknown')[0].UpdatedUtc
            $failedAt | Should -Not -BeNullOrEmpty

            $second = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile -Options { param($o) $o.ReconcileDelaySeconds = 2 }
            $second.Result.ReconciledCreated | Should -Be 1
            $second.Result.Unknown | Should -Be 0
            $second.Result.JobUnknown | Should -Be 0
            $second.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
            ($second.Messages | Where-Object { $_ -like 'reconcile: 1 Unknown rows left by an earlier run*' }).Count | Should -Be 1
            $listing = @(Get-TargetListings $svc)
            ($svc.Times[$listing[0]] - $failedAt).TotalSeconds | Should -BeGreaterOrEqual 1.9 -Because 'the delay counts from the end of the request that left the row Unknown'
        } finally { $store.Dispose() }
    }
    It 'imports again at the start of a run an earlier Unknown row for which nothing was created' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc7'
        try {
            $svc.ImportFailures.Enqueue(504)
            $null = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc'
            $second = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile
            $second.Result.ReconciledReimported | Should -Be 1
            $second.Result.Done | Should -Be 1
            $second.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst).Count | Should -Be 10
        } finally { $store.Dispose() }
    }
    It 'keeps today''s behaviour with AutoReconcile off: the row stays Unknown and the folder is not listed' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc8'
        try {
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile -Options { param($o) $o.AutoReconcile = $false }
            $r.Result.Done | Should -Be 9
            $r.Result.Unknown | Should -Be 1
            $r.Result.JobUnknown | Should -Be 1
            $r.Result.JobComplete | Should -BeFalse
            $r.Result.ReconcileSummary() | Should -BeNullOrEmpty
            @(Get-TargetListings $svc).Count | Should -Be 0
        } finally { $store.Dispose() }
    }
    It 'stops after -ReconcileRounds lookups for an item that keeps ending Unknown' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc9'
        try {
            1..3 | ForEach-Object { $svc.UploadDrops.Enqueue(-1) }   # the same row three times: never created
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile -Options { param($o) $o.MaxAttempts = 10; $o.ReconcileRounds = 2 }
            $r.Result.ReconciledReimported | Should -Be 2
            $r.Result.ReconcileExhausted | Should -Be 1
            $r.Result.Unknown | Should -Be 1
            $r.Result.JobUnknown | Should -Be 1
            $r.Result.Done | Should -Be 2
            ($r.Result.Problems | Where-Object { $_ -like '*it was looked up 2 time(s)*' }).Count | Should -Be 1
        } finally { $store.Dispose() }
    }
    It 'respects -MaxAttempts: a row whose attempts are used and that is not in the target ends Failed (nothing created), not Unknown' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc10'
        try {
            1..2 | ForEach-Object { $svc.UploadDrops.Enqueue(-1) }
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile -Options { param($o) $o.MaxAttempts = 2 }
            $r.Result.ReconciledReimported | Should -Be 1
            $r.Result.Failed | Should -Be 1
            $r.Result.Unknown | Should -Be 0
            $r.Result.JobFailed | Should -Be 1
            $r.Result.JobUnknown | Should -Be 0
            $svc.ItemsOf($script:Dst).Count | Should -Be 2
        } finally { $store.Dispose() }
    }
    It 'looks the Unknown rows up while the queue goes on once their delay is over' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 8
        Add-Targets $svc
        $store = New-ExportedStore $svc 'rc11'
        $svc.LatencyMs = 80
        try {
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc' -Reconcile -Options { param($o) $o.ReconcileDelaySeconds = 0.3 }
            $r.Result.ReconciledCreated | Should -Be 1
            $r.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst).Count | Should -Be 8
            $imports = @(0..($svc.Log.Count - 1) | Where-Object { $svc.Log[$_] -like 'POST*/importItem*' })
            @(Get-TargetListings $svc)[0] | Should -BeLessThan $imports[-1] -Because 'the lookup did not wait for the end of the queue'
        } finally { $store.Dispose() }
    }
    It 'settles several lost items with four workers: one lookup at a time, no duplicate, no wrong match' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 12
        Add-Targets $svc
        foreach ($i in 1..12) { $svc.ChangeItem($script:Src, "I$i", (New-Message "Inbox $i" ($i * 100)), "ck-I$i") }   # distinct sizes
        $store = New-ExportedStore $svc 'rc12'
        $svc.LatencyMs = 20
        try {
            $svc.CreateOnLostAnswer = $true
            1..3 | ForEach-Object { $svc.UploadDrops.Enqueue(-1) }
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 4 -Name 'rc' -Reconcile
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.ReconciledCreated | Should -Be 3
            $r.Result.Unknown | Should -Be 0
            $r.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst).Count | Should -Be 12
            $ids = @(Get-Row $store $r 'Done' | ForEach-Object TargetItemId)
            @($ids | Select-Object -Unique).Count | Should -Be 12
        } finally { $store.Dispose() }
    }
    It 'Get-MbtImportUnknown -Check does not list as a candidate an older item of the same size' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $null = $svc.AddItem($script:Dst, 'D-INBOX', 'OLD1', (New-Message 'Inbox 1'), 'ck-old')
        $store = New-ExportedStore $svc 'rc13'
        try {
            $null = Set-MbtSession -Graph (New-FakeClient -Service $svc)
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDrops.Enqueue(-1)
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'rc'
            $r.Result.Unknown | Should -Be 1
            $rows = @(Get-MbtImportUnknown -Store $store -Check)
            $rows.Count | Should -Be 1
            $rows[0].Candidates.Count | Should -Be 1
            $rows[0].Candidates | Should -Not -Contain 'OLD1'
        } finally { $store.Dispose() }
    }
}

Describe 'Auto-expanding archive: export follows the expanded folders into the auxiliary archives' {
    BeforeAll {
        function Get-ExportedCount($store, [string]$Mailbox) { (@($store.GetItemStats($Mailbox) | Where-Object State -EQ 'Exported')[0]).Count }
    }
    It 'exports the main archive and the items of its expanded folders, each recorded with the mailbox that really holds it' {
        $svc = New-ExpandedArchive
        $store = New-TestStoreAt 'ax1'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Completed | Should -BeTrue
            $r.Result.Failed | Should -Be 0
            $r.Result.Exported | Should -Be 9
            $r.Result.Inventoried | Should -Be 9
            $r.Result.FoldersMoved | Should -Be 2
            @($r.Result.AuxMailboxes) | Should -Be @($script:Aux1)
            @($store.GetMailboxes() | ForEach-Object { "$($_.Kind) $($_.Id)" }) | Should -Be @("Archive $($script:Arch)", "AuxArchive $($script:Aux1)")
            # the real mailbox of each item
            $store.GetItem($script:Aux1, 'M1').MailboxId | Should -Be $script:Aux1
            $store.GetItem($script:Aux1, 'S2').MailboxId | Should -Be $script:Aux1
            $store.GetItem($script:Arch, 'A1').MailboxId | Should -Be $script:Arch
            $store.FindItemKey($script:Arch, 'M1') | Should -BeNullOrEmpty -Because 'the item is not recorded under the main archive as well'
            Get-ExportedCount $store $script:Aux1 | Should -Be 6
            Get-ExportedCount $store $script:Arch | Should -Be 3
            # the folder is recorded once, in the main archive tree, and counts the items that live in the auxiliary archive
            $tree = $store.GetFolderTree($script:Arch)
            $old = Find-Node $tree 'Old mail'
            $old.Own.Exported | Should -Be 4
            $old.Total.Exported | Should -Be 6
            (Find-Node $tree 'Y2019').Own.Exported | Should -Be 2
            $store.GetItem($script:Aux1, 'M1').FolderKey | Should -Be $old.FolderKey
            @($store.GetFolderTree($script:Aux1)).Count | Should -Be 0 -Because 'the folder is not recorded a second time under the auxiliary archive'
            # the auxiliary archive was asked for its own items (no export of them from the main archive)
            (@($svc.ExportCalls | Where-Object Mailbox -EQ $script:Aux1 | ForEach-Object Items) | Measure-Object -Sum).Sum | Should -Be 6
            (@($svc.ExportCalls | Where-Object Mailbox -EQ $script:Arch | ForEach-Object Items) | Measure-Object -Sum).Sum | Should -Be 3
            # locations: how and when
            $locs = @($store.GetLocations())
            $aux = $locs | Where-Object MailboxId -EQ $script:Aux1
            $aux.Kind | Should -Be 'AuxArchive'
            $aux.Source | Should -Be 'Redirect308'
            $aux.Detail | Should -BeLike "*$($script:Arch)*"
            ($locs | Where-Object MailboxId -EQ $script:Arch).Kind | Should -Be 'Archive'
            ($r.Messages | Where-Object { $_ -like "auxiliary archive discovered: $($script:Aux1)*" }).Count | Should -Be 1
            ($r.Messages | Where-Object { $_ -like "folder 'Old mail' is an expanded folder*" }).Count | Should -Be 1
            # a second run does nothing and records nothing twice
            $again = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $again.Result.Exported | Should -Be 0
            $store.GetStats().Items | Should -Be 9
            @($store.GetLocations()).Count | Should -Be 2
        } finally { $store.Dispose() }
    }
    It 'lists every folder of the main archive, whatever its announced item count (an expanded folder announces none)' {
        $svc = New-ExpandedArchive
        $store = New-TestStoreAt 'ax2'
        try {
            $null = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $svc.RequestsOf("mailboxes/$($script:Arch)/folders/F-MOVED/items") | Should -BeGreaterThan 0
        } finally { $store.Dispose() }
        # a primary mailbox keeps skipping the folders that announce no item
        $svc2 = New-SourceService
        $store2 = New-TestStoreAt 'ax2b'
        try { $null = Invoke-ExportJob -Service $svc2 -Store $store2; $svc2.RequestsOf('F-EMPTY/items') | Should -Be 0 } finally { $store2.Dispose() }
    }
    It 'learns that an item lives in an auxiliary archive from ErrorArchiveFolderMovedPermanently alone, and moves its record there' {
        $svc = New-ExpandedArchive
        $svc.RedirectFolders = $false; $svc.ListMovedFromMain = $true   # the main archive lists the items itself, no 308
        $store = New-TestStoreAt 'ax3'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $r.Result.Failed | Should -Be 0
            $r.Result.Exported | Should -Be 9
            $r.Result.ItemsMoved | Should -Be 6
            $store.FindItemKey($script:Arch, 'M1') | Should -BeNullOrEmpty
            $m1 = $store.GetItem($script:Aux1, 'M1')
            $m1.State | Should -Be 'Exported'
            $m1.FolderId | Should -Be 'F-MOVED'
            ($store.GetLocations() | Where-Object MailboxId -EQ $script:Aux1).Source | Should -Be 'ExportMoved'
            (Find-Node $store.GetFolderTree($script:Arch) 'Old mail').Own.Exported | Should -Be 4
            # listed through the main archive again (now with the 308): the same items, no duplicate record
            $svc.RedirectFolders = $true; $svc.ListMovedFromMain = $false
            $again = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $again.Result.Exported | Should -Be 0
            $store.GetStats().Items | Should -Be 9
        } finally { $store.Dispose() }
    }
    It 'does not record an item twice when it was first recorded under the main archive (store written before the auxiliary archive was known)' {
        $store = New-TestStoreAt 'ax4'
        try {
            $null = $store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:Arch; Kind = 'Archive' })
            $null = $store.UpsertFolders([MailboxTransfer.StoreFolder[]]@([MailboxTransfer.StoreFolder]@{ MailboxId = $script:Arch; FolderId = 'F-MOVED'; DisplayName = 'Old mail' }))
            $null = $store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@([MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:Arch; ItemId = 'M1'; FolderId = 'F-MOVED'; ChangeKey = 'c'; Size = 10 }))
            $null = $store.EnsureMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:Aux1; Kind = 'AuxArchive' })
            $null = $store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@([MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:Aux1; ItemId = 'M1'; FolderId = 'F-MOVED'; FolderMailboxId = $script:Arch; ChangeKey = 'c'; Size = 10 }))
            $store.GetStats().Items | Should -Be 1
            $store.FindItemKey($script:Arch, 'M1') | Should -BeNullOrEmpty
            $store.FindItemKey($script:Aux1, 'M1') | Should -Not -BeNullOrEmpty
            # EnsureMailbox never changes the kind of a mailbox that exists
            $null = $store.EnsureMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:Arch; Kind = 'AuxArchive' })
            ($store.GetMailboxes() | Where-Object Id -EQ $script:Arch).Kind | Should -Be 'Archive'
        } finally { $store.Dispose() }
    }
    It 'exports one auxiliary archive by itself; the folders the main archive tree records are not recorded again' {
        $svc = New-ExpandedArchive
        $store = New-TestStoreAt 'ax5'
        try {
            $null = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Arch -Kind 'Archive'
            $r = Invoke-ExportJob -Service $svc -Store $store -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $r.Result.Completed | Should -BeTrue
            $r.Result.Folders | Should -Be 2
            $r.Result.Exported | Should -Be 0 -Because 'its items are already exported'
            @($store.GetFolderTree($script:Aux1)).Count | Should -Be 0
            $store.GetStats().Items | Should -Be 9
        } finally { $store.Dispose() }
        # alone in a store, it records its own folders
        $alone = New-TestStoreAt 'ax5b'
        try {
            $r = Invoke-ExportJob -Service $svc -Store $alone -Mailbox $script:Aux1 -Kind 'AuxArchive'
            $r.Result.Exported | Should -Be 6
            $alone.GetMailboxes()[0].Kind | Should -Be 'AuxArchive'
            (Find-Node $alone.GetFolderTree($script:Aux1) 'Old mail').Total.Exported | Should -Be 6
            ($alone.GetLocations() | Where-Object MailboxId -EQ $script:Aux1).Source | Should -Be 'ExportTarget'
        } finally { $alone.Dispose() }
    }
}

Describe 'Auto-expanding archive: import follows the service into the auxiliary archives' {
    BeforeAll {
        function Get-DstSessionImports($svc) { @($svc.Log | Where-Object { $_ -like "POST *Mailboxes('$($script:Dst)')/importItem*" }).Count }
        function Get-Rows($store, $result, [string]$State) { @($store.GetImportItems($result.Result.JobId, [MailboxTransfer.StoreImportState]::$State, 100)) }
    }
    It 'lands the items in the auxiliary archive, records it, keeps the route and sends the next items there directly' {
        $svc = New-SourceService
        Add-Targets $svc
        $svc.MovedFolderToMailbox['D-INBOX'] = $script:Aux
        $store = New-ExportedStore $svc 'ai1'
        try {
            $svc.RedirectFolders = $true
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Done | Should -Be 10
            $r.Result.Failed | Should -Be 0
            $r.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Dst).Count | Should -Be 0
            $svc.ItemsOf($script:Aux).Count | Should -Be 10 -Because 'every item exists once, in the auxiliary archive'
            # the child folder requested under the expanded folder was created where the service redirected (308), the items followed
            $svc.FoldersOf($script:Aux).Count | Should -Be 2
            # one refused request (409) per folder, then straight to the auxiliary archive
            Get-DstSessionImports $svc | Should -Be 2
            $r.Result.SessionsCreated | Should -Be 2
            $r.Result.FoldersInAux | Should -Be 2
            @($r.Result.AuxArchives) | Should -Be @($script:Aux)
            # the rows target the mailbox that holds the folder, and the routes are kept for a resumed run
            @(Get-Rows $store $r 'Done' | Where-Object TargetMailbox -NE $script:Aux).Count | Should -Be 0
            $store.GetImportRoutes($r.Result.JobId).Count | Should -Be 2
            # the location is recorded, with how it was discovered
            $loc = $store.GetLocations() | Where-Object MailboxId -EQ $script:Aux
            $loc.Kind | Should -Be 'AuxArchive'
            $loc.Source | Should -BeIn 'Redirect308', 'Import409'
        } finally { $store.Dispose() }
    }
    It 'records the auxiliary archive named by a 409 as Import409 and says so' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        Add-Targets $svc
        $svc.MovedFolderToMailbox['D-INBOX'] = $script:Aux
        $store = New-ExportedStore $svc 'ai1b'
        try {
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 3
            $loc = $store.GetLocations() | Where-Object MailboxId -EQ $script:Aux
            $loc.Source | Should -Be 'Import409'
            $loc.Kind | Should -Be 'AuxArchive'
            ($r.Messages | Where-Object { $_ -like "the folder lives in the auxiliary archive $($script:Aux)*" }).Count | Should -Be 1 -Because 'only the first item of the folder is refused, the next ones go straight there'
            ($r.Messages | Where-Object { $_ -like "auxiliary archive discovered: $($script:Aux) (Import409)" }).Count | Should -Be 1
            Get-DstSessionImports $svc | Should -Be 1
        } finally { $store.Dispose() }
    }
    It 'looks up an Unknown row in the auxiliary archive that holds the folder: found created, no duplicate' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 4
        Add-Targets $svc
        $svc.MovedFolderToMailbox['D-INBOX'] = $script:Aux
        $store = New-ExportedStore $svc 'ai2'
        try {
            $svc.CreateOnLostAnswer = $true
            $svc.UploadDropMailbox = $script:Aux
            $svc.UploadDrops.Enqueue(-1)   # the answer of the first import into the auxiliary archive is lost, the item exists
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Reconcile
            $r.Job.Error | Should -BeNullOrEmpty
            $r.Result.Done | Should -Be 4
            $r.Result.Unknown | Should -Be 0
            $r.Result.ReconciledCreated | Should -Be 1
            $r.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Aux, 'D-INBOX').Count | Should -Be 4 -Because 'the item that exists is not imported a second time'
            $svc.ItemsOf($script:Dst).Count | Should -Be 0
            $svc.RequestsOf("mailboxes/$($script:Aux)/folders/D-INBOX/items") | Should -BeGreaterThan 0 -Because 'the reconciliation lists the folder in the mailbox that holds it'
        } finally { $store.Dispose() }
    }
    It 'resumes with the routes learned, and sends a retried row straight to the auxiliary archive' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 5
        Add-Targets $svc
        $svc.MovedFolderToMailbox['D-INBOX'] = $script:Aux
        $store = New-ExportedStore $svc 'ai3'
        $root = $store.Root
        try {
            # the first item succeeds after the redirect (2 calls: refused, then accepted); the next two are refused (400, nothing created) and stay Failed
            $svc.ImportFailures.Enqueue(0); $svc.ImportFailures.Enqueue(0); $svc.ImportFailures.Enqueue(400); $svc.ImportFailures.Enqueue(400)
            $r1 = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1 -Name 'res'
            $r1.Result.Done | Should -Be 3
            $r1.Result.Failed | Should -Be 2
            Get-DstSessionImports $svc | Should -Be 1
        } finally { $store.Dispose() }
        $again = [MailboxTransfer.ExportStore]::Open($root)
        try {
            $r2 = Invoke-ImportJob -Service $svc -Store $again -Concurrency 1 -Name 'res'
            $r2.Result.JobComplete | Should -BeTrue
            $svc.ItemsOf($script:Aux, 'D-INBOX').Count | Should -Be 5
            Get-DstSessionImports $svc | Should -Be 1 -Because 'the second run knew the route and never asked the main archive again'
            ($r2.Messages | Where-Object { $_ -like 'routes: 1 folder(s)*' }).Count | Should -Be 1
        } finally { $again.Dispose() }
    }
    It 'finds the folder in the auxiliary archive by its path when that mailbox lists it with another id' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        $null = $svc.AddMailbox($script:Dst); $null = $svc.AddFolder($script:Dst, 'D-INBOX', 'Inbox')
        $null = $svc.AddMailbox($script:Aux); $null = $svc.AddFolder($script:Aux, 'X-INBOX', 'Inbox')
        $svc.MovedFolderToMailbox['D-INBOX'] = $script:Aux
        $store = New-ExportedStore $svc 'ai4'
        try {
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 3
            $svc.ItemsOf($script:Aux, 'X-INBOX').Count | Should -Be 3
            @(Get-Rows $store $r 'Done' | Where-Object TargetFolder -NE 'X-INBOX').Count | Should -Be 0
            ($r.Messages | Where-Object { $_ -like "the auxiliary archive $($script:Aux) lists the folder 'Inbox' with another id: X-INBOX" }).Count | Should -Be 1
            $store.GetImportRoutes($r.Result.JobId)["$($script:Dst)|D-INBOX"].Value | Should -Be 'X-INBOX'
        } finally { $store.Dispose() }
    }
    It 'never creates a folder in an auxiliary archive to receive an item: the row fails with nothing created' {
        $svc = New-SourceService -ProjectItems 0 -InboxItems 3
        $null = $svc.AddMailbox($script:Dst); $null = $svc.AddFolder($script:Dst, 'D-INBOX', 'Inbox')
        $null = $svc.AddMailbox($script:Aux); $null = $svc.AddFolder($script:Aux, 'X-OTHER', 'Other')
        $svc.MovedFolderToMailbox['D-INBOX'] = $script:Aux
        $store = New-ExportedStore $svc 'ai5'
        try {
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 1
            $r.Result.Done | Should -Be 0
            $r.Result.Failed | Should -Be 3
            $r.Result.Unknown | Should -Be 0
            @($svc.FoldersOf($script:Aux).Name) | Should -Be @('Other')
            $svc.ItemsOf($script:Aux).Count | Should -Be 0
            $svc.ItemsOf($script:Dst).Count | Should -Be 0
            (Get-Rows $store $r 'Failed')[0].LastError | Should -BeLike "*lives in the auxiliary archive $($script:Aux)*does not list a folder of that path*folders are not created in an auxiliary archive*Nothing was created*"
        } finally { $store.Dispose() }
    }
    It 'imports into an auxiliary archive named explicitly as the target' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = New-ExportedStore $svc 'ai6'
        try {
            $r = Invoke-ImportJob -Service $svc -Store $store -Concurrency 2 -Options { param($o) $o.TargetMailboxId = $script:Aux }
            $r.Result.Done | Should -Be 10
            $svc.ItemsOf($script:Aux).Count | Should -Be 10
            $svc.ItemsOf($script:Dst).Count | Should -Be 0
            $r.Result.SessionsCreated | Should -Be 1
            $r.Result.FoldersInAux | Should -Be 0
        } finally { $store.Dispose() }
    }
}

Describe 'Mailbox locations: discovery, record, status' {
    BeforeAll {
        $script:LocTenant = $script:Tenant
        $script:LocP = '1a2b3c03-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $script:LocM = '1a2b3c04-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $script:LocA1 = '11111111-aaaa-bbbb-cccc-111111111111'
        $script:LocA2 = '22222222-aaaa-bbbb-cccc-222222222222'
        function Get-LocId([string]$Guid) { "MBX:$Guid@$($script:LocTenant)" }
        $script:LocText = @"
MailboxLocationType MailboxGuid
------------------- -----------
Primary             $($script:LocP)
MainArchive         $($script:LocM)
AuxArchive          $($script:LocA1)
AuxArchive          $($script:LocA2)
Aggregated          33333333-aaaa-bbbb-cccc-333333333333
"@
    }
    It 'reads the output of Get-MailboxLocation as a table, a CSV and JSON' {
        $table = Join-Path $script:Tmp 'loc-table.txt'; Set-Content -LiteralPath $table -Value $script:LocText
        $csv = Join-Path $script:Tmp 'loc.csv'
        Set-Content -LiteralPath $csv -Value ("MailboxGuid,MailboxLocationType`r`n$($script:LocP),Primary`r`n$($script:LocM),MainArchive`r`n$($script:LocA1),AuxArchive`r`n$($script:LocA2),AuxArchive")
        $json = Join-Path $script:Tmp 'loc.json'
        Set-Content -LiteralPath $json -Value (@(
            @{ MailboxGuid = $script:LocP; MailboxLocationType = 'Primary' }, @{ MailboxGuid = $script:LocM; MailboxLocationType = 'MainArchive' },
            @{ MailboxGuid = $script:LocA1; MailboxLocationType = 'AuxArchive' }, @{ MailboxGuid = $script:LocA2; MailboxLocationType = 'AuxArchive' }) | ConvertTo-Json)
        foreach ($file in $table, $csv, $json) {
            $locs = @(Read-MbtMailboxLocationFile -Path $file -TenantId $script:LocTenant)
            $locs.Count | Should -Be 4 -Because $file
            @($locs | ForEach-Object { "$($_.Kind) $($_.MailboxId)" }) | Should -Be @(
                "Primary $(Get-LocId $script:LocP)", "Archive $(Get-LocId $script:LocM)", "AuxArchive $(Get-LocId $script:LocA1)", "AuxArchive $(Get-LocId $script:LocA2)") -Because $file
        }
        { Read-MbtMailboxLocationFile -Path (Join-Path $script:Tmp 'nope.csv') -TenantId $script:LocTenant } | Should -Throw '*not found*'
    }
    It 'merges the operator file, the parameters and the store, and records them with their source' {
        $store = New-TestStoreAt 'loc1'
        try {
            $file = Join-Path $script:Tmp 'loc-merge.txt'; Set-Content -LiteralPath $file -Value $script:LocText
            # auxiliary archives learned from redirects earlier
            $null = $store.RecordLocation([MailboxTransfer.StoreMailboxLocation]@{ MailboxId = (Get-LocId $script:LocA1); Kind = 'AuxArchive'; Source = 'Redirect308'; Detail = 'learned'; Rank = 1 })
            $null = $store.RecordLocation([MailboxTransfer.StoreMailboxLocation]@{ MailboxId = (Get-LocId '99999999-aaaa-bbbb-cccc-999999999999'); Kind = 'AuxArchive'; Source = 'Import409'; Detail = 'learned'; Rank = 1 })
            $locs = @(Get-MbtMailboxLocation -UserPrincipalName 'adele@contoso.test' -TenantId $script:LocTenant -Store $store -LocationsFile $file)
            @($locs | ForEach-Object Kind) | Should -Be @('Primary', 'Archive', 'AuxArchive', 'AuxArchive', 'AuxArchive')
            ($locs | Where-Object MailboxId -EQ (Get-LocId $script:LocA1)).Source | Should -Be 'OperatorFile' -Because 'what the operator gives beats what was learned'
            ($locs | Where-Object MailboxId -EQ (Get-LocId '99999999-aaaa-bbbb-cccc-999999999999')).Learned | Should -BeTrue
            # recorded in the store, with the better source for the one that was only learned
            $saved = @($store.GetLocations())
            $saved.Count | Should -Be 5
            ($saved | Where-Object MailboxId -EQ (Get-LocId $script:LocA1)).Source | Should -Be 'OperatorFile'
            ($saved | Where-Object MailboxId -EQ (Get-LocId $script:LocM)).Kind | Should -Be 'Archive'
            ($saved | Where-Object MailboxId -EQ (Get-LocId $script:LocP)).Upn | Should -Be 'adele@contoso.test'
            # a redirect never overrules what was given
            $learner = [MailboxTransfer.AuxLocationLearner]::new($store, $null, [string[]]@(), $null)
            $learner.Learn((Get-LocId $script:LocM), 'Redirect308', 'x') | Should -BeFalse
            ($store.GetLocations() | Where-Object MailboxId -EQ (Get-LocId $script:LocM)).Kind | Should -Be 'Archive'
            # piped objects (what Get-MailboxLocation returns)
            $piped = [pscustomobject]@{ MailboxGuid = [guid]'44444444-aaaa-bbbb-cccc-444444444444'; MailboxLocationType = 'AuxArchive' } | Get-MbtMailboxLocation -TenantId $script:LocTenant
            $piped.MailboxId | Should -Be (Get-LocId '44444444-aaaa-bbbb-cccc-444444444444')
            $piped.Source | Should -Be 'OperatorInput'
        } finally { $store.Dispose() }
    }
    It 'reads the primary and the archive from Graph beta, falls back to v1.0 (primary only), and says nothing without User.Read.All' {
        $svc = New-SourceService
        $null = Set-MbtSession -Graph (New-FakeClient -Service $svc)
        $upn = 'adele@contoso.test'
        # no permission: 403 on both
        @(Get-MbtMailboxLocation -UserPrincipalName $upn -TenantId $script:LocTenant).Count | Should -Be 0
        $svc.RequestsOf('/beta/users/') | Should -BeGreaterThan 0
        # beta: primary and archive
        $svc.ExchangeSettingsBetaJson = '{"primaryMailboxId":"' + (Get-LocId $script:LocP) + '","inPlaceArchiveMailboxId":"' + (Get-LocId $script:LocM) + '"}'
        $locs = @(Get-MbtMailboxLocation -UserPrincipalName $upn -TenantId $script:LocTenant)
        @($locs | ForEach-Object { "$($_.Kind) $($_.Source)" }) | Should -Be @('Primary Graph', 'Archive Graph')
        # beta refused, v1.0 answers: the primary only (v1.0 has no archive property)
        $svc.ExchangeSettingsBetaJson = $null
        $svc.ExchangeSettingsV1Json = '{"primaryMailboxId":"' + (Get-LocId $script:LocP) + '"}'
        $locs = @(Get-MbtMailboxLocation -UserPrincipalName $upn -TenantId $script:LocTenant)
        @($locs | ForEach-Object Kind) | Should -Be @('Primary')
        # -SkipGraph asks nothing
        $before = $svc.Log.Count
        $null = Get-MbtMailboxLocation -UserPrincipalName $upn -TenantId $script:LocTenant -SkipGraph
        $svc.Log.Count | Should -Be $before
    }
    It 'Resolve-MbtMailboxId: an explicit id wins; the kind is taken from the store or the file; an unknown auxiliary archive is explained' {
        $store = New-TestStoreAt 'loc2'
        try {
            $upn = 'adele@contoso.test'
            (Resolve-MbtMailboxId -Mailbox AuxArchive -TenantId $script:LocTenant -UserPrincipalName $upn -MailboxId $script:LocA1) | Should -Be (Get-LocId $script:LocA1)
            { Resolve-MbtMailboxId -Mailbox AuxArchive -TenantId $script:LocTenant -UserPrincipalName $upn -Store $store } | Should -Throw '*No auxiliary archive*Get-MailboxLocation*'
            $null = $store.RecordLocation([MailboxTransfer.StoreMailboxLocation]@{ MailboxId = (Get-LocId $script:LocA1); Kind = 'AuxArchive'; Source = 'Redirect308'; Rank = 1 })
            (Resolve-MbtMailboxId -Mailbox AuxArchive -TenantId $script:LocTenant -UserPrincipalName $upn -Store $store) | Should -Be (Get-LocId $script:LocA1)
            $null = $store.RecordLocation([MailboxTransfer.StoreMailboxLocation]@{ MailboxId = (Get-LocId $script:LocA2); Kind = 'AuxArchive'; Source = 'Redirect308'; Rank = 1 })
            { Resolve-MbtMailboxId -Mailbox AuxArchive -TenantId $script:LocTenant -UserPrincipalName $upn -Store $store } | Should -Throw '*2 AuxArchive mailboxes*Choose one with -MailboxId*'
            $file = Join-Path $script:Tmp 'loc-resolve.txt'; Set-Content -LiteralPath $file -Value $script:LocText
            (Resolve-MbtMailboxId -Mailbox Archive -TenantId $script:LocTenant -UserPrincipalName $upn -LocationsFile $file) | Should -Be (Get-LocId $script:LocM)
        } finally { $store.Dispose() }
    }
    It 'shows the locations in the store status' {
        $store = New-TestStoreAt 'loc3'
        try {
            $null = $store.RecordLocation([MailboxTransfer.StoreMailboxLocation]@{ MailboxId = (Get-LocId $script:LocM); Kind = 'Archive'; Source = 'Parameter'; Rank = 2 })
            $null = $store.RecordLocation([MailboxTransfer.StoreMailboxLocation]@{ MailboxId = (Get-LocId $script:LocA1); Kind = 'AuxArchive'; Source = 'Redirect308'; Detail = 'asked X'; Rank = 1 })
            $text = Show-MbtStoreStatus -Status (Get-MbtStoreStatus -Store $store) *>&1 | Out-String
            $text | Should -Match 'Mailbox locations'
            $text | Should -Match "Archive\s+$([regex]::Escape((Get-LocId $script:LocM)))\s+\[Parameter"
            $text | Should -Match "AuxArchive\s+$([regex]::Escape((Get-LocId $script:LocA1)))\s+\[Redirect308.*asked X"
        } finally { $store.Dispose() }
    }
    It 'a store written before the location tables existed is upgraded when opened for writing and stays readable read-only, schema version unchanged' {
        $path = Join-Path $script:Tmp 'loc-old'
        $s = [MailboxTransfer.ExportStore]::Create($path, 'old', 'tests'); $s.Dispose()
        $c = [Microsoft.Data.Sqlite.SqliteConnection]::new("Data Source=$(Join-Path $path 'catalog.db');Pooling=False"); $c.Open()
        foreach ($sql in 'DROP TABLE mailboxLocation', 'DROP TABLE importRoute') { $cmd = $c.CreateCommand(); $cmd.CommandText = $sql; [void]$cmd.ExecuteNonQuery() }
        $c.Dispose()
        $ro = [MailboxTransfer.ExportStore]::OpenReadOnly($path)
        try { @($ro.GetLocations()).Count | Should -Be 0; $ro.GetImportRoutes(1).Count | Should -Be 0 } finally { $ro.Dispose() }
        $rw = [MailboxTransfer.ExportStore]::Open($path)
        try {
            $null = $rw.RecordLocation([MailboxTransfer.StoreMailboxLocation]@{ MailboxId = (Get-LocId $script:LocA1); Kind = 'AuxArchive'; Source = 'Parameter'; Rank = 2 })
            @($rw.GetLocations()).Count | Should -Be 1
        } finally { $rw.Dispose() }
        $c = [Microsoft.Data.Sqlite.SqliteConnection]::new("Data Source=$(Join-Path $path 'catalog.db');Pooling=False"); $c.Open()
        $cmd = $c.CreateCommand(); $cmd.CommandText = 'PRAGMA user_version'; $version = $cmd.ExecuteScalar(); $c.Dispose()
        $version | Should -Be 1 -Because 'the schema version is not raised: an older build still opens the store'
    }
}

Describe 'Auto-expanding archive: Invoke-MailboxTransfer.ps1' {
    BeforeAll {
        $script:Entry = Join-Path $script:Root 'Invoke-MailboxTransfer.ps1'
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
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtEntryClient, MbtEntryExit -Scope Global -ErrorAction SilentlyContinue }
    It '-Mailbox Archive exports the expanded folders too; Status shows the locations' {
        $svc = New-ExpandedArchive
        $store = Join-Path $script:Tmp 'ea-arch'
        $text = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'Archive'; ArchiveMailboxId = $script:Arch; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ea-tmp1'); Concurrency = 2 }
        $text | Should -Match 'export finished: 9 items exported'
        $text | Should -Match 'auto-expanding archive: 2 expanded folder\(s\).*1 auxiliary archive\(s\)'
        $text | Should -Not -Match '(?m)^ {6}x {2}\S'
        $status = Invoke-Entry $svc @{ Mode = 'Status'; StorePath = $store }
        $status | Should -Match 'Mailbox locations'
        $status | Should -Match "AuxArchive\s+$([regex]::Escape($script:Aux1))\s+\[Redirect308"
        $status | Should -Match "AuxArchive mailbox $([regex]::Escape($script:Aux1))"
        $status | Should -Match 'Old mail\s+6 items\s+exported\s+6'
    }
    It '-Mailbox AuxArchive -MailboxId exports one explicit location; the kind of -MailboxId alone comes from the store' {
        $svc = New-ExpandedArchive
        $store = Join-Path $script:Tmp 'ea-aux'
        $text = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'AuxArchive'; MailboxId = $script:Aux1; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ea-tmp2') }
        $text | Should -Match 'export finished: 6 items exported'
        $text | Should -Match "exporting the AuxArchive mailbox $([regex]::Escape($script:Aux1))"
        # -MailboxId alone: unknown to the store -> asks what it is; known -> its kind is used
        { Invoke-Entry $svc @{ Mode = 'Export'; MailboxId = $script:Arch; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ea-tmp2') } } | Should -Throw '*is not a known location*'
        $again = Invoke-Entry $svc @{ Mode = 'Export'; MailboxId = $script:Aux1; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ea-tmp2') }
        $again | Should -Match "exporting the AuxArchive mailbox $([regex]::Escape($script:Aux1))"
    }
    It 'Status -Locations signs in and lists Primary, MainArchive and AuxArchive from Graph, the operator file and the store, with their source' {
        $svc = New-ExpandedArchive
        $store = Join-Path $script:Tmp 'ea-loc'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'Archive'; ArchiveMailboxId = $script:Arch; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ea-tmp7') }
        $file = Join-Path $script:Tmp 'ea-loc.txt'
        Set-Content -LiteralPath $file -Value ("Primary $('1a2b3c03-5e7f-4a1b-9c2d-3e4f5a6b7c8d')`nMainArchive $($script:Arch.Substring(4, 36))`nAuxArchive 77777777-aaaa-bbbb-cccc-777777777777")
        $svc.ExchangeSettingsBetaJson = '{"primaryMailboxId":"MBX:1a2b3c03-5e7f-4a1b-9c2d-3e4f5a6b7c8d@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d"}'
        $text = Invoke-Entry $svc @{ Mode = 'Status'; Locations = $true; StorePath = $store; LocationsFile = $file; UserPrincipalName = 'adele.vance@contoso.com' }
        $text | Should -Match 'Primary\s+MBX:1a2b3c03-5e7f-4a1b-9c2d-3e4f5a6b7c8d@\S+\s+\[Graph\]'
        $text | Should -Match 'AuxArchive\s+MBX:77777777-aaaa-bbbb-cccc-777777777777@\S+\s+\[OperatorFile\]'
        $text | Should -Match "AuxArchive\s+$([regex]::Escape($script:Aux1))\s+\[Redirect308, learned from a redirect\]"
        $text | Should -Match "Archive\s+$([regex]::Escape($script:Arch))"
    }
    It '-Mailbox All exports the primary, the archive with its expanded folders, then the auxiliary archives learned' {
        $svc = New-SourceService
        Add-ExpandedArchive $svc
        $store = Join-Path $script:Tmp 'ea-all'
        $text = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'All'; PrimaryMailboxId = $script:Src; ArchiveMailboxId = $script:Arch; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ea-tmp3'); Concurrency = 2 }
        $text | Should -Match 'exporting the Primary mailbox'
        $text | Should -Match 'exporting the Archive mailbox'
        $text | Should -Match 'exporting the AuxArchive mailbox'
        $text | Should -Not -Match '(?m)^ {6}x {2}\S'
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try {
            @($s.GetMailboxes() | ForEach-Object { $_.Kind.ToString() }) | Should -Be @('Primary', 'Archive', 'AuxArchive')
            $s.GetStats().Items | Should -Be 19
        } finally { $s.Dispose() }
        $global:MbtEntryExit | Should -Not -Be 1
    }
    It '-Mailbox All with an archive id that cannot be listed reports it and still exports the others' {
        $svc = New-SourceService
        $store = Join-Path $script:Tmp 'ea-all2'
        $text = Invoke-Entry $svc @{ Mode = 'Export'; Mailbox = 'All'; PrimaryMailboxId = $script:Src; ArchiveMailboxId = $script:Arch; StorePath = $store; TempPath = (Join-Path $script:Tmp 'ea-tmp4'); Concurrency = 2 }
        $text | Should -Match 'export finished: 10 items exported'
        $text | Should -Match '(?m)^ {6}x {2}.*export of the Archive mailbox failed'
        $global:MbtEntryExit | Should -Be 1
    }
    It 'Import -Mailbox AuxArchive -MailboxId imports into that auxiliary archive and records it; -Mailbox All is refused' {
        $svc = New-SourceService
        Add-Targets $svc
        $store = Join-Path $script:Tmp 'ea-imp'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:Src; TempPath = (Join-Path $script:Tmp 'ea-tmp5') }
        $text = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; Mailbox = 'AuxArchive'; MailboxId = $script:Aux; TargetRoot = 'MBT-Aux'; Concurrency = 2 }
        $text | Should -Match "into the AuxArchive mailbox $([regex]::Escape($script:Aux))"
        $text | Should -Match 'import finished: 10 items imported'
        $svc.ItemsOf($script:Aux).Count | Should -Be 10
        $svc.ItemsOf($script:Dst).Count | Should -Be 0
        $s = Open-MbtExportStore -Path $store -ReadOnly
        try { ($s.GetLocations() | Where-Object MailboxId -EQ $script:Aux).Kind | Should -Be 'AuxArchive' } finally { $s.Dispose() }
        { Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; Mailbox = 'All'; PrimaryMailboxId = $script:Dst } } | Should -Throw '*-Mailbox All is for Export*'
    }
    It 'Import into the main archive follows the service into the auxiliary archive by itself' {
        $svc = New-SourceService
        Add-Targets $svc
        $svc.MovedFolderToMailbox['D-INBOX'] = $script:Aux
        $store = Join-Path $script:Tmp 'ea-imp2'
        $null = Invoke-Entry $svc @{ Mode = 'Export'; StorePath = $store; PrimaryMailboxId = $script:Src; TempPath = (Join-Path $script:Tmp 'ea-tmp6') }
        $svc.RedirectFolders = $true
        $text = Invoke-Entry $svc @{ Mode = 'Import'; StorePath = $store; Mailbox = 'Archive'; ArchiveMailboxId = $script:Dst; Concurrency = 1 }
        $text | Should -Match 'import finished: 10 items imported'
        $text | Should -Match 'auto-expanding archive: 2 folder\(s\) of the target live in 1 auxiliary archive\(s\)'
        $svc.ItemsOf($script:Aux).Count | Should -Be 10
        $status = Invoke-Entry $svc @{ Mode = 'Status'; StorePath = $store }
        $status | Should -Match "AuxArchive\s+$([regex]::Escape($script:Aux))"
    }
}
