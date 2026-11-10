#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Real tests of package\src\Engine.Store.cs (catalog.db + packs): lifecycle and schema, 20,000 items + large items written
    by concurrent workers (one of them >= 300 MiB streamed from a generated stream), SHA-256 re-read, process memory,
    crash recovery (in-process abandon, damaged pack, and real kills of a child pwsh with Stop-Process), pack rotation,
    folder tree / pagination / previews, import cycle (InFlight -> Unknown), throughput.

    Run:  Invoke-Pester -Path .\tests\Store.Tests.ps1 -Output Detailed
    Data goes to $env:MAILBOXTRANSFER_TEST_DIR (default: the scratchpad store-tests folder) and is deleted at the end.
#>

BeforeAll {
    $script:RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $script:PackageRoot = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:PackageRoot -WithTestSupport

    $script:Base = if ($env:MAILBOXTRANSFER_TEST_DIR) { $env:MAILBOXTRANSFER_TEST_DIR } else { Join-Path ([IO.Path]::GetTempPath()) 'MailboxTransfer-store-tests' }
    if (Test-Path -LiteralPath $script:Base) { Remove-Item -LiteralPath $script:Base -Recurse -Force }
    [void](New-Item -ItemType Directory -Path $script:Base -Force)

    $script:Metrics = [ordered]@{}
    $script:Children = [System.Collections.Generic.List[System.Diagnostics.Process]]::new()
    $script:StoreNS = 'MailboxTransfer'
    $script:NoCancel = [Threading.CancellationToken]::None

    function New-TestStore {
        param([string]$Name, [MailboxTransfer.StoreOptions]$Options)
        $path = Join-Path $script:Base $Name
        if ($Options) { [MailboxTransfer.ExportStore]::Create($path, $Name, 'tests', $Options) } else { [MailboxTransfer.ExportStore]::Create($path, $Name, 'tests') }
    }

    function Add-TestMailbox {
        param($Store, [string]$Id = 'mb1', [string]$Kind = 'Primary')
        $Store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $Id; Kind = $Kind; Upn = "$Id@contoso.test"; DisplayName = "Mailbox $Id" }) | Out-Null
    }

    function New-Jobs {
        # Jobs for item keys; seed = key; length from $Sizes (index = key - 1 - $KeyBase)
        param([long[]]$Keys, $Sizes, [long]$KeyBase = 0)
        $l = [System.Collections.Generic.List[MailboxTransfer.TestSupport.Job]]::new()
        foreach ($k in $Keys) {
            $j = [MailboxTransfer.TestSupport.Job]::new()
            $j.Key = $k; $j.Seed = $k; $j.Length = [long]$Sizes[$k - 1 - $KeyBase]
            $l.Add($j)
        }
        , $l
    }

    function New-CrashStore {
        # Store with $Count pending items (keys 1..$Count) in mailbox mb1, root folder only.
        param([string]$Name, [int]$Count)
        $s = New-TestStore $Name
        Add-TestMailbox $s
        $s.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'f0'; DisplayName = 'Root' })) | Out-Null
        $entries = [System.Collections.Generic.List[MailboxTransfer.StoreItemEntry]]::new()
        for ($i = 0; $i -lt $Count; $i++) {
            $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = [MailboxTransfer.TestSupport.Workload]::ItemId($i); FolderId = 'f0'; ChangeKey = "c$i"; ItemType = 'IPM.Note'; Size = [MailboxTransfer.TestSupport.Workload]::CrashSize($i + 1) })
        }
        $s.UpsertItems($entries) | Out-Null
        $s
    }

    function Get-CommittedEnd {
        # Largest end offset (offset+length) of the Exported items of a pack, from the catalog (read-only connection).
        param([string]$ExportPath)
        $c = [Microsoft.Data.Sqlite.SqliteConnection]::new("Data Source=$(Join-Path $ExportPath 'catalog.db');Mode=ReadOnly;Pooling=False")
        $c.Open()
        try {
            $cmd = $c.CreateCommand()
            $cmd.CommandText = 'SELECT pack_id, MAX(pack_offset + pack_length) FROM item WHERE state=1 GROUP BY pack_id'
            $r = $cmd.ExecuteReader()
            $h = @{}
            while ($r.Read()) { $h[[int]$r.GetInt32(0)] = $r.GetInt64(1) }
            $r.Dispose()
            $h
        }
        finally { $c.Dispose() }
    }

    function Invoke-Sql {
        param([string]$ExportPath, [string]$Sql)
        $c = [Microsoft.Data.Sqlite.SqliteConnection]::new("Data Source=$(Join-Path $ExportPath 'catalog.db');Mode=ReadOnly;Pooling=False")
        $c.Open()
        try {
            $cmd = $c.CreateCommand(); $cmd.CommandText = $Sql
            $r = $cmd.ExecuteReader(); $rows = [System.Collections.Generic.List[object[]]]::new()
            while ($r.Read()) { $v = [object[]]::new($r.FieldCount); [void]$r.GetValues($v); $rows.Add($v) }
            $r.Dispose()
            , $rows
        }
        finally { $c.Dispose() }
    }

    function Start-CrashChild {
        param([string]$ExportPath, [string]$Mode, [string]$Signal, [int]$Threads = 4)
        $pwsh = (Get-Process -Id $PID).Path
        $script1 = Join-Path $PSScriptRoot 'support\crash-writer.ps1'
        # Start-Process does not quote the elements of -ArgumentList: quote the paths (the repository path contains a space)
        $args1 = @('-NoProfile', '-NonInteractive', '-File', "`"$script1`"", '-PackageRoot', "`"$($script:PackageRoot)`"", '-ExportPath', "`"$ExportPath`"", '-SignalFile', "`"$Signal`"", '-Mode', $Mode, '-Threads', $Threads)
        $p = Start-Process -FilePath $pwsh -ArgumentList $args1 -PassThru -WindowStyle Hidden `
            -RedirectStandardOutput "$Signal.out" -RedirectStandardError "$Signal.err"
        $script:Children.Add($p)
        $p
    }

    function Wait-ForFile {
        param([string]$Path, $Process, [int]$TimeoutSec = 120)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        while (-not (Test-Path -LiteralPath $Path)) {
            if ($Process.HasExited) { throw "Child process exited (code $($Process.ExitCode)) before writing $Path : $(Get-Content -Raw -ErrorAction SilentlyContinue ($Path + '.err'))" }
            if ($sw.Elapsed.TotalSeconds -gt $TimeoutSec) { throw "Timeout waiting for $Path" }
            Start-Sleep -Milliseconds 50
        }
    }

    function Get-Nodes { param($Roots) $st = [System.Collections.Generic.Stack[object]]::new(); foreach ($r in $Roots) { $st.Push($r) }; while ($st.Count) { $n = $st.Pop(); $n; foreach ($c in $n.Children) { $st.Push($c) } } }

    function Assert-AllExportedVerify {
        # Every Exported item of the store re-hashes to the generator's SHA-256 (seed = key, length = $LengthOf(key)).
        param($Store, [string]$MailboxId, [scriptblock]$LengthOf)
        $keys = [System.Collections.Generic.List[long]]::new()
        $after = 0L
        # Exported items = everything that is not pending: enumerate through the folder pages
        $folder = $Store.FindFolderKey($MailboxId, 'f0')
        $page = 0
        while ($true) {
            $p = $Store.GetItems($folder, $page, 500, [MailboxTransfer.StoreItemSort]::Size, $false, [MailboxTransfer.StoreItemState]::Exported)
            foreach ($i in $p.Items) { $keys.Add($i.Key) }
            $page++
            if ($page * 500 -ge $p.Total) { break }
        }
        $jobs = [System.Collections.Generic.List[MailboxTransfer.TestSupport.Job]]::new()
        foreach ($k in $keys) { $j = [MailboxTransfer.TestSupport.Job]::new(); $j.Key = $k; $j.Seed = $k; $j.Length = [long](& $LengthOf $k); $jobs.Add($j) }
        $r = [MailboxTransfer.TestSupport.Runner]::VerifyAll($Store, $jobs, 4)
        [pscustomobject]@{ Count = $keys.Count; Errors = $r.Errors }
    }
}

AfterAll {
    foreach ($p in $script:Children) { try { if (-not $p.HasExited) { $p.Kill($true) } } catch { } }
    Start-Sleep -Milliseconds 300
    if ($script:Base -and (Test-Path -LiteralPath $script:Base)) {
        for ($i = 0; $i -lt 5; $i++) {
            try { Remove-Item -LiteralPath $script:Base -Recurse -Force -ErrorAction Stop; break } catch { Start-Sleep -Milliseconds 500 }
        }
    }
    Write-Host ''
    Write-Host '=== Store test metrics ===' -ForegroundColor Cyan
    foreach ($k in $script:Metrics.Keys) { Write-Host ('  {0,-34} {1}' -f $k, $script:Metrics[$k]) }
    Write-Host ('  test data removed: {0}' -f (-not (Test-Path -LiteralPath $script:Base)))
}

Describe 'ExportStore' {

    Context 'Creation, reopening, schema' {
        BeforeAll {
            $script:S1Path = Join-Path $script:Base 'schema'
            $s = New-TestStore 'schema'
            Add-TestMailbox $s 'mb1' 'Primary'
            Add-TestMailbox $s 'mb2' 'Archive'
            $s.UpsertFolders(@(
                    [MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'root'; DisplayName = 'Top' },
                    [MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'inbox'; ParentFolderId = 'root'; DisplayName = 'Inbox'; FolderType = 'IPF.Note'; SourceItemCount = 12; SourceSize = 3456 },
                    [MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'sub'; ParentFolderId = 'inbox'; DisplayName = 'Sub'; FolderType = 'IPF.Note' }
                )) | Out-Null
            $script:S1Open = $s
        }
        AfterAll { try { $script:S1Open.Dispose() } catch { } }

        It 'creates catalog.db, packs\ and the lock file' {
            Test-Path (Join-Path $script:S1Path 'catalog.db') | Should -BeTrue
            Test-Path (Join-Path $script:S1Path 'packs') -PathType Container | Should -BeTrue
            Test-Path (Join-Path $script:S1Path 'export.lock') | Should -BeTrue
            $script:S1Open.Name | Should -Be 'schema'
            $script:S1Open.Source | Should -Be 'tests'
        }

        It 'refuses to create over an existing export and to open the same export twice for writing' {
            { [MailboxTransfer.ExportStore]::Create($script:S1Path, 'x', 'y') } | Should -Throw
            { [MailboxTransfer.ExportStore]::Open($script:S1Path) } | Should -Throw '*already open*'
        }

        It 'allows a read-only open while the export is open for writing, and refuses writes through it' {
            $ro = [MailboxTransfer.ExportStore]::OpenReadOnly($script:S1Path)
            try {
                $ro.GetMailboxes().Count | Should -Be 2
                { $ro.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = 'zz' }) } | Should -Throw '*read-only*'
            }
            finally { $ro.Dispose() }
        }

        It 'computes the full folder paths' {
            $f = $script:S1Open.GetFolders('mb1')
            ($f | Where-Object FolderId -EQ 'sub').Path | Should -Be 'Top/Inbox/Sub'
            ($f | Where-Object FolderId -EQ 'inbox').Path | Should -Be 'Top/Inbox'
            ($f | Where-Object FolderId -EQ 'root').Path | Should -Be 'Top'
        }

        It 'recomputes the paths of the descendants when a folder is renamed' {
            $script:S1Open.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'inbox'; ParentFolderId = 'root'; DisplayName = 'Boite' })) | Out-Null
            ($script:S1Open.GetFolders('mb1') | Where-Object FolderId -EQ 'sub').Path | Should -Be 'Top/Boite/Sub'
        }

        It 'rejects items for unknown mailboxes or folders' {
            { $script:S1Open.UpsertItems(@([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'nope'; ItemId = 'a'; FolderId = 'root' })) } | Should -Throw '*Unknown mailbox*'
            { $script:S1Open.UpsertItems(@([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = 'a'; FolderId = 'nope' })) } | Should -Throw '*Unknown folder*'
        }

        It 'reopens with the same content' {
            $script:S1Open.Dispose()
            $s = [MailboxTransfer.ExportStore]::Open($script:S1Path)
            try {
                $s.Name | Should -Be 'schema'
                $m = $s.GetMailboxes()
                $m.Count | Should -Be 2
                ($m | Where-Object Id -EQ 'mb2').Kind | Should -Be 'Archive'
                $f = $s.GetFolders('mb1')
                $f.Count | Should -Be 3
                ($f | Where-Object FolderId -EQ 'inbox').SourceItemCount | Should -BeNullOrEmpty   # replaced by the rename upsert (null)
                $s.LastRecovery.TruncatedPacks | Should -Be 0
                $s.LastRecovery.LostItems | Should -Be 0
            }
            finally { $s.Dispose() }
        }

        It 'has the expected tables, indexes, WAL and user_version' {
            $tables = (Invoke-Sql $script:S1Path "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'") | ForEach-Object { $_[0] }
            foreach ($t in 'export', 'mailbox', 'folder', 'pack', 'item', 'preview', 'importJob', 'importItem') { $tables | Should -Contain $t }
            $idx = (Invoke-Sql $script:S1Path "SELECT name FROM sqlite_master WHERE type='index'") | ForEach-Object { $_[0] }
            foreach ($i in 'ix_item_pending', 'ix_item_agg', 'ix_item_size', 'ix_item_date', 'ix_item_subject', 'ix_importitem_state') { $idx | Should -Contain $i }
            (Invoke-Sql $script:S1Path 'PRAGMA user_version')[0][0] | Should -Be 1
            (Invoke-Sql $script:S1Path 'PRAGMA journal_mode')[0][0] | Should -Be 'wal'
            (Invoke-Sql $script:S1Path 'SELECT schema_version, name FROM export')[0][0] | Should -Be 1
        }

        It 'refuses a catalog of a newer schema' {
            $p = Join-Path $script:Base 'newer'
            $s = New-TestStore 'newer'; $s.Dispose()
            $c = [Microsoft.Data.Sqlite.SqliteConnection]::new("Data Source=$(Join-Path $p 'catalog.db');Pooling=False"); $c.Open()
            $cmd = $c.CreateCommand(); $cmd.CommandText = 'PRAGMA user_version=99'; [void]$cmd.ExecuteNonQuery(); $c.Dispose()
            { [MailboxTransfer.ExportStore]::Open($p) } | Should -Throw '*schema version 99*'
        }
    }

    Context 'Volume: 20,000 items + large items, concurrent writers' {
        BeforeAll {
            $script:VolWriters = 6
            $opts = [MailboxTransfer.StoreOptions]::new()
            $opts.PackTargetSize = 256MB
            $script:Vol = New-TestStore 'volume' $opts
            Add-TestMailbox $script:Vol 'mb1'
            $fids = [MailboxTransfer.TestSupport.Workload]::FolderIds(40)
            $script:VolFolders = $fids
            [void]$script:Vol.UpsertFolders([MailboxTransfer.TestSupport.Workload]::MakeFolders('mb1', $fids))

            $big = [long[]]@((320MB + 1234), 60MB, 45MB, 30MB)
            $entries = $null; $sizes = $null
            [MailboxTransfer.TestSupport.Workload]::Build('mb1', $fids, 20000, 40, $big, 42, [ref]$entries, [ref]$sizes)
            $script:VolSizes = $sizes
            $script:VolTotal = ($sizes | Measure-Object -Sum).Sum
            $sw = [Diagnostics.Stopwatch]::StartNew()
            $script:VolUpserted = $script:Vol.UpsertItems($entries)
            $script:Metrics['UpsertItems 20k (ms)'] = $sw.ElapsedMilliseconds

            # jobs: the 4 big ones first (they run concurrently with the small ones), the rest shuffled
            $order = [MailboxTransfer.TestSupport.Workload]::ShuffledOrder(20000, 7) | Where-Object { $_ -ge 4 }
            $keys = @(1, 2, 3, 4) + @($order | ForEach-Object { $_ + 1 })
            $script:VolJobs = New-Jobs -Keys $keys -Sizes $sizes

            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
            $script:Mem = [MailboxTransfer.TestSupport.MemorySampler]::new()
            $script:Mem.Start(100)
            $script:Bg = [MailboxTransfer.TestSupport.BackgroundReader]::new($script:Vol, 'mb1')
            $script:Bg.Start()
            $script:WriteRun = [MailboxTransfer.TestSupport.Runner]::WriteAll($script:Vol, $script:VolJobs, $script:VolWriters, $false)
            $script:Bg.Stop()
            $script:VerifyRun = [MailboxTransfer.TestSupport.Runner]::VerifyAll($script:Vol, $script:VolJobs, 8)
            $script:Mem.Stop()
            $mb = 1MB
            $script:Metrics['Write 20k items'] = '{0:N0} MB in {1:N1} s = {2:N0} MB/s, {3:N0} items/s ({4} writers)' -f ($script:WriteRun.Bytes / $mb), $script:WriteRun.Seconds, $script:WriteRun.MBps, $script:WriteRun.ItemsPerSec, $script:VolWriters
            $script:Metrics['Read + SHA256 + generator'] = '{0:N0} MB in {1:N1} s = {2:N0} MB/s (8 readers, includes 2 hashes per byte)' -f ($script:VerifyRun.Bytes / $mb), $script:VerifyRun.Seconds, $script:VerifyRun.MBps
            $script:Metrics['Memory (process) baseline'] = '{0:N0} MB private' -f ($script:Mem.Baseline / $mb)
            $script:Metrics['Memory peak private / WS / managed'] = '{0:N0} / {1:N0} / {2:N0} MB' -f ($script:Mem.PeakPrivate / $mb), ($script:Mem.PeakWorkingSet / $mb), ($script:Mem.PeakManaged / $mb)
            $script:Metrics['Memory growth (peak - baseline)'] = '{0:N0} MB' -f (($script:Mem.PeakPrivate - $script:Mem.Baseline) / $mb)
        }
        AfterAll { try { $script:Vol.Dispose() } catch { } }

        It 'registered the 20,000 items (including one of at least 300 MiB)' {
            $script:VolUpserted | Should -Be 20000
            $script:VolSizes.Count | Should -Be 20000
            ($script:VolSizes | Measure-Object -Maximum).Maximum | Should -BeGreaterOrEqual 300MB
        }

        It 'wrote everything without error, with concurrent readers running' {
            $script:WriteRun.Errors | Select-Object -First 5 | Should -BeNullOrEmpty
            $script:WriteRun.Count | Should -Be 20000
            $script:Bg.Errors.Count | Should -Be 0
            $script:Bg.Loops | Should -BeGreaterThan 0
            $script:Metrics['Reader loops during the writes'] = $script:Bg.Loops
        }

        It 'reads back every item with the right length and SHA-256 (catalog and independent generator)' {
            $script:VerifyRun.Errors | Select-Object -First 5 | Should -BeNullOrEmpty
            $script:VerifyRun.Count | Should -Be 20000
            $script:VerifyRun.Bytes | Should -Be $script:VolTotal
        }

        It 'reports consistent statistics and no dead bytes in the packs' {
            $st = $script:Vol.GetStats()
            $st.Items | Should -Be 20000
            ($st.ItemsByState | Where-Object State -EQ 'Exported').Count | Should -Be 20000
            ($st.ItemsByState | Where-Object State -EQ 'Pending').Count | Should -Be 0
            $st.StoredBytes | Should -Be $script:VolTotal
            $st.PackBytesOnDisk | Should -Be $script:VolTotal
            $st.Packs | Should -BeGreaterThan 1
            $script:Metrics['Packs / catalog'] = '{0} packs, {1:N0} MB data, catalog.db {2:N1} MB' -f $st.Packs, ($st.PackBytesOnDisk / 1MB), ((Get-Item (Join-Path $script:Vol.Root 'catalog.db')).Length / 1MB)
        }

        It 'uses one pack per concurrent writer at most (+ rotations), never one file per item' {
            $files = Get-ChildItem (Join-Path $script:Vol.Root 'packs') -File
            $files.Count | Should -BeLessThan 40
            (Get-ChildItem $script:Vol.Root -Recurse -File).Count | Should -BeLessThan 50
        }

        It 'kept the process memory low' {
            $growth = $script:Mem.PeakPrivate - $script:Mem.Baseline
            $growth | Should -BeLessThan 300MB -Because 'streams must never be loaded in memory (a 320 MiB item was written)'
        }

        It 'measures the raw read throughput of the packs (sequential per item, 8 readers)' {
            $keys = [long[]](1..20000)
            $r = [MailboxTransfer.TestSupport.Runner]::ReadAllRaw($script:Vol, $keys, 8)
            $r.Errors.Count | Should -Be 0
            $r.Bytes | Should -Be $script:VolTotal
            $script:Metrics['Raw read (OS cache likely hot)'] = '{0:N0} MB in {1:N1} s = {2:N0} MB/s' -f ($r.Bytes / 1MB), $r.Seconds, $r.MBps
        }

        It 'survives close and reopen: same stats, sample of items verifies, VerifyItem agrees' {
            $before = $script:Vol.GetStats()
            $script:Vol.Dispose()
            $s = [MailboxTransfer.ExportStore]::Open((Join-Path $script:Base 'volume'))
            try {
                $s.LastRecovery.TruncatedPacks | Should -Be 0
                $s.LastRecovery.LostItems | Should -Be 0
                $after = $s.GetStats()
                $after.Items | Should -Be $before.Items
                $after.StoredBytes | Should -Be $before.StoredBytes
                $rnd = [Random]::new(5)
                for ($n = 0; $n -lt 100; $n++) {
                    $k = $rnd.Next(1, 20001)
                    $v = $s.VerifyItem($k)
                    $v.Ok | Should -BeTrue -Because "item $k : $($v.Reason)"
                }
                ($s.VerifyItem(1)).Ok | Should -BeTrue    # the 320 MiB item
                $sw = [Diagnostics.Stopwatch]::StartNew()
                $bad = $s.VerifyAll([Threading.CancellationToken]::None)
                $script:Metrics['VerifyAll (sequential, 1 thread)'] = '{0:N1} s for {1:N0} MB = {2:N0} MB/s' -f $sw.Elapsed.TotalSeconds, ($after.StoredBytes / 1MB), (($after.StoredBytes / 1MB) / $sw.Elapsed.TotalSeconds)
                $bad.Count | Should -Be 0
            }
            finally { $s.Dispose() }
        }
    }

    Context 'Async writes' {
        It 'WriteItemAsync stores items with the same guarantees' {
            $s = New-TestStore 'async'
            try {
                Add-TestMailbox $s
                $s.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'f0'; DisplayName = 'Root' })) | Out-Null
                $entries = [System.Collections.Generic.List[MailboxTransfer.StoreItemEntry]]::new()
                for ($i = 0; $i -lt 200; $i++) { $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = "a$i"; FolderId = 'f0'; Size = 100000 + $i }) }
                $s.UpsertItems($entries) | Out-Null
                $jobs = New-Jobs -Keys (1..200) -Sizes ([long[]](100000..100199))
                $r = [MailboxTransfer.TestSupport.Runner]::WriteAll($s, $jobs, 4, $true)
                $r.Errors.Count | Should -Be 0
                $r.Count | Should -Be 200
                $v = [MailboxTransfer.TestSupport.Runner]::VerifyAll($s, $jobs, 4)
                $v.Errors.Count | Should -Be 0
                { $s.WriteItem(1, [MailboxTransfer.TestSupport.GenStream]::new(1, 10), $script:NoCancel) } | Should -Throw '*already exported*'
                $s.WriteItem(1, [MailboxTransfer.TestSupport.GenStream]::new(1, 100000), $script:NoCancel, $true).Length | Should -Be 100000
                $s.VerifyItem(1).Ok | Should -BeTrue
            }
            finally { $s.Dispose() }
        }
    }

    Context 'Failures during a write (process alive)' {
        It 'rolls the pack back on a source error, on cancellation and on an unknown item; the other items stay intact' {
            $s = New-CrashStore 'fail' 20
            try {
                $good = [MailboxTransfer.TestSupport.GenStream]::new(1, 3MB)
                $s.WriteItem(1, $good, $script:NoCancel).Length | Should -Be 3MB
                $packPath = $s.GetPackPath(1)
                $len0 = (Get-Item $packPath).Length

                { $s.WriteItem(2, [MailboxTransfer.TestSupport.FailingStream]::new(2, 10MB, 4MB), $script:NoCancel) } | Should -Throw '*simulated source failure*'
                (Get-Item $packPath).Length | Should -Be $len0 -Because 'the partial bytes are truncated immediately'
                $s.GetItem(2).State | Should -Be 'Pending'

                $cts = [Threading.CancellationTokenSource]::new(); $cts.Cancel()
                { $s.WriteItem(3, [MailboxTransfer.TestSupport.GenStream]::new(3, 5MB), $cts.Token) } | Should -Throw
                (Get-Item $packPath).Length | Should -Be $len0
                $s.GetItem(3).State | Should -Be 'Pending'

                { $s.WriteItem(99999, [MailboxTransfer.TestSupport.GenStream]::new(1, 10), $script:NoCancel) } | Should -Throw '*Unknown item key*'
                (Get-Item $packPath).Length | Should -Be $len0

                # the pack is reused and the next item lands exactly where the failed one started
                $r = $s.WriteItem(2, [MailboxTransfer.TestSupport.GenStream]::new(2, 2MB), $script:NoCancel)
                $r.PackId | Should -Be 1
                $r.Offset | Should -Be $len0
                $s.VerifyItem(1).Ok | Should -BeTrue
                $s.VerifyItem(2).Ok | Should -BeTrue

                $s.MarkItemFailed(4, 'timeout while exporting')
                $s.MarkItemSkipped(5, 'not exportable')
                $i4 = $s.GetItem(4); $i4.State | Should -Be 'Failed'; $i4.Attempts | Should -Be 1; $i4.LastError | Should -Be 'timeout while exporting'
                $s.GetItem(5).State | Should -Be 'Skipped'
                { $s.MarkItemFailed(1, 'x') } | Should -Throw '*already exported*'
                $s.RequeueFailed('mb1', 3) | Should -Be 1
                $s.GetItem(4).State | Should -Be 'Pending'
            }
            finally { $s.Dispose() }
        }

        It 'puts an exported item back to Pending when its changeKey changes upstream' {
            $s = New-CrashStore 'changekey' 3
            try {
                $s.WriteItem(1, [MailboxTransfer.TestSupport.GenStream]::new(1, 1000), $script:NoCancel) | Out-Null
                $same = [MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = ([MailboxTransfer.TestSupport.Workload]::ItemId(0)); FolderId = 'f0'; ChangeKey = 'c0'; ItemType = 'IPM.Note'; Size = 1000 }
                $s.UpsertItems(@($same)) | Out-Null
                $s.GetItem(1).State | Should -Be 'Exported'
                $same.ChangeKey = 'c0-modified'
                $s.UpsertItems(@($same)) | Out-Null
                $i = $s.GetItem(1); $i.State | Should -Be 'Pending'; $i.PackId | Should -BeNullOrEmpty; $i.ChangeKey | Should -Be 'c0-modified'
            }
            finally { $s.Dispose() }
        }
    }

    Context 'Crash recovery' {
        It 'in-process abandon in the middle of a write: pack truncated, item Pending again, nothing else corrupted' {
            $s = New-CrashStore 'abandon' 40
            $path = $s.Root
            $stall = $null; $t = $null
            try {
                for ($k = 1; $k -le 10; $k++) {
                    $s.WriteItem($k, [MailboxTransfer.TestSupport.GenStream]::new($k, [MailboxTransfer.TestSupport.Workload]::CrashSize($k)), $script:NoCancel) | Out-Null
                }
                $committedEnd = (Get-CommittedEnd $path)[1]
                $stall = [MailboxTransfer.TestSupport.StallStream]::new(11, 40MB, 12MB, $null)
                $t = [MailboxTransfer.TestSupport.Runner]::StartWrite($s, 11, $stall)
                $stall.Reached.Wait(30000) | Should -BeTrue
                (Get-Item $s.GetPackPath(1)).Length | Should -BeGreaterThan ($committedEnd + 10MB) -Because 'the partial item is on disk'
                $s.Abandon()
                $stall.Release.Set()
                [void]$t.Wait(30000)
                $onDisk = (Get-Item $s.GetPackPath(1)).Length     # what the "dead" process left behind
            }
            finally { try { $s.Dispose() } catch { } }

            $r = [MailboxTransfer.ExportStore]::Open($path)
            try {
                $r.LastRecovery.TruncatedPacks | Should -Be 1
                $r.LastRecovery.TruncatedBytes | Should -Be ($onDisk - $committedEnd)
                (Get-Item $r.GetPackPath(1)).Length | Should -Be $committedEnd
                $r.GetItem(11).State | Should -Be 'Pending'
                $r.GetPendingItems('mb1', 100, 0).Count | Should -Be 30
                $chk = Assert-AllExportedVerify $r 'mb1' { param($k) [MailboxTransfer.TestSupport.Workload]::CrashSize($k) }
                $chk.Count | Should -Be 10
                $chk.Errors.Count | Should -Be 0
                # and the export can be completed afterwards
                $jobs = [System.Collections.Generic.List[MailboxTransfer.TestSupport.Job]]::new()
                foreach ($k in 11..40) { $j = [MailboxTransfer.TestSupport.Job]::new(); $j.Key = $k; $j.Seed = $k; $j.Length = [MailboxTransfer.TestSupport.Workload]::CrashSize($k); $jobs.Add($j) }
                [MailboxTransfer.TestSupport.Runner]::WriteAll($r, $jobs, 3, $false).Errors.Count | Should -Be 0
                $chk = Assert-AllExportedVerify $r 'mb1' { param($k) [MailboxTransfer.TestSupport.Workload]::CrashSize($k) }
                $chk.Count | Should -Be 40
                $chk.Errors.Count | Should -Be 0
                $script:Metrics['Abandon recovery'] = 'truncated {0:N1} MB, item 11 Pending, 40/40 verify after completion' -f ($r.LastRecovery.TruncatedBytes / 1MB)
            }
            finally { $r.Dispose() }
        }

        It 'garbage appended to a pack after the last commit is truncated at reopen' {
            $s = New-CrashStore 'garbage' 5
            $path = $s.Root
            $s.WriteItem(1, [MailboxTransfer.TestSupport.GenStream]::new(1, [MailboxTransfer.TestSupport.Workload]::CrashSize(1)), $script:NoCancel) | Out-Null
            $s.Dispose()
            $end = (Get-CommittedEnd $path)[1]
            $pack = Join-Path $path 'packs\pack-000001.bin'
            $fs = [IO.File]::Open($pack, 'Append', 'Write'); $fs.Write([byte[]]::new(1234567), 0, 1234567); $fs.Dispose()
            (Get-Item $pack).Length | Should -Be ($end + 1234567)
            $r = [MailboxTransfer.ExportStore]::Open($path)
            try {
                $r.LastRecovery.TruncatedBytes | Should -Be 1234567
                (Get-Item $pack).Length | Should -Be $end
                $r.VerifyItem(1).Ok | Should -BeTrue
            }
            finally { $r.Dispose() }
        }

        It 'a pack shorter than the catalog (external damage): the missing items go back to Pending, the others are fine' {
            $s = New-CrashStore 'damaged' 10
            $path = $s.Root
            foreach ($k in 1..10) { $s.WriteItem($k, [MailboxTransfer.TestSupport.GenStream]::new($k, 1MB + $k * 1000), $script:NoCancel) | Out-Null }
            $s.Dispose()
            $pack = Join-Path $path 'packs\pack-000001.bin'
            $full = (Get-Item $pack).Length
            $fs = [IO.File]::Open($pack, 'Open', 'Write'); $fs.SetLength($full - 2MB - 5000); $fs.Dispose()    # cuts items 9 and 10 (and part of 8 or not)
            $r = [MailboxTransfer.ExportStore]::Open($path)
            try {
                $r.LastRecovery.DamagedPacks | Should -Be 1
                $r.LastRecovery.LostItems | Should -BeGreaterOrEqual 2
                $pending = $r.GetPendingItems('mb1', 100, 0)
                $pending.Count | Should -Be $r.LastRecovery.LostItems
                foreach ($i in $pending) { $i.State | Should -Be 'Pending' }
                foreach ($k in 1..(10 - $pending.Count)) { $r.VerifyItem($k).Ok | Should -BeTrue }
                (Get-Item $pack).Length | Should -Be ((Get-CommittedEnd $path)[1])
            }
            finally { $r.Dispose() }
        }

        It 'a killed process (Stop-Process) in the middle of an item write: pack truncated, item Pending, the rest intact' {
            $s = New-CrashStore 'kill-stall' 30
            $path = $s.Root
            $s.Dispose()
            $sig = Join-Path $script:Base 'kill-stall.signal'
            $p = Start-CrashChild -ExportPath $path -Mode Stall -Signal $sig
            Wait-ForFile $sig $p
            $p.HasExited | Should -BeFalse
            $committed = Get-CommittedEnd $path       # read-only open of the catalog while the child holds the export
            Stop-Process -Id $p.Id -Force
            $p.WaitForExit(30000) | Should -BeTrue
            $onDisk = (Get-Item (Join-Path $path 'packs\pack-000001.bin')).Length
            $onDisk | Should -BeGreaterThan ($committed[1] + 25MB) -Because 'the partial item (>= 30 MiB) is on disk when the process is killed'

            $r = [MailboxTransfer.ExportStore]::Open($path)
            try {
                $r.LastRecovery.TruncatedPacks | Should -Be 1
                $r.LastRecovery.TruncatedBytes | Should -Be ($onDisk - $committed[1])
                (Get-Item $r.GetPackPath(1)).Length | Should -Be $committed[1]
                $r.GetPendingItems('mb1', 100, 0).Count | Should -Be 25
                $r.GetItem(6).State | Should -Be 'Pending' -Because 'item 6 was the one being written'
                $chk = Assert-AllExportedVerify $r 'mb1' { param($k) [MailboxTransfer.TestSupport.Workload]::CrashSize($k) }
                $chk.Count | Should -Be 5
                $chk.Errors.Count | Should -Be 0
                $script:Metrics['Kill (stalled write) recovery'] = 'truncated {0:N1} MB, item 6 Pending, 5/5 verify' -f ($r.LastRecovery.TruncatedBytes / 1MB)
            }
            finally { $r.Dispose() }
        }

        It 'a process killed under load (4 writers) twice in a row: every Exported item verifies, every pack ends on an item boundary, and the export completes' {
            $s = New-CrashStore 'kill-load' 800
            $path = $s.Root
            $s.Dispose()
            $total = 800
            $lengthOf = { param($k) [MailboxTransfer.TestSupport.Workload]::CrashSize($k) }
            $truncated = 0L
            foreach ($round in 1, 2) {
                $sig = Join-Path $script:Base "kill-load-$round.signal"
                $p = Start-CrashChild -ExportPath $path -Mode Load -Signal $sig -Threads 4
                Wait-ForFile $sig $p
                Start-Sleep -Milliseconds (1200 + 400 * $round)
                if ($p.HasExited) { Set-ItResult -Skipped -Because 'child finished before being killed'; return }
                Stop-Process -Id $p.Id -Force
                $p.WaitForExit(30000) | Should -BeTrue
                $r = [MailboxTransfer.ExportStore]::Open($path)
                try {
                    $truncated += $r.LastRecovery.TruncatedBytes
                    $r.LastRecovery.DamagedPacks | Should -Be 0
                    $r.LastRecovery.LostItems | Should -Be 0
                    $st = $r.GetStats()
                    $exported = ($st.ItemsByState | Where-Object State -EQ 'Exported').Count
                    $exported | Should -BeGreaterThan 0
                    $exported | Should -BeLessThan $total
                    # every pack ends exactly at the end of its last committed item
                    $ends = Get-CommittedEnd $path
                    foreach ($f in Get-ChildItem (Join-Path $path 'packs') -File) {
                        $id = [int]($f.BaseName -replace '\D', '')
                        $expected = if ($ends.ContainsKey($id)) { $ends[$id] } else { 0 }
                        $f.Length | Should -Be $expected -Because "pack $id must end on an item boundary"
                    }
                    $chk = Assert-AllExportedVerify $r 'mb1' $lengthOf
                    $chk.Count | Should -Be $exported
                    $chk.Errors | Select-Object -First 3 | Should -BeNullOrEmpty
                    $script:Metrics["Kill under load, round $round"] = '{0} items committed, {1:N1} MB of partial writes truncated, all verify' -f $exported, ($r.LastRecovery.TruncatedBytes / 1MB)
                }
                finally { $r.Dispose() }
            }
            # finish the export with the survivors
            $r = [MailboxTransfer.ExportStore]::Open($path)
            try {
                $pend = $r.GetPendingItems('mb1', 1000, 0)
                $jobs = [System.Collections.Generic.List[MailboxTransfer.TestSupport.Job]]::new()
                foreach ($i in $pend) { $j = [MailboxTransfer.TestSupport.Job]::new(); $j.Key = $i.Key; $j.Seed = $i.Key; $j.Length = [MailboxTransfer.TestSupport.Workload]::CrashSize($i.Key); $jobs.Add($j) }
                [MailboxTransfer.TestSupport.Runner]::WriteAll($r, $jobs, 4, $false).Errors.Count | Should -Be 0
                $chk = Assert-AllExportedVerify $r 'mb1' $lengthOf
                $chk.Count | Should -Be $total
                $chk.Errors.Count | Should -Be 0
                $r.GetStats().PackBytesOnDisk | Should -Be $r.GetStats().StoredBytes -Because 'no dead bytes left after the recoveries'
            }
            finally { $r.Dispose() }
        }
    }

    Context 'Pack rotation' {
        It 'rotates at the (reduced) target size, reuses the open packs after a reopen' {
            $opts = [MailboxTransfer.StoreOptions]::new(); $opts.PackTargetSize = 2MB
            $s = New-TestStore 'rotation' $opts
            $path = $s.Root
            try {
                Add-TestMailbox $s
                $s.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'f0'; DisplayName = 'Root' })) | Out-Null
                $entries = [System.Collections.Generic.List[MailboxTransfer.StoreItemEntry]]::new()
                $item = 700KB
                for ($i = 0; $i -lt 70; $i++) { $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = "r$i"; FolderId = 'f0'; Size = $item }) }
                $s.UpsertItems($entries) | Out-Null
                $jobs = New-Jobs -Keys (1..60) -Sizes ([long[]](@([long]$item) * 60))
                [MailboxTransfer.TestSupport.Runner]::WriteAll($s, $jobs, 3, $false).Errors.Count | Should -Be 0
                # How the 60 items spread over the 3 writers is not deterministic (60 items can close every pack: 3 x 700 KB >= 2 MB).
                # A small item written alone guarantees that at least one pack stays open (unsealed) for the reopen below:
                # an open pack holds at most 2 items (1.4 MB) and 100 KB more stays under the 2 MB target.
                $small = 100KB
                $jobsSmall = New-Jobs -Keys @(61) -Sizes ([long[]](@([long]$item) * 60 + @([long]$small)))
                [MailboxTransfer.TestSupport.Runner]::WriteAll($s, $jobsSmall, 1, $false).Errors.Count | Should -Be 0
                $sizes = Get-ChildItem (Join-Path $path 'packs') -File | ForEach-Object Length
                $sizes.Count | Should -BeGreaterOrEqual 14
                ($sizes | Measure-Object -Maximum).Maximum | Should -BeLessOrEqual (2MB + $item) -Because 'a pack closes as soon as it reaches the target (an item is never split)'
                $sealed = (Invoke-Sql $path 'SELECT COUNT(*) FROM pack WHERE sealed=1')[0][0]
                $sealed | Should -BeGreaterOrEqual ($sizes.Count - 3) -Because 'at most one open (unsealed) pack per writer'
                $packsBefore = $s.GetStats().Packs
                $openBefore = @((Invoke-Sql $path 'SELECT pack_id FROM pack WHERE sealed=0') | ForEach-Object { [int]$_[0] })
                $openBefore.Count | Should -BeGreaterOrEqual 1
                $s.Dispose()

                $s = [MailboxTransfer.ExportStore]::Open($path, $opts)
                $sizes65 = [long[]](@([long]$item) * 60 + @([long]$small) + @([long]$item) * 4)
                $jobs2 = New-Jobs -Keys (62..65) -Sizes $sizes65
                [MailboxTransfer.TestSupport.Runner]::WriteAll($s, $jobs2, 1, $false).Errors.Count | Should -Be 0
                $s.GetItem(62).PackId | Should -BeIn $openBefore -Because 'an unsealed pack of the previous session is reused (the first item goes to one of them)'
                $s.GetStats().Packs | Should -BeLessOrEqual ($packsBefore + 2) -Because '4 items of 700 KB need at most 2 packs beyond the open ones'
                $all = New-Jobs -Keys (1..65) -Sizes $sizes65
                [MailboxTransfer.TestSupport.Runner]::VerifyAll($s, $all, 4).Errors.Count | Should -Be 0
                $script:Metrics['Rotation (target 2 MB)'] = '{0} packs for {1} items of {2} KB' -f $s.GetStats().Packs, 65, ($item / 1KB)
            }
            finally { $s.Dispose() }
        }
    }

    Context 'Folder tree, aggregates, pagination, previews' {
        BeforeAll {
            $script:T = New-TestStore 'tree'
            Add-TestMailbox $script:T 'mb1' 'Primary'
            Add-TestMailbox $script:T 'mb2' 'Archive'
            $script:Fids = [MailboxTransfer.TestSupport.Workload]::FolderIds(13)   # F0 root; F1..F4 children of F0; F5..F8 of F1; F9..F12 of F2
            [void]$script:T.UpsertFolders([MailboxTransfer.TestSupport.Workload]::MakeFolders('mb1', $script:Fids))
            [void]$script:T.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = 'mb2'; FolderId = 'arch-root'; DisplayName = 'Archive root' }))
            $rnd = [Random]::new(11)
            $script:Exp = @{}    # folderId -> list of [pscustomobject] key,size,state,subject,date
            $entries = [System.Collections.Generic.List[MailboxTransfer.StoreItemEntry]]::new()
            $meta = [System.Collections.Generic.List[object]]::new()
            $n = 900
            for ($i = 0; $i -lt $n; $i++) {
                $fid = $script:Fids[$rnd.Next($script:Fids.Count)]
                $size = 500 + $rnd.Next(60000)
                $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = "t$i"; FolderId = $fid; ChangeKey = "k$i"; ItemType = 'IPM.Note'; Size = $size })
                $meta.Add([pscustomobject]@{ Key = [long]($i + 1); Folder = $fid; Size = [long]$size; State = 'Pending' })
            }
            for ($i = 0; $i -lt 40; $i++) { $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb2'; ItemId = "arch$i"; FolderId = 'arch-root'; Size = 1000 }) }
            [void]$script:T.UpsertItems($entries)
            $script:Meta = $meta
            # export 300 items, fail 20, skip 10
            $jobs = [System.Collections.Generic.List[MailboxTransfer.TestSupport.Job]]::new()
            for ($i = 0; $i -lt 300; $i++) { $j = [MailboxTransfer.TestSupport.Job]::new(); $j.Key = $meta[$i].Key; $j.Seed = $j.Key; $j.Length = $meta[$i].Size; $jobs.Add($j); $meta[$i].State = 'Exported' }
            [void][MailboxTransfer.TestSupport.Runner]::WriteAll($script:T, $jobs, 4, $false)
            for ($i = 300; $i -lt 320; $i++) { $script:T.MarkItemFailed($meta[$i].Key, 'boom'); $meta[$i].State = 'Failed' }
            for ($i = 320; $i -lt 330; $i++) { $script:T.MarkItemSkipped($meta[$i].Key, 'skip'); $meta[$i].State = 'Skipped' }
            # previews with ties on the dates (hours granularity), mixed-case ASCII subjects
            $base = [datetime]::new(2024, 1, 1, 0, 0, 0, [DateTimeKind]::Utc)
            $words = 'alpha', 'Bravo', 'charlie', 'Delta', 'echo', 'Foxtrot', 'golf', 'Hotel'
            $prev = [System.Collections.Generic.List[MailboxTransfer.StorePreview]]::new()
            foreach ($m in $meta) {
                $m | Add-Member Subject ("{0} {1} #{2}" -f $words[$rnd.Next(8)], $words[$rnd.Next(8)], $rnd.Next(30)) -Force
                $m | Add-Member Date ($base.AddHours($rnd.Next(0, 200))) -Force
                $prev.Add([MailboxTransfer.StorePreview]@{ ItemKey = $m.Key; Subject = $m.Subject; From = 'a@b.test'; To = 'c@d.test'; ReceivedUtc = $m.Date; SentUtc = $m.Date.AddMinutes(-1); MessageClass = 'IPM.Note'; HasAttachments = $false; AttachmentCount = 0; BodyPreview = 'x' })
            }
            [void]$script:T.SetPreviews($prev)
            $script:TreeExport = $script:T.Root
        }
        AfterAll { try { $script:T.Dispose() } catch { } }


        It 'builds the tree with the right shape and paths' {
            $roots = $script:T.GetFolderTree('mb1')
            $roots.Count | Should -Be 1
            $roots[0].DisplayName | Should -Be 'Root'
            $roots[0].Children.Count | Should -Be 4
            $nodes = @(Get-Nodes $roots)
            $nodes.Count | Should -Be 13
            ($nodes | Where-Object FolderId -EQ 'folder-005').Path | Should -Be 'Root/Folder 1/Folder 5'
            ($nodes | Where-Object FolderId -EQ 'folder-005').Depth | Should -Be 2
            $script:T.GetFolderTree('mb2').Count | Should -Be 1
        }

        It 'aggregates counts, announced sizes, states and stored bytes per folder (own and recursive)' {
            $roots = $script:T.GetFolderTree('mb1')
            $nodes = @(Get-Nodes $roots)
            foreach ($f in $script:Fids) {
                $mine = @($script:Meta | Where-Object Folder -EQ $f)
                $node = $nodes | Where-Object FolderId -EQ $f
                $node.Own.Count | Should -Be $mine.Count
                $node.Own.Size | Should -Be (($mine | Measure-Object Size -Sum).Sum ?? 0)
                $node.Own.Exported | Should -Be @($mine | Where-Object State -EQ 'Exported').Count
                $node.Own.Pending | Should -Be @($mine | Where-Object State -EQ 'Pending').Count
                $node.Own.Failed | Should -Be @($mine | Where-Object State -EQ 'Failed').Count
                $node.Own.Skipped | Should -Be @($mine | Where-Object State -EQ 'Skipped').Count
                $node.Own.StoredBytes | Should -Be (($mine | Where-Object State -EQ 'Exported' | Measure-Object Size -Sum).Sum ?? 0)
            }
            $root = $roots[0]
            $root.Total.Count | Should -Be $script:Meta.Count
            $root.Total.Size | Should -Be (($script:Meta | Measure-Object Size -Sum).Sum)
            $root.Total.Exported | Should -Be 300
            $root.Total.Failed | Should -Be 20
            $root.Total.Skipped | Should -Be 10
            $root.Total.Pending | Should -Be ($script:Meta.Count - 330)
            $f1 = $roots[0].Children | Where-Object FolderId -EQ 'folder-001'
            $sub = @('folder-001', 'folder-005', 'folder-006', 'folder-007', 'folder-008')
            $f1.Total.Count | Should -Be @($script:Meta | Where-Object { $_.Folder -in $sub }).Count
        }

        It 'keeps the mailboxes (primary / archive) separate' {
            $a = $script:T.GetFolderTree('mb2')[0]
            $a.Own.Count | Should -Be 40
            $a.Own.Pending | Should -Be 40
        }

        It 'paginates a folder sorted by date, size and subject, ascending and descending, with stable pages and a correct total' {
            $folderId = ($script:Meta | Group-Object Folder | Sort-Object Count -Descending | Select-Object -First 1).Name
            $fk = $script:T.FindFolderKey('mb1', $folderId)
            $mine = @($script:Meta | Where-Object Folder -EQ $folderId)
            $pageSize = 17
            foreach ($sort in 'ReceivedDate', 'Size', 'Subject') {
                foreach ($desc in $false, $true) {
                    $got = [System.Collections.Generic.List[long]]::new()
                    $p = 0
                    while ($true) {
                        $page = $script:T.GetItems($fk, $p, $pageSize, [MailboxTransfer.StoreItemSort]$sort, $desc)
                        $page.Total | Should -Be $mine.Count
                        $page.Items.Count | Should -BeLessOrEqual $pageSize
                        foreach ($i in $page.Items) { $got.Add($i.Key) }
                        $p++
                        if ($p -ge $page.PageCount) { break }
                    }
                    $got.Count | Should -Be $mine.Count
                    @($got | Select-Object -Unique).Count | Should -Be $mine.Count -Because "no duplicates across pages ($sort desc=$desc)"
                    # Expected order, computed independently: key, then item key as tie-break. SQLite NOCASE folds ASCII only
                    # (the subjects are ASCII) and compares bytes: ordinal comparison of the lower-cased text.
                    $list = [System.Collections.Generic.List[object]]::new($mine)
                    $cmpBlock = {
                        param($a, $b)
                        $c = switch ($sort) {
                            'ReceivedDate' { $a.Date.CompareTo($b.Date) }
                            'Size' { $a.Size.CompareTo($b.Size) }
                            'Subject' { [string]::CompareOrdinal($a.Subject.ToLowerInvariant(), $b.Subject.ToLowerInvariant()) }
                        }
                        if ($c -eq 0) { $c = $a.Key.CompareTo($b.Key) }
                        if ($desc) { -$c } else { $c }
                    }.GetNewClosure()
                    $list.Sort([Comparison[object]]$cmpBlock)
                    ($got -join ',') | Should -Be (($list | ForEach-Object Key) -join ',') -Because "order for $sort desc=$desc"
                }
            }
        }

        It 'filters a page by state' {
            $fk = $script:T.FindFolderKey('mb1', 'folder-000')
            $p = $script:T.GetItems($fk, 0, 1000, [MailboxTransfer.StoreItemSort]::Size, $true, [MailboxTransfer.StoreItemState]::Exported)
            $p.Total | Should -Be @($script:Meta | Where-Object { $_.Folder -eq 'folder-000' -and $_.State -eq 'Exported' }).Count
            foreach ($i in $p.Items) { $i.State | Should -Be 'Exported'; $i.Length | Should -Be $i.Size; $i.Sha256.Length | Should -Be 64 }
        }

        It 'uses the indexes for the hot queries (no table scan, no sort of a whole folder)' {
            $fk = $script:T.FindFolderKey('mb1', 'folder-000')
            $plan = { param($sql) ((Invoke-Sql $script:TreeExport "EXPLAIN QUERY PLAN $sql") | ForEach-Object { $_[3] }) -join ' | ' }
            (& $plan "SELECT item_key FROM item WHERE mailbox_key=1 AND state=0 AND item_key>0 ORDER BY item_key LIMIT 100") | Should -Match 'ix_item_pending'
            (& $plan "SELECT item_key FROM item WHERE folder_key=$fk ORDER BY received_utc, item_key LIMIT 50") | Should -Match 'ix_item_date'
            (& $plan "SELECT item_key FROM item WHERE folder_key=$fk ORDER BY size DESC, item_key DESC LIMIT 50") | Should -Match 'ix_item_size'
            (& $plan "SELECT item_key FROM item WHERE folder_key=$fk ORDER BY subject COLLATE NOCASE, item_key LIMIT 50") | Should -Match 'ix_item_subject'
            (& $plan "SELECT folder_key, state, COUNT(*), SUM(size), SUM(pack_length) FROM item WHERE mailbox_key=1 AND folder_key IS NOT NULL GROUP BY folder_key, state") | Should -Match 'COVERING INDEX ix_item_agg'
            (& $plan "SELECT item_key FROM item WHERE folder_key=$fk ORDER BY received_utc, item_key LIMIT 50") | Should -Not -Match 'TEMP B-TREE'
        }

        It 'stores and reads back previews (1 KiB body limit, UTF-8 safe, replacement)' {
            $key = $script:Meta[0].Key
            $body = ('é€' * 2000) + ' fin'
            $p = [MailboxTransfer.StorePreview]@{ Subject = 'Sujet accentué'; From = 'Alice <alice@contoso.test>'; To = 'bob@contoso.test; carol@contoso.test'; ReceivedUtc = [datetime]::new(2025, 3, 4, 5, 6, 7, [DateTimeKind]::Utc); SentUtc = [datetime]::new(2025, 3, 4, 5, 0, 0, [DateTimeKind]::Utc); MessageClass = 'IPM.Note'; HasAttachments = $true; AttachmentCount = 2; AttachmentNames = "a.pdf`nb.xlsx"; BodyPreview = $body }
            $script:T.SetPreview($key, $p)
            $g = $script:T.GetPreview($key)
            $g.Subject | Should -Be 'Sujet accentué'
            $g.From | Should -Be 'Alice <alice@contoso.test>'
            $g.To | Should -Be 'bob@contoso.test; carol@contoso.test'
            $g.ReceivedUtc | Should -Be ([datetime]::new(2025, 3, 4, 5, 6, 7, [DateTimeKind]::Utc))
            $g.SentUtc.Kind | Should -Be 'Utc'
            $g.HasAttachments | Should -BeTrue
            $g.AttachmentCount | Should -Be 2
            $g.AttachmentNames | Should -Be "a.pdf`nb.xlsx"
            [Text.Encoding]::UTF8.GetByteCount($g.BodyPreview) | Should -BeLessOrEqual 1024
            [Text.Encoding]::UTF8.GetByteCount($g.BodyPreview) | Should -BeGreaterThan 1000
            $g.BodyPreview | Should -Not -Match ([string][char]0xFFFD)
            $script:T.GetItem($key).Subject | Should -Be 'Sujet accentué'    # denormalised sort column follows
            $script:T.GetPreview(999999) | Should -BeNullOrEmpty
        }
    }

    Context 'Import cycle' {
        It 'InFlight items become Unknown after a crash and are never replayed automatically' {
            $s = New-CrashStore 'import' 40
            $path = $s.Root
            try {
                $jobsW = [System.Collections.Generic.List[MailboxTransfer.TestSupport.Job]]::new()
                foreach ($k in 1..30) { $j = [MailboxTransfer.TestSupport.Job]::new(); $j.Key = $k; $j.Seed = $k; $j.Length = 10000 + $k; $jobsW.Add($j) }
                [MailboxTransfer.TestSupport.Runner]::WriteAll($s, $jobsW, 3, $false).Errors.Count | Should -Be 0

                $job = $s.CreateImportJob('restore to target', 'target@contoso.test')
                $plan = [System.Collections.Generic.List[MailboxTransfer.StoreImportPlan]]::new()
                foreach ($k in 1..40) { $plan.Add([MailboxTransfer.StoreImportPlan]@{ ItemKey = $k; TargetFolder = "Inbox/$($k % 3)" }) }
                $s.AddImportItems($job, $plan) | Should -Be 30 -Because 'only Exported items can be imported (31..40 are Pending)'
                $s.AddImportItems($job, $plan) | Should -Be 0 -Because 'duplicates are ignored'
                $s.GetImportJob($job).TargetMailbox | Should -Be 'target@contoso.test'

                $a = $s.NextImportItems($job, 10)
                $a.Count | Should -Be 10
                foreach ($i in $a) { $i.State | Should -Be 'InFlight'; $i.Attempts | Should -Be 1; $i.TargetMailbox | Should -Be 'target@contoso.test' }
                $a[0].TargetFolder | Should -Be "Inbox/$($a[0].ItemKey % 3)"
                $a[0].Length | Should -Be (10000 + $a[0].ItemKey)
                # the stream of an item to import can be opened from the info
                $st = $s.OpenItem($a[0].ItemKey); try { $st.Length | Should -Be $a[0].Length } finally { $st.Dispose() }
                foreach ($i in $a[0..3]) { $s.MarkImportDone($i.ImportKey, "target-id-$($i.ItemKey)") }
                foreach ($i in $a[4..5]) { $s.MarkImportFailed($i.ImportKey, 'throttled') }
                # $a[6..9] stay InFlight
                $b = $s.NextImportItems($job, 10)
                $b.Count | Should -Be 10
                (@($a.ImportKey) + @($b.ImportKey) | Select-Object -Unique).Count | Should -Be 20 -Because 'an item is handed out once'

                $stats = $s.GetImportStats($job)
                ($stats | Where-Object State -EQ 'Done').Count | Should -Be 4
                ($stats | Where-Object State -EQ 'Failed').Count | Should -Be 2
                ($stats | Where-Object State -EQ 'InFlight').Count | Should -Be 14
                ($stats | Where-Object State -EQ 'Pending').Count | Should -Be 10

                $s.Abandon()    # crash
            }
            finally { try { $s.Dispose() } catch { } }

            $r = [MailboxTransfer.ExportStore]::Open($path)
            try {
                $r.LastRecovery.ImportInFlightToUnknown | Should -Be 14
                $stats = $r.GetImportStats($job)
                ($stats | Where-Object State -EQ 'Unknown').Count | Should -Be 14
                ($stats | Where-Object State -EQ 'InFlight').Count | Should -Be 0
                ($stats | Where-Object State -EQ 'Done').Count | Should -Be 4
                ($stats | Where-Object State -EQ 'Failed').Count | Should -Be 2
                ($r.GetStats().ImportByState | Where-Object State -EQ 'Unknown').Count | Should -Be 14

                $next = $r.NextImportItems($job, 100)
                $next.Count | Should -Be 10 -Because 'only the Pending ones: Unknown and Failed are not replayed automatically'
                $unknown = $r.GetImportItems($job, [MailboxTransfer.StoreImportState]::Unknown, 100)
                $unknown.Count | Should -Be 14
                @($next.ImportKey | Where-Object { $_ -in $unknown.ImportKey }).Count | Should -Be 0
                # the operator checks the target and resolves them
                $r.ResolveImportUnknown($unknown[0].ImportKey, $true, 'found-on-target')
                $r.ResolveImportUnknown($unknown[1].ImportKey, $false, $null)
                { $r.ResolveImportUnknown($unknown[0].ImportKey, $true, 'x') } | Should -Throw '*not Unknown*'
                $r.GetImportItems($job, [MailboxTransfer.StoreImportState]::Done, 100).Count | Should -Be 5
                $again = $r.NextImportItems($job, 100)
                $again.Count | Should -Be 1
                $again[0].ImportKey | Should -Be $unknown[1].ImportKey
                $again[0].Attempts | Should -Be 2 -Because 'handed out in the first session, again after the resolution: attempts counted'
                # failed items are retried on request only
                $r.RequeueImportFailed($job, 5) | Should -Be 2
                $r.NextImportItems($job, 100).Count | Should -Be 2
                $r.GetImportItems($job, [MailboxTransfer.StoreImportState]::Done, 100)[0].TargetItemId | Should -Not -BeNullOrEmpty
                $script:Metrics['Import crash'] = '14 InFlight -> Unknown, not replayed; 10 Pending resumed'
            }
            finally { $r.Dispose() }
        }
    }
}
