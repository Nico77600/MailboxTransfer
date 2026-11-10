#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - automated tests (Pester 5 or later).
    Author  : Nicolas Fabert
    Version : 1.0.0

    Run:  Invoke-Pester -Path .\tests\MailboxTransfer.Tests.ps1 -Output Detailed

    Nothing here reaches the network or a mailbox: the Graph client talks to a fake HTTP handler (tests\Fixtures.cs) and
    the FastTransfer reader reads streams written by FxBuilder from the layout of MS-OXCFXICS, plus the byte samples
    published by Microsoft (Learn: importItem; Open Specifications blog: how to parse the FastTransfer stream).
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
    # The engine (SQLite + C#) is compiled once per process; the test helpers of Store.Tests.ps1 must be in the same unit
    # (an in-memory assembly cannot be referenced later), so this file goes through the shared loader too.
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-tests-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp

    function New-TestClient {
        param([Parameter(Mandatory)]$Handler, [int]$MaxAttempts = 6)
        $o = [MailboxTransfer.GraphOptions]::new()
        $o.DelayScale = 0
        $o.UseJitter = $false
        $o.MaxAttempts = $MaxAttempts
        return [MailboxTransfer.GraphClient]::new([Func[string]] { 'test-token' }, $o, $Handler)
    }

    function Get-Sha256Hex([byte[]]$Bytes) {
        # Should -Be on a large byte array takes about a minute per MB: big payloads are compared by hash.
        return [BitConverter]::ToString([Security.Cryptography.SHA256]::HashData($Bytes)).Replace('-', '').ToLowerInvariant()
    }

    function ConvertFrom-HexText([string]$Text) {
        return [byte[]]@(($Text -split '\s+' | Where-Object { $_ }) | ForEach-Object { [Convert]::ToByte($_, 16) })
    }

    function Read-Fx([byte[]]$Bytes) {
        $read = [MailboxTransfer.FxReader]::Read([IO.MemoryStream]::new($Bytes))
        return [pscustomobject]@{ Read = $read; Summary = [MailboxTransfer.FxMessage]::Summarize($read) }
    }

    # The structure observed on a real export of the lab (synthetic values): properties in code page strings, an extended property
    # (type | 0x4000 and two bytes), a named property by name, the body last, FXDelProp before the recipients and the attachments,
    # then one attachment whose data (the large part) is in the middle of its properties.
    function Write-RealShapeMessage {
        param([Parameter(Mandatory)][IO.Stream]$Stream, [long]$AttachBytes = 1MB, [string]$AttachName = 'lot1-00001-1.bin')
        $cp = [Text.Encoding]::Unicode
        $t = [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)
        $b = [MailboxTransfer.Tests.FxBuilder]::new($Stream)
        $null = $b.Int32(0x0017, 1).CodePageStr(0x001A, 0x84B0, 'IPM.Note', $cp).Bool(0x0023, $false).Int32(0x0036, 0).CodePageStr(0x0037, 0x84B0, '[lot1-00001] Export - Alpha', $cp)
        $null = $b.Time(0x0039, $t).Bin(0x003B, [byte[]](1..40)).CodePageStr(0x0C1A, 0x84B0, 'Sender Name', $cp).CodePageStr(0x5D01, 0x84B0, 'sender@lab.example', $cp)
        $null = $b.Time(0x0E06, $t).Int32(0x0E07, 1).Time(0x3007, $t).Bin(0x65E1, [byte[]](1..22)).Int32(0x3FFD, 65001)
        $null = $b.NamedNameStr(0x8EB2, [guid]'23239608-685d-4732-9c55-4c95cb4e8e33', 'FocusedInboxSubjectEmbeddings', ('[0.5,' * 100) + ']')
        $null = $b.ExtBin(0x010B, [byte[]](0..255 * 6)).ExtBin(0x010C, [byte[]](0..255 * 29)).ExtTime(0x0FE3, $t).ExtBool(0x1705, $false).ExtInt16(0x1714, 2).ExtTime(0x3309, $t)
        $null = $b.CodePageStr(0x1000, 0x84B0, "Bonjour,`r`nCi-joint le document.", $cp)
        $null = $b.DelProp(0x0E12000D)
        $null = $b.Marker(0x40030003).Int32(0x3000, 0).Int32(0x0C15, 1).Bin(0x0FF9, [byte[]](1..157)).CodePageStr(0x3001, 0x84B0, 'Alice', $cp).CodePageStr(0x39FE, 0x84B0, 'alice@lab.example', $cp).Marker(0x40040003)
        $null = $b.DelProp(0x0E13000D)
        $null = $b.Marker(0x40000003).Int32(0x0E21, 0).Int32(0x0E20, [int]($AttachBytes + 177)).Bin(0x0FF9, [byte[]](1..16)).CodePageStr(0x3001, 0x84B0, $AttachName, $cp)
        $null = $b.BinLarge(0x3701, $AttachBytes)
        $null = $b.CodePageStr(0x3704, 0x84B0, $AttachName, $cp).Int32(0x370B, -1).Int32(0x3714, 0).CodePageStr(0x3A0C, 0x84B0, 'EnUs', $cp).Bool(0x7FFE, $false).Marker(0x400E0003)
    }

    $script:Mbx = 'MBX:11111111-2222-3333-4444-555555555555@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:Base = 'https://graph.microsoft.com/v1.0'

    # A complete message written the way MS-OXCFXICS describes it.
    function New-TestMessage {
        param([switch]$WithFolderInfo)
        $b = [MailboxTransfer.Tests.FxBuilder]::new()
        $named = [guid]'00020329-0000-0000-c000-000000000046'
        $null = $b.Int32(0x0017, 1).Str(0x001A, 'IPM.Note').Str(0x0037, 'Quarterly report').Str8(0x0C1A, 'Adele Vance')
        $null = $b.Str(0x5D01, 'adele@contoso.test').Time(0x0039, [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)).Time(0x3007, [datetime]::new(2026, 10, 21, 9, 31, 0, [DateTimeKind]::Utc))
        $null = $b.Int32(0x0E08, 71133).Bool(0x0E1B, $true).Double(0x7005, 1.5).Int16(0x7001, 7)
        $null = $b.Bin(0x0FF9, [byte[]](1..16)).GuidValue(0x7002, [guid]'11111111-2222-3333-4444-555555555555')
        $null = $b.MvInt32(0x7003, @(10, 20, 30)).MvStr(0x7004, @('red', 'blue'))
        $null = $b.NamedLidInt32(0x8005, $named, 0x8503, 5).NamedNameStr(0x8006, $named, 'Keywords', 'lab')
        $null = $b.Str(0x1000, ('hello ' * 100000))                                   # 1.2 million characters of body
        $null = $b.Bin(0x1013, [Text.Encoding]::UTF8.GetBytes('<html><body>Hi</body></html>')).Bin(0x1009, [byte[]](1..50))
        if ($WithFolderInfo) { $null = $b.Int64(0x6749, -863846703525003263).Int64(0x6748, 12345) }
        # recipients: To then Cc
        $null = $b.Marker(0x40030003).Int32(0x0C15, 1).Str(0x3001, 'Alice').Str(0x3003, '/o=ex/cn=alice').Str(0x39FE, 'alice@contoso.test').Marker(0x40040003)
        $null = $b.Marker(0x40030003).Int32(0x0C15, 2).Str(0x3001, 'Bob').Str(0x39FE, 'bob@contoso.test').Marker(0x40040003)
        # attachment 1: a file
        $null = $b.Marker(0x40000003).Str(0x3707, 'report.pdf').Int32(0x0E20, 2048).Int32(0x3705, 1).Str(0x370E, 'application/pdf').Bin(0x3701, [byte[]](0..255 * 8)).Marker(0x400E0003)
        # attachment 2: an embedded message
        $null = $b.Marker(0x40000003).Str(0x3001, 'forwarded.msg').Int32(0x3705, 5).Object(0x3701).Marker(0x40010003).Str(0x0037, 'Inner subject').Marker(0x40020003).Marker(0x400E0003)
        return $b.ToArray()
    }
}

AfterAll {
    if ($script:Tmp -and (Test-Path -LiteralPath $script:Tmp)) { Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

Describe 'Package' {
    It 'every PowerShell file parses' {
        $files = Get-ChildItem -LiteralPath $script:RepoRoot -Recurse -Include *.ps1, *.psm1, *.psd1 -File | Where-Object FullName -NotMatch '\\(\.git|out)\\'
        foreach ($file in $files) {
            $errors = $null
            [void][Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$null, [ref]$errors)
            $errors | Should -BeNullOrEmpty -Because $file.FullName
        }
    }
    It 'the manifest is valid and exports the functions of the module' {
        $manifest = Test-ModuleManifest -Path (Join-Path $script:Root 'MailboxTransfer.psd1')
        $manifest.ExportedFunctions.Keys | Should -Contain 'Export-MbtItems'
        (Get-Command -Module MailboxTransfer).Count | Should -Be $manifest.ExportedFunctions.Count
    }
    It 'the engine is compiled' {
        'MailboxTransfer.GraphClient' -as [type] | Should -Not -BeNullOrEmpty
        'MailboxTransfer.FxReader' -as [type] | Should -Not -BeNullOrEmpty
    }
    It 'reads the configuration over the defaults and refuses unknown keys' {
        $cfg = Import-MbtConfiguration
        $cfg.Graph.ExportBatchSize | Should -Be 20
        $cfg.Auth.CacheFile | Should -BeLike '*MailboxTransfer*token.cache'
        $bad = Join-Path $script:Tmp 'bad.psd1'
        Set-Content -LiteralPath $bad -Value "@{ Graph = @{ Nonsense = 1 } }"
        { Import-MbtConfiguration -Path $bad } | Should -Throw '*unknown key*'
        Set-Content -LiteralPath $bad -Value "@{ Graph = @{ ExportBatchSize = 21 } }"
        { Import-MbtConfiguration -Path $bad } | Should -Throw '*between 1 and 20*'
    }
    It 'Format-MbtBytes / Format-MbtRate' {
        Format-MbtBytes 512 | Should -Be '512 B'
        Format-MbtBytes (20MB) | Should -Match '^20[.,]0 MiB$'
        Format-MbtRate -Bytes 10MB -Seconds 2 | Should -Match '^5[.,]00 MiB/s$'
    }
}

Describe 'Engine and local store' {
    It 'Initialize-MbtEngine loads SQLite and compiles the store with the rest of the engine' {
        ('MailboxTransfer.ExportStore' -as [type]) | Should -Not -BeNullOrEmpty
        ('MailboxTransfer.GraphClient' -as [type]) | Should -Not -BeNullOrEmpty
        { Initialize-MbtEngine } | Should -Not -Throw -Because 'it is a no-op once the engine is compiled'
    }
    It 'creates a store, writes an item stream in a pack, reads it back and verifies it' {
        $path = Join-Path $script:Tmp 'store-smoke'
        $store = [MailboxTransfer.ExportStore]::Create($path, 'smoke', 'tests')
        try {
            $store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:Mbx; Kind = 'Primary'; Upn = 'adele@contoso.test'; DisplayName = 'Adele Vance' }) | Out-Null
            $store.UpsertFolders(@([MailboxTransfer.StoreFolder]@{ MailboxId = $script:Mbx; FolderId = 'F1'; DisplayName = 'Lab' })) | Out-Null
            $store.UpsertItems(@([MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:Mbx; ItemId = 'I1'; FolderId = 'F1'; ItemType = 'IPM.Note'; Size = 10 })) | Out-Null
            $bytes = [byte[]](1..200)
            $r = $store.WriteItem($script:Mbx, 'I1', [IO.MemoryStream]::new($bytes), [Threading.CancellationToken]::None)
            $r.Length | Should -Be 200
            $in = $store.OpenItem($script:Mbx, 'I1')
            try { $back = [byte[]]::new(200); [void]$in.Read($back, 0, 200) } finally { $in.Dispose() }
            $back | Should -Be $bytes
            $store.VerifyAll([Threading.CancellationToken]::None).Ok | Should -Not -Contain $false
            Get-ChildItem (Join-Path $path 'packs') -File | Should -HaveCount 1
        } finally { $store.Dispose() }
    }
}

Describe 'Graph client: pagination' {
    It 'follows @odata.nextLink as returned, page after page' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $next1 = "$($script:Base)/admin/exchange/mailboxes/MBX:x@y/folders?%24skip=10"
        $next2 = "$($script:Base)/admin/exchange/mailboxes/MBX:x@y/folders?%24skip=20&%24top=10"
        $h.Enqueue(200, ('{"value":[{"id":"a"},{"id":"b"}],"@odata.nextLink":"' + $next1 + '"}'))
        $h.Enqueue(200, ('{"value":[{"id":"c"}],"@odata.nextLink":"' + $next2 + '"}'))
        $h.Enqueue(200, '{"value":[{"id":"d"}]}')
        $client = New-TestClient $h
        $ids = @($client.EnumerateValues("$($script:Base)/admin/exchange/mailboxes/MBX:x@y/folders") | ForEach-Object { $_.GetProperty('id').GetString() })
        $ids | Should -Be @('a', 'b', 'c', 'd')
        $h.Requests.Count | Should -Be 3
        $h.Requests[1].Url | Should -Be $next1
        $h.Requests[2].Url | Should -Be $next2
        $h.Requests.HasAuthorization | Should -Not -Contain $false
        $client.PagesRead | Should -Be 3
    }
    It 'accepts an empty collection' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, '{"value":[]}')
        @((New-TestClient $h).EnumerateValues("$($script:Base)/x")).Count | Should -Be 0
    }
    It 'refuses a nextLink that is not https on graph.microsoft.com' -ForEach @(
        @{ Link = 'http://graph.microsoft.com/v1.0/x?%24skip=10' }
        @{ Link = 'https://evil.example.com/v1.0/x?%24skip=10' }
        @{ Link = 'https://graph.microsoft.com.evil.example/v1.0/x' }
    ) {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, ('{"value":[{"id":"a"}],"@odata.nextLink":"' + $Link + '"}'))
        { @((New-TestClient $h).EnumerateValues("$($script:Base)/x")) } | Should -Throw '*nextLink*'
        $h.Requests.Count | Should -Be 1
    }
}

Describe 'Graph client: retry' {
    It 'honours Retry-After on 429 and retries' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(429, '{"error":{"code":"TooManyRequests","message":"slow down"}}', '7')
        $h.Enqueue(429, '{}', '3')
        $h.Enqueue(200, '{"value":[]}')
        $client = New-TestClient $h
        $doc = $client.GetJson("$($script:Base)/x")
        $doc.Dispose()
        $h.Requests.Count | Should -Be 3
        $client.RetryCount | Should -Be 2
        $client.Delays[0].TotalSeconds | Should -Be 7
        $client.Delays[1].TotalSeconds | Should -Be 3
    }
    It 'backs off exponentially on 503 / 504 without Retry-After' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(503); $h.Enqueue(504); $h.Enqueue(503); $h.Enqueue(200, '{}')
        $client = New-TestClient $h
        ($client.GetJson("$($script:Base)/x")).Dispose()
        ($client.Delays | ForEach-Object TotalSeconds) | Should -Be @(2, 4, 8)
    }
    It 'gives up after MaxAttempts and reports the Graph error' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        1..3 | ForEach-Object { $h.Enqueue(503, '{"error":{"code":"ServiceUnavailable","message":"busy"}}') }
        $client = New-TestClient $h -MaxAttempts 3
        try { $client.GetJson("$($script:Base)/x"); throw 'not reached' } catch { $ge = Get-MbtGraphException $_ }
        $ge.StatusCode | Should -Be 503
        $ge.ErrorCode | Should -Be 'ServiceUnavailable'
        $ge.Message | Should -BeLike '*busy*'
        $h.Requests.Count | Should -Be 3
    }
    It 'does not retry a client error and parses the error body' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(404, '{"error":{"code":"ErrorItemNotFound","message":"nope"}}')
        $client = New-TestClient $h
        try { $client.GetJson("$($script:Base)/x"); throw 'not reached' } catch { $ge = Get-MbtGraphException $_ }
        $ge.StatusCode | Should -Be 404
        $ge.ErrorCode | Should -Be 'ErrorItemNotFound'
        $h.Requests.Count | Should -Be 1
    }
    It 'retries a network failure of an idempotent request' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.EnqueueFailure(); $h.Enqueue(200, '{}')
        ($client = New-TestClient $h).GetJson("$($script:Base)/x").Dispose()
        $h.Requests.Count | Should -Be 2
    }
    It 'follows a 308 redirect to graph.microsoft.com and refuses another host' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $moved = "$($script:Base)/admin/exchange/mailboxes/MBX:other@y/folders/F"
        $h.Enqueue(308, '', $null, $moved); $h.Enqueue(200, '{"id":"F"}')
        ($client = New-TestClient $h).GetJson("$($script:Base)/admin/exchange/mailboxes/MBX:x@y/folders/F").Dispose()
        $h.Requests[1].Url | Should -Be $moved
        $h2 = [MailboxTransfer.Tests.FakeHandler]::new()
        $h2.Enqueue(308, '', $null, 'https://evil.example.com/x')
        { (New-TestClient $h2).GetJson("$($script:Base)/x") } | Should -Throw '*not on graph.microsoft.com*'
    }
    It 'never retries a 504 on the import (the item may exist) but does retry a 429' {
        $file = Join-Path $script:Tmp 'imp.fxs'
        [IO.File]::WriteAllBytes($file, [byte[]](1..100))
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(504, '{}')
        $client = New-TestClient $h
        $api = [MailboxTransfer.MailboxApi]::new($client)
        $session = [MailboxTransfer.ImportSession]@{ MailboxId = $script:Mbx; ImportUrl = "https://outlook.office365.com/api/gv1.0/Mailboxes('x')/importItem?authtoken=SECRET123" }
        try { $api.ImportItem($session, 'F1', $file); throw 'not reached' } catch { $ge = Get-MbtGraphException $_ }
        $ge.StatusCode | Should -Be 504
        $ge.Message | Should -Not -BeLike '*SECRET123*'
        $h.Requests.Count | Should -Be 1
        $h.Enqueue(429, '{}', '1'); $h.Enqueue(200, '{"itemId":"I1","changeKey":"C1"}')
        $result = $api.ImportItem($session, 'F1', $file)
        $result.ItemId | Should -Be 'I1'
        $h.Requests.Count | Should -Be 3
    }
    It 'sends the import body again when the upload is interrupted before its last byte (nothing can have been created)' {
        $file = Join-Path $script:Tmp 'imp-cut.fxs'
        [IO.File]::WriteAllBytes($file, [byte[]](1..100))
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.EnqueueInterruptedUpload(60); $h.EnqueueInterruptedUpload(100)
        $h.Enqueue(200, '{"itemId":"I2","changeKey":"C2"}')
        $client = New-TestClient $h
        $script:Notices = [System.Collections.Generic.List[string]]::new()
        $client.Notice = [Action[string]] { param($m) $script:Notices.Add($m) }
        $session = [MailboxTransfer.ImportSession]@{ MailboxId = $script:Mbx; ImportUrl = "https://outlook.office365.com/api/gv1.0/Mailboxes('x')/importItem?authtoken=SECRET123" }
        $result = ([MailboxTransfer.MailboxApi]::new($client)).ImportItem($session, 'F1', $file)
        $result.ItemId | Should -Be 'I2'
        $h.Requests.Count | Should -Be 3
        $h.Requests | ForEach-Object { $_.HasAuthorization | Should -BeFalse }
        $client.RetryCount | Should -Be 2
        $script:Notices.Count | Should -Be 2
        $script:Notices[0] | Should -BeLike 'upload interrupted after 41 of 179 bytes, nothing created, retrying*'
        ($script:Notices -join ' ') | Should -Not -BeLike '*SECRET123*'
    }
    It 'gives up as "nothing created" (UploadInterrupted) when every attempt is interrupted, and as network error when the whole body was sent' {
        $file = Join-Path $script:Tmp 'imp-cut2.fxs'
        [IO.File]::WriteAllBytes($file, [byte[]](1..100))
        $session = [MailboxTransfer.ImportSession]@{ MailboxId = $script:Mbx; ImportUrl = "https://outlook.office365.com/api/gv1.0/Mailboxes('x')/importItem?authtoken=SECRET123" }
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        1..3 | ForEach-Object { $h.EnqueueInterruptedUpload(60) }
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h -MaxAttempts 3))
        try { $api.ImportItem($session, 'F1', $file); throw 'not reached' } catch { $ge = Get-MbtGraphException $_ }
        $ge.ErrorCode | Should -Be 'UploadInterrupted'
        $ge.Message | Should -BeLike '*nothing created*3 attempts*'
        $ge.Message | Should -Not -BeLike '*SECRET123*'
        $h.Requests.Count | Should -Be 3
        [MailboxTransfer.ImportHelper]::NothingCreated($ge) | Should -BeTrue

        $h2 = [MailboxTransfer.Tests.FakeHandler]::new()
        $h2.EnqueueFailure('The response ended prematurely.')
        $api2 = [MailboxTransfer.MailboxApi]::new((New-TestClient $h2))
        try { $api2.ImportItem($session, 'F1', $file); throw 'not reached' } catch { $ge2 = Get-MbtGraphException $_ }
        $ge2.ErrorCode | Should -Be 'NetworkError'
        $h2.Requests.Count | Should -Be 1 -Because 'the whole body was sent: the item may exist, so it is not sent again'
        [MailboxTransfer.ImportHelper]::NothingCreated($ge2) | Should -BeFalse
    }
    It 'sends the import again when not a single byte of the body reached the transport (connect / TLS / DNS failure: nothing can have been created)' {
        $file = Join-Path $script:Tmp 'imp-notsent.fxs'
        [IO.File]::WriteAllBytes($file, [byte[]](1..100))
        $session = [MailboxTransfer.ImportSession]@{ MailboxId = $script:Mbx; ImportUrl = "https://outlook.office365.com/api/gv1.0/Mailboxes('x')/importItem?authtoken=SECRET123" }
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.EnqueueNotSent(); $h.EnqueueNotSent('A connection attempt failed.')
        $h.Enqueue(200, '{"itemId":"I3","changeKey":"C3"}')
        $client = New-TestClient $h
        $script:Notices = [System.Collections.Generic.List[string]]::new()
        $client.Notice = [Action[string]] { param($m) $script:Notices.Add($m) }
        $result = ([MailboxTransfer.MailboxApi]::new($client)).ImportItem($session, 'F1', $file)
        $result.ItemId | Should -Be 'I3'
        $h.Requests.Count | Should -Be 3
        $client.RetryCount | Should -Be 2
        $script:Notices[0] | Should -BeLike 'request not sent (no byte of the body reached the transport), nothing created, retrying*'
        ($script:Notices -join ' ') | Should -Not -BeLike '*SECRET123*'

        $h2 = [MailboxTransfer.Tests.FakeHandler]::new()
        1..3 | ForEach-Object { $h2.EnqueueNotSent() }
        $api2 = [MailboxTransfer.MailboxApi]::new((New-TestClient $h2 -MaxAttempts 3))
        try { $api2.ImportItem($session, 'F1', $file); throw 'not reached' } catch { $ge = Get-MbtGraphException $_ }
        $ge.ErrorCode | Should -Be 'NotSent'
        $ge.Message | Should -BeLike '*not sent*nothing created*3 attempts*'
        $ge.Message | Should -Not -BeLike '*SECRET123*'
        $h2.Requests.Count | Should -Be 3
        [MailboxTransfer.ImportHelper]::NothingCreated($ge) | Should -BeTrue
    }
    It 'does not treat as "nothing created" a failure after the body was handed over, nor any HTTP status received' {
        $file = Join-Path $script:Tmp 'imp-after.fxs'
        [IO.File]::WriteAllBytes($file, [byte[]](1..100))
        $session = [MailboxTransfer.ImportSession]@{ MailboxId = $script:Mbx; ImportUrl = "https://outlook.office365.com/api/gv1.0/Mailboxes('x')/importItem?authtoken=SECRET123" }
        # one attempt not sent, then the whole body is sent and the answer is lost: Unknown, and the body is not sent a third time
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.EnqueueNotSent(); $h.EnqueueFailure('The response ended prematurely.')
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        try { $api.ImportItem($session, 'F1', $file); throw 'not reached' } catch { $ge = Get-MbtGraphException $_ }
        $ge.ErrorCode | Should -Be 'NetworkError'
        $h.Requests.Count | Should -Be 2
        [MailboxTransfer.ImportHelper]::NothingCreated($ge) | Should -BeFalse
        foreach ($status in 500, 502, 504) {
            $h2 = [MailboxTransfer.Tests.FakeHandler]::new()
            $h2.Enqueue($status, '{"error":{"code":"x","message":"y"}}')
            $api2 = [MailboxTransfer.MailboxApi]::new((New-TestClient $h2))
            try { $api2.ImportItem($session, 'F1', $file); throw 'not reached' } catch { $ge2 = Get-MbtGraphException $_ }
            $ge2.StatusCode | Should -Be $status
            $h2.Requests.Count | Should -Be 1
            [MailboxTransfer.ImportHelper]::NothingCreated($ge2) | Should -BeFalse -Because "a $status was received after the body was sent"
        }
    }
}

Describe 'Mailbox API: folders' {
    It 'completes a first-level listing with childFolders until the tree is complete' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $mb = "$($script:Base)/admin/exchange/mailboxes/$($script:Mbx)"
        $h.Enqueue(200, '{"value":[
            {"id":"R","displayName":"Top","parentFolderId":"ROOT","childFolderCount":2,"totalItemCount":0,"type":"IPF.Note"},
            {"id":"A","displayName":"Alpha","parentFolderId":"R","childFolderCount":1,"totalItemCount":3,"type":"IPF.Note"},
            {"id":"B","displayName":"Beta","parentFolderId":"R","childFolderCount":0,"totalItemCount":0,"type":"IPF.Note"}]}')
        $h.Enqueue(200, '{"value":[{"id":"C","displayName":"Gamma","parentFolderId":"A","childFolderCount":0,"totalItemCount":7,"type":"IPF.Note"}]}')
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        $tree = @($api.GetFolderTree($script:Mbx))
        ($tree | ForEach-Object Path) | Should -Be @('Top', 'Top/Alpha', 'Top/Alpha/Gamma', 'Top/Beta')
        $h.Requests.Count | Should -Be 2
        $h.Requests[0].Url | Should -Be "$mb/folders"
        $h.Requests[1].Url | Should -Be "$mb/folders/A/childFolders"
        ($tree | Where-Object Id -EQ 'C').Depth | Should -Be 2
    }
    It 'does not ask again when the listing already holds every folder' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, '{"value":[
            {"id":"R","displayName":"Top","childFolderCount":1,"totalItemCount":0},
            {"id":"A","displayName":"Alpha","parentFolderId":"R","childFolderCount":0,"totalItemCount":3}]}')
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        @($api.GetFolderTree($script:Mbx)).Count | Should -Be 2
        $h.Requests.Count | Should -Be 1
    }
    It 'resolves a folder by id, path, or end of path, and refuses ambiguity' {
        $mk = { param($id, $name, $path) [pscustomobject]@{ Id = $id; DisplayName = $name; Path = $path } }
        $tree = @((& $mk 'i1' 'Lab' 'Top/Lab'), (& $mk 'i2' 'Attachments' 'Top/Lab/Attachments'), (& $mk 'i3' 'Attachments' 'Top/Other/Attachments'), (& $mk 'i4' 'Top' 'Top'))
        (Resolve-MbtFolder -Tree $tree -Name 'i2').Path | Should -Be 'Top/Lab/Attachments'
        (Resolve-MbtFolder -Tree $tree -Name 'Top/Lab').Id | Should -Be 'i1'
        (Resolve-MbtFolder -Tree $tree -Name 'Lab/Attachments').Id | Should -Be 'i2'
        { Resolve-MbtFolder -Tree $tree -Name 'Attachments' } | Should -Throw '*ambiguous*'
        { Resolve-MbtFolder -Tree $tree -Name 'Missing' } | Should -Throw '*not found*'
    }
    It 'lists items with size and type, requesting $top only when asked' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, '{"value":[{"id":"i1","type":"IPM.Note","size":71133,"changeKey":"ck","createdDateTime":"2026-10-21T09:00:00Z"}]}')
        $h.Enqueue(200, '{"value":[]}')
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        $items = @($api.EnumerateItems($script:Mbx, 'F=1'))
        $items[0].Size | Should -Be 71133
        $items[0].Created.Year | Should -Be 2026
        $h.Requests[0].Url | Should -BeLike '*/folders/F%3D1/items'
        @($api.EnumerateItems($script:Mbx, 'F', 100)).Count | Should -Be 0
        $h.Requests[1].Url | Should -BeLike '*/items?$top=100'
    }
    It 'does not trust a mailbox id: checks syntax, tenant, Graph lookup (denied or not) and reachability' {
        $tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(403, '{"error":{"code":"Authorization_RequestDenied","message":"no"}}')       # users/{upn}/settings/exchange
        $h.Enqueue(200, '{"value":[{"id":"R","displayName":"Top"}]}')                            # folders?$top=1
        $null = Set-MbtSession -Graph (New-TestClient $h)
        $r = Test-MbtMailboxId -MailboxId $script:Mbx -TenantId $tenant -UserPrincipalName 'adele@contoso.test'
        $r.Syntax | Should -BeTrue
        $r.TenantMatches | Should -BeTrue
        $r.GraphLookup | Should -Be 'HTTP 403 Authorization_RequestDenied'
        $r.Reachable | Should -BeTrue
        $h.Requests[1].Url | Should -BeLike '*/admin/exchange/mailboxes/MBX:11111111-*/folders?$top=1'
        $h2 = [MailboxTransfer.Tests.FakeHandler]::new()
        $h2.Enqueue(200, ('{"primaryMailboxId":"' + $script:Mbx + '"}'))
        $h2.Enqueue(404, '{"error":{"code":"ErrorMailboxNotFound","message":"x"}}')
        $null = Set-MbtSession -Graph (New-TestClient $h2)
        $r2 = Test-MbtMailboxId -MailboxId $script:Mbx -TenantId 'another-tenant' -UserPrincipalName 'adele@contoso.test'
        $r2.IsPrimaryPerGraph | Should -BeTrue
        $r2.TenantMatches | Should -BeFalse
        $r2.Reachable | Should -BeFalse
        $r2.Error | Should -BeLike '*ErrorMailboxNotFound*'
        $bad = Test-MbtMailboxId -MailboxId 'not-an-id' -TenantId $tenant
        $bad.Syntax | Should -BeFalse
        Disconnect-MbtGraph
    }
    It 'measures folder sizes by adding up the item sizes' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, '{"value":[{"id":"A","displayName":"A","childFolderCount":0,"totalItemCount":2},{"id":"B","displayName":"B","childFolderCount":0,"totalItemCount":0}]}')
        $h.Enqueue(200, ('{"value":[{"id":"i1","size":1000}],"@odata.nextLink":"' + $script:Base + '/next"}'))
        $h.Enqueue(200, '{"value":[{"id":"i2","size":2500}]}')
        $null = Set-MbtSession -Graph (New-TestClient $h)
        $tree = Get-MbtMailboxTree -MailboxId $script:Mbx -Sizes
        ($tree | Where-Object Id -EQ 'A').TotalSize | Should -Be 3500
        ($tree | Where-Object Id -EQ 'B').TotalSize | Should -Be 0
        $h.Requests.Count | Should -Be 3
        Disconnect-MbtGraph
    }
    It 'splits item ids in groups of 20' {
        $groups = [MailboxTransfer.MailboxApi]::Chunk([string[]](1..45))
        ($groups | ForEach-Object Count) | Should -Be @(20, 20, 5)
        [MailboxTransfer.MailboxApi]::Chunk([string[]]@()).Count | Should -Be 0
        ([MailboxTransfer.MailboxApi]::Chunk([string[]](1..40)) | ForEach-Object Count) | Should -Be @(20, 20)
    }
}

Describe 'Export: streaming of the exportItems answer' {
    BeforeAll {
        function ConvertTo-EscapedBase64([byte[]]$Bytes) {
            # what a JSON serializer may write: "/" as "\/" and "+" as "\u002B"
            return [Convert]::ToBase64String($Bytes).Replace('/', '\/').Replace('+', '\u002B')
        }
        $script:Rng = [Random]::new(42)
        $script:DataA = [byte[]]::new(300000); $script:Rng.NextBytes($script:DataA)
        $script:DataB = [byte[]]::new(1000); $script:Rng.NextBytes($script:DataB)
        $script:DataC = [byte[]]::new(2); $script:DataC[0] = 0xFB; $script:DataC[1] = 0xEF
    }
    It 'writes each item to a file, with JSON escapes, any property order, errors and null data' {
        $dir = Join-Path $script:Tmp 'exp1'
        $json = '{"@odata.context":"https://graph.microsoft.com/v1.0/$metadata#Collection(microsoft.graph.exportItemResponse)","value":[' +
            '{"data":"' + (ConvertTo-EscapedBase64 $script:DataA) + '","itemId":"ITEM-A/1+x","changeKey":"CK-A"},' +
            '{"itemId":"ITEM-B","changeKey":"CK-B","data":"' + (ConvertTo-EscapedBase64 $script:DataB) + '"},' +
            '{"itemId":"ITEM-C","data":"' + ([Convert]::ToBase64String($script:DataC)) + '","note":{"nested":["x","y\"z"]}},' +
            '{"itemId":"ITEM-D","Error":{"message":"The item was not found","code":"ErrorItemNotFound"}},' +
            '{"itemId":"ITEM-E","data":null}' +
            ']}'
        $entries = @([MailboxTransfer.ExportResponseReader]::Read([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($json)), $dir))
        $entries.Count | Should -Be 5
        $entries[0].ItemId | Should -Be 'ITEM-A/1+x'
        $entries[0].Bytes | Should -Be 300000
        Get-Sha256Hex ([IO.File]::ReadAllBytes($entries[0].FilePath)) | Should -Be (Get-Sha256Hex $script:DataA)
        [IO.File]::ReadAllBytes($entries[1].FilePath) | Should -Be $script:DataB
        [IO.File]::ReadAllBytes($entries[2].FilePath) | Should -Be $script:DataC
        $entries[0].Sha256 | Should -Be ([BitConverter]::ToString([Security.Cryptography.SHA256]::HashData($script:DataA)).Replace('-', '').ToLowerInvariant())
        $entries[3].HasError | Should -BeTrue
        $entries[3].ErrorCode | Should -Be 'ErrorItemNotFound'
        $entries[3].FilePath | Should -BeNullOrEmpty
        $entries[4].FilePath | Should -BeNullOrEmpty
        (Get-ChildItem $dir -Filter *.part) | Should -BeNullOrEmpty
        (Get-ChildItem $dir -Filter *.fxs).Count | Should -Be 3
    }
    It 'cleans up the unfinished file when the answer is cut' {
        $dir = Join-Path $script:Tmp 'exp2'
        $b64 = [Convert]::ToBase64String($script:DataA)
        $json = '{"value":[{"itemId":"X","data":"' + $b64.Substring(0, 5000)    # cut in the middle of the data
        { [MailboxTransfer.ExportResponseReader]::Read([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($json)), $dir) } | Should -Throw
        (Get-ChildItem $dir -File) | Should -BeNullOrEmpty
    }
    It 'rejects invalid base64' {
        $dir = Join-Path $script:Tmp 'exp3'
        $json = '{"value":[{"itemId":"X","data":"AAAA!AAA"}]}'
        { [MailboxTransfer.ExportResponseReader]::Read([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($json)), $dir) } | Should -Throw
        (Get-ChildItem $dir -File) | Should -BeNullOrEmpty
    }
    It 'asks again at the URL of ErrorArchiveFolderMovedPermanently' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $other = 'https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/MBX:aux@y/exportItems'
        $h.Enqueue(200, ('{"value":[{"itemId":"I1","changeKey":"c1","data":"' + [Convert]::ToBase64String($script:DataB) + '"},{"itemId":"I2","Error":{"code":"ErrorArchiveFolderMovedPermanently","message":"' + $other + '"}}]}'))
        $h.Enqueue(200, ('{"value":[{"itemId":"I2","changeKey":"c2","data":"' + [Convert]::ToBase64String($script:DataC) + '"}]}'))
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        $entries = @($api.ExportItems($script:Mbx, [string[]]@('I1', 'I2'), (Join-Path $script:Tmp 'exp4')))
        $h.Requests.Count | Should -Be 2
        $h.Requests[1].Url | Should -Be $other
        ($h.Requests[1].Body | ConvertFrom-Json).itemIds | Should -Be @('I2')
        ($entries | ForEach-Object ItemId | Sort-Object) | Should -Be @('I1', 'I2')
        $entries | ForEach-Object { $_.FilePath | Should -Not -BeNullOrEmpty }
        $h.Requests.HasAuthorization | Should -Not -Contain $false
    }
    It 'refuses more than 20 ids in one request' {
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient ([MailboxTransfer.Tests.FakeHandler]::new())))
        { $api.ExportItems($script:Mbx, [string[]](1..21), $script:Tmp) } | Should -Throw '*1 to 20*'
    }
    It 'Export-MbtItems chunks by 20, writes the files and the manifest' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, ('{"value":[{"itemId":"id-01","changeKey":"k1","data":"' + [Convert]::ToBase64String($script:DataB) + '"}]}'))
        $h.Enqueue(200, ('{"value":[{"itemId":"id-21","changeKey":"k21","data":"' + [Convert]::ToBase64String($script:DataC) + '"}]}'))
        $null = Set-MbtSession -Graph (New-TestClient $h)
        $out = Join-Path $script:Tmp 'exp5'
        $info = [pscustomobject]@{ Id = 'id-01'; Type = 'IPM.Note'; Size = 1234 }
        $r = Export-MbtItems -MailboxId $script:Mbx -ItemId (1..25 | ForEach-Object { 'id-{0:00}' -f $_ }) -OutputPath $out -FolderId 'F1' -FolderPath 'Top/Lab' -ItemInfo $info
        $h.Requests.Count | Should -Be 2
        ($h.Requests[0].Body | ConvertFrom-Json).itemIds.Count | Should -Be 20
        ($h.Requests[1].Body | ConvertFrom-Json).itemIds.Count | Should -Be 5
        $r.Batches.Count | Should -Be 2
        $manifest = Get-Content -LiteralPath $r.ManifestPath -Raw | ConvertFrom-Json
        @($manifest.Items).Count | Should -Be 2
        $manifest.Items[0].SourceFolderPath | Should -Be 'Top/Lab'
        $manifest.Items[0].ListedSize | Should -Be 1234
        $manifest.Items[0].Bytes | Should -Be 1000
        Test-Path (Join-Path $out $manifest.Items[1].File) | Should -BeTrue
        Disconnect-MbtGraph
    }
}

Describe 'Import: createImportSession and importItem' {
    It 'creates the session with the Graph token and returns the pre-authenticated URL' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, '{"importUrl":"https://outlook.office365.com/api/gv1.0/Mailboxes(''MBX:a@b'')/importItem?authtoken=TOK","expirationDateTime":"2026-10-28T19:00:48.1052906Z"}')
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        $session = $api.CreateImportSession($script:Mbx)
        $h.Requests[0].Method | Should -Be 'POST'
        $h.Requests[0].Url | Should -BeLike '*/createImportSession'
        $h.Requests[0].HasAuthorization | Should -BeTrue
        $session.ImportUrl | Should -BeLike '*authtoken=TOK'
        $session.Expiration.Year | Should -Be 2026
    }
    It 'posts the file in create mode, without Authorization, with an exact Content-Length (<Size> bytes)' -ForEach @(
        @{ Size = 0 }, @{ Size = 1 }, @{ Size = 2 }, @{ Size = 3 }, @{ Size = 49152 }, @{ Size = 49153 }, @{ Size = 100000 }
    ) {
        $file = Join-Path $script:Tmp "imp-$Size.fxs"
        $bytes = [byte[]]::new($Size); [Random]::new($Size).NextBytes($bytes)
        [IO.File]::WriteAllBytes($file, $bytes)
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, '{"@odata.context":"x","itemId":"NEW-ID=","changeKey":"CK"}')
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        $url = "https://outlook.office365.com/api/gv1.0/Mailboxes('MBX:a@b')/importItem?authtoken=eyJ.TOK"
        $result = $api.ImportItem([MailboxTransfer.ImportSession]@{ MailboxId = $script:Mbx; ImportUrl = $url }, 'EDSVrdi3lRAAEED0yTAAA=', $file)
        $req = $h.Requests[0]
        $req.Url | Should -Be $url
        $req.HasAuthorization | Should -BeFalse
        $req.ContentType | Should -Be 'application/json'
        $req.ContentLength | Should -Be $req.BodyBytes.Length
        $body = $req.Body | ConvertFrom-Json
        $body.FolderId | Should -Be 'EDSVrdi3lRAAEED0yTAAA='
        $body.Mode | Should -Be 'create'
        Get-Sha256Hex ([Convert]::FromBase64String($body.Data)) | Should -Be (Get-Sha256Hex $bytes)
        $result.ItemId | Should -Be 'NEW-ID='
        $result.Bytes | Should -Be $Size
    }
    It 'Import-MbtItem does nothing with -WhatIf' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $null = Set-MbtSession -Graph (New-TestClient $h)
        $file = Join-Path $script:Tmp 'whatif.fxs'
        [IO.File]::WriteAllBytes($file, [byte[]](1..10))
        Import-MbtItem -MailboxId $script:Mbx -FolderId 'F' -Path $file -WhatIf
        $h.Requests.Count | Should -Be 0
        Disconnect-MbtGraph
    }
}

Describe 'Sign-in helpers' {
    It 'decodes the claims of an access token' {
        $enc = { param($o) [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($o | ConvertTo-Json -Compress))).TrimEnd('=').Replace('+', '-').Replace('/', '_') }
        $jwt = (& $enc @{ alg = 'RS256' }) + '.' + (& $enc @{ upn = 'adele@contoso.test'; scp = 'MailboxFolder.Read MailboxItem.ImportExport'; tid = 'T'; appid = 'A'; aud = 'https://graph.microsoft.com'; exp = 1793000000 }) + '.sig'
        $info = Get-MbtTokenInfo -AccessToken $jwt
        $info.Account | Should -Be 'adele@contoso.test'
        $info.Scopes | Should -Be @('MailboxFolder.Read', 'MailboxItem.ImportExport')
        $info.ExpiresOn.Year | Should -Be 2026
    }
    It 'compiles the MSAL engine and builds a sign-in object (no network)' -Skip:(-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        Initialize-MbtAuthEngine
        $auth = [MailboxTransfer.GraphAuth]::new('1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d', '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d', [string[]]@('https://graph.microsoft.com/.default'), (Join-Path $script:Tmp 'token.cache'))
        $auth | Should -Not -BeNullOrEmpty
        $auth.Provider('adele@contoso.test') | Should -BeOfType [Func[string]]
    }
}

Describe 'Mailbox ids and folder creation' {
    BeforeAll {
        $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $script:Upn = 'adele.vance@contoso.com'
        $script:MailboxUrl = "$($script:Base)/admin/exchange/mailboxes/$($script:Mbx)"
    }
    It 'takes an id MBX:<guid>@<tenant> as is, completes a bare guid with the tenant, refuses anything else and points to Get-MailboxLocation' {
        ConvertTo-MbtMailboxId -Value $script:Mbx -TenantId $script:Tenant | Should -Be $script:Mbx
        ConvertTo-MbtMailboxId -Value '11111111-2222-3333-4444-555555555555' -TenantId $script:Tenant | Should -Be "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
        { ConvertTo-MbtMailboxId -Value 'adele' -TenantId $script:Tenant } | Should -Throw '*Get-MailboxLocation*'
        { ConvertTo-MbtMailboxId -Value 'MBX:nope@nope' -TenantId $script:Tenant } | Should -Throw '*not a mailbox id*'
    }
    It 'reads the primary mailbox id from GET /users/{upn}/settings/exchange' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, ('{"primaryMailboxId":"' + $script:Mbx + '"}'))
        $null = Set-MbtSession -Graph (New-TestClient $h)
        Get-MbtPrimaryMailboxId -UserPrincipalName $script:Upn -TenantId $script:Tenant | Should -Be $script:Mbx
        $h.Requests[0].Url | Should -Be "$($script:Base)/users/adele.vance%40contoso.com/settings/exchange"
        Disconnect-MbtGraph
    }
    It 'says to give -PrimaryMailboxId and how to get it when Graph refuses (permission missing)' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(403, '{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}')
        $null = Set-MbtSession -Graph (New-TestClient $h)
        $caught = $null
        try { Get-MbtPrimaryMailboxId -UserPrincipalName $script:Upn -TenantId $script:Tenant } catch { $caught = $_.Exception.Message }
        $caught | Should -BeLike '*HTTP 403*'
        $caught | Should -BeLike '*User.Read.All*'
        $caught | Should -BeLike '*-PrimaryMailboxId MBX:<mailbox guid>@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d*'
        $caught | Should -BeLike "*Get-MailboxLocation -User $($script:Upn)*"
        Disconnect-MbtGraph
    }
    It 'tells an SMTP address / UPN from a physical id or a bare guid' {
        Test-MbtMailboxAddress 'MBT-Shared-Projects@contoso.test' | Should -BeTrue
        Test-MbtMailboxAddress $script:Mbx | Should -BeFalse
        Test-MbtMailboxAddress '11111111-2222-3333-4444-555555555555' | Should -BeFalse
        Test-MbtMailboxAddress '' | Should -BeFalse
    }
    It 'resolves the address of a shared mailbox through Graph settings/exchange when the token allows it' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, ('{"primaryMailboxId":"' + $script:Mbx + '"}'))
        $null = Set-MbtSession -Graph (New-TestClient $h)
        Resolve-MbtMailboxAddress -Address 'MBT-Shared-Projects@contoso.com' -TenantId $script:Tenant | Should -Be $script:Mbx
        $h.Requests.Count | Should -Be 1
        Disconnect-MbtGraph
    }
    It 'resolves the address through the import-export API (parentMailboxUrl of a folder) when Graph refuses settings/exchange (no User.Read.All)' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(403, '{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}')
        $h.Enqueue(200, ('{"value":[{"id":"F1","displayName":"Archive","parentMailboxUrl":"https://graph.microsoft.com/v1.0/admin/Exchange/Mailboxes/' + $script:Mbx + '","childFolderCount":0,"totalItemCount":0}]}'))
        $null = Set-MbtSession -Graph (New-TestClient $h)
        Resolve-MbtMailboxAddress -Address 'MBT-Shared-Projects@contoso.com' -TenantId $script:Tenant | Should -Be $script:Mbx
        $h.Requests[1].Url | Should -BeLike "$($script:Base)/admin/exchange/mailboxes/MBT-Shared-Projects%40contoso.com/folders*"
        Disconnect-MbtGraph
    }
    It 'Resolve-MbtMailboxId resolves an address given as -MailboxId and still takes an id as is' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(200, ('{"primaryMailboxId":"' + $script:Mbx + '"}'))
        $null = Set-MbtSession -Graph (New-TestClient $h)
        Resolve-MbtMailboxId -Mailbox Primary -TenantId $script:Tenant -UserPrincipalName $script:Upn -MailboxId 'MBT-Shared-Projects@contoso.com' | Should -Be $script:Mbx
        Resolve-MbtMailboxId -Mailbox Primary -TenantId $script:Tenant -UserPrincipalName $script:Upn -MailboxId $script:Mbx | Should -Be $script:Mbx
        $h.Requests.Count | Should -Be 1
        Disconnect-MbtGraph
    }
    It 'explains what is missing when neither Graph nor the import-export API can resolve the address (no FullAccess yet)' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(403, '{"error":{"code":"ErrorAccessDenied","message":"Access is denied."}}')
        $h.Enqueue(404, '{"error":{"code":"ErrorItemNotFound","message":"The specified object was not found in the store., Default folder Root not found."}}')
        $null = Set-MbtSession -Graph (New-TestClient $h)
        $caught = $null
        try { Resolve-MbtMailboxAddress -Address 'MBT-Shared-Projects@contoso.com' -TenantId $script:Tenant } catch { $caught = $_.Exception.Message }
        $caught | Should -BeLike '*Graph settings/exchange*HTTP 403*'
        $caught | Should -BeLike '*import-export API: HTTP 404 ErrorItemNotFound*'
        $caught | Should -BeLike '*-MailboxId MBX:<mailbox guid>@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d*'
        $caught | Should -BeLike '*FullAccess*'
        $caught | Should -BeLike '*User.Read.All*'
        Disconnect-MbtGraph
    }
    It 'names the missing right and permissions on an access denial for another mailbox, and says nothing otherwise' {
        $denied = 'POST https://graph.microsoft.com/v1.0/admin/exchange/mailboxes/MBX:x/folders -> 403 ErrorAccessDenied: Access is denied.'
        Get-MbtAccessHint -Message $denied -Mode Import -OtherMailbox -Scopes @('MailboxFolder.Read') | Should -BeLike '*Missing in the token: MailboxFolder.ReadWrite, MailboxItem.ImportExport*FullAccess*'
        Get-MbtAccessHint -Message $denied -Mode Import -OtherMailbox -Scopes @('MailboxFolder.ReadWrite', 'MailboxItem.ImportExport') | Should -BeLike '*needed are in the token, so the refusal comes from Exchange*FullAccess*60 minutes*'
        Get-MbtAccessHint -Message $denied -Mode Browse -OtherMailbox | Should -BeLike '*FullAccess*Mail.Read.Shared*'
        Get-MbtAccessHint -Message $denied -Mode Export -OtherMailbox -Scopes @('MailboxFolder.Read') | Should -BeLike '*Missing in the token: MailboxItem.Read*FullAccess*'
        Get-MbtAccessHint -Message $denied -Mode Import | Should -Be ''
        Get-MbtAccessHint -Message 'HTTP 500 ErrorInternalServerError' -Mode Import -OtherMailbox | Should -Be ''
    }
    It 'reports the permissions the service names in innererror.internalexception, and puts them in the hint' {
        $body = '{"error":{"code":"ErrorAccessDenied","message":"Access is denied. Check credentials and try again.","innererror":{"message":"Access is denied. Check credentials and try again.","type":"X","internalexception":{"message":"Delegate access to this API requires the following permissions: ''MailboxFolder.ReadWrite.All,MailboxFolder.ReadWrite.Shared''. However, the application only has the following permissions granted: ''MailboxFolder.ReadWrite''.","type":"Y"}}}}'
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(403, $body)
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        $caught = $null
        try { $api.CreateFolder($script:Mbx, 'Lab') } catch { $caught = $_.Exception.InnerException.Message; if (-not $caught) { $caught = $_.Exception.Message } }
        $caught | Should -BeLike '*403 ErrorAccessDenied*MailboxFolder.ReadWrite.All,MailboxFolder.ReadWrite.Shared*'
        $hint = Get-MbtAccessHint -Message $caught -Mode Import -OtherMailbox -Scopes @('MailboxFolder.ReadWrite', 'MailboxItem.ImportExport')
        $hint | Should -BeLike '*needs: MailboxFolder.ReadWrite.All,MailboxFolder.ReadWrite.Shared*do not exist as delegated*application permissions*FullAccess*'
        $hint | Should -Not -BeLike '*refusal comes from Exchange*'
    }
    It 'creates a top-level folder with POST .../folders (displayName and type) and a child with .../childFolders' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(201, '{"id":"NEW1","displayName":"Lab","parentFolderId":"ROOT","childFolderCount":0,"totalItemCount":0,"type":"IPF.Note"}')
        $h.Enqueue(201, '{"id":"NEW2","displayName":"Sub","parentFolderId":"NEW1","childFolderCount":0,"totalItemCount":0,"type":"IPF.Note"}')
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        $top = $api.CreateFolder($script:Mbx, 'Lab')
        $sub = $api.CreateFolder($script:Mbx, 'Sub', $top.Id, 'IPF.Note', $top.Path)
        $h.Requests[0].Method | Should -Be 'POST'
        $h.Requests[0].Url | Should -Be "$($script:MailboxUrl)/folders"
        $b = $h.Requests[0].Body | ConvertFrom-Json
        $b.displayName | Should -Be 'Lab'
        $b.type | Should -Be 'IPF.Note'
        $h.Requests[1].Url | Should -Be "$($script:MailboxUrl)/folders/NEW1/childFolders"
        $sub.Path | Should -Be 'Lab/Sub'
        $sub.Depth | Should -Be 1
    }
    It 'New-MbtFolderPath creates only the missing folders of a path, never touches the existing ones, and does nothing with -WhatIf' {
        $tree = @([pscustomobject]@{ Id = 'T'; DisplayName = 'Lab'; Path = 'Lab'; ParentFolderId = 'ROOT' })
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $null = Set-MbtSession -Graph (New-TestClient $h)
        $existing = New-MbtFolderPath -MailboxId $script:Mbx -Tree $tree -Path 'lab' -Confirm:$false
        $existing.Id | Should -Be 'T'
        $h.Requests.Count | Should -Be 0
        $plan = New-MbtFolderPath -MailboxId $script:Mbx -Tree $tree -Path 'Lab/Import/Deep' -WhatIf
        $plan | Should -BeNullOrEmpty
        $h.Requests.Count | Should -Be 0
        $h.Enqueue(201, '{"id":"N1","displayName":"Import","parentFolderId":"T","childFolderCount":0,"totalItemCount":0,"type":"IPF.Note"}')
        $h.Enqueue(201, '{"id":"N2","displayName":"Deep","parentFolderId":"N1","childFolderCount":0,"totalItemCount":0,"type":"IPF.Note"}')
        $made = New-MbtFolderPath -MailboxId $script:Mbx -Tree $tree -Path 'Lab/Import/Deep' -Confirm:$false
        $made.Id | Should -Be 'N2'
        $made.Path | Should -Be 'Lab/Import/Deep'
        $h.Requests.Count | Should -Be 2
        $h.Requests[0].Url | Should -Be "$($script:MailboxUrl)/folders/T/childFolders"
        $h.Requests[1].Url | Should -Be "$($script:MailboxUrl)/folders/N1/childFolders"
        Disconnect-MbtGraph
    }
    It 'Resolve-MbtFolder -NullIfMissing returns $null for a folder that does not exist' {
        Resolve-MbtFolder -Tree @() -Name 'Nope' -NullIfMissing | Should -BeNullOrEmpty
    }
    It 'the shipped configuration holds placeholders only; the local file is read after it and the lab defaults come from there' {
        $shipped = Import-PowerShellDataFile -LiteralPath (Join-Path $script:Root 'config\MailboxTransfer.config.psd1')
        foreach ($k in @($shipped.Lab.Keys)) { $shipped.Lab[$k] | Should -BeNullOrEmpty -Because "Lab.$k of the shipped file is a placeholder" }
        $lab = (Import-MbtConfiguration).Lab                         # the synthetic local file of tests\support\Import-StoreEngine.ps1
        $lab.UserPrincipalName | Should -Be $script:Upn
        $lab.UserObjectId | Should -Be '1a2b3c05-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $lab.SourceFolder | Should -Be 'MBT-Fill-20261028'
        $saved = $env:MAILBOXTRANSFER_LOCAL_CONFIG
        try {
            $env:MAILBOXTRANSFER_LOCAL_CONFIG = Join-Path $script:Tmp 'no-such-local-file.psd1'
            (Import-MbtConfiguration).Lab.TenantId | Should -BeNullOrEmpty
            $bad = Join-Path $script:Tmp 'bad-local.psd1'
            Set-Content -LiteralPath $bad -Value "@{ Lab = @{ Nonsense = 1 } }"
            $env:MAILBOXTRANSFER_LOCAL_CONFIG = $bad
            { Import-MbtConfiguration } | Should -Throw '*unknown key*'
        } finally { $env:MAILBOXTRANSFER_LOCAL_CONFIG = $saved }
    }
    It 'no tenant name other than contoso is written in the tree, and the local denylist (tests\privacy-denylist.local.txt, git-ignored, optional) finds nothing' {
        $files = Get-ChildItem -LiteralPath $script:RepoRoot -Recurse -Include *.ps1, *.psm1, *.psd1, *.cs, *.md, *.xaml, *.json, *.html, *.txt -File |
            Where-Object { $_.FullName -notmatch '[\\/](\.git|out|lib)[\\/]' -and $_.Name -notmatch '\.local\.' }
        $tenants = @($files | Select-String -Pattern '(?i)\b(?!contoso\b)[a-z0-9-]+\.onmicrosoft\.com' | Where-Object { $_.Matches[0].Value -notmatch '(?i)^(tenant|yourtenant|your-tenant|example)\.' })
        $tenants | Should -BeNullOrEmpty -Because (($tenants | ForEach-Object { "$($_.Path):$($_.LineNumber)" }) -join ', ')
        $deny = Join-Path $PSScriptRoot 'privacy-denylist.local.txt'
        if (Test-Path -LiteralPath $deny) {
            $terms = @(Get-Content -LiteralPath $deny | Where-Object { $_.Trim() -and -not $_.StartsWith('#') } | ForEach-Object { $_.Trim() })
            $hits = @($files | Where-Object { $_.FullName -ne $deny } | Select-String -SimpleMatch -Pattern $terms)
            $hits | Should -BeNullOrEmpty -Because (($hits | ForEach-Object { "$($_.Path):$($_.LineNumber)" }) -join ', ')
        }
    }
}

Describe 'Device code sign-in' -Skip:(-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    BeforeAll { Initialize-MbtAuthEngine }
    It 'writes a line starting with DEVICE CODE: (code and URL) to the standard output at once, and to a file that is overwritten each time' {
        $file = Join-Path $script:Tmp 'out\device-code.txt'
        $line1 = [MailboxTransfer.GraphAuth]::DeviceCodeLine('ABC123XYZ', 'https://microsoft.com/devicelogin', [DateTimeOffset]::UtcNow.AddMinutes(15))
        $line2 = [MailboxTransfer.GraphAuth]::DeviceCodeLine('SECOND77', 'https://microsoft.com/devicelogin', [DateTimeOffset]::UtcNow.AddMinutes(15))
        $old = [Console]::Out
        $sw = [IO.StringWriter]::new()
        try {
            [Console]::SetOut($sw)
            [MailboxTransfer.GraphAuth]::PublishDeviceCode($line1, $file)
            $first = [IO.File]::ReadAllText($file)
            [MailboxTransfer.GraphAuth]::PublishDeviceCode($line2, $file)
        } finally { [Console]::SetOut($old) }
        $sw.ToString() | Should -BeLike 'DEVICE CODE: ABC123XYZ  URL: https://microsoft.com/devicelogin*SECOND77*'
        $line1 | Should -BeLike 'DEVICE CODE: ABC123XYZ  URL: https://microsoft.com/devicelogin  (valid until *)'
        $first.TrimEnd() | Should -Be $line1
        [IO.File]::ReadAllText($file).TrimEnd() | Should -Be $line2 -Because 'the file is overwritten, not appended'
    }
    It 'creates the folder of the file and accepts no file' {
        $file = Join-Path $script:Tmp 'a\b\c\device-code.txt'
        $old = [Console]::Out
        try {
            [Console]::SetOut([IO.StringWriter]::new())
            [MailboxTransfer.GraphAuth]::PublishDeviceCode('DEVICE CODE: X', $file)
            [MailboxTransfer.GraphAuth]::PublishDeviceCode('DEVICE CODE: X', $null)
        } finally { [Console]::SetOut($old) }
        Test-Path -LiteralPath $file | Should -BeTrue
    }
    It 'builds a device code sign-in object (no network) with the file of the console line' {
        $auth = [MailboxTransfer.GraphAuth]::new('1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d', '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d', [string[]]@('https://graph.microsoft.com/.default'), (Join-Path $script:Tmp 'dc.cache'), 'http://localhost', $true, (Join-Path $script:Tmp 'dc.txt'))
        $auth.DeviceCode | Should -BeTrue
        $auth.DeviceCodeFile | Should -BeLike '*dc.txt'
    }
    It 'does not fall back to another cached account: without a cache the silent call asks for a sign-in, not for the browser, when interactive is false' {
        $auth = [MailboxTransfer.GraphAuth]::new('1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d', '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d', [string[]]@('https://graph.microsoft.com/.default'), (Join-Path $script:Tmp 'none.cache'))
        { $auth.GetToken('adele@contoso.test', $false, [Threading.CancellationToken]::None) } | Should -Throw '*Not signed in*'
    }
}

Describe 'Phase 0 on the store: export, preview, read, import' {
    BeforeAll {
        $script:Msg = New-TestMessage
        $script:MsgSha = [BitConverter]::ToString([Security.Cryptography.SHA256]::HashData($script:Msg)).Replace('-', '').ToLowerInvariant()
        $script:Other = [byte[]]::new(5000); [Random]::new(7).NextBytes($script:Other)
        function New-PhaseStore([string]$Name) {
            $store = Open-MbtExportStore -Path (Join-Path $script:Tmp $Name) -Name $Name
            $tree = @(
                [pscustomobject]@{ Id = 'ROOTF'; DisplayName = 'Top'; ParentFolderId = 'ROOT'; Type = 'IPF.Note'; TotalItemCount = 0; TotalSize = -1 }
                [pscustomobject]@{ Id = 'F1'; DisplayName = 'MBT-Fill-20261028'; ParentFolderId = 'ROOTF'; Type = 'IPF.Note'; TotalItemCount = 3; TotalSize = -1 })
            Save-MbtFolderTree -Store $store -MailboxId $script:Mbx -Kind Primary -UserPrincipalName 'adele.vance@contoso.com' -Tree $tree
            Save-MbtItemList -Store $store -MailboxId $script:Mbx -FolderId 'F1' -Items @(
                [pscustomobject]@{ Id = 'I1'; ChangeKey = 'c1'; Type = 'IPM.Note'; Size = $script:Msg.Length }
                [pscustomobject]@{ Id = 'I2'; ChangeKey = 'c2'; Type = 'IPM.Note'; Size = 5000 }
                [pscustomobject]@{ Id = 'I3'; ChangeKey = 'c3'; Type = 'IPM.Note'; Size = 10 })
            return $store
        }
    }
    It 'records the folder tree and the listed items (Pending) in the store' {
        $store = New-PhaseStore 'ps-tree'
        try {
            $store.GetMailboxes().Count | Should -Be 1
            $node = $store.GetFolderTree($script:Mbx)[0]
            $node.DisplayName | Should -Be 'Top'
            $node.Children[0].Path | Should -Be 'Top/MBT-Fill-20261028'
            $node.Children[0].Own.Pending | Should -Be 3
        } finally { $store.Dispose() }
    }
    It 'exports into packs, decodes the preview, marks the failed item, deletes the temporary file and compares the SHA-256' {
        $store = New-PhaseStore 'ps-export'
        try {
            $h = [MailboxTransfer.Tests.FakeHandler]::new()
            $json = '{"value":[{"itemId":"I1","changeKey":"c1","data":"' + [Convert]::ToBase64String($script:Msg) + '"},' +
                '{"itemId":"I2","changeKey":"c2","data":"' + [Convert]::ToBase64String($script:Other) + '"},' +
                '{"itemId":"I3","Error":{"code":"ErrorItemNotFound","message":"gone"}}]}'
            $h.Enqueue(200, $json)
            $null = Set-MbtSession -Graph (New-TestClient $h)
            $r = Export-MbtItemsToStore -Store $store -MailboxId $script:Mbx -ItemId 'I1', 'I2', 'I3' -TempPath (Join-Path $script:Tmp 'ps-export-tmp') -FolderId 'F1'
            $r.Batches.Count | Should -Be 1
            $r.Entries[0].Sha256Matches | Should -BeTrue
            $r.Entries[0].Sha256 | Should -Be $script:MsgSha
            $r.Entries[2].ErrorCode | Should -Be 'ErrorItemNotFound'
            $k1 = $store.FindItemKey($script:Mbx, 'I1')
            $i1 = $store.GetItem($k1)
            $i1.State | Should -Be 'Exported'
            $i1.Sha256 | Should -Be $script:MsgSha
            $store.GetItem($store.FindItemKey($script:Mbx, 'I3')).State | Should -Be 'Failed'
            $store.GetItem($store.FindItemKey($script:Mbx, 'I3')).LastError | Should -BeLike '*ErrorItemNotFound*'
            $p = $store.GetPreview($k1)
            $p.Subject | Should -Be 'Quarterly report'
            $p.MessageClass | Should -Be 'IPM.Note'
            $p.HasAttachments | Should -BeTrue
            $p.AttachmentNames | Should -BeLike '*report.pdf*'
            $p.To | Should -BeLike '*Alice*'
            Get-ChildItem (Join-Path $script:Tmp 'ps-export-tmp') -Recurse -File -ErrorAction SilentlyContinue | Should -BeNullOrEmpty
            Get-ChildItem (Join-Path $script:Tmp 'ps-export\packs') -File | Should -HaveCount 1
            @($store.VerifyAll([Threading.CancellationToken]::None) | Where-Object { -not $_.Ok }) | Should -BeNullOrEmpty
            @(Get-MbtStoreItems -Store $store).Count | Should -Be 2
            (Get-MbtStoreItems -Store $store | Select-Object -First 1).FolderPath | Should -Be 'Top/MBT-Fill-20261028'
            @(Get-MbtStoreItems -Store $store -State Failed).Count | Should -Be 1
        } finally { $store.Dispose(); Disconnect-MbtGraph }
    }
    It 'decodes an item of the store with the same reader as a file' {
        $store = New-PhaseStore 'ps-read'
        try {
            $h = [MailboxTransfer.Tests.FakeHandler]::new()
            $h.Enqueue(200, ('{"value":[{"itemId":"I1","data":"' + [Convert]::ToBase64String($script:Msg) + '"}]}'))
            $null = Set-MbtSession -Graph (New-TestClient $h)
            $null = Export-MbtItemsToStore -Store $store -MailboxId $script:Mbx -ItemId 'I1' -TempPath (Join-Path $script:Tmp 'ps-read-tmp')
            $d = Read-MbtFxStream -Store $store -ItemKey $store.FindItemKey($script:Mbx, 'I1')
            $d.Summary.Subject | Should -Be 'Quarterly report'
            $d.Summary.Attachments.Count | Should -Be 2
            $d.Read.BytesRead | Should -Be $script:Msg.Length
            $file = Join-Path $script:Tmp 'ps-read.fxs'
            [IO.File]::WriteAllBytes($file, $script:Msg)
            (Read-MbtFxStream -Path $file).Summary.Subject | Should -Be $d.Summary.Subject
        } finally { $store.Dispose(); Disconnect-MbtGraph }
    }
    It 'registers on the fly an item the listing did not bring when the folder is given, and refuses it otherwise' {
        $store = New-PhaseStore 'ps-fly'
        try {
            $h = [MailboxTransfer.Tests.FakeHandler]::new()
            $h.Enqueue(200, ('{"value":[{"itemId":"NEW","data":"' + [Convert]::ToBase64String($script:Other) + '"}]}'))
            $null = Set-MbtSession -Graph (New-TestClient $h)
            { Export-MbtItemsToStore -Store $store -MailboxId $script:Mbx -ItemId 'NEW' -TempPath (Join-Path $script:Tmp 'ps-fly-tmp') } | Should -Throw '*not in the store*'
            $null = Export-MbtItemsToStore -Store $store -MailboxId $script:Mbx -ItemId 'NEW' -TempPath (Join-Path $script:Tmp 'ps-fly-tmp') -FolderId 'F1'
            $store.GetItem($script:Mbx, 'NEW').State | Should -Be 'Exported'
        } finally { $store.Dispose(); Disconnect-MbtGraph }
    }
    It 'imports an item of the store: job Done, body streamed from the pack, no Authorization header' {
        $store = New-PhaseStore 'ps-import'
        try {
            $h = [MailboxTransfer.Tests.FakeHandler]::new()
            $h.Enqueue(200, ('{"value":[{"itemId":"I1","data":"' + [Convert]::ToBase64String($script:Msg) + '"}]}'))
            $h.Enqueue(200, '{"importUrl":"https://outlook.office365.com/api/gv1.0/Mailboxes(''MBX:a@b'')/importItem?authtoken=TOK","expirationDateTime":"2026-10-28T19:00:48Z"}')
            $h.Enqueue(200, '{"itemId":"NEW-ID=","changeKey":"CK"}')
            $null = Set-MbtSession -Graph (New-TestClient $h)
            $null = Export-MbtItemsToStore -Store $store -MailboxId $script:Mbx -ItemId 'I1' -TempPath (Join-Path $script:Tmp 'ps-import-tmp')
            $key = $store.FindItemKey($script:Mbx, 'I1')
            $arch = 'MBX:99999999-8888-7777-6666-555555555555@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
            $r = Import-MbtStoreItem -Store $store -ItemKey $key -MailboxId $arch -FolderId 'TARGET=' -Confirm:$false
            $r.ItemId | Should -Be 'NEW-ID='
            $r.Bytes | Should -Be $script:Msg.Length
            $req = $h.Requests[2]
            $req.Url | Should -BeLike '*authtoken=TOK'
            $req.HasAuthorization | Should -BeFalse
            $req.ContentLength | Should -Be $req.BodyBytes.Length
            $body = $req.Body | ConvertFrom-Json
            $body.FolderId | Should -Be 'TARGET='
            $body.Mode | Should -Be 'create'
            Get-Sha256Hex ([Convert]::FromBase64String($body.Data)) | Should -Be $script:MsgSha
            $job = @($store.GetImportItems(1, [MailboxTransfer.StoreImportState]::Done, 10))
            $job.Count | Should -Be 1
            $job[0].TargetItemId | Should -Be 'NEW-ID='
            $job[0].TargetFolder | Should -Be 'TARGET='
        } finally { $store.Dispose(); Disconnect-MbtGraph }
    }
    It 'explains the 409 of an auxiliary archive (target session expected in another mailbox), for a store item and for a file' {
        $store = New-PhaseStore 'ps-aux'
        try {
            $h = [MailboxTransfer.Tests.FakeHandler]::new()
            $h.Enqueue(200, ('{"value":[{"itemId":"I1","data":"' + [Convert]::ToBase64String($script:Msg) + '"}]}'))
            $null = Set-MbtSession -Graph (New-TestClient $h)
            $null = Export-MbtItemsToStore -Store $store -MailboxId $script:Mbx -ItemId 'I1' -TempPath (Join-Path $script:Tmp 'ps-aux-tmp')
            $key = $store.FindItemKey($script:Mbx, 'I1')
            $aux = 'MBX:aaaaaaaa-1111-2222-3333-444444444444@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
            $conflict = '{"error":{"code":"ErrorInvalidSession","message":"Invalid import session. The target session is expected in mailbox ' + $aux + '."}}'
            $url = '{"importUrl":"https://outlook.office365.com/api/gv1.0/Mailboxes(''MBX:a@b'')/importItem?authtoken=TOK"}'
            $h.Enqueue(200, $url); $h.Enqueue(409, $conflict)
            $caught = $null
            try { Import-MbtStoreItem -Store $store -ItemKey $key -MailboxId $script:Mbx -FolderId 'T' -Confirm:$false } catch { $caught = $_.Exception.Message }
            $caught | Should -BeLike "*auxiliary archive*$aux*AuxArchive*Get-MailboxLocation*nothing was imported*"
            @($store.GetImportItems(1, [MailboxTransfer.StoreImportState]::Failed, 10)).Count | Should -Be 1
            $file = Join-Path $script:Tmp 'auxiliary-item.fxs'
            [IO.File]::WriteAllBytes($file, $script:Msg)
            $h.Enqueue(200, $url); $h.Enqueue(409, $conflict)
            $caught = $null
            try { Import-MbtItem -MailboxId $script:Mbx -FolderId 'T' -Path $file -Confirm:$false } catch { $caught = $_.Exception.Message }
            $caught | Should -BeLike "*auxiliary archive*$aux*"
        } finally { $store.Dispose(); Disconnect-MbtGraph }
    }
    It 'marks the import Failed on a refusal of the service, does nothing with -WhatIf, and refuses an item that is not exported' {
        $store = New-PhaseStore 'ps-import2'
        try {
            $h = [MailboxTransfer.Tests.FakeHandler]::new()
            $h.Enqueue(200, ('{"value":[{"itemId":"I1","data":"' + [Convert]::ToBase64String($script:Msg) + '"}]}'))
            $null = Set-MbtSession -Graph (New-TestClient $h)
            $null = Export-MbtItemsToStore -Store $store -MailboxId $script:Mbx -ItemId 'I1' -TempPath (Join-Path $script:Tmp 'ps-import2-tmp')
            $key = $store.FindItemKey($script:Mbx, 'I1')
            $before = $h.Requests.Count
            Import-MbtStoreItem -Store $store -ItemKey $key -MailboxId $script:Mbx -FolderId 'T' -WhatIf
            $h.Requests.Count | Should -Be $before
            { Import-MbtStoreItem -Store $store -ItemKey $store.FindItemKey($script:Mbx, 'I2') -MailboxId $script:Mbx -FolderId 'T' -Confirm:$false } | Should -Throw '*not exported*'
            $h.Enqueue(200, '{"importUrl":"https://outlook.office365.com/api/gv1.0/Mailboxes(''MBX:a@b'')/importItem?authtoken=TOK"}')
            $h.Enqueue(400, '{"error":{"code":"ErrorInvalidIdMalformed","message":"bad folder"}}')
            { Import-MbtStoreItem -Store $store -ItemKey $key -MailboxId $script:Mbx -FolderId 'T' -Confirm:$false } | Should -Throw '*bad folder*'
            @($store.GetImportItems(1, [MailboxTransfer.StoreImportState]::Failed, 10)).Count | Should -Be 1
        } finally { $store.Dispose(); Disconnect-MbtGraph }
    }
}

Describe 'Invoke-Phase0.ps1 (against a fake service, store included)' {
    BeforeAll {
        $script:Phase0 = Join-Path $script:RepoRoot 'tools\Invoke-Phase0.ps1'
        Remove-Item -LiteralPath (Join-Path $script:Tmp 'phase0') -Recurse -Force -ErrorAction SilentlyContinue   # the tests of this block build on each other (same store)
        $script:PrimaryId = 'MBX:1a2b3c03-5e7f-4a1b-9c2d-3e4f5a6b7c8d@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $script:ArchiveId = 'MBX:1a2b3c04-5e7f-4a1b-9c2d-3e4f5a6b7c8d@1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $script:Fx = New-TestMessage
        function New-FakeService {
            $r = [MailboxTransfer.Tests.RouteHandler]::new()
            $p = "$($script:PrimaryId)"
            $a = "$($script:ArchiveId)"
            $item = { param($id, $size) '{"id":"' + $id + '","type":"IPM.Note","size":' + $size + ',"changeKey":"ck' + $id + '","createdDateTime":"2026-10-28T09:00:00Z","lastModifiedDateTime":"2026-10-28T09:00:00Z"}' }
            $r.Add('GET', '/users/', 403, '{"error":{"code":"Authorization_RequestDenied","message":"Insufficient privileges"}}')
            foreach ($id in $p, $a) { $r.Add('GET', "$id/folders?`$top=1", 200, '{"value":[]}') }
            $r.Add('GET', "$p/folders/F1/items", 200, '{"value":[' + ((& $item 'I1' 500) + ',' + (& $item 'I2' $script:Fx.Length) + ',' + (& $item 'I3' 800)) + ']}')
            $r.Add('GET', "$p/folders", 200, '{"value":[{"id":"R","displayName":"Top","parentFolderId":"ROOT","childFolderCount":1,"totalItemCount":0,"type":"IPF.Note"},{"id":"F1","displayName":"MBT-Fill-20261028","parentFolderId":"R","childFolderCount":0,"totalItemCount":3,"type":"IPF.Note"}]}')
            $r.Add('GET', "$a/folders/NEWF/items", 200, '{"value":[]}')
            $r.Add('GET', "$a/folders/NEWF/items", 200, '{"value":[' + (& $item 'NEW-ID' $script:Fx.Length) + ']}')
            $r.Add('GET', "$a/folders", 200, '{"value":[{"id":"AR","displayName":"Top","parentFolderId":"AROOT","childFolderCount":0,"totalItemCount":0,"type":"IPF.Note"}]}')
            $r.Add('POST', "$p/exportItems", 200, '{"value":[{"itemId":"I2","changeKey":"ckI2","data":"' + [Convert]::ToBase64String($script:Fx) + '"}]}')
            $r.Add('POST', "$a/exportItems", 200, '{"value":[{"itemId":"NEW-ID","changeKey":"ckN","data":"' + [Convert]::ToBase64String($script:Fx) + '"}]}')
            $r.Add('POST', "$a/folders", 201, '{"id":"NEWF","displayName":"MBT-Import-20261028","parentFolderId":"AROOT","childFolderCount":0,"totalItemCount":0,"type":"IPF.Note"}')
            $r.Add('POST', "$a/createImportSession", 200, '{"importUrl":"https://outlook.office365.com/api/gv1.0/Mailboxes(''MBX:a@b'')/importItem?authtoken=TOK","expirationDateTime":"2026-10-28T19:00:48Z"}')
            $r.Add('POST', '/importItem?', 200, '{"itemId":"NEW-ID","changeKey":"CK"}')
            return $r
        }
        function Invoke-FakePhase0 {
            param([hashtable]$Parameters = @{})
            $global:MbtRoute = New-FakeService
            $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
            $global:MbtClient = [MailboxTransfer.GraphClient]::new([Func[string]] { 'test-token' }, $o, $global:MbtRoute)
            $text = & $script:Phase0 -OutputPath (Join-Path $script:Tmp 'phase0') @Parameters *>&1 | Out-String
            return $text
        }
    }
    BeforeEach {
        Mock Connect-MbtGraph {
            $session = Set-MbtSession -Graph $global:MbtClient
            $session.Token = [pscustomobject]@{
                Account = 'adele.vance@contoso.com'; Name = 'Adele Vance'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; AppId = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
                Audience = 'https://graph.microsoft.com'; Scopes = @('MailboxFolder.ReadWrite', 'MailboxItem.Read', 'MailboxItem.ImportExport'); ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1)
            }
            return $session
        }
    }
    AfterAll { Remove-Variable -Name MbtRoute, MbtClient -Scope Global -ErrorAction SilentlyContinue }
    It 'reads the trees of adele.vance@contoso.com (ids of the configuration, no id lookup needed), shows them and records them in the store' {
        $text = Invoke-FakePhase0 @{ Sizes = $true }
        $text | Should -Match 'signed in as adele.vance@contoso.com'
        $text | Should -Match 'MBT-Fill-20261028'
        $text | Should -Match '2 folders, 3 items'
        $text | Should -Match 'folders recorded in the export store'
        $text | Should -Not -Match 'not as adele.vance@contoso.com'
        $store = Open-MbtExportStore -Path (Join-Path $script:Tmp 'phase0\export') -ReadOnly
        try {
            @($store.GetMailboxes()).Count | Should -Be 2
            ($store.GetFolderTree($script:PrimaryId))[0].Children[0].Path | Should -Be 'Top/MBT-Fill-20261028'
        } finally { $store.Dispose() }
    }
    It 'exports the largest message of the default source folder into the store, then decodes it from its pack' {
        $text = Invoke-FakePhase0 @{ SkipTree = $true; Export = $true; Largest = 1 }
        $text | Should -Match "Items of 'MBT-Fill-20261028'"
        $text | Should -Match 'items recorded in the export store \(Pending\)'
        $text | Should -Match 'stored \(key \d+, pack \d+, offset 0\)'
        $text | Should -Match "preview: 'Quarterly report'"
        $text | Should -Match 'Quarterly report'
        $text | Should -Not -Match 'differs from the one computed'
        $text | Should -Not -Match '(?m)^ {6}x {2}\S'
        ($global:MbtRoute.Requests | Where-Object { $_.Method -eq 'POST' -and $_.Url -like '*exportItems' }).Count | Should -Be 1
        (($global:MbtRoute.Requests | Where-Object { $_.Url -like '*exportItems' }).Body | ConvertFrom-Json).itemIds | Should -Be @('I2')
        Get-ChildItem (Join-Path $script:Tmp 'phase0') -Recurse -Include *.fxs, manifest.json, items.csv | Should -BeNullOrEmpty
        $store = Open-MbtExportStore -Path (Join-Path $script:Tmp 'phase0\export')
        try {
            @(Get-MbtStoreItems -Store $store).Count | Should -Be 1
            $store.GetStats().Items | Should -Be 3
        } finally { $store.Dispose() }
        $report = Get-Content -LiteralPath (Join-Path $script:Tmp 'phase0\phase0-report.json') -Raw | ConvertFrom-Json
        $report.Steps.Export.Items[0].Sha256Matches | Should -BeTrue
    }
    It 'with -WhatIf plans the import (the folder would be created) and sends nothing to the target mailbox' {
        $text = Invoke-FakePhase0 @{ SkipTree = $true; Import = $true; CreateTargetFolder = $true; WhatIf = $true }
        $text | Should -Match "WhatIf: the folder 'MBT-Import-20261028' does not exist"
        $text | Should -Match 'store   : item'
        $text | Should -Match 'WhatIf: nothing is created'
        @($global:MbtRoute.Requests | Where-Object { $_.Method -eq 'POST' }).Count | Should -Be 0
    }
    It 'refuses a missing target folder without -CreateTargetFolder, before any write' {
        { Invoke-FakePhase0 @{ SkipTree = $true; Import = $true } } | Should -Throw '*-CreateTargetFolder*'
        @($global:MbtRoute.Requests | Where-Object { $_.Method -eq 'POST' }).Count | Should -Be 0
    }
    It 'creates the target folder and imports the stored item from its pack, then verifies the round trip' {
        $text = Invoke-FakePhase0 @{ SkipTree = $true; Import = $true; CreateTargetFolder = $true; VerifyRoundTrip = $true }
        $text | Should -Match 'folder created in the Archive mailbox: MBT-Import-20261028'
        $text | Should -Match 'imported in .*new item id'
        $text | Should -Match 'target folder: 0 -> 1 items'
        $text | Should -Match 'item found in the target folder'
        $text | Should -Match 'Subject +source / re-exported: same'
        $text | Should -Not -Match 'DIFFERENT'
        $posts = @($global:MbtRoute.Requests | Where-Object Method -eq 'POST')
        ($posts | Where-Object { $_.Url -like '*/folders' }).Count | Should -Be 1
        $imp = $posts | Where-Object { $_.Url -like '*importItem*' }
        $imp.HasAuthorization | Should -BeFalse
        $body = $imp.Body | ConvertFrom-Json
        $body.FolderId | Should -Be 'NEWF'
        $body.Mode | Should -Be 'create'
        Get-Sha256Hex ([Convert]::FromBase64String($body.Data)) | Should -Be (Get-Sha256Hex $script:Fx)
        (($posts | Where-Object { $_.Url -like '*/folders' }).Body | ConvertFrom-Json).displayName | Should -Be 'MBT-Import-20261028'
        $store = Open-MbtExportStore -Path (Join-Path $script:Tmp 'phase0\export')
        try {
            $done = @($store.GetImportItems(1, [MailboxTransfer.StoreImportState]::Done, 10))
            $done.Count | Should -Be 1
            $done[0].TargetItemId | Should -Be 'NEW-ID'
            $done[0].TargetFolder | Should -Be 'NEWF'
        } finally { $store.Dispose() }
    }
}

Describe 'FastTransfer reader' {
    It 'reads fixed, variable and multi-valued properties, markers, named properties and nested items' {
        $d = Read-Fx (New-TestMessage)
        $r = $d.Read; $m = $d.Summary
        $r.Envelope | Should -Be 'None'
        $r.Error | Should -BeNullOrEmpty
        $r.Truncated | Should -BeFalse
        $byId = { param($id) $r.Elements | Where-Object { -not $_.IsMarker -and $_.Id -eq $id -and $_.Path -eq 'message' } | Select-Object -First 1 }
        (& $byId 0x0017).Value | Should -Be 1
        (& $byId 0x0017).TypeName | Should -Be 'PtypInteger32'
        (& $byId 0x0017).Name | Should -Be 'PidTagImportance'
        (& $byId 0x7001).Value | Should -Be 7                       # PtypInteger16
        (& $byId 0x7005).Value | Should -Be 1.5                     # PtypFloating64
        (& $byId 0x0E1B).Value | Should -BeTrue                     # PtypBoolean, 2 bytes
        (& $byId 0x0FF9).ValueBytes | Should -Be 16                 # PtypBinary
        (& $byId 0x7002).Value | Should -Be ([guid]'11111111-2222-3333-4444-555555555555')
        (& $byId 0x7003).Value | Should -Be @(10, 20, 30)           # PtypMultipleInteger32
        (& $byId 0x7004).Value | Should -Be @('red', 'blue')        # PtypMultipleString
        $m.Subject | Should -Be 'Quarterly report'
        $m.MessageClass | Should -Be 'IPM.Note'
        $m.SenderName | Should -Be 'Adele Vance'                          # PtypString8
        $m.SenderSmtp | Should -Be 'adele@contoso.test'
        $m.ClientSubmitTime.UtcDateTime | Should -Be ([datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc))
        $m.CreationTime.Minute | Should -Be 31
        $m.MessageSize | Should -Be 71133
        $m.HasAttachments | Should -BeTrue
        # named properties: GUID + lid, GUID + name
        $named = @($r.Elements | Where-Object { $_.IsNamed })
        $named.Count | Should -Be 2
        $named[0].NamedGuid | Should -Be ([guid]'00020329-0000-0000-c000-000000000046')
        $named[0].NamedLid | Should -Be 0x8503
        $named[0].Value | Should -Be 5
        $named[1].NamedString | Should -Be 'Keywords'
        $named[1].Value | Should -Be 'lab'
        $m.NamedProperties.Count | Should -Be 2
    }
    It 'extracts recipients, attachments (without loading them) and bodies' {
        $m = (Read-Fx (New-TestMessage)).Summary
        $m.Recipients.Count | Should -Be 2
        $m.Recipients[0].TypeName | Should -Be 'To'
        $m.Recipients[0].Smtp | Should -Be 'alice@contoso.test'
        $m.Recipients[0].DisplayName | Should -Be 'Alice'
        $m.Recipients[1].TypeName | Should -Be 'Cc'
        $m.Attachments.Count | Should -Be 2
        $m.Attachments[0].FileName | Should -Be 'report.pdf'
        $m.Attachments[0].MimeTag | Should -Be 'application/pdf'
        $m.Attachments[0].Size | Should -Be 2048
        $m.Attachments[0].DataBytes | Should -Be 2048
        $m.Attachments[0].EmbeddedMessage | Should -BeFalse
        $m.Attachments[1].DisplayName | Should -Be 'forwarded.msg'
        $m.Attachments[1].EmbeddedMessage | Should -BeTrue
        $m.Attachments[1].EmbeddedSubject | Should -Be 'Inner subject'
        $m.Subject | Should -Be 'Quarterly report'                  # the inner subject does not replace it
        $m.BodyTextBytes | Should -Be (600000 * 2 + 2)
        $m.BodyTextPreview.Length | Should -BeLessOrEqual 403
        $m.BodyTextPreview | Should -BeLike 'hello hello*'
        $m.BodyHtmlPreview | Should -Be '<html><body>Hi</body></html>'
        $m.HasCompressedRtf | Should -BeTrue
        $m.RtfBytes | Should -Be 50
    }
    It 'says whether the folder is in the stream' {
        (Read-Fx (New-TestMessage)).Summary.HasFolderInfo | Should -BeFalse
        $with = (Read-Fx (New-TestMessage -WithFolderInfo)).Summary
        $with.HasFolderInfo | Should -BeTrue
        ($with.FolderInfoProperties -join ' ') | Should -Match 'PidTagParentFolderId'
        ($with.FolderInfoProperties -join ' ') | Should -BeLike '*-863846703525003263*'
    }
    It 'reads a string with a code page type (0x84B0 = 1200 = UTF-16LE, 0xFDE9 = UTF-8)' {
        $b = [MailboxTransfer.Tests.FxBuilder]::new()
        $null = $b.CodePageStr(0x0037, 0x84B0, 'Sujet accentue: cafe', [Text.Encoding]::Unicode).CodePageStr(0x001A, 0xFDE9, 'IPM.Note', [Text.Encoding]::UTF8)
        $d = Read-Fx $b.ToArray()
        $d.Read.Error | Should -BeNullOrEmpty
        $d.Summary.Subject | Should -Be 'Sujet accentue: cafe'
        $d.Summary.MessageClass | Should -Be 'IPM.Note'
        $d.Read.Elements[0].TypeName | Should -Be 'CodePage(1200)'
    }
    It 'detects and documents the envelope in front of the stream' {
        $data = New-TestMessage
        $d = Read-Fx ([MailboxTransfer.Tests.FxBuilder]::Envelope($data))
        $d.Read.Envelope | Should -Be 'Tlv'
        @($d.Read.EnvelopeRecords).Count | Should -Be 4
        (@($d.Read.EnvelopeRecords) | ForEach-Object Type) | Should -Be @(1, 3, 4, 2)
        @($d.Read.EnvelopeRecords | Where-Object IsData).Count | Should -Be 1
        $d.Read.DataLength | Should -Be $data.Length
        $d.Read.DataOffset | Should -Be (8 + 8 + 8 + 24 + 8 + 6 + 8)
        $d.Read.Error | Should -BeNullOrEmpty
        $d.Summary.Subject | Should -Be 'Quarterly report'
        $d.Summary.Attachments.Count | Should -Be 2
    }
    It 'reads the sample of the importItem page of Microsoft Learn (envelope, importance, class cut by the sample)' {
        $sample = ConvertFrom-HexText '01 00 00 00 08 00 00 00 00 00 00 00 01 00 00 00 03 00 00 00 18 00 00 00 01 00 00 00 07 03 02 00 00 00 00 00 c0 00 00 00 00 00 00 46 00 24 00 80 04 00 00 00 06 00 00 00 14 0f d6 91 86 17 02 00 00 00 c0 7b 00 00 03 00 17 00 01 00 00 00 b0 84 1a 00 12 00 00 00 49 00 50 00 4d 00 2e'
        $d = Read-Fx $sample
        $d.Read.Envelope | Should -Be 'Tlv'
        (@($d.Read.EnvelopeRecords) | ForEach-Object Length) | Should -Be @(8, 24, 6, 31680)
        $d.Read.DataLength | Should -Be 31680
        $d.Read.Truncated | Should -BeTrue
        $d.Read.Elements.Count | Should -Be 2
        $d.Read.Elements[0].Name | Should -Be 'PidTagImportance'
        $d.Read.Elements[0].Value | Should -Be 1
        $d.Read.Elements[1].TypeName | Should -Be 'CodePage(1200)'
        $d.Read.Elements[1].Incomplete | Should -BeTrue
        $d.Read.Elements[1].ValueBytes | Should -Be 18
    }
    It 'reads the sample of the Open Specifications blog (folder properties, raw stream, clean end)' {
        $sample = ConvertFrom-HexText @'
            03 00 12 40 02 01 E1 65 00 00 00 00 02 01 E0 65 16 00 00 00 52 F6 85 EC 7D 43 2E 4A A9 60 34 50
            88 53 D9 0A 00 00 00 00 03 F5 40 00 08 30 00 D0 9E A8 81 85 D0 01 02 01 E2 65 16 00 00 00 52 F6
            85 EC 7D 43 2E 4A A9 60 34 50 88 53 D9 0A 00 00 00 00 20 7C 02 01 E3 65 17 00 00 00 16 52 F6 85
            EC 7D 43 2E 4A A9 60 34 50 88 53 D9 0A 00 00 00 00 20 7C 1F 00 01 30 0C 00 00 00 49 00 4E 00 42
            00 4F 00 58 00 00 00 14 00 49 67 01 00 00 00 00 00 03 F4 03 00 39 66 FB 03 00 00 40 00 07 30 80
            39 06 A8 81 85 D0 01 0B 00 F4 10 00 00
'@
        $d = Read-Fx $sample
        $r = $d.Read
        $r.Envelope | Should -Be 'None'
        $r.Error | Should -BeNullOrEmpty
        $r.Truncated | Should -BeFalse
        $r.MarkerCount | Should -Be 1
        $r.Elements[0].Name | Should -Be 'IncrSyncChg'
        ($r.Elements | Where-Object { -not $_.IsMarker } | ForEach-Object TagText) |
            Should -Be @('0x65E10102', '0x65E00102', '0x30080040', '0x65E20102', '0x65E30102', '0x3001001F', '0x67490014', '0x66390003', '0x30070040', '0x10F4000B')
        ($r.Elements | Where-Object Id -EQ 0x65E0).ValueBytes | Should -Be 22
        ($r.Elements | Where-Object Id -EQ 0x3001).Value | Should -Be 'INBOX'
        ($r.Elements | Where-Object Id -EQ 0x6749).Value | Should -Be -863846703525003263
        ($r.Elements | Where-Object Id -EQ 0x6639).Value | Should -Be 0x3FB
        ($r.Elements | Where-Object Id -EQ 0x3008).Value.Year | Should -Be 2015
        ($r.Elements | Where-Object Id -EQ 0x10F4).Value | Should -BeFalse
    }
    It 'never crashes on a stream cut at any length' {
        $data = [MailboxTransfer.Tests.FxBuilder]::Envelope((New-TestMessage -WithFolderInfo))
        $step = [math]::Max(1, [int]($data.Length / 400))
        foreach ($len in (@(0..60) + @(60..($data.Length - 1) | Where-Object { $_ % $step -eq 0 }))) {
            $cut = [byte[]]::new($len)
            [Array]::Copy($data, $cut, $len)
            { $null = Read-Fx $cut } | Should -Not -Throw -Because "length $len"
        }
        $plain = New-TestMessage
        foreach ($len in 0..300) {
            $cut = [byte[]]::new($len)
            [Array]::Copy($plain, $cut, $len)
            { $null = Read-Fx $cut } | Should -Not -Throw -Because "plain length $len"
        }
        (Read-Fx $plain[0..100]).Read.Truncated -or (Read-Fx $plain[0..100]).Read.Error | Should -BeTrue
    }
    It 'never crashes on garbage and keeps what it understood before an unknown type' {
        $rng = [Random]::new(7)
        foreach ($i in 1..300) {
            $junk = [byte[]]::new($rng.Next(0, 600)); $rng.NextBytes($junk)
            { $null = Read-Fx $junk } | Should -Not -Throw
        }
        $b = [MailboxTransfer.Tests.FxBuilder]::new()
        $null = $b.Str(0x0037, 'Before').Raw([byte[]](0x77, 0x07, 0x01, 0x30)).Str(0x001A, 'After')   # type 0x0777 does not exist
        $d = Read-Fx $b.ToArray()
        $d.Summary.Subject | Should -Be 'Before'
        $d.Read.Error | Should -BeLike '*Unknown property type 0x0777*'
        $d.Summary.MessageClass | Should -BeNullOrEmpty
        # an absurd multi-valued count and an absurd length do not allocate or loop
        $d2 = Read-Fx ([byte[]](0x03, 0x10, 0x01, 0x70, 0xFF, 0xFF, 0xFF, 0x7F))
        $d2.Read.Error | Should -BeLike '*absurd count*'
        $d3 = Read-Fx ([byte[]](0x1F, 0x00, 0x01, 0x30, 0xFF, 0xFF, 0xFF, 0xFF, 1, 2, 3))
        $d3.Read.Truncated | Should -BeTrue
    }
    It 'does not load a 64 MiB attachment in memory' {
        $file = Join-Path $script:Tmp 'big.fxs'
        $fs = [IO.File]::Create($file)
        try {
            $b = [MailboxTransfer.Tests.FxBuilder]::new($fs)
            $null = $b.Str(0x0037, 'Big one').Marker(0x40000003).Str(0x3707, 'big.bin').Int32(0x0E20, 67108864).BinLarge(0x3701, 64MB).Marker(0x400E0003)
            $null = $b.Marker(0x40000003).Str(0x3707, 'small.txt').Bin(0x3701, [byte[]](1..10)).Marker(0x400E0003)
        } finally { $fs.Dispose() }
        [GC]::Collect()
        $before = [GC]::GetAllocatedBytesForCurrentThread()
        $read = [MailboxTransfer.FxReader]::ReadFile($file)
        $allocated = [GC]::GetAllocatedBytesForCurrentThread() - $before
        $m = [MailboxTransfer.FxMessage]::Summarize($read)
        $read.Error | Should -BeNullOrEmpty
        $m.Attachments.Count | Should -Be 2
        $m.Attachments[0].DataBytes | Should -Be 64MB
        $m.Attachments[1].FileName | Should -Be 'small.txt'
        $m.Subject | Should -Be 'Big one'
        $allocated | Should -BeLessThan (8MB) -Because 'the attachment is read and discarded, only its first bytes are kept'
        ($read.Elements | Where-Object { $_.Id -eq 0x3701 -and $_.ValueBytes -eq 64MB }).Preview.Length | Should -BeLessOrEqual 16384
    }
    It 'concatenates the records of type 2 of the envelope: a record may end in the middle of a property' {
        $data = New-TestMessage
        $whole = Read-Fx ([MailboxTransfer.Tests.FxBuilder]::Envelope($data))
        foreach ($chunk in 1, 7, 31680, 4096) {
            $d = Read-Fx ([MailboxTransfer.Tests.FxBuilder]::Envelope($data, $chunk, $false))
            $d.Read.Envelope | Should -Be 'Tlv'
            $d.Read.Error | Should -BeNullOrEmpty -Because "chunk $chunk"
            $d.Read.Truncated | Should -BeFalse -Because "chunk $chunk"
            $d.Read.DataRecords | Should -Be ([math]::Ceiling($data.Length / $chunk))
            $d.Read.DataLength | Should -Be $data.Length
            $d.Read.PropertyCount | Should -Be $whole.Read.PropertyCount
            $d.Read.MarkerCount | Should -Be $whole.Read.MarkerCount
            $d.Summary.Subject | Should -Be 'Quarterly report'
            $d.Summary.Recipients.Count | Should -Be 2
            $d.Summary.Attachments.Count | Should -Be 2
            $d.Summary.Attachments[1].EmbeddedSubject | Should -Be 'Inner subject'
        }
        $d = Read-Fx ([MailboxTransfer.Tests.FxBuilder]::Envelope($data, 100, $false))
        @($d.Read.EnvelopeRecords).Count | Should -BeLessOrEqual 64 -Because 'only the first records are kept'
        $d.Read.EnvelopeRecordCount | Should -Be ($d.Read.DataRecords + 2)
    }
    It 'reports a truncation when the last record of the envelope is shorter than announced, and when the file ends between two records' {
        $full = [MailboxTransfer.Tests.FxBuilder]::Envelope((New-TestMessage), 5000, $false)
        $cut = [byte[]]::new($full.Length - 100); [Array]::Copy($full, $cut, $cut.Length)
        (Read-Fx $cut).Read.Truncated | Should -BeTrue
        (Read-Fx $full).Read.Truncated | Should -BeFalse
    }
    It 'reads a stream built like a real export: envelope in 31,680-byte records, extended properties, recipients, a large attachment, no warning' {
        $data = Join-Path $script:Tmp 'real-data.fxs'
        $file = Join-Path $script:Tmp 'real-shape.fxs'
        $fs = [IO.File]::Create($data)
        try { Write-RealShapeMessage -Stream $fs -AttachBytes 40MB } finally { $fs.Dispose() }
        $in = [IO.File]::OpenRead($data); $out = [IO.File]::Create($file)
        try { [MailboxTransfer.Tests.FxBuilder]::WriteEnvelope($out, $in, 31680, $false) } finally { $in.Dispose(); $out.Dispose() }
        [GC]::Collect()
        $before = [GC]::GetAllocatedBytesForCurrentThread()
        $read = [MailboxTransfer.FxReader]::ReadFile($file)
        $allocated = [GC]::GetAllocatedBytesForCurrentThread() - $before
        $m = [MailboxTransfer.FxMessage]::Summarize($read)
        $read.Error | Should -BeNullOrEmpty
        $read.Truncated | Should -BeFalse
        $read.DataRecords | Should -BeGreaterThan 1000
        $read.MarkerCount | Should -Be 6      # FXDelProp x2, StartRecip, EndToRecip, NewAttach, EndAttach
        ($read.Elements | Where-Object Extended).Count | Should -Be 6
        ($read.Elements | Where-Object { $_.Extended -and $_.Id -eq 0x1714 }).Value | Should -Be 2
        ($read.Elements | Where-Object { $_.Extended -and $_.Id -eq 0x1714 }).ExtendedFlags | Should -Be 0x0200
        ($read.Elements | Where-Object { $_.Id -eq 0x010B }).TagText | Should -Be '0x010B0102' -Because 'the type is shown without the 0x4000 bit'
        ($read.Elements | Where-Object { $_.Id -eq 0x8EB2 }).NamedString | Should -Be 'FocusedInboxSubjectEmbeddings'
        $m.Subject | Should -Be '[lot1-00001] Export - Alpha'
        $m.MessageClass | Should -Be 'IPM.Note'
        $m.BodyTextPreview | Should -BeLike 'Bonjour*'
        $m.Recipients.Count | Should -Be 1
        $m.Recipients[0].Smtp | Should -Be 'alice@lab.example'
        $m.Attachments.Count | Should -Be 1
        $m.Attachments[0].FileName | Should -BeLike '*.bin'
        $m.Attachments[0].DataBytes | Should -Be 40MB
        $m.Attachments[0].Size | Should -Be (40MB + 177)
        $m.Attachments[0].DataTruncated | Should -BeFalse
        $m.Attachments[0].EmbeddedMessage | Should -BeFalse
        $allocated | Should -BeLessThan (8MB) -Because 'the attachment is skipped, never loaded'
        Remove-Item -LiteralPath $data, $file -Force
    }
}

Describe 'Grant-LabConsent.ps1 (against a fake tenant)' {
    BeforeAll {
        $script:GrantScript = Join-Path $script:RepoRoot 'tools\Grant-LabConsent.ps1'
        $script:LabApp = '1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        $script:adele = '1a2b3c05-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
        function New-FakeTenant {
            $app = '{"id":"app-obj","appId":"1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d","displayName":"MBT lab app","signInAudience":"AzureADMyOrg",
                "requiredResourceAccess":[
                    {"resourceAppId":"00000003-0000-0000-c000-000000000000","resourceAccess":[{"id":"u1","type":"Scope"}]},
                    {"resourceAppId":"other-resource","resourceAccess":[{"id":"x1","type":"Role"}]}],
                "publicClient":{"redirectUris":["https://login.microsoftonline.com/common/oauth2/nativeclient"]}}' | ConvertFrom-Json
            $graph = '{"id":"graph-sp","displayName":"Microsoft Graph","oauth2PermissionScopes":[
                {"id":"u1","value":"User.Read"},{"id":"s1","value":"MailboxFolder.Read"},{"id":"s2","value":"MailboxFolder.ReadWrite"},
                {"id":"s3","value":"MailboxItem.Read"},{"id":"s4","value":"MailboxItem.ImportExport"},{"id":"o1","value":"openid"},{"id":"o2","value":"profile"},{"id":"o3","value":"offline_access"}]}' | ConvertFrom-Json
            $sp = '{"id":"app-sp","appId":"1a2b3c02-5e7f-4a1b-9c2d-3e4f5a6b7c8d","displayName":"MBT lab app"}' | ConvertFrom-Json
            return @{ App = $app; Graph = $graph; Sp = $sp; Grants = [System.Collections.Generic.List[object]]::new(); Writes = [System.Collections.Generic.List[string]]::new(); UserLookup = $true; Lookups = [System.Collections.Generic.List[string]]::new() }
        }
    }
    BeforeEach {
        $global:MbtFake = New-FakeTenant
        Mock Connect-MgGraph { }
        Mock Get-MgContext { [pscustomobject]@{ Account = 'admin@lab.test'; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; Scopes = @('Application.ReadWrite.All', 'DelegatedPermissionGrant.ReadWrite.All', 'User.ReadBasic.All') } }
        Mock Invoke-MgGraphRequest {
            $f = $global:MbtFake
            $obj = if ($Body) { $Body | ConvertFrom-Json } else { $null }
            switch -Wildcard ("$Method $Uri") {
                'GET */users/*' {
                    $f.Lookups.Add($Uri)
                    if (-not $f.UserLookup) { throw 'Insufficient privileges to complete the operation.' }
                    return [pscustomobject]@{ id = '1a2b3c05-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; userPrincipalName = 'adele.vance@contoso.com' }
                }
                'GET */applications?*' { return [pscustomobject]@{ value = @($f.App) } }
                "GET */servicePrincipals?*appId eq '1a2b3c02*" { return [pscustomobject]@{ value = @(if ($f.Sp) { $f.Sp }) } }
                "GET */servicePrincipals?*appId eq '00000003*" { return [pscustomobject]@{ value = @($f.Graph) } }
                'GET */oauth2PermissionGrants?*' { return [pscustomobject]@{ value = $f.Grants.ToArray() } }
                'PATCH */applications/app-obj' {
                    $f.Writes.Add("PATCH application: $($obj.PSObject.Properties.Name -join ',')")
                    if ($obj.PSObject.Properties['requiredResourceAccess']) { $f.App.requiredResourceAccess = @($obj.requiredResourceAccess) }
                    if ($obj.PSObject.Properties['publicClient']) { $f.App.publicClient = $obj.publicClient }
                    if ($obj.PSObject.Properties['isFallbackPublicClient']) { $f.App | Add-Member -NotePropertyName isFallbackPublicClient -NotePropertyValue $obj.isFallbackPublicClient -Force }
                    return $null
                }
                'POST */servicePrincipals' { $f.Writes.Add('POST servicePrincipal'); $f.Sp = '{"id":"app-sp","appId":"x","displayName":"new"}' | ConvertFrom-Json; return $f.Sp }
                'POST */oauth2PermissionGrants' {
                    $f.Writes.Add("POST grant: $($obj.consentType) $($obj.principalId)")
                    $f.Grants.Add([pscustomobject]@{ id = 'grant-new'; clientId = $obj.clientId; consentType = $obj.consentType; principalId = $obj.principalId; resourceId = $obj.resourceId; scope = $obj.scope })
                    return $null
                }
                'PATCH */oauth2PermissionGrants/*' {
                    $id = ($Uri -split '/')[-1]
                    $f.Writes.Add("PATCH grant $id")
                    ($f.Grants | Where-Object id -EQ $id).scope = $obj.scope
                    return $null
                }
                default { throw "unexpected call: $Method $Uri" }
            }
        }
    }

    It 'adds the permissions, the redirect URI and the consent of one user, keeping everything else' {
        $out = & $script:GrantScript *>&1 | Out-String
        $graphAccess = ($global:MbtFake.App.requiredResourceAccess | Where-Object resourceAppId -EQ '00000003-0000-0000-c000-000000000000').resourceAccess
        ($graphAccess | ForEach-Object id | Sort-Object) | Should -Be @('s1', 's2', 's3', 's4', 'u1')
        ($graphAccess | ForEach-Object type | Sort-Object -Unique) | Should -Be @('Scope')
        ($global:MbtFake.App.requiredResourceAccess | Where-Object resourceAppId -EQ 'other-resource').resourceAccess[0].id | Should -Be 'x1'
        @($global:MbtFake.App.publicClient.redirectUris) | Should -Be @('https://login.microsoftonline.com/common/oauth2/nativeclient', 'http://localhost')
        $global:MbtFake.Grants.Count | Should -Be 1
        $grant = $global:MbtFake.Grants[0]
        $grant.consentType | Should -Be 'Principal'
        $grant.principalId | Should -Be $script:adele
        $grant.resourceId | Should -Be 'graph-sp'
        ($grant.scope -split ' ' | Sort-Object) | Should -Be @('MailboxFolder.Read', 'MailboxFolder.ReadWrite', 'MailboxItem.ImportExport', 'MailboxItem.Read', 'offline_access', 'openid', 'profile')
        $out | Should -Match 'State BEFORE'
        $out | Should -Match 'State AFTER'
        $out | Should -Match 'OK: the application declares'
        $global:MbtFake.App.isFallbackPublicClient | Should -BeTrue -Because 'the device code flow needs "Allow public client flows"'
        $global:MbtFake.Writes.Count | Should -Be 2                    # one PATCH of the application, one POST of the grant
    }
    It 'resolves the user principal name of the configuration (adele.vance@contoso.com) to its object id' {
        $null = & $script:GrantScript *>&1
        $global:MbtFake.Lookups.Count | Should -Be 1
        $global:MbtFake.Lookups[0] | Should -BeLike '*/users/adele.vance%40contoso.com?*'
        $global:MbtFake.Grants[0].principalId | Should -Be $script:adele
    }
    It 'takes the object id of -PrincipalObjectId without any lookup' {
        $null = & $script:GrantScript -PrincipalObjectId '99999999-0000-0000-0000-000000000001' *>&1
        $global:MbtFake.Lookups.Count | Should -Be 0
        $global:MbtFake.Grants[0].principalId | Should -Be '99999999-0000-0000-0000-000000000001'
    }
    It 'falls back to the object id of the configuration for the configured user when the lookup is refused, and asks for -PrincipalObjectId for anyone else' {
        $global:MbtFake.UserLookup = $false
        $null = & $script:GrantScript *>&1
        $global:MbtFake.Grants[0].principalId | Should -Be $script:adele
        $global:MbtFake.Grants.Clear()
        { & $script:GrantScript -UserPrincipalName 'someone@lab.test' *>&1 | Out-Null } | Should -Throw '*-PrincipalObjectId*'
        $global:MbtFake.Grants.Count | Should -Be 0
    }
    It 'leaves "Allow public client flows" alone with -NoPublicClientFlows' {
        $null = & $script:GrantScript -NoPublicClientFlows *>&1
        $global:MbtFake.App.PSObject.Properties['isFallbackPublicClient'] | Should -BeNullOrEmpty
    }
    It 'with -DeviceCode gets the administrator token from the module (Graph Command Line Tools client) and gives it to Connect-MgGraph -AccessToken' {
        $global:MbtFake.TokenCalls = [System.Collections.Generic.List[object]]::new()
        $global:MbtFake.ContextCalls = 0
        Mock Get-MgContext { if ($global:MbtFake.ContextCalls++ -eq 0) { $null } else { [pscustomobject]@{ Account = $null; TenantId = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'; Scopes = @() } } }
        Mock Get-MbtAccessToken {
            $global:MbtFake.TokenCalls.Add(@{ ClientId = $ClientId; Scopes = $Scopes; DeviceCode = [bool]$DeviceCode; File = $DeviceCodeFile })
            [pscustomobject]@{ AccessToken = 'admin-token'; Account = 'admin@lab.test'; ExpiresOn = [DateTimeOffset]::UtcNow.AddHours(1) }
        }
        $out = & $script:GrantScript -DeviceCode *>&1 | Out-String
        $global:MbtFake.TokenCalls.Count | Should -Be 1
        $call = $global:MbtFake.TokenCalls[0]
        $call.ClientId | Should -Be '14d82eec-204b-4c2f-b7e8-296a70dab67e'
        $call.DeviceCode | Should -BeTrue
        $call.File | Should -BeLike '*\out\device-code.txt'
        $call.Scopes | Should -Contain 'https://graph.microsoft.com/Application.ReadWrite.All'
        $call.Scopes | Should -Contain 'https://graph.microsoft.com/DelegatedPermissionGrant.ReadWrite.All'
        Should -Invoke Connect-MgGraph -Times 1 -Exactly -ParameterFilter { $AccessToken -and -not $Scopes }
        $out | Should -Match 'Signed in as admin@lab.test'
        $global:MbtFake.Grants.Count | Should -Be 1
    }
    It 'is idempotent: a second run writes nothing' {
        $null = & $script:GrantScript *>&1
        $global:MbtFake.Writes.Clear()
        $out = & $script:GrantScript *>&1 | Out-String
        $global:MbtFake.Writes.Count | Should -Be 0
        $out | Should -Match 'nothing to change'
        $global:MbtFake.Grants.Count | Should -Be 1
    }
    It 'with -WhatIf shows the plan and changes nothing' {
        $out = & $script:GrantScript -WhatIf *>&1 | Out-String
        $global:MbtFake.Writes.Count | Should -Be 0
        $global:MbtFake.Grants.Count | Should -Be 0
        $out | Should -Match 'permissions to declare'
        $out | Should -Match 'WhatIf'
    }
    It 'completes an existing consent of the user and never touches an AllPrincipals consent' {
        $global:MbtFake.Grants.Add([pscustomobject]@{ id = 'g-adele'; clientId = 'app-sp'; consentType = 'Principal'; principalId = $script:adele; resourceId = 'graph-sp'; scope = 'User.Read MailboxItem.Read' })
        $global:MbtFake.Grants.Add([pscustomobject]@{ id = 'g-all'; clientId = 'app-sp'; consentType = 'AllPrincipals'; principalId = $null; resourceId = 'graph-sp'; scope = 'User.Read' })
        $null = & $script:GrantScript *>&1
        $global:MbtFake.Grants.Count | Should -Be 2
        (($global:MbtFake.Grants | Where-Object id -EQ 'g-adele').scope -split ' ') | Should -Contain 'User.Read'
        (($global:MbtFake.Grants | Where-Object id -EQ 'g-adele').scope -split ' ') | Should -Contain 'MailboxItem.ImportExport'
        ($global:MbtFake.Grants | Where-Object id -EQ 'g-all').scope | Should -Be 'User.Read'
        ($global:MbtFake.Writes | Where-Object { $_ -like 'PATCH grant*' }) | Should -Be @('PATCH grant g-adele')
    }
    It 'refuses a permission that Microsoft Graph does not publish, before any write' {
        { & $script:GrantScript -Permissions 'MailboxFolder.Read', 'MailboxThing.Nope' *>&1 | Out-Null } | Should -Throw '*do not exist*'
        $global:MbtFake.Writes.Count | Should -Be 0
    }
    AfterAll { Remove-Variable -Name MbtFake -Scope Global -ErrorAction SilentlyContinue }
}

Describe 'Engine left in the process by an older build' {
    # A compiled engine cannot be unloaded: a reloaded module must refuse it at import, not fail at the end of a run.
    BeforeAll {
        $script:StaleRoot = Join-Path $script:Tmp 'stale-package'
        $null = New-Item -ItemType Directory -Path $script:StaleRoot
        Copy-Item -Path (Join-Path $script:Root '*') -Destination $script:StaleRoot -Recurse -Exclude 'config', 'docs', 'PackageIntegrity.psd1'
        function Invoke-StaleChild {
            param([Parameter(Mandatory)][string]$Body)
            $psd1 = Join-Path $script:StaleRoot 'MailboxTransfer.psd1'
            $script = "`$ErrorActionPreference = 'Stop'; Import-Module '$psd1'; $Body"
            $out = & (Get-Process -Id $PID).Path -NoProfile -NonInteractive -Command $script 2>&1 | Out-String
            [pscustomobject]@{ Output = $out; ExitCode = $LASTEXITCODE }
        }
    }
    It 'reloading the module with unchanged sources reuses the loaded engine' {
        $r = Invoke-StaleChild "Initialize-MbtEngine; Import-Module '$(Join-Path $script:StaleRoot 'MailboxTransfer.psd1')' -Force; Initialize-MbtEngine; 'RELOAD-OK'"
        $r.Output | Should -Match 'RELOAD-OK'
        $r.ExitCode | Should -Be 0
    }
    It 'refuses at module import an engine compiled from other sources, asking for a new session' {
        $graph = Join-Path $script:StaleRoot 'src\Engine.Graph.cs'
        $r = Invoke-StaleChild "Initialize-MbtEngine; Add-Content -LiteralPath '$graph' -Value '// updated'; Import-Module '$(Join-Path $script:StaleRoot 'MailboxTransfer.psd1')' -Force; 'NOT-REACHED'"
        $r.Output | Should -Not -Match 'NOT-REACHED'
        $r.Output | Should -Match 'older build'
        $r.Output | Should -Match 'open a new PowerShell session'
        $r.ExitCode | Should -Not -Be 0
        # restore for the following test
        Set-Content -LiteralPath $graph -Value (Get-Content -LiteralPath (Join-Path $script:Root 'src\Engine.Graph.cs') -Raw) -NoNewline
    }
    It 'refuses a changed Auth engine the same way' -Skip:(-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        $auth = Join-Path $script:StaleRoot 'src\Engine.Auth.cs'
        $r = Invoke-StaleChild "Initialize-MbtAuthEngine; Add-Content -LiteralPath '$auth' -Value '// updated'; Import-Module '$(Join-Path $script:StaleRoot 'MailboxTransfer.psd1')' -Force; 'NOT-REACHED'"
        $r.Output | Should -Not -Match 'NOT-REACHED'
        $r.Output | Should -Match 'auth engine .* older build'
        $r.Output | Should -Match 'open a new PowerShell session'
    }
}
