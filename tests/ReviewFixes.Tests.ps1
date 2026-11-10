#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - regression tests of the findings of the independent review (probes in the review folder, one Describe per finding).
    Nothing here reaches the network or a mailbox.

    Run:  Invoke-Pester -Path .\tests\ReviewFixes.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-review-tests-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp

    function New-TestClient {
        param([Parameter(Mandatory)]$Handler, [int]$MaxAttempts = 6, [double]$DelayScale = 0)
        $o = [MailboxTransfer.GraphOptions]::new()
        $o.DelayScale = $DelayScale
        $o.UseJitter = $false
        $o.MaxAttempts = $MaxAttempts
        return [MailboxTransfer.GraphClient]::new([Func[string]] { 'test-token' }, $o, $Handler)
    }

    if (-not ('ReviewHangingStream' -as [type])) {
        Add-Type -ReferencedAssemblies 'System.Net.Http', 'System.Runtime', 'System.Threading.Tasks', 'System.Threading', 'netstandard', 'System.Net.Primitives', 'System.Memory' -TypeDefinition @'
using System; using System.IO; using System.Net; using System.Net.Http; using System.Threading; using System.Threading.Tasks;
// A body that delivers "ab" and then goes silent. IgnoresToken: the read does not look at its token (only Dispose releases it).
public class ReviewHangingStream : Stream {
    public bool IgnoresToken; int _sent; readonly ManualResetEventSlim _released = new ManualResetEventSlim(false);
    public override int Read(byte[] b, int o, int c) { if (_sent == 0) { _sent = 2; b[o] = 97; b[o + 1] = 98; return 2; } _released.Wait(); throw new ObjectDisposedException("hang"); }
    public override async Task<int> ReadAsync(byte[] b, int o, int c, CancellationToken ct) {
        if (_sent == 0) { _sent = 2; b[o] = 97; b[o + 1] = 98; return 2; }
        if (IgnoresToken) { await Task.Run(() => _released.Wait()); throw new ObjectDisposedException("hang"); }
        await Task.Delay(Timeout.Infinite, ct); return 0;
    }
    protected override void Dispose(bool d) { _released.Set(); base.Dispose(d); }
    public override bool CanRead { get { return true; } } public override bool CanSeek { get { return false; } } public override bool CanWrite { get { return false; } }
    public override long Length { get { throw new NotSupportedException(); } } public override long Position { get { throw new NotSupportedException(); } set { throw new NotSupportedException(); } }
    public override void Flush() { } public override long Seek(long a, SeekOrigin b) { throw new NotSupportedException(); } public override void SetLength(long v) { throw new NotSupportedException(); }
    public override void Write(byte[] b, int o, int c) { throw new NotSupportedException(); }
}
public class ReviewHangingHandler : HttpMessageHandler {
    public bool IgnoresToken;
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage r, CancellationToken ct) {
        var m = new HttpResponseMessage(HttpStatusCode.OK); m.Content = new StreamContent(new ReviewHangingStream { IgnoresToken = IgnoresToken }); return Task.FromResult(m);
    }
}
'@
    }
}

AfterAll {
    Remove-Item -LiteralPath $script:Tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Describe 'Review 5: a stop is honoured during the retry waits' {
    It 'a Retry-After of 30 s ends within a few seconds once the token is cancelled' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.Enqueue(429, '{}', '30')
        $client = New-TestClient $h -DelayScale 1
        $cts = [Threading.CancellationTokenSource]::new(500)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $thrown = $null
        try { $null = $client.Send('GET', 'https://graph.microsoft.com/v1.0/x', $null, $true, $true, [Net.Http.HttpCompletionOption]::ResponseContentRead, $cts.Token) } catch { $thrown = $_.Exception }
        $sw.Stop()
        $thrown | Should -Not -BeNullOrEmpty
        ($thrown.InnerException ?? $thrown) | Should -BeOfType ([OperationCanceledException])
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 5
    }
    It 'the backoff after a network error is interrupted too' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new()
        $h.EnqueueFailure()
        $o = [MailboxTransfer.GraphOptions]::new(); $o.UseJitter = $false; $o.BaseBackoffSeconds = 30
        $client = [MailboxTransfer.GraphClient]::new([Func[string]] { 't' }, $o, $h)
        $cts = [Threading.CancellationTokenSource]::new(500)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        try { $null = $client.Send('GET', 'https://graph.microsoft.com/v1.0/x', $null, $true, $true, [Net.Http.HttpCompletionOption]::ResponseContentRead, $cts.Token) } catch { $thrown = $_.Exception }
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 5
        ($thrown.InnerException ?? $thrown) | Should -BeOfType ([OperationCanceledException])
    }
}

Describe 'Review 8: a streamed response does not stall' {
    It 'a body that goes silent ends with an inactivity error (<Mode>)' -ForEach @(@{ Mode = 'read honours its token'; Ignores = $false }, @{ Mode = 'read ignores its token'; Ignores = $true }) {
        $h = [ReviewHangingHandler]::new(); $h.IgnoresToken = $Ignores
        $o = [MailboxTransfer.GraphOptions]::new(); $o.ReadInactivityTimeout = [TimeSpan]::FromMilliseconds(400)
        $client = [MailboxTransfer.GraphClient]::new([Func[string]] { 't' }, $o, $h)
        $resp = $client.Send('GET', 'https://graph.microsoft.com/v1.0/x/$value', $null, $true, $true, [Net.Http.HttpCompletionOption]::ResponseHeadersRead)
        $s = $client.ReadBody($resp)
        $buf = [byte[]]::new(16)
        $s.Read($buf, 0, 16) | Should -Be 2
        $sw = [Diagnostics.Stopwatch]::StartNew()
        { $s.Read($buf, 0, 16) } | Should -Throw '*No data received*'
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 5
        $s.Dispose(); $resp.Dispose()
    }
    It 'a cancellation releases a silent read (<Mode>)' -ForEach @(@{ Mode = 'read honours its token'; Ignores = $false }, @{ Mode = 'read ignores its token'; Ignores = $true }) {
        $h = [ReviewHangingHandler]::new(); $h.IgnoresToken = $Ignores
        $client = [MailboxTransfer.GraphClient]::new([Func[string]] { 't' }, [MailboxTransfer.GraphOptions]::new(), $h)
        $cts = [Threading.CancellationTokenSource]::new()
        $resp = $client.Send('GET', 'https://graph.microsoft.com/v1.0/x/$value', $null, $true, $true, [Net.Http.HttpCompletionOption]::ResponseHeadersRead, $cts.Token)
        $s = $client.ReadBody($resp, $cts.Token)
        $buf = [byte[]]::new(16)
        $null = $s.Read($buf, 0, 16)
        $cts.CancelAfter(300)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $thrown = $null
        try { $null = $s.Read($buf, 0, 16) } catch { $thrown = $_.Exception }
        $sw.Elapsed.TotalSeconds | Should -BeLessThan 5
        ($thrown.InnerException ?? $thrown) | Should -BeOfType ([OperationCanceledException])
        $resp.Dispose()
    }
    It 'the export response reader reads through the guarded stream' {
        $json = '{"value":[{"itemId":"A","changeKey":"C","data":"AQID"}]}'
        $dir = Join-Path $script:Tmp 'reader'
        $e = [MailboxTransfer.ExportResponseReader]::Read([MailboxTransfer.GuardedReadStream]::new([IO.MemoryStream]::new([Text.Encoding]::UTF8.GetBytes($json)), [Threading.CancellationToken]::None, [TimeSpan]::FromSeconds(5)), $dir)
        $e.Count | Should -Be 1
        [IO.File]::ReadAllBytes($e[0].FilePath) | Should -Be ([byte[]](1, 2, 3))
    }
}

Describe 'Review 16: the request timeout follows the size of the body' {
    It 'is the request timeout for a small body and grows with the body' {
        $o = [MailboxTransfer.GraphOptions]::new()
        $o.RequestTimeout = [TimeSpan]::FromMinutes(15); $o.MinUploadBytesPerSecond = 100000
        $c = [MailboxTransfer.GraphClient]::new([Func[string]] { 't' }, $o)
        $c.TimeoutFor(0).TotalSeconds | Should -Be 900
        $c.TimeoutFor(1000000000).TotalSeconds | Should -Be (900 + 10000)
        $o.MinUploadBytesPerSecond = 0
        $c.TimeoutFor(1000000000).TotalSeconds | Should -Be 900
    }
    It 'has a default minimum throughput' {
        [MailboxTransfer.GraphOptions]::new().MinUploadBytesPerSecond | Should -BeGreaterThan 0
    }
}

Describe 'Review 18: the whole query string of a pre-authenticated URL is redacted' {
    It '<Url>' -ForEach @(
        @{ Url = "https://outlook.office365.com/api/gv1.0/Mailboxes('x')/importItem?authtoken=SECRET"; Expected = "https://outlook.office365.com/api/gv1.0/Mailboxes('x')/importItem?***" }
        @{ Url = "https://outlook.office365.com/api/gv1.0/x?a=1&AuthToken=SECRET&b=2"; Expected = 'https://outlook.office365.com/api/gv1.0/x?***' }
        @{ Url = 'https://outlook.office365.com/import?sig=SECRET&exp=1'; Expected = 'https://outlook.office365.com/import?***' }
        @{ Url = 'https://graph.microsoft.com/v1.0/me/messages?$top=5'; Expected = 'https://graph.microsoft.com/v1.0/me/messages?$top=5' }
        @{ Url = 'https://graph.microsoft.com/v1.0/x?authtoken=SECRET'; Expected = 'https://graph.microsoft.com/v1.0/x?***' }
        @{ Url = 'https://outlook.office365.com/import'; Expected = 'https://outlook.office365.com/import' }
    ) {
        [MailboxTransfer.GraphClient]::Redact($Url) | Should -Be $Expected
        [MailboxTransfer.GraphClient]::Redact($null) | Should -BeNullOrEmpty
    }
    It 'an error on the import URL shows no part of the query' {
        $h = [MailboxTransfer.Tests.FakeHandler]::new(); $h.Enqueue(400, '{"error":{"code":"X","message":"bad"}}')
        $api = [MailboxTransfer.MailboxApi]::new((New-TestClient $h))
        $file = Join-Path $script:Tmp 'i.fxs'; [IO.File]::WriteAllBytes($file, [byte[]](1, 2, 3))
        $session = [MailboxTransfer.ImportSession]@{ MailboxId = 'MBX:a@b'; ImportUrl = "https://outlook.office365.com/api/gv1.0/x/importItem?sig=SECRET123&authtoken=SECRET456" }
        $m = $null; $u = $null
        try { $null = $api.ImportItem($session, 'F', $file) } catch { $e = $_.Exception.InnerException ?? $_.Exception; $m = $e.Message; $u = $e.Url }
        $m | Should -Not -BeNullOrEmpty
        "$m $u" | Should -Not -BeLike '*SECRET*'
    }
}

Describe 'Review 6: the PST export deletes only the files it wrote' {
    BeforeAll {
        $script:PG = '11111111-2222-3333-4444-555555555555'
        $script:P = "MBX:$($script:PG)@cccccccc-dddd-eeee-ffff-000000000000"
        function New-SplitStore {
            param([string]$Name)
            $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp $Name), $Name, 'tests')
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:P; Kind = 'Primary'; Upn = 'u@x.test'; DisplayName = 'Primary' })
            $fid = [MailboxTransfer.Tests.FakeIds]::Folder($script:PG, 1)
            [void]$store.UpsertFolders([MailboxTransfer.StoreFolder[]]@([MailboxTransfer.StoreFolder]@{ MailboxId = $script:P; FolderId = $fid; DisplayName = 'Inbox'; FolderType = 'IPF.Note'; SourceItemCount = 0 }))
            $P = [MailboxTransfer.FxProperty]
            foreach ($n in 1..8) {
                $m = [MailboxTransfer.FxMessage]::new()
                $m.Properties.Add($P::FromString(0x001A, 'IPM.Note')); $m.Properties.Add($P::FromString(0x0037, "Item $n"))
                $a = [MailboxTransfer.FxAttachment]::new(); $b = [byte[]]::new(1200000); [Random]::new($n).NextBytes($b)
                $a.Properties.Add($P::FromInt32(0x3705, 1)); $a.Properties.Add($P::FromString(0x3707, 'd.bin')); $a.Properties.Add($P::FromBinary(0x3701, $b)); $m.Attachments.Add($a)
                $bytes = [MailboxTransfer.FxCodec]::EncodePayload($m)
                $id = "I-$n"
                [void]$store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@([MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:P; ItemId = $id; FolderId = $fid; Size = $bytes.Length; ItemType = 'IPM.Note' }))
                [void]$store.WriteItem($store.FindItemKey($script:P, $id), [IO.MemoryStream]::new($bytes), [Threading.CancellationToken]::None)
            }
            return $store
        }
        function New-ExportOptions([string]$Dest, [switch]$Force) {
            $o = [MailboxTransfer.PstExportOptions]::new(); $o.Destination = $Dest; $o.MaxPartBytes = 5MB; $o.Check = $false; $o.Force = [bool]$Force
            return $o
        }
        function Add-UserFiles([string]$Dir) {
            foreach ($n in 'Archive-2019.pst', 'Archive-2020.pst', 'Archive-old.pst', 'Archive-12345.pst', 'Archive.bak.pst') { [IO.File]::WriteAllBytes((Join-Path $Dir $n), [byte[]](1, 2, 3)) }
        }
    }

    It '-Force replaces the parts of the manifest and keeps the user files that only look like parts' {
        $store = New-SplitStore 'six1'
        try {
            $dir = Join-Path $script:Tmp 'six1\out'; $dest = Join-Path $dir 'Archive.pst'
            $r1 = [MailboxTransfer.PstExporter]::Run($store, (New-ExportOptions $dest), $null, [Threading.CancellationToken]::None)
            $r1.Ok | Should -BeTrue
            $r1.Parts.Count | Should -BeGreaterThan 1
            Add-UserFiles $dir
            $r2 = [MailboxTransfer.PstExporter]::Run($store, (New-ExportOptions $dest -Force), $null, [Threading.CancellationToken]::None)
            $r2.Ok | Should -BeTrue
            foreach ($n in 'Archive-2019.pst', 'Archive-2020.pst', 'Archive-old.pst', 'Archive-12345.pst', 'Archive.bak.pst') {
                Test-Path -LiteralPath (Join-Path $dir $n) | Should -BeTrue -Because "$n is not a file of ours"
                (Get-Item -LiteralPath (Join-Path $dir $n)).Length | Should -Be 3
            }
        } finally { $store.Dispose() }
    }

    It 'a resume after a cancelled run keeps the user files too, and the cleanup of a stale part still happens' {
        $store = New-SplitStore 'six2'
        try {
            $dir = Join-Path $script:Tmp 'six2\out'; $dest = Join-Path $dir 'Archive.pst'
            $cts = [Threading.CancellationTokenSource]::new(); $state = @{ Done = $false }
            $sink = [MailboxTransfer.PstExportProgressSink]::new([Action[MailboxTransfer.PstExportProgress]]{
                    param($e) if ($e.Phase -eq 'write' -and $e.PartIndex -eq 2 -and $e.ItemsDone -ge 4 -and -not $state.Done) { $state.Done = $true; $cts.Cancel() }
                }.GetNewClosure())
            $first = [MailboxTransfer.PstExporter]::Run($store, (New-ExportOptions $dest), $sink, $cts.Token)
            $first.Cancelled | Should -BeTrue
            Add-UserFiles $dir
            [IO.File]::WriteAllBytes((Join-Path $dir 'Archive-002.pst'), [byte[]](9, 9, 9))      # the half-written part of the killed run
            $second = [MailboxTransfer.PstExporter]::Run($store, (New-ExportOptions $dest), $null, [Threading.CancellationToken]::None)
            $second.Ok | Should -BeTrue
            (Get-Item -LiteralPath (Join-Path $dir 'Archive-002.pst')).Length | Should -BeGreaterThan 3
            foreach ($n in 'Archive-2019.pst', 'Archive-2020.pst', 'Archive-old.pst', 'Archive-12345.pst', 'Archive.bak.pst') {
                (Get-Item -LiteralPath (Join-Path $dir $n)).Length | Should -Be 3 -Because "$n is not a file of ours"
            }
        } finally { $store.Dispose() }
    }

    It 'a new destination next to files such as Archive-2019.pst is not refused for them' {
        $store = New-SplitStore 'six3'
        try {
            $dir = Join-Path $script:Tmp 'six3\out'; $null = New-Item -ItemType Directory -Path $dir
            Add-UserFiles $dir
            $r = [MailboxTransfer.PstExporter]::Run($store, (New-ExportOptions (Join-Path $dir 'Archive.pst')), $null, [Threading.CancellationToken]::None)
            $r.Ok | Should -BeTrue
            (Get-Item -LiteralPath (Join-Path $dir 'Archive-2019.pst')).Length | Should -Be 3
        } finally { $store.Dispose() }
    }
}

Describe 'Review 7: the console starts its children without manual quoting' {
    BeforeAll {
        foreach ($f in 'Core', 'Run', 'Jobs') { . (Join-Path $script:Root "module\MailboxTransfer.Gui.$f.ps1") }
        $script:Echo = Join-Path $script:Tmp 'echo-args.ps1'
        Set-Content -LiteralPath $script:Echo -Value 'param([Parameter(ValueFromRemainingArguments)]$a); $a | ConvertTo-Json -AsArray | Set-Content -LiteralPath (Join-Path $PSScriptRoot "args.json") -Encoding utf8; "to stdout"; [Console]::Error.WriteLine("to stderr")'
    }
    It 'every argument reaches the child as typed (spaces, quotes, trailing backslash, injection attempt, empty, dash)' {
        $nasty = @('C:\Out Folder\', 'a" -ConfigPath evil b', 'C:\My Dir\sub\', '', 'He said "hi"', 'back\slash\"', 'end\', '-ConfigPath', 'x y z', "caf$([char]0xE9) $([char]0x4E2D)")
        $dir = Join-Path $script:Tmp 'job7'; $null = New-Item -ItemType Directory -Path $dir
        $job = New-MbtGuiJob -WorkFolder $dir -Task export -EntryScript $script:Echo
        $job.Queue.Enqueue([pscustomobject]@{ Title = 't'; Arguments = $nasty })
        $null = Start-MbtGuiProcess -Job $job
        $job.Process.WaitForExit(60000) | Should -BeTrue
        Close-MbtGuiProcessPipes -Job $job
        $got = @(Get-Content -LiteralPath (Join-Path $script:Tmp 'args.json') -Raw -Encoding utf8 | ConvertFrom-Json)
        # the console puts its own options first (-LogFolder <work>\logs, and -DevelopmentMode from a source checkout), then the arguments of the task
        $got[0] | Should -BeExactly '-LogFolder'
        $got = @($got | Select-Object -Skip ($got.Count - $nasty.Count))
        $got.Count | Should -Be $nasty.Count
        for ($i = 0; $i -lt $nasty.Count; $i++) { $got[$i] | Should -BeExactly $nasty[$i] }
    }
    It 'the output and the error of the child go to the files of the job' {
        $dir = Join-Path $script:Tmp 'job7b'; $null = New-Item -ItemType Directory -Path $dir
        $job = New-MbtGuiJob -WorkFolder $dir -Task export -EntryScript $script:Echo
        $job.Queue.Enqueue([pscustomobject]@{ Title = 't'; Arguments = @('x') })
        $null = Start-MbtGuiProcess -Job $job
        $job.Process.WaitForExit(60000) | Should -BeTrue
        Close-MbtGuiProcessPipes -Job $job
        (Get-Content -LiteralPath $job.StdOut -Raw) | Should -Match 'to stdout'
        (Get-Content -LiteralPath $job.StdErr -Raw) | Should -Match 'to stderr'
    }
    It 'refuses a folder name the engine could take for a parameter or that has a control character' -ForEach @(
        @{ Name = '-ConfigPath'; Bad = $true }, @{ Name = '  -x'; Bad = $true }, @{ Name = ''; Bad = $true }, @{ Name = "a`nb"; Bad = $true }, @{ Name = 'a//b'; Bad = $true },
        @{ Name = '../x'; Bad = $true }, @{ Name = ('a' * 300); Bad = $true },
        @{ Name = 'Imported 2019-11-09'; Bad = $false }, @{ Name = 'Imported/archive 2019'; Bad = $false }, @{ Name = 'a" b'; Bad = $false }, @{ Name = "R$([char]0xE9)sum$([char]0xE9)"; Bad = $false }
    ) {
        [bool](Test-MbtGuiFolderName $Name) | Should -Be $Bad
    }
    It 'the command of an import throws for a refused folder name' {
        $s = @{ Task = 'import'; Mailbox = 'a@contoso.com'; ImShared = ''; Dest = 'pst'; OutFolder = ''; OutName = ''; SplitGB = 20; CheckEach = $true; SepAux = $false; Conc = 4; Tries = 6; Reconcile = 5; Awake = $false; ImSrc = 'pst'; ImPath = 'D:\a.pst'; TgtLoc = 'Primary'; TgtMode = 'new'; TgtName = '-ConfigPath'; Mapping = ''; Changed = @{} }
        $j = @{ Id = 'j'; Folder = 'D:\J'; EventsPath = 'D:\J\e'; StopPath = 'D:\J\s'; ConfigPath = $null; StorePath = 'D:\S'; SelectionFiles = @(); JobName = 'import-1' }
        { Get-MbtGuiCommand -Task import -S $s -Job $j } | Should -Throw '*refused*'
        $s.TgtName = 'Imported "x" \'
        $a = (Get-MbtGuiCommand -Task import -S $s -Job $j)[0].Arguments
        $a[($a.IndexOf('-TargetRoot') + 1)] | Should -BeExactly 'Imported "x" \'
    }
}

Describe 'Review 17: report.csv does not carry formulas' {
    BeforeAll { foreach ($f in 'Core', 'Run', 'Jobs') { . (Join-Path $script:Root "module\MailboxTransfer.Gui.$f.ps1") } }
    It 'a message that starts with = + - @ gets an apostrophe in the CSV and stays as it is in the JSON' {
        $r = New-MbtGuiRun -Task import
        $r.Warnings.Add(@{ Step = 's'; Key = '=HYPERLINK("http://evil.example","x")'; Code = '+cmd'; Msg = '@SUM(A1)' })
        $r.Warnings.Add(@{ Step = 's'; Key = 'ok folder'; Code = 'W1'; Msg = '-2+3' })
        $folder = Join-Path $script:Tmp 'rep17'
        $files = Write-MbtGuiRunReport -Run $r -Folder $folder -Task import
        $csv = @(Import-Csv -LiteralPath (Join-Path $folder 'report.csv') | Where-Object { $_.Kind -eq 'warning' })
        $csv.Count | Should -Be 2
        $csv[0].Key | Should -BeExactly "'=HYPERLINK(""http://evil.example"",""x"")"
        $csv[0].State | Should -BeExactly "'+cmd"
        $csv[0].Message | Should -BeExactly "'@SUM(A1)"
        $csv[1].Message | Should -BeExactly "'-2+3"
        $csv[1].Key | Should -BeExactly 'ok folder'
        $json = Get-Content -LiteralPath (Join-Path $folder 'report.json') -Raw | ConvertFrom-Json
        @($json.rows | Where-Object { $_.Kind -eq 'warning' })[0].Message | Should -BeExactly '@SUM(A1)'
    }
}

Describe 'Review 9, 11, 15: the import journal' {
    BeforeAll {
        $script:Tenant = 'cccccccc-dddd-eeee-ffff-000000000000'
        $script:Dst = "MBX:99999999-8888-7777-6666-555555555555@$($script:Tenant)"
        function New-Msg([string]$Subject) {
            $P = [MailboxTransfer.FxProperty]
            $m = [MailboxTransfer.FxMessage]::new()
            $m.Properties.Add($P::FromString(0x001A, 'IPM.Note')); $m.Properties.Add($P::FromString(0x0037, $Subject)); $m.Properties.Add($P::FromString(0x1000, 'body'))
            $r = [MailboxTransfer.FxRecipient]::new()
            $r.Properties.Add($P::FromInt32(0x0C15, 1)); $r.Properties.Add($P::FromString(0x3001, 'Bob')); $m.Recipients.Add($r)
            return $m
        }
        function New-SmallPst([string]$Path, [string[]]$Subjects) {
            $null = New-Item -ItemType Directory -Force -Path (Split-Path $Path)
            $w = [MailboxTransfer.PstWriter]::Create($Path, 'Old')
            $inbox = $w.CreateFolder($null, 'Inbox', 'IPF.Note')
            foreach ($s in $Subjects) { [void]$w.AddMessage($inbox, (New-Msg $s)) }
            $w.Close()
        }
        function Invoke-SmallJob([string]$Pst, [string]$Journal, [string]$Work, [string]$Name = 'job1') {
            $svc = [MailboxTransfer.Tests.FakeMailboxService]::new(); $null = $svc.AddMailbox($script:Dst)
            $f = $svc.AddFolder($script:Dst, 'F-INBOX', 'Inbox'); $f.WellKnown = 'inbox'
            $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
            $client = [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $svc)
            $po = [MailboxTransfer.PstImportOptions]::new(); $po.PstPaths.Add($Pst); $po.TargetMailboxId = $script:Dst; $po.TargetFolder = 'Imported/old'; $po.JobName = $Name
            $po.Concurrency = 2; $po.ReconcileDelaySeconds = 0; $po.WorkPath = $Work
            $j = [MailboxTransfer.PstImportJournal]::new($Journal)
            try {
                $job = [MailboxTransfer.PstImportJob]::Start([MailboxTransfer.MailboxApi]::new($client), $j, $po, $Work, $null)
                $messages = [System.Collections.Generic.List[string]]::new()
                $deadline = [DateTime]::UtcNow.AddSeconds(90)
                while (-not $job.Wait(100)) { foreach ($m in $job.TakeMessages()) { $messages.Add($m) }; if ([DateTime]::UtcNow -gt $deadline) { $job.Cancel(); throw 'timeout' } }
                foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
                return [pscustomobject]@{ Messages = $messages; Result = $job.Result; Error = $job.Error; Service = $svc }
            } finally { $j.Dispose() }
        }
    }

    It '15: a second open of a journal that is open is refused, and works once it is closed' {
        $path = Join-Path $script:Tmp 'lock\j.db'
        $j1 = [MailboxTransfer.PstImportJournal]::new($path)
        try {
            { [MailboxTransfer.PstImportJournal]::new($path) } | Should -Throw '*already open by another process*'
        } finally { $j1.Dispose() }
        $j2 = [MailboxTransfer.PstImportJournal]::new($path); $j2.Dispose()
        Test-Path -LiteralPath "$path.lock" | Should -BeFalse
    }
    It '15: the lock file of an open journal is refused to another pwsh process too' {
        $path = Join-Path $script:Tmp 'lock2\j.db'
        $j1 = [MailboxTransfer.PstImportJournal]::new($path)
        try {
            $probe = 'param($p); try { $fs = [IO.File]::Open("$p.lock", "Open", "ReadWrite", "None"); $fs.Dispose(); "OPENED" } catch [IO.IOException] { "REFUSED" }'
            $f = Join-Path $script:Tmp 'probe15.ps1'; Set-Content -LiteralPath $f -Value $probe
            (& (Get-Process -Id $PID).Path -NoProfile -File $f $path) | Should -Be 'REFUSED'
        } finally { $j1.Dispose() }
    }
    It '9: the journal writes with synchronous=FULL' {
        $path = Join-Path $script:Tmp 'sync\j.db'
        $j = [MailboxTransfer.PstImportJournal]::new($path)
        try {
            $c = [MailboxTransfer.PstImportJournal].GetField('_c', [Reflection.BindingFlags]'NonPublic,Instance').GetValue($j)
            $cmd = $c.CreateCommand(); $cmd.CommandText = 'PRAGMA synchronous;'
            [int]$cmd.ExecuteScalar() | Should -Be 2
        } finally { $j.Dispose() }
    }
    It '9: the import rows of the store go through FULL and come back to NORMAL (ImportSyncFull = <Full>)' -ForEach @(@{ Full = $true }, @{ Full = $false }) {
        $dir = Join-Path $script:Tmp ("st9-$Full")
        $opt = [MailboxTransfer.StoreOptions]::new(); $opt.ImportSyncFull = $Full
        $store = [MailboxTransfer.ExportStore]::Create($dir, 'st9', 'tests', $opt)
        try {
            $mb = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $mb; Kind = 'Primary'; Upn = 'u@x.test'; DisplayName = 'Primary' })
            $fid = [MailboxTransfer.Tests.FakeIds]::Folder('11111111-2222-3333-4444-555555555555', 1)
            [void]$store.UpsertFolders([MailboxTransfer.StoreFolder[]]@([MailboxTransfer.StoreFolder]@{ MailboxId = $mb; FolderId = $fid; DisplayName = 'Inbox'; FolderType = 'IPF.Note'; SourceItemCount = 0 }))
            $bytes = [MailboxTransfer.FxCodec]::EncodePayload((New-Msg 'x'))
            [void]$store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@([MailboxTransfer.StoreItemEntry]@{ MailboxId = $mb; ItemId = 'I1'; FolderId = $fid; Size = $bytes.Length; ItemType = 'IPM.Note' }))
            $key = $store.FindItemKey($mb, 'I1')
            [void]$store.WriteItem($key, [IO.MemoryStream]::new($bytes), [Threading.CancellationToken]::None)
            $job = $store.CreateImportJob('j', $mb)
            $store.AddImportItems($job, [MailboxTransfer.StoreImportPlan[]]@([MailboxTransfer.StoreImportPlan]@{ ItemKey = $key; TargetMailbox = $mb; TargetFolder = 'F' })) | Should -Be 1
            $next = $store.NextImportItems($job, 5)
            $next.Count | Should -Be 1
            $store.MarkImportDone($next[0].ImportKey, 'NEW')
            @($store.GetImportItems($job, [MailboxTransfer.StoreImportState]::Done, 10)).Count | Should -Be 1
            $w = [MailboxTransfer.ExportStore].GetField('_w', [Reflection.BindingFlags]'NonPublic,Instance').GetValue($store)
            $cmd = $w.CreateCommand(); $cmd.CommandText = 'PRAGMA synchronous;'
            [int]$cmd.ExecuteScalar() | Should -Be 1 -Because 'the connection is back on NORMAL after an import commit'
        } finally { $store.Dispose() }
    }
    It '11: two different PST files with the same name and the same size do not get the same identity' {
        $a = Join-Path $script:Tmp 'id11\a\Archive.pst'; $b = Join-Path $script:Tmp 'id11\b\Archive.pst'
        New-SmallPst $a @('aaaa 1', 'aaaa 2'); New-SmallPst $b @('bbbb 1', 'bbbb 2')
        (Get-Item $a).Length | Should -Be (Get-Item $b).Length
        [MailboxTransfer.PstImportFiles]::LegacyIdOf($a) | Should -Be ([MailboxTransfer.PstImportFiles]::LegacyIdOf($b))
        [MailboxTransfer.PstImportFiles]::IdOf($a) | Should -Not -Be ([MailboxTransfer.PstImportFiles]::IdOf($b))
        [MailboxTransfer.PstImportFiles]::IdOf($a) | Should -Be ([MailboxTransfer.PstImportFiles]::IdOf($a))
    }
    It '11: a journal that knows a file by name and length keeps that key' {
        $a = Join-Path $script:Tmp 'id11b\Old.pst'; New-SmallPst $a @('x')
        $legacy = [MailboxTransfer.PstImportFiles]::LegacyIdOf($a)
        $known = [Func[string, MailboxTransfer.PstImportSource]] { param($id) if ($id -eq $legacy) { [MailboxTransfer.PstImportSource]@{ PstId = $legacy; SourcePath = $a } } else { $null } }
        $o = [MailboxTransfer.PstImportOptions]::new(); $o.PstPaths.Add($a); $o.Check = $false
        $s = [MailboxTransfer.PstImportPrepare]::Run($o, $known, (Join-Path $script:Tmp 'id11b\work'), $false, $null, [Threading.CancellationToken]::None)
        $s[0].PstId | Should -Be $legacy
        $unknown = [Func[string, MailboxTransfer.PstImportSource]] { param($id) $null }
        ([MailboxTransfer.PstImportPrepare]::Run($o, $unknown, (Join-Path $script:Tmp 'id11b\work'), $false, $null, [Threading.CancellationToken]::None))[0].PstId | Should -Be ([MailboxTransfer.PstImportFiles]::IdOf($a))
    }
    It '11: a PST that changed between two runs of the same job is reported' {
        $pst = Join-Path $script:Tmp 'id11c\Old.pst'; $journal = Join-Path $script:Tmp 'id11c\j.db'; $work = Join-Path $script:Tmp 'id11c\work'
        New-SmallPst $pst @('one', 'two')
        $first = Invoke-SmallJob $pst $journal $work
        $first.Result.Completed | Should -BeTrue
        ($first.Messages -join "`n") | Should -Not -Match 'not the file this job imported before'
        $same = Invoke-SmallJob $pst $journal $work
        ($same.Messages -join "`n") | Should -Not -Match 'not the file this job imported before'
        Remove-Item -LiteralPath $pst -Force; New-SmallPst $pst @('one', 'two', 'three')
        $changed = Invoke-SmallJob $pst $journal $work
        ($changed.Messages -join "`n") | Should -Match 'WARNING: .*Old\.pst is not the file this job imported before'
    }
}

Describe 'Review 13: the token cache file is safe across processes' -Skip:(-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
    BeforeAll {
        Initialize-MbtAuthEngine
        if (-not ('ReviewFileHolder' -as [type])) {
            Add-Type -TypeDefinition 'using System; using System.IO; using System.Threading; public static class ReviewFileHolder { public static void Hold(string path, int ms) { var fs = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None); var t = new Thread(() => { Thread.Sleep(ms); fs.Dispose(); }); t.IsBackground = true; t.Start(); } }'
        }
        $script:Hashes = foreach ($v in 0, 65, 66) { $b = [byte[]]::new(300000); [Array]::Fill($b, [byte]$v); [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($b)) }
        $script:ChildScript = Join-Path $script:Tmp 'cache-writer.ps1'
        Set-Content -LiteralPath $script:ChildScript -Value @'
param($Root, $Cache, $Count)
Import-Module (Join-Path $Root 'MailboxTransfer.psd1')
Initialize-MbtAuthEngine -Root $Root
$f = [MailboxTransfer.TokenCacheFile]::new($Cache)
$bytes = [byte[]]::new(300000); [Array]::Fill($bytes, [byte]66)
for ($i = 0; $i -lt $Count; $i++) { $f.Enter(); try { $null = $f.Read(); $f.Write($bytes) } finally { $f.Leave() } }
'done'
'@
    }
    It 'writes through a temporary file, reads it back, and leaves no temporary file' {
        $p = Join-Path $script:Tmp 'tc1\token.cache'
        $f = [MailboxTransfer.TokenCacheFile]::new($p)
        $f.Read() | Should -BeNullOrEmpty
        $f.Write([byte[]](1, 2, 3, 4))
        [IO.File]::ReadAllBytes($p) | Should -Not -Be ([byte[]](1, 2, 3, 4)) -Because 'the file is protected'
        $f.Read() | Should -Be ([byte[]](1, 2, 3, 4))
        $f.Write([byte[]](9, 9))
        $f.Read() | Should -Be ([byte[]](9, 9))
        @(Get-ChildItem -LiteralPath (Split-Path $p) -Filter '*.tmp').Count | Should -Be 0
    }
    It 'a file that stays unreadable is tried again, then put aside and reported (never silently dropped)' {
        $p = Join-Path $script:Tmp 'tc2\token.cache'
        $null = New-Item -ItemType Directory -Path (Split-Path $p)
        [IO.File]::WriteAllBytes($p, [byte[]](1..200))
        $f = [MailboxTransfer.TokenCacheFile]::new($p); $f.ReadAttempts = 2
        $f.Read() | Should -BeNullOrEmpty
        $f.Warning | Should -Match 'could not be read.*kept as .*unreadable'
        Test-Path -LiteralPath "$p.unreadable" | Should -BeTrue
        [IO.File]::ReadAllBytes("$p.unreadable").Length | Should -Be 200
        Test-Path -LiteralPath $p | Should -BeFalse
    }
    It 'a read that fails once (file held by another writer) succeeds on the retry' {
        $p = Join-Path $script:Tmp 'tc3\token.cache'
        $f = [MailboxTransfer.TokenCacheFile]::new($p)
        $f.Write([byte[]](5, 6, 7))
        [ReviewFileHolder]::Hold($p, 250)
        $f.Read() | Should -Be ([byte[]](5, 6, 7))
        $f.Warning | Should -BeNullOrEmpty
    }
    It 'two processes writing the same cache never leave a torn file for the reader' {
        $p = Join-Path $script:Tmp 'tc4\token.cache'
        $f = [MailboxTransfer.TokenCacheFile]::new($p)
        $f.Write([byte[]]::new(300000))
        $psi = [Diagnostics.ProcessStartInfo]::new((Get-Process -Id $PID).Path)
        foreach ($a in '-NoProfile', '-File', $script:ChildScript, $script:Root, $p, '150') { $psi.ArgumentList.Add($a) }
        $psi.UseShellExecute = $false; $psi.CreateNoWindow = $true; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
        $child = [Diagnostics.Process]::Start($psi)
        $bad = 0; $reads = 0; $mine = [byte[]]::new(300000); [Array]::Fill($mine, [byte]65)
        try {
            $deadline = [DateTime]::UtcNow.AddSeconds(90)
            while (-not $child.HasExited -and [DateTime]::UtcNow -lt $deadline) {
                $f.Enter()
                try {
                    $d = $f.Read(); $reads++
                    if ($null -eq $d -or $d.Length -ne 300000 -or [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($d)) -notin $script:Hashes) { $bad++ }
                    $f.Write($mine)
                } finally { $f.Leave() }
            }
            $child.WaitForExit(30000) | Should -BeTrue
        } finally { if (-not $child.HasExited) { $child.Kill() } }
        $reads | Should -BeGreaterThan 5
        $bad | Should -Be 0
        $f.Warning | Should -BeNullOrEmpty
        @(Get-ChildItem -LiteralPath (Split-Path $p) -Filter '*.tmp').Count | Should -Be 0
    }
}

Describe 'Review 14: a message moved during the export is recorded once on the primary mailbox' {
    BeforeAll {
        $script:PG = '11111111-2222-3333-4444-555555555555'
        $script:Prim = "MBX:$($script:PG)@cccccccc-dddd-eeee-ffff-000000000000"
        function Fid14([long]$N) { [MailboxTransfer.Tests.FakeIds]::Folder($script:PG, $N) }
        function Mid14([long]$Folder, [long]$N) { [MailboxTransfer.Tests.FakeIds]::Message($script:PG, $Folder, $N) }
        function New-Msg14([string]$Subject) {
            $b = [MailboxTransfer.Tests.FxBuilder]::new()
            $null = $b.Str(0x001A, 'IPM.Note').Str(0x0037, $Subject)
            return $b.ToArray()
        }
        # message 7 and 8; -Moved: message 7 is now in the second folder (new item id, same message counter)
        function New-Service14([switch]$Moved) {
            $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
            $null = $svc.AddMailbox($script:Prim)
            $null = $svc.AddFolder($script:Prim, (Fid14 1), 'Inbox')
            $null = $svc.AddFolder($script:Prim, (Fid14 2), 'Projects')
            $null = $svc.AddItem($script:Prim, $(if ($Moved) { Fid14 2 } else { Fid14 1 }), $(if ($Moved) { Mid14 2 7 } else { Mid14 1 7 }), (New-Msg14 'Moved one'), 'ck-7')
            $null = $svc.AddItem($script:Prim, (Fid14 1), (Mid14 1 8), (New-Msg14 'Stays'), 'ck-8')
            return $svc
        }
        function Invoke-Export14($Service, $Store, [scriptblock]$Options) {
            $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
            $client = [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
            $eo = [MailboxTransfer.ExportOptions]::new(); $eo.MailboxId = $script:Prim; $eo.Kind = 'Primary'; $eo.Upn = 'adil@contoso.test'; $eo.Concurrency = 2
            $eo.TempPath = Join-Path $script:Tmp ('tmp14-' + [guid]::NewGuid().ToString('N'))
            if ($Options) { & $Options $eo }
            $job = [MailboxTransfer.ExportJob]::Start([MailboxTransfer.MailboxApi]::new($client), $Store, $eo)
            $deadline = [DateTime]::UtcNow.AddSeconds(60)
            while (-not $job.Wait(100)) { if ([DateTime]::UtcNow -gt $deadline) { $job.Cancel(); throw 'timeout' } }
            return $job.Result
        }
    }
    It 'a message exported from its first folder and moved before the next run is not exported again' {
        $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp 'mv1'), 'mv1', 'tests')
        try {
            $a = Invoke-Export14 (New-Service14) $store
            $a.Exported | Should -Be 2
            $svc2 = New-Service14 -Moved
            $b = Invoke-Export14 $svc2 $store
            $b.Exported | Should -Be 0
            $b.ItemsDeduplicated | Should -Be 1
            $store.GetStats().Items | Should -Be 2
            @($svc2.ExportCalls).Count | Should -Be 0 -Because 'nothing is asked of the service again'
        } finally { $store.Dispose() }
    }
    It 'a message recorded but not exported yet follows its new id, so that the export asks for the id that exists now' {
        $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp 'mv2'), 'mv2', 'tests')
        try {
            # an earlier inventory recorded message 7 under its old id and was stopped before the export
            [void]$store.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = $script:Prim; Kind = 'Primary'; Upn = 'adil@contoso.test'; DisplayName = 'Primary' })
            [void]$store.UpsertFolders([MailboxTransfer.StoreFolder[]]@([MailboxTransfer.StoreFolder]@{ MailboxId = $script:Prim; FolderId = (Fid14 1); DisplayName = 'Inbox'; FolderType = 'IPF.Note'; SourceItemCount = 0 }))
            [void]$store.UpsertItems([MailboxTransfer.StoreItemEntry[]]@([MailboxTransfer.StoreItemEntry]@{ MailboxId = $script:Prim; ItemId = (Mid14 1 7); FolderId = (Fid14 1); Size = 10; ItemType = 'IPM.Note' }))
            $svc2 = New-Service14 -Moved
            $b = Invoke-Export14 $svc2 $store
            $b.ItemsRehomed | Should -Be 1
            $b.ItemsDeduplicated | Should -Be 0
            $store.GetStats().Items | Should -Be 2
            $store.GetItem($script:Prim, (Mid14 2 7)).State | Should -Be 'Exported'
        } finally { $store.Dispose() }
    }
    It 'items that only look alike are kept: two messages of different counters are both recorded' {
        $store = [MailboxTransfer.ExportStore]::Create((Join-Path $script:Tmp 'mv3'), 'mv3', 'tests')
        try {
            $r = Invoke-Export14 (New-Service14) $store
            $r.Exported | Should -Be 2
            $r.ItemsDeduplicated | Should -Be 0
        } finally { $store.Dispose() }
    }
}

Describe 'Review 19: the Exchange conversion cleans up and streams the base64' {
    BeforeAll {
        $script:Src19 = "MBX:11111111-2222-3333-4444-555555555555@cccccccc-dddd-eeee-ffff-000000000000"
        function New-World19 {
            $ie = [MailboxTransfer.Tests.FakeMailboxService]::new()
            $ie.RootOnlyFolders = $true
            $mail = [MailboxTransfer.Tests.FakeMailApi]::new($ie)
            $null = $ie.AddMailbox($script:Src19); $mail.AddMailbox($script:Src19)
            $fake = [MailboxTransfer.Tests.EmlFakeMail]::new($mail, $ie)
            $null = $ie.AddFolder($script:Src19, 'ROOT', 'Dummy')
            $fake.DraftFolder = 'ROOT'
            $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
            $client = [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $fake)
            return [pscustomobject]@{ Ie = $ie; Fake = $fake; Api = [MailboxTransfer.MailboxApi]::new($client) }
        }
        function Get-DraftData([byte[]]$Eml) {
            $m = [MailboxTransfer.MimeToMapi]::Build([MailboxTransfer.MimeParser]::Parse($Eml))
            $ms = [IO.MemoryStream]::new(); $s = [MailboxTransfer.FxMessageSource]::new($m, [MailboxTransfer.FxImportStream]::NewEnvelope()).OpenRead(); $s.CopyTo($ms); $s.Dispose()
            return $ms.ToArray()
        }
    }
    It 'sends the MIME as base64 text of the exact size, whatever its length (<Size> bytes), and leaves no temporary folder after Dispose' -ForEach @(
        @{ Size = 200 }, @{ Size = 201 }, @{ Size = 202 }, @{ Size = 49152 }, @{ Size = 49153 }, @{ Size = 150001 }
    ) {
        $w = New-World19
        $head = [Text.Encoding]::ASCII.GetBytes("From: a@x.fr`r`nTo: b@x.fr`r`nSubject: S`r`nDate: Mon, 18 Mar 2024 10:00:00 +0000`r`nContent-Type: text/plain`r`n`r`n")
        $eml = [byte[]]::new($Size); [Array]::Copy($head, $eml, $head.Length); for ($i = $head.Length; $i -lt $Size; $i++) { $eml[$i] = [byte](65 + ($i % 26)) }
        $w.Fake.NextDraftData = Get-DraftData $eml
        $temp = Join-Path $script:Tmp ("conv19-$Size"); $null = New-Item -ItemType Directory -Path $temp
        $conv = [MailboxTransfer.EmlExchangeConverter]::new($w.Api, $script:Src19, $script:Src19); $conv.TempDirectory = $temp
        $notes = [Collections.Generic.List[string]]::new()
        $c = $conv.Convert($eml, [DateTimeOffset]::UtcNow, $null, $false, [Threading.CancellationToken]::None, $notes)
        try {
            Test-Path -LiteralPath $c.TempFile | Should -BeTrue
            (Get-FileHash -InputStream ([IO.MemoryStream]::new($w.Fake.LastDraftBody)) -Algorithm SHA256).Hash | Should -Be (Get-FileHash -InputStream ([IO.MemoryStream]::new($eml)) -Algorithm SHA256).Hash
            $w.Fake.LastDraftBody.Length | Should -Be $Size
        } finally { $c.Dispose() }
        @(Get-ChildItem -LiteralPath $temp).Count | Should -Be 0 -Because 'neither the file nor the folder of the conversion is left'
    }
    It 'a conversion that fails leaves no temporary folder' {
        $w = New-World19
        $w.Fake.DenyDrafts = $true
        $temp = Join-Path $script:Tmp 'conv19-fail'; $null = New-Item -ItemType Directory -Path $temp
        $conv = [MailboxTransfer.EmlExchangeConverter]::new($w.Api, $script:Src19, $script:Src19); $conv.TempDirectory = $temp
        { $conv.Convert([Text.Encoding]::ASCII.GetBytes("Subject: x`r`n`r`nhi"), [DateTimeOffset]::UtcNow, $null, $false, [Threading.CancellationToken]::None, [Collections.Generic.List[string]]::new()) } | Should -Throw '*Mail.ReadWrite*'
        @(Get-ChildItem -LiteralPath $temp).Count | Should -Be 0
    }
    It 'a draft that cannot be exported leaves no temporary folder' {
        $w = New-World19      # no NextDraftData: the fake answers an error for the draft
        $temp = Join-Path $script:Tmp 'conv19-fail2'; $null = New-Item -ItemType Directory -Path $temp
        $conv = [MailboxTransfer.EmlExchangeConverter]::new($w.Api, $script:Src19, $script:Src19); $conv.TempDirectory = $temp
        { $conv.Convert([Text.Encoding]::ASCII.GetBytes("Subject: x`r`n`r`nhi"), [DateTimeOffset]::UtcNow, $null, $false, [Threading.CancellationToken]::None, [Collections.Generic.List[string]]::new()) } | Should -Throw
        @(Get-ChildItem -LiteralPath $temp).Count | Should -Be 0
    }
}

Describe 'Review 3: no silent default to the lab identity' {
    BeforeAll {
        Import-Module (Join-Path $script:Root 'MailboxTransfer.psd1') -Force
        $script:Cfg = [pscustomobject]@{ Lab = @{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '00000000-1111-2222-3333-444444444444'; UserPrincipalName = 'Adele@contoso.com'
                PrimaryMailboxId = 'MBX:11111111-2222-3333-4444-555555555555@t'; ArchiveMailboxId = 'MBX:66666666-7777-8888-9999-000000000000@t' } }
        . (Join-Path $script:Root 'module\MailboxTransfer.Gui.Core.ps1')
        . (Join-Path $script:Root 'module\MailboxTransfer.Gui.Pages.Export.ps1')
    }
    It 'the configured mailbox ids are given for the user of the configuration only (any case)' {
        $mine = Get-MbtConfiguredMailboxId -Configuration $script:Cfg -UserPrincipalName 'ADELE@contoso.com'
        $mine.Primary | Should -Be 'MBX:11111111-2222-3333-4444-555555555555@t'
        $mine.Archive | Should -Be 'MBX:66666666-7777-8888-9999-000000000000@t'
    }
    It 'another user, no user, or a configuration without ids gets none: the ids are discovered' {
        foreach ($upn in 'bob@contoso.com', '', $null) {
            $r = Get-MbtConfiguredMailboxId -Configuration $script:Cfg -UserPrincipalName $upn
            $r.Primary | Should -BeNullOrEmpty; $r.Archive | Should -BeNullOrEmpty
        }
        $empty = [pscustomobject]@{ Lab = @{ UserPrincipalName = 'adele@contoso.com'; PrimaryMailboxId = ''; ArchiveMailboxId = '' } }
        (Get-MbtConfiguredMailboxId -Configuration $empty -UserPrincipalName 'adele@contoso.com').Primary | Should -BeNullOrEmpty
        (Get-MbtConfiguredMailboxId -Configuration ([pscustomobject]@{}) -UserPrincipalName 'adele@contoso.com').Primary | Should -BeNullOrEmpty
    }
    It 'the sign-in refuses a missing or malformed tenant or application with a message that says what to give' {
        { Assert-MbtSignInSettings -TenantId '' -ClientId '' } | Should -Throw '*tenant*application*-TenantId and -ClientId*'
        { Assert-MbtSignInSettings -TenantId 'contoso.onmicrosoft.com' -ClientId '' } | Should -Throw '*application (client) id*'
        { Assert-MbtSignInSettings -TenantId '' -ClientId '00000000-1111-2222-3333-444444444444' } | Should -Throw '*tenant*'
        { Assert-MbtSignInSettings -TenantId 'not a tenant' -ClientId '00000000-1111-2222-3333-444444444444' } | Should -Throw '*valid tenant*'
        { Assert-MbtSignInSettings -TenantId 'contoso.onmicrosoft.com' -ClientId 'abc' } | Should -Throw '*valid application*'
        { Assert-MbtSignInSettings -TenantId 'cccccccc-dddd-eeee-ffff-000000000000' -ClientId '00000000-1111-2222-3333-444444444444' } | Should -Not -Throw
        { Assert-MbtSignInSettings -TenantId 'contoso.onmicrosoft.com' -ClientId '00000000-1111-2222-3333-444444444444' } | Should -Not -Throw
    }
    It 'the entry script stops before any network call when it has no tenant or application' {
        $cfgFile = Join-Path $script:Tmp 'empty-lab.psd1'
        Set-Content -LiteralPath $cfgFile -Value "@{ Lab = @{ TenantId = ''; ClientId = ''; UserPrincipalName = ''; UserObjectId = ''; PrimaryMailboxId = ''; ArchiveMailboxId = ''; SourceFolder = ''; TargetFolder = '' } }"
        # the git-ignored local configuration of the maintainer must not give the child a tenant: it reads a missing file instead
        $savedLocal = $env:MAILBOXTRANSFER_LOCAL_CONFIG
        try {
            $env:MAILBOXTRANSFER_LOCAL_CONFIG = Join-Path $script:Tmp 'no-such-local-file.psd1'
            $out = & (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $script:Root 'Invoke-MailboxTransfer.ps1') -DevelopmentMode -Mode Browse -ConfigPath $cfgFile -UserPrincipalName 'bob@contoso.com' -StorePath (Join-Path $script:Tmp 'lab-store') 2>&1 | Out-String
        } finally { $env:MAILBOXTRANSFER_LOCAL_CONFIG = $savedLocal }
        $out | Should -Match 'Sign-in needs the tenant .* and the application'
        $out | Should -Not -Match 'signing in as'
    }
    It 'no script of the package takes a mailbox id from the Lab section without the user check' {
        $hits = Get-ChildItem -LiteralPath $script:Root -Include *.ps1, *.psm1 -Recurse | Select-String -Pattern 'Lab\.(Primary|Archive)MailboxId' |
            Where-Object { $_.Path -notmatch 'MailboxTransfer\.psm1$' -or $_.Line -notmatch 'lab\.PrimaryMailboxId|lab\.ArchiveMailboxId' }
        @($hits | Where-Object { $_.Line -notmatch '^\s*(#|<#)' -and $_.Line -notmatch '\.PARAMETER|configuration \(' }).Count | Should -Be 0 -Because (($hits | ForEach-Object { "$($_.Path):$($_.LineNumber)" }) -join ', ')
    }
    It 'the Sign-in card asks for the tenant and the application only when the configuration has none' {
        $script:Gui = @{ S = @{ TenantId = ''; ClientId = '' }; Engine = [pscustomobject]@{ SignInSettings = { [pscustomobject]@{ TenantId = ''; ClientId = '' } } } }
        $s = Get-MbtGuiSignInSettings
        $s.Asks | Should -BeTrue
        $script:Gui.S.TenantId = 'contoso.onmicrosoft.com'; $script:Gui.S.ClientId = '00000000-1111-2222-3333-444444444444'
        $s = Get-MbtGuiSignInSettings
        $s.TenantId | Should -Be 'contoso.onmicrosoft.com'; $s.ClientId | Should -Be '00000000-1111-2222-3333-444444444444'
        $script:Gui = @{ S = @{ TenantId = ''; ClientId = '' }; Engine = [ordered]@{ SignInSettings = { [pscustomobject]@{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '00000000-1111-2222-3333-444444444444' } } } }
        $s = Get-MbtGuiSignInSettings
        $s.Asks | Should -BeFalse
        $s.TenantId | Should -Be 'contoso.onmicrosoft.com'
        $script:Gui = @{ S = @{ TenantId = ''; ClientId = '' }; Engine = [pscustomobject]@{} }
        (Get-MbtGuiSignInSettings).Asks | Should -BeFalse -Because 'an engine that does not say what it has is not asked questions'
    }
}
