#Requires -Version 7.4
<#
.SYNOPSIS
    Measures the memory used to move large items between FastTransfer streams and PST files with the streaming code of the engine:
    the working set and the private bytes of this process, sampled every 10 ms, for each phase.

.DESCRIPTION
    The memory used to move an item must not depend on its size. Each phase starts after a full garbage collection and reports the
    growth of the working set and of the private bytes over that baseline (the peak minus the value at the start).

    Scenario Corpus (needs the git-ignored corpus, default out\fxics-corpus: 1.3 GB, 100 MiB attachments):
      corpus-fx-to-pst     every .fxs is decoded lazily and written into one PST
      corpus-pst-read      every message is read back and compared with its original (values compared in pieces)
      corpus-pst-to-fx     every message of the PST is turned into a FastTransfer stream written to a sink
    Scenario Synthetic (one item built on the fly: message, 8 MiB body, an attachment of -SizeMiB, an embedded message with a
    16 MiB attachment; nothing of it is ever in memory as a whole):
      synthetic-create     the item is written as a Graph-style .fxs (envelope, records of 31,680 bytes) from FxMessageSource
      synthetic-fx-to-pst  .fxs -> PST (lazy decode, data trees written block by block)
      synthetic-pst-read   PST -> FxMessage (lazy), compared with the original and the attachments with the generator
      synthetic-pst-to-fx  PST message -> FastTransfer stream (FxMessageSource) -> file, then checked (round trip, attachment bytes)
      synthetic-import     the same stream through ImportContent (the body of importItem: base64 in a JSON envelope) to a sink

    Exit code: 0 when every check passed and every growth is under -LimitMiB, 1 otherwise.

.PARAMETER Scenario
    Corpus, Synthetic or All (default).

.PARAMETER CorpusPath
    Folder of .fxs files for the Corpus scenario (default <repository>\out\fxics-corpus).

.PARAMETER SizeMiB
    Size of the big attachment of the Synthetic scenario (default 1024 = 1 GiB; the PST and the .fxs files take that much disk each).

.PARAMETER WorkDir
    Where the files are written (default a new folder in the temporary directory, removed at the end).

.PARAMETER LimitMiB
    Growth of the working set above which a phase is reported as failed (default 150).

.PARAMETER Json
    Writes the phases as JSON on the output (for the tests) instead of the table.

.EXAMPLE
    .\tools\Measure-PstMemory.ps1
    .\tools\Measure-PstMemory.ps1 -Scenario Synthetic -SizeMiB 4096

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.
#>
[CmdletBinding()]
param(
    [ValidateSet('Corpus', 'Synthetic', 'All')][string]$Scenario = 'All',
    [string]$CorpusPath = (Join-Path $PSScriptRoot '..\out\fxics-corpus'),
    [ValidateRange(1, 3800)][int]$SizeMiB = 1024,
    [string]$WorkDir,
    [int]$LimitMiB = 150,
    [switch]$Json
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $repo 'tests\support\Import-StoreEngine.ps1')
Import-StoreEngine -PackageRoot (Join-Path $repo 'package') -WithTestSupport

if (-not $WorkDir) { $WorkDir = Join-Path ([IO.Path]::GetTempPath()) ('mbt-pst-memory-' + [guid]::NewGuid().ToString('N').Substring(0, 8)) }
$null = New-Item -ItemType Directory -Force -Path $WorkDir
$WorkDir = (Resolve-Path $WorkDir).Path

$phases = [Collections.Generic.List[object]]::new()
$failed = $false

function Measure-Phase([string]$Name, [scriptblock]$Body) {
    $probe = [MailboxTransfer.TestSupport.MemoryProbe]::new()
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $detail = $null
    try { $detail = & $Body } finally { $sw.Stop(); $probe.Stop() }
    $row = [ordered]@{
        Phase = $Name
        Seconds = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        BaseWorkingSetMiB = [int]($probe.BaseWorkingSet / 1MB)
        PeakWorkingSetMiB = [int]($probe.PeakWorkingSet / 1MB)
        WorkingSetGrowthMiB = [int]($probe.WorkingSetGrowth / 1MB)
        PrivateGrowthMiB = [int]($probe.PrivateGrowth / 1MB)
        Bytes = if ($detail -and $detail.ContainsKey('Bytes')) { [long]$detail.Bytes } else { 0 }
        Ok = $true
        Note = ''
    }
    if ($detail -and $detail.ContainsKey('Problems') -and @($detail.Problems).Count) { $row.Ok = $false; $row.Note = (@($detail.Problems) -join '; ') }
    if ($row.WorkingSetGrowthMiB -gt $LimitMiB) { $row.Ok = $false; $row.Note = ($row.Note + " working set grew by $($row.WorkingSetGrowthMiB) MiB (limit $LimitMiB)").Trim() }
    $script:phases.Add([pscustomobject]$row)
    if (-not $row.Ok) { $script:failed = $true }
    if (-not $Json) {
        $color = if ($row.Ok) { 'Green' } else { 'Red' }
        Write-Host ('{0,-22} {1,7:N1} s  working set +{2,4} MiB (peak {3,5}), private +{4,4} MiB  {5:N0} bytes  {6}' -f $row.Phase, $row.Seconds, $row.WorkingSetGrowthMiB, $row.PeakWorkingSetMiB, $row.PrivateGrowthMiB, $row.Bytes, $row.Note) -ForegroundColor $color
    }
}

function Get-ItemFolder([MailboxTransfer.FxMessage]$m) {
    $cls = $m.GetString(0x001A)
    if (-not $cls) { $cls = 'IPM.Note' }
    if ($cls -like 'IPM.Appointment*' -or $cls -like 'IPM.Schedule*') { return 'Calendar', 'IPF.Appointment' }
    if ($cls -like 'IPM.Contact*') { return 'Contacts', 'IPF.Contact' }
    if ($cls -like 'IPM.Task*') { return 'Tasks', 'IPF.Task' }
    if ($cls -like 'IPM.StickyNote*') { return 'Notes', 'IPF.StickyNote' }
    return 'Mail', 'IPF.Note'
}

$sink = [IO.Stream]::Null
$P = [MailboxTransfer.FxProperty]

try {
    # ------------------------------------------------------------------------------------------------ corpus
    if ($Scenario -in 'Corpus', 'All') {
        $corpus = if (Test-Path -LiteralPath $CorpusPath) { (Resolve-Path $CorpusPath).Path } else { $null }
        $files = if ($corpus) { @(Get-ChildItem -LiteralPath $corpus -Filter '*.fxs' -File | Sort-Object Name) } else { @() }
        if ($files.Count -eq 0) { if (-not $Json) { Write-Host "No corpus in ${CorpusPath}: scenario Corpus skipped." -ForegroundColor Yellow } }
        else {
            $pst = Join-Path $WorkDir 'corpus.pst'
            $where = @{}
            Measure-Phase 'corpus-fx-to-pst' {
                $split = [MailboxTransfer.PstSplitWriter]::new($pst, 'Corpus', 51200MB)
                $bytes = 0L; $problems = @()
                try {
                    foreach ($f in $files) {
                        $doc = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromFile($f.FullName))
                        if ($doc.Error) { $problems += "$($f.Name): $($doc.Error)"; continue }
                        $folder, $cls = Get-ItemFolder $doc.Message
                        $nid = $split.AddMessage($folder, $cls, $doc.Message)
                        $where[$f.Name] = @{ Nid = $nid; File = $split.Files[$split.Files.Count - 1] }
                        $bytes += $f.Length
                    }
                    $split.Close()
                } finally { $split.Dispose() }
                @{ Bytes = $bytes; Problems = $problems }
            }
            Measure-Phase 'corpus-pst-read' {
                $readers = @{}
                $problems = @(); $bytes = 0L
                try {
                    foreach ($f in $files) {
                        $w = $where[$f.Name]
                        if (-not $w) { continue }
                        if (-not $readers.ContainsKey($w.File)) { $readers[$w.File] = [MailboxTransfer.PstReader]::Open($w.File) }
                        $orig = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromFile($f.FullName)).Message
                        $copy = $readers[$w.File].ReadMessage([uint32]$w.Nid)
                        $d = [Collections.Generic.List[string]]::new()
                        [void][MailboxTransfer.PstFxCompare]::Compare($orig, $copy, $d)
                        if ($d.Count) { $problems += "$($f.Name): $($d.Count) difference(s), first: $($d[0])" }
                        $bytes += $f.Length
                    }
                } finally { foreach ($r in $readers.Values) { $r.Dispose() } }
                @{ Bytes = $bytes; Problems = $problems }
            }
            Measure-Phase 'corpus-pst-to-fx' {
                $readers = @{}
                $problems = @(); $bytes = 0L
                try {
                    foreach ($f in $files) {
                        $w = $where[$f.Name]
                        if (-not $w) { continue }
                        if (-not $readers.ContainsKey($w.File)) { $readers[$w.File] = [MailboxTransfer.PstReader]::Open($w.File) }
                        $src = $readers[$w.File].OpenMessageSource([uint32]$w.Nid)
                        $s = $src.OpenRead()
                        try { $s.CopyTo($sink, 81920) } finally { $s.Dispose() }
                        $bytes += $src.Length
                    }
                } finally { foreach ($r in $readers.Values) { $r.Dispose() } }
                @{ Bytes = $bytes; Problems = $problems }
            }
            Remove-Item -LiteralPath (Join-Path $WorkDir 'corpus*.pst') -Force -ErrorAction SilentlyContinue
        }
    }

    # ------------------------------------------------------------------------------------------------ synthetic
    if ($Scenario -in 'Synthetic', 'All') {
        $attach = [long]$SizeMiB * 1MB
        $bodyBytes = 8MB
        $innerBytes = 16MB
        $seedAttach = [uint64]1001; $seedBody = [uint64]1002; $seedInner = [uint64]1003

        function New-BigMessage {
            $m = [MailboxTransfer.FxMessage]::new()
            $m.Properties.Add($P::FromString(0x001A, 'IPM.Note'))
            $m.Properties.Add($P::FromString(0x0037, "Synthetic item with $SizeMiB MiB attachment"))
            $m.Properties.Add($P::FromInt32(0x0E07, 0x09))
            $m.Properties.Add($P::FromTime(0x0039, [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)))
            $body = [MailboxTransfer.FxConcatBlob]::new([MailboxTransfer.TestSupport.PatternBlob]::new($bodyBytes, $seedBody), [MailboxTransfer.FxBytesBlob]::Terminator)
            $m.Properties.Add($P::CreateLazy(0x1000, 0x001F, $body))
            $r = [MailboxTransfer.FxRecipient]::new()
            $r.Properties.Add($P::FromInt32(0x0C15, 1)); $r.Properties.Add($P::FromString(0x3001, 'Recipient')); $r.Properties.Add($P::FromString(0x39FE, 'r@example.test'))
            $m.Recipients.Add($r)
            $a = [MailboxTransfer.FxAttachment]::new()
            $a.Properties.Add($P::FromInt32(0x3705, 1)); $a.Properties.Add($P::FromString(0x3704, 'big.bin'))
            $a.Properties.Add($P::CreateLazy(0x3701, 0x0102, [MailboxTransfer.TestSupport.PatternBlob]::new($attach, $seedAttach)))
            $m.Attachments.Add($a)
            $e = [MailboxTransfer.FxMessage]::new()
            $e.Properties.Add($P::FromString(0x001A, 'IPM.Note')); $e.Properties.Add($P::FromString(0x0037, 'Embedded'))
            $ea = [MailboxTransfer.FxAttachment]::new()
            $ea.Properties.Add($P::FromInt32(0x3705, 1)); $ea.Properties.Add($P::FromString(0x3704, 'inner.bin'))
            $ea.Properties.Add($P::CreateLazy(0x3701, 0x0102, [MailboxTransfer.TestSupport.PatternBlob]::new($innerBytes, $seedInner)))
            $e.Attachments.Add($ea)
            $holder = [MailboxTransfer.FxAttachment]::new()
            $holder.Properties.Add($P::FromInt32(0x3705, 5)); $holder.Properties.Add($P::Create(0x3701, 0x000D, [byte[]]::new(0)))
            $holder.Embedded = $e
            $m.Attachments.Add($holder)
            $m
        }

        # a Graph-style envelope: type 1 (8 bytes), type 4 (6 bytes), then the data in records of 31,680 bytes
        $envelope = [MailboxTransfer.FxEnvelope]::new()
        foreach ($rec in @(@(1, 8), @(4, 6), @(2, 1))) {
            $e = [MailboxTransfer.FxEnvelopeEntry]::new()
            $e.Type = $rec[0]; $e.Length = $rec[1]
            if ($rec[0] -ne 2) { $e.Payload = [byte[]]::new($rec[1]) }
            $envelope.Entries.Add($e)
        }

        function Test-Attachments([MailboxTransfer.FxMessage]$m, [string]$what) {
            $bad = @()
            $data = $m.Attachments[0].Properties | Where-Object { $_.Id -eq 0x3701 }
            if (-not $data -or -not [MailboxTransfer.TestSupport.Streams]::Equal($data.Blob, [MailboxTransfer.TestSupport.PatternBlob]::new($attach, $seedAttach))) { $bad += "$what : the big attachment differs" }
            $inner = $m.Attachments[1].Embedded.Attachments[0].Properties | Where-Object { $_.Id -eq 0x3701 }
            if (-not $inner -or -not [MailboxTransfer.TestSupport.Streams]::Equal($inner.Blob, [MailboxTransfer.TestSupport.PatternBlob]::new($innerBytes, $seedInner))) { $bad += "$what : the embedded attachment differs" }
            $bad
        }

        $fxs = Join-Path $WorkDir 'synthetic.fxs'
        $pst = Join-Path $WorkDir 'synthetic.pst'
        $back = Join-Path $WorkDir 'synthetic-back.fxs'
        $expected = 0L
        $nid = 0

        Measure-Phase 'synthetic-create' {
            $src = [MailboxTransfer.FxMessageSource]::new((New-BigMessage), $envelope)
            $n = [MailboxTransfer.TestSupport.Streams]::CopyToFile($src, $fxs)
            $script:expected = $n
            @{ Bytes = $n; Problems = $(if ($n -ne $src.Length) { "length announced $($src.Length), written $n" } else { @() }) }
        }
        Measure-Phase 'synthetic-fx-to-pst' {
            $doc = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromFile($fxs))
            $problems = @()
            if ($doc.Error -or $doc.Warnings.Count) { $problems += "decode: $($doc.Error) $($doc.Warnings -join ' ')" }
            $w = [MailboxTransfer.PstWriter]::Create($pst, 'Synthetic')
            try { $script:nid = $w.AddMessage($w.IpmSubtree, $doc.Message) } finally { $w.Close(); $w.Dispose() }
            @{ Bytes = (Get-Item -LiteralPath $pst).Length; Problems = $problems }
        }
        Measure-Phase 'synthetic-pst-read' {
            $r = [MailboxTransfer.PstReader]::Open($pst)
            $problems = @()
            try {
                $copy = $r.ReadMessage([uint32]$script:nid)
                $orig = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromFile($fxs)).Message
                $d = [Collections.Generic.List[string]]::new()
                [void][MailboxTransfer.PstFxCompare]::Compare($orig, $copy, $d)
                if ($d.Count) { $problems += "$($d.Count) difference(s), first: $($d[0])" }
                $problems += Test-Attachments $copy 'from the PST'
            } finally { $r.Dispose() }
            @{ Bytes = (Get-Item -LiteralPath $pst).Length; Problems = $problems }
        }
        Measure-Phase 'synthetic-pst-to-fx' {
            $r = [MailboxTransfer.PstReader]::Open($pst)
            $problems = @()
            try {
                $src = $r.OpenMessageSource([uint32]$script:nid, $envelope)
                $n = [MailboxTransfer.TestSupport.Streams]::CopyToFile($src, $back)
                if ($n -ne $src.Length) { $problems += "length announced $($src.Length), written $n" }
            } finally { $r.Dispose() }
            $rt = [MailboxTransfer.FxCodec]::RoundTripFile($back)
            if (-not $rt.Identical) { $problems += "round trip of the stream made from the PST: $($rt.Note)" }
            $problems += Test-Attachments $rt.Document.Message 'from the stream made from the PST'
            @{ Bytes = $n; Problems = $problems }
        }
        Measure-Phase 'synthetic-import' {
            $r = [MailboxTransfer.PstReader]::Open($pst)
            $problems = @()
            try {
                $src = $r.OpenMessageSource([uint32]$script:nid, $envelope)
                $err = [MailboxTransfer.TestSupport.ImportBody]::Check($src, 'folder-id')
                if ($err) { $problems += $err }
            } finally { $r.Dispose() }
            @{ Bytes = $src.Length; Problems = $problems }
        }
    }
} finally {
    Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
}

if ($Json) { $phases | ConvertTo-Json -Depth 4 -AsArray }
elseif ($failed) { Write-Host 'FAILED' -ForegroundColor Red }
else { Write-Host 'All phases passed.' -ForegroundColor Green }
if ($failed) { exit 1 }
