#Requires -Version 7.4
<#
.SYNOPSIS
    Writes the FastTransfer streams of one or more folders into PST files with our own MS-PST writer, reads them back with our own
    reader and compares every message property by property. Prints one line per item and writes the PST files and a CSV report.

.DESCRIPTION
    For each corpus folder (default out\fxics-corpus and out\fxics-corpus2, git-ignored):
      1. every *.fxs is decoded with FxCodec (the exact decoder; a value above 64 KiB stays in the file and is read again in pieces,
         so the memory used does not depend on the size of the items),
      2. the messages are written into one PST (PstSplitWriter: folders per item class, split by size), in
         <OutDir>\<corpus name>.pst (and <name>-0002.pst, ... when the size limit is reached),
      3. the PST is opened with PstReader (an independent implementation of the reader side: it checks every CRC, signature and
         B-tree order it meets), every message is read back and compared with the decoded original (PstFxCompare),
      4. optionally (-Check) the integrity checker of the PST is run on each file.
    The report lists, per item: file, class, folder, properties, recipients, attachments, embedded depth, size, status (OK / DIFF /
    ERROR) and the first differences. What a PST intentionally drops or recomputes is NOT a difference (see PstFxCompare and
    PstConvert.NotStored): the properties computed from the node (entry ids, record keys, object type, has-attachments flag), the
    Exchange identifiers (folder id, mid, change number), the size properties (recomputed), the attachment number (position),
    string8 / code page strings (stored as Unicode), and the default values added when missing.

    Exit code: 0 if every item is OK, 1 otherwise.

.PARAMETER Path
    Corpus folders (default <repository>\out\fxics-corpus and \out\fxics-corpus2).

.PARAMETER OutDir
    Where the PST files and the report go (default <repository>\out\pst-validation, git-ignored).

.PARAMETER MaxSizeMiB
    Maximum size of one PST in MiB (default 51200 = 50 GiB). Lower it to test the split.

.PARAMETER Check
    Also run the integrity checker (Test-MbtPst) on each PST written.

.EXAMPLE
    .\tools\Test-PstRoundTrip.ps1
    .\tools\Test-PstRoundTrip.ps1 -Path .\out\fxics-corpus2 -MaxSizeMiB 1

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.
#>
[CmdletBinding()]
param(
    [string[]]$Path = @((Join-Path $PSScriptRoot '..\out\fxics-corpus'), (Join-Path $PSScriptRoot '..\out\fxics-corpus2')),
    [string]$OutDir = (Join-Path $PSScriptRoot '..\out\pst-validation'),
    [long]$MaxSizeMiB = 51200,
    [switch]$Check
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$package = Join-Path $PSScriptRoot '..\package'
Import-Module (Join-Path $package 'MailboxTransfer.psd1') -Force
Initialize-MbtEngine -Root (Resolve-Path $package).Path
$null = New-Item -ItemType Directory -Force -Path $OutDir
$OutDir = (Resolve-Path $OutDir).Path

function Get-ItemFolder([MailboxTransfer.FxMessage]$m) {
    $cls = $m.GetString(0x001A)
    if (-not $cls) { $cls = 'IPM.Note' }
    if ($cls -like 'IPM.Appointment*' -or $cls -like 'IPM.Schedule*') { return 'Calendar', 'IPF.Appointment' }
    if ($cls -like 'IPM.Contact*') { return 'Contacts', 'IPF.Contact' }
    if ($cls -like 'IPM.Task*') { return 'Tasks', 'IPF.Task' }
    if ($cls -like 'IPM.StickyNote*') { return 'Notes', 'IPF.StickyNote' }
    return 'Mail', 'IPF.Note'
}

$allOk = $true
$report = [Collections.Generic.List[object]]::new()
foreach ($corpus in $Path) {
    if (-not (Test-Path -LiteralPath $corpus)) { Write-Warning "Not found: $corpus"; continue }
    $corpus = (Resolve-Path $corpus).Path
    $name = Split-Path $corpus -Leaf
    $files = @(Get-ChildItem -LiteralPath $corpus -Filter '*.fxs' -File | Sort-Object Name)
    if ($files.Count -eq 0) { Write-Warning "No *.fxs in $corpus"; continue }
    Write-Host ''
    Write-Host "== $name : $($files.Count) streams ($([math]::Round(($files | Measure-Object Length -Sum).Sum / 1MB, 1)) MiB) ==" -ForegroundColor Cyan
    $pstPath = Join-Path $OutDir (($name -replace '^fxics-', '' -replace '-', '') + '.pst')
    $stem = [IO.Path]::GetFileNameWithoutExtension($pstPath)
    foreach ($old in Get-ChildItem -LiteralPath $OutDir -Filter '*.pst' | Where-Object { $_.BaseName -eq $stem -or $_.BaseName -match ('^' + [regex]::Escape($stem) + '-\d{4}$') }) { Remove-Item -LiteralPath $old.FullName -Force }

    $sw = [Diagnostics.Stopwatch]::StartNew()
    $split = [MailboxTransfer.PstSplitWriter]::new($pstPath, "MailboxTransfer $name", $MaxSizeMiB * 1MB)
    $where = @{}
    try {
        foreach ($f in $files) {
            $doc = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromFile($f.FullName))
            if ($doc.Error) { $where[$f.Name] = @{ Error = "decode: $($doc.Error)" }; continue }
            $folder, $cls = Get-ItemFolder $doc.Message
            $nid = $split.AddMessage($folder, $cls, $doc.Message)
            $where[$f.Name] = @{ Nid = $nid; File = $split.Files[$split.Files.Count - 1]; Folder = $folder
                Class = $doc.Message.GetString(0x001A); Props = $doc.Message.CountProperties(); Recipients = $doc.Message.Recipients.Count
                Attachments = $doc.Message.Attachments.Count; Depth = $doc.Message.EmbeddedDepth(); Bytes = $f.Length }
            $doc = $null
        }
        $split.Close()
    } finally { $split.Dispose() }
    $writeSeconds = $sw.Elapsed.TotalSeconds
    $peak = [math]::Round([Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64 / 1MB)
    $pstFiles = @($split.Files)
    Write-Host ("written {0} PST file(s) in {1:N1} s (peak working set {2} MiB):" -f $pstFiles.Count, $writeSeconds, $peak)
    foreach ($p in $pstFiles) { Write-Host ("  {0}  {1:N0} bytes" -f $p, (Get-Item -LiteralPath $p).Length) }

    $sw.Restart()
    $readers = @{}
    foreach ($p in $pstFiles) { $readers[$p] = [MailboxTransfer.PstReader]::Open($p) }
    try {
        $n = 0
        foreach ($f in $files) {
            $n++
            $w = $where[$f.Name]
            $row = [ordered]@{ Corpus = $name; Item = $f.Name; Class = ''; Folder = ''; Pst = ''; Props = 0; Recipients = 0; Attachments = 0; Depth = 0; Bytes = $f.Length; Status = 'OK'; Differences = '' }
            if ($w.ContainsKey('Error')) { $row.Status = 'ERROR'; $row.Differences = $w.Error }
            else {
                $row.Class = $w.Class; $row.Folder = $w.Folder; $row.Pst = Split-Path $w.File -Leaf; $row.Props = $w.Props; $row.Recipients = $w.Recipients
                $row.Attachments = $w.Attachments; $row.Depth = $w.Depth
                try {
                    $orig = [MailboxTransfer.FxCodec]::Decode([MailboxTransfer.FxStreamSource]::FromFile($f.FullName)).Message
                    $copy = $readers[$w.File].ReadMessage([uint32]$w.Nid)
                    $d = [Collections.Generic.List[string]]::new()
                    [void][MailboxTransfer.PstFxCompare]::Compare($orig, $copy, $d)
                    if ($d.Count) { $row.Status = 'DIFF'; $row.Differences = "$($d.Count): " + (($d | Select-Object -First 5) -join ' | ') }
                    $orig = $null; $copy = $null
                } catch { $row.Status = 'ERROR'; $row.Differences = $_.Exception.Message }
            }
            if ($row.Status -ne 'OK') { $allOk = $false }
            $report.Add([pscustomobject]$row)
            $color = if ($row.Status -eq 'OK') { 'Green' } else { 'Red' }
            Write-Host ("  {0,-18} {1,-6} {2,-9} {3,-26} props {4,4} rcpt {5,3} att {6,2} depth {7} {8,11:N0} B  {9}" -f $row.Item, $row.Status, $row.Folder, $row.Class, $row.Props, $row.Recipients, $row.Attachments, $row.Depth, $row.Bytes, $row.Differences) -ForegroundColor $color
        }
        # folder counts: the contents tables agree with what was written
        foreach ($p in $pstFiles) {
            $r = $readers[$p]
            foreach ($fo in $r.AllFolders()) {
                $rows = $r.MessageNids($fo, $false).Count
                if ($rows -ne $fo.ContentCount) { $allOk = $false; Write-Host "  folder '$($fo.Path)' in $(Split-Path $p -Leaf): PR_CONTENT_COUNT $($fo.ContentCount) but $rows rows" -ForegroundColor Red }
            }
        }
    } finally { foreach ($r in $readers.Values) { $r.Dispose() } }
    Write-Host ("read back and compared in {0:N1} s" -f $sw.Elapsed.TotalSeconds)

    if ($Check) {
        foreach ($p in $pstFiles) {
            $res = Test-MbtPst -Path $p
            Write-Host ("  check {0}: {1} error(s), {2} warning(s)" -f (Split-Path $p -Leaf), $res.Errors, $res.Warnings)
            if ($res.Errors -gt 0) { $allOk = $false }
        }
    }
}

$csv = Join-Path $OutDir 'roundtrip-report.csv'
$report | Export-Csv -LiteralPath $csv -NoTypeInformation -Encoding UTF8
$ok = @($report | Where-Object Status -EQ 'OK').Count
Write-Host ''
Write-Host ("{0} of {1} items identical (report: {2})" -f $ok, $report.Count, $csv) -ForegroundColor $(if ($allOk) { 'Green' } else { 'Red' })
if (-not $allOk) { exit 1 }
