#Requires -Version 7.4
<#
.SYNOPSIS
    Repairs a big damaged synthetic PST (2 GB and more) and measures the memory of the repair: the working set and the private bytes of
    this process, sampled every 10 ms, for each repair (the checks of the source and of the result are inside the measure).

.DESCRIPTION
    The memory used by a repair must not depend on the size of the file. The tool writes a PST made of many small messages and a few
    messages with big attachments (generated on the fly, nothing of it is ever in memory as a whole), damages it in place (random bytes
    in the data, then a leaf page of the node B-tree), and repairs it:
      repair-recompute   level Recompute: the damaged values are dropped, the other items kept (the file opens fine)
      repair-salvage     the node B-tree is damaged too: Recompute cannot open the file, the repair goes on with Salvage (scan of the file)
    Each phase checks the result with the integrity checker of the engine (no error) and the number of items kept, and reports the growth of the
    working set after a full GC at the start. Exit code 0 when everything passed and every growth is under -LimitMiB.

.PARAMETER SizeMiB
    Total size of the big attachments (default 2200 = 2.15 GiB; the source, the repaired file and the check take that much disk each).

.PARAMETER SmallMessages
    Number of small messages (default 20000).

.PARAMETER WorkDir
    Where the files are written (default a new folder in the temporary directory, removed at the end).

.PARAMETER LimitMiB
    Growth of the working set above which a phase is reported as failed (default 250).

.EXAMPLE
    .\tools\Measure-PstRepair.ps1
    .\tools\Measure-PstRepair.ps1 -SizeMiB 6000 -WorkDir D:\scratch

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.
#>
[CmdletBinding()]
param(
    [ValidateRange(300, 100000)][int]$SizeMiB = 2200,
    [ValidateRange(0, 1000000)][int]$SmallMessages = 20000,
    [string]$WorkDir,
    [int]$LimitMiB = 250,
    [switch]$Json
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
. (Join-Path $repo 'tests\support\Import-StoreEngine.ps1')
Import-StoreEngine -PackageRoot (Join-Path $repo 'package') -WithTestSupport

if (-not $WorkDir) { $WorkDir = Join-Path ([IO.Path]::GetTempPath()) ('mbt-pst-repair-' + [guid]::NewGuid().ToString('N').Substring(0, 8)) }
$null = New-Item -ItemType Directory -Force -Path $WorkDir
$WorkDir = (Resolve-Path $WorkDir).Path

$phases = [Collections.Generic.List[object]]::new()
$failed = $false
$P = [MailboxTransfer.FxProperty]

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
    if ($detail -and $detail.ContainsKey('Note')) { $row.Note = ($detail.Note + ' ' + $row.Note).Trim() }
    if ($row.WorkingSetGrowthMiB -gt $LimitMiB) { $row.Ok = $false; $row.Note = ($row.Note + " working set grew by $($row.WorkingSetGrowthMiB) MiB (limit $LimitMiB)").Trim() }
    $script:phases.Add([pscustomobject]$row)
    if (-not $row.Ok) { $script:failed = $true }
    if (-not $Json) {
        $color = if ($row.Ok) { 'Green' } else { 'Red' }
        Write-Host ('{0,-18} {1,7:N1} s  working set +{2,4} MiB (peak {3,5}), private +{4,4} MiB  {5:N0} bytes  {6}' -f $row.Phase, $row.Seconds, $row.WorkingSetGrowthMiB, $row.PeakWorkingSetMiB, $row.PrivateGrowthMiB, $row.Bytes, $row.Note) -ForegroundColor $color
    }
}

function Set-Bytes([string]$Path, [long[]]$Positions) {
    $fs = [IO.File]::Open($Path, 'Open', 'ReadWrite')
    try { foreach ($pos in $Positions) { $fs.Position = $pos; $b = $fs.ReadByte(); $fs.Position = $pos; $fs.WriteByte([byte]($b -bxor 0x5A)) } } finally { $fs.Dispose() }
}

# the first leaf page of the node B-tree, found from the header with plain reads
function Get-NbtLeafOffset([string]$Path) {
    $fs = [IO.File]::OpenRead($Path)
    try {
        $buf = [byte[]]::new(512)
        $fs.Position = 224; $fs.ReadExactly($buf, 0, 8); $ib = [BitConverter]::ToInt64($buf, 0)
        while ($true) {
            $fs.Position = $ib; $fs.ReadExactly($buf, 0, 512)
            if ($buf[491] -eq 0) { return $ib }
            $ib = [BitConverter]::ToInt64($buf, 16)
        }
    } finally { $fs.Dispose() }
}

$src = Join-Path $WorkDir 'big.pst'
$expected = $SmallMessages + 3
$perAttach = [long]($SizeMiB * 1MB / 3)

try {
    Measure-Phase 'build' {
        $w = [MailboxTransfer.PstWriter]::Create($src, 'Big')
        try {
            $mail = $w.CreateFolder($null, 'Mail', 'IPF.Note')
            for ($i = 1; $i -le $SmallMessages; $i++) {
                $m = [MailboxTransfer.FxMessage]::new()
                $m.Properties.Add($P::FromString(0x001A, 'IPM.Note')); $m.Properties.Add($P::FromString(0x0037, "mail $i"))
                $m.Properties.Add($P::FromString(0x1000, ('body of message ' + $i + ' ') * 40)); $m.Properties.Add($P::FromInt32(0x0E07, $(if ($i % 5 -eq 0) { 0 } else { 1 })))
                $r = [MailboxTransfer.FxRecipient]::new(); $r.Properties.Add($P::FromInt32(0x0C15, 1)); $r.Properties.Add($P::FromString(0x3001, 'Bob'))
                $m.Recipients.Add($r)
                [void]$w.AddMessage($mail, $m)
            }
            $big = $w.CreateFolder($null, 'Big', 'IPF.Note')
            for ($i = 1; $i -le 3; $i++) {
                $m = [MailboxTransfer.FxMessage]::new()
                $m.Properties.Add($P::FromString(0x001A, 'IPM.Note')); $m.Properties.Add($P::FromString(0x0037, "big $i"))
                $a = [MailboxTransfer.FxAttachment]::new()
                $a.Properties.Add($P::FromInt32(0x3705, 1)); $a.Properties.Add($P::FromString(0x3704, "big$i.bin"))
                $a.Properties.Add($P::CreateLazy(0x3701, 0x0102, [MailboxTransfer.TestSupport.PatternBlob]::new($perAttach, [uint64](2000 + $i))))
                $m.Attachments.Add($a)
                [void]$w.AddMessage($big, $m)
            }
        } finally { $w.Close(); $w.Dispose() }
        @{ Bytes = (Get-Item -LiteralPath $src).Length; Problems = @() }
    }
    $len = (Get-Item -LiteralPath $src).Length
    # damage 1: 30 random bytes among the small messages (the first 40 MB), 1 byte in the middle of the second big attachment (that value
    # is dropped, the two others - 1.4 GB - are verified and copied block by block)
    $rnd = [Random]::new(31)
    $positions = [long[]]((1..30 | ForEach-Object { $rnd.NextInt64(0x4400, 40MB) }) + @([long](40MB + 1.5 * $perAttach)))
    Set-Bytes $src $positions

    $o = [MailboxTransfer.PstRepairOptions]::new(); $o.Overwrite = $true; $o.MaxItemRecords = 1000
    $out = Join-Path $WorkDir 'fixed1.pst'
    Measure-Phase 'repair-recompute' {
        $o.Level = 'Recompute'
        $r = [MailboxTransfer.PstRepair]::Repair($src, $out, $o)
        $pr = @()
        if (-not $r.Success) { $pr += "repair failed: $($r.Error) after=$($r.After.Errors) " + (($r.After.Findings | Where-Object Severity -NE 'Info' | Select-Object -First 3 | ForEach-Object { $_.ToString() }) -join ' | ') }
        if ($r.ItemsTotal -ne $expected -and $r.ItemsTotal -ne $expected - 0) { $pr += "items $($r.ItemsTotal), expected $expected" }
        if ($r.ItemsRecovered + $r.ItemsPartial -lt $expected - 100) { $pr += "only $($r.ItemsRecovered + $r.ItemsPartial) items kept" }
        if ($r.TargetLength -lt 1.3GB) { $pr += "the result is only $($r.TargetLength) bytes: the intact big attachments were not copied" }
        @{ Bytes = $r.TargetLength; Problems = $pr; Note = "level $($r.Used): $($r.ItemsRecovered) recovered, $($r.ItemsPartial) partial, $($r.ItemsLost) lost; before $($r.Before.Errors)e/$($r.Before.Warnings)w, after $($r.After.Errors)e/$($r.After.Warnings)w" }
    }
    Remove-Item -LiteralPath $out -Force -ErrorAction SilentlyContinue

    # damage 2: a leaf page of the node B-tree
    $leaf = Get-NbtLeafOffset $src
    Set-Bytes $src ([long[]]@($leaf + 5))
    $out2 = Join-Path $WorkDir 'fixed2.pst'
    Measure-Phase 'repair-salvage' {
        $o.Level = 'Recompute'
        $r = [MailboxTransfer.PstRepair]::Repair($src, $out2, $o)
        $pr = @()
        if (-not $r.Success) { $pr += "repair failed: $($r.Error) after=$($r.After.Errors) " + (($r.After.Findings | Where-Object Severity -NE 'Info' | Select-Object -First 3 | ForEach-Object { $_.ToString() }) -join ' | ') }
        if ($r.Used -ne 'Salvage') { $pr += "level used $($r.Used), expected Salvage" }
        if ($r.ItemsRecovered + $r.ItemsPartial -lt $expected - 100) { $pr += "only $($r.ItemsRecovered + $r.ItemsPartial) items kept of $expected" }
        if ($r.TargetLength -lt 1.3GB) { $pr += "the result is only $($r.TargetLength) bytes: the intact big attachments were not copied" }
        @{ Bytes = $r.TargetLength; Problems = $pr; Note = "level $($r.Used): $($r.ItemsRecovered) recovered, $($r.ItemsPartial) partial, $($r.ItemsLost) lost; before $($r.Before.Errors)e/$($r.Before.Warnings)w, after $($r.After.Errors)e/$($r.After.Warnings)w" }
    }
} finally {
    Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue
}

if ($Json) { $phases | ConvertTo-Json -Depth 4 -AsArray }
elseif ($failed) { Write-Host 'FAILED' -ForegroundColor Red }
else { Write-Host 'All phases passed.' -ForegroundColor Green }
if ($failed) { exit 1 }
