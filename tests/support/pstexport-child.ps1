# Child process of tests\PstExport.Tests.ps1 and tools\Measure-PstExport.ps1: converts a store into PST files in a fresh pwsh (so that the
# process can be killed in the middle of a part, and its memory measured). Prints one line "RESULT {json}" at the end.
param(
    [Parameter(Mandatory)][string]$StorePath,
    [Parameter(Mandatory)][string]$Dest,
    [long]$MaxMiB = 51200,
    # sleep this long after each item written (the tests kill the process while a part is being written)
    [int]$DelayMs = 0,
    [switch]$SkipCheck,
    [switch]$SeparateAux,
    [switch]$Force
)
Set-StrictMode -Version Latest
[AppDomain]::CurrentDomain.SetData('MailboxTransfer.UnverifiedTreeWarned', $true)   # a source checkout has no PackageIntegrity.psd1: no warning in the output that the test reads
$ErrorActionPreference = 'Stop'
$package = Join-Path (Split-Path (Split-Path $PSScriptRoot -Parent) -Parent) 'package'
Import-Module (Join-Path $package 'MailboxTransfer.psd1')
Initialize-MbtEngine -Root $package
$store = Open-MbtExportStore -Path $StorePath -ReadOnly
try {
    $o = [MailboxTransfer.PstExportOptions]::new()
    $o.Destination = $Dest
    $o.MaxPartBytes = $MaxMiB * 1MB
    $o.Check = -not $SkipCheck
    $o.MergeAuxIntoArchive = -not $SeparateAux
    $o.Force = [bool]$Force
    $script:lastDone = -1
    $script:peakDuring = 0
    $sink = [MailboxTransfer.PstExportProgressSink]::new([Action[MailboxTransfer.PstExportProgress]]{
            param($e)
            if ($e.Phase -eq 'write' -and $e.ItemsDone -ne $script:lastDone) {
                $script:lastDone = $e.ItemsDone
                if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
            }
        })
    $baseline = [Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64
    $watch = [Diagnostics.Stopwatch]::StartNew()
    $r = [MailboxTransfer.PstExporter]::Run($store, $o, $sink, [Threading.CancellationToken]::None)
    $peak = [Diagnostics.Process]::GetCurrentProcess().PeakWorkingSet64
    $out = [ordered]@{
        Complete = $r.Complete; Cancelled = $r.Cancelled; Error = $r.Error; Parts = $r.Parts.Count; ItemsWritten = $r.ItemsWritten; ItemsReused = $r.ItemsReused
        PartsReused = $r.PartsReused; PartsWithErrors = $r.PartsWithErrors; Seconds = [math]::Round($watch.Elapsed.TotalSeconds, 1)
        BaselinePeakMiB = [math]::Round($baseline / 1MB); PeakMiB = [math]::Round($peak / 1MB); PeakPrivateMiB = [math]::Round([Diagnostics.Process]::GetCurrentProcess().PeakPagedMemorySize64 / 1MB)
        Oversize = $r.Oversize.Count; Failed = $r.Failed.Count
        PartDetail = @($r.Parts | ForEach-Object { [ordered]@{ File = $_.File; Bytes = $_.Bytes; Items = $_.ItemCount; Folders = $_.FolderCount; WriteSeconds = [math]::Round($_.Seconds, 1); CheckSeconds = [math]::Round($_.CheckSeconds, 1); Checked = $_.Checked; CheckErrors = $_.CheckErrors; CheckWarnings = $_.CheckWarnings; Reused = $_.Reused } })
    }
    'RESULT ' + ($out | ConvertTo-Json -Compress -Depth 5)
} finally { $store.Dispose() }
