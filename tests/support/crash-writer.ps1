# Child process of Store.Tests.ps1: writes into an export until the parent kills it (Stop-Process -Force).
#   -Mode Stall : commits $Committed items, then starts one more item whose source blocks after 30 MiB (after writing the signal file).
#   -Mode Load  : writes every pending item with $Threads concurrent writers; writes the signal file as soon as it starts.
# Item #key has seed = key and length = [Workload]::CrashSize(key) (1 to 16 MiB), except the stalled one.
param(
    [Parameter(Mandatory)][string]$PackageRoot,
    [Parameter(Mandatory)][string]$ExportPath,
    [Parameter(Mandatory)][string]$SignalFile,
    [ValidateSet('Stall', 'Load')][string]$Mode = 'Stall',
    [string]$Mailbox = 'mb1',
    [int]$Threads = 4,
    [int]$Committed = 5
)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'Import-StoreEngine.ps1')
Import-StoreEngine -PackageRoot $PackageRoot -WithTestSupport

$store = [MailboxTransfer.ExportStore]::Open($ExportPath)
$jobs = [System.Collections.Generic.List[MailboxTransfer.TestSupport.Job]]::new()
$after = 0L
while ($true) {
    $page = $store.GetPendingItems($Mailbox, 500, $after)
    if ($page.Count -eq 0) { break }
    foreach ($i in $page) {
        $j = [MailboxTransfer.TestSupport.Job]::new()
        $j.Key = $i.Key; $j.Seed = $i.Key; $j.Length = [MailboxTransfer.TestSupport.Workload]::CrashSize($i.Key)
        $jobs.Add($j)
    }
    $after = $page[$page.Count - 1].Key
}

if ($Mode -eq 'Stall') {
    for ($n = 0; $n -lt $Committed; $n++) {
        $j = $jobs[$n]
        $s = [MailboxTransfer.TestSupport.GenStream]::new($j.Seed, $j.Length)
        [void]$store.WriteItem($j.Key, $s, [Threading.CancellationToken]::None)
    }
    $j = $jobs[$Committed]
    $stall = [MailboxTransfer.TestSupport.StallStream]::new($j.Seed, 100MB, 30MB, $SignalFile)
    [void]$store.WriteItem($j.Key, $stall, [Threading.CancellationToken]::None)   # blocks until killed
}
else {
    Set-Content -LiteralPath $SignalFile -Value 'started'
    $r = [MailboxTransfer.TestSupport.Runner]::WriteAll($store, $jobs, $Threads, $false)
    Set-Content -LiteralPath "$SignalFile.finished" -Value "finished before being killed: $($r.Count) items"
}
