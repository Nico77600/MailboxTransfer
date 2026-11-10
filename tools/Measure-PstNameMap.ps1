#Requires -Version 7.4
<#
.SYNOPSIS
    Measures the PST repair (level Salvage) on randomly damaged copies of a PST: items kept, partial items, and how many partial items
    are partial only because of named properties (the name-to-id map was lost or damaged).

.DESCRIPTION
    The campaign of tests\PstRepair.Corpus.Tests.ps1: for each copy i (0 .. Copies-1) a Random seeded with 900 + i changes -Bytes bytes of
    the file at random places, the copy is repaired with Salvage, and the report is read. Same seeds as the test, so the figures are
    comparable from one version of the engine to the next. Nothing is written outside -WorkDir (removed at the end).
    Columns: kept = recovered + partial; perfect = copies with every item recovered; partial items split into "named only" (every loss of
    the item is a named property without entry in the name map), "named + other" and "other"; the source of the name map per copy
    (Original / RecoveredByScan / RebuiltFromUsage / None) when the engine reports it.

.PARAMETER Path
    The clean PST to damage (default out\pst-validation\corpus2.pst).

.PARAMETER Bytes
    Number of bytes changed per copy (default 5 and 30, one campaign each).

.PARAMETER Copies
    Number of copies per campaign (default 60 for 5 bytes, 100 for 30 bytes).

.EXAMPLE
    .\tools\Measure-PstNameMap.ps1
    .\tools\Measure-PstNameMap.ps1 -Bytes 30 -Copies 20

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.
#>
[CmdletBinding()]
param(
    [string]$Path,
    [int[]]$Bytes = @(5, 30),
    [int[]]$Copies = @(60, 100),
    [string]$WorkDir
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$repo = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if (-not $Path) { $Path = Join-Path $repo 'out\pst-validation\corpus2.pst' }
. (Join-Path $repo 'tests\support\Import-StoreEngine.ps1')
Import-StoreEngine -PackageRoot (Join-Path $repo 'package')
if (-not $WorkDir) { $WorkDir = Join-Path ([IO.Path]::GetTempPath()) ('mbt-pst-namemap-' + [guid]::NewGuid().ToString('N').Substring(0, 8)) }
$null = New-Item -ItemType Directory -Force -Path $WorkDir
function Get-NamedIdentities([string]$File) {
    # every (GUID, LID or name) a named property of a message of the file stands for
    $ids = [Collections.Generic.HashSet[string]]::new()
    $r = [MailboxTransfer.PstReader]::Open($File)
    try {
        foreach ($f in $r.AllFolders()) {
            foreach ($assoc in $false, $true) {
                foreach ($nid in $r.MessageNids($f, $assoc)) {
                    foreach ($p in $r.ReadMessage($nid).Properties) {
                        if ($p.IsNamed) { [void]$ids.Add($(if ($null -ne $p.NamedLid) { '{0}/L{1:X}' -f $p.NamedGuid, $p.NamedLid } else { '{0}/N{1}' -f $p.NamedGuid, $p.NamedName })) }
                    }
                }
            }
        }
    } finally { $r.Dispose() }
    , $ids
}
try {
    $cleanIds = Get-NamedIdentities $Path
    $total = [MailboxTransfer.PstReader]::Open($Path)
    $n = 0; foreach ($f in $total.AllFolders()) { $n += @($total.MessageNids($f, $false)).Count + @($total.MessageNids($f, $true)).Count }
    $total.Dispose()
    for ($c = 0; $c -lt $Bytes.Count; $c++) {
        $nb = $Bytes[$c]; $nCopies = $Copies[[math]::Min($c, $Copies.Count - 1)]
        $wrong = 0; $kept = 0; $perfect = 0; $lost = 0; $partial = 0; $pNamed = 0; $pMixed = 0; $pOther = 0; $recovered = 0
        $sources = @{}
        for ($i = 0; $i -lt $nCopies; $i++) {
            $p = Join-Path $WorkDir "rd-$nb-$i.pst"
            $raw = [IO.File]::ReadAllBytes($Path); $rnd = [Random]::new(900 + $i)
            for ($k = 0; $k -lt $nb; $k++) { $raw[$rnd.Next(0, $raw.Length)] = [byte]$rnd.Next(0, 256) }
            [IO.File]::WriteAllBytes($p, $raw)
            $o = [MailboxTransfer.PstRepairOptions]::new(); $o.Level = 'Salvage'; $o.Overwrite = $true; $o.CheckBefore = $false; $o.CheckAfter = $false
            $out = Join-Path $WorkDir "rd-$nb-$i-fixed.pst"
            $r = [MailboxTransfer.PstRepair]::Repair($p, $out, $o)
            if ($r.Error) { Write-Warning "copy $i : $($r.Error)"; continue }
            $kept += $r.ItemsRecovered + $r.ItemsPartial; $lost += $r.ItemsLost; $recovered += $r.ItemsRecovered
            if ($r.ItemsRecovered -eq $n) { $perfect++ }
            foreach ($it in $r.Items) {
                if ($it.Status -ne 'Partial') { continue }
                $partial++
                $named = @($it.Lost | Where-Object { $_ -match 'named property' }).Count
                if ($named -eq $it.Lost.Count) { $pNamed++ } elseif ($named -gt 0) { $pMixed++ } else { $pOther++ }
            }
            $src = if ($r.PSObject.Properties['NameMapSource']) { [string]$r.NameMapSource } else { 'n/a' }
            $sources[$src] = 1 + $(if ($sources.ContainsKey($src)) { $sources[$src] } else { 0 })
            foreach ($id in (Get-NamedIdentities $out)) { if (-not $cleanIds.Contains($id)) { $wrong++ } }
            Remove-Item -LiteralPath $p, $out -Force -ErrorAction SilentlyContinue
        }
        $tot = [double]($n * $nCopies)
        Write-Host ("{0} bytes x {1} copies, Salvage, {2} items per copy: kept {3:P1} (recovered fully {4:P1}, partial {5:P1}), lost {6:P1}, {7} copies with every item recovered" -f $nb, $nCopies, $n, ($kept / $tot), ($recovered / $tot), ($partial / $tot), ($lost / $tot), $perfect)
        Write-Host ("    partial items: {0}  named properties only {1}, named + other {2}, other only {3}" -f $partial, $pNamed, $pMixed, $pOther)
        Write-Host ("    named-property identities in the results that are not in the clean file (a wrong GUID, number or name): {0}" -f $wrong)
        Write-Host ("    name map: " + (($sources.GetEnumerator() | Sort-Object Name | ForEach-Object { "$($_.Name) $($_.Value)" }) -join ', '))
    }
} finally { Remove-Item -LiteralPath $WorkDir -Recurse -Force -ErrorAction SilentlyContinue }
