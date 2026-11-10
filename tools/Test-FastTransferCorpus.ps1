#Requires -Version 7.4
<#
.SYNOPSIS
    Decodes every FastTransfer stream of a folder with the exact decoder (FxCodec), encodes it again and checks that the bytes
    are identical. Reports, per item, the number of properties, recipients, attachments, the embedded depth and the round
    trip (OK, or KO with the first differing offset).

.DESCRIPTION
    The streams are the files exported by Graph (envelope + MS-OXCFXICS stream), for instance copied out of the packs of an
    export store into a folder that is NOT committed (default <repository>\out\fxics-corpus, git-ignored).

    For each file: FxCodec.RoundTripFile decodes it (a value above 64 KiB stays in the file, read again in pieces: a 100 MiB attachment takes no memory),
    encodes the document again and compares with the original while writing. A decoder warning (anything the syntax of
    MS-OXCFXICS 2.2.4.2 does not allow) or error makes the item KO even when the bytes are identical.

    At the end: the totals, and the features seen in the corpus (property types, multi-valued, named properties by LID and by
    name, code pages, extended properties, markers, embedded messages, wrapper, envelope records), so that what the corpus
    does NOT cover is known.

    Exit code: 0 if every item is OK (and at least one item was found), 1 otherwise.

.PARAMETER Path
    Folder with the streams (default <repository>\out\fxics-corpus), or one file.

.PARAMETER Filter
    File name filter (default *.fxs).

.PARAMETER NoCoverage
    Do not print the features seen.

.EXAMPLE
    .\tools\Test-FastTransferCorpus.ps1

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.
#>
[CmdletBinding()]
param(
    [string]$Path = (Join-Path $PSScriptRoot '..\out\fxics-corpus'),
    [string]$Filter = '*.fxs',
    [switch]$NoCoverage
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$package = Join-Path $PSScriptRoot '..\package'
Import-Module (Join-Path $package 'MailboxTransfer.psd1') -Force
Initialize-MbtEngine -Root (Resolve-Path $package).Path

if (-not (Test-Path -LiteralPath $Path)) { Write-Warning "Nothing at $Path"; exit 1 }
$files = @(if ((Get-Item -LiteralPath $Path).PSIsContainer) { Get-ChildItem -LiteralPath $Path -Filter $Filter -File | Sort-Object Name } else { Get-Item -LiteralPath $Path })
if ($files.Count -eq 0) { Write-Warning "No file matching $Filter in $Path"; exit 1 }

$seen = @{
    Types = [Collections.Generic.SortedDictionary[string, int]]::new()
    Named = [Collections.Generic.SortedDictionary[string, int]]::new()
    Markers = [Collections.Generic.SortedDictionary[string, int]]::new()
    Other = [Collections.Generic.SortedDictionary[string, int]]::new()
}
function Add-Seen([string]$Kind, [string]$Key) {
    $d = $seen[$Kind]
    if ($d.ContainsKey($Key)) { $d[$Key]++ } else { $d[$Key] = 1 }
}

function Measure-Message($Message, [int]$Depth) {
    $all = [Collections.Generic.List[object]]::new()
    $all.AddRange([object[]]@($Message.Properties))
    if ($Message.DelPropRecipients) { $all.Add($Message.DelPropRecipients); Add-Seen Other 'FXDelProp (recipients)' }
    if ($Message.DelPropAttachments) { $all.Add($Message.DelPropAttachments); Add-Seen Other 'FXDelProp (attachments)' }
    foreach ($r in $Message.Recipients) { $all.AddRange([object[]]@($r.Properties)); Add-Seen Markers 'StartRecip/EndToRecip' }
    foreach ($a in $Message.Attachments) {
        $all.AddRange([object[]]@($a.Properties)); Add-Seen Markers 'NewAttach/EndAttach'
        if ($a.Embedded) { Add-Seen Markers ('StartEmbed/EndEmbed (level ' + ($Depth + 1) + ')'); Measure-Message $a.Embedded ($Depth + 1) }
    }
    foreach ($p in $all) {
        $kind = if ($p.TypeInfo.CodePageString) { 'CodePage ' + $p.TypeInfo.CodePage } else { '0x' + $p.Type.ToString('X4') + ' ' + [MailboxTransfer.FxNames]::TypeName($p.Type) }
        if ($p.TypeInfo.CodePageString -and $p.TypeInfo.MultiValued) { $kind += ' (multi-valued)' }
        Add-Seen Types $kind
        if ($p.IsNamed) { Add-Seen Named ($(if ($null -ne $p.NamedLid) { 'LID' } else { 'name' })) }
        if ($p.Extended) { Add-Seen Other ('extended type (0x4000) flags 0x' + $p.ExtendedFlags.ToString('X4')) }
        if ($p.BareObject) { Add-Seen Other 'PtypObject without length' }
    }
    if ($Message.StartMarker) { Add-Seen Other ('wrapper 0x' + $Message.StartMarker.ToString('X8')) }
}

$rows = foreach ($f in $files) {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $rt = [MailboxTransfer.FxCodec]::RoundTripFile($f.FullName)
    $doc = $rt.Document
    $sw.Stop()
    $msg = $doc.Message
    Measure-Message $msg 0
    if ($doc.Envelope) {
        Add-Seen Other 'envelope (Graph records)'
        foreach ($e in $doc.Envelope.Entries | Where-Object { -not $_.IsData }) { Add-Seen Other ('envelope record type ' + $e.Type + ' (' + $e.Length + ' bytes)') }
    }
    else { Add-Seen Other 'bare stream (no envelope)' }
    $ok = $rt.Identical -and $doc.Warnings.Count -eq 0
    [pscustomobject]@{
        Item        = $f.Name
        Bytes       = $f.Length
        Properties  = $msg.CountProperties()
        Recipients  = $msg.Recipients.Count
        Attachments = $msg.Attachments.Count
        Embedded    = $msg.EmbeddedDepth()
        Warnings    = $doc.Warnings.Count
        RoundTrip   = if ($ok) { 'OK' } else { 'KO' }
        FirstDiff   = if ($rt.FirstDifference -ge 0) { $rt.FirstDifference } else { $null }
        Seconds     = [math]::Round($sw.Elapsed.TotalSeconds, 1)
        Note        = $rt.Note
        Warning     = @($doc.Warnings | Select-Object -First 3)
    }
}

$rows | Format-Table Item, Bytes, Properties, Recipients, Attachments, Embedded, Warnings, RoundTrip, FirstDiff, Seconds, Note -AutoSize | Out-String -Width 220 | Write-Host
foreach ($r in $rows | Where-Object { $_.Warning.Count -gt 0 }) { foreach ($w in $r.Warning) { Write-Host "  $($r.Item): $w" -ForegroundColor Yellow } }

$bad = @($rows | Where-Object RoundTrip -NE 'OK')
Write-Host ("{0} item(s): {1} round-trip OK, {2} KO, {3} decoder warning(s)." -f $rows.Count, ($rows.Count - $bad.Count), $bad.Count, ($rows | Measure-Object Warnings -Sum).Sum)

if (-not $NoCoverage) {
    Write-Host ''
    Write-Host 'Seen in this corpus (what is not listed is not covered by it):'
    foreach ($kind in 'Types', 'Named', 'Markers', 'Other') {
        foreach ($k in $seen[$kind].Keys) { Write-Host ('  {0,-8} {1}  x{2}' -f $kind, $k, $seen[$kind][$k]) }
    }
}
if ($bad.Count -gt 0) { exit 1 }
exit 0
