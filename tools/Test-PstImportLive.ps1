#Requires -Version 7.4
<#
.SYNOPSIS
    Lab tool: does importItem accept a FastTransfer stream WE built from a PST message, and what does Exchange keep of it?

.DESCRIPTION
    Imports the messages of a PST (read with our PST reader, converted to FastTransfer by our encoder) into a dedicated folder of the
    PRIMARY mailbox of the lab user (never the archives), exports the imported items again and lists, property by property, what
    Exchange changed, dropped or added (FxDiff). Create only: nothing existing is changed or deleted.

    Steps (each one is a separate run, state in <OutDir>\state.json):
      Probe     sign in (cached token only: never prompts), read the PST, create the target folders, list what is there.
      Baseline  import ONE exported stream (an .fxs of the corpus, as Exchange produced it) unchanged, export it again.
      Variants  import ONE message of the PST built with each envelope variant (bare, header, header + key, ...).
      Import    import -Take messages (0 = all) of the PST with -Envelope; items already imported are skipped (state).
      Verify    export the imported items again and write the differences (CSV + summary).

    The signed-in user must be the lab user; the sign-in is silent only unless -DeviceCode is given (without it, when no valid token is cached the
    tool stops and says so).
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('Probe', 'Baseline', 'Variants', 'Import', 'Verify')][string]$Step,
    [string]$PstPath = 'D:\MailboxTransfer\out\pst-validation\corpus2.pst',
    [string]$CorpusDir = 'D:\MailboxTransfer\out\fxics-corpus2',
    [string]$TargetName = 'MBT-PstImport-Test-20261109',
    [string]$Generation = '',
    [switch]$NoDelProp,
    [string[]]$Experiment = @(),
    [string]$Only,
    [string]$OutDir,
    [ValidateSet('Bare', 'Header', 'HeaderKey', 'HeaderZeroKey')][string]$Envelope = 'Header',
    [int]$Take = 0,
    [int]$Skip = 0,
    [switch]$DeviceCode,
    [string]$ConfigPath
)
Set-StrictMode -Version Latest
$Experiment = @($Experiment | ForEach-Object { $_ -split ',' } | Where-Object { $_ })
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
if (-not $OutDir) { $OutDir = Join-Path $repo 'out\pstimport-live' }
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null
Import-Module (Join-Path $repo 'package\MailboxTransfer.psd1') -Force
$config = if ($ConfigPath) { Import-MbtConfiguration -Path $ConfigPath } else { Import-MbtConfiguration }
Initialize-MbtEngine
$statePath = Join-Path $OutDir "state$Generation.json"
$state = if (Test-Path $statePath) { Get-Content $statePath -Raw | ConvertFrom-Json -AsHashtable } else { @{ Folders = @{}; Imported = @{} } }
function Save-State { $state | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $statePath -Encoding utf8 }

# ---------------------------------------------------------------- sign in (silent only)
try {
    $session = Connect-MbtGraph -TenantId $config.Lab.TenantId -ClientId $config.Lab.ClientId -UserPrincipalName $config.Lab.UserPrincipalName -NoPrompt:(-not $DeviceCode) -DeviceCode:$DeviceCode -DeviceCodeFile (Join-Path $OutDir 'device-code.txt')
} catch {
    Write-Host "NO VALID CACHED TOKEN ($($_.Exception.Message)): stopping, a sign-in is needed (nothing was asked)." -ForegroundColor Red
    exit 9
}
$api = $session.Api
$mbx = $config.Lab.PrimaryMailboxId
Write-Host ("signed in as {0}; scopes: {1}" -f $session.Token.Account, ($session.Token.Scopes -join ' '))
if ($session.Token.Account -ine $config.Lab.UserPrincipalName) { throw "Signed in as $($session.Token.Account), not the lab user." }

# ---------------------------------------------------------------- target folders (primary only)
function Ensure-Folder([string]$Name, [string]$ParentId, [string]$Type) {
    $existing = if ($ParentId) { $api.GetChildFolders($mbx, $ParentId) } else { $api.GetRootFolders($mbx) }
    $f = @($existing | Where-Object { $_.DisplayName -eq $Name }) | Select-Object -First 1
    if ($f) { return $f }
    Write-Host "creating folder $Name ($Type)"
    return $api.CreateFolder($mbx, $Name, $(if ($ParentId) { $ParentId } else { [NullString]::Value }), $Type)
}
$rootF = Ensure-Folder $TargetName $null 'IPF.Note'
$folderDefs = [ordered]@{ Mail = 'IPF.Note'; Calendar = 'IPF.Appointment'; Contacts = 'IPF.Contact'; Variants = 'IPF.Note' }
$folderIds = @{}
foreach ($k in $folderDefs.Keys) { $folderIds[$k] = (Ensure-Folder ("$Generation$k") $rootF.Id $folderDefs[$k]).Id }
$state.Folders = $folderIds
$state.RootId = $rootF.Id
Save-State

# ---------------------------------------------------------------- PST
$reader = [MailboxTransfer.PstReader]::Open($PstPath)
$items = [System.Collections.Generic.List[object]]::new()
foreach ($f in $reader.AllFolders()) {
    $cls = $f.ContainerClass
    $target = switch -Wildcard ($cls) { 'IPF.Appointment' { 'Calendar' } 'IPF.Contact' { 'Contacts' } 'IPF.Note' { if ($f.Path -like '*/Mail') { 'Mail' } else { $null } } default { $null } }
    if (-not $target) { continue }
    foreach ($nid in $reader.MessageNids($f, $false)) {
        $m = $reader.ReadMessage($nid)
        $items.Add([pscustomobject]@{ Nid = $nid; Key = ('{0:X}' -f $nid); Target = $target; Class = $m.GetString(0x001A); Subject = $m.GetString(0x0037); Length = [MailboxTransfer.FxCodec]::PayloadLength($m) })
    }
}
Write-Host ("PST: {0} importable messages" -f $items.Count)

# ---------------------------------------------------------------- envelope variants
function New-Envelope([string]$Kind) {
    if ($Kind -eq 'Bare') { return $null }
    $env = [MailboxTransfer.FxEnvelope]::new()
    $e1 = [MailboxTransfer.FxEnvelopeEntry]::new(); $e1.Type = 1; $e1.Payload = [byte[]](0, 0, 0, 0, 1, 0, 0, 0); $e1.Length = 8
    $env.Entries.Add($e1)
    if ($Kind -ne 'Header') {
        $e4 = [MailboxTransfer.FxEnvelopeEntry]::new(); $e4.Type = 4; $e4.Length = 6
        $e4.Payload = if ($Kind -eq 'HeaderZeroKey') { [byte[]](0, 0, 0, 0, 0, 0) } else { [byte[]](0x15, 0x0F, 0x39, 0x82, 0x72, 0x17) }
        $env.Entries.Add($e4)
    }
    return $env
}

function Apply-Experiment($m, [string[]]$Names) {
    $all = {
        param($x, $f)
        & $f $x.Properties
        foreach ($r in $x.Recipients) { & $f $r.Properties }
        foreach ($a in $x.Attachments) { & $f $a.Properties; if ($a.Embedded) { & $all $a.Embedded $f } }
    }
    foreach ($n in $Names) {
        switch ($n) {
            'NoAtt' { $m.Attachments.Clear() }
            'NoRcp' { $m.Recipients.Clear() }
            'Types' { & $all $m { param($ps) foreach ($p in $ps) { if ($p.Type -eq 0x001F) { $p.Type = 0x84B0 } elseif ($p.Type -eq 0x101F) { $p.Type = 0x94B0 } } } }
            'Keys2' {
                foreach ($r in $m.Recipients) {
                    $one = $r.Properties | Where-Object { $_.Id -eq 0x5FF7 } | Select-Object -First 1
                    if ($one) {
                        $r.Properties.Add([MailboxTransfer.FxProperty]::FromBinary(0x0FF9, $one.Raw))
                        $r.Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x0FFE, 6))
                        $r.Properties.Add([MailboxTransfer.FxProperty]::FromBinary(0x0FFF, $one.Raw))
                    }
                }
            }
            'Order' {
                foreach ($r in $m.Recipients) {
                    $ps = @($r.Properties | Sort-Object { if ($_.Id -eq 0x3000) { -1 } else { $_.Id } })
                    $r.Properties.Clear(); foreach ($p in $ps) { $r.Properties.Add($p) }
                }
            }
            'EmbStrip' { foreach ($a in $m.Attachments) { if ($a.Embedded) { $a.Embedded.Properties.RemoveAll([Predicate[MailboxTransfer.FxProperty]]{ param($p) $p.Id -in 0x0E02,0x0E03,0x0E04,0x0E08,0x0E17 }) | Out-Null } } }
            'EmbMid' { foreach ($a in $m.Attachments) { if ($a.Embedded) { $a.Embedded.Properties.Insert(0, [MailboxTransfer.FxProperty]::FromInt64(0x674A, 1234567)) } } }
            'EmbNoRcp' { foreach ($a in $m.Attachments) { if ($a.Embedded) { $a.Embedded.Recipients.Clear() } } }
            'EmbNoAtt' { foreach ($a in $m.Attachments) { if ($a.Embedded) { $a.Embedded.Attachments.Clear() } } }
            'EmbNoProps' { foreach ($a in $m.Attachments) { if ($a.Embedded) { $a.Embedded.Recipients.Clear(); $a.Embedded.Attachments.Clear() } } }
            'Keys' {
                foreach ($r in $m.Recipients) {
                    $r.Properties.Add([MailboxTransfer.FxProperty]::FromBinary(0x0FF9, [Guid]::NewGuid().ToByteArray()))
                    $r.Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x0FFE, 6))
                    $r.Properties.Add([MailboxTransfer.FxProperty]::FromBinary(0x0FFF, [byte[]](0, 0, 0, 0) + [Guid]::NewGuid().ToByteArray()))
                }
                foreach ($a in $m.Attachments) { $a.Properties.Add([MailboxTransfer.FxProperty]::FromBinary(0x0FF9, [Guid]::NewGuid().ToByteArray())) }
            }
        }
    }
}
function Import-One($pstItem, [string]$FolderId, [string]$EnvKind, $Pool) {
    $msg = $reader.ReadMessage($pstItem.Nid); if (-not $NoDelProp) { [void][MailboxTransfer.FxImportStream]::Prepare($msg) }
    Apply-Experiment $msg $Experiment
    $src = [MailboxTransfer.FxMessageSource]::new($msg, (New-Envelope $EnvKind))
    $sw = [Diagnostics.Stopwatch]::StartNew()
    try {
        $r = $api.ImportItem($Pool, $FolderId, $src)
        return [pscustomobject]@{ Ok = $true; ItemId = $r.ItemId; Bytes = $src.Length; Seconds = $sw.Elapsed.TotalSeconds; Error = $null }
    } catch {
        $msg = $_.Exception.Message
        $ge = Get-MbtGraphException $_
        if ($ge) { $msg = "HTTP $($ge.StatusCode) $($ge.ErrorCode): $($ge.Message)" }
        return [pscustomobject]@{ Ok = $false; ItemId = $null; Bytes = $src.Length; Seconds = $sw.Elapsed.TotalSeconds; Error = $msg }
    }
}

function Export-Back([string[]]$Ids, [string]$Dir) {
    New-Item -ItemType Directory -Force -Path $Dir | Out-Null
    $res = @()
    for ($i = 0; $i -lt $Ids.Count; $i += 20) {
        $chunk = [string[]]@($Ids[$i..([Math]::Min($i + 19, $Ids.Count - 1))])
        $res += @($api.ExportItems($mbx, [string[]]$chunk, $Dir, $i))
    }
    return $res
}

switch ($Step) {
    'Probe' {
        foreach ($i in $items) { Write-Host ("  {0,-8} {1,-9} {2,-16} {3,8}  {4}" -f $i.Key, $i.Target, $i.Class, $i.Length, $i.Subject.Substring(0, [Math]::Min(60, $i.Subject.Length))) }
        foreach ($k in $folderIds.Keys) { $fi = $api.GetFolder($mbx, $folderIds[$k]); Write-Host ("folder {0}: {1} items" -f $k, $fi.TotalItemCount) }
    }
    'Baseline' {
        $file = Join-Path $CorpusDir 'item-000003.fxs'
        $pool = $api.CreateImportSession($mbx)
        Write-Host "importing $file unchanged (Exchange's own stream)"
        $sw = [Diagnostics.Stopwatch]::StartNew()
        if ($state.Imported.ContainsKey('baseline')) { Write-Host 'already imported (state)' }
        else { try { $r = $api.ImportItem($pool, $folderIds.Variants, $file); Write-Host ("OK itemId {0} in {1:n1}s" -f $r.ItemId.Substring(0, 24), $sw.Elapsed.TotalSeconds); $state.Imported['baseline'] = @{ ItemId = $r.ItemId; Source = $file } }
        catch { $ge = Get-MbtGraphException $_; Write-Host "FAILED: $(if ($ge) { "HTTP $($ge.StatusCode) $($ge.ErrorCode) $($ge.Message)" } else { $_.Exception.Message })" -ForegroundColor Red } }
        Save-State
        if ($state.Imported.ContainsKey('baseline')) {
            $back = Export-Back @($state.Imported['baseline'].ItemId) (Join-Path $OutDir 'export-baseline')
            $b = $back[0]; Write-Host "re-exported: $($b.FilePath) $($b.Bytes) bytes"
            $orig = [MailboxTransfer.FxCodec]::DecodeFile($file); $got = [MailboxTransfer.FxCodec]::DecodeFile($b.FilePath)
            Write-Host ("original envelope: {0}; re-export envelope: {1}" -f (($orig.Envelope.Entries | ForEach-Object { "$($_.Type):$($_.Length)" }) -join ' '), (($got.Envelope.Entries | ForEach-Object { "$($_.Type):$($_.Length)" }) -join ' '))
            [MailboxTransfer.FxDiff]::Compare($orig.Message, $got.Message) | ForEach-Object { Write-Host "  $_" }
        }
    }
    'Variants' {
        $pool = $api.CreateImportSession($mbx)
        $pick = $items | Where-Object { $_.Subject -like '*Empty body*' } | Select-Object -First 1
        foreach ($v in 'Bare', 'Header', 'HeaderZeroKey', 'HeaderKey') {
            $r = Import-One $pick $folderIds.Variants $v $pool
            Write-Host ("{0,-14} {1}  {2}" -f $v, $(if ($r.Ok) { 'OK ' + $r.ItemId.Substring(0, 20) } else { 'FAILED ' + $r.Error }), ('{0:n1}s' -f $r.Seconds))
            if ($r.Ok) { $state.Imported["variant-$v"] = @{ ItemId = $r.ItemId; Nid = $pick.Key; Variant = $v } }
        }
        Save-State
    }
    'Import' {
        $pool = $api.CreateImportSession($mbx)
        $todo = @($items | Where-Object { (-not $state.Imported.ContainsKey($_.Key) -or $Only) -and (-not $Only -or $_.Key -eq $Only) } | Select-Object -Skip $Skip)
        if ($Take -gt 0) { $todo = @($todo | Select-Object -First $Take) }
        foreach ($i in $todo) {
            $r = Import-One $i $folderIds[$i.Target] $Envelope $pool
            $name = $i.Subject.Substring(0, [Math]::Min(40, $i.Subject.Length))
            Write-Host ("{0,-8} {1,-9} {2}  {3}  {4}" -f $i.Key, $i.Target, $(if ($r.Ok) { 'OK    ' } else { 'FAILED' }), ('{0:n1}s' -f $r.Seconds), $(if ($r.Ok) { $name } else { "$name -> $($r.Error)" }))
            if ($r.Ok) { $state.Imported[$i.Key] = @{ ItemId = $r.ItemId; Nid = $i.Key; Target = $i.Target; Envelope = $Envelope; Bytes = $r.Bytes } ; Save-State }
        }
    }
    'Verify' {
        $keys = @($state.Imported.Keys | Where-Object { $_ -notlike 'variant-*' -and $_ -ne 'baseline' } | Sort-Object)
        $ids = @($keys | ForEach-Object { $state.Imported[$_].ItemId })
        $dir = Join-Path $OutDir 'export-verify'
        $back = Export-Back $ids $dir
        $script:countMismatch = 0
        $rows = [System.Collections.Generic.List[object]]::new()
        for ($n = 0; $n -lt $keys.Count; $n++) {
            $k = $keys[$n]; $e = $back | Where-Object { $_.ItemId -eq $ids[$n] } | Select-Object -First 1
            if (-not $e -or $e.HasError) { Write-Host "$k : export failed $($e.ErrorCode) $($e.ErrorMessage)"; continue }
            $sent = $reader.ReadMessage([Convert]::ToUInt32($k, 16))
            $got = [MailboxTransfer.FxCodec]::DecodeFile($e.FilePath)
            $subject = $sent.GetString(0x0037)
            $diffs = [MailboxTransfer.FxDiff]::Compare($sent, $got.Message)
            $ts = [MailboxTransfer.FxDiff]::TotalsText($sent); $tg = [MailboxTransfer.FxDiff]::TotalsText($got.Message)
            if ($ts -ne $tg) { $script:countMismatch++; Write-Host ("  COUNT MISMATCH {0}: sent [{1}] exported [{2}]" -f $k, $ts, $tg) -ForegroundColor Red }
            foreach ($d in $diffs) { $rows.Add([pscustomobject]@{ Item = $k; Subject = $subject.Substring(0, [Math]::Min(30, $subject.Length)); Kind = $d.Kind; Where = $d.Where; Property = $d.PropertyName; Sent = $d.Sent; Got = $d.Got; Detail = $d.Detail }) }
            Write-Host ("{0,-8} sent {1,7} B, re-exported {2,7} B, {3} differences" -f $k, $state.Imported[$k].Bytes, $e.Bytes, $diffs.Count)
        }
        if ($script:countMismatch -eq 0) { Write-Host "counts OK: recipients, attachments and embedded messages of every item are the same as sent" -ForegroundColor Green } else { Write-Host "$($script:countMismatch) ITEM(S) LOST RECIPIENTS OR ATTACHMENTS" -ForegroundColor Red }
        $rows | Export-Csv -LiteralPath (Join-Path $OutDir "verify-diffs$Generation.csv") -NoTypeInformation -Encoding utf8
        Write-Host "--- by property (how many items)"
        $rows | Group-Object Kind, Property | Sort-Object Name | ForEach-Object { '{0,-70} {1} (items {2})' -f $_.Name, $_.Count, (@($_.Group | Select-Object -ExpandProperty Item -Unique).Count) } | Write-Host
    }
}
$reader.Dispose()
