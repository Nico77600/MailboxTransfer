#Requires -Version 7.4
<#
.SYNOPSIS
    Live check of the EML import (lab tenant, adele.vance@contoso.com's PRIMARY mailbox only): imports .eml files into a test folder, reads them back
    (Mail API $value and message properties) and compares them with the source .eml files.

.DESCRIPTION
    Steps (-Step):
      Import   Import-MbtEml of -Path into the top-level folder -TargetRoot (default MBT-EmlImport-Test-<yyyyMMdd>), -MimeConversion Own|Exchange.
               Create only; the folder is created when missing. Never writes to the archive or to an auxiliary archive.
      Verify   For every file imported (import-state.jsonl): GET message properties (receivedDateTime, sentDateTime, isRead, isDraft,
               hasAttachments, subject) and GET $value, then EmlCompare of the source .eml (left) and Exchange's MIME of the imported item (right).
    Uses the cached token only (Connect-MbtGraph -NoPrompt, or -DeviceCode: the code is written to the console and the script waits).
    Nothing is deleted.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('Import', 'Verify', 'Compare')][string]$Step,
    [Parameter(Mandatory)][string]$Path,
    [string]$TargetRoot = ('MBT-EmlImport-Test-' + (Get-Date -Format 'yyyyMMdd')),
    [ValidateSet('Own', 'Exchange')][string]$MimeConversion = 'Own',
    [string]$StatePath,
    [string[]]$MorePath,
    [string]$PathExchange,
    [switch]$Unread,
    [string]$KeepDir,
    [switch]$DeviceCode,
    [string]$ConfigPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $repo 'package\MailboxTransfer.psd1') -Force
$config = if ($ConfigPath) { Import-MbtConfiguration -Path $ConfigPath } else { Import-MbtConfiguration }
Initialize-MbtEngine
try {
    $session = Connect-MbtGraph -TenantId $config.Lab.TenantId -ClientId $config.Lab.ClientId -UserPrincipalName $config.Lab.UserPrincipalName -NoPrompt:(-not $DeviceCode) -DeviceCode:$DeviceCode
} catch { Write-Host "NO VALID CACHED TOKEN ($($_.Exception.Message)): stopping, a sign-in is needed (nothing was asked)." -ForegroundColor Red; exit 9 }
if ($session.Token.Account -ine $config.Lab.UserPrincipalName) { throw "Signed in as $($session.Token.Account), not the lab user." }
$mbx = $config.Lab.PrimaryMailboxId   # the primary mailbox ONLY
Write-Host ("signed in as {0}; permissions: {1}" -f $session.Token.Account, ($session.Token.Scopes -join ' '))

switch ($Step) {
    'Import' {
        $args2 = @{ Path = $Path; MailboxId = $mbx; TargetRoot = $TargetRoot; MimeConversion = $MimeConversion; Unread = $Unread; Confirm = $false }
        if ($StatePath) { $args2.StatePath = $StatePath }
        $r = Import-MbtEml @args2
        Write-Host ("imported {0}, resumed {1}, failed {2}, unknown {3}, folders created {4}" -f $r.Done, $r.Resumed, $r.Failed, $r.Unknown, $r.FoldersCreated)
        $r.Entries | Where-Object State -ne 'Done' | ForEach-Object { Write-Host ("  {0} {1}: {2}" -f $_.State, $_.File, $_.Error) -ForegroundColor Yellow }
    }
    'Compare' {
        # own conversion (-Path, -MorePath) against Exchange's conversion (-PathExchange): both read back from the mailbox, files matched by name
        function Read-Imported([string]$Dir) {
            $state = Join-Path (Get-Item -LiteralPath $Dir).FullName 'import-state.jsonl'
            foreach ($e in @([MailboxTransfer.EmlImporter]::ReadState($state) | Where-Object { $_.State -eq 'Done' -and $_.TargetRoot -ieq "$mbx|$TargetRoot" })) {
                [pscustomobject]@{ Name = [IO.Path]::GetFileName($e.File); Dir = (Get-Item -LiteralPath $Dir).FullName; Entry = $e }
            }
        }
        function Read-Item($Imp) {
            $g = $session.Graph
            $id = [uri]::EscapeDataString($Imp.Entry.ItemId)
            $doc = $g.GetJson($g.Url("/users/$mbx/messages/$id`?`$select=receivedDateTime,isRead,isDraft,hasAttachments")); $p = $doc.RootElement.Clone(); $doc.Dispose()
            $resp = $g.Send('GET', $g.Url("/users/$mbx/messages/$id/`$value"))
            try { $ms = [IO.MemoryStream]::new(); $resp.Content.ReadAsStream().CopyTo($ms) } finally { $resp.Dispose() }
            [pscustomobject]@{ Props = $p; Mime = $ms.ToArray(); Source = (Join-Path $Imp.Dir ($Imp.Entry.File -replace '/', '\')) }
        }
        $own = @(@($Path) + @($MorePath) | Where-Object { $_ } | ForEach-Object { Read-Imported $_ })
        $exc = @(Read-Imported $PathExchange)
        Write-Host ("own conversion: {0} item(s); Exchange conversion: {1} item(s)" -f $own.Count, $exc.Count)
        $rows = foreach ($x in $exc) {
            $o = $own | Where-Object Name -eq $x.Name | Select-Object -First 1
            $a = Read-Item $x
            $srcBytes = [IO.File]::ReadAllBytes($a.Source)
            $srcMime = [MailboxTransfer.MimeParser]::Parse($srcBytes)
            $expected = [MailboxTransfer.MimeToMapi]::ReceivedFromHeaders($srcMime)
            if (-not $expected) { $expected = [MailboxTransfer.MimeText]::ParseDate($srcMime.Header('Date')) }
            $shape = {
                param($item)
                $c = [MailboxTransfer.EmlCompare]::Compare($srcBytes, $item.Mime)
                $recv = [DateTimeOffset]::Parse($item.Props.GetProperty('receivedDateTime').GetString())
                [pscustomobject]@{
                    Content = $(if ($c.ContentEqual) { 'equal' } else { 'DIFFERS' }); Structure = $(if ($c.LeftStructure -eq $c.RightStructure) { 'same' } else { 'differs' })
                    HdrLost = @($c.Differences | Where-Object { $_.Kind -eq 'Header' -and $_.Result -eq 'OnlyLeft' } | ForEach-Object Name).Count
                    HdrLostNames = (@($c.Differences | Where-Object { $_.Kind -eq 'Header' -and $_.Result -eq 'OnlyLeft' } | ForEach-Object Name) -join ',')
                    ReceivedOk = [bool]($expected -and [Math]::Abs(($recv - $expected).TotalSeconds) -lt 2)
                    Read = $item.Props.GetProperty('isRead').GetBoolean(); Draft = $item.Props.GetProperty('isDraft').GetBoolean(); Cmp = $c
                }
            }
            $sx = & $shape $a
            $so = if ($o) { & $shape (Read-Item $o) } else { $null }
            $between = if ($o) { $ob = Read-Item $o; [MailboxTransfer.EmlCompare]::Compare($ob.Mime, $a.Mime) } else { $null }
            [pscustomobject]@{
                File = $x.Name
                Own = if ($so) { "$($so.Content)/$($so.Structure)/recv:$($so.ReceivedOk)/read:$($so.Read)/draft:$($so.Draft)/hdrlost:$($so.HdrLost)" } else { 'n/a' }
                Exchange = "$($sx.Content)/$($sx.Structure)/recv:$($sx.ReceivedOk)/read:$($sx.Read)/draft:$($sx.Draft)/hdrlost:$($sx.HdrLost)"
                OwnVsExchange = if ($between) { $(if ($between.ContentEqual) { 'content equal' } else { 'content DIFFERS' }) + ", $(@($between.Differences).Count) difference(s)" } else { 'n/a' }
                LostOwn = if ($so) { $so.HdrLostNames } else { '' }; LostExchange = $sx.HdrLostNames
            }
        }
        $rows | Format-Table File, Own, Exchange, OwnVsExchange -AutoSize -Wrap | Out-String -Width 300 | Write-Host
        $rows | Where-Object { $_.LostOwn -or $_.LostExchange } | ForEach-Object { Write-Host ("headers not returned by `$value - {0}: own [{1}] exchange [{2}]" -f $_.File, $_.LostOwn, $_.LostExchange) }
    }
    'Verify' {
        $state = if ($StatePath) { $StatePath } else { Join-Path (Get-Item -LiteralPath $Path).FullName 'import-state.jsonl' }
        $entries = @([MailboxTransfer.EmlImporter]::ReadState($state) | Where-Object { $_.State -eq 'Done' -and $_.TargetRoot -ieq "$mbx|$TargetRoot" })
        Write-Host ("{0} imported item(s) in the state file" -f $entries.Count)
        $g = $session.Graph
        $root = (Get-Item -LiteralPath $Path).FullName
        $rowsOut = foreach ($e in $entries) {
            $id = [uri]::EscapeDataString($e.ItemId)
            $props = $null
            $url = $g.Url("/users/$mbx/messages/$id`?`$select=subject,receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,internetMessageId,importance")
            $doc = $g.GetJson($url); $props = $doc.RootElement.Clone(); $doc.Dispose()
            $resp = $g.Send('GET', $g.Url("/users/$mbx/messages/$id/`$value"))
            try { $ms = [IO.MemoryStream]::new(); $resp.Content.ReadAsStream().CopyTo($ms); $got = $ms.ToArray() } finally { $resp.Dispose() }
            $srcFile = Join-Path $root ($e.File -replace '/', '\')
            if ($KeepDir) {
                New-Item -ItemType Directory -Force $KeepDir | Out-Null
                [IO.File]::WriteAllBytes((Join-Path $KeepDir ([IO.Path]::GetFileName($e.File) -replace '[^A-Za-z0-9.-]', '_')), $got)
            }
            $cmp = [MailboxTransfer.EmlCompare]::Compare([IO.File]::ReadAllBytes($srcFile), $got)
            $srcMime = [MailboxTransfer.MimeParser]::ParseFile($srcFile)
            $expected = [MailboxTransfer.MimeToMapi]::ReceivedFromHeaders($srcMime)
            if (-not $expected) { $expected = [MailboxTransfer.MimeText]::ParseDate($srcMime.Header('Date')) }
            $recv = [DateTimeOffset]::Parse($props.GetProperty('receivedDateTime').GetString())
            [pscustomobject]@{
                File = $e.File; Subject = $e.Subject; Received = $recv.UtcDateTime.ToString('u'); ExpectedReceived = $(if ($expected) { $expected.UtcDateTime.ToString('u') } else { '(none: import time)' })
                ReceivedOk = [bool]($expected -and [Math]::Abs(($recv - $expected).TotalSeconds) -lt 2)
                IsRead = $props.GetProperty('isRead').GetBoolean(); IsDraft = $props.GetProperty('isDraft').GetBoolean()
                Content = $(if ($cmp.ContentEqual) { 'equal' } else { 'DIFFERS' }); Structure = $(if ($cmp.LeftStructure -eq $cmp.RightStructure) { 'same' } else { "$($cmp.LeftStructure) -> $($cmp.RightStructure)" })
                HeaderDiffs = (@($cmp.Differences | Where-Object { $_.Kind -eq 'Header' -and $_.Result -ne 'OnlyRight' } | ForEach-Object { "$($_.Name):$($_.Result)" }) -join ' ')
                Cmp = $cmp
            }
        }
        $rowsOut | Select-Object File, Received, ExpectedReceived, ReceivedOk, IsRead, IsDraft, Content, Structure | Format-Table -AutoSize -Wrap | Out-String -Width 250 | Write-Host
        foreach ($r in $rowsOut) {
            $d = @($r.Cmp.Differences | Where-Object { $_.Kind -ne 'Header' -or $_.Result -ne 'OnlyRight' })
            if ($d) { Write-Host ("-- {0}" -f $r.File); $d | ForEach-Object { Write-Host ("     {0}" -f $_) } }
        }
    }
}
Disconnect-MbtGraph
