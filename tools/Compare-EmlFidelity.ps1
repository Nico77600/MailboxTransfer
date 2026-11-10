#Requires -Version 7.4
<#
.SYNOPSIS
    Fidelity of the MAPI -> MIME writer against Exchange's own MIME (read only).

.DESCRIPTION
    For the messages of one folder of a live mailbox (Primary or Archive, Mail API):
      1. GET .../messages/{id}/$value                       Exchange's MIME (the reference)
      2. exportItems -> FastTransfer stream -> FxCodec.Decode -> MapiToMime   this tool's MIME from the MAPI properties
    and compares the two (EmlCompare: headers, structure, text bodies, attachment hashes). Prints one line per message and a summary
    table by header and by kind of difference. Nothing is created or changed in the mailbox; both MIME files can be kept with -KeepDir.

    The same comparison works for a store or a PST item when its live twin exists (use Export-MbtEml for both and Compare-MbtEml).
    Uses the cached token (Connect-MbtGraph -NoPrompt): it stops when none is valid.

.EXAMPLE
    .\Compare-EmlFidelity.ps1 -Folder Inbox/EML-Import-Test
    .\Compare-EmlFidelity.ps1 -Mailbox Archive -Folder 'Deleted Items' -KeepDir D:\cmp
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Folder,
    [ValidateSet('Primary', 'Archive')][string]$Mailbox = 'Primary',
    [int]$Top = 20,
    [string]$KeepDir,
    [string]$ConfigPath
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
Import-Module (Join-Path $repo 'package\MailboxTransfer.psd1') -Force
$config = if ($ConfigPath) { Import-MbtConfiguration -Path $ConfigPath } else { Import-MbtConfiguration }
Initialize-MbtEngine
try {
    $session = Connect-MbtGraph -TenantId $config.Lab.TenantId -ClientId $config.Lab.ClientId -UserPrincipalName $config.Lab.UserPrincipalName -NoPrompt
} catch { Write-Host "NO VALID CACHED TOKEN ($($_.Exception.Message)): stopping." -ForegroundColor Red; exit 9 }
$mbx = if ($Mailbox -eq 'Primary') { $config.Lab.PrimaryMailboxId } else { $config.Lab.ArchiveMailboxId }
$loc = New-MbtBrowseLocation -Kind $Mailbox -MailboxId $mbx
$browser = Get-MbtBrowser -NoNonMailFolders
$f = Find-MbtBrowseFolder -Location $loc -Path $Folder
$rows = @($browser.EnumerateItems($loc, $f, [MailboxTransfer.BrowseItemsQuery]@{ PageSize = 100; Recipients = $false }, [Threading.CancellationToken]::None, $null) | Select-Object -First $Top)
Write-Host ("{0}: {1} message(s) compared ({2})" -f $f.Path, $rows.Count, $Mailbox)
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('mbt-cmp-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
$all = [System.Collections.Generic.List[object]]::new()
try {
    $ids = @($rows | ForEach-Object Id)
    $exports = @{}
    for ($i = 0; $i -lt $ids.Count; $i += 20) {
        $chunk = [string[]]@($ids[$i..([Math]::Min($i + 19, $ids.Count - 1))])
        foreach ($e in $session.Api.ExportItems($mbx, $chunk, $tmp, $i)) { $exports[$e.ItemId] = $e }
    }
    $n = 0
    foreach ($row in $rows) {
        $n++
        $ref = $null
        $url = $session.Graph.Url("/users/$mbx/messages/$([uri]::EscapeDataString($row.Id))/`$value")
        $resp = $session.Graph.Send('GET', $url)
        try { $ms = [IO.MemoryStream]::new(); $resp.Content.ReadAsStream().CopyTo($ms); $ref = $ms.ToArray() } finally { $resp.Dispose() }
        $e = $exports[$row.Id]
        if (-not $e -or $e.HasError) { Write-Host ("  {0,2} {1}: export failed {2}" -f $n, $row.Subject, $e.ErrorMessage) -ForegroundColor Yellow; continue }
        $doc = [MailboxTransfer.FxCodec]::DecodeFile($e.FilePath)
        $notes = [System.Collections.Generic.List[string]]::new()
        $own = [MailboxTransfer.MapiToMime]::ToBytes($doc.Message, $null, $notes)
        if ($KeepDir) {
            New-Item -ItemType Directory -Force -Path $KeepDir | Out-Null
            [IO.File]::WriteAllBytes((Join-Path $KeepDir ("{0:00}-exchange.eml" -f $n)), $ref)
            [IO.File]::WriteAllBytes((Join-Path $KeepDir ("{0:00}-own.eml" -f $n)), $own)
        }
        $cmp = [MailboxTransfer.EmlCompare]::Compare($own, $ref)   # left = ours, right = Exchange
        $all.Add([pscustomobject]@{ N = $n; Subject = $row.Subject; Cmp = $cmp; Notes = $notes })
        $diff = @($cmp.Differences)
        $kinds = ($diff | Group-Object Kind | ForEach-Object { "$($_.Name)=$($_.Count)" }) -join ' '
        Write-Host ("  {0,2} {1,-48} content {2}  structure {3}  diffs: {4}" -f $n, ($row.Subject ?? '').PadRight(48).Substring(0, 48), $(if ($cmp.ContentEqual) { 'equal   ' } else { 'DIFFERS ' }), $(if ($cmp.LeftStructure -eq $cmp.RightStructure) { 'same' } else { 'differ' }), $(if ($kinds) { $kinds } else { 'none' }))
    }
} finally { Remove-Item -Recurse -Force $tmp -ErrorAction SilentlyContinue }

Write-Host "`nBy header (L = this tool, R = Exchange): number of messages"
$all | ForEach-Object { $_.Cmp.Items | Where-Object { $_.Kind -eq 'Header' } } | Group-Object Name, Result | Sort-Object Name |
    ForEach-Object { '{0,-34} {1}' -f $_.Name, $_.Count } | Write-Host
Write-Host "`nBodies and attachments"
$all | ForEach-Object { $_.Cmp.Items | Where-Object { $_.Kind -in 'Body', 'Attachment', 'Structure' } } | Group-Object Kind, Result | ForEach-Object { '{0,-34} {1}' -f $_.Name, $_.Count } | Write-Host
$notesAll = @($all | ForEach-Object { $_.Notes } | Group-Object | ForEach-Object { '{0}  x{1}' -f $_.Name, $_.Count })
if ($notesAll) { Write-Host "`nNotes of the writer:"; $notesAll | Write-Host }
Disconnect-MbtGraph
