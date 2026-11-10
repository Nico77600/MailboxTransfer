#Requires -Version 7.4
<#
.SYNOPSIS
    Copies the FastTransfer streams of an export store into a folder of .fxs files (one file per item) for
    tools\Test-FastTransferCorpus.ps1, with an index of what each file is. Reads the store, never changes it.

.DESCRIPTION
    The store (catalog.db + packs\) is opened READ-ONLY with the module's store API, so it can be read while another process exports
    into it. Each Exported item of the selected folders is copied from its pack to <OutputPath>\item-<key>.fxs and its SHA-256 is
    compared with the one the store recorded when the item arrived (a file that does not match is reported and left out of the
    result). An existing file of the right size and hash is kept as is; nothing in the output folder is ever deleted.

    index.json (in the output folder) lists, per file: the item key, the folder path, the item type, the subject, the size and the
    hash. The subjects of the lab corpus start with [fxc-NN] (tools\New-LabFxCorpus.ps1), which tells what each stream is.

    Exit code: 0 if every selected item was copied (or already there) and checked, 1 otherwise (or nothing selected).

.PARAMETER StorePath
    Folder of the store (default <repository>\out\fxics-corpus-store).

.PARAMETER OutputPath
    Folder that receives the .fxs files and index.json (default <repository>\out\fxics-corpus2; out\ is git-ignored).

.PARAMETER Folder
    Only the items whose folder path is this one or below it (for instance MBT-FXICS-Corpus/Mail). Default: every folder of the store.
    Several folders: give the parameter an array.

.PARAMETER Mailbox
    Id of the mailbox of the store (default: every mailbox).

.EXAMPLE
    .\tools\Export-FxCorpusFromStore.ps1 -WhatIf

.EXAMPLE
    .\tools\Export-FxCorpusFromStore.ps1 -Folder 'MBT-FXICS-Corpus/Mail', 'Calendar/MBT-FXICS-Corpus', 'Contacts/MBT-FXICS-Corpus'
    .\tools\Test-FastTransferCorpus.ps1 -Path .\out\fxics-corpus2

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$StorePath = (Join-Path $PSScriptRoot '..\out\fxics-corpus-store'),
    [string]$OutputPath = (Join-Path $PSScriptRoot '..\out\fxics-corpus2'),
    [string[]]$Folder,
    [string]$Mailbox
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# -WhatIf only means "do not write the files": the module functions that read the store must not see it.
$dryRun = [bool]$WhatIfPreference
$WhatIfPreference = $false

$package = Join-Path $PSScriptRoot '..\package'
Import-Module (Join-Path $package 'MailboxTransfer.psd1') -Force
Initialize-MbtEngine -Root (Resolve-Path $package).Path

$StorePath = [IO.Path]::GetFullPath($StorePath)
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if (-not (Test-Path -LiteralPath (Join-Path $StorePath 'catalog.db'))) { Write-Warning "No store in $StorePath"; exit 1 }

function Test-InFolder([string]$Path) {
    if (-not $Folder) { return $true }
    foreach ($f in $Folder) {
        $p = $f.Trim('/')
        if ($Path -ieq $p -or $Path.StartsWith($p + '/', [StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    $false
}

$store = Open-MbtExportStore -Path $StorePath -ReadOnly
$index = [Collections.Generic.List[object]]::new()
$problems = 0
try {
    $items = @(Get-MbtStoreItems -Store $store -MailboxId $Mailbox | Where-Object { Test-InFolder $_.FolderPath } | Sort-Object Key)
    if (-not $items) { Write-Warning "No exported item in $StorePath$(if ($Folder) { " for the folder(s) $($Folder -join ', ')" })."; exit 1 }
    Write-Host ("{0} item(s) to copy from {1} to {2}" -f $items.Count, $StorePath, $OutputPath)
    if (-not $dryRun -and -not (Test-Path -LiteralPath $OutputPath)) { $null = New-Item -ItemType Directory -Path $OutputPath }

    foreach ($item in $items) {
        $name = 'item-{0:D6}.fxs' -f $item.Key
        $target = Join-Path $OutputPath $name
        $state = 'copied'
        if ($dryRun) { $state = 'planned' }
        else {
            if ((Test-Path -LiteralPath $target) -and (Get-Item -LiteralPath $target).Length -eq $item.Length -and
                (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ieq $item.Sha256) { $state = 'kept' }
            else {
                $in = $store.OpenItem($item.Key)
                try { $out = [IO.File]::Create($target); try { $in.CopyTo($out) } finally { $out.Dispose() } } finally { $in.Dispose() }
                if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ine $item.Sha256) {
                    Write-Warning "Item $($item.Key): the copied file does not match the hash of the store; removed from the result."
                    Remove-Item -LiteralPath $target -Force
                    $problems++; continue
                }
            }
        }
        Write-Host ("  {0,-8} {1}  {2}  {3}" -f $state, $name, $item.ItemType, $item.Subject)
        $index.Add([ordered]@{ File = $name; Key = $item.Key; ItemId = $item.ItemId; FolderPath = $item.FolderPath; ItemType = $item.ItemType; Subject = $item.Subject; Size = $item.Length; Sha256 = $item.Sha256 })
    }
} finally { $store.Dispose() }

if (-not $dryRun) {
    $index | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $OutputPath 'index.json') -Encoding utf8
    Write-Host ("{0} file(s) in {1}, index: {2}" -f $index.Count, $OutputPath, (Join-Path $OutputPath 'index.json'))
}
if ($problems) { exit 1 }
