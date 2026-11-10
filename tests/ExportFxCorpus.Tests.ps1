#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of tools\Export-FxCorpusFromStore.ps1 (streams of a store copied to .fxs files, read only).
    The store is built here with the store API; the script runs in a child process like the owner runs it. Nothing reaches the network.

    Run:  Invoke-Pester -Path .\tests\ExportFxCorpus.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot (Join-Path $script:RepoRoot 'package') -WithTestSupport
    $script:Tool = Join-Path $script:RepoRoot 'tools\Export-FxCorpusFromStore.ps1'
    $script:Base = Join-Path ([IO.Path]::GetTempPath()) ('fxcexport-' + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Base
    $script:StorePath = Join-Path $script:Base 'store'

    $s = [MailboxTransfer.ExportStore]::Create($script:StorePath, 'fxc', 'tests')
    $s.UpsertMailbox([MailboxTransfer.StoreMailbox]@{ Id = 'mb1'; Kind = 'Primary'; Upn = 'lab@contoso.test'; DisplayName = 'Lab' }) | Out-Null
    $s.UpsertFolders(@(
            [MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'corpus'; DisplayName = 'MBT-FXICS-Corpus'; FolderType = 'IPF.Note' },
            [MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'mail'; ParentFolderId = 'corpus'; DisplayName = 'Mail'; FolderType = 'IPF.Note' },
            [MailboxTransfer.StoreFolder]@{ MailboxId = 'mb1'; FolderId = 'other'; DisplayName = 'Other'; FolderType = 'IPF.Note' }
        )) | Out-Null
    $entries = [Collections.Generic.List[MailboxTransfer.StoreItemEntry]]::new()
    foreach ($i in 1..3) { $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = "m$i"; FolderId = 'mail'; ChangeKey = "c$i"; ItemType = 'IPM.Note'; Size = 100 * $i }) }
    $entries.Add([MailboxTransfer.StoreItemEntry]@{ MailboxId = 'mb1'; ItemId = 'o1'; FolderId = 'other'; ChangeKey = 'c9'; ItemType = 'IPM.Note'; Size = 50 })
    $s.UpsertItems($entries) | Out-Null
    $script:Bytes = @{}
    foreach ($id in 'm1', 'm2', 'm3', 'o1') {
        $b = [byte[]]::new($(if ($id -eq 'o1') { 50 } else { 100 * [int]$id.Substring(1) })); [Random]::new([int]$id[1]).NextBytes($b)
        $script:Bytes[$id] = $b
        $key = $s.FindItemKey('mb1', $id)
        $s.WriteItem($key, [IO.MemoryStream]::new($b), [Threading.CancellationToken]::None) | Out-Null
    }
    $s.Dispose()

    function Invoke-Tool([string[]]$Arguments) {
        $out = & pwsh -NoProfile -File $script:Tool -StorePath $script:StorePath @Arguments 2>&1 | Out-String
        [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $out }
    }
}

AfterAll { Remove-Item -LiteralPath $script:Base -Recurse -Force -ErrorAction SilentlyContinue }

Describe 'Export-FxCorpusFromStore.ps1' {
    It 'copies the streams of the chosen folder byte for byte and writes the index' {
        $dest = Join-Path $script:Base 'out1'
        $r = Invoke-Tool @('-OutputPath', $dest, '-Folder', 'MBT-FXICS-Corpus')
        $r.ExitCode | Should -Be 0 -Because $r.Output
        $files = @(Get-ChildItem -LiteralPath $dest -Filter '*.fxs')
        $files.Count | Should -Be 3
        $index = Get-Content (Join-Path $dest 'index.json') -Raw | ConvertFrom-Json
        @($index).Count | Should -Be 3
        foreach ($e in $index) {
            $e.FolderPath | Should -Be 'MBT-FXICS-Corpus/Mail'
            (Get-FileHash (Join-Path $dest $e.File) -Algorithm SHA256).Hash | Should -Be $e.Sha256
            [IO.File]::ReadAllBytes((Join-Path $dest $e.File)) | Should -Be $script:Bytes[$e.ItemId]
        }
    }

    It 'takes every folder by default, keeps files already there, and does nothing with -WhatIf' {
        $dest = Join-Path $script:Base 'out2'
        $w = Invoke-Tool @('-OutputPath', $dest, '-WhatIf')
        $w.ExitCode | Should -Be 0 -Because $w.Output
        Test-Path -LiteralPath $dest | Should -BeFalse
        $r = Invoke-Tool @('-OutputPath', $dest)
        @(Get-ChildItem -LiteralPath $dest -Filter '*.fxs').Count | Should -Be 4
        $again = Invoke-Tool @('-OutputPath', $dest)
        $again.ExitCode | Should -Be 0 -Because $again.Output
        $again.Output | Should -Match 'kept'
        $again.Output | Should -Not -Match 'copied'
    }

    It 'fails when the folder holds nothing' {
        (Invoke-Tool @('-OutputPath', (Join-Path $script:Base 'out3'), '-Folder', 'Nope')).ExitCode | Should -Be 1
    }
}
