#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the choice of folders (Engine.FolderFilter.cs), the preview of a PST message (Engine.PstPreview.cs) and the copy of chosen folders and
    messages of a PST into new PST files (Engine.PstCopy.cs, Copy-MbtPst). Nothing here reaches the network.

    Run:  Invoke-Pester -Path .\tests\PstCopy.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    $script:P = [MailboxTransfer.FxProperty]

    function New-Msg {
        param([string]$Subject, [string]$Class = 'IPM.Note', [int]$AttachBytes = 0, [switch]$Embedded, [string]$Html, [string]$Text = 'plain body of the message')
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($script:P::FromString(0x001A, $Class))
        $m.Properties.Add($script:P::FromString(0x0037, $Subject))
        if (-not $Html) { $m.Properties.Add($script:P::FromString(0x1000, $Text)) }
        else { $m.Properties.Add($script:P::FromBinary(0x1013, [Text.Encoding]::UTF8.GetBytes($Html))); $m.Properties.Add($script:P::FromInt32(0x3FDE, 65001)) }
        $m.Properties.Add($script:P::FromString(0x0C1A, 'Alice Martin'))
        $m.Properties.Add($script:P::FromString(0x5D01, 'alice@contoso.example'))
        $m.Properties.Add($script:P::FromTime(0x0039, [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)))
        $m.Properties.Add($script:P::FromInt32(0x0E07, 1))
        foreach ($rcp in @(@(1, 'Bob', 'bob@contoso.example'), @(2, 'Carol', 'carol@contoso.example'))) {
            $r = [MailboxTransfer.FxRecipient]::new()
            $r.Properties.Add($script:P::FromInt32(0x0C15, $rcp[0])); $r.Properties.Add($script:P::FromString(0x3001, $rcp[1])); $r.Properties.Add($script:P::FromString(0x39FE, $rcp[2]))
            $m.Recipients.Add($r)
        }
        if ($AttachBytes -gt 0) {
            $a = [MailboxTransfer.FxAttachment]::new()
            $b = [byte[]]::new($AttachBytes); [Random]::new(7).NextBytes($b)
            $a.Properties.Add($script:P::FromInt32(0x3705, 1)); $a.Properties.Add($script:P::FromBinary(0x3701, $b)); $a.Properties.Add($script:P::FromString(0x3707, 'report.bin')); $a.Properties.Add($script:P::FromString(0x370E, 'application/octet-stream'))
            $m.Attachments.Add($a)
        }
        if ($Embedded) {
            $e = [MailboxTransfer.FxMessage]::new()
            $e.Properties.Add($script:P::FromString(0x001A, 'IPM.Note')); $e.Properties.Add($script:P::FromString(0x0037, 'Inner'))
            $ea = [MailboxTransfer.FxAttachment]::new()
            $ea.Properties.Add($script:P::FromInt32(0x3705, 5)); $ea.Properties.Add($script:P::Create(0x3701, 0x000D, [byte[]]@()))
            $ea.Embedded = $e
            $m.Attachments.Add($ea)
        }
        return $m
    }

    # Inbox (3), Inbox/Projects (2), Sent Items (1), Deleted Items (1), Calendar (2): 9 messages.
    function New-SourcePst {
        param([string]$Name, [int]$AttachBytes = 0)
        $path = Join-Path $TestDrive ($Name + '.pst')
        $w = [MailboxTransfer.PstWriter]::Create($path, 'Old archive')
        $inbox = $w.CreateFolder($null, 'Inbox', 'IPF.Note'); $proj = $w.CreateFolder($inbox, 'Projects', 'IPF.Note')
        $sent = $w.CreateFolder($null, 'Sent Items', 'IPF.Note'); $cal = $w.CreateFolder($null, 'Calendar', 'IPF.Appointment')
        foreach ($i in 1..3) { [void]$w.AddMessage($inbox, (New-Msg "in $i" -AttachBytes $AttachBytes)) }
        foreach ($i in 1..2) { [void]$w.AddMessage($proj, (New-Msg "proj $i" -AttachBytes $AttachBytes)) }
        [void]$w.AddMessage($sent, (New-Msg 'sent 1' -AttachBytes $AttachBytes))
        [void]$w.AddMessage($w.DeletedItems, (New-Msg 'deleted 1' -AttachBytes $AttachBytes))
        foreach ($i in 1..2) { [void]$w.AddMessage($cal, (New-Msg "meeting $i" -Class 'IPM.Appointment')) }
        $w.Close()
        return $path
    }

    # path -> subjects, of every PST file of a copy
    function Get-Tree([string[]]$Files) {
        $map = @{}
        foreach ($f in $Files) {
            $r = [MailboxTransfer.PstReader]::Open($f)
            try {
                $top = [MailboxTransfer.PstImportPlanner]::TopFolder($r)
                $walk = $null
                $walk = { param($node, $path)
                    $subjects = @($r.MessageNids($node, $false) | ForEach-Object { $r.ReadMessage($_).GetString(0x0037) })
                    if (-not $map.ContainsKey($path)) { $map[$path] = [System.Collections.Generic.List[string]]::new() }
                    foreach ($s in $subjects) { $map[$path].Add($s) }
                    foreach ($c in $node.Children) { & $walk $c $(if ($path) { "$path/$($c.Name)" } else { $c.Name }) }
                }
                & $walk $top ''
            } finally { $r.Dispose() }
        }
        return $map
    }

    function Get-Nid([string]$Pst, [string]$Subject) {
        $r = [MailboxTransfer.PstReader]::Open($Pst)
        try {
            foreach ($f in $r.AllFolders()) { foreach ($n in $r.MessageNids($f, $false)) { if ($r.ReadMessage($n).GetString(0x0037) -eq $Subject) { return $n } } }
        } finally { $r.Dispose() }
    }
}

Describe 'the choice of folders (FolderFilter)' {
    It 'a path takes the folder and what is below it; wildcards; case does not matter' {
        $f = [MailboxTransfer.FolderFilter]::new([string[]]@('inbox'), [string[]]@())
        $f.Allows('Inbox') | Should -BeTrue
        $f.Allows('Inbox/Projects/2019') | Should -BeTrue
        $f.Allows('Sent Items') | Should -BeFalse
        $f.Allows('Inbox2') | Should -BeFalse
        $w = [MailboxTransfer.FolderFilter]::new([string[]]@('*/2019', 'Arch*'), [string[]]@())
        $w.Allows('Mail/2019') | Should -BeTrue
        $w.Allows('Mail/2019/Q1') | Should -BeTrue
        $w.Allows('Archive/Old') | Should -BeTrue
        $w.Allows('Mail/2020') | Should -BeFalse
        $q = [MailboxTransfer.FolderFilter]::new([string[]]@('Folder?'), [string[]]@())
        $q.Allows('Folder1') | Should -BeTrue
        $q.Allows('Folder12') | Should -BeFalse
    }
    It 'an exclusion wins; Inbox without its subfolders is Inbox with Inbox/* excluded; only excluding takes everything else' {
        $f = [MailboxTransfer.FolderFilter]::new([string[]]@('Inbox'), [string[]]@('Inbox/*'))
        $f.Allows('Inbox') | Should -BeTrue
        $f.Allows('Inbox/Projects') | Should -BeFalse
        $f.Allows('Inbox/Projects/Deep') | Should -BeFalse
        $e = [MailboxTransfer.FolderFilter]::new([string[]]@(), [string[]]@('Deleted Items', 'Junk*'))
        $e.Allows('Inbox') | Should -BeTrue
        $e.Allows('Deleted Items/Old') | Should -BeFalse
        $e.Allows('Junk Email') | Should -BeFalse
        [MailboxTransfer.FolderFilter]::Create([string[]]@(), [string[]]@()) | Should -BeNullOrEmpty
        $both = [MailboxTransfer.FolderFilter]::new([string[]]@('Inbox'), [string[]]@('Inbox'))
        $both.Allows('Inbox') | Should -BeFalse
    }
    It 'tells which patterns matched no folder' {
        $f = [MailboxTransfer.FolderFilter]::new([string[]]@('Inbox', 'Inbx'), [string[]]@('Nope'))
        $null = $f.Allows('Inbox'); $null = $f.Allows('Sent Items')
        @($f.Unmatched()) | Should -Be @("-IncludeFolder 'Inbx'", "-ExcludeFolder 'Nope'")
    }
}

Describe 'the preview of a PST message' {
    BeforeAll {
        $script:Src = Join-Path $TestDrive 'preview.pst'
        $w = [MailboxTransfer.PstWriter]::Create($script:Src, 'Preview')
        $f = $w.CreateFolder($null, 'Inbox', 'IPF.Note')
        [void]$w.AddMessage($f, (New-Msg 'with text' -AttachBytes 5000 -Embedded -Text ('x' * 50)))
        [void]$w.AddMessage($f, (New-Msg 'with html' -Html '<html><head><style>p{}</style></head><body><p>Hello <b>world</b></p><p>second</p></body></html>'))
        [void]$w.AddMessage($f, (New-Msg 'long' -Text ('0123456789' * 500)))
        $w.Close()
    }
    It 'gives the header, the recipients, the text body and the attachments (name, type, size, embedded message)' {
        $p = Get-MbtPstPreview -Path $script:Src -Nid (Get-Nid $script:Src 'with text')
        $p.Subject | Should -Be 'with text'
        $p.FromName | Should -Be 'Alice Martin'
        $p.FromAddress | Should -Be 'alice@contoso.example'
        $p.FromDisplay | Should -Be 'Alice Martin <alice@contoso.example>'
        @($p.To) | Should -Be @('Bob <bob@contoso.example>')
        @($p.Cc) | Should -Be @('Carol <carol@contoso.example>')
        $p.SentUtc | Should -Be ([datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc))
        $p.BodySource | Should -Be 'text'
        $p.BodyText | Should -Be ('x' * 50)
        $p.Attachments.Count | Should -Be 2
        $p.Attachments[0].Name | Should -Be 'report.bin'
        $p.Attachments[0].Size | Should -Be 5000
        $p.Attachments[0].MimeType | Should -Be 'application/octet-stream'
        $p.Attachments[0].IsMessage | Should -BeFalse
        $p.Attachments[1].IsMessage | Should -BeTrue
        $p.Attachments[1].EmbeddedSubject | Should -Be 'Inner'
        $p.Attachments[1].Name | Should -Be 'Inner'
    }
    It 'shows the text of an HTML-only message and keeps the HTML; cuts a long body' {
        $p = Get-MbtPstPreview -Path $script:Src -Nid (Get-Nid $script:Src 'with html')
        $p.BodySource | Should -Be 'html'
        $p.BodyText | Should -BeLike '*Hello world*'
        $p.BodyText | Should -Not -BeLike '*<*'
        $p.BodyHtml | Should -BeLike '<html>*'
        $long = Get-MbtPstPreview -Path $script:Src -Nid (Get-Nid $script:Src 'long') -BodyChars 100
        $long.BodyText.Length | Should -Be 100
        $long.BodyTruncated | Should -BeTrue
        (Get-MbtPstPreview -Path $script:Src -Nid (Get-Nid $script:Src 'long')).BodyTruncated | Should -BeFalse
    }
}

Describe 'the copy of chosen folders and messages of a PST (Copy-MbtPst)' {
    BeforeAll { $script:Pst = New-SourcePst 'source' }
    It 'copies everything when nothing is chosen, with the folder tree, and the new file is clean for the checker' {
        $dest = Join-Path $TestDrive 'all.pst'
        $r = Copy-MbtPst -Path $script:Pst -Destination $dest -Confirm:$false
        $r.ExitCode | Should -Be 0
        $r.Result.ItemsWritten | Should -Be 9
        $r.Result.Parts.Count | Should -Be 1
        $r.Result.Parts[0].Checked | Should -BeTrue
        $r.Result.Parts[0].CheckErrors | Should -Be 0
        (Test-MbtPst -Path $dest).Errors | Should -Be 0
        $t = Get-Tree @($dest)
        @($t['Inbox']) | Should -Be @('in 1', 'in 2', 'in 3')
        @($t['Inbox/Projects']) | Should -Be @('proj 1', 'proj 2')
        @($t['Sent Items']) | Should -Be @('sent 1')
        @($t['Calendar'].Count) | Should -Be 2
        # one Deleted Items, the source's (the writer makes its own)
        @($t.Keys | Where-Object { $_ -eq 'Deleted Items' }).Count | Should -Be 1
        @($t['Deleted Items']) | Should -Be @('deleted 1')
    }
    It 'copies the chosen folders only (a path takes the subfolders; an exclusion wins) and reports a pattern that matched nothing' {
        $dest = Join-Path $TestDrive 'inbox.pst'
        $r = Copy-MbtPst -Path $script:Pst -Destination $dest -IncludeFolder 'Inbox', 'Nope' -ExcludeFolder 'Inbox/Projects' -Confirm:$false
        $r.ExitCode | Should -Be 0
        $r.Result.ItemsWritten | Should -Be 3
        $r.Plan.Notes -join ' ' | Should -BeLike "*-IncludeFolder 'Nope' matches no folder*"
        $t = Get-Tree @($dest)
        @($t['Inbox']) | Should -Be @('in 1', 'in 2', 'in 3')
        $t.ContainsKey('Sent Items') | Should -BeFalse
        $t.ContainsKey('Inbox/Projects') | Should -BeFalse
        $dest2 = Join-Path $TestDrive 'proj.pst'
        $r2 = Copy-MbtPst -Path $script:Pst -Destination $dest2 -IncludeFolder 'Inbox/Projects' -Confirm:$false
        $r2.Result.ItemsWritten | Should -Be 2
        $t2 = Get-Tree @($dest2)
        @($t2['Inbox'].Count) | Should -Be 0     # created because the tree needs it, holding nothing
        @($t2['Inbox/Projects']) | Should -Be @('proj 1', 'proj 2')
    }
    It 'copies messages chosen one by one (node ids), in their folders, with the folders above them' {
        $dest = Join-Path $TestDrive 'items.pst'
        $nids = [uint32[]]@((Get-Nid $script:Pst 'in 2'), (Get-Nid $script:Pst 'sent 1'))
        $r = Copy-MbtPst -Path $script:Pst -Destination $dest -ItemNid $nids -Confirm:$false
        $r.Result.ItemsWritten | Should -Be 2
        $t = Get-Tree @($dest)
        @($t['Inbox']) | Should -Be @('in 2')
        @($t['Sent Items']) | Should -Be @('sent 1')
        $t.ContainsKey('Calendar') | Should -BeFalse
        # items in addition to a folder
        $dest2 = Join-Path $TestDrive 'items2.pst'
        $r2 = Copy-MbtPst -Path $script:Pst -Destination $dest2 -IncludeFolder 'Calendar' -ItemNid ([uint32[]]@(Get-Nid $script:Pst 'sent 1')) -Confirm:$false
        $r2.Result.ItemsWritten | Should -Be 3
    }
    It 'splits on item boundaries at the limit and checks every part' {
        $big = New-SourcePst 'big' -AttachBytes 600000
        $dest = Join-Path $TestDrive 'split.pst'
        $r = Copy-MbtPst -Path $big -Destination $dest -MaxSizeMiB 2 -Confirm:$false
        $r.Result.ItemsWritten | Should -Be 9
        $r.Result.Parts.Count | Should -BeGreaterThan 1
        foreach ($p in $r.Result.Parts) { $p.Checked | Should -BeTrue; $p.CheckErrors | Should -Be 0 }
        Test-Path (Join-Path $TestDrive 'split-002.pst') | Should -BeTrue
        $t = Get-Tree @($r.Result.Parts | ForEach-Object { $_.Path })
        @($t['Inbox']).Count | Should -Be 3
        @($t['Sent Items']).Count | Should -Be 1
    }
    It 'never overwrites a file, leaves no file when it is stopped, and refuses the source as destination' {
        $dest = Join-Path $TestDrive 'exists.pst'
        $null = Copy-MbtPst -Path $script:Pst -Destination $dest -Confirm:$false
        { Copy-MbtPst -Path $script:Pst -Destination $dest -Confirm:$false } | Should -Throw '*never overwrites*'
        (Copy-MbtPst -Path $script:Pst -Destination $dest -Force -Confirm:$false).ExitCode | Should -Be 0
        { Copy-MbtPst -Path $script:Pst -Destination $script:Pst -Force -Confirm:$false } | Should -Throw '*source*'
        $cts = [Threading.CancellationTokenSource]::new(); $cts.Cancel()
        $stopped = Join-Path $TestDrive 'stopped.pst'
        $r = Copy-MbtPst -Path $script:Pst -Destination $stopped -CancellationToken $cts.Token -Confirm:$false
        $r.ExitCode | Should -Be 2
        Test-Path $stopped | Should -BeFalse
        { Copy-MbtPst -Path $script:Pst -Destination (Join-Path $TestDrive 'aux.pst') -Confirm:$false } | Should -Throw '*Windows*'
    }
    It 'with -WhatIf plans and writes nothing; the source is never changed' {
        $before = (Get-FileHash -LiteralPath $script:Pst).Hash
        $dest = Join-Path $TestDrive 'whatif.pst'
        $r = Copy-MbtPst -Path $script:Pst -Destination $dest -IncludeFolder 'Inbox' -WhatIf
        $r.Plan.Items | Should -Be 5
        $r.Plan.Folders.Count | Should -BeGreaterThan 1
        $r.Result | Should -BeNullOrEmpty
        Test-Path $dest | Should -BeFalse
        (Get-FileHash -LiteralPath $script:Pst).Hash | Should -Be $before
    }
}
