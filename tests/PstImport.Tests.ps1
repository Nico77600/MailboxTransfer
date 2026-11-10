#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the import from PST (Engine.FxImport.cs, Engine.PstImportPlan.cs, Engine.PstImportJob.cs, Engine.FxDiff.cs):
    the conversion of a PST message for importItem, the files and the folder plan, the journal, the job (resume, throttling, refusals,
    reconciliation, cancellation), damaged files, split PSTs and the public function Import-MbtPst. Nothing here reaches the network:
    the job talks to FakeMailboxService (tests\FakeMailboxService.cs).

    Run:  Invoke-Pester -Path .\tests\PstImport.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    . (Join-Path $PSScriptRoot 'support\PstDamage.ps1')
    $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:Dst = "MBX:99999999-8888-7777-6666-555555555555@$($script:Tenant)"
    $script:P = [MailboxTransfer.FxProperty]

    # A message of the PST: class, subject, one recipient, a named property, optionally a file attachment, an embedded message, a big body.
    function New-PstMsg {
        param([string]$Subject, [string]$Class = 'IPM.Note', [int]$AttachBytes = 0, [switch]$Embedded, [int]$BodyChars = 400)
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($script:P::FromString(0x001A, $Class))
        $m.Properties.Add($script:P::FromString(0x0037, $Subject))
        $m.Properties.Add($script:P::FromString(0x1000, ('body ' * $BodyChars)))
        $m.Properties.Add($script:P::FromTime(0x0039, [datetime]::new(2026, 10, 21, 9, 30, 0, [DateTimeKind]::Utc)))
        $m.Properties.Add($script:P::FromInt32(0x0E07, 1))
        $m.Properties.Add($script:P::FromInt32(0x8001, 5).Named([guid]'11111111-2222-3333-4444-555555555555', [uint32]0x8501))
        $r = [MailboxTransfer.FxRecipient]::new()
        $r.Properties.Add($script:P::FromInt32(0x0C15, 1)); $r.Properties.Add($script:P::FromString(0x3001, 'Bob')); $r.Properties.Add($script:P::FromString(0x39FE, 'bob@contoso.example'))
        $m.Recipients.Add($r)
        if ($AttachBytes -gt 0) {
            $a = [MailboxTransfer.FxAttachment]::new()
            $b = [byte[]]::new($AttachBytes); [Random]::new(7).NextBytes($b)
            $a.Properties.Add($script:P::FromInt32(0x3705, 1)); $a.Properties.Add($script:P::FromBinary(0x3701, $b)); $a.Properties.Add($script:P::FromString(0x3707, 'file.bin'))
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

    # Inbox (3, one with a 200 KB attachment, one with an embedded message), Inbox/Projects (2), Sent Items (1), Calendar (2), Contacts (1) : 9 messages.
    function New-ImportPst {
        param([string]$Name, [string]$Store = 'Old archive')
        $path = Join-Path $TestDrive ($Name + '.pst')
        $w = [MailboxTransfer.PstWriter]::Create($path, $Store)
        $inbox = $w.CreateFolder($null, 'Inbox', 'IPF.Note'); $proj = $w.CreateFolder($inbox, 'Projects', 'IPF.Note')
        $sent = $w.CreateFolder($null, 'Sent Items', 'IPF.Note'); $cal = $w.CreateFolder($null, 'Calendar', 'IPF.Appointment'); $con = $w.CreateFolder($null, 'Contacts', 'IPF.Contact')
        [void]$w.AddMessage($inbox, (New-PstMsg 'in 1'))
        [void]$w.AddMessage($inbox, (New-PstMsg 'in 2 with attachment' -AttachBytes 200000))
        [void]$w.AddMessage($inbox, (New-PstMsg 'in 3 with embedded' -Embedded))
        [void]$w.AddMessage($proj, (New-PstMsg 'proj 1')); [void]$w.AddMessage($proj, (New-PstMsg 'proj 2'))
        [void]$w.AddMessage($sent, (New-PstMsg 'sent 1'))
        [void]$w.AddMessage($cal, (New-PstMsg 'meeting 1' -Class 'IPM.Appointment')); [void]$w.AddMessage($cal, (New-PstMsg 'meeting 2' -Class 'IPM.Appointment'))
        [void]$w.AddMessage($con, (New-PstMsg 'Contact One' -Class 'IPM.Contact'))
        $w.Close()
        return $path
    }

    function New-FakeClient {
        param([Parameter(Mandatory)]$Service)
        $o = [MailboxTransfer.GraphOptions]::new()
        $o.DelayScale = 0
        $o.UseJitter = $false
        return [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Service)
    }

    function New-TargetService {
        $svc = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $null = $svc.AddMailbox($script:Dst)
        $f = $svc.AddFolder($script:Dst, 'F-INBOX', 'Inbox'); $f.WellKnown = 'inbox'
        $f = $svc.AddFolder($script:Dst, 'F-SENT', 'Sent Items'); $f.WellKnown = 'sentitems'
        $f = $svc.AddFolder($script:Dst, 'F-CAL', 'Calendar', 'ROOT', 'IPF.Appointment'); $f.WellKnown = 'calendar'
        return $svc
    }

    function New-Options {
        param([string[]]$Paths, [string]$TargetFolder = 'Imported/old')
        $o = [MailboxTransfer.PstImportOptions]::new()
        foreach ($p in $Paths) { $o.PstPaths.Add($p) }
        $o.TargetMailboxId = $script:Dst; $o.TargetFolder = $TargetFolder
        $o.Concurrency = 2; $o.ReconcileDelaySeconds = 0; $o.WorkPath = Join-Path $TestDrive 'work'
        return $o
    }

    # Runs a PstImportJob to its end; returns the job, its messages and the client.
    function Invoke-PstJob {
        param([Parameter(Mandatory)]$Service, [Parameter(Mandatory)]$Options, [Parameter(Mandatory)][string]$Journal, [int]$TimeoutSeconds = 90, [int]$CancelAfterMs = 0)
        $client = New-FakeClient -Service $Service
        $j = [MailboxTransfer.PstImportJournal]::new($Journal)
        try {
            $job = [MailboxTransfer.PstImportJob]::Start([MailboxTransfer.MailboxApi]::new($client), $j, $Options, (Join-Path $TestDrive 'work'), $null)
            $messages = [System.Collections.Generic.List[string]]::new()
            $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
            $cancelAt = if ($CancelAfterMs -gt 0) { [DateTime]::UtcNow.AddMilliseconds($CancelAfterMs) } else { $null }
            while (-not $job.Wait(100)) {
                foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
                if ($cancelAt -and [DateTime]::UtcNow -gt $cancelAt) { $job.Cancel(); $cancelAt = $null }
                if ([DateTime]::UtcNow -gt $deadline) { $job.Cancel(); throw 'the PST import job did not finish in time' }
            }
            foreach ($m in $job.TakeMessages()) { $messages.Add($m) }
            return [pscustomobject]@{ Job = $job; Messages = $messages; Client = $client; Result = $job.Result; Error = $job.Error }
        } finally { $j.Dispose() }
    }

    function Get-Subjects($svc, [string]$FolderId) {
        @($svc.ItemsOf($script:Dst, $FolderId) | ForEach-Object { [MailboxTransfer.FxCodec]::Decode($_.Data).Message.GetString(0x0037) } | Sort-Object)
    }
    function Get-FolderIdByName($svc, [string]$Name, [string]$ParentId) {
        (@($svc.FoldersOf($script:Dst) | Where-Object { $_.Name -eq $Name -and ((-not $ParentId) -or $_.ParentId -eq $ParentId) }) | Select-Object -First 1).Id
    }
}

Describe 'the conversion of a PST message for importItem' {
    It 'adds FXDelProp before the recipients and the attachments of the message, puts PidTagRowid first in every recipient, starts an embedded message with PidTagMid and drops the empty PidTagAttachDataObject' {
        $m = New-PstMsg 'x' -Embedded -AttachBytes 100
        $m.Recipients[0].Properties.Add($script:P::FromInt32(0x3000, 0))   # at the end, as a PST table lists the columns
        [void][MailboxTransfer.FxImportStream]::Prepare($m)
        $m.DelPropRecipients | Should -Not -BeNullOrEmpty
        $m.DelPropAttachments | Should -Not -BeNullOrEmpty
        $m.Recipients[0].Properties[0].Id | Should -Be 0x3000
        $emb = $m.Attachments | Where-Object { $_.Embedded } | Select-Object -First 1
        $emb.Embedded.Properties[0].Id | Should -Be 0x674A
        $emb.Embedded.DelPropRecipients | Should -BeNullOrEmpty
        @($emb.Properties | Where-Object { $_.Id -eq 0x3701 }).Count | Should -Be 0
    }
    It 'gives a stream that starts with the header record of the envelope and decodes back to the same message' {
        $m = [MailboxTransfer.FxImportStream]::Prepare((New-PstMsg 'round trip' -AttachBytes 5000))
        $src = [MailboxTransfer.FxMessageSource]::new($m, [MailboxTransfer.FxImportStream]::NewEnvelope())
        $ms = [IO.MemoryStream]::new(); $s = $src.OpenRead(); $s.CopyTo($ms); $s.Dispose()
        $bytes = $ms.ToArray()
        $bytes.Length | Should -Be $src.Length
        [BitConverter]::ToUInt32($bytes, 0) | Should -Be 1          # record type 1 ...
        [BitConverter]::ToUInt32($bytes, 4) | Should -Be 8          # ... of 8 bytes
        $doc = [MailboxTransfer.FxCodec]::Decode($bytes)
        $doc.Ok | Should -BeTrue
        $doc.Message.GetString(0x0037) | Should -Be 'round trip'
        $doc.Message.Recipients.Count | Should -Be 1
        $doc.Message.Attachments.Count | Should -Be 1
        $doc.Message.DelPropRecipients | Should -Not -BeNullOrEmpty
    }
}

Describe 'FxDiff' {
    It 'lists what changed, what was dropped and what was added' {
        $a = New-PstMsg 'one'; $b = New-PstMsg 'two'
        $b.Properties.RemoveAll([Predicate[MailboxTransfer.FxProperty]]{ param($p) $p.Id -eq 0x0E07 }) | Out-Null
        $b.Properties.Add($script:P::FromInt32(0x6000, 1))
        $d = @([MailboxTransfer.FxDiff]::Compare($a, $b))
        ($d | Where-Object { $_.Kind -eq 'Changed' -and $_.PropertyKey -eq '0037' }).Sent | Should -Be 'one'
        ($d | Where-Object { $_.Kind -eq 'Dropped' -and $_.PropertyKey -eq '0E07' }) | Should -Not -BeNullOrEmpty
        ($d | Where-Object { $_.Kind -eq 'Added' -and $_.PropertyKey -eq '6000' }) | Should -Not -BeNullOrEmpty
        @([MailboxTransfer.FxDiff]::Compare($a, $a)).Count | Should -Be 0
    }
}

Describe 'a silent loss of recipients or attachments can never pass the verification' {
    It 'counts recipients, attachments and embedded messages, embedded messages included' {
        $m = New-PstMsg 'x' -Embedded -AttachBytes 100
        $m.Attachments | Where-Object { $_.Embedded } | ForEach-Object { $r = [MailboxTransfer.FxRecipient]::new(); $r.Properties.Add($script:P::FromString(0x3001, 'inner rcpt')); $_.Embedded.Recipients.Add($r) }
        [MailboxTransfer.FxDiff]::TotalsText($m) | Should -Be '2 recipients, 2 attachments, 1 embedded messages'
    }
    It 'FxDiff reports a message that came back without its recipients and attachments (the service accepted it and dropped them)' {
        $sent = New-PstMsg 'x' -AttachBytes 100 -Embedded
        $got = New-PstMsg 'x'
        $got.Recipients.Clear()
        $d = @([MailboxTransfer.FxDiff]::Compare($sent, $got))
        ($d | Where-Object PropertyKey -EQ 'recipients').Sent | Should -Be '1'
        ($d | Where-Object PropertyKey -EQ 'recipients').Got | Should -Be '0'
        ($d | Where-Object PropertyKey -EQ 'attachments').Sent | Should -Be '2'
        [MailboxTransfer.FxDiff]::TotalsText($sent) | Should -Not -Be ([MailboxTransfer.FxDiff]::TotalsText($got))
    }
    It 'every message the job sends carries FXDelProp for its recipients and its attachments (as an export does), also when it has none' {
        $pst = New-ImportPst 'delprop'
        $svc = New-TargetService
        $run = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal (Join-Path $TestDrive 'delprop.db')
        $run.Result.JobDone | Should -Be 9
        foreach ($it in @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' })) {
            $m = [MailboxTransfer.FxCodec]::Decode($it.Data).Message
            $m.DelPropRecipients | Should -Not -BeNullOrEmpty
            $m.DelPropAttachments | Should -Not -BeNullOrEmpty
            $m.Recipients[0].Properties[0].Id | Should -Be 0x3000
        }
    }
}
Describe 'the files of an import' {
    It 'takes the parts of a split PST in order (name.pst, name-0002.pst, ..., name-0010.pst), folder by folder' {
        $dir = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'parts')
        foreach ($n in 'b.pst', 'a-0010.pst', 'a-0003.pst', 'a.pst', 'a-0002.pst') { Set-Content -LiteralPath (Join-Path $dir $n) -Value 'x' }
        $names = @([MailboxTransfer.PstImportFiles]::Resolve([string[]]@($dir.FullName)) | ForEach-Object { Split-Path $_ -Leaf })
        $names | Should -Be @('a.pst', 'a-0002.pst', 'a-0003.pst', 'a-0010.pst', 'b.pst')
    }
    It 'refuses a missing file and a folder without PST' {
        { [MailboxTransfer.PstImportFiles]::Resolve([string[]]@((Join-Path $TestDrive 'nope.pst'))) } | Should -Throw '*not found*'
        $empty = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'emptydir')
        { [MailboxTransfer.PstImportFiles]::Resolve([string[]]@($empty.FullName)) } | Should -Throw '*No .pst file*'
    }
}

Describe 'the plan' {
    BeforeAll { $script:Pst = New-ImportPst 'plan' }
    It 'maps the folder tree below the target folder, with the container classes, the number and the size of the items, and the estimated duration' {
        $o = New-Options @($script:Pst)
        $plan = [MailboxTransfer.PstImportPlanner]::Build($o, $null, $null, [Threading.CancellationToken]::None)
        $plan.Items | Should -Be 9
        $plan.Bytes | Should -BeGreaterThan 200000
        $plan.Blocked | Should -BeFalse
        $targets = @($plan.Folders | ForEach-Object TargetPath)
        $targets | Should -Contain 'Imported/old'
        $targets | Should -Contain 'Imported/old/Inbox'
        $targets | Should -Contain 'Imported/old/Inbox/Projects'
        $targets | Should -Contain 'Imported/old/Calendar'
        ($plan.Folders | Where-Object PstPath -EQ 'Calendar').ContainerClass | Should -Be 'IPF.Appointment'
        ($plan.Folders | Where-Object PstPath -EQ 'Inbox/Projects').Items | Should -Be 2
        $plan.PathsToCreate[0] | Should -Be 'Imported'
        $plan.EstimatedSeconds | Should -BeGreaterThan 0
        $plan.EstimatedSeconds | Should -BeLessThan 1   # 230 KB at 4 GiB per hour: 0.2 s
        $big = [MailboxTransfer.PstImportPlanner]::Build((New-Options @($script:Pst)), $null, $null, [Threading.CancellationToken]::None)
        [math]::Abs($big.EstimatedSeconds - $big.Bytes / (4 * 1GB) * 3600) | Should -BeLessThan 1e-9
    }
    It 'recognises the well-known folders of the PST and, with MergeWellKnown, maps them to the folders of the mailbox' {
        $svc = New-TargetService
        $api = [MailboxTransfer.MailboxApi]::new((New-FakeClient -Service $svc))
        $tree = $api.GetFolderTree($script:Dst)
        $o = New-Options @($script:Pst); $o.MergeWellKnown = $true
        $plan = [MailboxTransfer.PstImportPlanner]::Build($o, $tree, $null, [Threading.CancellationToken]::None)
        ($plan.Folders | Where-Object PstPath -EQ 'Inbox').WellKnown | Should -Be 'inbox'
        ($plan.Folders | Where-Object PstPath -EQ 'Inbox').TargetPath | Should -Be 'Inbox'
        ($plan.Folders | Where-Object PstPath -EQ 'Inbox/Projects').TargetPath | Should -Be 'Inbox/Projects'
        ($plan.Folders | Where-Object PstPath -EQ 'Sent Items').TargetPath | Should -Be 'Sent Items'
        ($plan.Folders | Where-Object PstPath -EQ 'Calendar').TargetPath | Should -Be 'Calendar'
        ($plan.Folders | Where-Object PstPath -EQ 'Contacts').WellKnown | Should -Be 'contacts'
        ($plan.Folders | Where-Object PstPath -EQ 'Contacts').TargetPath | Should -Be 'Imported/old/Contacts'   # no Contacts folder in this mailbox: below the target folder
        $plain = [MailboxTransfer.PstImportPlanner]::Build((New-Options @($script:Pst)), $tree, $null, [Threading.CancellationToken]::None)
        ($plain.Folders | Where-Object PstPath -EQ 'Inbox').TargetPath | Should -Be 'Imported/old/Inbox'   # never the real Inbox by default
    }
    It 'applies the folder map (longest prefix) and the per-file subfolder' {
        $o = New-Options @($script:Pst)
        $o.FolderMap = [System.Collections.Generic.Dictionary[string, string]]::new()
        $o.FolderMap['Inbox'] = 'Mail'; $o.FolderMap['Inbox/Projects'] = 'Work'
        $plan = [MailboxTransfer.PstImportPlanner]::Build($o, $null, $null, [Threading.CancellationToken]::None)
        ($plan.Folders | Where-Object PstPath -EQ 'Inbox').TargetPath | Should -Be 'Imported/old/Mail'
        ($plan.Folders | Where-Object PstPath -EQ 'Inbox/Projects').TargetPath | Should -Be 'Imported/old/Work'
        $o2 = New-Options @($script:Pst); $o2.PerFileSubfolder = $true
        $plan2 = [MailboxTransfer.PstImportPlanner]::Build($o2, $null, $null, [Threading.CancellationToken]::None)
        ($plan2.Folders | Where-Object PstPath -EQ 'Sent Items').TargetPath | Should -Be 'Imported/old/Old archive/Sent Items'
    }
    It 'plans only the folders chosen, says which pattern matched nothing' {
        $o = New-Options @($script:Pst); $o.IncludeFolders.Add('Sent Items'); $o.IncludeFolders.Add('Calendar'); $o.IncludeFolders.Add('Nope*')
        $plan = [MailboxTransfer.PstImportPlanner]::Build($o, $null, $null, [Threading.CancellationToken]::None)
        $plan.Items | Should -Be 3
        @($plan.Folders | ForEach-Object PstPath | Sort-Object) | Should -Be @('Calendar', 'Sent Items')
        $plan.Notes -join ' ' | Should -BeLike "*-IncludeFolder 'Nope*' matches no folder*"
        $o2 = New-Options @($script:Pst); $o2.ExcludeFolders.Add('Inbox'); $o2.ExcludeFolders.Add('Calendar')
        $plan2 = [MailboxTransfer.PstImportPlanner]::Build($o2, $null, $null, [Threading.CancellationToken]::None)
        $plan2.Items | Should -Be 2   # Sent Items and Contacts
    }
    It 'says a damaged file is blocked, with the findings' {
        $bad = New-TestPst $TestDrive 'plan-bad' -Faults @('dangling-nid')
        $plan = [MailboxTransfer.PstImportPlanner]::Build((New-Options @($bad)), $null, $null, [Threading.CancellationToken]::None)
        $plan.Blocked | Should -BeTrue
        $plan.Sources[0].CheckErrors | Should -BeGreaterThan 0
        $plan.Notes -join ' ' | Should -BeLike '*damaged*'
        $plan.Items | Should -Be 0
    }
}

Describe 'the import job' {
    It 'recreates the folder tree with the container classes and imports every message with its recipients, attachments and embedded message' {
        $pst = New-ImportPst 'basic'
        $svc = New-TargetService
        $run = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal (Join-Path $TestDrive 'basic.db')
        $run.Error | Should -BeNullOrEmpty
        $r = $run.Result
        $r.JobComplete | Should -BeTrue
        $r.JobDone | Should -Be 9
        $r.ExitCode | Should -Be 0
        $r.Done | Should -Be 9
        $names = @($svc.FoldersOf($script:Dst) | ForEach-Object Name)
        foreach ($n in 'Imported', 'old', 'Inbox', 'Projects', 'Sent Items', 'Calendar', 'Contacts') { $names | Should -Contain $n }
        $root = Get-FolderIdByName $svc 'old'
        $inbox = Get-FolderIdByName $svc 'Inbox' $root
        (Get-FolderIdByName $svc 'Calendar' $root | ForEach-Object { ($svc.FoldersOf($script:Dst) | Where-Object Id -EQ $_).Type }) | Should -Be 'IPF.Appointment'
        (($svc.FoldersOf($script:Dst) | Where-Object { $_.Name -eq 'Contacts' -and $_.ParentId -eq $root }).Type) | Should -Be 'IPF.Contact'
        Get-Subjects $svc $inbox | Should -Be @('in 1', 'in 2 with attachment', 'in 3 with embedded')
        # the real Inbox of the mailbox was not touched
        @($svc.ItemsOf($script:Dst, 'F-INBOX')).Count | Should -Be 0
        # the stream received: envelope, recipients, attachments with their bytes, embedded message
        $att = $svc.ItemsOf($script:Dst, $inbox) | Where-Object { [MailboxTransfer.FxCodec]::Decode($_.Data).Message.GetString(0x0037) -eq 'in 2 with attachment' }
        $doc = [MailboxTransfer.FxCodec]::Decode($att.Data)
        $doc.Envelope | Should -Not -BeNullOrEmpty
        $doc.Message.Recipients.Count | Should -Be 1
        $doc.Message.Attachments[0].Properties.Where({ $_.Id -eq 0x3701 })[0].RawLength | Should -Be 200000
        $emb = $svc.ItemsOf($script:Dst, $inbox) | Where-Object { [MailboxTransfer.FxCodec]::Decode($_.Data).Message.GetString(0x0037) -eq 'in 3 with embedded' }
        [MailboxTransfer.FxCodec]::Decode($emb.Data).Message.Attachments[0].Embedded.GetString(0x0037) | Should -Be 'Inner'
        $svc.RequestsOf('/importItem?') | Should -Be 9
    }
    It 'resumes: a second run with the same job creates nothing again, and a new part of the same set is the only thing imported' {
        $pst = New-ImportPst 'resume'
        $svc = New-TargetService
        $journal = Join-Path $TestDrive 'resume.db'
        $first = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $first.Result.JobDone | Should -Be 9
        $again = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $again.Result.Done | Should -Be 0
        $again.Result.JobComplete | Should -BeTrue
        $again.Result.JobTotal | Should -Be 9
        $svc.RequestsOf('/importItem?') | Should -Be 9
        @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 9
        # another file in the same job: only its messages are added
        $pst2 = Join-Path $TestDrive 'resume2.pst'
        $w = [MailboxTransfer.PstWriter]::Create($pst2, 'Second'); $f = $w.CreateFolder($null, 'Inbox', 'IPF.Note'); [void]$w.AddMessage($f, (New-PstMsg 'later')); $w.Close()
        $more = Invoke-PstJob -Service $svc -Options (New-Options @($pst, $pst2)) -Journal $journal
        $more.Result.Done | Should -Be 1
        $more.Result.JobTotal | Should -Be 10
        $svc.RequestsOf('/importItem?') | Should -Be 10
    }
    It 'imports the parts of a split PST in order as one tree' {
        $dir = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'split')
        $w = [MailboxTransfer.PstSplitWriter]::new((Join-Path $dir 'big.pst'), 'Big', 120000)
        foreach ($i in 1..8) { [void]$w.AddMessage('Inbox', 'IPF.Note', (New-PstMsg "msg $i" -AttachBytes 60000)) }
        $w.Close()
        $w.Files.Count | Should -BeGreaterThan 1
        $svc = New-TargetService
        $run = Invoke-PstJob -Service $svc -Options (New-Options @($dir.FullName)) -Journal (Join-Path $TestDrive 'split.db')
        $run.Result.JobComplete | Should -BeTrue
        $run.Result.JobDone | Should -Be 8
        $inbox = Get-FolderIdByName $svc 'Inbox' (Get-FolderIdByName $svc 'old')
        @($svc.FoldersOf($script:Dst) | Where-Object { $_.Name -eq 'Inbox' -and $_.ParentId -ne 'ROOT' }).Count | Should -Be 1   # one Inbox for all the parts
        (Get-Subjects $svc $inbox).Count | Should -Be 8
        $run.Result.Sources.Count | Should -Be $w.Files.Count
    }
    It 'obeys a 429 (Retry-After) and still imports every item once' {
        $pst = New-ImportPst 'throttle'
        $svc = New-TargetService
        $svc.Throttle(2, 0.1, '/importItem')
        $run = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal (Join-Path $TestDrive 'throttle.db')
        $run.Result.JobComplete | Should -BeTrue
        $svc.Throttled | Should -Be 2
        @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 9
    }
    It 'reports an item the service refuses (400) with the reason, goes on with the others; the job is complete (nothing to run again) and exits 1' {
        $pst = New-ImportPst 'refuse'
        $svc = New-TargetService
        $svc.ImportFailures.Enqueue(400)
        $run = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal (Join-Path $TestDrive 'refuse.db')
        $r = $run.Result
        $r.Done | Should -Be 8
        $r.Refused | Should -Be 1
        $r.JobRefused | Should -Be 1
        $r.JobComplete | Should -BeTrue      # every row is Done or Refused for good: running again changes nothing
        $r.JobClean | Should -BeFalse
        $r.ExitCode | Should -Be 1
        $r.JobSummary() | Should -BeLike '*complete, 1 item(s) refused for good*'
        $r.NotImported.Count | Should -Be 1
        $r.NotImported[0].State | Should -Be 'Refused'
        $r.NotImported[0].Reason | Should -BeLike 'refused by the service*'
        $r.NotImported[0].Subject | Should -Not -BeNullOrEmpty
        $r.NotImported[0].Folder | Should -Not -BeNullOrEmpty
        # a Refused item is not sent again by the next run
        $again = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal (Join-Path $TestDrive 'refuse.db')
        $again.Result.Done | Should -Be 0
        $svc.RequestsOf('/importItem?') | Should -Be 9
    }
    It 'tries a Failed item (refused with nothing created: 404) again in the next run' {
        $pst = New-ImportPst 'failed'
        $svc = New-TargetService
        $svc.ImportFailures.Enqueue(404)
        $journal = Join-Path $TestDrive 'failed.db'
        $first = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $first.Result.JobFailed | Should -Be 1
        $first.Result.ExitCode | Should -Be 2
        $second = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $second.Result.JobComplete | Should -BeTrue
        $second.Result.Done | Should -Be 1
        @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 9
    }
    It 'refuses for good an item that failed with nothing created on every attempt allowed (-MaxAttempts): the job is complete, exit 1' {
        $pst = New-ImportPst 'giveup'
        $svc = New-TargetService
        $svc.ImportFailures.Enqueue(404)
        $o = New-Options @($pst); $o.MaxAttempts = 1
        $run = Invoke-PstJob -Service $svc -Options $o -Journal (Join-Path $TestDrive 'giveup.db')
        $r = $run.Result
        $r.JobRefused | Should -Be 1
        $r.JobFailed | Should -Be 0
        $r.JobComplete | Should -BeTrue
        $r.ExitCode | Should -Be 1
        $r.NotImported[0].State | Should -Be 'Refused'
        $r.NotImported[0].Reason | Should -BeLike 'gave up after 1 attempt(s), nothing was created*'
        # with two attempts allowed the first failure is Failed (exit 2) and the second one is final
        $journal = Join-Path $TestDrive 'giveup2.db'
        $svc2 = New-TargetService; $svc2.ImportFailures.Enqueue(404)
        $o2 = New-Options @($pst); $o2.MaxAttempts = 2
        $first = Invoke-PstJob -Service $svc2 -Options $o2 -Journal $journal
        $first.Result.JobFailed | Should -Be 1
        $first.Result.ExitCode | Should -Be 2
        $svc2.ImportFailures.Enqueue(404)
        $second = Invoke-PstJob -Service $svc2 -Options $o2 -Journal $journal
        $second.Result.JobRefused | Should -Be 1
        $second.Result.ExitCode | Should -Be 1
    }
    It 'a folder of the target that cannot be created leaves its items as Failed rows (job not complete, exit 2), and the next run imports them once the folder exists' {
        $pst = New-ImportPst 'nofolder'
        $svc = New-TargetService
        $svc.RefuseFolderNames.Add('Sent Items'); $svc.RefuseFolderStatus = 400
        $journal = Join-Path $TestDrive 'nofolder.db'
        $first = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $r = $first.Result
        $r.JobDone | Should -Be 8
        $r.JobFailed | Should -Be 1
        $r.JobTotal | Should -Be 9
        $r.JobComplete | Should -BeFalse
        $r.ExitCode | Should -Be 2
        ($r.Problems -join ' ') | Should -BeLike "*could not be created*"
        $r.NotImported[0].Reason | Should -BeLike "*could not be created*"
        $svc.RefuseFolderNames.Clear()
        $second = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $second.Result.Done | Should -Be 1
        $second.Result.JobComplete | Should -BeTrue
        $second.Result.ExitCode | Should -Be 0
        $sent = Get-FolderIdByName $svc 'Sent Items' (Get-FolderIdByName $svc 'old')
        Get-Subjects $svc $sent | Should -Be @('sent 1')
    }
    It 'refuses to resume a job with other target options (target folder, folder map, well-known folders): the rest would go to other folders' {
        $pst = New-ImportPst 'options'
        $svc = New-TargetService
        $journal = Join-Path $TestDrive 'options.db'
        $first = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $first.Result.JobComplete | Should -BeTrue
        $other = Invoke-PstJob -Service $svc -Options (New-Options @($pst) 'Elsewhere') -Journal $journal
        $other.Error.Message | Should -BeLike "*other target options*Give the same options, or another -JobName*"
        $merge = New-Options @($pst); $merge.MergeWellKnown = $true
        (Invoke-PstJob -Service $svc -Options $merge -Journal $journal).Error.Message | Should -BeLike '*other target options*'
        $map = New-Options @($pst); $map.FolderMap = [System.Collections.Generic.Dictionary[string, string]]::new(); $map.FolderMap['Inbox'] = 'Mail'
        (Invoke-PstJob -Service $svc -Options $map -Journal $journal).Error.Message | Should -BeLike '*other target options*'
        # the same options resume; choosing other folders of the same PST is allowed
        $same = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $same.Error | Should -BeNullOrEmpty
        $same.Result.JobComplete | Should -BeTrue
        $svc.RequestsOf('/importItem?') | Should -Be 9
    }
    It 'a 2xx answer of importItem that cannot be read leaves the item Unknown and never sends it again; an empty answer is a Done item without id' {
        $pst = New-ImportPst 'answer'
        $svc = New-TargetService
        $svc.ImportAnswers.Enqueue('garbage')
        $o = New-Options @($pst); $o.Concurrency = 1; $o.AutoReconcile = $false
        $run = Invoke-PstJob -Service $svc -Options $o -Journal (Join-Path $TestDrive 'answer.db')
        $r = $run.Result
        $r.JobUnknown | Should -Be 1
        $r.JobFailed | Should -Be 0
        $r.JobDone | Should -Be 8
        $r.ExitCode | Should -Be 3
        $r.NotImported[0].State | Should -Be 'Unknown'
        $r.NotImported[0].Reason | Should -BeLike '*answer*'
        $svc.RequestsOf('/importItem?') | Should -Be 9      # the item that was created is not sent again
        @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 9
        $svc2 = New-TargetService
        $svc2.ImportAnswers.Enqueue('')
        $o2 = New-Options @($pst); $o2.Concurrency = 1
        $run2 = Invoke-PstJob -Service $svc2 -Options $o2 -Journal (Join-Path $TestDrive 'answer2.db')
        $run2.Result.JobDone | Should -Be 9
        $run2.Result.ExitCode | Should -Be 0
        $svc2.RequestsOf('/importItem?') | Should -Be 9
    }
    It 'imports only the folders chosen (-IncludeFolder / -ExcludeFolder), creates the folders above them, and stops when none is chosen' {
        $pst = New-ImportPst 'choice'
        $svc = New-TargetService
        $o = New-Options @($pst); $o.IncludeFolders.Add('Inbox'); $o.ExcludeFolders.Add('Inbox/Projects')
        $run = Invoke-PstJob -Service $svc -Options $o -Journal (Join-Path $TestDrive 'choice.db')
        $run.Result.JobTotal | Should -Be 3
        $run.Result.JobComplete | Should -BeTrue
        $root = Get-FolderIdByName $svc 'old'
        Get-Subjects $svc (Get-FolderIdByName $svc 'Inbox' $root) | Should -Be @('in 1', 'in 2 with attachment', 'in 3 with embedded')
        @($svc.FoldersOf($script:Dst) | Where-Object Name -EQ 'Projects').Count | Should -Be 0
        @($svc.FoldersOf($script:Dst) | Where-Object Name -EQ 'Sent Items' | Where-Object ParentId -EQ $root).Count | Should -Be 0
        # a subfolder alone: its parent is created (empty) because the tree needs it
        $svc2 = New-TargetService
        $o2 = New-Options @($pst); $o2.IncludeFolders.Add('Inbox/Projects')
        $run2 = Invoke-PstJob -Service $svc2 -Options $o2 -Journal (Join-Path $TestDrive 'choice2.db')
        $run2.Result.JobTotal | Should -Be 2
        @($svc2.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 2
        $inbox2 = Get-FolderIdByName $svc2 'Inbox' (Get-FolderIdByName $svc2 'old')
        @($svc2.ItemsOf($script:Dst, $inbox2)).Count | Should -Be 0
        # nothing matches: the job stops before it creates anything
        $svc3 = New-TargetService
        $o3 = New-Options @($pst); $o3.IncludeFolders.Add('Inbx')
        $run3 = Invoke-PstJob -Service $svc3 -Options $o3 -Journal (Join-Path $TestDrive 'choice3.db')
        $run3.Error.Message | Should -BeLike '*No folder of the PST file(s) is chosen*'
        @($svc3.FoldersOf($script:Dst) | Where-Object Name -EQ 'Imported').Count | Should -Be 0
    }
    It 'reconciles an item whose answer was lost: found created (no duplicate)' {
        $pst = New-ImportPst 'lost-created'
        $svc = New-TargetService
        $svc.UploadDrops.Enqueue(-1); $svc.CreateOnLostAnswer = $true
        $o = New-Options @($pst); $o.Concurrency = 1
        $run = Invoke-PstJob -Service $svc -Options $o -Journal (Join-Path $TestDrive 'lost-created.db')
        $r = $run.Result
        $r.ReconciledCreated | Should -Be 1
        $r.JobComplete | Should -BeTrue
        $r.JobDone | Should -Be 9
        @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 9
        $svc.RequestsOf('/importItem?') | Should -Be 8   # the lost one (not logged: the connection dropped) was not sent again
    }
    It 'reconciles an item whose answer was lost: not created, so it is sent again' {
        $pst = New-ImportPst 'lost-absent'
        $svc = New-TargetService
        $svc.UploadDrops.Enqueue(-1)
        $o = New-Options @($pst); $o.Concurrency = 1
        $run = Invoke-PstJob -Service $svc -Options $o -Journal (Join-Path $TestDrive 'lost-absent.db')
        $r = $run.Result
        $r.ReconciledReimported | Should -Be 1
        $r.JobComplete | Should -BeTrue
        @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 9
    }
    It 'leaves an Unknown item Unknown (exit 3) without automatic reconciliation, and settles it at the next run' {
        $pst = New-ImportPst 'unknown'
        $svc = New-TargetService
        $svc.UploadDrops.Enqueue(-1); $svc.CreateOnLostAnswer = $true
        $o = New-Options @($pst); $o.Concurrency = 1; $o.AutoReconcile = $false
        $journal = Join-Path $TestDrive 'unknown.db'
        $first = Invoke-PstJob -Service $svc -Options $o -Journal $journal
        $first.Result.JobUnknown | Should -Be 1
        $first.Result.ExitCode | Should -Be 3
        $first.Result.NotImported[0].State | Should -Be 'Unknown'
        $o2 = New-Options @($pst); $o2.Concurrency = 1
        $second = Invoke-PstJob -Service $svc -Options $o2 -Journal $journal
        $second.Result.JobComplete | Should -BeTrue
        $second.Result.ReconciledCreated | Should -Be 1
        @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 9
    }
    It 'stops on Cancel and resumes later without duplicates' {
        $pst = New-ImportPst 'cancel'
        $svc = New-TargetService
        $svc.LatencyMs = 250
        $o = New-Options @($pst); $o.Concurrency = 1
        $journal = Join-Path $TestDrive 'cancel.db'
        $first = Invoke-PstJob -Service $svc -Options $o -Journal $journal -CancelAfterMs 1200
        $first.Result.Completed | Should -BeFalse
        $first.Result.JobDone | Should -BeLessThan 9
        $svc.LatencyMs = 5
        $second = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $second.Result.JobComplete | Should -BeTrue
        @($svc.ItemsOf($script:Dst) | Where-Object { $_.FolderId -ne 'F-INBOX' }).Count | Should -Be 9
    }
    It 'refuses a damaged file: nothing is imported and the error says what the checker found' {
        $bad = New-TestPst $TestDrive 'job-bad' -Faults @('dangling-nid')
        $svc = New-TargetService
        $run = Invoke-PstJob -Service $svc -Options (New-Options @($bad)) -Journal (Join-Path $TestDrive 'bad.db')
        $run.Error | Should -Not -BeNullOrEmpty
        $run.Error.Message | Should -BeLike '*damaged*'
        $run.Error.Message | Should -BeLike '*-Repair*'
        $run.Result.JobComplete | Should -BeFalse
        @($svc.ItemsOf($script:Dst)).Count | Should -Be 0
        $svc.RequestsOf('/importItem?') | Should -Be 0
        (Get-Item -LiteralPath $bad).Length | Should -BeGreaterThan 0
    }
    It 'with Repair, repairs into a temporary copy, leaves the file alone, imports the copy and reports it' {
        $bad = New-TestPst $TestDrive 'job-repair' -Faults @('dangling-nid')
        $hashBefore = (Get-FileHash -LiteralPath $bad).Hash
        $svc = New-TargetService
        $o = New-Options @($bad); $o.Repair = $true
        $journal = Join-Path $TestDrive 'repair.db'
        $run = Invoke-PstJob -Service $svc -Options $o -Journal $journal
        $run.Error | Should -BeNullOrEmpty
        $run.Result.Sources[0].Repaired | Should -BeTrue
        $run.Result.Sources[0].WorkPath | Should -BeLike '*.repaired.pst'
        (Test-Path -LiteralPath $run.Result.Sources[0].WorkPath) | Should -BeTrue
        $run.Result.Sources[0].Note | Should -BeLike '*repaired*'
        (Get-FileHash -LiteralPath $bad).Hash | Should -Be $hashBefore
        $run.Result.JobComplete | Should -BeTrue
        $run.Result.JobDone | Should -Be 56
        # the same copy is read again by the next run: nothing is imported twice
        $again = Invoke-PstJob -Service $svc -Options $o -Journal $journal
        $again.Result.Done | Should -Be 0
        $again.Result.JobTotal | Should -Be 56
        $svc.RequestsOf('/importItem?') | Should -Be 56
    }
    It 'reports an item that cannot be read from the PST and imports the others (check skipped)' {
        $good = New-TestPst $TestDrive 'job-readgood'
        $bad = New-DamagedPst $good $TestDrive 'flip-message'
        $svc = New-TargetService
        $o = New-Options @($bad); $o.Check = $false
        $run = Invoke-PstJob -Service $svc -Options $o -Journal (Join-Path $TestDrive 'unreadable.db')
        $r = $run.Result
        $r.Refused | Should -BeGreaterOrEqual 1
        $r.Done | Should -BeGreaterThan 40
        $r.NotImported[0].State | Should -Be 'Refused'
        $r.NotImported[0].Reason | Should -BeLike 'cannot be read from the PST*'
        $r.ExitCode | Should -Be 1   # finished, the unreadable items are refused for good
    }
    It 'refuses a job whose name belongs to another mailbox' {
        $pst = New-ImportPst 'othertarget'
        $svc = New-TargetService
        $other = "MBX:aaaaaaaa-0000-0000-0000-00000000000a@$($script:Tenant)"
        $null = $svc.AddMailbox($other)
        $journal = Join-Path $TestDrive 'other.db'
        $null = Invoke-PstJob -Service $svc -Options (New-Options @($pst)) -Journal $journal
        $o = New-Options @($pst); $o.TargetMailboxId = $other
        $run = Invoke-PstJob -Service $svc -Options $o -Journal $journal
        $run.Error.Message | Should -BeLike '*targets another mailbox*'
    }
}

Describe 'Import-MbtPst' {
    It 'shows the plan with -WhatIf and creates nothing, then imports through the function' {
        $pst = New-ImportPst 'cmdlet'
        $svc = New-TargetService
        $null = Set-MbtSession -Graph (New-FakeClient -Service $svc)
        $journal = Join-Path $TestDrive 'cmdlet.db'
        $dry = Import-MbtPst -Path $pst -MailboxId $script:Dst -TargetFolder 'Imported/cmd' -JournalPath $journal -WhatIf 6>$null
        $dry.Plan.Items | Should -Be 9
        $dry.Result | Should -BeNullOrEmpty
        @($svc.FoldersOf($script:Dst) | Where-Object Name -EQ 'cmd').Count | Should -Be 0
        $svc.RequestsOf('/importItem?') | Should -Be 0
        $r = Import-MbtPst -Path $pst -MailboxId $script:Dst -TargetFolder 'Imported/cmd' -JournalPath $journal -Concurrency 2 -ReconcileDelaySeconds 0 -Confirm:$false 6>$null
        $r.Result.JobComplete | Should -BeTrue
        $r.Result.ExitCode | Should -Be 0
        $r.Result.JobDone | Should -Be 9
        { Import-MbtPst -Path $pst -MailboxId $script:Dst -JournalPath $journal -Confirm:$false } | Should -Throw '*-TargetFolder is required*'
        Disconnect-MbtGraph
    }
}
