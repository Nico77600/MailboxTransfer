#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the one construction of the FastTransfer stream that importItem takes (Engine.FxImport.cs, FxImportStream),
    used by the import from PST and by the import of .eml files. One test per live-proven rule, the self-check that decodes the stream built
    (a silent loss of recipients or attachments can never be produced by us), and both features going through the helper. Nothing here
    reaches the network.

    Run:  Invoke-Pester -Path .\tests\FxImport.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    $script:P = [MailboxTransfer.FxProperty]
    $script:Fx = [MailboxTransfer.FxImportStream]

    function New-Rcpt([string]$Name, [switch]$RowIdLast, [switch]$NoRowId) {
        $r = [MailboxTransfer.FxRecipient]::new()
        if (-not $NoRowId -and -not $RowIdLast) { $r.Properties.Add($script:P::FromInt32(0x3000, 7)) }
        $r.Properties.Add($script:P::FromInt32(0x0C15, 1)); $r.Properties.Add($script:P::FromString(0x3001, $Name))
        if ($RowIdLast) { $r.Properties.Add($script:P::FromInt32(0x3000, 7)) }
        $r
    }
    # A message as a PST reader or a MIME converter gives it: no FXDelProp, recipients as they come, an embedded message holding a message
    # (depth 2), the empty PidTagAttachDataObject marker of a PST reader on the embedded attachments. Deterministic bytes (no random data).
    function New-Model {
        param([int]$Recipients = 2, [int]$FileAttachments = 1, [switch]$Embedded, [switch]$RowIdLast, [switch]$NoRowId)
        $m = [MailboxTransfer.FxMessage]::new()
        $m.Properties.Add($script:P::FromString(0x001A, 'IPM.Note')); $m.Properties.Add($script:P::FromString(0x0037, 'subject'))
        for ($i = 0; $i -lt $Recipients; $i++) { $m.Recipients.Add((New-Rcpt "rcpt$i" -RowIdLast:$RowIdLast -NoRowId:$NoRowId)) }
        for ($i = 0; $i -lt $FileAttachments; $i++) {
            $a = [MailboxTransfer.FxAttachment]::new()
            $a.Properties.Add($script:P::FromInt32(0x3705, 1)); $a.Properties.Add($script:P::FromBinary(0x3701, [byte[]]::new(300))); $a.Properties.Add($script:P::FromString(0x3707, "file$i.bin"))
            $m.Attachments.Add($a)
        }
        if ($Embedded) {
            $inner = [MailboxTransfer.FxMessage]::new()
            $inner.Properties.Add($script:P::FromString(0x001A, 'IPM.Note')); $inner.Properties.Add($script:P::FromString(0x0037, 'inner'))
            $inner.Recipients.Add((New-Rcpt 'inner rcpt' -RowIdLast:$RowIdLast -NoRowId:$NoRowId))
            $deep = [MailboxTransfer.FxMessage]::new()
            $deep.Properties.Add($script:P::FromString(0x001A, 'IPM.Note')); $deep.Properties.Add($script:P::FromString(0x0037, 'deep'))
            $da = [MailboxTransfer.FxAttachment]::new()
            $da.Properties.Add($script:P::FromInt32(0x3705, 5)); $da.Properties.Add($script:P::Create(0x3701, 0x000D, [byte[]]@())); $da.Embedded = $deep
            $inner.Attachments.Add($da)
            $ea = [MailboxTransfer.FxAttachment]::new()
            $ea.Properties.Add($script:P::FromInt32(0x3705, 5)); $ea.Properties.Add($script:P::Create(0x3701, 0x000D, [byte[]]@())); $ea.Embedded = $inner
            $m.Attachments.Add($ea)
        }
        $m
    }
    function Get-Bytes($source) { $ms = [IO.MemoryStream]::new(); $s = $source.OpenRead(); $s.CopyTo($ms); $s.Dispose(); , $ms.ToArray() }
    function Count-Pattern([byte[]]$Bytes, [byte[]]$Pattern) {
        $n = 0
        for ($i = 0; $i -le $Bytes.Length - $Pattern.Length; $i++) {
            $hit = $true
            for ($k = 0; $k -lt $Pattern.Length; $k++) { if ($Bytes[$i + $k] -ne $Pattern[$k]) { $hit = $false; break } }
            if ($hit) { $n++; $i += $Pattern.Length - 1 }
        }
        $n
    }
}

Describe 'the envelope' {
    It 'is one type-1 header record of 8 bytes (00000000 01000000), no type-4 record' {
        $env = $script:Fx::NewEnvelope()
        $env.Entries.Count | Should -Be 1
        $env.Entries[0].Type | Should -Be 1
        $env.Entries[0].Length | Should -Be 8
        ($env.Entries[0].Payload -join ',') | Should -Be '0,0,0,0,1,0,0,0'
    }
    It 'starts the stream built with the header record, the data records (type 2) follow and nothing else' {
        $src = $script:Fx::Build((New-Model -Embedded))
        $bytes = Get-Bytes $src
        $bytes.Length | Should -Be $src.Length
        ($bytes[0..15] -join ',') | Should -Be '1,0,0,0,8,0,0,0,0,0,0,0,1,0,0,0'
        [BitConverter]::ToUInt32($bytes, 16) | Should -Be 2         # the first data record
        $doc = [MailboxTransfer.FxCodec]::Decode($bytes)
        @($doc.Envelope.Entries | ForEach-Object Type | Select-Object -Unique) | Should -Be @(1, 2)
    }
}

Describe 'FXDelProp' {
    It 'is added before the recipients and before the attachments of the top-level message, also when it has neither' {
        $m = $script:Fx::Prepare((New-Model -Recipients 0 -FileAttachments 0))
        $m.DelPropRecipients.AsInt32() | Should -Be 0x0E12000D
        $m.DelPropAttachments.AsInt32() | Should -Be 0x0E13000D
    }
    It 'is written twice in the stream of a message with an embedded message, never inside the embedded message' {
        $bytes = Get-Bytes ($script:Fx::Build((New-Model -Embedded)))
        Count-Pattern $bytes ([byte[]](0x03, 0x00, 0x16, 0x40)) | Should -Be 2     # PtypInteger32 0x4016 = FXDelProp, LE
        $doc = [MailboxTransfer.FxCodec]::Decode($bytes)
        $doc.Message.DelPropRecipients | Should -Not -BeNullOrEmpty
        $doc.Message.DelPropAttachments | Should -Not -BeNullOrEmpty
        $inner = ($doc.Message.Attachments | Where-Object Embedded).Embedded
        $inner.DelPropRecipients | Should -BeNullOrEmpty
        $inner.DelPropAttachments | Should -BeNullOrEmpty
        $inner.Attachments[0].Embedded.DelPropAttachments | Should -BeNullOrEmpty
    }
    It 'is taken away from an embedded message that came with it (as a stream of another source could have)' {
        $m = New-Model -Embedded
        $inner = ($m.Attachments | Where-Object Embedded).Embedded
        $inner.DelPropRecipients = $script:P::Create(0x4016, 0x0003, [BitConverter]::GetBytes([uint32]0x0E12000D))
        $inner.DelPropAttachments = $script:P::Create(0x4016, 0x0003, [BitConverter]::GetBytes([uint32]0x0E13000D))
        [void]$script:Fx::Prepare($m)
        $inner.DelPropRecipients | Should -BeNullOrEmpty
        $inner.DelPropAttachments | Should -BeNullOrEmpty
    }
    It 'is kept when the message has it already (the same object, no second marker)' {
        $m = $script:Fx::Prepare((New-Model))
        $first = $m.DelPropRecipients
        [void]$script:Fx::Prepare($m)
        [object]::ReferenceEquals($m.DelPropRecipients, $first) | Should -BeTrue
        $m.CountProperties() | Should -Be ((New-Model).CountProperties() + 2)    # the 2 FXDelProp only: every recipient had its PidTagRowid
    }
}

Describe 'PidTagRowid' {
    It 'is moved to the first place of a recipient that has it elsewhere (value kept)' {
        $m = $script:Fx::Prepare((New-Model -RowIdLast))
        foreach ($r in $m.Recipients) { $r.Properties[0].Id | Should -Be 0x3000; $r.Properties[0].AsInt32() | Should -Be 7 }
        @($m.Recipients[0].Properties | Where-Object Id -EQ 0x3000).Count | Should -Be 1
    }
    It 'is added, numbered by the position of the recipient, to a recipient that has none' {
        $m = $script:Fx::Prepare((New-Model -Recipients 3 -NoRowId))
        for ($i = 0; $i -lt 3; $i++) { $m.Recipients[$i].Properties[0].Id | Should -Be 0x3000; $m.Recipients[$i].Properties[0].AsInt32() | Should -Be $i }
    }
    It 'applies to the recipients of embedded messages too' {
        $m = $script:Fx::Prepare((New-Model -Embedded -NoRowId))
        $inner = ($m.Attachments | Where-Object Embedded).Embedded
        $inner.Recipients[0].Properties[0].Id | Should -Be 0x3000
    }
}

Describe 'embedded messages' {
    It 'start with PidTagMid (0x674A), an embedded message of an embedded message too, each with its own value' {
        $m = $script:Fx::Prepare((New-Model -Embedded))
        $inner = ($m.Attachments | Where-Object Embedded).Embedded
        $deep = ($inner.Attachments | Where-Object Embedded).Embedded
        $inner.Properties[0].Id | Should -Be 0x674A
        $deep.Properties[0].Id | Should -Be 0x674A
        $inner.Properties[0].AsInt64() | Should -Not -Be $deep.Properties[0].AsInt64()
        $m.Properties[0].Id | Should -Not -Be 0x674A       # the top-level message gets none
    }
    It 'get PidTagMid in the first place when they carry it elsewhere (not duplicated)' {
        $m = New-Model -Embedded
        $inner = ($m.Attachments | Where-Object Embedded).Embedded
        $inner.Properties.Add($script:P::FromInt64(0x674A, 99))
        [void]$script:Fx::Prepare($m)
        $inner.Properties[0].Id | Should -Be 0x674A
        $inner.Properties[0].AsInt64() | Should -Be 99
        @($inner.Properties | Where-Object Id -EQ 0x674A).Count | Should -Be 1
    }
    It 'have no empty PidTagAttachDataObject (0x3701) marker on their attachment, a file attachment keeps its data' {
        $m = $script:Fx::Prepare((New-Model -Embedded))
        $ea = $m.Attachments | Where-Object Embedded
        @($ea.Properties | Where-Object Id -EQ 0x3701).Count | Should -Be 0
        $file = $m.Attachments | Where-Object { -not $_.Embedded } | Select-Object -First 1
        @($file.Properties | Where-Object Id -EQ 0x3701).Count | Should -Be 1
    }
}

Describe 'the self-check of the stream built' {
    It 'gives the same recipients, attachments and embedded messages after decoding as the message has (embedded included)' {
        $m = New-Model -Recipients 3 -FileAttachments 2 -Embedded
        $src = $script:Fx::Build($m)
        $want = [MailboxTransfer.FxImportCounts]::Of($m)
        $want.ToString() | Should -Be '4 recipients, 4 attachments, 2 embedded messages'
        $doc = [MailboxTransfer.FxCodec]::Decode((Get-Bytes $src))
        [MailboxTransfer.FxImportCounts]::Of($doc.Message).ToString() | Should -Be $want.ToString()
    }
    It 'passes a message without recipients, attachments or embedded messages' {
        { [void]$script:Fx::Build((New-Model -Recipients 0 -FileAttachments 0)) } | Should -Not -Throw
    }
    It 'refuses a stream without FXDelProp (the service would accept it and drop the recipients and attachments)' {
        $m = New-Model                                                   # not prepared
        [void]$script:Fx::Prepare($m); $m.DelPropRecipients = $null; $m.DelPropAttachments = $null
        $src = [MailboxTransfer.FxMessageSource]::new($m, $script:Fx::NewEnvelope())
        { $script:Fx::Verify($src, $m) } | Should -Throw '*no FXDelProp*'
    }
    It 'refuses a stream that holds fewer recipients than the message (a silent drop)' {
        $m = $script:Fx::Prepare((New-Model))
        $src = [MailboxTransfer.FxMessageSource]::new($m, $script:Fx::NewEnvelope())
        $bigger = $script:Fx::Prepare((New-Model -Recipients 3))
        { $script:Fx::Verify($src, $bigger) } | Should -Throw '*2 recipients*3 recipients*'
    }
    It 'refuses a stream that holds fewer attachments, or fewer embedded messages, than the message' {
        $m = $script:Fx::Prepare((New-Model -Embedded))
        $src = [MailboxTransfer.FxMessageSource]::new($m, $script:Fx::NewEnvelope())
        { $script:Fx::Verify($src, (New-Model -FileAttachments 4 -Embedded)) } | Should -Throw '*3 attachments*where the message has*6 attachments*'
        { $script:Fx::Verify($src, $script:Fx::Prepare((New-Model))) } | Should -Throw '*2 embedded messages*where the message has*0 embedded messages*'     # the model has none, the stream has two
    }
    It 'refuses a stream without the header record of the envelope' {
        $m = $script:Fx::Prepare((New-Model))
        $bare = [MailboxTransfer.FxMessageSource]::new($m)
        { $script:Fx::Verify($bare, $m) } | Should -Throw '*type-1 header record*'
    }
    It 'refuses a recipient that does not start with PidTagRowid, an embedded message without PidTagMid or with FXDelProp, a marker before StartEmbed' {
        $src = { param($m) [MailboxTransfer.FxMessageSource]::new($m, $script:Fx::NewEnvelope()) }
        $m = $script:Fx::Prepare((New-Model -Embedded)); $m.Recipients[0].Properties.Reverse()
        { $script:Fx::Verify((& $src $m), $m) } | Should -Throw '*PidTagRowid*'
        $m = $script:Fx::Prepare((New-Model -Embedded)); ($m.Attachments | Where-Object Embedded).Embedded.Properties.RemoveAt(0)
        { $script:Fx::Verify((& $src $m), $m) } | Should -Throw '*PidTagMid*'
        $m = $script:Fx::Prepare((New-Model -Embedded)); ($m.Attachments | Where-Object Embedded).Embedded.DelPropRecipients = $script:P::Create(0x4016, 0x0003, [BitConverter]::GetBytes([uint32]0x0E12000D))
        { $script:Fx::Verify((& $src $m), $m) } | Should -Throw '*FXDelProp*'
        $m = $script:Fx::Prepare((New-Model -Embedded)); ($m.Attachments | Where-Object Embedded).Properties.Add($script:P::Create(0x3701, 0x000D, [byte[]]@()))
        { $script:Fx::Verify((& $src $m), $m) } | Should -Throw '*PidTagAttachDataObject*'
    }
    It 'Build reports a property that cannot be expressed before anything else (ArgumentException)' {
        $m = New-Model
        $m.Properties.Add($script:P::Create(0x6000, 0x0003, [byte[]]@(1, 2)))     # a 4-byte type with 2 bytes
        { [void]$script:Fx::Build($m) } | Should -Throw
    }
}

Describe 'both features build their stream with the helper' {
    It 'the EML conversion gives the same structure as the PST path: DelProp at the top, Rowid first, Mid in the embedded message, no marker, header record' {
        $eml = @(
            'From: Alice <alice@contoso.example>', 'To: Bob <bob@contoso.example>', 'Cc: carol@contoso.example', 'Subject: outer', 'MIME-Version: 1.0',
            'Content-Type: multipart/mixed; boundary="B"', '',
            '--B', 'Content-Type: text/plain; charset=utf-8', '', 'hello', '--B',
            'Content-Type: application/octet-stream; name="a.bin"', 'Content-Disposition: attachment; filename="a.bin"', 'Content-Transfer-Encoding: base64', '', 'AAECAwQF', '--B',
            'Content-Type: message/rfc822', 'Content-Disposition: attachment; filename="inner.eml"', '',
            'From: Dan <dan@contoso.example>', 'To: Eve <eve@contoso.example>', 'Subject: inner', 'Content-Type: text/plain', '', 'inner body', '--B--', '') -join "`r`n"
        $notes = [System.Collections.Generic.List[string]]::new()
        $msg = $null
        $src = [MailboxTransfer.EmlImporter]::ConvertOwn([Text.Encoding]::ASCII.GetBytes($eml), [MailboxTransfer.EmlImportOptions]::new(), $notes, [ref]$msg)
        $bytes = Get-Bytes $src
        ($bytes[0..7] -join ',') | Should -Be '1,0,0,0,8,0,0,0'
        $doc = [MailboxTransfer.FxCodec]::Decode($bytes)
        $doc.Ok | Should -BeTrue
        $doc.Message.Recipients.Count | Should -Be 2
        foreach ($r in $doc.Message.Recipients) { $r.Properties[0].Id | Should -Be 0x3000 }
        $doc.Message.DelPropRecipients | Should -Not -BeNullOrEmpty
        $doc.Message.DelPropAttachments | Should -Not -BeNullOrEmpty
        $inner = ($doc.Message.Attachments | Where-Object Embedded).Embedded
        $inner.Properties[0].Id | Should -Be 0x674A
        $inner.DelPropRecipients | Should -BeNullOrEmpty
        $inner.Recipients[0].Properties[0].Id | Should -Be 0x3000
        [MailboxTransfer.FxImportCounts]::Of($doc.Message).ToString() | Should -Be '3 recipients, 2 attachments, 1 embedded messages'
        [MailboxTransfer.FxDiff]::TotalsText($doc.Message) | Should -Be '3 recipients, 2 attachments, 1 embedded messages'
    }
}
