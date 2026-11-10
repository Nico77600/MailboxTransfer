#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of tools\New-LabFxCorpus.ps1 (lab corpus of real Exchange items for the FastTransfer codec).
    Offline only: the plan and the shape of the Graph payloads (tools\LabFxCorpus.Lib.ps1), -WhatIf, and the whole run against a local
    stand-in of Graph (tests\support\FakeGraphServer.cs, 127.0.0.1) with -AccessToken / -GraphBaseUrl. Nothing here reaches the cloud.

    Run:  Invoke-Pester -Path .\tests\LabFxCorpus.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:Tools = Join-Path (Split-Path $PSScriptRoot -Parent) 'tools'
    . (Join-Path $script:Tools 'LabFxCorpus.Lib.ps1')
    $script:Plan = Get-FxcPlan -Mailbox 'lab.user@contoso.example'
    function Step([string]$Key) { $script:Plan | Where-Object Key -eq $Key }
    function Get-Json($Value) { $Value | ConvertTo-Json -Depth 20 -Compress | ConvertFrom-Json -Depth 30 }
    if (-not ('MailboxTransfer.Tests.FakeGraphServer' -as [type])) {
        Add-Type -LiteralPath (Join-Path $PSScriptRoot 'support\FakeGraphServer.cs') -ReferencedAssemblies 'System.Net.HttpListener', 'System.Text.Json', 'System.Linq', 'System.Collections', 'System.Net.Primitives', 'System.Threading', 'System.Threading.Thread', 'System.Runtime', 'System.Net.Sockets', 'System.Private.Uri', 'System.Memory', 'System.Text.Encodings.Web', 'System.Collections.Concurrent', 'System.Collections.Specialized', 'System.Net.Primitives', 'netstandard' -IgnoreWarnings -WarningAction SilentlyContinue
    }
}

Describe 'The plan of the corpus' {
    It 'has 23 items with unique keys, a marker at the start of every subject and a payload' {
        $script:Plan.Count | Should -Be 23
        ($script:Plan.Key | Select-Object -Unique).Count | Should -Be 23
        foreach ($s in $script:Plan) {
            $s.Subject.StartsWith($s.Marker) | Should -BeTrue -Because $s.Key
            Get-FxcMarkerFromSubject $s.Subject | Should -Be $s.Key
            ($s.Variants.Count -gt 0 -or $s.Mime) | Should -BeTrue -Because $s.Key
        }
        ($script:Plan | Where-Object Kind -eq 'Mail').Count | Should -Be 15
        ($script:Plan | Where-Object Kind -eq 'Event').Count | Should -Be 5
        ($script:Plan | Where-Object Kind -eq 'Contact').Count | Should -Be 3
    }

    It 'has JSON payloads that survive serialisation (nothing cut by the depth) and stay small' {
        $total = 0
        foreach ($s in $script:Plan) {
            foreach ($v in $s.Variants) {
                $json = $v | ConvertTo-Json -Depth 20 -Compress -WarningAction Stop
                $json | Should -Not -Match 'System\.Collections|System\.Object'
                $total += $json.Length
            }
            foreach ($a in $s.Attachments) { foreach ($v in $a.Variants) { ($v | ConvertTo-Json -Depth 20 -Compress -WarningAction Stop) | Should -Not -Match 'System\.Collections|System\.Object' } }
        }
        $total | Should -BeLessThan 200000
    }

    It 'fxc-01: HTML body, inline image attachment with the contentId of the body, external link' {
        $s = Step 'fxc-01'
        $m = Get-Json $s.Variants[0]
        $a = Get-Json $s.Attachments[0].Variants[0]
        $a.'@odata.type' | Should -Be '#microsoft.graph.fileAttachment'
        $a.isInline | Should -BeTrue
        $m.body.contentType | Should -Be 'html'
        $m.body.content | Should -Match ('cid:' + [regex]::Escape($a.contentId))
        $m.body.content | Should -Match 'https://www.contoso.example'
        [Convert]::FromBase64String($a.contentBytes)[0..3] | Should -Be @(0x89, 0x50, 0x4E, 0x47)
    }

    It 'fxc-02: 5 To, 3 Cc, 2 Bcc with accented, CJK and emoji display names' {
        $m = Get-Json (Step 'fxc-02').Variants[0]
        @($m.toRecipients).Count | Should -Be 5
        @($m.ccRecipients).Count | Should -Be 3
        @($m.bccRecipients).Count | Should -Be 2
        $names = @($m.toRecipients + $m.ccRecipients + $m.bccRecipients).emailAddress.name -join ''
        $names | Should -Match "[$([char]0xE9)$([char]0xFC)]"
        $names | Should -Match ([char]0x65E5)
        $names | Should -Match ([char]::ConvertFromUtf32(0x1F600))
    }

    It 'fxc-03 and fxc-04: embedded messages, with a nested file and with a second level (and a weaker variant each)' {
        $a3 = (Step 'fxc-03').Attachments[0].Variants
        $a3.Count | Should -Be 2
        (Get-Json $a3[0]).item.'@odata.type' | Should -Be 'microsoft.graph.message'
        @((Get-Json $a3[0]).item.attachments)[0].'@odata.type' | Should -Be '#microsoft.graph.fileAttachment'
        (Get-Json $a3[1]).item.PSObject.Properties.Name | Should -Not -Contain 'attachments'
        $a4 = (Step 'fxc-04').Attachments[0].Variants
        $level1 = (Get-Json $a4[0]).item
        @($level1.attachments)[0].'@odata.type' | Should -Be '#microsoft.graph.itemAttachment'
        @($level1.attachments)[0].item.subject | Should -Be 'Level 2 message'
        (Get-Json $a4[1]).item.PSObject.Properties.Name | Should -Not -Contain 'attachments'
    }

    It 'fxc-05: an event and a contact as item attachments' {
        $att = (Step 'fxc-05').Attachments
        (Get-Json $att[0].Variants[0]).item.'@odata.type' | Should -Be 'microsoft.graph.event'
        (Get-Json $att[1].Variants[0]).item.'@odata.type' | Should -Be 'microsoft.graph.contact'
    }

    It 'fxc-06: real file formats (PNG, PDF, ZIP, text, 64 KiB binary) and a Unicode file name' {
        $att = @((Step 'fxc-06').Attachments | ForEach-Object { Get-Json $_.Variants[0] })
        $att.Count | Should -Be 6
        $bytes = @{}; foreach ($a in $att) { $bytes[$a.name] = [Convert]::FromBase64String($a.contentBytes) }
        $bytes['pixel.png'][0..7] | Should -Be @(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)
        [Text.Encoding]::ASCII.GetString($bytes['doc.pdf'], 0, 8) | Should -Be '%PDF-1.1'
        $zip = [IO.Compression.ZipArchive]::new([IO.MemoryStream]::new($bytes['bundle.zip']))
        $zip.Entries.Count | Should -Be 2
        $zip.Dispose()
        $bytes['random.bin'].Length | Should -Be 65536
        @($att.name | Where-Object { $_ -cmatch '[^\x20-\x7E]' }).Count | Should -Be 1
    }

    It 'fxc-07: categories, importance, flag with dates, read, receipts; sensitivity is the MAPI property, not a message property' {
        $s = Step 'fxc-07'
        $m = Get-Json $s.Variants[0]
        @($m.categories).Count | Should -Be 2
        $m.importance | Should -Be 'high'
        $m.isRead | Should -BeTrue
        $m.isReadReceiptRequested | Should -BeTrue
        $m.isDeliveryReceiptRequested | Should -BeTrue
        $m.flag.flagStatus | Should -BeIn 'notFlagged', 'complete', 'flagged'
        $m.flag.flagStatus | Should -Be 'flagged'
        $m.flag.startDateTime.dateTime | Should -Not -BeNullOrEmpty
        $m.flag.startDateTime.timeZone | Should -Not -BeNullOrEmpty
        $m.flag.dueDateTime.dateTime | Should -Not -BeNullOrEmpty
        $m.flag.dueDateTime.timeZone | Should -Not -BeNullOrEmpty
        $s.Variants.Count | Should -BeGreaterThan 1
        # sensitivity: PidTagSensitivity 0x0036, Integer, 3 = confidential, set by a PATCH of singleValueExtendedProperties
        $e = @($s.Ext | Where-Object Id -eq 'Integer 0x0036')
        $e.Count | Should -Be 1
        $e[0].Variants[0] | Should -Be '3'
        $body = Get-Json (Get-FxcExtBody -Ext $e[0] -Variant 0)
        $body.singleValueExtendedProperties[0].id | Should -Be 'Integer 0x0036'
        $body.singleValueExtendedProperties[0].value | Should -Be '3'
    }

    It 'message payloads only use properties of the Graph v1.0 message resource (no sensitivity, which only events have)' {
        $allowed = 'subject', 'body', 'toRecipients', 'ccRecipients', 'bccRecipients', 'from', 'sender', 'replyTo', 'categories', 'importance', 'isRead',
            'isReadReceiptRequested', 'isDeliveryReceiptRequested', 'flag', 'internetMessageHeaders', 'singleValueExtendedProperties', 'multiValueExtendedProperties'
        foreach ($s in $script:Plan | Where-Object { $_.Kind -eq 'Mail' }) {
            foreach ($v in $s.Variants) {
                foreach ($k in $v.Keys) { $k | Should -BeIn $allowed -Because "$($s.Key): '$k' is not a property of the message resource" }
                $v.Contains('sensitivity') | Should -BeFalse -Because $s.Key
            }
        }
    }
    It 'fxc-08 and fxc-09: extended properties of every type, by name and by id, each with its own PATCH body' {
        $e8 = (Step 'fxc-08').Ext
        $types8 = $e8 | ForEach-Object { ($_.Id -split ' ')[0] }
        foreach ($t in 'Integer', 'Long', 'Double', 'Currency', 'Boolean', 'SystemTime', 'Binary', 'String', 'CLSID') { $types8 | Should -Contain $t }
        ($e8.Id | Where-Object { $_ -match ' Name ' }).Count | Should -BeGreaterThan 5
        ($e8.Id | Where-Object { $_ -match ' Id 0x8' }).Count | Should -BeGreaterOrEqual 2
        $body = Get-Json (Get-FxcExtBody -Ext ($e8 | Where-Object Id -like 'Currency*') -Variant 1)
        $body.singleValueExtendedProperties[0].value | Should -Be '12345600'
        $e9 = (Step 'fxc-09').Ext
        $types9 = $e9 | ForEach-Object { ($_.Id -split ' ')[0] }
        foreach ($t in 'StringArray', 'IntegerArray', 'LongArray', 'DoubleArray', 'SystemTimeArray', 'BinaryArray', 'CLSIDArray', 'CurrencyArray') { $types9 | Should -Contain $t }
        $mv = Get-Json (Get-FxcExtBody -Ext ($e9 | Where-Object Id -like 'StringArray*') -Variant 0)
        @($mv.multiValueExtendedProperties[0].value).Count | Should -Be 4
        $cur = ($e9 | Where-Object Id -like 'CurrencyArray*')
        $cur.Variants.Count | Should -Be 2
        @((Get-Json (Get-FxcExtBody -Ext $cur -Variant 1)).multiValueExtendedProperties[0].value) | Should -Be @('12500', '25000')
    }

    It 'fxc-10 and fxc-11: custom X- headers, long lines, a subject of at least 255 characters' {
        $s10 = Step 'fxc-10'
        $m = Get-Json $s10.Variants[0]
        @($m.internetMessageHeaders).Count | Should -BeLessOrEqual 5
        $m.internetMessageHeaders.name | ForEach-Object { $_ | Should -BeLike 'x-*' }
        $s10.Variants[1].Contains('internetMessageHeaders') | Should -BeFalse
        ($m.body.content -split "`r`n" | Measure-Object -Property Length -Maximum).Maximum | Should -BeGreaterOrEqual 4000
        $m.body.content | Should -Match ([char]::ConvertFromUtf32(0x1F680))
        $s11 = Step 'fxc-11'
        $s11.Variants[0].subject.Length | Should -Be 300
        $s11.Variants[1].subject.Length | Should -Be 255
        Get-FxcMarkerFromSubject $s11.Variants[1].subject | Should -Be 'fxc-11'
    }

    It 'fxc-12 to fxc-14: empty body, attachment only, received-style message with message flags' {
        (Step 'fxc-12').Variants[0].body.content | Should -Be ''
        (Step 'fxc-13').Variants[0].body.content | Should -Be ''
        (Step 'fxc-13').Attachments.Count | Should -Be 1
        $v = (Step 'fxc-14').Variants
        ($v[0].singleValueExtendedProperties.id) | Should -Contain 'Integer 0x0E07'
        $v[2].Contains('from') | Should -BeFalse
    }

    It 'fxc-15: a MIME message with an rfc822 attachment that holds another rfc822 attachment, encoded words and Bcc' {
        $mime = (Step 'fxc-15').Mime
        ([regex]::Matches($mime, '(?m)^Content-Type: message/rfc822')).Count | Should -Be 2
        $mime | Should -Match '(?m)^Bcc: '
        $mime | Should -Match '(?m)^Subject: \[fxc-15\] '
        $mime | Should -Match '=\?utf-8\?B\?'
        $mime | Should -Match "filename\*=utf-8''"
        $mime | Should -Not -Match "[^\x00-\x7F]"
        ($mime -split "`r`n").Where({ $_.Length -gt 998 }).Count | Should -Be 0
    }

    It 'events: attendees of the three types, all-day, weekly recurrence of 6, attachments, private' {
        $e = Get-Json (Step 'fxc-20').Variants[0]
        @($e.attendees.type | Sort-Object) | Should -Be @('optional', 'required', 'resource')
        $e.location.displayName | Should -Not -BeNullOrEmpty
        $e.body.contentType | Should -Be 'html'
        (Get-Json (Step 'fxc-21').Variants[0]).isAllDay | Should -BeTrue
        $r = Get-Json (Step 'fxc-22').Variants[0]
        $r.recurrence.pattern.type | Should -Be 'weekly'
        $r.recurrence.range.type | Should -BeIn 'endDate', 'noEnd', 'numbered'
        $r.recurrence.range.type | Should -Be 'numbered'
        $r.recurrence.range.numberOfOccurrences | Should -Be 6
        (Step 'fxc-22').Variants[0].recurrence.range.numberOfOccurrences | Should -BeOfType [int]
        $r.recurrence.range.startDate | Should -Match '^\d{4}-\d{2}-\d{2}$'
        (Step 'fxc-22').Special | Should -Be 'RecurringExceptions'
        (Step 'fxc-23').Attachments.Count | Should -Be 2
        $p = Get-Json (Step 'fxc-24').Variants[0]
        $p.sensitivity | Should -Be 'private'
        $p.isOnlineMeeting | Should -BeFalse
        $p.PSObject.Properties.Name | Should -Not -Contain 'attendees'
    }

    It 'contacts: a full one (many fields), a minimal one, one with a JPEG photo when System.Drawing exists' {
        $c = Get-Json (Step 'fxc-30').Variants[0]
        @($c.emailAddresses).Count | Should -Be 3
        @($c.businessPhones).Count | Should -Be 2
        @($c.homePhones).Count | Should -Be 2
        $c.homeAddress.city | Should -Not -BeNullOrEmpty
        $c.businessAddress.city | Should -Not -BeNullOrEmpty
        $c.otherAddress.city | Should -Not -BeNullOrEmpty
        $c.birthday | Should -Be '1984-02-29'
        @($c.categories).Count | Should -Be 2
        $c.givenName | Should -Match ([char]0xE9)
        (Get-Json (Step 'fxc-31').Variants[0]).PSObject.Properties.Name.Count | Should -BeLessOrEqual 3
        $photo = (Step 'fxc-32').Photo
        if ($photo) { $photo[0..2] | Should -Be @(0xFF, 0xD8, 0xFF) }
    }

    It 'plan summary lists one row per step' {
        @(Get-FxcPlanSummary -Plan $script:Plan).Count | Should -Be 23
    }
}

Describe 'The end of the run' {
    It 'records the id and the name of a folder, and does not guess a path Graph did not give' {
        $r = New-FxcFolderReport -Kind Event -Id 'AAMkCAL=' -Name 'MBT-FXICS-Corpus'
        $r.GraphId | Should -Be 'AAMkCAL='
        $r.Name | Should -Be 'MBT-FXICS-Corpus'
        $r.PathResolved | Should -BeFalse
        $r.ExportPath | Should -BeNullOrEmpty
        $r.HowToFindPath | Should -Match 'Invoke-Phase0'
        $r.HowToFindPath | Should -Match "'MBT-FXICS-Corpus'"
        $r.HowToFindPath | Should -Match 'IPF.Appointment'
        $ok = New-FxcFolderReport -Kind Mail -Id 'AAMkMAIL=' -Name 'Mail' -ExportPath 'MBT-FXICS-Corpus/Mail'
        $ok.PathResolved | Should -BeTrue
        $ok.ExportPath | Should -Be 'MBT-FXICS-Corpus/Mail'
        $ok.HowToFindPath | Should -BeNullOrEmpty
    }

    It 'builds the manifest without modifying what it reads (failed items, unresolved path), also with no item and no folder' {
        $results = @(
            [pscustomobject]@{ Key = 'fxc-01'; Kind = 'Mail'; Status = 'Created'; Id = 'm1'; Title = 'one'; Notes = [Collections.Generic.List[string]]::new() }
            [pscustomobject]@{ Key = 'fxc-22'; Kind = 'Event'; Status = 'Failed'; Id = $null; Title = 'two'; Notes = @('HTTP 400') })
        $folders = [ordered]@{ Mail = 'fm'; Event = 'fe' }
        $paths = [ordered]@{ Mail = 'MBT-FXICS-Corpus/Mail'; Event = $null }
        $names = [ordered]@{ Mail = 'Mail'; Event = 'MBT-FXICS-Corpus' }
        $m = New-FxcManifest -User 'u@contoso.example' -Root 'MBT-FXICS-Corpus' -Results $results -Folders $folders -Paths $paths -FolderNames $names
        $folders.Count | Should -Be 2
        $paths.Count | Should -Be 2
        @($m.Items).Count | Should -Be 2
        ($m.Items | Where-Object Key -eq 'fxc-22').Status | Should -Be 'Failed'
        $m.ExportFolders.Mail | Should -Be 'MBT-FXICS-Corpus/Mail'
        $m.ExportFolders.Event | Should -BeNullOrEmpty
        $m.Folders.Event.GraphId | Should -Be 'fe'
        $m.Folders.Event.PathResolved | Should -BeFalse
        { $m | ConvertTo-Json -Depth 8 } | Should -Not -Throw
        $empty = New-FxcManifest -User 'u@contoso.example' -Root 'R' -RunError 'boom'
        $empty.Error | Should -Be 'boom'
        @($empty.Items).Count | Should -Be 0
    }
}
Describe 'New-LabFxCorpus.ps1 -WhatIf' {
    It 'prints the plan, signs in nowhere and sends nothing' {
        $out = & (Join-Path $script:Tools 'New-LabFxCorpus.ps1') -WhatIf *>&1 | Out-String -Width 250
        $out | Should -Match 'fxc-01'
        $out | Should -Match 'fxc-32'
        $out | Should -Match 'MBT-FXICS-Corpus\\Mail'
        $out | Should -Match 'nothing was sent'
        $out | Should -Not -Match 'DEVICE CODE'
    }

    It 'limits the plan with -Only (keys, numbers, kinds)' {
        $out = & (Join-Path $script:Tools 'New-LabFxCorpus.ps1') -WhatIf -Only 7, 'fxc-22', 'Contact' *>&1 | Out-String -Width 250
        foreach ($k in 'fxc-07', 'fxc-22', 'fxc-30', 'fxc-31', 'fxc-32') { $out | Should -Match $k }
        $out | Should -Not -Match 'fxc-01'
        { & (Join-Path $script:Tools 'New-LabFxCorpus.ps1') -WhatIf -Only 'nothing' *>&1 } | Should -Throw '*matches no item*'
    }

    It 'refuses a Graph address that is neither Microsoft nor local' {
        { & (Join-Path $script:Tools 'New-LabFxCorpus.ps1') -WhatIf -GraphBaseUrl 'https://evil.example/v1.0' *>&1 } | Should -Throw '*GraphBaseUrl*'
    }
}

Describe 'New-LabFxCorpus.ps1 against a local stand-in of Graph' {
    BeforeAll {
        $script:Script = Join-Path $script:Tools 'New-LabFxCorpus.ps1'
        $script:Work = Join-Path ([IO.Path]::GetTempPath()) ('fxcorpus-' + [guid]::NewGuid().ToString('N'))
        $null = New-Item -ItemType Directory -Path $script:Work
        function Invoke-Corpus($Server, [string]$Name, [string[]]$More = @()) {
            $manifest = Join-Path $script:Work "$Name.json"
            $out = & pwsh -NoProfile -File $script:Script -AccessToken 'fake-token' -GraphBaseUrl $Server.Url -OutputPath $manifest -UserPrincipalName 'lab.user@contoso.example' @More 2>&1 | Out-String
            [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $out; Manifest = $(if (Test-Path $manifest) { Get-Content $manifest -Raw | ConvertFrom-Json -Depth 20 }) }
        }
    }
    AfterAll { Remove-Item -LiteralPath $script:Work -Recurse -Force -ErrorAction SilentlyContinue }

    It 'creates the folders and the 23 items, then does nothing on a second run' {
        $server = [MailboxTransfer.Tests.FakeGraphServer]::new()
        try {
            $r = Invoke-Corpus $server 'run1'
            $r.ExitCode | Should -Be 0 -Because $r.Output
            $r.Manifest.Items.Count | Should -Be 23
            @($r.Manifest.Items | Where-Object Status -ne 'Created') | Should -BeNullOrEmpty -Because $r.Output
            @($server.Requested('POST me/mailFolders/fld')).Count | Should -BeGreaterThan 0
            @($server.Requested('POST me/mailFolders/fld2/messages')).Count | Should -Be 15
            @($server.Requested('POST me/calendars/fld')).Count | Should -Be 5
            @($server.Requested('POST me/contactFolders/fld')).Count | Should -Be 3
            @($server.Requested('POST me/messages/')).Count | Should -BeGreaterThan 10
            @($server.Requested('PATCH me/messages/')).Count | Should -Be (15 + 10 + 3 + 1)
            @($server.Requested('PUT me/contactFolders/')).Count | Should -Be 1
            @($server.Requested('DELETE me/events/inst-')).Count | Should -Be 1
            @($server.Requested('PATCH me/events/inst-')).Count | Should -Be 1
            $r.Output | Should -Match 'Mail +MBT-FXICS-Corpus/Mail'
            $r.Manifest.ExportFolders.Mail | Should -Be 'MBT-FXICS-Corpus/Mail'
            $r.Manifest.ExportFolders.Contact | Should -Be 'Contacts/MBT-FXICS-Corpus'

            $posts = @($server.Requested('POST ')).Count
            $r2 = Invoke-Corpus $server 'run2'
            $r2.ExitCode | Should -Be 0 -Because $r2.Output
            @($r2.Manifest.Items | Where-Object Status -ne 'Skipped') | Should -BeNullOrEmpty -Because $r2.Output
            @($server.Requested('POST ')).Count | Should -Be $posts -Because 'a second run creates nothing'
            @($server.Requested('DELETE me/events/inst-')).Count | Should -Be 1
        } finally { $server.Dispose() }
    }

    It 'writes the manifest, with the id and the name of the calendar, when Graph cannot resolve its path (no crash at the end)' {
        $server = [MailboxTransfer.Tests.FakeGraphServer]::new()
        $server.CalendarsAreNotMailFolders = $true
        try {
            $r = Invoke-Corpus $server 'run5'
            $r.Output | Should -Not -Match 'Collection was modified'
            $r.ExitCode | Should -Be 0 -Because $r.Output
            $r.Manifest | Should -Not -BeNullOrEmpty -Because $r.Output
            $r.Manifest.Items.Count | Should -Be 23
            $r.Manifest.ExportFolders.Event | Should -BeNullOrEmpty
            $r.Manifest.ExportFolders.Mail | Should -Be 'MBT-FXICS-Corpus/Mail'
            $r.Manifest.Folders.Event.PathResolved | Should -BeFalse
            $r.Manifest.Folders.Event.GraphId | Should -Be $r.Manifest.FolderIds.Event
            $r.Manifest.Folders.Event.Name | Should -Be 'MBT-FXICS-Corpus'
            $r.Manifest.Folders.Event.HowToFindPath | Should -Match 'Invoke-Phase0'
            $r.Output | Should -Match 'path NOT resolved by Graph'
            $r.Output | Should -Not -Match 'probably'
        } finally { $server.Dispose() }
    }

    It 'writes the manifest and exits 1 when an item is refused by every variant; the other items are created' {
        $server = [MailboxTransfer.Tests.FakeGraphServer]::new()
        $server.RefuseBodiesContaining.Add('Weekly recurring event')
        try {
            $r = Invoke-Corpus $server 'run6'
            $r.ExitCode | Should -Be 1
            $r.Output | Should -Not -Match 'Collection was modified'
            $r.Manifest | Should -Not -BeNullOrEmpty -Because $r.Output
            $r.Manifest.Items.Count | Should -Be 23
            @($r.Manifest.Items | Where-Object Status -eq 'Failed').Key | Should -Be @('fxc-22')
            @($r.Manifest.Items | Where-Object Status -eq 'Created').Count | Should -Be 22
            $r.Manifest.Folders.Event.GraphId | Should -Not -BeNullOrEmpty
        } finally { $server.Dispose() }
    }

    It 'creates fxc-07 and fxc-22 with the payloads the service accepts (no sensitivity in the body, range numbered)' {
        $server = [MailboxTransfer.Tests.FakeGraphServer]::new()
        try {
            $r = Invoke-Corpus $server 'run7' @('-Only', 'fxc-07,fxc-22')
            $r.ExitCode | Should -Be 0 -Because $r.Output
            ($r.Manifest.Items | Where-Object Key -eq 'fxc-07').Status | Should -Be 'Created'
            ($r.Manifest.Items | Where-Object Key -eq 'fxc-22').Status | Should -Be 'Created'
            ($r.Manifest.Items | Where-Object Key -eq 'fxc-07').Notes | Should -BeNullOrEmpty
        } finally { $server.Dispose() }
    }
    It 'falls back to a weaker variant, and reports a refused property as Partial without stopping' {
        $server = [MailboxTransfer.Tests.FakeGraphServer]::new()
        $server.RefuseBodiesContaining.Add('internetMessageHeaders')
        $server.RefuseBodiesContaining.Add('Name FxcShort"')
        try {
            $r = Invoke-Corpus $server 'run3'
            $r.ExitCode | Should -Be 0 -Because $r.Output
            ($r.Manifest.Items | Where-Object Key -eq 'fxc-10').Status | Should -Be 'Partial'
            ($r.Manifest.Items | Where-Object Key -eq 'fxc-08').Status | Should -Be 'Partial'
            ($r.Manifest.Items | Where-Object Key -eq 'fxc-08').Notes -join ' ' | Should -Match 'FxcShort'
            ($r.Manifest.Items | Where-Object Key -eq 'fxc-09').Status | Should -Be 'Created'
            ($r.Manifest.Items | Where-Object Status -eq 'Failed') | Should -BeNullOrEmpty
        } finally { $server.Dispose() }
    }

    It 'tries the next value of a property the service refused' {
        $server = [MailboxTransfer.Tests.FakeGraphServer]::new()
        $server.RefuseBodiesContaining.Add('"value":"1234.56"')
        try {
            $r = Invoke-Corpus $server 'run4' @('-Only', 'fxc-08')
            $r.ExitCode | Should -Be 0 -Because $r.Output
            ($r.Manifest.Items | Where-Object Key -eq 'fxc-08').Notes -join ' ' | Should -Match 'Currency.*value variant 2 accepted'
        } finally { $server.Dispose() }
    }
}

Describe 'LabFxAuth.cs' {
    It 'compiles against MSAL.NET' -Skip:(-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        $module = Get-Module -ListAvailable -Name Microsoft.Graph.Authentication | Sort-Object Version -Descending | Select-Object -First 1
        $deps = Join-Path $module.ModuleBase 'Dependencies'
        $psi = [Diagnostics.ProcessStartInfo]::new('pwsh', @('-NoProfile', '-Command', @"
`$ErrorActionPreference = 'Stop'
Add-Type -LiteralPath '$deps\Microsoft.IdentityModel.Abstractions.dll'
Add-Type -LiteralPath '$deps\Core\Microsoft.Identity.Client.dll'
Add-Type -AssemblyName System.Security.Cryptography.ProtectedData
`$refs = '$deps\Core\Microsoft.Identity.Client.dll', 'System.Security.Cryptography.ProtectedData', 'System.Linq', 'System.Runtime', 'System.Threading', 'System.Threading.Tasks', 'System.Threading.Thread', 'System.Collections', 'System.Memory', 'netstandard', 'System.Security.Cryptography', 'System.Text.Encoding.Extensions', 'System.Runtime.Extensions', 'System.IO.FileSystem', 'System.Console', 'System.Private.Uri', 'System.ComponentModel.Primitives'
Add-Type -LiteralPath '$($script:Tools)\LabFxAuth.cs' -ReferencedAssemblies `$refs -IgnoreWarnings -WarningAction SilentlyContinue
if (-not ('LabFx.Auth' -as [type])) { exit 3 }
"@))
        $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
        $p = [Diagnostics.Process]::Start($psi)
        $err = $p.StandardError.ReadToEnd() + $p.StandardOutput.ReadToEnd()
        $p.WaitForExit()
        $p.ExitCode | Should -Be 0 -Because $err
    }
}

Describe 'The PDF of the corpus' {
    It 'is a valid minimal PDF: the xref table points at every "N 0 obj", startxref at the table, /Size and /Length are right' {
        $bytes = New-FxcPdfBytes
        $text = [Text.Encoding]::ASCII.GetString($bytes)
        $text.StartsWith('%PDF-1.') | Should -BeTrue
        $text.EndsWith("%%EOF`n") | Should -BeTrue
        # startxref gives the offset of the table
        ($text -match 'startxref\n(\d+)\n%%EOF\n$') | Should -BeTrue
        $xref = [int]$Matches[1]
        $text.Substring($xref, 5) | Should -Be "xref`n"
        $lines = $text.Substring($xref) -split "`n"
        $lines[1] | Should -Be '0 6'
        $lines[2] | Should -Be '0000000000 65535 f '
        for ($n = 1; $n -le 5; $n++) {
            $entry = $lines[2 + $n]
            $entry.Length + 1 | Should -Be 20 -Because 'an entry is 20 bytes with its line feed'
            $entry | Should -Match '^\d{10} 00000 n $'
            $off = [int]$entry.Substring(0, 10)
            $text.Substring($off).StartsWith("$n 0 obj`n") | Should -BeTrue -Because "object $n is at offset $off"
        }
        $text | Should -Match 'trailer\n<</Root 1 0 R/Size 6>>\nstartxref'
        # the declared length of the content stream is its real length
        ($text -match '/Length (\d+)>>\nstream\n([^\n]*)\nendstream') | Should -BeTrue
        [int]$Matches[1] | Should -Be $Matches[2].Length
    }
}
