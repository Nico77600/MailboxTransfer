#Requires -Version 7.4
#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0.0' }
<#
    Mailbox Transfer - tests of the EML export and import (Engine.Mime.cs, Engine.EmlMapi.cs, Engine.Eml.cs, Engine.EmlImport.cs,
    MailboxTransfer.Eml.ps1). Nothing here reaches the network: the Mail API is FakeMailApi + EmlFakeMail (tests\EmlFakes.cs), the
    import-export API is FakeMailboxService, MIME fixtures are written in the tests.

    Run:  Invoke-Pester -Path .\tests\Eml.Tests.ps1 -Output Detailed
#>

BeforeAll {
    $script:RepoRoot = Split-Path $PSScriptRoot -Parent
    $script:Root = Join-Path $script:RepoRoot 'package'
    . (Join-Path $PSScriptRoot 'support\Import-StoreEngine.ps1')
    Import-StoreEngine -PackageRoot $script:Root -WithTestSupport
    . (Join-Path $PSScriptRoot 'support\Import-TestFixtures.ps1')
    Import-TestFixtures -TestsRoot $PSScriptRoot
    $script:Tenant = '1a2b3c01-5e7f-4a1b-9c2d-3e4f5a6b7c8d'
    $script:Src = "MBX:11111111-2222-3333-4444-555555555555@$($script:Tenant)"
    $script:Tmp = Join-Path ([IO.Path]::GetTempPath()) ("mbt-eml-tests-" + [guid]::NewGuid().ToString('N'))
    $null = New-Item -ItemType Directory -Path $script:Tmp

    function B([string]$Text) { [Text.Encoding]::UTF8.GetBytes(($Text -replace "(?<!`r)`n", "`r`n")) }
    function Parse([string]$Text) { [MailboxTransfer.MimeParser]::Parse((B $Text)) }
    function Shape([MailboxTransfer.MimeEntity]$E) { [MailboxTransfer.EmlCompare]::Structure($E) }

    $script:Rich = @'
From: =?UTF-8?Q?Andr=C3=A9_Dupont?= <andre@example.com>
To: "Martin, Alice" <alice@example.com>, bob@example.com (Bob)
Cc: Group: carol@example.com, dan@example.com;
Subject: =?UTF-8?B?UsOpdW5pb24gw6l0w6kg4oKs?=
Date: Tue, 19 Mar 2024 08:15:00 +0100 (CET)
Message-ID: <rich-1@example.com>
In-Reply-To: <orig@example.com>
Received: from mx2.example.com by mx1.example.com; Tue, 19 Mar 2024 07:15:10 +0000
Received: from a.example.com by mx2.example.com; Tue, 19 Mar 2024 07:15:05 +0000
Importance: high
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="MIX"

--MIX
Content-Type: multipart/alternative; boundary="ALT"

--ALT
Content-Type: text/plain; charset=utf-8
Content-Transfer-Encoding: quoted-printable

Caf=C3=A9 =E2=82=AC 100 and a long line that is soft broken =
here.
--ALT
Content-Type: multipart/related; boundary="REL"

--REL
Content-Type: text/html; charset=iso-8859-1
Content-Transfer-Encoding: quoted-printable

<html><body><p>Caf=E9</p><img src=3D"cid:logo@example.com"></body></html>
--REL
Content-Type: image/png; name="logo.png"
Content-Transfer-Encoding: base64
Content-ID: <logo@example.com>
Content-Disposition: inline; filename="logo.png"

iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmM
IQAAAABJRU5ErkJggg==
--REL--
--ALT--
--MIX
Content-Type: application/pdf; name*0*=UTF-8''r%C3%A9sum%C3%A9; name*1*=%20final.pdf
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename*0*=UTF-8''r%C3%A9sum%C3%A9; filename*1*=%20final.pdf

JVBERi0xLjQKJSBzbWFsbCBwZGY=
--MIX
Content-Type: message/rfc822
Content-Disposition: attachment; filename="inner.eml"

From: inner@example.com
To: you@example.com
Subject: Inner message
Date: Mon, 18 Mar 2024 10:00:00 +0000

Inner body.
--MIX--
'@

    function New-FakeClient {
        param([Parameter(Mandatory)]$Handler)
        $o = [MailboxTransfer.GraphOptions]::new(); $o.DelayScale = 0; $o.UseJitter = $false
        return [MailboxTransfer.GraphClient]::new([MailboxTransfer.Tests.FixedToken]::Provider, $o, $Handler)
    }

    <# A mailbox with Inbox (3 messages + 1 item that is not a message), Inbox/Projects (1), a Calendar (non-mail, 2 items) and an empty folder. #>
    function New-World {
        $ie = [MailboxTransfer.Tests.FakeMailboxService]::new()
        $ie.RootOnlyFolders = $true
        $mail = [MailboxTransfer.Tests.FakeMailApi]::new($ie)
        $null = $ie.AddMailbox($script:Src); $mail.AddMailbox($script:Src)
        $inbox = $mail.AddFolder($script:Src, 'F-INBOX', 'Inbox', $null, 'inbox', $false, 1)
        $null = $mail.AddFolder($script:Src, 'F-PROJ', 'Projects', 'F-INBOX')
        $null = $mail.AddFolder($script:Src, 'F-AUX', 'AUX', $null)          # a reserved device name
        $null = $ie.AddFolder($script:Src, 'F-INBOX', 'Inbox'); $null = $ie.AddFolder($script:Src, 'F-PROJ', 'Projects', 'F-INBOX')
        $null = $ie.AddFolder($script:Src, 'F-AUX', 'AUX')
        $null = $ie.AddFolder($script:Src, 'F-CAL', 'Calendar', 'ROOT', 'IPF.Appointment')
        $fake = [MailboxTransfer.Tests.EmlFakeMail]::new($mail, $ie)
        $w = [pscustomobject]@{ Ie = $ie; Mail = $mail; Fake = $fake; Client = $null }
        $add = {
            param($folder, $id, $subject, $day)
            $mime = B ("From: a@example.com`nTo: b@example.com`nSubject: $subject`nDate: Mon, $day Oct 2026 08:00:00 +0000`nMessage-ID: <$id@example.com>`n`nBody of $id`n")
            $m = $mail.AddMessage($script:Src, $folder, $id, $subject, [datetime]::new(2026, 10, $day, 8, 0, 0, [DateTimeKind]::Utc), $mime.Length)
            $fake.Mime["$($script:Src)|$id"] = $mime
            $m
        }
        $null = & $add 'F-INBOX' 'M1' 'Re: First: with colon?' 1
        $null = & $add 'F-INBOX' 'M2' 'Second / slash' 2
        $null = & $add 'F-INBOX' 'M3' 'Third message' 3
        $inbox.NonMessageItems = 1     # an appointment kept in the folder: counted by the folder, not a message
        $null = & $add 'F-PROJ' 'P1' 'Project one' 5
        $null = & $add 'F-AUX' 'A1' 'In a folder named AUX' 6
        $null = $ie.AddItem($script:Src, 'F-CAL', 'C1', [byte[]]::new(300), 'ck', [datetime]::new(2026, 10, 7, 8, 0, 0, [DateTimeKind]::Utc))
        $w.Client = New-FakeClient -Handler $fake
        return $w
    }

    function New-Source($World) {
        $browser = [MailboxTransfer.MailboxBrowser]::new([MailboxTransfer.MailboxApi]::new($World.Client))
        $loc = [MailboxTransfer.MailboxBrowser]::Location('Primary', $script:Src, 'test')
        return [MailboxTransfer.LiveEmlSource]::new($browser, $loc)
    }

    function Invoke-Export($Source, [string]$Dir, [switch]$WhatIf, [switch]$NoResume, [int]$Concurrency = 1) {
        $o = [MailboxTransfer.EmlExportOptions]::new()
        $o.OutputDirectory = $Dir; $o.WhatIf = [bool]$WhatIf; $o.Resume = -not $NoResume; $o.Concurrency = $Concurrency
        return [MailboxTransfer.EmlExporter]::Run($Source, $o)
    }
}

AfterAll { Remove-Item -Recurse -Force $script:Tmp -ErrorAction SilentlyContinue }

Describe 'MIME parser' {
    It 'reads headers, structure, encodings and parameters of a rich message' {
        $e = Parse $script:Rich
        Shape $e | Should -Be 'mixed(alternative(plain,related(html,image/png)),application/pdf,rfc822[plain])'
        $e.HeaderDecoded('Subject') | Should -Be "Réunion été €"
        $alt = $e.Parts[0]
        $alt.Parts[0].Text() | Should -Be "Café € 100 and a long line that is soft broken here."
        $html = $alt.Parts[1].Parts[0]
        $html.ContentType['charset'] | Should -Be 'iso-8859-1'
        $html.Text() | Should -BeLike '*<p>Café</p>*cid:logo@example.com*'
        $png = $alt.Parts[1].Parts[1]
        $png.ContentId | Should -Be 'logo@example.com'
        $png.Content.Length | Should -Be 70
        $pdf = $e.Parts[1]
        $pdf.FileName | Should -Be 'résumé final.pdf'
        $e.Parts[2].Message.Header('Subject') | Should -Be 'Inner message'
    }
    It 'parses address lists: quoted names, comments, groups, encoded names' {
        $l = [MailboxTransfer.MimeText]::ParseAddressList('"Martin, Alice" <alice@example.com>, bob@example.com (Bob), =?UTF-8?Q?Andr=C3=A9?= <a@x.fr>, Undisclosed: ;, Grp: c@x.org, d@x.org;')
        @($l | % Address) | Should -Be @('alice@example.com', 'bob@example.com', 'a@x.fr', 'c@x.org', 'd@x.org')
        $l[0].Name | Should -Be 'Martin, Alice'
        $l[1].Name | Should -Be 'Bob'
        $l[2].Name | Should -Be 'André'
    }
    It 'parses dates: zones, comments, obsolete forms, two-digit years' {
        $d = { param($s) [MailboxTransfer.MimeText]::ParseDate($s).UtcDateTime.ToString('yyyy-MM-dd HH:mm:ss') }
        & $d 'Tue, 19 Mar 2024 08:15:00 +0100 (CET)' | Should -Be '2024-03-19 07:15:00'
        & $d '19 Mar 24 8:15 EST' | Should -Be '2024-03-19 13:15:00'
        & $d 'Wed, 20 Mar 2024 08:15 GMT' | Should -Be '2024-03-20 08:15:00'
        & $d 'Thu, 21 Mar 2024 08:15:30 -0000' | Should -Be '2024-03-21 08:15:30'
        [MailboxTransfer.MimeText]::ParseDate('not a date') | Should -BeNullOrEmpty
    }
    It 'decodes encoded words split across words and charsets' {
        [MailboxTransfer.MimeText]::DecodeWords('=?UTF-8?B?w6k=?= =?UTF-8?B?dMOp?=') | Should -Be 'été'
        [MailboxTransfer.MimeText]::DecodeWords('=?ISO-8859-1?Q?Caf=E9_cr=E8me?=') | Should -Be 'Café crème'
        [MailboxTransfer.MimeText]::DecodeWords('plain =?UTF-8?Q?x?= text') | Should -Be 'plain x text'
        [MailboxTransfer.MimeText]::DecodeWords('=?bogus-charset?Q?abc?=') | Should -Be 'abc'
    }
    It 'tolerates bare LF, a missing closing boundary and a missing blank line' {
        $e = Parse "Content-Type: multipart/mixed; boundary=B`nSubject: s`n`n--B`nContent-Type: text/plain`n`none`n--B`nContent-Type: text/plain`n`ntwo"
        $e.Parts.Count | Should -Be 2
        $e.Parts[1].Text() | Should -Be 'two'
        $e.Warnings -join '|' | Should -BeLike '*not closed*'
        $e2 = Parse "Subject: only headers"
        $e2.Header('Subject') | Should -Be 'only headers'
    }
    It 'keeps a multipart without boundary, an unknown transfer encoding and a headerless text' {
        $e = Parse "Content-Type: multipart/mixed`n`nbody"
        $e.IsMultipart | Should -BeFalse
        $e.Warnings.Count | Should -BeGreaterThan 0
        $u = Parse "Content-Type: text/plain`nContent-Transfer-Encoding: x-weird`n`nraw text"
        $u.Text() | Should -Be 'raw text'
        $h = Parse "this is not a header line`nnor this`n"
        $h.Headers.Count | Should -Be 0
        $h.Text() | Should -BeLike 'this is not*'
    }
    It 'tolerates damaged base64 and quoted-printable' {
        $b = Parse "Content-Type: application/octet-stream`nContent-Transfer-Encoding: base64`n`nSGVs bG8*gd29y`nbGQ"
        [Text.Encoding]::ASCII.GetString($b.Content.ToArray()) | Should -Be 'Hello world'
        $q = Parse "Content-Type: text/plain; charset=iso-8859-1`nContent-Transfer-Encoding: quoted-printable`n`nprice =3D 5 =E9 and =ZZ stays and a=`nb"
        $q.Text() | Should -Be 'price = 5 é and =ZZ stays and ab'
    }
    It 'does not overflow on deep nesting' {
        $sb = [Text.StringBuilder]::new('MIME-Version: 1.0' + "`n")
        $inner = 'Content-Type: text/plain' + "`n`n" + 'leaf'
        for ($i = 0; $i -lt 70; $i++) { $inner = "Content-Type: multipart/mixed; boundary=b$i`n`n--b$i`n$inner`n--b$i--" }
        $e = Parse $inner
        $e | Should -Not -BeNullOrEmpty
        (@($e.Walk()).Count) | Should -BeGreaterThan 40
    }
    It 'falls back to Windows-1252 for 8-bit headers that are not UTF-8, and sniffs UTF-8 declared as us-ascii' {
        $raw = [byte[]]([Text.Encoding]::ASCII.GetBytes("Subject: caf") + 0xE9 + [Text.Encoding]::ASCII.GetBytes("`r`n`r`nbody`r`n"))
        [MailboxTransfer.MimeParser]::Parse($raw).Header('Subject') | Should -Be 'café'
        $u = Parse "Content-Type: text/plain; charset=us-ascii`nContent-Transfer-Encoding: 8bit`n`ncafé"
        $u.Text() | Should -Be 'café'
    }
}

Describe 'MIME writer' {
    It 'writes a message that parses back to the same structure, headers and content' {
        $e = Parse $script:Rich
        $bytes = [MailboxTransfer.MimeWriter]::ToBytes($e)
        $e2 = [MailboxTransfer.MimeParser]::Parse($bytes)
        Shape $e2 | Should -Be (Shape $e)
        $e2.HeaderDecoded('Subject') | Should -Be $e.HeaderDecoded('Subject')
        $e2.Parts[1].FileName | Should -Be 'résumé final.pdf'
        (Get-FileHash -InputStream ([IO.MemoryStream]::new($e2.Parts[0].Parts[1].Parts[1].Content.ToArray()))).Hash | Should -Be (Get-FileHash -InputStream ([IO.MemoryStream]::new($e.Parts[0].Parts[1].Parts[1].Content.ToArray()))).Hash
    }
    It 'folds long headers at 78 columns and encodes non-ASCII text (UTF-8 encoded words of at most 75 characters)' {
        $subject = 'Réunion équipe très importante - prévoir un café et des croissants pour tout le monde, merci beaucoup de votre ponctualité'
        $h = [MailboxTransfer.MimeText]::FoldHeader('Subject', [MailboxTransfer.MimeText]::EncodeText($subject))
        foreach ($line in $h -split "`r`n") { $line.Length | Should -BeLessOrEqual 78 }
        [MailboxTransfer.MimeText]::DecodeWords(($h -replace "`r`n ", ' ').Substring(9)) | Should -Be $subject
        foreach ($w in [regex]::Matches($h, '=\?[^ ]+\?=')) { $w.Length | Should -BeLessOrEqual 75 }
    }
    It 'writes quoted-printable lines of at most 76 characters and base64 lines of 76 characters' {
        $e = [MailboxTransfer.MimeEntity]::new()
        $e.Add('Content-Type', 'text/plain; charset="utf-8"'); $e.Add('Content-Transfer-Encoding', 'quoted-printable')
        $e.ContentType = [MailboxTransfer.MimeContentType]::Parse('text/plain; charset=utf-8'); $e.TransferEncoding = 'quoted-printable'
        $e.Content = [MailboxTransfer.FxBytesBlob]::new([Text.Encoding]::UTF8.GetBytes(('é' * 200) + "`n" + 'trailing space ' + "`n"))
        $text = [Text.Encoding]::ASCII.GetString([MailboxTransfer.MimeWriter]::ToBytes($e))
        foreach ($line in ($text -split "`r`n")) { $line.Length | Should -BeLessOrEqual 76 }
        $text | Should -BeLike '*trailing space=20*'
        $b = [MailboxTransfer.MimeEntity]::new()
        $b.ContentType = [MailboxTransfer.MimeContentType]::Parse('application/octet-stream'); $b.TransferEncoding = 'base64'
        $b.Add('Content-Type', 'application/octet-stream'); $b.Add('Content-Transfer-Encoding', 'base64')
        $b.Content = [MailboxTransfer.FxBytesBlob]::new([byte[]](0..255 * 3))
        $lines = [Text.Encoding]::ASCII.GetString([MailboxTransfer.MimeWriter]::ToBytes($b)) -split "`r`n" | Where-Object { $_ -and $_ -notmatch ':' }
        $lines[0].Length | Should -Be 76
    }
    It 'formats parameters and addresses' {
        [MailboxTransfer.MimeText]::FormatParameter('filename', 'a b.pdf') | Should -Be 'filename="a b.pdf"'
        [MailboxTransfer.MimeText]::FormatParameter('filename', 'é.pdf') | Should -Be "filename*=UTF-8''%C3%A9.pdf"
        [MailboxTransfer.MimeText]::FormatAddress('Martin, Alice', 'a@x.fr') | Should -Be '"Martin, Alice" <a@x.fr>'
        [MailboxTransfer.MimeText]::FormatAddress('André', 'a@x.fr') | Should -BeLike '=?UTF-8?*?= <a@x.fr>'
        [MailboxTransfer.MimeText]::FormatDate([DateTimeOffset]::new(2024, 3, 19, 8, 15, 0, [TimeSpan]::FromHours(1))) | Should -Be 'Tue, 19 Mar 2024 08:15:00 +0100'
    }
}

Describe 'File names' {
    It 'makes safe components: reserved device names, invalid characters, trailing dots, length' {
        [MailboxTransfer.EmlNames]::Component('AUX', 50) | Should -Be '_AUX'
        [MailboxTransfer.EmlNames]::Component('nul.txt', 50) | Should -Be '_nul.txt'
        [MailboxTransfer.EmlNames]::Component('COM1', 50) | Should -Be '_COM1'
        [MailboxTransfer.EmlNames]::Component('a<b>c:d"e/f\g|h?i*j', 50) | Should -Be 'a_b_c_d_e_f_g_h_i_j'
        [MailboxTransfer.EmlNames]::Component('trailing. . ', 50) | Should -Be 'trailing'
        [MailboxTransfer.EmlNames]::Component('', 50) | Should -Be '_'
        ([MailboxTransfer.EmlNames]::Component(('x' * 300), 40)).Length | Should -Be 40
        [MailboxTransfer.EmlNames]::Component('Auxiliary', 50) | Should -Be 'Auxiliary'
    }
    It 'keeps file names inside the path budget and distinguishes folders that differ only by case' {
        $name = [MailboxTransfer.EmlNames]::FileName([DateTimeOffset]::new(2026, 10, 1, 8, 0, 0, [TimeSpan]::Zero), ('s' * 500), 'id-1', 200, 240)
        (200 + 1 + $name.Length) | Should -BeLessOrEqual 240
        $name | Should -BeLike '20261001-080000 *.eml'
        $dirs = [MailboxTransfer.EmlNames]::FolderDirectories([string[]]@('Inbox/Report', 'Inbox/report', 'Inbox/Other', 'CON/aux'), 200)
        $dirs['Inbox/Report'] | Should -Not -Be $dirs['Inbox/report']
        $dirs['Inbox/Report'].ToLowerInvariant() | Should -Not -Be $dirs['Inbox/report'].ToLowerInvariant()
        $dirs['Inbox/Other'] | Should -Be 'Inbox\Other'
        $dirs['CON/aux'] | Should -Be '_CON\_aux'
        $long = 'a' * 100 + '/' + 'b' * 100 + '/' + 'c' * 100
        ([MailboxTransfer.EmlNames]::FolderDirectories([string[]]@($long), 120)[$long]).Length | Should -BeLessOrEqual 120
    }
}

Describe 'Bodies: RTF, HTML to text' {
    It 'reads an uncompressed (MELA) and a compressed (LZFu) RTF container' {
        $rtf = [Text.Encoding]::ASCII.GetBytes('{\rtf1\ansi Hello \par world}')
        $mela = [byte[]]([BitConverter]::GetBytes([int]($rtf.Length + 12)) + [BitConverter]::GetBytes([int]$rtf.Length) + [BitConverter]::GetBytes([uint32]0x414C454D) + [byte[]]::new(4) + $rtf)
        [Text.Encoding]::ASCII.GetString([MailboxTransfer.Rtf]::Decompress($mela)) | Should -Be '{\rtf1\ansi Hello \par world}'
        # LZFu with literals only (flag byte 0 = eight literals) and the end marker (a reference to the write position)
        $pre = 207
        $lit = [Collections.Generic.List[byte]]::new()
        $i = 0
        while ($i -lt $rtf.Length) {
            $n = [Math]::Min(8, $rtf.Length - $i)
            $flags = 0
            $lit.Add(0)   # flag byte patched below when the end marker follows
            for ($k = 0; $k -lt $n; $k++) { $lit.Add($rtf[$i + $k]) }
            $i += $n
        }
        # end marker: set flag bit n of the last group and add the 2 bytes (offset = write position, length 0)
        $lastStart = $lit.Count - ($rtf.Length % 8 -eq 0 ? 9 : ($rtf.Length % 8 + 1))
        $count = $rtf.Length % 8 -eq 0 ? 8 : $rtf.Length % 8
        if ($count -lt 8) { $lit[$lastStart] = [byte](1 -shl $count) } else { $lit.Add([byte]1) }
        $wr = ($pre + $rtf.Length) % 4096
        $lit.Add([byte](($wr -shr 4) -band 0xFF)); $lit.Add([byte]((($wr -band 0xF) -shl 4) -bor 0))
        $body = $lit.ToArray()
        $lz = [byte[]]([BitConverter]::GetBytes([int]($body.Length + 12)) + [BitConverter]::GetBytes([int]$rtf.Length) + [BitConverter]::GetBytes([uint32]0x75465A4C) + [byte[]]::new(4) + $body)
        [Text.Encoding]::ASCII.GetString([MailboxTransfer.Rtf]::Decompress($lz)) | Should -Be '{\rtf1\ansi Hello \par world}'
        [MailboxTransfer.Rtf]::Decompress([byte[]](1..20)) | Should -BeNullOrEmpty
    }
    It 'extracts the HTML encapsulated in an RTF and the text of a plain RTF' {
        $enc = '{\rtf1\ansi\ansicpg1252\fromhtml1 {\*\htmltag64 <html>}\htmlrtf {\b \htmlrtf0 Caf\''e9 {\*\htmltag <b>}bold{\*\htmltag </b>}\htmlrtf }\htmlrtf0 {\*\htmltag </html>}}'
        $html = [Text.Encoding]::UTF8.GetString([MailboxTransfer.Rtf]::ExtractHtml([Text.Encoding]::ASCII.GetBytes($enc)))
        $html | Should -Be '<html>Café <b>bold</b></html>'
        [MailboxTransfer.Rtf]::ExtractHtml([Text.Encoding]::ASCII.GetBytes('{\rtf1\ansi plain}')) | Should -BeNullOrEmpty
        [MailboxTransfer.Rtf]::ToText([Text.Encoding]::ASCII.GetBytes('{\rtf1\ansi{\fonttbl{\f0 Arial;}}Caf\''e9\par second ' + '\u8364? line}')) | Should -Be "Café`nsecond € line"
    }
    It 'builds a text alternative from HTML' {
        [MailboxTransfer.HtmlText]::FromHtml('<html><head><style>p{}</style></head><body><p>Hello&nbsp;<b>world</b></p><ul><li>a</li><li>b</li></ul><br>x &amp; y</body></html>') | Should -Be "Hello world`n`n* a`n* b`n`nx & y"
    }
}

Describe 'MIME to MAPI' {
    It 'builds the message importItem needs: delete markers, recipients, bodies, attachments, embedded message' {
        $notes = [Collections.Generic.List[string]]::new()
        $m = [MailboxTransfer.MimeToMapi]::Build((Parse $script:Rich), $null, $notes)
        $m.DelPropRecipients | Should -Not -BeNullOrEmpty
        $m.DelPropAttachments | Should -Not -BeNullOrEmpty
        $m.GetString(0x0037) | Should -Be 'Réunion été €'
        $m.GetString(0x0E1D) | Should -Be 'Réunion été €'
        $m.GetString(0x1000) | Should -BeLike 'Café €*'
        [Text.Encoding]::UTF8.GetString($m.Find(0x1013).Raw) | Should -BeLike '*<p>Café</p>*'
        $m.Find(0x3FDE).AsInt32() | Should -Be 65001
        $m.Recipients.Count | Should -Be 4
        $m.Recipients[0].Properties[0].Id | Should -Be 0x3000
        @($m.Recipients | % { [MailboxTransfer.MapiToMime]::Int($_.Properties, 0x0C15) }) | Should -Be @(1, 1, 2, 2)
        $m.Find(0x0017).AsInt32() | Should -Be 2
        $m.Attachments.Count | Should -Be 3
        $png = $m.Attachments[0]
        [MailboxTransfer.MapiToMime]::Str($png.Properties, 0x3712) | Should -Be 'logo@example.com'
        [MailboxTransfer.MapiToMime]::Str($m.Attachments[1].Properties, 0x3707) | Should -Be 'résumé final.pdf'
        $emb = $m.Attachments[2]
        $emb.Embedded | Should -Not -BeNullOrEmpty
        $emb.Embedded.Properties[0].Id | Should -Be 0x674A
        $emb.Embedded.GetString(0x0037) | Should -Be 'Inner message'
        $emb.Properties | Where-Object { $_.Id -eq 0x3701 } | Should -BeNullOrEmpty
        $m.Find(0x007D).AsString() | Should -BeLike 'From: =?UTF-8?Q?Andr*'
        ($m.Find(0x0E07).AsInt32() -band 1) | Should -Be 1
        ($m.Find(0x0E07).AsInt32() -band 0x8) | Should -Be 0
    }
    It 'takes the received time from the topmost Received header, else Date, else the override; the read state from the option' {
        $m = [MailboxTransfer.MimeToMapi]::Build((Parse $script:Rich))
        $m.Find(0x0E06).AsTime().UtcDateTime.ToString('s') | Should -Be '2024-03-19T07:15:10'
        $m.Find(0x0039).AsTime().UtcDateTime.ToString('s') | Should -Be '2024-03-19T07:15:00'
        $d = [MailboxTransfer.MimeToMapi]::Build((Parse "From: a@x.fr`nDate: Mon, 18 Mar 2024 10:00:00 +0000`nSubject: x`n`nhi"))
        $d.Find(0x0E06).AsTime().UtcDateTime.ToString('s') | Should -Be '2024-03-18T10:00:00'
        $o = [MailboxTransfer.EmlImportOptions]::new(); $o.Read = $false; $o.Received = [DateTimeOffset]::new(2020, 1, 2, 3, 4, 5, [TimeSpan]::Zero)
        $u = [MailboxTransfer.MimeToMapi]::Build((Parse "Subject: x`n`nhi"), $o)
        $u.Find(0x0E06).AsTime().UtcDateTime.ToString('s') | Should -Be '2020-01-02T03:04:05'
        ($u.Find(0x0E07).AsInt32() -band 1) | Should -Be 0
    }
    It 'encodes and decodes (FastTransfer) without warning, with the importItem envelope' {
        $m = [MailboxTransfer.MimeToMapi]::Build((Parse $script:Rich))
        $src = [MailboxTransfer.FxMessageSource]::new($m, [MailboxTransfer.FxImportStream]::NewEnvelope())
        $ms = [IO.MemoryStream]::new(); $s = $src.OpenRead(); $s.CopyTo($ms); $s.Dispose()
        $ms.Length | Should -Be $src.Length
        $bytes = $ms.ToArray()
        ($bytes[0..7] -join ',') | Should -Be '1,0,0,0,8,0,0,0'
        $doc = [MailboxTransfer.FxCodec]::Decode($bytes)
        $doc.Error | Should -BeNullOrEmpty
        $doc.Warnings.Count | Should -Be 0
        $doc.Message.Recipients.Count | Should -Be 4
        $doc.Message.Attachments.Count | Should -Be 3
        $doc.Message.Attachments[2].Embedded.GetString(0x0037) | Should -Be 'Inner message'
    }
}

Describe 'MAPI to MIME' {
    It 'writes a message built from MAPI properties, and parses back to the same content' {
        $m = [MailboxTransfer.MimeToMapi]::Build((Parse $script:Rich))
        $notes = [Collections.Generic.List[string]]::new()
        $bytes = [MailboxTransfer.MapiToMime]::ToBytes($m, $null, $notes)
        $cmp = [MailboxTransfer.EmlCompare]::Compare((B $script:Rich), $bytes)
        $cmp.ContentEqual | Should -BeTrue -Because (($cmp.Differences | Out-String))
        Shape ([MailboxTransfer.MimeParser]::Parse($bytes)) | Should -Be 'mixed(alternative(plain,related(html,image/png)),application/pdf,rfc822[plain])'
        # the headers of the original come from the transport headers, verbatim
        $e = [MailboxTransfer.MimeParser]::Parse($bytes)
        @($e.HeaderAll('Received')).Count | Should -Be 2
        $e.Header('Message-ID') | Should -Be '<rich-1@example.com>'
        $e.Header('X-MailboxTransfer-Received') | Should -BeNullOrEmpty
        $e.Parts[2].Message.Header('Subject') | Should -Be 'Inner message'
    }
    It 'builds every header from the properties when there are no transport headers, and records the received time' {
        $m = [MailboxTransfer.MimeToMapi]::Build((Parse "From: Zoé <z@x.fr>`nTo: a@x.fr, B <b@x.fr>`nCc: c@x.fr`nSubject: Hé`nDate: Mon, 18 Mar 2024 10:00:00 +0000`nMessage-ID: <q@x.fr>`nSensitivity: private`n`nhi"), (& { $o = [MailboxTransfer.EmlImportOptions]::new(); $o.KeepHeaders = $false; $o }))
        $e = [MailboxTransfer.MimeParser]::Parse([MailboxTransfer.MapiToMime]::ToBytes($m))
        $e.HeaderDecoded('From') | Should -Be 'Zoé <z@x.fr>'
        $e.Header('To') | Should -Be 'a@x.fr, B <b@x.fr>'
        $e.Header('Cc') | Should -Be 'c@x.fr'
        $e.HeaderDecoded('Subject') | Should -Be 'Hé'
        $e.Header('Sensitivity') | Should -Be 'private'
        $e.Header('X-MailboxTransfer-Received') | Should -Not -BeNullOrEmpty
        [MailboxTransfer.MimeToMapi]::ReceivedFromHeaders($e).UtcDateTime.ToString('s') | Should -Be '2024-03-18T10:00:00'
    }
    It 'builds the text alternative from an HTML-only message and keeps non-ASCII text intact' {
        $m = [MailboxTransfer.MimeToMapi]::Build((Parse "Subject: x`nContent-Type: text/html; charset=utf-8`n`n<p>Café <b>été</b></p>"))
        $m.Properties.RemoveAll([Predicate[MailboxTransfer.FxProperty]] { param($p) $p.Id -eq 0x1000 }) | Out-Null
        $notes = [Collections.Generic.List[string]]::new()
        $e = [MailboxTransfer.MimeParser]::Parse([MailboxTransfer.MapiToMime]::ToBytes($m, $null, $notes))
        Shape $e | Should -Be 'alternative(plain,html)'
        $e.Parts[0].Text() | Should -Be 'Café été'
        $e.Parts[1].Text() | Should -BeLike '*<b>été</b>*'
        $notes -join '|' | Should -BeLike '*text alternative*'
    }
    It 'reports what it cannot carry over (OLE and reference attachments) and skips non-mail classes' {
        $m = [MailboxTransfer.MimeToMapi]::Build((Parse $script:Rich))
        $m.Attachments[1].Properties.RemoveAll([Predicate[MailboxTransfer.FxProperty]] { param($p) $p.Id -eq 0x3705 }) | Out-Null
        $m.Attachments[1].Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x3705, 6))
        $notes = [Collections.Generic.List[string]]::new()
        $null = [MailboxTransfer.MapiToMime]::ToBytes($m, $null, $notes)
        $notes -join '|' | Should -BeLike '*attach method 6*'
        [MailboxTransfer.MapiToMime]::IsMailClass('IPM.Note.SMIME') | Should -BeTrue
        [MailboxTransfer.MapiToMime]::IsMailClass('IPM.Schedule.Meeting.Request') | Should -BeTrue
        [MailboxTransfer.MapiToMime]::IsMailClass('IPM.Appointment') | Should -BeFalse
        [MailboxTransfer.MapiToMime]::IsMailClass('IPM.Contact') | Should -BeFalse
    }
}

Describe 'Attachments lost in the mailbox' {
    It 'MAPI to MIME notes that the MIME skeleton lists attachments the message does not hold' {
        $skeleton = "Content-Type: multipart/mixed; boundary=o`n`n--o`nContent-Type: text/plain`n`n--o`nX-Exchange-Mime-Skeleton-Content-Id: a@x`nContent-Type: message/rfc822`nContent-Disposition: attachment; filename=middle.eml`n`n--o`nX-Exchange-Mime-Skeleton-Content-Id: b@x`nContent-Type: application/zip`nContent-Disposition: attachment; filename=a.zip`n`n--o--"
        $m = [MailboxTransfer.MimeToMapi]::Build((Parse "Subject: x`n`nhi"))
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromBinary(0x64F0, (B $skeleton)))
        $notes = [Collections.Generic.List[string]]::new()
        $null = [MailboxTransfer.MapiToMime]::ToBytes($m, $null, $notes)
        $notes -join '|' | Should -BeLike '*skeleton lists 2 attachment(s), the message holds 0*'
        # no note when the message holds them
        $ok = [MailboxTransfer.MimeToMapi]::Build((Parse $script:Rich))
        $ok.Properties.Add([MailboxTransfer.FxProperty]::FromBinary(0x64F0, (B $skeleton)))
        $n2 = [Collections.Generic.List[string]]::new()
        $null = [MailboxTransfer.MapiToMime]::ToBytes($ok, $null, $n2)
        $n2 -join '|' | Should -Not -BeLike '*skeleton*'
    }
    It 'the live export notes a MIME with attachment parts for a message the Mail API says has none' {
        $w = New-World
        $mime = B "From: a@example.com`nSubject: Phantom`nMIME-Version: 1.0`nContent-Type: multipart/mixed; boundary=b`n`n--b`nContent-Type: text/plain`n`nbody`n--b`nContent-Type: application/zip`nContent-Disposition: attachment; filename=a.zip`n`n--b--`n"
        $null = $w.Mail.AddMessage($script:Src, 'F-PROJ', 'P9', 'Phantom', [datetime]::new(2026, 10, 9, 8, 0, 0, [DateTimeKind]::Utc), $mime.Length)
        $w.Fake.Mime["$($script:Src)|P9"] = $mime
        $dir = Join-Path $script:Tmp 'live6'
        $o = [MailboxTransfer.EmlExportOptions]::new(); $o.OutputDirectory = $dir
        $o.FolderFilter = [Func[string, bool]] { param($p) $p -eq 'Inbox/Projects' }
        $r = [MailboxTransfer.EmlExporter]::Run((New-Source $w), $o)
        $r.Exported | Should -Be 2
        $phantom = [MailboxTransfer.EmlExporter]::ReadManifest($dir) | ? Id -eq 'P9'
        ($phantom.Notes -join '|') | Should -BeLike '*attachment content was lost in the mailbox*'
        $clean = [MailboxTransfer.EmlExporter]::ReadManifest($dir) | ? Id -eq 'P1'
        $clean.Notes | Should -BeNullOrEmpty
    }
}

Describe 'Comparison of two messages' {
    It 'reports identical, changed, missing headers, bodies and attachments' {
        $a = B $script:Rich
        ([MailboxTransfer.EmlCompare]::Compare($a, $a)).Identical | Should -BeTrue
        $changed = (Parse $script:Rich)
        $changed.Headers.RemoveAll([Predicate[MailboxTransfer.MimeHeader]] { param($h) $h.Name -eq 'Importance' }) | Out-Null
        $changed.Add('X-New', 'v')
        $changed.Parts.RemoveAt(1)
        $cmp = [MailboxTransfer.EmlCompare]::Compare($a, [MailboxTransfer.MimeWriter]::ToBytes($changed))
        ($cmp.Differences | ? { $_.Kind -eq 'Header' -and $_.Name -eq 'Importance' }).Result | Should -Be 'OnlyLeft'
        ($cmp.Differences | ? { $_.Kind -eq 'Header' -and $_.Name -eq 'X-New' }).Result | Should -Be 'OnlyRight'
        ($cmp.Differences | ? { $_.Kind -eq 'Attachment' }).Result | Should -Contain 'OnlyLeft'
        $cmp.ContentEqual | Should -BeFalse
    }
    It 'treats re-serialized HTML and equivalent address forms as the same' {
        $x = Parse "From: A <a@x.fr>`nContent-Type: text/html`n`n<P CLASS=a>hi</P>"
        $y = Parse "From: a@x.fr`nContent-Type: text/html`n`n<p class=`"a`">hi</p>`n"
        $cmp = [MailboxTransfer.EmlCompare]::Compare($x, $y)
        $cmp.Identical | Should -BeTrue -Because ($cmp.Differences | Out-String)
    }
}

Describe 'Export from a live mailbox (Mail API $value)' {
    It 'turns the folder tree into a directory tree, one .eml per message, with a manifest, and reports what is not a message' {
        $w = New-World
        $dir = Join-Path $script:Tmp 'live1'
        $r = Invoke-Export (New-Source $w) $dir
        $r.Exported | Should -Be 5
        $r.Failed | Should -Be 0
        $r.Skipped | Should -Be 1
        $r.Complete | Should -BeTrue
        (Get-ChildItem $dir -Recurse -Filter *.eml).Count | Should -Be 5
        Test-Path (Join-Path $dir 'Inbox\Projects') | Should -BeTrue
        Test-Path (Join-Path $dir '_AUX') | Should -BeTrue -Because 'AUX is a reserved device name'
        ($r.FolderNotes -join '|') | Should -BeLike '*Calendar*'
        ($r.Problems | % Reason) -join '|' | Should -BeLike '*1 item(s) of this folder are not messages*'
        $entries = [MailboxTransfer.EmlExporter]::ReadManifest($dir)
        @($entries | ? State -eq 'Exported').Count | Should -Be 5
        foreach ($e in @($entries | ? State -eq 'Exported')) {
            $f = Join-Path $dir ($e.File -replace '/', '\')
            (Get-FileHash -LiteralPath $f -Algorithm SHA256).Hash | Should -Be $e.Sha256
            (Get-Item -LiteralPath $f).Length | Should -Be $e.Size
            $e.Subject | Should -Not -BeNullOrEmpty
            $e.Id | Should -Not -BeNullOrEmpty
        }
        # the file is exactly what the Mail API returned
        $m1 = $entries | ? Id -eq 'M1'
        [IO.File]::ReadAllBytes((Join-Path $dir ($m1.File -replace '/', '\'))) | Should -Be $w.Fake.Mime["$($script:Src)|M1"]
        # a subject with a colon, a question mark and a slash gives a valid name
        ($entries | ? Id -eq 'M2').File | Should -Not -Match '[<>:"|?*]'
        @(Get-ChildItem $dir -Recurse -Filter *.eml | % { $_.Name }) | Should -Contain ((Get-ChildItem $dir -Recurse -Filter *.eml | ? Name -like '*Second*' | Select-Object -First 1).Name)
    }
    It 'resumes: a message already exported is not read again, a failed one is tried again, nothing is exported twice' {
        $w = New-World
        $dir = Join-Path $script:Tmp 'live2'
        $w.Fake.MissingMime.Add('M2') | Out-Null
        $r1 = Invoke-Export (New-Source $w) $dir
        $r1.Exported | Should -Be 4
        $r1.Failed | Should -Be 1
        $r1.Complete | Should -BeFalse
        $r1.Problems | ? State -eq 'Failed' | % Reason | Should -BeLike '*404*'
        $w.Fake.MissingMime.Clear()
        $w.Fake.Log.Clear()
        $r2 = Invoke-Export (New-Source $w) $dir
        $r2.Exported | Should -Be 1
        $r2.Resumed | Should -Be 4
        $r2.Failed | Should -Be 0
        @($w.Fake.Log | ? { $_ -like '*$value' }).Count | Should -Be 1
        @(Get-ChildItem $dir -Recurse -Filter *.eml).Count | Should -Be 5
        @(Get-ChildItem $dir -Recurse -Filter *.part).Count | Should -Be 0
        $r3 = Invoke-Export (New-Source $w) $dir
        $r3.Exported | Should -Be 0
        $r3.Resumed | Should -Be 5
    }
    It 'with WhatIf writes nothing and reads no message' {
        $w = New-World
        $dir = Join-Path $script:Tmp 'live3'
        $r = Invoke-Export (New-Source $w) $dir -WhatIf
        $r.Exported | Should -Be 5
        Test-Path $dir | Should -BeFalse
        @($w.Fake.Log | ? { $_ -like '*$value' }).Count | Should -Be 0
    }
    It 'reads with several workers' {
        $w = New-World
        $dir = Join-Path $script:Tmp 'live4'
        $r = Invoke-Export (New-Source $w) $dir -Concurrency 4
        $r.Exported | Should -Be 5
        @([MailboxTransfer.EmlExporter]::ReadManifest($dir) | ? State -eq 'Exported').Count | Should -Be 5
    }
    It 'refuses an auxiliary archive with a message that says what to do' {
        $w = New-World
        $browser = [MailboxTransfer.MailboxBrowser]::new([MailboxTransfer.MailboxApi]::new($w.Client))
        $loc = [MailboxTransfer.MailboxBrowser]::Location('AuxArchive', "MBX:bbbbbbbb-0000-0000-0000-00000000000b@$($script:Tenant)", 'test')
        { [MailboxTransfer.LiveEmlSource]::new($browser, $loc) } | Should -Throw '*auxiliary archive*'
    }
    It 'limits the export to a folder and its subfolders' {
        $w = New-World
        $dir = Join-Path $script:Tmp 'live5'
        $o = [MailboxTransfer.EmlExportOptions]::new(); $o.OutputDirectory = $dir
        $o.FolderFilter = [Func[string, bool]] { param($p) $p -eq 'Inbox/Projects' }
        $r = [MailboxTransfer.EmlExporter]::Run((New-Source $w), $o)
        $r.Exported | Should -Be 1
        (Get-ChildItem $dir -Recurse -Filter *.eml).Directory.Name | Should -Be 'Projects'
    }
}

Describe 'Export from a PST (any location)' {
    It 'writes the .eml files of a PST built from MAPI messages, skipping non-mail items, and the files read back as the originals' {
        $pst = Join-Path $script:Tmp 'src.pst'
        $w = [MailboxTransfer.PstSplitWriter]::new($pst, 'Test archive', [long]1GB)
        $m1 = [MailboxTransfer.MimeToMapi]::Build((Parse $script:Rich))
        $null = $w.AddMessage('Inbox/Rich', 'IPF.Note', $m1)
        $m2 = [MailboxTransfer.MimeToMapi]::Build((Parse "From: a@x.fr`nTo: b@x.fr`nSubject: Hello`nDate: Mon, 18 Mar 2024 10:00:00 +0000`n`nPlain body`n"))
        $null = $w.AddMessage('Inbox', 'IPF.Note', $m2)
        $cal = [MailboxTransfer.MimeToMapi]::Build((Parse "Subject: Lunch`n`nx"), (& { $o = [MailboxTransfer.EmlImportOptions]::new(); $o.MessageClass = 'IPM.Appointment'; $o }))
        $null = $w.AddMessage('Calendar', 'IPF.Appointment', $cal)
        $w.Close()
        $reader = [MailboxTransfer.PstReader]::Open($pst)
        try {
            $dir = Join-Path $script:Tmp 'pst1'
            $r = Invoke-Export ([MailboxTransfer.PstEmlSource]::new($reader, 'src')) $dir
        } finally { $reader.Dispose() }
        $r.Exported | Should -Be 2
        $r.Failed | Should -Be 0
        $files = @(Get-ChildItem $dir -Recurse -Filter *.eml)
        $files.Count | Should -Be 2
        $rich = $files | ? { $_.FullName -like '*Rich*' }
        $cmp = [MailboxTransfer.EmlCompare]::Compare((B $script:Rich), [IO.File]::ReadAllBytes($rich.FullName))
        $cmp.ContentEqual | Should -BeTrue -Because ($cmp.Differences | Out-String)
        ($cmp.Items | ? { $_.Kind -eq 'Header' -and $_.Name -eq 'Subject' }).Result | Should -Be 'Same'
    }
}

Describe 'Import' {
    BeforeAll {
        function New-ImportWorld {
            $ie = [MailboxTransfer.Tests.FakeMailboxService]::new()
            $ie.RootOnlyFolders = $true
            $mail = [MailboxTransfer.Tests.FakeMailApi]::new($ie)
            $null = $ie.AddMailbox($script:Src); $mail.AddMailbox($script:Src)
            $fake = [MailboxTransfer.Tests.EmlFakeMail]::new($mail, $ie)
            return [pscustomobject]@{ Ie = $ie; Fake = $fake; Client = (New-FakeClient -Handler $fake) }
        }
        function New-EmlDir([string]$Name) {
            $d = Join-Path $script:Tmp $Name
            New-Item -ItemType Directory -Force (Join-Path $d 'Sub folder') | Out-Null
            [IO.File]::WriteAllBytes((Join-Path $d 'rich.eml'), (B $script:Rich))
            [IO.File]::WriteAllBytes((Join-Path $d 'Sub folder\plain.eml'), (B "From: a@x.fr`nTo: b@x.fr`nSubject: Plain`nDate: Mon, 18 Mar 2024 10:00:00 +0000`n`nPlain body`n"))
            return $d
        }
        function Invoke-Import($World, [string]$Dir, [string]$Root = 'MBT-EmlImport-Test', [switch]$NoResume, [switch]$WhatIf, [switch]$Exchange, [string[]]$Include, [string[]]$Exclude) {
            $run = [MailboxTransfer.EmlImportRun]::new()
            $run.Path = $Dir; $run.Api = [MailboxTransfer.MailboxApi]::new($World.Client); $run.MailboxId = $script:Src; $run.TargetRoot = $Root
            foreach ($i in $Include) { $run.IncludeFolders.Add($i) }
            foreach ($x in $Exclude) { $run.ExcludeFolders.Add($x) }
            $run.Resume = -not $NoResume; $run.WhatIf = [bool]$WhatIf
            if ($Exchange) { $run.Conversion = [MailboxTransfer.EmlConversion]::Exchange; $run.Exchange = [MailboxTransfer.EmlExchangeConverter]::new($run.Api, $script:Src, $script:Src) }
            return [MailboxTransfer.EmlImporter]::Run($run)
        }
    }
    It 'creates the folder tree below the root and one item per file with the stream importItem needs' {
        $w = New-ImportWorld
        $dir = New-EmlDir 'imp1'
        $r = Invoke-Import $w $dir
        $r.Done | Should -Be 2
        $r.Failed | Should -Be 0
        $r.FoldersCreated | Should -Be 2
        $names = @($w.Ie.FoldersOf($script:Src) | % Name)
        $names | Should -Contain 'MBT-EmlImport-Test'
        $names | Should -Contain 'Sub folder'
        $items = @($w.Ie.ItemsOf($script:Src))
        $items.Count | Should -Be 2
        $richItem = $items | ? { $_.Data.Length -gt 3000 }
        $doc = [MailboxTransfer.FxCodec]::Decode($richItem.Data)
        $doc.Error | Should -BeNullOrEmpty
        $doc.Message.Recipients.Count | Should -Be 4
        $doc.Message.Attachments.Count | Should -Be 3
        $doc.Envelope.Entries[0].Type | Should -Be 1
        # the folder of the plain message is the subfolder
        $sub = $w.Ie.FoldersOf($script:Src) | ? Name -eq 'Sub folder'
        @($w.Ie.ItemsOf($script:Src, $sub.Id)).Count | Should -Be 1
        # nothing was changed or deleted in the source folder, and a state file was written
        Test-Path (Join-Path $dir 'import-state.jsonl') | Should -BeTrue
    }
    It 'resumes: files done are not imported again; Unknown (the request may have created the item) is not replayed' {
        $w = New-ImportWorld
        $dir = New-EmlDir 'imp2'
        $w.Ie.ImportFailures.Enqueue(0); $w.Ie.ImportFailures.Enqueue(400)
        $r1 = Invoke-Import $w $dir
        $r1.Done | Should -Be 1
        $r1.Failed | Should -Be 1
        $r1.Complete | Should -BeFalse
        $r2 = Invoke-Import $w $dir
        $r2.Done | Should -Be 1       # the failed file is tried again
        $r2.Resumed | Should -Be 1
        @($w.Ie.ItemsOf($script:Src)).Count | Should -Be 2
        $r3 = Invoke-Import $w $dir
        $r3.Done | Should -Be 0
        $r3.Resumed | Should -Be 2
        # a server error is Unknown and stays so
        $w2 = New-ImportWorld
        $dir2 = New-EmlDir 'imp2b'
        $w2.Ie.ImportFailures.Enqueue(503); $w2.Ie.ImportFailures.Enqueue(0)
        $w2.Client.Options.MaxAttempts = 1
        $u = Invoke-Import $w2 $dir2
        $u.Unknown + $u.Failed | Should -BeGreaterThan 0
    }
    It 'imports only the folders chosen (-IncludeFolder / -ExcludeFolder), counts the files left out and reports a pattern that matched nothing' {
        $w = New-ImportWorld
        $dir = New-EmlDir 'imp-choice'
        $r = Invoke-Import $w $dir -Include 'Sub folder', 'Nope'
        $r.Done | Should -Be 1
        $r.Filtered | Should -Be 1
        @($r.UnmatchedPatterns) | Should -Be @("-IncludeFolder 'Nope'")
        @($w.Ie.ItemsOf($script:Src)).Count | Should -Be 1
        $w2 = New-ImportWorld
        $r2 = Invoke-Import $w2 (New-EmlDir 'imp-choice2') -Exclude 'sub*'
        $r2.Done | Should -Be 1
        $r2.Filtered | Should -Be 1
        @($w2.Ie.FoldersOf($script:Src) | % Name) | Should -Not -Contain 'Sub folder'
        $r3 = Invoke-Import (New-ImportWorld) (New-EmlDir 'imp-choice3') -Include 'Inbx'
        $r3.Done | Should -Be 0
        $r3.Filtered | Should -Be 2
    }
    It 'a failure before the request leaves (folder creation) is Failed, not Unknown: nothing can have been created' {
        $w = New-ImportWorld
        $dir = New-EmlDir 'imp-early'
        $w.Ie.RefuseFolderNames.Add('Sub folder'); $w.Ie.RefuseFolderStatus = 500
        $w.Client.Options.MaxAttempts = 1
        $r = Invoke-Import $w $dir
        $r.Done | Should -Be 1
        $r.Failed | Should -Be 1
        $r.Unknown | Should -Be 0
        (@($r.Entries | ? State -EQ 'Failed')[0]).File | Should -BeLike '*plain.eml'
        # tried again by the next run
        $w.Ie.RefuseFolderNames.Clear()
        $again = Invoke-Import $w $dir
        $again.Done | Should -Be 1
        $again.Resumed | Should -Be 1
    }
    It 'an answer that cannot be read after the request left is Unknown (InFlight first in the state), never sent again, and settled by hand with Resolve-MbtEmlUnknown' {
        $w = New-ImportWorld
        $dir = New-EmlDir 'imp-unknown'
        $w.Ie.ImportAnswers.Enqueue('garbage')
        $r = Invoke-Import $w $dir
        $r.Unknown | Should -Be 1
        $r.Complete | Should -BeFalse
        $lines = @(Get-Content -LiteralPath (Join-Path $dir 'import-state.jsonl') | % { $_ | ConvertFrom-Json })
        $lines[0].State | Should -Be 'InFlight'
        $lines[1].State | Should -Be 'Unknown'
        $items = @($w.Ie.ItemsOf($script:Src)).Count
        # the next run reports it, and does not send it again
        $again = Invoke-Import $w $dir
        $again.Unknown | Should -Be 1
        $again.Done | Should -Be 0
        @($w.Ie.ItemsOf($script:Src)).Count | Should -Be $items
        # a person looked in the mailbox: it is there
        $settled = Resolve-MbtEmlUnknown -Path $dir -As Imported -Confirm:$false
        $settled.Count | Should -Be 1
        $done = Invoke-Import $w $dir
        $done.Unknown | Should -Be 0
        $done.Resumed | Should -Be 2
        # or it is not: the next run imports it
        $w2 = New-ImportWorld
        $dir2 = New-EmlDir 'imp-unknown2'
        $w2.Ie.ImportAnswers.Enqueue('garbage')
        $null = Invoke-Import $w2 $dir2
        @(Resolve-MbtEmlUnknown -Path $dir2 -As NotImported -Confirm:$false).Count | Should -Be 1
        $retry = Invoke-Import $w2 $dir2
        $retry.Done | Should -Be 1
        $retry.Unknown | Should -Be 0
    }
    It 'a run killed inside the request leaves an InFlight line: the file is Unknown at the next run, not imported again' {
        $w = New-ImportWorld
        $dir = New-EmlDir 'imp-inflight'
        $key = "$($script:Src)|MBT-EmlImport-Test"
        Set-Content -LiteralPath (Join-Path $dir 'import-state.jsonl') -Value (@{ State = 'InFlight'; File = 'rich.eml'; Folder = ''; TargetRoot = $key; Bytes = 10 } | ConvertTo-Json -Compress)
        $r = Invoke-Import $w $dir
        $r.Unknown | Should -Be 1
        $r.Done | Should -Be 1
        $r.Entries.Where({ $_.State -eq 'Unknown' })[0].Error | Should -BeLike '*in flight*'
        @($w.Ie.ItemsOf($script:Src)).Count | Should -Be 1
    }
    It 'with WhatIf creates nothing' {
        $w = New-ImportWorld
        $dir = New-EmlDir 'imp3'
        $r = Invoke-Import $w $dir -WhatIf
        $r.Done | Should -Be 2
        @($w.Ie.ItemsOf($script:Src)).Count | Should -Be 0
        @($w.Ie.FoldersOf($script:Src)).Count | Should -Be 0
        Test-Path (Join-Path $dir 'import-state.jsonl') | Should -BeFalse
    }
    It 'uses the manifest of an export: read state, received time and the original folder names' {
        $w = New-World
        $dir = Join-Path $script:Tmp 'imp4'
        $null = Invoke-Export (New-Source $w) $dir
        $iw = New-ImportWorld
        $r = Invoke-Import $iw $dir 'MBT-Back'
        $r.Done | Should -Be 5
        $names = @($iw.Ie.FoldersOf($script:Src) | % Name)
        $names | Should -Contain 'Projects'
        $names | Should -Contain 'AUX'      # the real folder name, not the made-safe directory name
        $m1 = $iw.Ie.ItemsOf($script:Src) | % { [MailboxTransfer.FxCodec]::Decode($_.Data) } | ? { $_.Message.GetString(0x0037) -like 'Re: First*' }
        $m1.Message.Find(0x0E06).AsTime().UtcDateTime.ToString('s') | Should -Be '2026-10-01T08:00:00'
        ($m1.Message.Find(0x0E07).AsInt32() -band 1) | Should -Be 1
    }
    It 'Exchange conversion: posts the MIME as a draft, patches flags and received time of the exported stream, imports it, deletes the draft' {
        $w = New-ImportWorld
        $dir = Join-Path $script:Tmp 'imp5'
        New-Item -ItemType Directory -Force $dir | Out-Null
        $eml = B "From: a@x.fr`nTo: b@x.fr`nSubject: Via Exchange`nDate: Mon, 18 Mar 2024 10:00:00 +0000`nReceived: by x; Mon, 18 Mar 2024 10:00:05 +0000`n`nbody`n"
        [IO.File]::WriteAllBytes((Join-Path $dir 'one.eml'), $eml)
        # what Exchange would export for the draft: flags 0x0E07 = draft + attachments bit, delivery time = now
        $m = [MailboxTransfer.MimeToMapi]::Build([MailboxTransfer.MimeParser]::Parse($eml))
        $m.Properties.RemoveAll([Predicate[MailboxTransfer.FxProperty]] { param($p) $p.Id -eq 0x0E07 }) | Out-Null
        $m.Properties.Add([MailboxTransfer.FxProperty]::FromInt32(0x0E07, 0x8 -bor 0x400))
        $ms = [IO.MemoryStream]::new(); $s = [MailboxTransfer.FxMessageSource]::new($m, [MailboxTransfer.FxImportStream]::NewEnvelope()).OpenRead(); $s.CopyTo($ms); $s.Dispose()
        $w.Fake.NextDraftData = $ms.ToArray()
        $null = $w.Ie.AddFolder($script:Src, 'ROOT', 'Dummy')
        $w.Fake.DraftFolder = 'ROOT'
        $r = Invoke-Import $w $dir -Exchange
        $r.Done | Should -Be 1 -Because ($r.Entries | % Error | Out-String)
        $w.Fake.LastDraftContentType | Should -Be 'text/plain'
        [IO.File]::ReadAllBytes((Join-Path $dir 'one.eml')) | Should -Be $w.Fake.LastDraftBody
        $w.Fake.DraftsDeleted.Count | Should -Be 1
        $item = @($w.Ie.ItemsOf($script:Src) | ? { $_.Id -like 'NEW*' })
        $item.Count | Should -Be 1
        $doc = [MailboxTransfer.FxCodec]::Decode($item[0].Data)
        $flags = $doc.Message.Find(0x0E07).AsInt32()
        ($flags -band 0x8) | Should -Be 0 -Because 'not a draft'
        ($flags -band 1) | Should -Be 1 -Because 'read'
        ($flags -band 0x400) | Should -Be 0x400 -Because 'the other bits are kept'
        $doc.Message.Find(0x0E06).AsTime().UtcDateTime.ToString('s') | Should -Be '2024-03-18T10:00:05'
    }
    It 'Exchange conversion without Mail.ReadWrite fails with a message that names the permission, and leaves nothing behind' {
        $w = New-ImportWorld
        $w.Fake.DenyDrafts = $true
        $dir = Join-Path $script:Tmp 'imp6'
        New-Item -ItemType Directory -Force $dir | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $dir 'one.eml'), (B "Subject: x`n`nhi"))
        $r = Invoke-Import $w $dir -Exchange
        $r.Done | Should -Be 0
        $r.Failed | Should -Be 1
        $r.Entries[0].Error | Should -BeLike '*Mail.ReadWrite*'
        @($w.Ie.ItemsOf($script:Src)).Count | Should -Be 0
        $w.Fake.DraftsDeleted.Count | Should -Be 0
    }
    It 'the PowerShell function refuses -MimeConversion Exchange when the token lacks Mail.ReadWrite' {
        $w = New-ImportWorld
        $null = Set-MbtSession -Graph $w.Client
        $s = Get-MbtSession
        $s.Token = [pscustomobject]@{ Scopes = @('Mail.Read', 'MailboxItem.ImportExport') }
        $dir = Join-Path $script:Tmp 'imp7'
        New-Item -ItemType Directory -Force $dir | Out-Null
        [IO.File]::WriteAllBytes((Join-Path $dir 'one.eml'), (B "Subject: x`n`nhi"))
        { Import-MbtEml -Path $dir -MailboxId $script:Src -TargetRoot T -MimeConversion Exchange -Confirm:$false } | Should -Throw '*Mail.ReadWrite*'
    }
}
