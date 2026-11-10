# LabFxCorpus.Lib.ps1 - the OFFLINE part of New-LabFxCorpus.ps1 (dot-source it; it never runs alone and never touches the network):
# the content of the corpus as a plan (steps with their Graph payloads), the small binaries used as attachments and the MIME builder.
# Everything here is deterministic, so the plan can be shown with -WhatIf and checked by tests\LabFxCorpus.Tests.ps1.
#
# A step is one item to create:
#   Key         fxc-NN (also the marker "[fxc-NN]" at the start of the subject: an existing marker means "already created")
#   Kind        Mail | Event | Contact
#   Variants    JSON bodies for the creation, the FIRST one the service accepts is used (the next ones are weaker fallbacks)
#   Mime        (Mail) a complete MIME message, created with Content-Type text/plain + base64 (instead of Variants)
#   Attachments @{ Name; Variants } posted after the creation, first accepted variant used
#   Ext         extended properties set one by one after the creation: @{ Id; Multi; Variants = values to try in order }
#   Photo       (Contact) JPEG bytes sent with PUT .../photo/$value
#   Special     RecurringExceptions: the recurring event gets one modified and one cancelled occurrence

Set-StrictMode -Version Latest

$script:FxcPropSet = '{7c3a52f0-4b21-4e6d-9d2e-5a1f0c8b3e41}'    # GUID of the named properties of the corpus (made up, not a Microsoft one)

#region Small helpers -------------------------------------------------------------------------------------------------------------

function Get-FxcMarker {
    param([Parameter(Mandatory)][string]$Key)
    "[$Key]"
}

function New-FxcAddress {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Address)
    @{ emailAddress = [ordered]@{ name = $Name; address = $Address } }
}

function New-FxcDateTime {
    param([Parameter(Mandatory)][string]$Local, [string]$TimeZone = 'UTC')
    [ordered]@{ dateTime = $Local; timeZone = $TimeZone }
}

function New-FxcFileAttachment {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$ContentType, [Parameter(Mandatory)][byte[]]$Bytes, [string]$ContentId)
    $a = [ordered]@{
        '@odata.type' = '#microsoft.graph.fileAttachment'
        name          = $Name
        contentType   = $ContentType
        contentBytes  = [Convert]::ToBase64String($Bytes)
        isInline      = [bool]$ContentId
    }
    if ($ContentId) { $a['contentId'] = $ContentId }
    $a
}

function New-FxcExt {
    # One extended property to set. -Value: the values to try in order (for a multi-valued one, each is an array of strings).
    param([Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)][object[]]$Value, [switch]$Multi)
    [pscustomobject]@{ Id = $Id; Multi = [bool]$Multi; Variants = [object[]]$Value }
}

function New-FxcStep {
    param(
        [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][ValidateSet('Mail', 'Event', 'Contact')][string]$Kind,
        [Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Subject,
        [object[]]$Variants = @(), [string]$Mime, [object[]]$Attachments = @(), [object[]]$Ext = @(), [byte[]]$Photo,
        [string]$Special, [switch]$ReportSubject
    )
    [pscustomobject]@{
        Key = $Key; Kind = $Kind; Marker = (Get-FxcMarker $Key); Title = $Title; Subject = $Subject
        Variants = [object[]]$Variants; Mime = $Mime; Attachments = [object[]]$Attachments; Ext = [object[]]$Ext
        Photo = $Photo; Special = $Special; ReportSubject = [bool]$ReportSubject
    }
}

#endregion

#region Binaries ------------------------------------------------------------------------------------------------------------------

function Get-FxcCrc32 {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    if (-not (Get-Variable -Scope Script -Name FxcCrcTable -ErrorAction SilentlyContinue)) {
        $t = [long[]]::new(256)
        for ($n = 0; $n -lt 256; $n++) {
            [long]$c = $n
            for ($k = 0; $k -lt 8; $k++) { if (($c -band 1) -eq 1) { $c = 3988292384 -bxor ($c -shr 1) } else { $c = $c -shr 1 } }
            $t[$n] = $c
        }
        $script:FxcCrcTable = $t
    }
    [long]$crc = 4294967295
    foreach ($b in $Bytes) { $crc = $script:FxcCrcTable[($crc -bxor $b) -band 255] -bxor ($crc -shr 8) }
    $crc -bxor 4294967295
}

function ConvertTo-FxcBigEndian {
    param([Parameter(Mandatory)][long]$Value)
    [byte[]]@((($Value -shr 24) -band 255), (($Value -shr 16) -band 255), (($Value -shr 8) -band 255), ($Value -band 255))
}

function New-FxcPng {
    # A real 16x16 RGB PNG (gradient), a few hundred bytes.
    param([int]$Size = 16)
    $raw = [IO.MemoryStream]::new()
    for ($y = 0; $y -lt $Size; $y++) {
        $raw.WriteByte(0)                                        # filter: none
        for ($x = 0; $x -lt $Size; $x++) {
            $raw.WriteByte([byte](($x * 255) / ($Size - 1))); $raw.WriteByte([byte](($y * 255) / ($Size - 1))); $raw.WriteByte(160)
        }
    }
    $zipped = [IO.MemoryStream]::new()
    $z = [IO.Compression.ZLibStream]::new($zipped, [IO.Compression.CompressionLevel]::SmallestSize, $true)
    $z.Write($raw.ToArray(), 0, [int]$raw.Length); $z.Dispose()
    $out = [IO.MemoryStream]::new()
    $out.Write([byte[]]@(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A), 0, 8)
    $chunk = {
        param([string]$Type, [byte[]]$Data)
        $body = [byte[]]([Text.Encoding]::ASCII.GetBytes($Type) + $Data)
        $out.Write((ConvertTo-FxcBigEndian $Data.Length), 0, 4)
        $out.Write($body, 0, $body.Length)
        $out.Write((ConvertTo-FxcBigEndian (Get-FxcCrc32 $body)), 0, 4)
    }
    $ihdr = [byte[]]((ConvertTo-FxcBigEndian $Size) + (ConvertTo-FxcBigEndian $Size) + [byte[]]@(8, 2, 0, 0, 0))
    & $chunk 'IHDR' $ihdr
    & $chunk 'IDAT' $zipped.ToArray()
    & $chunk 'IEND' ([byte[]]@())
    , $out.ToArray()
}

function New-FxcJpeg {
    # A small JPEG made by System.Drawing (Windows); $null when it is not available.
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        $bmp = [System.Drawing.Bitmap]::new(48, 48)
        for ($y = 0; $y -lt 48; $y++) { for ($x = 0; $x -lt 48; $x++) { $bmp.SetPixel($x, $y, [System.Drawing.Color]::FromArgb(255, ($x * 5) % 256, ($y * 5) % 256, 128)) } }
        $ms = [IO.MemoryStream]::new()
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $bmp.Dispose()
        , $ms.ToArray()
    } catch { $null }
}

function New-FxcPdfBytes {
    # A minimal, valid one-page PDF written as ASCII: header, five objects, a cross-reference table with the real byte offset of
    # every object (20-byte entries), the trailer with /Size, startxref and %%EOF. The offsets and the stream /Length are computed
    # from what is written, so a strict reader opens it without having to rebuild the table.
    $content = 'BT /F1 12 Tf 20 50 Td (fxc corpus) Tj ET'
    $objects = @(
        '<</Type/Catalog/Pages 2 0 R>>',
        '<</Type/Pages/Kids[3 0 R]/Count 1>>',
        '<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 100]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>',
        "<</Length $($content.Length)>>`nstream`n$content`nendstream",
        '<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>')
    $sb = [Text.StringBuilder]::new("%PDF-1.1`n")        # ASCII only: a character is a byte, so a length is an offset
    $offsets = [Collections.Generic.List[int]]::new()
    for ($i = 0; $i -lt $objects.Count; $i++) {
        $offsets.Add($sb.Length)
        [void]$sb.Append("$($i + 1) 0 obj`n$($objects[$i])`nendobj`n")
    }
    $xref = $sb.Length
    [void]$sb.Append("xref`n0 $($objects.Count + 1)`n0000000000 65535 f `n")
    foreach ($o in $offsets) { [void]$sb.Append(('{0:D10} 00000 n ' -f $o) + "`n") }
    [void]$sb.Append("trailer`n<</Root 1 0 R/Size $($objects.Count + 1)>>`nstartxref`n$xref`n%%EOF`n")
    , [Text.Encoding]::ASCII.GetBytes($sb.ToString())
}

function New-FxcZipBytes {
    # A real zip with two small entries and a fixed time stamp.
    $ms = [IO.MemoryStream]::new()
    $zip = [IO.Compression.ZipArchive]::new($ms, [IO.Compression.ZipArchiveMode]::Create, $true)
    foreach ($e in @(@('readme.txt', 'FastTransfer corpus zip entry 1'), @("d`u{E9}j`u{E0}/vu.txt", 'entry 2 with a Unicode path'))) {
        $entry = $zip.CreateEntry($e[0]); $entry.LastWriteTime = [DateTimeOffset]::new(2026, 1, 1, 0, 0, 0, [TimeSpan]::Zero)
        $w = [IO.StreamWriter]::new($entry.Open(), [Text.UTF8Encoding]::new($false)); $w.Write($e[1]); $w.Dispose()
    }
    $zip.Dispose()
    , $ms.ToArray()
}

function New-FxcPseudoBytes {
    param([int]$Seed = 1, [int]$Length = 65536)
    $b = [byte[]]::new($Length)
    [Random]::new($Seed).NextBytes($b)
    , $b
}

#endregion

#region MIME ----------------------------------------------------------------------------------------------------------------------

function ConvertTo-FxcMimeWord {
    # RFC 2047 encoded word for a header value that is not plain ASCII.
    param([Parameter(Mandatory)][AllowEmptyString()][string]$Text)
    if ($Text -cmatch '^[\x20-\x7E]*$') { return $Text }
    '=?utf-8?B?' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Text)) + '?='
}

function Format-FxcMimeAddress {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$Address)
    $n = ConvertTo-FxcMimeWord $Name
    if ($n -ceq $Name) { $n = '"' + $Name.Replace('\', '\\').Replace('"', '\"') + '"' }
    "$n <$Address>"
}

function ConvertTo-FxcBase64Lines {
    param([Parameter(Mandatory)][byte[]]$Bytes)
    $b64 = [Convert]::ToBase64String($Bytes)
    $lines = for ($i = 0; $i -lt $b64.Length; $i += 76) { $b64.Substring($i, [Math]::Min(76, $b64.Length - $i)) }
    ($lines -join "`r`n")
}

function New-FxcMimeText {
    param([Parameter(Mandatory)][ValidateSet('plain', 'html')][string]$Subtype, [Parameter(Mandatory)][string]$Text)
    @{
        Headers = [ordered]@{ 'Content-Type' = "text/$Subtype; charset=utf-8"; 'Content-Transfer-Encoding' = 'base64' }
        Content = ConvertTo-FxcBase64Lines ([Text.Encoding]::UTF8.GetBytes($Text))
    }
}

function New-FxcMimeAttachment {
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][string]$ContentType, [Parameter(Mandatory)][byte[]]$Bytes)
    $ascii = $Name -cmatch '^[\x20-\x7E]*$'
    $file = if ($ascii) { "filename=`"$Name`"" } else { "filename*=utf-8''" + [uri]::EscapeDataString($Name) }
    $nameParam = if ($ascii) { "name=`"$Name`"" } else { "name*=utf-8''" + [uri]::EscapeDataString($Name) }
    @{
        Headers = [ordered]@{ 'Content-Type' = "$ContentType; $nameParam"; 'Content-Transfer-Encoding' = 'base64'; 'Content-Disposition' = "attachment; $file" }
        Content = ConvertTo-FxcBase64Lines $Bytes
    }
}

function New-FxcMimeMultipart {
    param([Parameter(Mandatory)][string]$Subtype, [Parameter(Mandatory)][string]$Boundary, [Parameter(Mandatory)][object[]]$Parts)
    $sb = [Text.StringBuilder]::new()
    foreach ($p in $Parts) { [void]$sb.Append("--$Boundary`r`n").Append((ConvertTo-FxcMimeText $p)).Append("`r`n") }
    [void]$sb.Append("--$Boundary--")
    @{ Headers = [ordered]@{ 'Content-Type' = "multipart/$Subtype; boundary=`"$Boundary`"" }; Content = $sb.ToString() }
}

function ConvertTo-FxcMimeText {
    # An entity (Headers + Content) as MIME text, CRLF line ends.
    param([Parameter(Mandatory)]$Entity)
    $lines = foreach ($k in $Entity.Headers.Keys) { "${k}: $($Entity.Headers[$k])" }
    ($lines -join "`r`n") + "`r`n`r`n" + $Entity.Content
}

function New-FxcMimeMessageEntity {
    # A message (usable as a top-level message or as the content of a message/rfc822 attachment).
    param([Parameter(Mandatory)]$Headers, [Parameter(Mandatory)]$Body)
    $h = [ordered]@{}
    foreach ($k in $Headers.Keys) { $h[$k] = $Headers[$k] }
    $h['MIME-Version'] = '1.0'
    foreach ($k in $Body.Headers.Keys) { $h[$k] = $Body.Headers[$k] }
    @{ Headers = $h; Content = $Body.Content }
}

function New-FxcMimeRfc822 {
    # A message/rfc822 attachment (Exchange makes it an embedded message).
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)]$Message)
    @{
        Headers = [ordered]@{ 'Content-Type' = 'message/rfc822'; 'Content-Disposition' = "attachment; filename=`"$Name`"" }
        Content = ConvertTo-FxcMimeText $Message
    }
}

#endregion

#region The plan ------------------------------------------------------------------------------------------------------------------

function New-FxcLongText {
    param([Parameter(Mandatory)][string]$Unit, [Parameter(Mandatory)][int]$Length)
    $sb = [Text.StringBuilder]::new()
    while ($sb.Length -lt $Length) { [void]$sb.Append($Unit) }
    $sb.ToString(0, $Length)
}

function Get-FxcPlan {
    <#
    .SYNOPSIS
        The whole corpus as steps (see the header of this file). Pure: no network, no file.
    .PARAMETER Mailbox
        Address of the lab user (recipient of the received-style message).
    #>
    [CmdletBinding()]
    param([string]$Mailbox = 'lab.user@contoso.example')

    $self = New-FxcAddress 'Lab User' $Mailbox
    $alice = New-FxcAddress 'Alice Martin' 'alice.martin@contoso.example'
    $png = New-FxcPng
    $pdf = New-FxcPdfBytes
    $zip = New-FxcZipBytes
    $setId = $script:FxcPropSet
    $steps = [Collections.Generic.List[object]]::new()

    # ---- Mail -------------------------------------------------------------------------------------------------------------
    # 01: HTML body, inline image (contentId), external link
    $cid = 'fxc-logo@contoso.example'
    $steps.Add((New-FxcStep -Key 'fxc-01' -Kind Mail -Title 'HTML body, inline image (contentId), external link' -Subject '[fxc-01] HTML with inline image' `
        -Variants @([ordered]@{
            subject = '[fxc-01] HTML with inline image'
            body = [ordered]@{ contentType = 'html'; content = "<html><head><meta charset=`"utf-8`"></head><body><h1 style=`"color:#336699`">FastTransfer corpus</h1><p>Inline image: <img src=`"cid:$cid`" width=`"16`" height=`"16`" alt=`"logo`"> and a <a href=`"https://www.contoso.example/path?a=1&amp;b=2#frag`">link</a>.</p><table border=`"1`"><tr><td>caf`u{E9}</td><td>`u{65E5}`u{672C}`u{8A9E}</td></tr></table></body></html>" }
            toRecipients = @($alice, $self)
        }) `
        -Attachments @(@{ Name = 'logo.png (inline)'; Variants = @((New-FxcFileAttachment -Name 'logo.png' -ContentType 'image/png' -Bytes $png -ContentId $cid)) })))

    # 02: many recipients with Unicode display names
    $to = @(
        (New-FxcAddress "Zo`u{E9} Dubois" 'zoe.dubois@contoso.example'), (New-FxcAddress "J`u{FC}rgen M`u{FC}ller" 'juergen.mueller@fabrikam.example'),
        (New-FxcAddress "`u{65E5}`u{672C}`u{8A9E} Tanaka" 'tanaka@contoso.example'), (New-FxcAddress "Sam `u{1F600} Smith" 'sam.smith@fabrikam.example'),
        (New-FxcAddress "O'Brien, Pat (Finance)" 'pat.obrien@contoso.example'))
    $cc = @((New-FxcAddress "Mar`u{ED}a Garc`u{ED}a" 'maria.garcia@contoso.example'), (New-FxcAddress "Ole Nystr`u{F8}m" 'ole.nystrom@fabrikam.example'), (New-FxcAddress 'Plain Cc' 'plain.cc@contoso.example'))
    $bcc = @((New-FxcAddress "Bcc `u{C9}lodie" 'elodie.bcc@contoso.example'), (New-FxcAddress 'Bcc Two' 'bcc.two@fabrikam.example'))
    $steps.Add((New-FxcStep -Key 'fxc-02' -Kind Mail -Title '5 To, 3 Cc, 2 Bcc, Unicode display names' -Subject '[fxc-02] Many recipients' `
        -Variants @([ordered]@{
            subject = '[fxc-02] Many recipients'; body = [ordered]@{ contentType = 'text'; content = "Ten recipients with accented, CJK and emoji display names." }
            toRecipients = $to; ccRecipients = $cc; bccRecipients = $bcc
        })))

    # 03: embedded message (item attachment) holding a file attachment
    $inner = [ordered]@{
        '@odata.type' = 'microsoft.graph.message'; subject = 'Embedded message with a file'
        body = [ordered]@{ contentType = 'text'; content = 'This message is attached to the outer one.' }
        toRecipients = @($alice); from = (New-FxcAddress 'Bob Lefevre' 'bob.lefevre@fabrikam.example')
    }
    $innerWithFile = [ordered]@{}; foreach ($k in $inner.Keys) { $innerWithFile[$k] = $inner[$k] }
    $innerWithFile['attachments'] = @((New-FxcFileAttachment -Name 'inner.txt' -ContentType 'text/plain' -Bytes ([Text.Encoding]::UTF8.GetBytes('file inside the embedded message'))))
    $steps.Add((New-FxcStep -Key 'fxc-03' -Kind Mail -Title 'embedded message (itemAttachment) that holds a file attachment' -Subject '[fxc-03] Embedded message with file' `
        -Variants @([ordered]@{ subject = '[fxc-03] Embedded message with file'; body = [ordered]@{ contentType = 'text'; content = 'See the attached message.' }; toRecipients = @($alice) }) `
        -Attachments @(@{ Name = 'embedded message'; Variants = @(
            [ordered]@{ '@odata.type' = '#microsoft.graph.itemAttachment'; name = 'Embedded message with a file'; item = $innerWithFile },
            [ordered]@{ '@odata.type' = '#microsoft.graph.itemAttachment'; name = 'Embedded message (no nested file)'; item = $inner }) })))

    # 04: two levels of embedding
    $level2 = [ordered]@{
        '@odata.type' = 'microsoft.graph.message'; subject = 'Level 2 message'
        body = [ordered]@{ contentType = 'text'; content = 'Innermost message.' }; toRecipients = @($self)
    }
    $level1 = [ordered]@{
        '@odata.type' = 'microsoft.graph.message'; subject = 'Level 1 message'
        body = [ordered]@{ contentType = 'text'; content = 'Middle message, it holds another message.' }; toRecipients = @($alice)
        attachments = @([ordered]@{ '@odata.type' = '#microsoft.graph.itemAttachment'; name = 'Level 2 message'; item = $level2 })
    }
    $level1Flat = [ordered]@{}; foreach ($k in $level1.Keys) { if ($k -ne 'attachments') { $level1Flat[$k] = $level1[$k] } }
    $steps.Add((New-FxcStep -Key 'fxc-04' -Kind Mail -Title 'two-level embedding: message in message in message' -Subject '[fxc-04] Two-level embedded message' `
        -Variants @([ordered]@{ subject = '[fxc-04] Two-level embedded message'; body = [ordered]@{ contentType = 'text'; content = 'Outer message.' }; toRecipients = @($alice) }) `
        -Attachments @(@{ Name = 'level 1 message'; Variants = @(
            [ordered]@{ '@odata.type' = '#microsoft.graph.itemAttachment'; name = 'Level 1 message'; item = $level1 },
            [ordered]@{ '@odata.type' = '#microsoft.graph.itemAttachment'; name = 'Level 1 message (not nested)'; item = $level1Flat }) })))

    # 05: item attachments of an event and of a contact
    $evItem = [ordered]@{
        '@odata.type' = 'microsoft.graph.event'; subject = 'Attached event'
        start = (New-FxcDateTime '2026-12-01T09:00:00.0000000' 'UTC'); end = (New-FxcDateTime '2026-12-01T10:00:00.0000000' 'UTC')
        location = [ordered]@{ displayName = 'Room 1' }; body = [ordered]@{ contentType = 'text'; content = 'Event attached to a message.' }
    }
    $ctItem = [ordered]@{
        '@odata.type' = 'microsoft.graph.contact'; givenName = "Ren`u{E9}e"; surname = 'Attached'; companyName = 'Contoso'
        emailAddresses = @([ordered]@{ name = "Ren`u{E9}e Attached"; address = 'renee.attached@contoso.example' }); businessPhones = @('+33 1 23 45 67 89')
    }
    $steps.Add((New-FxcStep -Key 'fxc-05' -Kind Mail -Title 'item attachments: an event and a contact' -Subject '[fxc-05] Event and contact attached' `
        -Variants @([ordered]@{ subject = '[fxc-05] Event and contact attached'; body = [ordered]@{ contentType = 'text'; content = 'An event and a contact are attached.' }; toRecipients = @($alice) }) `
        -Attachments @(
            @{ Name = 'attached event'; Variants = @([ordered]@{ '@odata.type' = '#microsoft.graph.itemAttachment'; name = 'Attached event'; item = $evItem }) },
            @{ Name = 'attached contact'; Variants = @([ordered]@{ '@odata.type' = '#microsoft.graph.itemAttachment'; name = 'Attached contact'; item = $ctItem }) })))

    # 06: several file attachments of various types
    $steps.Add((New-FxcStep -Key 'fxc-06' -Kind Mail -Title 'file attachments: txt, png, pdf-like, zip, Unicode name, 64 KiB binary' -Subject '[fxc-06] Various file attachments' `
        -Variants @([ordered]@{ subject = '[fxc-06] Various file attachments'; body = [ordered]@{ contentType = 'text'; content = 'Six attachments.' }; toRecipients = @($alice) }) `
        -Attachments @(
            @{ Name = 'notes.txt'; Variants = @((New-FxcFileAttachment -Name 'notes.txt' -ContentType 'text/plain' -Bytes ([Text.Encoding]::UTF8.GetBytes("line 1`r`nline 2 caf`u{E9}`r`n")))) },
            @{ Name = 'pixel.png'; Variants = @((New-FxcFileAttachment -Name 'pixel.png' -ContentType 'image/png' -Bytes $png)) },
            @{ Name = 'doc.pdf'; Variants = @((New-FxcFileAttachment -Name 'doc.pdf' -ContentType 'application/pdf' -Bytes $pdf)) },
            @{ Name = 'bundle.zip'; Variants = @((New-FxcFileAttachment -Name 'bundle.zip' -ContentType 'application/zip' -Bytes $zip)) },
            @{ Name = 'Unicode name'; Variants = @((New-FxcFileAttachment -Name "r`u{E9}sum`u{E9}-`u{65E5}`u{672C}`u{8A9E}-`u{1F600}.txt" -ContentType 'text/plain' -Bytes ([Text.Encoding]::UTF8.GetBytes('Unicode file name')))) },
            @{ Name = 'random.bin'; Variants = @((New-FxcFileAttachment -Name 'random.bin' -ContentType 'application/octet-stream' -Bytes (New-FxcPseudoBytes -Seed 6 -Length 65536))) })))

    # 07: categories, importance, flag, read state, sensitivity, receipts.
    # Graph v1.0: "sensitivity" is NOT a property of the message resource (it belongs to event and others), so a body that carries it is
    # refused with "UnableToDeserializePostBody". The sensitivity of a message is the MAPI property PidTagSensitivity (0x0036, Integer:
    # 0 normal, 1 personal, 2 private, 3 confidential), set after the creation like the other extended properties (see -Ext below).
    # flag is a followupFlag: flagStatus (notFlagged | complete | flagged), startDateTime / dueDateTime as dateTimeTimeZone.
    $steps.Add((New-FxcStep -Key 'fxc-07' -Kind Mail -Title 'categories, importance high, follow-up flag, read, sensitivity (MAPI), receipts' -Subject '[fxc-07] Categories importance flag' `
        -Variants @([ordered]@{
            subject = '[fxc-07] Categories importance flag'; body = [ordered]@{ contentType = 'text'; content = 'Flagged, categorised, read.' }; toRecipients = @($alice)
            categories = @('Red category', "Caf`u{E9} projet"); importance = 'high'; isRead = $true
            isReadReceiptRequested = $true; isDeliveryReceiptRequested = $true; replyTo = @((New-FxcAddress 'Reply Here' 'reply.here@contoso.example'))
            flag = [ordered]@{ flagStatus = 'flagged'; startDateTime = (New-FxcDateTime '2026-11-02T08:00:00.0000000' 'UTC'); dueDateTime = (New-FxcDateTime '2026-11-06T17:00:00.0000000' 'UTC') }
        }, [ordered]@{
            subject = '[fxc-07] Categories importance flag'; body = [ordered]@{ contentType = 'text'; content = 'Flagged, categorised, read.' }; toRecipients = @($alice)
            categories = @('Red category', "Caf`u{E9} projet"); importance = 'high'; isRead = $true
            isReadReceiptRequested = $true; isDeliveryReceiptRequested = $true
        }, [ordered]@{
            subject = '[fxc-07] Categories importance flag'; body = [ordered]@{ contentType = 'text'; content = 'Flagged, categorised, read.' }; toRecipients = @($alice)
            categories = @('Red category', "Caf`u{E9} projet"); importance = 'high'
        }) `
        -Ext @((New-FxcExt 'Integer 0x0036' @('3')))))

    # 08: single-valued extended properties, every type, named by name and by id
    $ext1 = @(
        (New-FxcExt "Integer $setId Name FxcInteger" @('-123456'))
        (New-FxcExt "Long $setId Name FxcLong" @('9007199254740993'))
        (New-FxcExt "Double $setId Name FxcDouble" @('3.14159265358979'))
        (New-FxcExt "Currency $setId Name FxcCurrency" @('1234.56', '12345600'))
        (New-FxcExt "Boolean $setId Name FxcBoolean" @('true'))
        (New-FxcExt "SystemTime $setId Name FxcSystemTime" @('2026-03-01T10:20:30Z'))
        (New-FxcExt "Binary $setId Name FxcBinary" @([Convert]::ToBase64String([byte[]](0, 1, 2, 253, 254, 255) + [Text.Encoding]::ASCII.GetBytes('binary'))))
        (New-FxcExt "String $setId Name FxcString" @("Unicode: caf`u{E9} `u{FC}ber `u{65E5}`u{672C}`u{8A9E} `u{1F600}"))
        (New-FxcExt "CLSID $setId Name FxcClsid" @('00020329-0000-0000-C000-000000000046', '{00020329-0000-0000-C000-000000000046}'))
        (New-FxcExt "Float $setId Name FxcFloat" @('2.5'))
        (New-FxcExt "Short $setId Name FxcShort" @('-12'))
        (New-FxcExt "String $setId Id 0x8F10" @('named by id (string)'))
        (New-FxcExt "Integer $setId Id 0x8F11" @('42'))
        (New-FxcExt "Double $setId Id 0x8F12" @('0.5'))
        (New-FxcExt 'String {00020329-0000-0000-C000-000000000046} Name Keywords' @('well known GUID, by name'))
    )
    $steps.Add((New-FxcStep -Key 'fxc-08' -Kind Mail -Title 'single-valued extended properties of every type (named by name and by id)' -Subject '[fxc-08] Single-valued extended properties' `
        -Variants @([ordered]@{ subject = '[fxc-08] Single-valued extended properties'; body = [ordered]@{ contentType = 'text'; content = 'Extended properties of many types.' }; toRecipients = @($alice) }) -Ext $ext1))

    # 09: multi-valued extended properties
    $b64a = [Convert]::ToBase64String([byte[]](1, 2, 3)); $b64b = [Convert]::ToBase64String([byte[]](255, 0, 128, 7))
    $ext9 = @(
        (New-FxcExt -Multi "StringArray $setId Name FxcStringArray" @(, @('alpha', "Zo`u{E9}", "`u{65E5}`u{672C}`u{8A9E}", '')))
        (New-FxcExt -Multi "IntegerArray $setId Name FxcIntegerArray" @(, @('1', '-2', '3')))
        (New-FxcExt -Multi "LongArray $setId Name FxcLongArray" @(, @('9007199254740993', '-5')))
        (New-FxcExt -Multi "DoubleArray $setId Name FxcDoubleArray" @(, @('1.5', '2.25')))
        (New-FxcExt -Multi "SystemTimeArray $setId Name FxcSystemTimeArray" @(, @('2026-01-01T00:00:00Z', '2026-06-30T12:30:00Z')))
        (New-FxcExt -Multi "BinaryArray $setId Name FxcBinaryArray" @(, @($b64a, $b64b)))
        (New-FxcExt -Multi "CLSIDArray $setId Name FxcClsidArray" @(, @('00020329-0000-0000-C000-000000000046', '00062008-0000-0000-C000-000000000046')))
        (New-FxcExt -Multi "CurrencyArray $setId Name FxcCurrencyArray" @(@('1.25', '2.5'), @('12500', '25000')))
        (New-FxcExt -Multi "FloatArray $setId Name FxcFloatArray" @(, @('0.25', '0.5')))
        (New-FxcExt -Multi "ShortArray $setId Name FxcShortArray" @(, @('1', '2')))
    )
    $steps.Add((New-FxcStep -Key 'fxc-09' -Kind Mail -Title 'multi-valued extended properties of every type' -Subject '[fxc-09] Multi-valued extended properties' `
        -Variants @([ordered]@{ subject = '[fxc-09] Multi-valued extended properties'; body = [ordered]@{ contentType = 'text'; content = 'Multi-valued extended properties.' }; toRecipients = @($alice) }) -Ext $ext9))

    # 10: custom X- headers, long lines, many non-ASCII characters
    $cjk = -join (0..199 | ForEach-Object { [char](0x4E00 + $_ * 7) })
    $bodyText = (@(
        'Line 1: a very long line without any space follows.', (New-FxcLongText 'abcdefghij' 5000),
        'Line 3: a long CJK line follows.', ($cjk * 15),
        "Line 5: accents - `u{E0}`u{E2}`u{E4}`u{E7}`u{E8}`u{E9}`u{EA}`u{EB}`u{EE}`u{EF}`u{F4}`u{F6}`u{F9}`u{FB}`u{FC}`u{FF}`u{153}`u{20AC} - Greek `u{3B1}`u{3B2}`u{3B3} - Cyrillic `u{43F}`u{440}`u{438}`u{432}`u{435}`u{442} - emoji `u{1F600}`u{1F680}`u{1F9EA}",
        "Line 6: tabs`t`tand  double  spaces, trailing spaces   ", '', 'Line 8 after an empty line.') -join "`r`n")
    $hdr = @([ordered]@{ name = 'x-fxc-corpus'; value = 'fxc-10' }, [ordered]@{ name = 'x-fxc-long'; value = (New-FxcLongText 'headervalue-' 400) },
        [ordered]@{ name = 'x-fxc-spaces'; value = 'one two  three' }, [ordered]@{ name = 'x-fxc-empty-ish'; value = '-' })
    $m10 = [ordered]@{ subject = '[fxc-10] Headers, long lines, non-ASCII'; body = [ordered]@{ contentType = 'text'; content = $bodyText }; toRecipients = @($alice) }
    $m10h = [ordered]@{}; foreach ($k in $m10.Keys) { $m10h[$k] = $m10[$k] }; $m10h['internetMessageHeaders'] = $hdr
    $steps.Add((New-FxcStep -Key 'fxc-10' -Kind Mail -Title 'custom X- headers, 5000-character line, CJK / accents / emoji body' -Subject '[fxc-10] Headers, long lines, non-ASCII' -Variants @($m10h, $m10)))

    # 11: a very long subject (300 characters, then 255)
    $subj300 = '[fxc-11] ' + (New-FxcLongText "very long subject with accents `u{E9}`u{FC} and `u{65E5}`u{672C}`u{8A9E} " (300 - 9))
    $subj255 = $subj300.Substring(0, 255)
    $steps.Add((New-FxcStep -Key 'fxc-11' -Kind Mail -Title 'very long subject (300 characters; 255 as a fallback)' -Subject $subj300 -ReportSubject `
        -Variants @([ordered]@{ subject = $subj300; body = [ordered]@{ contentType = 'text'; content = 'Long subject.' }; toRecipients = @($alice) },
            [ordered]@{ subject = $subj255; body = [ordered]@{ contentType = 'text'; content = 'Long subject (255).' }; toRecipients = @($alice) })))

    # 12: empty body, no recipient
    $steps.Add((New-FxcStep -Key 'fxc-12' -Kind Mail -Title 'empty body, no recipient' -Subject '[fxc-12] Empty body' `
        -Variants @([ordered]@{ subject = '[fxc-12] Empty body'; body = [ordered]@{ contentType = 'text'; content = '' } })))

    # 13: only an attachment (empty body)
    $steps.Add((New-FxcStep -Key 'fxc-13' -Kind Mail -Title 'only an attachment, empty HTML body' -Subject '[fxc-13] Attachment only' `
        -Variants @([ordered]@{ subject = '[fxc-13] Attachment only'; body = [ordered]@{ contentType = 'html'; content = '' }; toRecipients = @($alice) }) `
        -Attachments @(@{ Name = 'only.zip'; Variants = @((New-FxcFileAttachment -Name 'only.zip' -ContentType 'application/zip' -Bytes $zip)) })))

    # 14: a message that looks received: from, delivery time, flags (not a draft), reply verb
    $recvExt = @(
        [ordered]@{ id = 'Integer 0x0E07'; value = '1' }, [ordered]@{ id = 'SystemTime 0x0E06'; value = '2026-10-01T09:15:00Z' }, [ordered]@{ id = 'SystemTime 0x0039'; value = '2026-10-01T09:14:50Z' })
    $m14 = [ordered]@{
        subject = '[fxc-14] Received-style message'; from = (New-FxcAddress 'Bob Lefevre' 'bob.lefevre@fabrikam.example'); toRecipients = @($self)
        body = [ordered]@{ contentType = 'html'; content = '<html><body><p>Looks received: <b>not</b> a draft.</p></body></html>' }
    }
    $m14e = [ordered]@{}; foreach ($k in $m14.Keys) { $m14e[$k] = $m14[$k] }; $m14e['singleValueExtendedProperties'] = $recvExt
    $m14s = [ordered]@{ subject = $m14.subject; body = $m14.body; toRecipients = @($self) }
    $steps.Add((New-FxcStep -Key 'fxc-14' -Kind Mail -Title 'received-style message (from, delivery time, message flags), reply verb' -Subject '[fxc-14] Received-style message' `
        -Variants @($m14e, $m14, $m14s) `
        -Ext @((New-FxcExt 'Integer 0x1080' @('261')), (New-FxcExt 'Integer 0x1081' @('102')), (New-FxcExt 'SystemTime 0x1082' @('2026-10-02T08:00:00Z')))))

    # 15: created from MIME: alternative bodies, Bcc, embedded rfc822 holding a PNG and another rfc822 (2 levels)
    $steps.Add((New-FxcStep -Key 'fxc-15' -Kind Mail -Title 'created from MIME: alternative bodies, embedded rfc822 in rfc822, RFC 2047 names' -Subject '[fxc-15] MIME import' `
        -Mime (New-FxcMimeCorpusMessage -Self $Mailbox -Png $png -Zip $zip)))

    # ---- Calendar ---------------------------------------------------------------------------------------------------------
    $attendees = @(
        [ordered]@{ emailAddress = [ordered]@{ name = 'Alice Martin'; address = 'alice.martin@contoso.example' }; type = 'required' },
        [ordered]@{ emailAddress = [ordered]@{ name = "J`u{FC}rgen M`u{FC}ller"; address = 'juergen.mueller@fabrikam.example' }; type = 'optional' },
        [ordered]@{ emailAddress = [ordered]@{ name = 'Room 12'; address = 'room12@contoso.example' }; type = 'resource' })
    # The event has attendees: its extended properties go in the creation (a later PATCH would update the attendees again).
    $e20 = [ordered]@{
        subject = '[fxc-20] Single event with attendees'
        body = [ordered]@{ contentType = 'html'; content = "<html><body><h2>Agenda</h2><ul><li>Caf`u{E9}</li><li><a href=`"https://www.contoso.example/agenda`">Link</a></li></ul></body></html>" }
        start = (New-FxcDateTime '2026-11-16T10:00:00.0000000' 'W. Europe Standard Time'); end = (New-FxcDateTime '2026-11-16T11:00:00.0000000' 'W. Europe Standard Time')
        location = [ordered]@{ displayName = "Salle Caf`u{E9} 1"; address = [ordered]@{ street = '1 Rue Exemple'; city = 'Paris'; countryOrRegion = 'France'; postalCode = '75001' } }
        attendees = $attendees; reminderMinutesBeforeStart = 30; importance = 'high'; showAs = 'busy'; categories = @('Blue category'); allowNewTimeProposals = $false
    }
    $e20x = [ordered]@{}; foreach ($k in $e20.Keys) { $e20x[$k] = $e20[$k] }
    $e20x['singleValueExtendedProperties'] = @([ordered]@{ id = "String $setId Name FxcEventString"; value = 'event property' }, [ordered]@{ id = "Integer $setId Id 0x8F20"; value = '7' })
    $e20x['multiValueExtendedProperties'] = @([ordered]@{ id = "StringArray $setId Name FxcEventArray"; value = @('a', 'b') })
    $steps.Add((New-FxcStep -Key 'fxc-20' -Kind Event -Title 'single event: attendees (required, optional, resource), location, HTML body, extended properties' -Subject '[fxc-20] Single event with attendees' `
        -Variants @($e20x, $e20, [ordered]@{
            subject = '[fxc-20] Single event with attendees'
            body = [ordered]@{ contentType = 'html'; content = '<html><body><p>Agenda</p></body></html>' }
            start = (New-FxcDateTime '2026-11-16T10:00:00.0000000' 'UTC'); end = (New-FxcDateTime '2026-11-16T11:00:00.0000000' 'UTC')
            location = [ordered]@{ displayName = 'Room 1' }; attendees = $attendees
        })))

    $steps.Add((New-FxcStep -Key 'fxc-21' -Kind Event -Title 'all-day event' -Subject '[fxc-21] All-day event' `
        -Variants @([ordered]@{
            subject = '[fxc-21] All-day event'; isAllDay = $true; showAs = 'free'
            start = (New-FxcDateTime '2026-11-18T00:00:00.0000000' 'UTC'); end = (New-FxcDateTime '2026-11-19T00:00:00.0000000' 'UTC')
            body = [ordered]@{ contentType = 'text'; content = 'All day.' }
        })))

    $recurring = [ordered]@{
        subject = '[fxc-22] Weekly recurring event'; body = [ordered]@{ contentType = 'text'; content = 'Every Monday, 6 times.' }
        start = (New-FxcDateTime '2026-11-09T14:00:00.0000000' 'Romance Standard Time'); end = (New-FxcDateTime '2026-11-09T14:30:00.0000000' 'Romance Standard Time')
        location = [ordered]@{ displayName = 'Room 2' }
        recurrence = [ordered]@{
            pattern = [ordered]@{ type = 'weekly'; interval = 1; daysOfWeek = @('monday'); firstDayOfWeek = 'sunday' }
            range = [ordered]@{ type = 'numbered'; startDate = '2026-11-09'; numberOfOccurrences = 6 }
        }
    }
    $steps.Add((New-FxcStep -Key 'fxc-22' -Kind Event -Title 'weekly recurring event, one modified and one cancelled occurrence' -Subject '[fxc-22] Weekly recurring event' -Variants @($recurring) -Special 'RecurringExceptions'))

    $steps.Add((New-FxcStep -Key 'fxc-23' -Kind Event -Title 'event with file attachments' -Subject '[fxc-23] Event with attachments' `
        -Variants @([ordered]@{
            subject = '[fxc-23] Event with attachments'; body = [ordered]@{ contentType = 'text'; content = 'Two attachments.' }
            start = (New-FxcDateTime '2026-11-20T09:00:00.0000000' 'UTC'); end = (New-FxcDateTime '2026-11-20T09:30:00.0000000' 'UTC')
        }) `
        -Attachments @(
            @{ Name = 'agenda.txt'; Variants = @((New-FxcFileAttachment -Name 'agenda.txt' -ContentType 'text/plain' -Bytes ([Text.Encoding]::UTF8.GetBytes("agenda caf`u{E9}")))) },
            @{ Name = 'map.png'; Variants = @((New-FxcFileAttachment -Name 'map.png' -ContentType 'image/png' -Bytes $png)) })))

    $steps.Add((New-FxcStep -Key 'fxc-24' -Kind Event -Title 'private event, no attendees, no online meeting, tentative, no reminder' -Subject '[fxc-24] Private event' `
        -Variants @([ordered]@{
            subject = '[fxc-24] Private event'; sensitivity = 'private'; showAs = 'tentative'; isOnlineMeeting = $false; isReminderOn = $false; responseRequested = $false
            body = [ordered]@{ contentType = 'text'; content = 'Private.' }; categories = @('Green category')
            start = (New-FxcDateTime '2026-11-23T16:00:00.0000000' 'UTC'); end = (New-FxcDateTime '2026-11-23T17:00:00.0000000' 'UTC')
        })))

    # ---- Contacts ---------------------------------------------------------------------------------------------------------
    $full = [ordered]@{
        displayName = "[fxc-30] Full Contact Ren`u{E9}e M`u{FC}ller-`u{65E5}`u{672C}"; givenName = "Ren`u{E9}e"; middleName = 'Marie'; surname = "M`u{FC}ller-`u{65E5}`u{672C}"; nickName = 'Reny'
        title = 'Dr.'; generation = 'Jr.'; initials = 'RMM'; fileAs = 'Full, Contact'; companyName = 'Contoso Ltd'; department = 'Research'; jobTitle = "Directrice de l'innovation"
        officeLocation = 'Building 4'; profession = 'Engineer'; assistantName = 'Assistant Name'; manager = 'Manager Name'; spouseName = 'Spouse Name'; children = @('Child One', "Child Tw`u{F6}")
        yomiGivenName = "`u{30EB}`u{30CD}"; yomiSurname = "`u{30DF}`u{30E5}`u{30E9}`u{30FC}"; yomiCompanyName = "`u{30B3}`u{30F3}`u{30C8}`u{30BD}"
        birthday = '1984-02-29'; personalNotes = "Notes with accents: caf`u{E9}, na`u{EF}ve, `u{65E5}`u{672C}`u{8A9E}, emoji `u{1F600}.`r`nSecond line."
        businessHomePage = 'https://www.contoso.example/people/renee'; imAddresses = @('renee@im.contoso.example', 'renee@im.fabrikam.example')
        emailAddresses = @(
            [ordered]@{ name = "Ren`u{E9}e (work)"; address = 'renee.mueller@contoso.example' }, [ordered]@{ name = 'Renee (home)'; address = 'renee.home@fabrikam.example' },
            [ordered]@{ name = 'Renee (other)'; address = 'renee.other@contoso.example' })
        businessPhones = @('+33 1 23 45 67 89', '+33 1 98 76 54 32'); homePhones = @('+33 4 11 22 33 44', '+33 4 55 66 77 88'); mobilePhone = '+33 6 12 34 56 78'
        homeAddress = [ordered]@{ street = "12 Rue de l'Exemple"; city = 'Lyon'; state = "Rh`u{F4}ne"; countryOrRegion = 'France'; postalCode = '69001' }
        businessAddress = [ordered]@{ street = '1 Contoso Way'; city = 'Paris'; state = 'IDF'; countryOrRegion = 'France'; postalCode = '75008' }
        otherAddress = [ordered]@{ street = 'Other Street 5'; city = "Z`u{FC}rich"; countryOrRegion = 'Switzerland'; postalCode = '8001' }
        categories = @('Red category', 'Blue category')
    }
    $fullPlain = [ordered]@{}; foreach ($k in $full.Keys) { if ($k -notin 'yomiGivenName', 'yomiSurname', 'yomiCompanyName', 'fileAs') { $fullPlain[$k] = $full[$k] } }
    $steps.Add((New-FxcStep -Key 'fxc-30' -Kind Contact -Title 'full contact: many fields, 3 e-mails, phones, 3 addresses, birthday, categories, Unicode' -Subject $full.displayName `
        -Variants @($full, $fullPlain) -Ext @((New-FxcExt "String $setId Name FxcContactString" @('contact property')), (New-FxcExt "Integer $setId Id 0x8F30" @('9')))))

    $steps.Add((New-FxcStep -Key 'fxc-31' -Kind Contact -Title 'minimal contact' -Subject '[fxc-31] Minimal' `
        -Variants @([ordered]@{ displayName = '[fxc-31] Minimal'; givenName = 'Minimal' })))

    $jpeg = New-FxcJpeg
    $steps.Add((New-FxcStep -Key 'fxc-32' -Kind Contact -Title $(if ($jpeg) { 'contact with a photo (JPEG)' } else { 'contact (no photo: System.Drawing is not available)' }) -Subject '[fxc-32] Contact with photo' `
        -Variants @([ordered]@{ displayName = '[fxc-32] Contact with photo'; givenName = 'Photo'; surname = 'Contact'; emailAddresses = @([ordered]@{ name = 'Photo Contact'; address = 'photo.contact@contoso.example' }) }) -Photo $jpeg))

    , $steps.ToArray()
}

function New-FxcMimeCorpusMessage {
    # fxc-15: multipart/mixed { multipart/alternative, rfc822 { text, png, rfc822 { text } }, zip }, Bcc, RFC 2047 names.
    param([Parameter(Mandatory)][string]$Self, [Parameter(Mandatory)][byte[]]$Png, [Parameter(Mandatory)][byte[]]$Zip)
    $deepest = New-FxcMimeMessageEntity -Headers ([ordered]@{
        From = (Format-FxcMimeAddress 'Deep Sender' 'deep@fabrikam.example'); To = (Format-FxcMimeAddress 'Lab User' $Self)
        Subject = (ConvertTo-FxcMimeWord "Innermost message `u{65E5}`u{672C}`u{8A9E}"); Date = 'Mon, 05 Oct 2026 08:00:00 +0000'; 'Message-ID' = '<fxc-15-deepest@fabrikam.example>' }) `
        -Body (New-FxcMimeText plain "The innermost message (level 2).`r`n")
    $middle = New-FxcMimeMessageEntity -Headers ([ordered]@{
        From = (Format-FxcMimeAddress 'Middle Sender' 'middle@contoso.example'); To = (Format-FxcMimeAddress 'Lab User' $Self)
        Subject = 'Middle message'; Date = 'Mon, 05 Oct 2026 08:30:00 +0000'; 'Message-ID' = '<fxc-15-middle@contoso.example>' }) `
        -Body (New-FxcMimeMultipart mixed '=_fxc15_mid' @(
            (New-FxcMimeText plain "The middle message (level 1).`r`n"),
            (New-FxcMimeAttachment -Name 'pixel.png' -ContentType 'image/png' -Bytes $Png),
            (New-FxcMimeRfc822 -Name 'deepest.eml' -Message $deepest)))
    $alt = New-FxcMimeMultipart alternative '=_fxc15_alt' @(
        (New-FxcMimeText plain "Plain alternative: caf`u{E9} `u{65E5}`u{672C}`u{8A9E}`r`n"),
        (New-FxcMimeText html "<html><body><p>HTML alternative: <b>caf`u{E9}</b> `u{65E5}`u{672C}`u{8A9E}</p></body></html>"))
    $outer = New-FxcMimeMessageEntity -Headers ([ordered]@{
        From = (Format-FxcMimeAddress 'Lab Sender' 'sender@fabrikam.example')
        To = ((Format-FxcMimeAddress "Zo`u{E9} Dubois" 'zoe.dubois@contoso.example'), (Format-FxcMimeAddress 'Lab User' $Self) -join ', ')
        Cc = (Format-FxcMimeAddress "J`u{FC}rgen M`u{FC}ller" 'juergen.mueller@fabrikam.example'); Bcc = (Format-FxcMimeAddress 'Hidden Bcc' 'hidden.bcc@contoso.example')
        Subject = '[fxc-15] MIME import: embedded messages and encoded words ' + (ConvertTo-FxcMimeWord "`u{E9}`u{FC}`u{65E5}`u{672C}`u{8A9E}")
        Date = 'Mon, 05 Oct 2026 09:00:00 +0000'; 'Message-ID' = '<fxc-15@contoso.example>'; 'X-Fxc-Corpus' = 'fxc-15'; Importance = 'high' }) `
        -Body (New-FxcMimeMultipart mixed '=_fxc15_outer' @(
            $alt,
            (New-FxcMimeRfc822 -Name 'middle.eml' -Message $middle),
            (New-FxcMimeAttachment -Name "archive-`u{E9}.zip" -ContentType 'application/zip' -Bytes $Zip)))
    ConvertTo-FxcMimeText $outer
}

#endregion

function Get-FxcPlanSummary {
    # One row per step for -WhatIf and the tests.
    param([Parameter(Mandatory)][object[]]$Plan)
    foreach ($s in $Plan) {
        $attBytes = 0
        foreach ($a in $s.Attachments) { $v = $a.Variants[0]; if ($v.Contains('contentBytes')) { $attBytes += [Math]::Floor($v['contentBytes'].Length * 3 / 4) } }
        $json = if ($s.Variants.Count) { ($s.Variants[0] | ConvertTo-Json -Depth 20 -Compress).Length } elseif ($s.Mime) { $s.Mime.Length } else { 0 }
        [pscustomobject]@{
            Key = $s.Key; Kind = $s.Kind; Title = $s.Title; Via = $(if ($s.Mime) { 'MIME' } else { 'JSON' }); BodyVariants = $s.Variants.Count
            Attachments = $s.Attachments.Count; ExtProps = $s.Ext.Count; PayloadChars = $json; AttachmentBytes = [long]$attBytes
            Extra = (@($(if ($s.Photo) { 'photo' }), $s.Special) | Where-Object { $_ }) -join ','
        }
    }
}

function Get-FxcMarkerFromSubject {
    # "[fxc-07] Something" -> "fxc-07", else $null.
    param([string]$Subject)
    if ($Subject -match '^\[(fxc-\d{2})\]') { return $Matches[1] }
    $null
}

function Get-FxcExtBody {
    # The PATCH body that sets one extended property with one of its value variants.
    param([Parameter(Mandatory)]$Ext, [Parameter(Mandatory)][int]$Variant)
    $v = $Ext.Variants[$Variant]
    if ($Ext.Multi) { return @{ multiValueExtendedProperties = @([ordered]@{ id = $Ext.Id; value = @([string[]]$v) }) } }
    @{ singleValueExtendedProperties = @([ordered]@{ id = $Ext.Id; value = [string]$v }) }
}

#region End of the run -----------------------------------------------------------------------------------------------------------

function New-FxcFolderReport {
    # What the end of the run says about one folder (and the manifest records): its Graph id and name, and the path to give to the export
    # when Graph told it. A path is never guessed: when Graph could not resolve it, ExportPath is $null and HowToFindPath says how to read it.
    # A secondary calendar is not a mail folder for Graph (GET me/mailFolders/{id} does not answer for it) and Graph v1.0 does not tell
    # its parent, so for it the answer is normally "not resolved".
    param([Parameter(Mandatory)][string]$Kind, [Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)][string]$Name, [string]$ExportPath)
    $resolved = -not [string]::IsNullOrEmpty($ExportPath)
    $how = if ($resolved) { $null } else {
        "Graph does not tell the path of this $($Kind.ToLowerInvariant()) folder. List the folder tree of the export API with " +
        ".\tools\Invoke-Phase0.ps1 -DeviceCode (read only; or the tree of the store with -Mode Status after an export), look for the folder named " +
        "'$Name' (its type: " + $(switch ($Kind) { 'Event' { 'IPF.Appointment' } 'Contact' { 'IPF.Contact' } default { 'IPF.Note' } }) +
        ") and give the path as shown, with / between the levels, to Invoke-MailboxTransfer.ps1 -Mode Export -Folder. The Graph id (below) is not the id of the export API: look for the name."
    }
    [ordered]@{ Kind = $Kind; Name = $Name; GraphId = $Id; ExportPath = $(if ($resolved) { $ExportPath }); PathResolved = $resolved; HowToFindPath = $how }
}

function New-FxcManifest {
    # The manifest of a run (a plain ordered hashtable, ready for ConvertTo-Json). Reads $Results / $Folders / $Paths, changes none.
    param(
        [Parameter(Mandatory)][string]$User, [Parameter(Mandatory)][string]$Root, [AllowEmptyCollection()][object[]]$Results = @(),
        [System.Collections.IDictionary]$Folders = @{}, [System.Collections.IDictionary]$Paths = @{}, [System.Collections.IDictionary]$FolderNames = @{}, [string]$RunError
    )
    $reports = [ordered]@{}
    foreach ($kind in @($Folders.Keys)) {
        $reports[$kind] = New-FxcFolderReport -Kind $kind -Id $Folders[$kind] -Name $FolderNames[$kind] -ExportPath $(if ($Paths.Contains($kind)) { $Paths[$kind] })
    }
    $exportFolders = [ordered]@{}
    foreach ($kind in @($reports.Keys)) { $exportFolders[$kind] = $reports[$kind].ExportPath }
    [ordered]@{
        CreatedUtc = [datetime]::UtcNow.ToString('o'); User = $User; Root = $Root
        Error = $(if ($RunError) { $RunError })
        FolderIds = $Folders; ExportFolders = $exportFolders; Folders = $reports
        Items = @($Results | ForEach-Object { [ordered]@{ Key = $_.Key; Kind = $_.Kind; Status = $_.Status; Id = $_.Id; Title = $_.Title; Notes = @($_.Notes) } })
    }
}

#endregion
