#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - the fake engine (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    Synthetic data and canned answers (contoso.com names only) behind the same interface as the real engine adapter
    (MailboxTransfer.Gui.Engine.ps1). It serves the documentation scenarios (-DocScenario), the screenshots and the headless tests: no network, no
    mailbox, no file outside the work folder. Every member answers at once through its OnDone callback.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

$script:FakePstState = 'damaged'

function New-MbtGuiFakeRandom {
    <# A small deterministic generator (the mockup's), so the screenshots are the same every time. #>
    param([string]$Seed)
    $a = 0
    foreach ($c in $Seed.ToCharArray()) { $a = [int](([long]$a * 31 + [int]$c) -band 0x7FFFFFFF) }
    $state = @{ A = $a }
    return { $state.A = [int](([long]$state.A * 1103515245 + 12345) -band 0x7FFFFFFF); return ($state.A / 2147483648.0) }.GetNewClosure()
}

function New-MbtGuiFakeNode {
    param([string]$Id, [string]$Name, [string]$Glyph, [long]$Count, [double]$SizeGiB, [object[]]$Kids = @(), [string]$Kind, [switch]$Root, [string]$Sub, [switch]$HeadersOnly, [switch]$Dim)
    $size = if ($SizeGiB -lt 0) { -1L } else { [long]($SizeGiB * 1GB) }
    $n = [MailboxTransferGui.FolderNode]::new($Id, $Name, $Count, $size)
    $n.Glyph = Get-MbtGlyph $Glyph
    $n.IsRoot = [bool]$Root; $n.Sub = $Sub; $n.Kind = $Kind; $n.HeadersOnly = [bool]$HeadersOnly; $n.Dim = [bool]$Dim
    if ($Root) { $n.Glyph = Get-MbtGlyph $Glyph }
    $kidNodes = @($Kids | Where-Object { $_ })
    $n.Tag = @{ Fake = $true; Count = $Count; HeadersOnly = [bool]$HeadersOnly }
    if ($kidNodes.Count) { $n.SetChildren([MailboxTransferGui.FolderNode[]]$kidNodes) }
    return $n
}

function Get-MbtGuiFakeTree {
    <# The mailbox of the mockup: Primary, Archive, Auxiliary archive 1 (being filled by Microsoft). -Recoverable: with Recoverable Items (the box "Include deleted items"). #>
    param([switch]$Recoverable)
    $f = { param($id, $name, $glyph, $count, $gib, $kids = @()) New-MbtGuiFakeNode -Id $id -Name $name -Glyph $glyph -Count $count -SizeGiB $gib -Kids $kids }
    $deleted = @()
    if ($Recoverable) { $deleted = @(& $f 'rec' 'Recoverable Items' folder 0 0 @((& $f 'rec-del' 'Deletions' trash 24 0.12), (& $f 'rec-pur' 'Purges' trash 4 0.02), (& $f 'rec-ver' 'Versions' folder 0 0), (& $f 'rec-dis' 'DiscoveryHolds' folder 0 0))) }
    $primary = New-MbtGuiFakeNode -Id 'r-p' -Name 'Adele Vance - Primary mailbox' -Glyph person -Count 0 -SizeGiB 0 -Root -Kind primary -Sub 'adele.vance@contoso.com' -Kids @(
        (& $f 'inbox' 'Inbox' mail 1284 2.1 @((& $f 'projects' 'Projects' folder 598 1.4 @((& $f 'p-c' 'Contoso Rollout' folder 410 1.0), (& $f 'p-f' 'Fabrikam Pilot' folder 188 0.4)))))
        (& $f 'sent' 'Sent Items' send 842 1.2), (& $f 'drafts' 'Drafts' pencil 3 0.02), (& $f 'cal' 'Calendar' cal 214 0.05), (& $f 'contacts' 'Contacts' people 96 0.01), (& $f 'del' 'Deleted Items' trash 57 0.2)
        $deleted)
    $archive = New-MbtGuiFakeNode -Id 'r-a' -Name 'Archive' -Glyph archive -Count 0 -SizeGiB 0 -Root -Kind archive -Sub 'Auto-expanding archive' -Kids @(
        (& $f 'old' 'Old Mail' folder 9210 31.4 @((& $f 'y19' '2019' folder 3120 10.1), (& $f 'y20' '2020' folder 3380 11.7), (& $f 'y21' '2021' folder 2710 9.6)))
        (& $f 'fin' 'Finance' folder 860 3.9 @((& $f 'fin-b' 'Budgets' folder 212 0.9))))
    $aux = New-MbtGuiFakeNode -Id 'r-x' -Name 'Auxiliary archive 1 (3f9c2a10)' -Glyph archive -Count 0 -SizeGiB 0 -Root -Kind aux -HeadersOnly -Sub 'Being filled by Microsoft - headers only until exported' -Kids @(
        (& $f 'x-old' 'Old Mail' folder 278 -1 @((& $f 'x-18' '2018' folder 204 -1), (& $f 'x-17' '2017' folder 74 -1))))
    foreach ($n in $aux.Descendants()) { $n.Tag.HeadersOnly = $true }
    return @($primary, $archive, $aux)
}

function Get-MbtGuiFakePstTree {
    $f = { param($id, $name, $glyph, $count, $gib, $kids = @()) New-MbtGuiFakeNode -Id $id -Name $name -Glyph $glyph -Count $count -SizeGiB $gib -Kids $kids }
    $root = New-MbtGuiFakeNode -Id 'r-o' -Name 'archive-2019.pst' -Glyph pst -Count 0 -SizeGiB 0 -Root -Kind pst -Sub '4 parts - 47.3 GiB - read only' -Kids @(
        (& $f 'o-in' 'Inbox' mail 1180 6.2)
        (& $f 'o-proj' 'Projects' folder 0 0 @((& $f 'o-p1' 'Contoso Rollout' folder 4300 21), (& $f 'o-p2' 'Fabrikam Pilot' folder 780 4.1)))
        (& $f 'o-old' 'Old Mail' folder 0 0 @((& $f 'o-19' '2019' folder 3120 12)))
        (& $f 'o-fin' 'Finance' folder 0 0 @((& $f 'o-bud' 'Budgets' folder 480 4))))
    # a PST gives no folder size before the import, and the engine takes folders by their path below the top folder (-IncludeFolder)
    $walk = $null
    $walk = { param($node, $path) foreach ($k in @($node.Children)) { if ($k.Id -eq "`0placeholder") { continue }; $p = if ($path) { "$path/$($k.Name)" } else { $k.Name }; $k.Tag.FolderPath = $p; $k.Tag.Nid = [uint32]0x8000; $k.SizeBytes = -1; & $walk $k $p } }
    & $walk $root ''
    return @($root)
}

$script:FakePeople = @(@('Alex Wilber', 'alex.wilber@contoso.com'), @('Diego Siciliani', 'diego.siciliani@contoso.com'), @('Grady Archie', 'grady.archie@contoso.com'),
    @('Isaiah Langford', 'isaiah.langford@contoso.com'), @('Lee Gu', 'lee.gu@contoso.com'), @('Megan Bowen', 'megan.bowen@contoso.com'))
$script:FakeSubjects = @('Rollout plan: week 46', 'Invoice 2026-0412 for review', 'Re: Meeting notes and next steps', 'Budget forecast Q4', 'Signed contract (final)', 'Training schedule',
    'Re: Mailbox migration window', 'Quarterly report draft', 'Shared folder access request', 'Site visit photos', 'Updated price list', 'Re: Holiday coverage')
$script:FakeBodies = @("Hello Adele,`n`nPlease find the latest version attached. I have kept the same structure as last time so the comparison is easy.`n`nCould you confirm by Friday?`n`nThanks,",
    "Hi,`n`nThe window we agreed is Saturday evening. Nothing is needed from your side except keeping Outlook closed until Sunday morning.`n`nBest regards,",
    "Adele,`n`nA short reminder about the review. The numbers are in the table below; the open points are marked in yellow.`n`nRegards,")

function Get-MbtGuiFakeItems {
    <# Up to 12 synthetic messages of a folder (the mockup's), as rows of the list. #>
    param([Parameter(Mandatory)]$Node)
    $rand = New-MbtGuiFakeRandom $Node.Id
    $n = [math]::Min([int]$Node.Count, 12)
    $headers = $Node.Tag -is [hashtable] -and $Node.Tag.HeadersOnly
    $rows = for ($i = 0; $i -lt $n; $i++) {
        $p = $script:FakePeople[[int][math]::Floor((& $rand) * $script:FakePeople.Count)]
        $s = $script:FakeSubjects[[int][math]::Floor((& $rand) * $script:FakeSubjects.Count)]
        $d = ([datetime]'2026-11-03T12:00:00Z').AddDays(-$i * (1 + [math]::Floor((& $rand) * 4))).AddSeconds(-[math]::Floor((& $rand) * 40000))
        $size = [long]([math]::Floor(((& $rand) * (& $rand) * 900 + 8) * 1KB))
        $att = (& $rand) -lt 0.3
        $body = $script:FakeBodies[[int][math]::Floor((& $rand) * 3)]
        $row = [MailboxTransferGui.MessageRow]::new()
        $row.Id = "$($Node.Id)-$i"
        $row.From = if ($headers) { '(sender not available)' } else { $p[0] }
        $row.Address = $p[1]
        $row.Subject = if ($headers) { '(subject not available before export)' } else { $s }
        $row.SubjectNa = [bool]$headers
        $row.SizeBytes = $size; $row.SizeText = Format-MbtGuiBytes $size
        $row.DateText = $d.ToString('yyyy-MM-dd HH:mm')
        $row.HasAttachment = $att -and -not $headers
        $row.Tag = @{ Body = $body + "`n$($p[0])"; Name = $p[0]; Address = $p[1]; Subject = $s; Date = $d; Attachment = $att }
        $row
    }
    return @($rows)
}

function Get-MbtGuiFakePreview {
    param($Row, $Node)
    $t = $Row.Tag
    $root = $Node; while ($root.Parent) { $root = $root.Parent }
    [pscustomobject]@{
        Subject = $t.Subject; From = $t.Name; Address = $t.Address; To = 'Adele Vance'; Received = $t.Date.ToString('yyyy-MM-dd HH:mm') + ' UTC'
        Location = "$(($root.Name -replace '^.* - ', '')) / $($Node.Name)"; Details = ('IPM.Note - {0}{1}' -f $Row.SizeText, $(if ($t.Attachment) { ' - 1 attachment' } else { '' }))
        Attachment = $(if ($t.Attachment) { 'Specification-v3.pdf   ' + (Format-MbtGuiBytes ($Row.SizeBytes * 0.7)) } else { '' }); Body = $t.Body; Available = -not $root.HeadersOnly
    }
}

function New-MbtGuiFakeEngine {
    <#
        The engine adapter with synthetic data. Same members as New-MbtGuiRealEngine; Worker is Inline. Capability: PstImport, Eml and SharedMailbox are
        switched on by the scenarios that show them.
    #>
    param([hashtable]$Capability = @{})
    $cap = @{ Events = $true; PstExport = $true; PstCheck = $true; PstRepair = $true; PstImport = $false; EmlExport = $false; SharedMailbox = $false; FolderChoice = $false; PstCopy = $false; PstPreview = $false }
    foreach ($k in $Capability.Keys) { $cap[$k] = $Capability[$k] }
    $engine = [ordered]@{
        Name = 'Fake'
        Worker = [pscustomobject]@{ Inline = $true }
        Capability = $cap
        State = @{ Signed = $false }
    }
    $engine.Readiness = { [pscustomobject]@{ Lights = @(
                @{ Name = 'PowerShell 7.5 or later'; Detail = '7.5.4'; Ok = $true }, @{ Name = 'Engine ready'; Detail = 'compiled in 3.1 s'; Ok = $true },
                @{ Name = 'Sign-in component (MSAL)'; Detail = 'found in Microsoft.Graph.Authentication 2.25'; Ok = $true },
                @{ Name = 'Work folder'; Detail = 'D:\MailboxTransfer - 412 GiB free'; Ok = $true }, @{ Name = 'Microsoft 365 reachable'; Detail = 'login.microsoftonline.com answers'; Ok = $true }
            ) | ForEach-Object { $_ } } }
    $engine.SignInSettings = { [pscustomobject]@{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '00000000-0000-0000-0000-000000000000' } }
    $engine.SignIn = { param($Mailbox, $UseCode, $CodeFile, $OnDone, $OnError, $Silent, $TenantId, $ClientId, $App)
        $engine.State.Signed = $true
        & $OnDone ([pscustomobject]@{ Account = $Mailbox; TenantId = 'contoso.onmicrosoft.com'; Scopes = @('MailboxFolder.Read', 'MailboxItem.Read', 'MailboxItem.ImportExport', 'Mail.Read', 'MailboxFolder.ReadWrite'); Missing = @(); CanReadMail = $true; Access = $(if ($App) { 'Application (certificate, no user)' } else { 'Delegated (you, with your rights)' }) }) }.GetNewClosure()
    $engine.SignOut = { $engine.State.Signed = $false }.GetNewClosure()
    $engine.Locations = { param($Mailbox, $LocationsFile, $OnDone, $OnError)
        & $OnDone @(
            [pscustomobject]@{ Kind = 'Primary'; Label = 'Primary mailbox'; Hint = '3,496 items - 5.1 GiB'; Root = 'r-p'; MailboxId = 'MBX:primary' }
            [pscustomobject]@{ Kind = 'Archive'; Label = 'Archive'; Hint = '10,070 items - 35.3 GiB'; Root = 'r-a'; MailboxId = 'MBX:archive' }
            [pscustomobject]@{ Kind = 'AuxArchive'; Label = 'Auxiliary archive 1'; Hint = '278 items - size not given by Microsoft'; Root = 'r-x'; MailboxId = 'MBX:aux1'; Filling = $true }) }
    $engine.Roots = { param($Locations, $OnDone, $OnError, $IncludeRecoverable)
        $all = Get-MbtGuiFakeTree -Recoverable:([bool]$IncludeRecoverable)
        $keep = @($all | Where-Object { $r = $_; @($Locations | Where-Object { $_.Root -eq $r.Id }).Count -gt 0 })
        & $OnDone $keep }
    $engine.Children = { param($Node, $OnDone, $OnError) & $OnDone @() }
    $engine.Items = { param($Node, $Token, $OnDone, $OnError) & $OnDone ([pscustomobject]@{ Rows = (Get-MbtGuiFakeItems $Node); Next = $null; Total = $Node.Count; Note = $null }) }
    $engine.Preview = { param($Row, $Node, $OnDone, $OnError) & $OnDone (Get-MbtGuiFakePreview $Row $Node) }
    $engine.OpenSource = { param($Kind, $Path, $OnDone, $OnError)
        if ($Kind -eq 'pst') { & $OnDone ([pscustomobject]@{ Ok = $true; Level = 'ok'; Message = '**4 parts found** (archive-2019.pst ... -004) - 47.3 GiB - checked: **clean**. The check only reads.'; Items = 9860; Bytes = [long](47.3GB); Folders = 7; Parts = 4; Roots = (Get-MbtGuiFakePstTree) }) }
        else { & $OnDone ([pscustomobject]@{ Ok = $true; Level = 'ok'; Message = '**Store verified:** 1,208 of 1,208 items match their fingerprint.'; Items = 1208; Bytes = [long](6.2GB); Folders = 12; Parts = 0; Roots = (Get-MbtGuiFakePstTree) }) } }
    $engine.CheckPst = { param($Path, $OnDone, $OnError)
        $bad = $script:FakePstState -ne 'clean'
        & $OnDone ([pscustomobject]@{ Path = $Path; SizeBytes = [long](4.2GB); Elapsed = '1 min 5 s'; Errors = $(if ($bad) { 6 } else { 0 }); Warnings = $(if ($bad) { 5 } else { 0 }); ExitCode = $(if ($bad) { 1 } else { 0 }); Recommended = 'Rebuild'
                Findings = $(if ($bad) { @(
                        @{ Severity = 'Error'; Text = '3 messages exist, but the folder does not list them'; Code = 'MSG-NOT-LISTED'; Count = 3; Level = 'Rebuild' }
                        @{ Severity = 'Error'; Text = 'A folder points to a parent that is gone'; Code = 'FOLDER-PARENT-MISMATCH'; Count = 1; Level = 'Rebuild' }
                        @{ Severity = 'Error'; Text = 'A folder line points to a message that no longer exists'; Code = 'DANGLING-ROW'; Count = 2; Level = 'Recompute' }
                        @{ Severity = 'Warning'; Text = 'Folder counters are out of date'; Code = 'CONTENTS-CELL'; Count = 4; Level = 'Recompute' }
                        @{ Severity = 'Warning'; Text = 'A message size differs from the sum of its parts'; Code = 'MSG-SIZE-COMPUTED'; Count = 1; Level = 'Recompute' }) } else { @() }) }) }
    $engine.SaveSelection = { param($Root, $Entries, $Path, $OnDone, $OnError) & $OnDone $Path }
    $engine.Jobs = { param($Work) @(
            [pscustomobject]@{ Task = 'import'; Title = 'Import ' + [char]0x00B7 + ' archive-2019.pst ' + [char]0x2192 + ' Adele Vance (Archive)'; Sub = '5,120 of 9,860 items done ' + [char]0x00B7 + ' stopped yesterday 22:41 ' + [char]0x00B7 + ' 3 items to check'; State = 'run'; Button = 'Continue'; Folder = ''; Demo = { Set-MbtGuiScenario 'RunImportStopped' } }
            [pscustomobject]@{ Task = 'export'; Title = 'Export ' + [char]0x00B7 + ' Primary + Archive ' + [char]0x2192 + ' D:\Exports\adele'; Sub = 'Not finished: 2 folders are not ready at Microsoft yet. Nothing is lost.'; State = 'warn'; Button = 'Continue'; Folder = ''; Demo = { Set-MbtGuiScenario 'DoneExportPartial' } }
            [pscustomobject]@{ Task = 'export'; Title = 'Export ' + [char]0x00B7 + ' finished 2026-11-02 ' + [char]0x00B7 + ' 3 PST files'; Sub = 'D:\Exports\fabrikam ' + [char]0x00B7 + ' all parts checked'; State = 'ok'; Button = 'Open'; Folder = ''; Demo = { Set-MbtGuiScenario 'Open' } }) }
    $engine.FreeSpace ={ param($Path) 412GB }
    # what the check of the plan would count: the ticks, as the service counts them (an auxiliary archive and a PST give no size)
    $engine.Plan = { param($Request, $OnDone, $OnError)
        $sel = Get-MbtGuiTaskSelection
        $totals = if ($sel -and $sel.Items -gt 0) { [pscustomobject]@{ Items = [long]$sel.Items; Bytes = [long]$sel.Bytes; Folders = [int]$sel.Folders; UnknownSize = [int]$sel.UnknownSize; FoldersToCreate = 0 } } else { $null }
        & $OnDone ([pscustomobject]@{ Ok = $true; Message = ''; Totals = $totals }) }
    return [pscustomobject]$engine
}

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCLRDdBQKaCSAMB
# L3sKf52HOpySM+5K//xr4tv6xA6pJqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYEDCCGAwCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAPFmAYWfkE
# vfFZafd80MZL5MgQ+oZDik054l02VyWIAjANBgkqhkiG9w0BAQEFAASCAgBaPSoq
# Ic6EC/PwL9ARyG1iAUxe5PUo+2ebM2s2YwWd4zyOdyPHx5cdu3C9kEDWdAhgsiYA
# 0IkAK+gKocHPwzL1h+xaCzcRaRr2RRtsrS/9kX5DU9ZcXoYzpBCYQ9nUDnkFN0Qe
# vARKeaoT+XuiAalXYv/89/WtFDdE44ViBjH84qh1uh9j+RueSWJZPXuoagG1tWPT
# +9EduOCUexe2hD+rp47DX6blskSETw+hwF9zpYE9W0p7rZwEl+23BYJLMWjNpuE6
# af5dDjqcsHB2cagH80mtP+SAz4gP8nY7rmhdcRzZHT5oxU2jMUP158aWIrTWQ16P
# CWHF33+6amFNZDjH8gzJpT5IxW21Z3NPtFSEPFcv0JXDpwG4YXhR7p9ZBMzeIwUk
# xhOn5Z1MXgkc0dcejjoXbe9oaj1/0YA0mlqgncocw9PvNvJ47F+lUpwNOJPtc6yY
# w4oCYfj2Nf0GVTEPs2GfcyRK+2s/2RUxOVmxUhiM4YPS4f6Eh1l/Oe9mM4RuFOo2
# XVzWZQW0jF9sGSDEUueBPPgXL9yOHKh85NpLYkweguFvpA1+ja40JRgBpaeKNQkm
# 0wVmkuop+7KAWen7JZTYi+KCfJ7kdNSSH2iu1HkxLhqJUe2iwviMsZkTYBEkV/a0
# J0150waa3OQtNcG3QpPm0M8fhnTjUP8pr42QuqGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEINYS/ZDY9HBCOkdOKrRHMSJ2Hxz/c8XkzAHOQJpNX/utAggYPOKrykzL
# pBgPMjAyNjExMTAwODIyMjBaMAMCAQECEFPmIfDw47tFeVXGL+noUyagghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlUwggJRAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjIyMFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEINc9
# /9zSqcMf+Q/3la6ZUg/7Mk1g6T8i8LuIinP+hGe0MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiBb4GfxrBYwsNMT5LBsGqbgXy48Zrm7GIMv+hbhd7UiPwIgWW8n
# EkawAgkN5uRXzQUdPg8a15mz+b5nkpGrCn2YYu4=
# SIG # End signature block
