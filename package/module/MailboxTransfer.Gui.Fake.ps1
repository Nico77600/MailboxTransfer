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
    <# A small deterministic generator, so the screenshots are the same every time. #>
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
    <# A mailbox with its three locations: Primary, Archive, Auxiliary archive 1 (being filled by Microsoft). -Recoverable: with Recoverable Items (the box "Include deleted items"). #>
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
    <# Up to 12 synthetic messages of a folder, as rows of the list. #>
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
    $cap = @{ Events = $true; AppOnly = $true; PstExport = $true; PstCheck = $true; PstRepair = $true; PstImport = $false; EmlExport = $false; SharedMailbox = $false; FolderChoice = $false; PstCopy = $false; PstPreview = $false }
    foreach ($k in $Capability.Keys) { $cap[$k] = $Capability[$k] }
    $engine = [ordered]@{
        Name = 'Fake'
        Worker = [pscustomobject]@{ Inline = $true }
        Capability = $cap
        State = @{ Signed = $false }
    }
    $engine.Readiness = { $work = $script:Gui.S.WorkFolder; [pscustomobject]@{ Lights = @(
                @{ Name = 'PowerShell 7.5 or later'; Detail = '7.5.4'; Ok = $true },
                @{ Name = 'Sign-in component (MSAL)'; Detail = 'found in Microsoft.Graph.Authentication 2.41.1; the Windows sign-in window is available'; Ok = $true },
                @{ Name = 'Work folder'; Detail = "$work - 412 GiB free"; Ok = $true },
                @{ Name = 'Network connection'; Detail = 'login.microsoftonline.com answers'; Ok = $true },
                @{ Name = 'Files unblocked'; Detail = 'no file of this tool is marked as downloaded'; Ok = $true }
            ) | ForEach-Object { $_ } } }
    $engine.WindowSignIn = { [pscustomobject]@{ Available = $true; Reason = ''; Directory = '' } }
    $engine.ConfigPath = { 'D:\MailboxTransfer\config\MailboxTransfer.config.psd1' }
    $engine.SignInSettings = { [pscustomobject]@{ TenantId = 'contoso.onmicrosoft.com'; ClientId = '7c0d1e2f-4a5b-4c6d-8e9f-0a1b2c3d4e5f'; AppTenantId = 'contoso.onmicrosoft.com'; AppClientId = '3f9a6b21-8d44-4c07-b5e2-91a7d0c4e6f8'; CertificateThumbprint = 'A1B2C3D4E5F60718293A4B5C6D7E8F9012345678' } }
    $engine.SignIn = { param($Mailbox, $UseCode, $CodeFile, $OnDone, $OnError, $Silent, $TenantId, $ClientId, $App, $ParentWindow)
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
    $engine.SaveSelections = { param($Groups, $Path, $OnDone, $OnError) & $OnDone $Path }
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

