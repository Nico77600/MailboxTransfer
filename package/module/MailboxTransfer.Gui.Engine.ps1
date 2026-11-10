#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - the real engine adapter (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    The same members as the fake engine (MailboxTransfer.Gui.Fake.ps1), over the module. In the console's own process, on one worker runspace (the
    sign-in session lives there): sign in, locations, browse (folders, messages, preview), selection files, the look into a PST or a store.
    Long work (export, PST write, import, check, repair) runs in a child pwsh: Gui.Jobs.ps1. Every call answers through callbacks that run on the
    UI thread. A feature the engine does not have (PST import, EML export and import, shared mailboxes) is reported by Capability and shown
    "coming in this version" by the pages; with the merged engine every one of them is there.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

function ConvertTo-MbtGuiFolderNode {
    <# BrowseFolder objects of the engine to nodes of the tree. #>
    param([object[]]$Folders)
    $glyphs = @{ inbox = 'mail'; sentitems = 'send'; drafts = 'pencil'; calendar = 'cal'; contacts = 'people'; deleteditems = 'trash'; archive = 'archive' }
    foreach ($f in $Folders) {
        $n = [MailboxTransferGui.FolderNode]::new([string]$f.Id, [string]$f.DisplayName, [long]$f.TotalItemCount, [long]$f.SizeBytes)
        $wk = [string]$f.WellKnownName
        $gname = if ($wk -and $glyphs.ContainsKey($wk.ToLowerInvariant())) { $glyphs[$wk.ToLowerInvariant()] } else { 'folder' }
        $n.Glyph = Get-MbtGlyph $gname
        $n.Tag = $f
        if ($f.MoveText) { $n.Sub = [string]$f.MoveText }
        if ($f.State -eq 'ContentNotAvailable') { $n.Dim = $true; $n.Sub = 'Not ready: Microsoft is still moving this folder' }
        if ($f.ChildFolderCount -gt 0) { $n.MarkNotRead() }
        $n
    }
}

function ConvertTo-MbtGuiMessageRow {
    param($Item, [bool]$HeadersOnly)
    $r = [MailboxTransferGui.MessageRow]::new()
    $r.Id = [string]$Item.Id; $r.Tag = $Item
    $r.SizeBytes = [long]$Item.Size; $r.SizeText = Format-MbtGuiBytes $Item.Size
    $d = if ($Item.Received) { $Item.Received } elseif ($Item.Created) { $Item.Created } else { $null }
    $r.DateText = if ($d) { ([datetimeoffset]$d).UtcDateTime.ToString('yyyy-MM-dd HH:mm') } else { '' }
    if ($HeadersOnly -or -not $Item.HasDetails) {
        $r.From = '(sender not available)'; $r.Subject = '(subject not available before export)'; $r.SubjectNa = $true
    } else {
        $r.From = if ($Item.FromName) { [string]$Item.FromName } else { [string]$Item.FromAddress }
        $r.Address = [string]$Item.FromAddress
        $r.Subject = if ($Item.Subject) { [string]$Item.Subject } else { '(no subject)' }
        $r.HasAttachment = [bool]$Item.HasAttachments
    }
    return $r
}

function New-MbtGuiRealEngine {
    <# The adapter over the module. #>
    param([Parameter(Mandatory)][string]$PackageRoot, [Parameter(Mandatory)][string]$EntryScript, [string]$WorkFolder)
    $modulePath = Join-Path $PackageRoot 'MailboxTransfer.psd1'
    Import-Module -Name $modulePath -Force -ErrorAction Stop
    $capObj = Get-MbtCapability -EntryScript $EntryScript
    $cap = @{}
    foreach ($p in $capObj.PSObject.Properties) { if ($p.Value -is [bool]) { $cap[$p.Name] = $p.Value } }
    $worker = New-MbtGuiWorker -ModulePath $modulePath
    $e = [ordered]@{ Name = 'Real'; Worker = $worker; Capability = $cap; PackageRoot = $PackageRoot; EntryScript = $EntryScript; State = @{} }

    $e.Readiness = {
        $lights = [Collections.Generic.List[object]]::new()
        $ps = $PSVersionTable.PSVersion
        $lights.Add(@{ Name = 'PowerShell 7.5 or later'; Detail = $ps.ToString(); Ok = ($ps -ge [version]'7.5') })
        $msal = $null; try { $msal = Find-MbtMsal } catch { $msal = $null }
        $lights.Add(@{ Name = 'Sign-in component (MSAL)'; Detail = $(if ($msal) { 'found' } else { 'not found: see the guide, "Start here"' }); Ok = [bool]$msal })
        $work = $script:Gui.S.WorkFolder
        $freeText = 'unknown'; $ok = $true
        try { $root = [IO.Path]::GetPathRoot($work); $d = [IO.DriveInfo]::new($root); $freeText = "$(Format-MbtGuiBytes $d.AvailableFreeSpace) free"; $ok = $d.AvailableFreeSpace -gt 5GB } catch { $ok = $true }
        $lights.Add(@{ Name = 'Work folder'; Detail = "$work - $freeText"; Ok = $ok })
        $net = [Net.NetworkInformation.NetworkInterface]::GetIsNetworkAvailable()
        $lights.Add(@{ Name = 'Network connection'; Detail = $(if ($net) { 'a connection is present' } else { 'no connection found' }); Ok = $net })
        $blocked = $false; try { $blocked = [bool](Get-Item -LiteralPath $PackageRoot -Stream Zone.Identifier -ErrorAction SilentlyContinue) -or [bool](Get-Item -LiteralPath $EntryScript -Stream Zone.Identifier -ErrorAction SilentlyContinue) } catch { }
        if ($blocked) { $lights.Add(@{ Name = 'Files unblocked'; Detail = 'Windows blocked a file of this tool. Unblock the folder (see the guide, "Start here").'; Ok = $false }) }
        [pscustomobject]@{ Lights = $lights.ToArray() }
    }.GetNewClosure()

    $e.Jobs = {
        param($Work)
        $out = foreach ($j in @(Get-MbtJob -WorkFolder $Work -Limit 8)) {
            $state = switch ($j.EffectiveState) { 'complete' { 'ok' } 'error' { 'warn' } 'needs-decision' { 'warn' } default { 'run' } }
            $counts = $j.Counts
            $sub = switch ($j.EffectiveState) {
                'complete' { 'finished ' + $j.Updated.ToLocalTime().ToString('yyyy-MM-dd') }
                'needs-decision' { 'Almost done: an item needs your decision.' }
                'error' { 'Stopped by an error. Nothing is lost.' }
                default { "stopped $($j.Updated.ToLocalTime().ToString('yyyy-MM-dd HH:mm'))" + $(if ($counts -and $counts.PSObject.Properties['done']) { " - $(Format-MbtGuiCount $counts.done) items done" } else { '' }) }
            }
            if ($j.Alive) { $sub = 'running elsewhere' }
            [pscustomobject]@{ Task = $j.Task; Title = $(if ($j.Title) { $j.Title } else { $j.JobId }); Sub = $sub; State = $state; Button = $(if ($j.EffectiveState -eq 'complete' -or $j.Alive) { 'Open' } else { 'Continue' }); Folder = $j.Folder }
        }
        @($out)
    }.GetNewClosure()

    $e.FreeSpace = { param($Path) try { $r = [IO.Path]::GetPathRoot([IO.Path]::GetFullPath($Path)); return ([IO.DriveInfo]::new($r)).AvailableFreeSpace } catch { return $null } }

    # -Silent: only the cached token (no browser, no device code): fails at once when there is none (the live check of tools\Test-ConsoleLive.ps1)
    # the tenant and the application of the configuration (empty: the Sign-in card asks for them)
    $e.SignInSettings = { $c = Import-MbtConfiguration; [pscustomobject]@{ TenantId = [string]$c.Lab.TenantId; ClientId = [string]$c.Lab.ClientId } }

    # -TenantId / -ClientId: typed in the Sign-in card when the configuration has none
    # -App: @{ Tenant; ClientId; Thumbprint } signs in as the application with its certificate (no window, no user: $Mailbox is then the mailbox it works on)
    $e.SignIn = { param($Mailbox, $UseCode, $CodeFile, $OnDone, $OnError, $Silent, $TenantId, $ClientId, $App)
        Start-MbtGuiWork -Worker $worker -Arguments @($Mailbox, [bool]$UseCode, $CodeFile, [bool]$Silent, [string]$TenantId, [string]$ClientId, $App) -OnDone $OnDone -OnError $OnError -Script {
            param($upn, $useCode, $codeFile, $silent, $tenantId, $clientId, $app)
            $cfg = Import-MbtConfiguration
            if ($app) {
                $session = Connect-MbtGraph -TenantId $app.Tenant -ClientId $app.ClientId -UserPrincipalName $upn -CertificateThumbprint $app.Thumbprint -Configuration $cfg
                $t = $session.Token
                $roles = @($t.Roles)
                # no role at all: the permissions are granted only through Exchange RBAC for Applications, which the token does not show (nothing is reported missing)
                $viaRbac = -not $roles.Count
                $missing = if ($viaRbac) { @() } else { @(Get-MbtMissingAppPermission -Token $t -Mode Import) + @(if (-not ('Mail.Read' -in $roles -or 'Mail.ReadWrite' -in $roles)) { 'Mail.Read' }) }
                return [pscustomobject]@{ Account = $upn; TenantId = $t.TenantId; Scopes = @($t.Scopes); Missing = @($missing | Select-Object -Unique); CanReadMail = ($viaRbac -or 'Mail.Read' -in $roles -or 'Mail.ReadWrite' -in $roles); Access = $(if ($viaRbac) { 'Application (certificate, permissions from Exchange RBAC)' } else { 'Application (certificate, no user)' }) }
            }
            if (-not $tenantId) { $tenantId = [string]$cfg.Lab.TenantId }
            if (-not $clientId) { $clientId = [string]$cfg.Lab.ClientId }
            Assert-MbtSignInSettings -TenantId $tenantId -ClientId $clientId
            $session = Connect-MbtGraph -TenantId $tenantId -ClientId $clientId -UserPrincipalName $upn -DeviceCode:$useCode -DeviceCodeFile $codeFile -NoPrompt:$silent -Configuration $cfg
            $t = $session.Token
            $scopes = @($t.Scopes)
            $canMail = ('Mail.Read' -in $scopes) -or ('Mail.ReadWrite' -in $scopes)
            $missing = @()
            if (-not ('MailboxFolder.Read' -in $scopes -or 'MailboxFolder.ReadWrite' -in $scopes)) { $missing += 'MailboxFolder.Read' }
            if (-not ('MailboxItem.Read' -in $scopes -or 'MailboxItem.ReadWrite' -in $scopes)) { $missing += 'MailboxItem.Read' }
            if (-not ('MailboxItem.ImportExport' -in $scopes)) { $missing += 'MailboxItem.ImportExport (needed to import)' }
            [pscustomobject]@{ Account = $(if ($t.Account) { $t.Account } else { $upn }); TenantId = $t.TenantId; Scopes = $scopes; Missing = $missing; CanReadMail = $canMail; Access = 'Delegated (you, with your rights)' }
        }
    }.GetNewClosure()

    $e.SignOut = { Start-MbtGuiWork -Worker $worker -Script { Disconnect-MbtGraph } }.GetNewClosure()

    # -Shared: the address of a shared mailbox or of another user's mailbox (resolved to its physical id, one location: the primary mailbox of it);
    # -AuxIds: auxiliary archives already known (ids or guids), as the -AuxArchiveMailboxId of the entry script
    $e.Locations = { param($Mailbox, $LocationsFile, $OnDone, $OnError, $Shared, $AuxIds)
        Start-MbtGuiWork -Worker $worker -Arguments @($Mailbox, $LocationsFile, $Shared, $AuxIds) -OnError $OnError -Script {
            param($upn, $file, $shared, $auxIds)
            if ($shared) {
                $id = Resolve-MbtMailboxAddress -Address $shared -TenantId (Get-MbtSession).TenantId
                return @(New-MbtBrowseLocation -Kind Primary -MailboxId $id -Source "address $shared")
            }
            $p = @{ UserPrincipalName = $upn }
            if ($file) { $p.LocationsFile = $file }
            if ($auxIds) { $p.AuxArchiveMailboxId = @($auxIds) }
            # the ids of the configuration (as the entry script uses them) when it is the user of the configuration
            $configured = Get-MbtConfiguredMailboxId -Configuration (Import-MbtConfiguration) -UserPrincipalName $upn
            if ($configured.Primary) { $p.PrimaryMailboxId = $configured.Primary }
            if ($configured.Archive) { $p.ArchiveMailboxId = $configured.Archive }
            $found = @(Get-MbtBrowseLocation @p)
            # Graph gives no mailbox id without User.Read.All: the primary mailbox is then found through the import-export API (its folder listing names the mailbox)
            if (-not @($found | Where-Object Kind -EQ 'Primary')) {
                try { $found = @(New-MbtBrowseLocation -Kind Primary -MailboxId (Resolve-MbtMailboxAddress -Address $upn -TenantId (Get-MbtSession).TenantId) -Source 'address') + $found } catch { Write-Verbose "primary by address: $_" }
            }
            $found
        } -OnDone {
            param($locs)
            $aux = 0
            $out = foreach ($l in @($locs)) {
                $label = switch ($l.Kind) { 'Primary' { $(if ($Shared) { "Shared mailbox $Shared" } else { 'Primary mailbox' }) } 'Archive' { 'Archive' } default { $aux++; "Auxiliary archive $aux" } }
                [pscustomobject]@{ Kind = [string]$l.Kind; Label = $label; Hint = $(if ($l.Kind -eq 'AuxArchive') { 'Counts appear as you open folders. Size not given by Microsoft.' } else { 'Counts appear as you open folders.' }); Root = [string]$l.MailboxId; MailboxId = [string]$l.MailboxId; Loc = $l; Filling = ($l.Kind -eq 'AuxArchive') }
            }
            & $OnDone @($out)
        }.GetNewClosure()
    }.GetNewClosure()

    # -IncludeRecoverable: the box "Include deleted items (Recoverable Items)" of the From page; the folders are read as the user opens them, so only Children needs it
    $e.Roots = { param($Locations, $OnDone, $OnError, $IncludeRecoverable)
        $roots = foreach ($l in @($Locations)) {
            $kind = switch ($l.Kind) { 'Primary' { 'primary' } 'Archive' { 'archive' } default { 'aux' } }
            $n = [MailboxTransferGui.FolderNode]::new([string]$l.MailboxId, $l.Label, 0, -1)
            $n.IsRoot = $true; $n.Kind = $kind; $n.Glyph = Get-MbtGlyph $(if ($kind -eq 'primary') { 'person' } else { 'archive' })
            $n.Sub = switch ($kind) { 'primary' { $(if ($script:Gui.S.Shared) { $script:Gui.S.Shared } else { $script:Gui.S.Mailbox }) } 'archive' { 'In-Place Archive' } default { 'Being filled by Microsoft - headers only until exported' } }
            $n.HeadersOnly = ($kind -eq 'aux')
            $n.Tag = @{ Location = $l.Loc }
            $n.MarkNotRead()
            $n
        }
        & $OnDone @($roots)
    }

    $e.Children = { param($Node, $OnDone, $OnError)
        $root = Get-MbtGuiRootNode $Node
        $loc = $root.Tag.Location
        $parent = if ($Node.IsRoot) { $null } else { $Node.Tag }
        if ($root.Kind -in 'pst', 'store', 'eml') { & $script:Gui.Engine.SourceChildren $root $Node $OnDone $OnError; return }
        # "Include deleted items": Recoverable Items appear at the first level, with the folders of the default set (Deletions, Purges, Versions, DiscoveryHolds) below them
        $deleted = [bool]$script:Gui.S.RecoverableItems
        Start-MbtGuiWork -Worker $worker -Arguments @($loc, $parent, $deleted) -OnError $OnError -Script {
            param($loc, $parent, $deleted)
            $names = [MailboxTransfer.RecoverableItems]::DefaultFolders
            if ($parent) { @(Get-MbtBrowseFolder -Location $loc -Lazy -Parent $parent -RecoverableItems:$deleted -RecoverableItemsFolders $names) } else { @(Get-MbtBrowseFolder -Location $loc -Lazy -RecoverableItems:$deleted -RecoverableItemsFolders $names) }
        } -OnDone { param($folders) & $OnDone @(ConvertTo-MbtGuiFolderNode $folders) }.GetNewClosure()
    }.GetNewClosure()

    $e.Items = { param($Node, $Token, $OnDone, $OnError)
        $root = Get-MbtGuiRootNode $Node
        if ($root.Kind -in 'pst', 'store', 'eml') { & $script:Gui.Engine.SourceItems $root $Node $Token $OnDone $OnError; return }
        $headersOnly = [bool]$root.HeadersOnly
        Start-MbtGuiWork -Worker $worker -Arguments @($root.Tag.Location, $Node.Tag, $Token) -OnError $OnError -Script {
            param($loc, $folder, $token)
            Get-MbtBrowseItem -Location $loc -Folder $folder -Page -PageSize 50 -Token $token
        } -OnDone {
            param($pages)
            $pg = @($pages)[0]
            $rows = @($pg.Items | ForEach-Object { ConvertTo-MbtGuiMessageRow $_ $headersOnly })
            $note = $null
            if ($pg.State -eq 'ContentNotAvailable') { $note = "Not ready. $($pg.StateMessage) The folder is not empty or lost; ask again later." }
            & $OnDone ([pscustomobject]@{ Rows = $rows; Next = $pg.NextToken; Total = $Node.Count; Note = $note })
        }.GetNewClosure()
    }.GetNewClosure()

    $e.Preview = { param($Row, $Node, $OnDone, $OnError)
        $root = Get-MbtGuiRootNode $Node
        if ($root.Kind -in 'pst', 'store', 'eml') { & $script:Gui.Engine.SourcePreview $root $Node $Row $OnDone $OnError; return }
        $label = "$($root.Name) / $($Node.Name)"
        Start-MbtGuiWork -Worker $worker -Arguments @($root.Tag.Location, $Row.Tag) -OnError $OnError -Script {
            param($loc, $item)
            Get-MbtBrowsePreview -Location $loc -Item $item -BodyChars 2000
        } -OnDone {
            param($res)
            $p = @($res)[0]; $h = $p.Header
            $att = @($p.Attachments | ForEach-Object { "$($_.Name)   $(Format-MbtGuiBytes $_.Size)" }) -join "`n"
            & $OnDone ([pscustomobject]@{
                    Subject = [string]$h.Subject; From = [string]$h.FromName; Address = [string]$h.FromAddress; To = (@($p.To | ForEach-Object { $_.ToString() }) -join '; ')
                    Received = $(if ($h.Received) { ([datetimeoffset]$h.Received).UtcDateTime.ToString('yyyy-MM-dd HH:mm') + ' UTC' } else { '' }); Location = $label
                    Details = ('{0} - {1}' -f $h.ItemClass, (Format-MbtGuiBytes $h.Size)); Attachment = $att; Body = $(if ($p.BodyText) { [string]$p.BodyText } else { [string]$p.BodyPreview }); Available = $true })
        }.GetNewClosure()
    }.GetNewClosure()

    $e.SaveSelection = { param($Root, $Entries, $Path, $OnDone, $OnError)
        $loc = $Root.Tag.Location
        $plain = @($Entries | ForEach-Object { @{ Folder = $_.Node.Tag; Kind = $_.Kind; Items = @($_.Node.PickRows.Values) } })
        Start-MbtGuiWork -Worker $worker -Arguments @($loc, $plain, $Path) -OnDone $OnDone -OnError $OnError -Script {
            param($loc, $entries, $path)
            $sel = New-MbtBrowseSelection -Location $loc
            foreach ($en in $entries) {
                switch ($en.Kind) {
                    'tree' { $null = Add-MbtBrowseSelection -Selection $sel -Folder $en.Folder -IncludeSubfolders }
                    'own' { $null = Add-MbtBrowseSelection -Selection $sel -Folder $en.Folder }
                    default { $null = Add-MbtBrowseSelection -Selection $sel -Folder $en.Folder -Item @($en.Items | ForEach-Object { $_ }) }
                }
            }
            Save-MbtBrowseSelection -Selection $sel -Path $path
        }
    }.GetNewClosure()

    $e.Plan = { param($Request, $OnDone, $OnError) & $OnDone ([pscustomobject]@{ Ok = $true; Message = '' }) }

    # ---- PST and store, read only -----------------------------------------------------------------------------------------------------------
    $e.OpenSource = { param($Kind, $Path, $OnDone, $OnError)
        switch ($Kind) { 'pst' { Open-MbtGuiPstSource $Path $OnDone $OnError } 'eml' { Open-MbtGuiEmlSource $Path $OnDone $OnError } default { Open-MbtGuiStoreSource $Path $OnDone $OnError } }
    }
    $e.SourceChildren = { param($Root, $Node, $OnDone, $OnError) & $OnDone @() }
    $e.SourceItems = { param($Root, $Node, $Token, $OnDone, $OnError) Read-MbtGuiSourceItems $Root $Node $Token $OnDone $OnError }
    $e.SourcePreview = { param($Root, $Node, $Row, $OnDone, $OnError) Read-MbtGuiSourcePreview $Root $Node $Row $OnDone $OnError }

    $e.CheckPst = { param($Path, $OnDone, $OnError) Start-MbtGuiPstCheckJob $Path $OnDone $OnError }
    return [pscustomobject]$e
}

#region Looking into a PST or a store ------------------------------------------------------------------------------------------------------------

function Find-MbtGuiPstParts {
    <# archive.pst, archive-002.pst, ...: the parts that are there. #>
    param([string]$Path)
    $dir = Split-Path $Path -Parent; $stem = [IO.Path]::GetFileNameWithoutExtension($Path) -replace '-\d{3}$', ''
    $parts = @(Get-ChildItem -LiteralPath $dir -Filter "$stem*.pst" -File -ErrorAction SilentlyContinue | Where-Object { $_.BaseName -eq $stem -or $_.BaseName -match "^$([regex]::Escape($stem))-\d{3}$" } | Sort-Object Name)
    return $parts
}

function Open-MbtGuiPstSource {
    param([string]$Path, [scriptblock]$OnDone, [scriptblock]$OnError)
    $g = $script:Gui
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { & $OnError 'The file was not found.'; return }
    $parts = Find-MbtGuiPstParts $Path
    $bytes = ($parts | Measure-Object -Property Length -Sum).Sum
    $worker = $g.Engine.Worker
    Start-MbtGuiWork -Worker $worker -Arguments @($Path) -OnError $OnError -Script {
        param($path)
        $reader = [MailboxTransfer.PstReader]::Open($path)
        try {
            # the tree below "Top of Personal Folders" (what the engine imports and copies); the path of a folder is the one the engine takes in -IncludeFolder
            $topFolder = [MailboxTransfer.PstImportPlanner]::TopFolder($reader)
            if (-not $topFolder) { throw 'This PST file has no folder tree (no "Top of Personal Folders").' }
            $build = $null
            $build = { param($f, $p) @{ Id = [string]$f.Nid; Name = [string]$f.Name; Count = [long]$f.ContentCount; Path = $p; Kids = @($f.Children | ForEach-Object { & $build $_ $(if ($p) { $p + '/' + [MailboxTransfer.PstImportPlanner]::Clean($_.Name) } else { [MailboxTransfer.PstImportPlanner]::Clean($_.Name) }) }) } }
            $top = & $build $topFolder ''
            $total = 0L; $folders = 0
            $stack = [Collections.Generic.Stack[object]]::new(); $stack.Push($topFolder)
            while ($stack.Count) { $f = $stack.Pop(); $total += $f.ContentCount; $folders++; foreach ($c in $f.Children) { $stack.Push($c) } }
            [pscustomobject]@{ Tree = $top; Items = $total; Folders = $folders; Name = $reader.DisplayName }
        } finally { $reader.Dispose() }
    } -OnDone {
        param($res)
        $r = @($res)[0]
        $root = [MailboxTransferGui.FolderNode]::new('pst-root', [IO.Path]::GetFileName($Path), 0, -1)
        $root.IsRoot = $true; $root.Kind = 'pst'; $root.Glyph = Get-MbtGlyph pst
        $root.Sub = "$($parts.Count) part$(if ($parts.Count -ne 1) { 's' }) - $(Format-MbtGuiBytes $bytes) - read only"
        $root.Tag = @{ Path = $Path; Kind = 'pst' }
        $mk = $null
        $mk = { param($d) $n = [MailboxTransferGui.FolderNode]::new([string]$d.Id, [string]$d.Name, [long]$d.Count, -1); $n.Tag = @{ Nid = [uint32]$d.Id; PstPath = $Path; FolderPath = [string]$d.Path }; $kids = @($d.Kids | ForEach-Object { & $mk $_ }); if ($kids.Count) { $n.SetChildren([MailboxTransferGui.FolderNode[]]$kids) }; $n }
        $kids = @($r.Tree.Kids | ForEach-Object { & $mk $_ })
        if ($kids.Count) { $root.SetChildren([MailboxTransferGui.FolderNode[]]$kids) }
        $verdict = '**{0} part{1} found** - {2} - the check of the file runs on the next page when you repair it. Reading it here only reads.' -f $parts.Count, $(if ($parts.Count -ne 1) { 's' } else { '' }), (Format-MbtGuiBytes $bytes)
        & $OnDone ([pscustomobject]@{ Ok = $true; Level = 'ok'; Message = $verdict; Items = $r.Items; Bytes = $bytes; Folders = $r.Folders; Parts = $parts.Count; Roots = @($root) })
    }.GetNewClosure()
}

function Open-MbtGuiStoreSource {
    param([string]$Path, [scriptblock]$OnDone, [scriptblock]$OnError)
    $g = $script:Gui
    if (-not (Test-Path -LiteralPath (Join-Path $Path 'catalog.db'))) { & $OnError 'There is no export store in this folder (catalog.db is missing).'; return }
    Start-MbtGuiWork -Worker $g.Engine.Worker -Arguments @($Path) -OnError $OnError -Script {
        param($path)
        $store = Open-MbtExportStore -Path $path -ReadOnly
        try {
            $mailboxes = foreach ($m in $store.GetMailboxes()) {
                $folders = @($store.GetFolders($m.Id))
                @{ Id = $m.Id; Kind = [string]$m.Kind; Folders = @($folders | ForEach-Object { @{ Key = $_.Key; Id = $_.FolderId; Parent = $_.ParentFolderId; Name = $_.DisplayName; Count = [long]$(if ($null -ne $_.SourceItemCount) { $_.SourceItemCount } else { 0 }); Size = [long]$(if ($null -ne $_.SourceSize) { $_.SourceSize } else { -1 }) } }) }
            }
            $status = Get-MbtStoreStatus -Store $store
            [pscustomobject]@{ Mailboxes = @($mailboxes); Status = $status }
        } finally { $store.Dispose() }
    } -OnDone {
        param($res)
        $r = @($res)[0]
        $roots = foreach ($m in $r.Mailboxes) {
            $root = [MailboxTransferGui.FolderNode]::new("store-$($m.Id)", $(if ($m.Kind -eq 'Archive') { 'Archive' } elseif ($m.Kind -eq 'Primary') { 'Primary mailbox' } else { 'Auxiliary archive' }), 0, -1)
            $root.IsRoot = $true; $root.Kind = 'store'; $root.Glyph = Get-MbtGlyph archive; $root.Sub = 'export store - read only'
            $root.Tag = @{ Path = $Path; Kind = 'store' }
            $byParent = @{}
            foreach ($f in $m.Folders) { $k = [string]$f.Parent; if (-not $byParent.ContainsKey($k)) { $byParent[$k] = [Collections.Generic.List[object]]::new() }; $byParent[$k].Add($f) }
            $mk = $null
            $mk = { param($f) $n = [MailboxTransferGui.FolderNode]::new([string]$f.Id, [string]$f.Name, [long]$f.Count, [long]$f.Size); $n.Tag = @{ Key = [long]$f.Key; StorePath = $Path }; $kids = @(); if ($byParent.ContainsKey([string]$f.Id)) { $kids = @($byParent[[string]$f.Id] | ForEach-Object { & $mk $_ }) }; if ($kids.Count) { $n.SetChildren([MailboxTransferGui.FolderNode[]]$kids) }; $n }
            $tops = @($m.Folders | Where-Object { -not $_.Parent -or -not ($m.Folders | Where-Object { $_.Id -eq $f.Parent }) } | ForEach-Object { & $mk $_ })
            $tops = @($m.Folders | Where-Object { $p = $_.Parent; -not $p -or -not (@($m.Folders | Where-Object { $_.Id -eq $p }).Count) } | ForEach-Object { & $mk $_ })
            if ($tops.Count) { $root.SetChildren([MailboxTransferGui.FolderNode[]]$tops) }
            $root
        }
        $items = 0L; $bytes = 0L; $folders = 0
        foreach ($m in $r.Mailboxes) { foreach ($f in $m.Folders) { $items += $f.Count; if ($f.Size -gt 0) { $bytes += $f.Size }; $folders++ } }
        & $OnDone ([pscustomobject]@{ Ok = $true; Level = 'ok'; Message = "**Store opened:** $(Format-MbtGuiCount $items) items in $folders folders. The items are checked against their fingerprints while they are imported."; Items = $items; Bytes = $bytes; Folders = $folders; Parts = 0; Roots = @($roots) })
    }.GetNewClosure()
}

function Open-MbtGuiEmlSource {
    <# A folder of .eml files (or one file), read only: the folders below it with the number of messages. The messages are read when they are imported. #>
    param([string]$Path, [scriptblock]$OnDone, [scriptblock]$OnError)
    $g = $script:Gui
    if (-not (Test-Path -LiteralPath $Path)) { & $OnError 'The folder was not found.'; return }
    Start-MbtGuiWork -Worker $g.Engine.Worker -Arguments @($Path) -OnError $OnError -Script {
        param($path)
        $root = (Resolve-Path -LiteralPath $path).Path
        $isDir = (Get-Item -LiteralPath $root).PSIsContainer
        $files = @(if ($isDir) { Get-ChildItem -LiteralPath $root -Filter '*.eml' -File -Recurse } else { Get-Item -LiteralPath $root })
        $base = if ($isDir) { $root } else { Split-Path $root -Parent }
        $dirs = @{}
        foreach ($f in $files) {
            $rel = [IO.Path]::GetRelativePath($base, $f.DirectoryName).Replace([string][char]92, '/'); if ($rel -eq '.') { $rel = '' }
            if (-not $dirs.ContainsKey($rel)) { $dirs[$rel] = @{ Count = 0L; Bytes = 0L } }
            $dirs[$rel].Count++; $dirs[$rel].Bytes += $f.Length
        }
        [pscustomobject]@{ Files = $files.Count; Bytes = [long](($files | Measure-Object -Property Length -Sum).Sum); Dirs = @($dirs.GetEnumerator() | ForEach-Object { @{ Path = [string]$_.Key; Count = [long]$_.Value.Count; Bytes = [long]$_.Value.Bytes } }); Name = (Split-Path $root -Leaf) }
    } -OnDone {
        param($res)
        $r = @($res)[0]
        if (-not $r.Files) { & $OnError 'There is no .eml file in this folder.'; return }
        $root = [MailboxTransferGui.FolderNode]::new('eml-root', [string]$r.Name, 0, -1)
        $root.IsRoot = $true; $root.Kind = 'eml'; $root.Glyph = Get-MbtGlyph folder
        $root.Sub = "$(Format-MbtGuiCount $r.Files) messages - $(Format-MbtGuiBytes $r.Bytes) - read only"
        $root.Tag = @{ Path = $Path; Kind = 'eml' }
        # one node per folder that holds .eml files, below its parent folder
        $nodes = @{ '' = $root }
        $get = $null
        $get = { param($rel)
            if ($nodes.ContainsKey($rel)) { return $nodes[$rel] }
            $parentRel = if ($rel.Contains('/')) { $rel.Substring(0, $rel.LastIndexOf('/')) } else { '' }
            $parent = & $get $parentRel
            $leaf = $rel.Substring($rel.LastIndexOf('/') + 1)
            $d = @($r.Dirs | Where-Object { $_.Path -eq $rel })
            $n = [MailboxTransferGui.FolderNode]::new("eml:$rel", $leaf, $(if ($d.Count) { [long]$d[0].Count } else { 0 }), $(if ($d.Count) { [long]$d[0].Bytes } else { -1 }))
            $n.Tag = @{ Rel = $rel; EmlPath = $Path }
            $nodes[$rel] = $n
            $kids = [Collections.Generic.List[MailboxTransferGui.FolderNode]]::new()
            foreach ($c in @($parent.Children)) { if ($c) { $kids.Add($c) } }
            $kids.Add($n)
            $parent.SetChildren($kids.ToArray())
            return $n
        }
        foreach ($d in ($r.Dirs | Sort-Object { $_.Path })) { if ($d.Path) { $null = & $get $d.Path } }
        $verdict = '**{0} message{1} found** - {2} - read only. The messages are read when they are imported; each folder below becomes a folder of the mailbox.' -f (Format-MbtGuiCount $r.Files), $(if ($r.Files -ne 1) { 's' } else { '' }), (Format-MbtGuiBytes $r.Bytes)
        & $OnDone ([pscustomobject]@{ Ok = $true; Level = 'ok'; Message = $verdict; Items = [long]$r.Files; Bytes = [long]$r.Bytes; Folders = $nodes.Count; Parts = 0; Roots = @($root) })
    }.GetNewClosure()
}

function Read-MbtGuiSourceItems {
    param($Root, $Node, $Token, [scriptblock]$OnDone, [scriptblock]$OnError)
    $g = $script:Gui
    if ($Root.Kind -eq 'eml') { & $OnDone ([pscustomobject]@{ Rows = @(); Next = $null; Total = $Node.Count; Note = 'The messages of a folder of .eml files are read when they are imported.' }); return }
    $page = if ($Token) { [int]$Token } else { 0 }
    if ($Root.Kind -eq 'store') {
        Start-MbtGuiWork -Worker $g.Engine.Worker -Arguments @($Node.Tag.StorePath, [long]$Node.Tag.Key, $page) -OnError $OnError -Script {
            param($path, $key, $page)
            $store = Open-MbtExportStore -Path $path -ReadOnly
            try { $p = $store.GetItems($key, $page, 50, [MailboxTransfer.StoreItemSort]::Received, $true); @{ Items = @($p.Items | ForEach-Object { @{ Key = $_.Key; Subject = $_.Subject; Size = $_.Size; Received = $_.ReceivedUtc } }); Total = $p.Total; PageCount = $p.PageCount } } finally { $store.Dispose() }
        } -OnDone {
            param($res)
            $r = @($res)[0]
            $rows = foreach ($i in $r.Items) {
                $row = [MailboxTransferGui.MessageRow]::new(); $row.Id = [string]$i.Key; $row.Subject = if ($i.Subject) { [string]$i.Subject } else { '(no subject)' }; $row.From = ''
                $row.SizeBytes = [long]$i.Size; $row.SizeText = Format-MbtGuiBytes $i.Size; $row.DateText = if ($i.Received) { ([datetime]$i.Received).ToString('yyyy-MM-dd HH:mm') } else { '' }
                $row.Tag = @{ Key = [long]$i.Key; StorePath = $Node.Tag.StorePath }
                $row
            }
            $next = if (($page + 1) -lt $r.PageCount) { [string]($page + 1) } else { $null }
            & $OnDone ([pscustomobject]@{ Rows = @($rows); Next = $next; Total = $r.Total; Note = $null })
        }.GetNewClosure()
        return
    }
    Start-MbtGuiWork -Worker $g.Engine.Worker -Arguments @($Node.Tag.PstPath, [uint32]$Node.Tag.Nid, $page) -OnError $OnError -Script {
        param($path, $nid, $page)
        $reader = [MailboxTransfer.PstReader]::Open($path)
        try {
            $folder = $reader.AllFolders() | Where-Object { $_.Nid -eq $nid } | Select-Object -First 1
            $nids = @($reader.MessageNids($folder, $false))
            $slice = @($nids | Select-Object -Skip ($page * 50) -First 50)
            $rows = foreach ($n in $slice) {
                try { $m = $reader.ReadMessage($n); @{ Nid = $n; Subject = $m.GetString(0x0037); From = $m.GetString(0x0C1A) } } catch { @{ Nid = $n; Subject = '(this message cannot be read)'; From = '' } }
            }
            @{ Rows = @($rows); Total = $nids.Count }
        } finally { $reader.Dispose() }
    } -OnDone {
        param($res)
        $r = @($res)[0]
        $rows = foreach ($i in $r.Rows) {
            $row = [MailboxTransferGui.MessageRow]::new(); $row.Id = [string]$i.Nid; $row.Subject = if ($i.Subject) { [string]$i.Subject } else { '(no subject)' }; $row.From = [string]$i.From
            $row.SizeText = ''; $row.DateText = ''; $row.Tag = @{ Nid = [uint32]$i.Nid; PstPath = $Node.Tag.PstPath }
            $row
        }
        $next = if ((($page + 1) * 50) -lt $r.Total) { [string]($page + 1) } else { $null }
        & $OnDone ([pscustomobject]@{ Rows = @($rows); Next = $next; Total = $r.Total; Note = $null })
    }.GetNewClosure()
}

function Read-MbtGuiSourcePreview {
    param($Root, $Node, $Row, [scriptblock]$OnDone, [scriptblock]$OnError)
    $g = $script:Gui
    $label = "$($Root.Name) / $($Node.Name)"
    if ($Root.Kind -eq 'store') {
        Start-MbtGuiWork -Worker $g.Engine.Worker -Arguments @($Row.Tag.StorePath, [long]$Row.Tag.Key) -OnError $OnError -Script {
            param($path, $key)
            $store = Open-MbtExportStore -Path $path -ReadOnly
            try { $p = $store.GetPreview($key); @{ Subject = $p.Subject; From = $p.From; To = $p.To; Body = $p.BodyPreview; Names = $p.AttachmentNames; Received = $p.ReceivedUtc; Class = $p.MessageClass } } finally { $store.Dispose() }
        } -OnDone {
            param($res)
            $p = @($res)[0]
            & $OnDone ([pscustomobject]@{ Subject = [string]$p.Subject; From = [string]$p.From; Address = ''; To = [string]$p.To; Received = $(if ($p.Received) { ([datetime]$p.Received).ToString('yyyy-MM-dd HH:mm') + ' UTC' } else { '' }); Location = $label; Details = [string]$p.Class; Attachment = [string]$p.Names; Body = [string]$p.Body; Available = $true })
        }.GetNewClosure()
        return
    }
    # a message of a PST: header, text and attachments read by the engine (read only; the message is imported with everything it holds, this is only what the pane shows)
    Start-MbtGuiWork -Worker $g.Engine.Worker -Arguments @($Row.Tag.PstPath, [uint32]$Row.Tag.Nid) -OnError $OnError -Script {
        param($path, $nid)
        $reader = [MailboxTransfer.PstReader]::Open($path)
        try { $reader.GetPreview($nid, 20000) } finally { $reader.Dispose() }
    } -OnDone {
        param($res)
        & $OnDone (ConvertTo-MbtGuiPstPreview @($res)[0] $label)
    }.GetNewClosure()
}

function ConvertTo-MbtGuiPstPreview {
    <# A PstMessagePreview of the engine as what the reading pane shows. #>
    param($Preview, [string]$Location)
    $p = $Preview
    $att = @($p.Attachments | ForEach-Object { $a = $_; if ($a.IsMessage) { "$($a.Name)   (attached message)" } else { "$($a.Name)   $(if ($a.Size -ge 0) { Format-MbtGuiBytes $a.Size } else { 'size unknown' })" } }) -join "`n"
    $body = [string]$p.BodyText
    if ($p.BodySource -eq 'html') { $body = "(the text of the HTML message)`n`n$body" } elseif ($p.BodySource -eq 'rtf') { $body = "(the text of the formatted message)`n`n$body" }
    if ($p.BodyTruncated) { $body += "`n`n(the message goes on: only the start is shown)" }
    foreach ($n in @($p.Notes)) { $body += "`n`n($n)" }
    if (-not $body.Trim()) { $body = '(this message has no text)' }
    $to = @(@($p.To) + @($p.Cc | ForEach-Object { "Cc: $_" })) -join '; '
    $when = if ($p.ReceivedUtc) { ([datetime]$p.ReceivedUtc).ToString('yyyy-MM-dd HH:mm') + ' UTC' } else { '' }
    [pscustomobject]@{ Subject = $(if ($p.Subject) { [string]$p.Subject } else { '(no subject)' }); From = [string]$p.FromName; Address = [string]$p.FromAddress; To = $to; Received = $when; Location = $Location
        Details = ('{0}{1}' -f $p.MessageClass, $(if ($p.SizeBytes -ge 0) { ' - ' + (Format-MbtGuiBytes $p.SizeBytes) } else { '' })); Attachment = $att; Body = $body; Available = $true }
}

#endregion

#region Check of a PST (a child process, read only) ----------------------------------------------------------------------------------------------------

function Start-MbtGuiPstCheckJob {
    <# -Mode CheckPst in a child, its events read by the tick (the poller) until the result. #>
    param([string]$Path, [scriptblock]$OnDone, [scriptblock]$OnError)
    $g = $script:Gui; $s = $g.S
    $job = New-MbtGuiJob -WorkFolder $s.WorkFolder -Task 'check' -EntryScript $g.EntryScript
    $run = New-MbtGuiRun -Task 'check'
    $saved = $s.RpPath; $s.RpPath = $Path
    foreach ($c in (Get-MbtGuiCommand -Task check -S $s -Job $job)) { $job.Queue.Enqueue($c) }
    $s.RpPath = $saved
    $null = Start-MbtGuiProcess -Job $job
    $started = [datetime]::UtcNow
    $g.Pollers.Add({
            $null = Step-MbtGuiJob -Job $job -Run $run
            if (-not $run.Result) { return $false }
            $res = $run.Result; $c = $res.Counts
            if ($run.Error -and -not $run.Items.Count) { & $OnError ($run.Error.Msg + " " + $run.Error.Tech); return $true }
            $findings = foreach ($i in $run.Items) { @{ Severity = $(if ($i.State -eq 'error') { 'Error' } else { 'Warning' }); Text = $i.Msg; Code = $i.Key; Count = [int]$i.Total; Level = (Get-MbtGuiFindingLevel $i.Key) } }
            $levels = @($findings | ForEach-Object { $_.Level })
            $rec = if ('Salvage' -in $levels) { 'Salvage' } elseif ('Rebuild' -in $levels) { 'Rebuild' } else { 'Recompute' }
            & $OnDone ([pscustomobject]@{ Path = $Path; SizeBytes = [long](Get-MbtGuiProp $c 'fileBytes' 0); Elapsed = (Format-MbtGuiDuration (([datetime]::UtcNow - $started).TotalSeconds)); Errors = [int](Get-MbtGuiProp $c 'errors' 0); Warnings = [int](Get-MbtGuiProp $c 'warnings' 0); ExitCode = [int]$res.Exit; Recommended = $rec; Findings = @($findings) })
            return $true
        }.GetNewClosure())
}

function Get-MbtGuiFindingLevel {
    <# The lightest repair level that deals with a finding of the checker. The checker does not say it per finding: counters, sizes and
       list lines are rewritten at every level (Recompute); a missing link between nodes needs the nodes to be read again (Rebuild); the
       block and page level (BLOCK-, BBT-, NBT-, HDR-) needs the scan of the file (Salvage). One place, so the rule can be corrected. #>
    param([string]$Code)
    switch -Regex ($Code) {
        '^(BBT|NBT|BLOCK|PAGE|HDR|AMAP|PMAP|FMAP|FPMAP|CRC)' { return 'Salvage' }
        '(NOT-LISTED|PARENT|ORPHAN|MISSING|HIERARCHY)' { return 'Rebuild' }
        default { return 'Recompute' }
    }
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxbJX50VExlo4m
# O1QSo2FpzrNITYNiUh4Fhl2Qc6DkMqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# wZislcdCZuS6MYIYETCCGA0CAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCBRRNt43kt
# zGTpCXXaz4eJvVaZ1vH165G5Odhf7eY2FDANBgkqhkiG9w0BAQEFAASCAgCsH1GC
# vuBkYCPUUANNPFJoBTesiVMGrA1zjEA5jU8bHXHl+fQf0/RSplMWfJ1GIDeLDJZv
# OGHbMBMF+XL5sPGiNSs+3psC+snUw2zrJoLRFyH78zKZuZu5Tg+DBoHOpLEYkGgr
# P7SolEqZIg4ZVJ6r1oGpEvc+tK2SP2hikyP4KHizO8R0DwbMXByujIjK62FRhwDp
# l9Z3p48tfgZa3Cb08DpaeOQW47yB0YxbT+wqm6lJCsDkb9QkpuTuBd2Qnd/3BgmT
# CUkrYYWKBZsojDw73k98bCtLnxIhspaVT7UmzFjnVAIXCCTagNr8Ox63eDqqMvBY
# wdGYeM1qsym+sqZApN6yYDD6o1hLa46dnS2blCepsVqx0rGl7o8KVNqMW9tl2SEd
# 75Rr3p/nMlcZ/47PHxRSXTwA/77TUPzo0eRopiRezuP143U9kSpFJZCBCIwzUIW9
# 2dyIq/VB2xaPdR7Zj+seELFyQRPEcyLMDQDYzhp7MydpKRI7wlq/RFykxhVJKxAU
# 8aTVn9RpbKD1MPAvFL8yRsYA/Nt15DYi6eO7edFouEupPAwXhN60J6zdiLUEzdkL
# ueF+Q1E9pfDxDviSBKY9zQbU8NUkakqVuxXwWVjirF/D4IMZxep4UuT74rK5HNi0
# xrLoUlUPl+jH/q9sndVGDO2P0VoVkxsPHWDtnKGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIKIkpPnuREEcIpsKjv/w2eVlKxkS8lDj5At+1MOUbJy1AghMC0bRPmFp
# +RgPMjAyNjExMTAwODIyMTRaMAMCAQECEF94SOXfAlKhUAzqpiLK2FmgghGnMIIE
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
# wzSPOzGCAlYwggJSAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjIxNFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIKWA
# SfVHcZ/34po45ZUdEfpDl9Q4tzSKjvdKfowpWo6/MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEA76cRNuvCn9FCLV7bktnPaCzOiYvzxhSdC5SdFUdkEtoCIAJI
# jHeaer2iE9KxoLz8LrW7RRU9J101KxUmVxZu5Xv7
# SIG # End signature block
