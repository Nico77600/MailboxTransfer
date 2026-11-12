#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - the browse card (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    Folder tree with tri-state boxes, message list, reading pane, selection tray (MailboxTransfer.Gui.Browse.xaml). One instance per use (export,
    import source, open); it lives as long as the window, so a tick survives a trip to another page. The data comes from the engine adapter
    ($script:Gui.Engine: Roots, Children, Items, Preview), read asynchronously, a level or a page at a time, never a blocking spinner.

    A selection is exactly what the engine's selection file holds: a folder ticked as a whole (with the folders below it), a folder's own messages,
    or messages picked one by one. Get-MbtGuiSelection walks the tree and returns those entries and the totals (an unknown size is counted
    apart, never summed).

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

$script:BrowserOf = $null
$script:BrowseWired = $false

function Initialize-MbtGuiBrowseEvents {
    <# The two static events of the tree: a node is opened (read its children), a tick changed (tray and footer follow). #>
    if ($script:BrowseWired) { return }
    $script:BrowseWired = $true
    $script:BrowserOf = [Collections.Generic.Dictionary[MailboxTransferGui.FolderNode, hashtable]]::new()
    [MailboxTransferGui.FolderNode]::add_ExpandRequested([Action[MailboxTransferGui.FolderNode]] { param($node) Read-MbtGuiChildren $node })
    [MailboxTransferGui.FolderNode]::add_SelectionChanged([Action] { Update-MbtGuiBrowseTrays })
}

function Get-MbtGuiRootNode {
    param([MailboxTransferGui.FolderNode]$Node)
    while ($Node.Parent) { $Node = $Node.Parent }
    return $Node
}

function New-MbtGuiBrowser {
    <# Loads the card. Title, Info and Badge (a pill text) sit in its header; TrayEmpty is the text of an empty tray (Markup **bold**). #>
    param([Parameter(Mandatory)][string]$Name, [string]$Title = 'Browse', [string]$Info = 'Tick folders or messages. The reading pane shows what is inside.', [string]$Badge, [string]$TrayEmpty = 'Nothing ticked yet.', [string]$Kind = 'live')
    Initialize-MbtGuiBrowseEvents
    $xaml = Get-Content -LiteralPath (Join-Path $script:GuiRoot 'module\MailboxTransfer.Gui.Browse.xaml') -Raw
    $root = [Windows.Markup.XamlReader]::Parse($xaml)
    $c = @{}
    foreach ($n in 'BrowseTitle', 'BrowseInfo', 'BrowseBadge', 'FolderTree', 'ListTitle', 'ListNote', 'ListHeader', 'PickAll', 'MessageList', 'ListMore', 'ListEmpty', 'ReadingHeader', 'ReadSubject', 'ReadFrom', 'ReadDetails', 'ReadAttachment', 'ReadAttachmentText', 'ReadBody', 'ReadEmpty', 'TrayPanel') {
        $c[$n] = $root.FindName($n)
    }
    $c.BrowseTitle.Text = $Title
    $c.BrowseInfo.Text = $Info
    if ($Badge) { $c.BrowseBadge.Content = New-GuiPill $Badge '' }
    $roots = [Collections.ObjectModel.ObservableCollection[MailboxTransferGui.FolderNode]]::new()
    $rows = [Collections.ObjectModel.ObservableCollection[MailboxTransferGui.MessageRow]]::new()
    $c.FolderTree.ItemsSource = $roots
    $c.MessageList.ItemsSource = $rows
    $b = @{ Name = $Name; Root = $root; C = $c; Roots = $roots; Rows = $rows; Current = $null; Next = $null; TrayEmpty = $TrayEmpty; Kind = $Kind; Ready = $false; OnChange = $null; Loading = $false; Pending = 0; Reading = $null; NoPicks = $false }
    $script:Gui.Browsers[$Name] = $b

    Add-MbtGuiHandler $c.FolderTree 'SelectedItemChanged' { param($sender, $e) if ($e.NewValue) { Open-MbtGuiFolder $e.NewValue } }
    Add-MbtGuiHandler $c.MessageList 'SelectionChanged' { param($sender, $e) if ($sender.SelectedItem) { Show-MbtGuiMessage $b $sender.SelectedItem } }
    Add-MbtGuiHandler $c.PickAll 'Click' { param($sender, $e) if ($b.Current) { $b.Current.TickOwnOnly([bool]$sender.IsChecked); Sync-MbtGuiRows $b } }
    Update-MbtGuiTray $b
    return $b
}

function Set-MbtGuiBrowserRoots {
    <# Gives the browser its roots (one per location, or the PST / store). #>
    param([Parameter(Mandatory)]$Browser, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Roots)
    $Browser.Roots.Clear(); $Browser.Rows.Clear(); $Browser.Current = $null
    foreach ($r in $Roots) {
        $script:BrowserOf[$r] = $Browser
        $Browser.Roots.Add($r)
        $r.IsExpanded = $true
    }
    $Browser.C.ListTitle.Text = ''
    $Browser.C.ListEmpty.Visibility = 'Visible'; $Browser.C.ListEmpty.Text = 'Pick a folder on the left.'
    Update-MbtGuiTray $Browser
}

function Read-MbtGuiChildren {
    <# A node is opened and its children are not read yet: ask the engine for the level below. #>
    param([MailboxTransferGui.FolderNode]$Node)
    $b = $script:BrowserOf[(Get-MbtGuiRootNode $Node)]
    if (-not $b) { return }
    $engine = $script:Gui.Engine
    $Node.Status = 'reading'
    & $engine.Children $Node { param($kids) $Node.SetChildren([MailboxTransferGui.FolderNode[]]@($kids)); $Node.Status = $null }.GetNewClosure() { param($msg) $Node.SetLoadFailed("Could not read the folders: $msg") }.GetNewClosure()
}

function Open-MbtGuiFolder {
    <# A folder is selected: read its first messages (a page at a time). #>
    param([MailboxTransferGui.FolderNode]$Node)
    $b = $script:BrowserOf[(Get-MbtGuiRootNode $Node)]
    if (-not $b) { return }
    $c = $b.C
    $b.Current = $Node; $b.Next = $null
    $b.Rows.Clear()
    $c.ReadingHeader.Visibility = 'Collapsed'; $c.ReadBody.Text = ''; $c.ReadEmpty.Visibility = 'Visible'; $c.ReadEmpty.Text = 'Select a message to read it.'
    if ($Node.IsRoot) {
        $c.ListTitle.Text = ''; $c.ListHeader.Visibility = 'Collapsed'; $c.ListNote.Content = $null
        $c.ListEmpty.Visibility = 'Visible'; $c.ListEmpty.Text = 'Pick a folder on the left.'
        return
    }
    $suffix = if ($Node.Count -gt 12) { ' ' + [char]0x00B7 + ' newest shown first' } else { '' }
    $c.ListTitle.Text = ('{0} {1} {2} items{3}' -f $Node.Name, [char]0x00B7, (Format-MbtGuiCount $Node.Count), $suffix)
    $c.ListHeader.Visibility = 'Visible'
    $c.PickAll.Visibility = $(if ($b.NoPicks) { 'Collapsed' } else { 'Visible' })
    $c.ListEmpty.Visibility = 'Collapsed'
    $headers = $false
    $root = Get-MbtGuiRootNode $Node
    if ($root.HeadersOnly) {
        $headers = $true
        $c.ListNote.Content = New-GuiMessage -Kind warn -Icon warn -Text 'The subject, sender and text of these messages cannot be read before the export (Microsoft offers headers only for this archive). They show after the export.'
        $c.ListNote.Content.Margin = [Windows.Thickness]::new(0, 0, 0, 4)
    } else { $c.ListNote.Content = $null }
    Read-MbtGuiItems $b $Node
}

function Read-MbtGuiItems {
    param($Browser, [MailboxTransferGui.FolderNode]$Node)
    $engine = $script:Gui.Engine
    $Browser.Loading = $true
    & $engine.Items $Node $Browser.Next {
        param($page)
        if ($Browser.Current -ne $Node) { return }
        $bw = $Browser; $nd = $Node
        foreach ($r in @($page.Rows)) {
            if ($Browser.NoPicks) { $r.CheckVisibility = 'Collapsed' }
            $r.SetCheckedQuietly((Test-MbtGuiRowTicked $nd $r))
            $r.Toggled = [Action[MailboxTransferGui.MessageRow]] { param($row) Set-MbtGuiRowTick $bw $row }.GetNewClosure()
            $Browser.Rows.Add($r)
        }
        $Browser.Next = $page.Next
        $Browser.Loading = $false
        $Browser.C.PickAll.IsChecked = $Node.OwnAll
        if ($page.Next) {
            $more = { Read-MbtGuiItems $bw $nd }.GetNewClosure()
            $Browser.C.ListMore.Content = New-GuiLink ("Load more ({0} left)" -f (Format-MbtGuiCount ([math]::Max(0, $Node.Count - $Browser.Rows.Count)))) $more
        } else { $Browser.C.ListMore.Content = $null }
        if ($Browser.Rows.Count -eq 0) { $Browser.C.ListEmpty.Visibility = 'Visible'; $Browser.C.ListEmpty.Text = 'No message in this folder.' }
        if ($page.Note) { $Browser.C.ListNote.Content = New-GuiMessage -Kind warn -Icon warn -Text $page.Note }
    }.GetNewClosure() {
        param($msg)
        $Browser.Loading = $false
        $Browser.C.ListEmpty.Visibility = 'Visible'; $Browser.C.ListEmpty.Text = "The messages could not be read: $msg"
    }.GetNewClosure()
}

function Test-MbtGuiRowTicked {
    param([MailboxTransferGui.FolderNode]$Node, [MailboxTransferGui.MessageRow]$Row)
    return ($Node.OwnAll -or $Node.Picks.ContainsKey($Row.Id))
}

function Set-MbtGuiRowTick {
    <# A message box was (un)ticked in the list. #>
    param($Browser, [MailboxTransferGui.MessageRow]$Row)
    $node = $Browser.Current
    if (-not $node) { return }
    if ($node.OwnAll) {
        # the folder was ticked as a whole and one message is unticked: from now on the listed messages are the picks
        $node.TickOwnOnly($false)
        foreach ($r in $Browser.Rows) { if ($r.IsChecked -or $r -eq $Row) { } }
        foreach ($r in $Browser.Rows) { if ($r.Id -ne $Row.Id) { $node.Picks[$r.Id] = $r.SizeBytes; $node.PickRows[$r.Id] = $r.Tag } }
        $Browser.C.ListNote.Content = New-GuiMessage -Kind info -Icon info -Text 'Only the messages listed here stay ticked.'
        $Browser.C.ListNote.Content.Margin = [Windows.Thickness]::new(0, 0, 0, 4)
    } elseif ($Row.IsChecked) { $node.Picks[$Row.Id] = $Row.SizeBytes; $node.PickRows[$Row.Id] = $Row.Tag }
    else { [void]$node.Picks.Remove($Row.Id); [void]$node.PickRows.Remove($Row.Id) }
    $node.Refresh()
    Sync-MbtGuiRows $Browser
    [MailboxTransferGui.FolderNode]::Notify()
}

function Sync-MbtGuiRows {
    <# The boxes of the list follow the node (after a tick in the tree, or "select all"). #>
    param($Browser)
    $node = $Browser.Current
    if (-not $node) { return }
    foreach ($r in $Browser.Rows) { $r.SetCheckedQuietly((Test-MbtGuiRowTicked $node $r)) }
    $Browser.C.PickAll.IsChecked = $node.OwnAll
}

function Show-MbtGuiMessage {
    param($Browser, [MailboxTransferGui.MessageRow]$Row)
    $c = $Browser.C
    $node = $Browser.Current
    if ($Row.SubjectNa) {
        $c.ReadingHeader.Visibility = 'Collapsed'; $c.ReadBody.Text = ''
        $c.ReadEmpty.Visibility = 'Visible'; $c.ReadEmpty.Text = "Preview available after the export.`nThis archive gives only sizes and dates before the export."
        return
    }
    $c.ReadEmpty.Visibility = 'Collapsed'
    $c.ReadingHeader.Visibility = 'Visible'
    $c.ReadSubject.Text = $Row.Subject; $c.ReadFrom.Text = $Row.From
    $c.ReadDetails.Text = 'Reading the message...'; $c.ReadBody.Text = ''; $c.ReadAttachment.Visibility = 'Collapsed'
    $engine = $script:Gui.Engine
    & $engine.Preview $Row $node {
        param($p)
        $c.ReadSubject.Text = $p.Subject
        $c.ReadFrom.Text = "$($p.From)  <$($p.Address)>"
        $c.ReadDetails.Text = "To: $($p.To)`nReceived: $($p.Received)`nLocation: $($p.Location)`nDetails: $($p.Details)"
        $c.ReadBody.Text = $p.Body
        if ($p.Attachment) { $c.ReadAttachment.Visibility = 'Visible'; $c.ReadAttachmentText.Text = $p.Attachment } else { $c.ReadAttachment.Visibility = 'Collapsed' }
    }.GetNewClosure() { param($msg) $c.ReadDetails.Text = "The message could not be read: $msg" }.GetNewClosure()
}

#region Selection ---------------------------------------------------------------------------

function Get-MbtGuiSelection {
    <#
        What is ticked in a set of root nodes, as the engine's selection file holds it: Entries (a folder as a whole with the folders below it, a folder's
        own messages, or picked messages) and the totals. A folder of unknown size is counted apart (UnknownSize), never summed; a ticked folder whose
        subfolders are not read yet is flagged (SubtreeUnread), its totals are then a minimum.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Roots)
    $entries = [Collections.Generic.List[object]]::new()
    $t = @{ Items = 0L; Bytes = 0L; Folders = 0; UnknownSize = 0; SubtreeUnread = $false }
    $add = {
        param($node, $items, $bytes, $unknown)
        $t.Items += $items; $t.Folders++
        if ($unknown) { $t.UnknownSize++ } else { $t.Bytes += $bytes }
    }
    $walk = $null
    $walk = {
        param([MailboxTransferGui.FolderNode]$n)
        if ($n.IsRoot) { foreach ($k in $n.Children) { & $walk $k }; return }
        if ($n.Check -eq $true) {
            $items = 0L; $bytes = 0L; $unknown = $false
            foreach ($d in $n.Descendants()) {
                if ($d.Count -le 0) { continue }
                $items += $d.Count
                if ($d.SizeBytes -lt 0) { $unknown = $true } else { $bytes += $d.SizeBytes }
                if ($d.ChildrenNotRead) { $t.SubtreeUnread = $true }
            }
            if ($n.ChildrenNotRead) { $t.SubtreeUnread = $true }
            $entries.Add(@{ Node = $n; Kind = 'tree'; Items = $items }); & $add $n $items $bytes $unknown
            return
        }
        if ($n.Check -eq $null) {
            if ($n.OwnAll -and $n.Count -gt 0) { $entries.Add(@{ Node = $n; Kind = 'own'; Items = $n.Count }); & $add $n $n.Count $(if ($n.SizeBytes -ge 0) { $n.SizeBytes } else { 0 }) ($n.SizeBytes -lt 0) }
            elseif ($n.Picks.Count -gt 0) { $sum = 0L; foreach ($v in $n.Picks.Values) { $sum += $v }; $entries.Add(@{ Node = $n; Kind = 'picks'; Items = $n.Picks.Count }); & $add $n $n.Picks.Count $sum $false }
            foreach ($k in $n.Children) { if ($k.Id -ne "`0placeholder") { & $walk $k } }
        }
    }
    foreach ($r in $Roots) { & $walk $r }
    [pscustomobject]@{ Entries = $entries.ToArray(); Items = $t.Items; Bytes = $t.Bytes; Folders = $t.Folders; UnknownSize = $t.UnknownSize; SubtreeUnread = $t.SubtreeUnread }
}

function Clear-MbtGuiSelection {
    param($Browser)
    foreach ($r in $Browser.Roots) { foreach ($k in @($r.Children)) { if ($k.Id -ne "`0placeholder") { $k.TickWholeFolder($false) } } }
    Sync-MbtGuiRows $Browser
}

function Update-MbtGuiTray {
    param($Browser)
    $p = $Browser.C.TrayPanel
    $p.Children.Clear()
    $sel = Get-MbtGuiSelection -Roots @($Browser.Roots)
    if ($sel.Items -gt 0 -or $sel.Folders -gt 0) {
        $txt = '**{0} items** in {1} folder{2} {3} {4}' -f (Format-MbtGuiCount $sel.Items), $sel.Folders, $(if ($sel.Folders -gt 1) { 's' } else { '' }), [char]0x00B7, (Format-MbtGuiSelectionSize $sel)
        if ($sel.SubtreeUnread) { $txt += ' (folders below are counted as you open them)' }
        $t = New-GuiText $txt 'Body' -NoWrap; $t.FontSize = 13; $t.VerticalAlignment = 'Center'
        $p.Children.Add($t) | Out-Null
        $clear = New-GuiButton -Text 'Clear' -Small -OnClick { Clear-MbtGuiSelection $Browser }.GetNewClosure()
        $clear.Margin = [Windows.Thickness]::new(14, 0, 0, 0)
        $p.Children.Add($clear) | Out-Null
    } else {
        $t = New-GuiText $Browser.TrayEmpty 'Muted' -NoWrap; $t.FontSize = 13
        $p.Children.Add($t) | Out-Null
    }
}

function Update-MbtGuiBrowseTrays {
    <# A tick changed anywhere: the trays, and the page's primary button and footer. #>
    $g = $script:Gui
    foreach ($b in $g.Browsers.Values) { Update-MbtGuiTray $b }
    if ($g.OnSelectionChanged) { & $g.OnSelectionChanged }
}

#endregion
