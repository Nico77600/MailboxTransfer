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

    $c.FolderTree.Add_SelectedItemChanged({ param($s, $e) if ($e.NewValue) { Open-MbtGuiFolder $e.NewValue } }.GetNewClosure())
    $c.MessageList.Add_SelectionChanged({ param($s, $e) if ($s.SelectedItem) { Show-MbtGuiMessage $b $s.SelectedItem } }.GetNewClosure())
    $c.PickAll.Add_Click({ param($s, $e) if ($b.Current) { $b.Current.TickOwnOnly([bool]$s.IsChecked); Sync-MbtGuiRows $b } }.GetNewClosure())
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

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAgZjvu7+qvflGO
# 8eALEogWxc+v64q9L1fJQ9oziiBm86CCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCABie0gBOPP
# 3NentbDZZw9vaP80uLUiwI78wLPSmmMi4jANBgkqhkiG9w0BAQEFAASCAgCb8WYn
# Ixm2V+AKBs8QtrtUEv8R7YcJ5ba7cr0/B7IlXvlOAPx21r49sgXLFD9Gucfs6k+s
# e1aJp3kUKVbIWIUV9bGmZ+cJ+OcJQ2A7mjJDloyrmnqYa/Gu4wLcgGwnl5exGOn4
# WGPY8ax0gOv2uNLKRNREk3gFSOFmkyulJ3Z8odI72NHenCOdkDbOgyg3Oudsf9dN
# eQeLvrfXcNlYsrFJVAY193Bm5ETMDQTTlbMINa5qcMtIN/VHwa7KtSLDR/y359QH
# YH5lz6t2cKmPTzIyfNPTaAEnTcoXKaINUknim9D7Z/8M4fZGQS2tEW/tIqvCYXU9
# eOqVcuFHyRaxSMXAm27vNqZa97cagQMDuWhAC6/Qkv21EDVQPkK7eqrKFCwMfNu8
# yq2sEMD1epLdd2VWt+LUWkYbwrmhtvgVtGg/j5Ann47Lwf2HrPuufOKHb8wllOnw
# Y2gbXbjvf57BdGPyNdlOG7rgINdWtiJ2aOKfPZP6jfLf81LpSoKH1t35qmcdOzs0
# AmAPqQv8dRejKVjNOatJgwr79+l7ApKWmh1h07ExqfNAmY6PYyRc/pw4gEHH5UAn
# vfFtMupR8Ys8dvgRoYEQiZEN/zhUf5WTRnmejgHyK237kTLTPRViXnfWeFBWwO5L
# 1x2L6WYUn5tIzZm2rsCuBoffP0HQuedAhWBYF6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIAZhbD/fMo7QSbRGxtWpDNEYtxd7UMJkhABjDroMafCzAghx8BAmz7Ks
# 3hgPMjAyNjExMTAwODIxNTlaMAMCAQECED+c5YmDuSwa24kD4gdTsS6gghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjE1OVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIE86
# FqU0RaW0Zh2Gn6y7lSm9sHHbc7pDdW+cdfsOmzsyMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEAqdNdnS6O44RyfzM17EbaeRBE6jt21Ypq+wgH3PFHXpICIDwz
# KYq9vJq64GMbMs1NEB6Ft+hKqDnAkcbLzKRBqeq1
# SIG # End signature block
