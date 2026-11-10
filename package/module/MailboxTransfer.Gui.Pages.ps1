#Requires -Version 7.4
<#
.SYNOPSIS
    Mailbox Transfer console - the pages Task, From, To, Plan (dot-sourced by Start-MailboxTransferConsole.ps1).

.DESCRIPTION
    One function per page, as the mockup MailboxTransfer-Console.html: each returns the element of the page and sets the action bar. The texts are
    the ones of CONSOLE-SPEC.md. A page is built again when the state changes (the browse cards live on, see Gui.Browse.ps1).
    Features that depend on the engine (PST import, EML, shared mailboxes) read $script:Gui.Engine.Capability: a missing one is shown disabled with
    "Coming in this version", or left out, so the page lights up by itself after the merge of the engine work.

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

$script:PageBuilders = @{
    start = { Build-MbtGuiStartPage }
    exFrom = { Build-MbtGuiExportFromPage }
    exTo = { Build-MbtGuiExportToPage }
    exPlan = { Build-MbtGuiExportPlanPage }
    imFrom = { Build-MbtGuiImportFromPage }
    imTo = { Build-MbtGuiImportToPage }
    imPlan = { Build-MbtGuiImportPlanPage }
    rpFrom = { Build-MbtGuiRepairFromPage }
    rpTo = { Build-MbtGuiRepairToPage }
    rpPlan = { Build-MbtGuiRepairPlanPage }
    opFrom = { Build-MbtGuiOpenPage }
    cpTo = { Build-MbtGuiCopyToPage }
    cpPlan = { Build-MbtGuiCopyPlanPage }
    run = { Build-MbtGuiRunPage }
    done = { Build-MbtGuiDonePage }
}

#region Shared pieces ------------------------------------------------------------------------------

function Get-MbtGuiCapability {
    param([string]$Name)
    $cap = $script:Gui.Engine.Capability
    return [bool]($cap.ContainsKey($Name) -and $cap[$Name])
}

function New-MbtGuiBrowseFolderRow {
    <# A text box with Browse...: a folder or a file; ReadOnly: no. #>
    param([string]$Value, [ValidateSet('Folder', 'File', 'SaveFile')][string]$Kind, [string]$Title, [string]$Filter, [scriptblock]$OnChange, [string]$Automation)
    $box = New-GuiTextBox -Text $Value -Automation $Automation
    $state = @{ Box = $box; OnChange = $OnChange }
    if ($OnChange) { $box.Add_LostFocus({ & $state.OnChange $state.Box.Text }.GetNewClosure()); $box.Add_KeyDown({ param($s, $e) if ($e.Key -eq 'Return') { & $state.OnChange $state.Box.Text; $e.Handled = $true } }.GetNewClosure()) }
    $btn = New-GuiButton -Text 'Browse...' -OnClick {
        $picked = switch ($Kind) {
            'Folder' { Select-MbtGuiFolder -Title $Title -InitialPath $state.Box.Text }
            'SaveFile' { Select-MbtGuiFile -Title $Title -Filter $Filter -InitialPath $state.Box.Text -Save }
            default { Select-MbtGuiFile -Title $Title -Filter $Filter -InitialPath $state.Box.Text }
        }
        if ($picked) { $state.Box.Text = $picked; if ($state.OnChange) { & $state.OnChange $picked } }
    }.GetNewClosure()
    $btn.Margin = [Windows.Thickness]::new(0)
    return (New-GuiRow -Box $box -Button $btn)
}

function Get-MbtGuiBrowser {
    <# The browse card of a use (export, import, open), created once. #>
    param([ValidateSet('export', 'import', 'open')][string]$Use)
    $g = $script:Gui
    if ($g.Browsers.ContainsKey($Use)) { return $g.Browsers[$Use] }
    switch ($Use) {
        'export' { return (New-MbtGuiBrowser -Name 'export' -Title 'Browse' -Info 'Live view of the mailbox. Read only.' -TrayEmpty 'Nothing ticked yet.') }
        'import' { return (New-MbtGuiBrowser -Name 'import' -Title 'What is inside' -Info 'Read only. The default is to import everything. Tick only if you want part of it.' -Badge 'Read only' -TrayEmpty 'Nothing ticked: **everything will be imported**.' -Kind source) }
        'open' { return (New-MbtGuiBrowser -Name 'open' -Title 'Browse' -Info 'Folders, messages and attachments, like a mailbox.' -Badge 'Read only' -TrayEmpty 'Nothing ticked yet.' -Kind source) }
    }
}

function Place-MbtGuiBrowser {
    <# Puts the card of a browser in a page (it may still belong to the page that was replaced). #>
    param($Browser, [Windows.Controls.Panel]$Into)
    $root = $Browser.Root
    $parent = $root.Parent
    if ($parent -is [Windows.Controls.Panel]) { $parent.Children.Remove($root) }
    elseif ($parent -is [Windows.Controls.Decorator]) { $parent.Child = $null }
    elseif ($parent -is [Windows.Controls.ContentControl]) { $parent.Content = $null }
    [void]$Into.Children.Add($root)
}

function Format-MbtGuiSelectionFooter {
    param($Selection, [bool]$Signed)
    if ($Selection.Items -gt 0) { return ('**{0} items** ticked in {1} folder{2}, {3}' -f (Format-MbtGuiSelectionItems $Selection), $Selection.Folders, $(if ($Selection.Folders -gt 1) { 's' } else { '' }), (Format-MbtGuiSelectionSize $Selection)) }
    if ($Signed) { return 'Tick what to keep, then continue.' }
    return 'Sign in first.'
}

function New-MbtGuiSelectionCard {
    param($Selection)
    if ($Selection.Items -gt 0) {
        $body = New-GuiFacts @(@('Items', (Format-MbtGuiSelectionItems $Selection)), @('Size', (Format-MbtGuiSelectionSize $Selection)), @('Folders', [string]$Selection.Folders))
    } else { $body = New-GuiText 'Tick a folder or some messages in the browser.' 'Hint' }
    return (New-GuiCard -Title 'Your selection' -Content @($body))
}

#endregion

#region Start ----------------------------------------------------------------------------------------

function New-MbtGuiTaskCard {
    param([string]$Key, [bool]$Selected)
    $t = $script:Tasks[$Key]
    $btn = [Windows.Controls.Primitives.ToggleButton]::new()
    $btn.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'MbtTaskCard')
    $btn.IsChecked = $Selected
    $btn.Tag = $Key
    [Windows.Automation.AutomationProperties]::SetName($btn, $t.Title)
    $g = [Windows.Controls.Grid]::new()
    $head = [Windows.Controls.StackPanel]::new(); $head.Orientation = 'Horizontal'
    $tile = [Windows.Controls.Border]::new(); $tile.Width = 40; $tile.Height = 40; $tile.CornerRadius = [Windows.CornerRadius]::new(8)
    Set-MbtGuiResource $tile ([Windows.Controls.Border]::BackgroundProperty) 'MbtAccentSoft'
    $ic = New-GuiIcon $t.Icon 22; $ic.HorizontalAlignment = 'Center'; $tile.Child = $ic
    $titles = [Windows.Controls.StackPanel]::new(); $titles.Margin = [Windows.Thickness]::new(12, 0, 0, 0); $titles.VerticalAlignment = 'Center'
    $eb = New-GuiText $t.Eyebrow 'Small' -NoWrap; $eb.FontWeight = [Windows.FontWeights]::SemiBold; Set-MbtGuiResource $eb ([Windows.Controls.TextBlock]::ForegroundProperty) 'MbtBrandText'
    $ti = New-GuiText $t.Title 'Body' -NoWrap; $ti.FontSize = 15; $ti.FontWeight = [Windows.FontWeights]::SemiBold
    $titles.Children.Add($eb) | Out-Null; $titles.Children.Add($ti) | Out-Null
    $head.Children.Add($tile) | Out-Null; $head.Children.Add($titles) | Out-Null
    $desc = New-GuiText $t.Text 'Muted'; $desc.FontSize = 13; $desc.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
    $col = New-GuiStack @($head, $desc)
    $g.Children.Add($col) | Out-Null
    $tick = [Windows.Controls.Border]::new(); $tick.Width = 22; $tick.Height = 22; $tick.CornerRadius = [Windows.CornerRadius]::new(11); $tick.HorizontalAlignment = 'Right'; $tick.VerticalAlignment = 'Top'
    Set-MbtGuiResource $tick ([Windows.Controls.Border]::BackgroundProperty) 'AccentFillColorDefaultBrush'
    $ti2 = New-GuiIcon 'check' 12 'TextOnAccentFillColorPrimaryBrush'; $ti2.HorizontalAlignment = 'Center'; $tick.Child = $ti2
    $tick.Visibility = $(if ($Selected) { 'Visible' } else { 'Collapsed' })
    $g.Children.Add($tick) | Out-Null
    $btn.Content = $g
    return $btn
}

function Build-MbtGuiStartPage {
    $g = $script:Gui; $s = $g.S; $engine = $g.Engine
    $sel = $s.Sel
    $grid = [Windows.Controls.Primitives.UniformGrid]::new(); $grid.Columns = 4
    foreach ($k in @($script:Tasks.Keys | Where-Object { -not (Get-MbtGuiProp $script:Tasks[$_] 'Hidden') })) {
        $card = New-MbtGuiTaskCard -Key $k -Selected ($sel -eq $k)
        $card.Margin = [Windows.Thickness]::new(0, 0, $(if ($k -eq 'open') { 0 } else { 12 }), 0)
        $card.Add_Click({ param($b, $e) $g.S.Sel = [string]$b.Tag; Show-MbtGuiPage })
        $grid.Children.Add($card) | Out-Null
    }
    $left = New-GuiCard -Title 'Continue a job' -Info ('Found in your work folder ' + $s.WorkFolder) -Content @((New-MbtGuiJobList))
    $right = if ($sel) {
        $t = $script:Tasks[$sel]
        $ol = [Windows.Controls.StackPanel]::new()
        $i = 1
        foreach ($w in $t.Will) { $row = New-GuiText ("$i.  $w") 'Body'; $row.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $ol.Children.Add($row) | Out-Null; $i++ }
        New-GuiCard -Title 'What will happen' -Content @($ol, (New-GuiText $t.Safe 'Hint'))
    } else { New-MbtGuiReadinessCard }
    $cols = New-GuiGrid -Children @($left, $right) -Widths @('1.4*', '1*')
    $stack = New-GuiStack @((New-GuiText 'What do you want to do?' 'H2'), $grid, (New-GuiStack @($cols) -Margin '0,12,0,0'))
    if ($sel) { $stack.Children.Add((New-MbtGuiReadinessCard)) | Out-Null }
    $footer = if ($sel) { "You chose **$($script:Tasks[$sel].Title)**." } else { 'Pick what you want to do. The console guides you from there.' }
    Set-MbtGuiActions -Back $false -Footer $footer -Label 'Continue' -Enabled ([bool]$sel) -OnClick {
        $s = $script:Gui.S
        $s.Task = $s.Sel; $s.Step = 1
        Initialize-MbtGuiTask
        Show-MbtGuiPage
    }
    return $stack
}

function New-MbtGuiJobList {
    $g = $script:Gui
    $jobs = @()
    try { $jobs = @(& $g.Engine.Jobs $g.S.WorkFolder) } catch { Write-Verbose "jobs: $_" }
    if (-not $jobs.Count) { return (New-GuiText 'No job yet. A job you stop, or that is not finished, is listed here so you can continue it.' 'Hint') }
    $stack = [Windows.Controls.StackPanel]::new()
    $first = $true
    foreach ($j in $jobs) {
        $row = [Windows.Controls.Grid]::new()
        $row.Margin = [Windows.Thickness]::new(0, 0, 0, 0)
        foreach ($w in 'Auto', '*', 'Auto') { $cd = [Windows.Controls.ColumnDefinition]::new(); $cd.Width = if ($w -eq '*') { [Windows.GridLength]::new(1, 'Star') } else { [Windows.GridLength]::Auto }; $row.ColumnDefinitions.Add($cd) }
        $glyphName = switch ($j.State) { 'ok' { 'check' } 'warn' { 'warn' } default { 'pause' } }
        $brush = switch ($j.State) { 'ok' { 'MbtSuccess' } 'warn' { 'MbtCaution' } default { 'MbtBrandText' } }
        $ic = New-GuiIcon $glyphName 14 $brush; $ic.Margin = [Windows.Thickness]::new(0, 0, 12, 0); $ic.VerticalAlignment = 'Center'
        $txt = [Windows.Controls.StackPanel]::new(); [Windows.Controls.Grid]::SetColumn($txt, 1)
        $t1 = New-GuiText $j.Title 'Body'; $t1.FontSize = 13; $t1.FontWeight = [Windows.FontWeights]::SemiBold
        $t2 = New-GuiText $j.Sub 'Small'
        $txt.Children.Add($t1) | Out-Null; $txt.Children.Add($t2) | Out-Null
        $btn = New-GuiButton -Text $j.Button -Small -OnClick { Resume-MbtGuiJob $j }.GetNewClosure()
        $btn.Margin = [Windows.Thickness]::new(12, 0, 0, 0); $btn.VerticalAlignment = 'Center'
        [Windows.Controls.Grid]::SetColumn($btn, 2)
        $row.Children.Add($ic) | Out-Null; $row.Children.Add($txt) | Out-Null; $row.Children.Add($btn) | Out-Null
        $wrap = [Windows.Controls.Border]::new(); $wrap.Child = $row; $wrap.Padding = [Windows.Thickness]::new(0, 9, 0, 9)
        if (-not $first) { $wrap.BorderThickness = [Windows.Thickness]::new(0, 1, 0, 0); Set-MbtGuiResource $wrap ([Windows.Controls.Border]::BorderBrushProperty) 'DividerStrokeColorDefaultBrush' }
        $stack.Children.Add($wrap) | Out-Null
        $first = $false
    }
    return $stack
}

function New-MbtGuiReadinessCard {
    $g = $script:Gui
    $lights = @()
    try { $lights = @((& $g.Engine.Readiness).Lights) } catch { $lights = @(@{ Name = 'The engine could not be checked'; Detail = $_.Exception.Message; Ok = $false }) }
    $red = @($lights | Where-Object { -not $_.Ok })
    $list = [Windows.Controls.StackPanel]::new()
    foreach ($l in $lights) {
        $row = [Windows.Controls.StackPanel]::new(); $row.Orientation = 'Horizontal'; $row.Margin = [Windows.Thickness]::new(0, 3, 0, 3)
        $dot = [Windows.Controls.Border]::new(); $dot.Width = 18; $dot.Height = 18; $dot.CornerRadius = [Windows.CornerRadius]::new(9)
        Set-MbtGuiResource $dot ([Windows.Controls.Border]::BackgroundProperty) $(if ($l.Ok) { 'MbtSuccessBackground' } else { 'MbtCriticalBackground' })
        $ic = New-GuiIcon $(if ($l.Ok) { 'check' } else { 'x' }) 10 $(if ($l.Ok) { 'MbtSuccess' } else { 'MbtCritical' }); $ic.HorizontalAlignment = 'Center'; $dot.Child = $ic
        $tx = New-GuiText ("**$($l.Name)**  " + [char]0x00B7 + " $($l.Detail)") 'Body' -NoWrap; $tx.FontSize = 13; $tx.Margin = [Windows.Thickness]::new(8, 0, 0, 0)
        $row.Children.Add($dot) | Out-Null; $row.Children.Add($tx) | Out-Null
        $list.Children.Add($row) | Out-Null
    }
    $list.Children.Add((New-GuiText 'A red light shows here by itself, with what to do. Check and repair, and Open, need no sign-in and no network.' 'Hint')) | Out-Null
    $more = New-GuiMore -Summary 'Show the lights' -Content @($list) -Open:($red.Count -gt 0)
    $intro = New-GuiText $(if ($red.Count) { 'Something needs your attention before the cloud tasks can run.' } else { 'Everything needed is in place. Only a red light needs your attention.' }) 'Hint'
    $intro.Margin = [Windows.Thickness]::new(0)
    return (New-GuiCard -Title 'This computer' -Right (New-GuiPill $(if ($red.Count) { 'Check' } else { 'Ready' }) $(if ($red.Count) { 'bad' } else { 'ok' })) -Content @($intro, $more))
}

function Initialize-MbtGuiTask {
    <# A task was chosen: what its pages start with. #>
    $g = $script:Gui; $s = $g.S
    $s.Outcome = $null; $s.PlanState = 'idle'; $s.PlanMsg = ''; $s.RStep = $null; $s.RObj = $null; $s.Confirm = ''
    $g.Run = $null; $g.Job = $null
    switch ($s.Task) {
        'export' { if (-not $s.OutFolder) { $s.OutFolder = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'MailboxTransfer' } }
        'import' { $s.Confirm = '' }
        'repair' { $s.RpCheck = $s.RpCheck }
    }
}

#endregion

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOcfZ+nmcO/XLY
# QJiAn3Mkbti45PTtjomJ7B4tymEafaCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDLdYWMCz8F
# OwZOkr0O8jOTEKXfgn+SDxwF+b8zNPbcZDANBgkqhkiG9w0BAQEFAASCAgCpemcm
# zfWxHf7AF5zBjAGhBN4f0hf9dPfCkQNykUdolUj21wWtnfk0OX9eZ4tKDjGj/HZx
# THpeWlWpov62HXBeROALOc6Mhyrgt1b2oUdCZIcwLzDv6I+dX+F9vK78PkvYEL21
# wQR3eW3rD0rvnhv+NsEp4HZYYtf+oAbbQ+t0rR6NySgml4ZdWQGqaHGs1PM9hahP
# m5eJWytxGxEOSA/mrcuoTAlgAmZs/RXtA0TE+s1e87NeRrRq7IjtO5Th9nZAMeqp
# WFyrBqe1mpD6pfiUP6WsZ401fI1UD7ZIJcR3XtobIWZMa1qnc8jqWWBaTniIrk9F
# xCOuDarj+bmZz8D+aDYvzhJwLv6/+dGFpX+xMyxW2rZxaPcisWzVUe15Ktc9Lgot
# 1UEzErnPv9o2P5ZoTb1Vfthq+TsJ6PjAIC3O+Rfr0HeT3ndzh3/NFSX+9FZ7aBmR
# 2UijKV3cxcmxQpG83/O06PPasYavQnpXkcrMtWGKuethCx43jE9sdMlViJ83mGhR
# RmT8bNGpsjDMDWCSXN7AsHNQG79hMjHdtbvNyfowljVIbZfE4o/PZg5lppGzi77D
# bREMdJOnRoVug9kzwWXkinWWRKJC8Irjtpg3Uemw/fpbnNSyTZr0qfTrJgxhl7yy
# otMuK2paJC1vE/dNDMM5ksojKF4ApJqzsXPY6KGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIPdDxRKMmRDD/A05/XX/P0/gcZzBwW0cIxk5AOIK8i/OAgh3h0sLyyEC
# VhgPMjAyNjExMTAwODIyMzJaMAMCAQECED2/MPE1bfc2V83rXvVXGwqgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDA4MjIzMlowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIBJ2
# ulEpNlz9ZkPpU53Jzz2Ji145ZuAbINAVUZxnUUDLMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiBuaUtEvrlSn91+K5U0rBb+mVU33sii+DRV8pZPB8bhhgIgOC/o
# 0RCs73e9+FEQpJKH+olyH01wNSPUDxIwvstZZm4=
# SIG # End signature block
