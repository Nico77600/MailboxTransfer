#Requires -Version 7.6
<#
.SYNOPSIS
    Mailbox Transfer console - the pages Import, Check and repair a PST, Open (dot-sourced by Start-MailboxTransferConsole.ps1).

.NOTES
    Author  : Nicolas Fabert
#>
Set-StrictMode -Version Latest

#region Source (PST or store) ---------------------------------------------------------------------------------

function Open-MbtGuiSource {
    <# Checks a PST or a store at once, read only, and gives its tree to the browser. Use: import or open. #>
    param([ValidateSet('import', 'open')][string]$Use)
    $g = $script:Gui; $s = $g.S
    $kind = if ($Use -eq 'import') { $s.ImSrc } else { $s.OpSrc }
    $path = if ($Use -eq 'import') { $s.ImPath } else { $s.OpPath }
    $browser = Get-MbtGuiBrowser $Use
    if ([string]::IsNullOrWhiteSpace($path)) { return }
    # a PST or .eml files to import: the folders are what is chosen, not single messages (the engine takes folders: -IncludeFolder)
    $browser.NoPicks = ($Use -eq 'import' -and $kind -in 'pst', 'eml')
    if ($Use -eq 'import') { $s.ImChecking = $true; $s.ImInfo = $null } else { $s.OpInfo = $null }
    Set-MbtGuiBrowserRoots -Browser $browser -Roots @()
    Show-MbtGuiPage
    & $g.Engine.OpenSource $kind $path {
        param($info)
        if ($Use -eq 'import') { $s.ImInfo = $info; $s.ImChecking = $false } else { $s.OpInfo = $info }
        if ($info.Ok) {
            Set-MbtGuiBrowserRoots -Browser $browser -Roots @($info.Roots)
            if ($Use -eq 'import') { $browser.TrayEmpty = "Nothing ticked: **everything will be imported** ($(Format-MbtGuiQuantity $info.Items 'item'), $(Format-MbtGuiBytes $info.Bytes))."; $browser.EverythingText = "**Everything** will be imported ($(Format-MbtGuiQuantity $info.Items 'item'), $(Format-MbtGuiBytes $info.Bytes))."; Update-MbtGuiTray $browser }
        }
        Show-MbtGuiPage
    }.GetNewClosure() {
        param($msg)
        $m = ConvertTo-MbtGuiMessage $msg
        $bad = [pscustomobject]@{ Ok = $false; Level = 'bad'; Message = $m.Text; Items = 0; Bytes = 0; Roots = @(); Tech = $m.Tech }
        if ($Use -eq 'import') { $s.ImInfo = $bad; $s.ImChecking = $false } else { $s.OpInfo = $bad }
        Show-MbtGuiPage
    }.GetNewClosure()
}

function New-MbtGuiSourceCheckMessage {
    param($Info, [bool]$Checking, [string]$Kind)
    if ($Checking) { return (New-GuiMessage -Kind info -Icon clock -Text 'Checking the file. This only reads.') }
    if (-not $Info) { return $null }
    $k = switch ($Info.Level) { 'ok' { 'ok' } 'warn' { 'warn' } 'bad' { 'bad' } default { 'info' } }
    $after = @()
    if ($Info.PSObject.Properties['Tech'] -and $Info.Tech) { $after = @(New-GuiMore -Summary 'Technical message' -Content @((New-GuiText $Info.Tech 'Small'))) }
    return (New-GuiMessage -Kind $k -Icon $(switch ($k) { 'ok' { 'check' } 'bad' { 'x' } default { 'warn' } }) -Text $Info.Message -After $after)
}

function New-MbtGuiSourceCard {
    <# Radio PST / store, path, result of the check. Use: import or open. #>
    param([ValidateSet('import', 'open')][string]$Use)
    $g = $script:Gui; $s = $g.S
    $kindKey = if ($Use -eq 'import') { 'ImSrc' } else { 'OpSrc' }; $pathKey = if ($Use -eq 'import') { 'ImPath' } else { 'OpPath' }
    $pstOk = ($Use -eq 'open') -or (Get-MbtGuiCapability 'PstImport')
    $content = [Collections.Generic.List[object]]::new()
    $pstTitle = if ($pstOk) { 'A PST file' } else { 'A PST file (coming in this version)' }
    $content.Add((New-GuiOption -Type Radio -Group src -Title $pstTitle -Hint 'One file, or the first of several parts (-002, -003 ... are found by themselves).' -Checked ($s[$kindKey] -eq 'pst') -Disabled:(-not $pstOk) -Margin '0' -OnChange { $s[$kindKey] = 'pst'; $s[$pathKey] = ''; Set-MbtGuiBrowserRoots -Browser (Get-MbtGuiBrowser $Use) -Roots @(); if ($Use -eq 'import') { $s.ImInfo = $null } else { $s.OpInfo = $null }; Show-MbtGuiPage }.GetNewClosure()))
    $content.Add((New-GuiOption -Type Radio -Group src -Title 'An export store' -Hint 'A folder made by Browse and export.' -Checked ($s[$kindKey] -eq 'store') -OnChange { $s[$kindKey] = 'store'; $s[$pathKey] = ''; Set-MbtGuiBrowserRoots -Browser (Get-MbtGuiBrowser $Use) -Roots @(); if ($Use -eq 'import') { $s.ImInfo = $null } else { $s.OpInfo = $null }; Show-MbtGuiPage }.GetNewClosure()))
    if ($Use -eq 'import') {
        $emlOk = Get-MbtGuiCapability 'EmlImport'
        $content.Add((New-GuiOption -Type Radio -Group src -Title $(if ($emlOk) { 'A folder of EML files' } else { 'A folder of EML files (coming in this version)' }) -Hint 'One .eml file per message. The folders below it become folders of the mailbox.' -Checked ($s[$kindKey] -eq 'eml') -Disabled:(-not $emlOk) -OnChange { $s[$kindKey] = 'eml'; $s[$pathKey] = ''; Set-MbtGuiBrowserRoots -Browser (Get-MbtGuiBrowser $Use) -Roots @(); $s.ImInfo = $null; $s.TgtMode = 'new'; Show-MbtGuiPage }.GetNewClosure()))
    }
    $isPst = $s[$kindKey] -eq 'pst'
    $isEml = $s[$kindKey] -eq 'eml'
    $srcLabel = if ($isPst) { 'PST file' } elseif ($isEml) { 'EML folder' } else { 'Store folder' }
    $content.Add((New-GuiText $srcLabel 'Label'))
    $content.Add((New-MbtGuiBrowseFolderRow -Value $s[$pathKey] -Kind $(if ($isPst) { 'File' } else { 'Folder' }) -Title $srcLabel -Filter 'PST files (*.pst)|*.pst|All files (*.*)|*.*' -Automation 'Source' -OnChange { param($v) if ($v -and $v -ne $s[$pathKey]) { $s[$pathKey] = $v; Open-MbtGuiSource -Use $Use } }.GetNewClosure()))
    $info = if ($Use -eq 'import') { $s.ImInfo } else { $s.OpInfo }
    $checking = ($Use -eq 'import') -and $s.ImChecking
    $m = New-MbtGuiSourceCheckMessage $info $checking $s[$kindKey]
    if ($m) { $content.Add($m) }
    if ($isPst -and $Use -eq 'import') {
        $line = [Windows.Controls.WrapPanel]::new(); $line.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
        $t = New-GuiText 'A damaged PST? ' 'Small'; $t.Margin = [Windows.Thickness]::new(0, 0, 4, 0)
        $line.Children.Add($t) | Out-Null
        $lk = New-GuiLink 'Check and repair it first' { $s.Task = 'repair'; $s.Sel = 'repair'; $s.Step = $script:Flows.repair.IndexOf('rpFrom'); if ($s.ImPath) { $s.RpPath = $s.ImPath }; Show-MbtGuiPage }; $lk.FontSize = 12
        $line.Children.Add($lk) | Out-Null
        $content.Add($line)
    }
    $title = if ($Use -eq 'import') { 'Source' } else { 'File or store' }
    return (New-GuiCard -Title $title -Content $content.ToArray())
}

#endregion

#region Import ----------------------------------------------------------------------------------------------------

function Build-MbtGuiImportFromPage {
    $g = $script:Gui; $s = $g.S
    $browser = Get-MbtGuiBrowser 'import'
    $left = New-GuiStack @((New-MbtGuiSourceCard -Use import))
    if ($s.ImSrc -in 'pst', 'eml') {
        $what = if ($s.ImSrc -eq 'pst') { 'a PST' } else { 'a folder of EML files' }
        $left.Children.Add((New-GuiMessage -Kind info -Icon info -Text "**Everything imports the whole of $what.** Choose **Only what I tick** to import only the folders you tick: a ticked folder takes the folders below it. Single messages are not chosen here; the folders the tick leaves out are not created.")) | Out-Null
    }
    $right = [Windows.Controls.Grid]::new()
    Update-MbtGuiModeRow $browser
    Update-MbtGuiTray $browser
    if ($s.ImInfo -and $s.ImInfo.Ok) { Place-MbtGuiBrowser $browser $right }
    else {
        $ico = New-GuiIcon pst 22 'MbtBrandText'; $ico.HorizontalAlignment = 'Center'
        $t1 = New-GuiText 'Pick the PST file or the export. It is checked at once, and you can look inside.' 'Muted'; $t1.TextAlignment = 'Center'; $t1.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
        $body = New-GuiStack @($ico, $t1); $body.Margin = [Windows.Thickness]::new(10, 60, 10, 60)
        $right = New-GuiCard -Title 'What is inside' -Content @($body)
    }
    $ok = [bool]($s.ImInfo -and $s.ImInfo.Ok)
    Set-MbtGuiActions -Footer 'Pick what to import. The file is checked first.' -Label 'Continue to the mailbox' -Enabled $ok -OnClick { Move-MbtGuiStep 1 }
    $stack = New-GuiFillPage -Title 'What do you want to import?' -Columns (New-GuiTwoColumns -Left $left -Right $right -Fill)
    return $stack
}

function Test-MbtGuiSourceHasDeletedItems {
    <# Does the export store or the PST being imported hold a top folder named Recoverable Items? (The import puts it in a normal folder.) #>
    $s = $script:Gui.S
    if ($s.Task -ne 'import' -or -not $s.ImInfo) { return $false }
    $info = $s.ImInfo
    $roots = if ($info -is [Collections.IDictionary]) { if ($info.Contains('Roots')) { $info['Roots'] } } elseif ($info.PSObject.Properties['Roots']) { $info.Roots }
    foreach ($r in @($roots)) {
        foreach ($n in @($r.Descendants())) { if ($n.Name -eq 'Recoverable Items') { return $true } }
    }
    return $false
}

function Build-MbtGuiImportToPage {
    $g = $script:Gui; $s = $g.S
    $left = New-GuiStack @((New-MbtGuiSignInCard 'import'))
    $right = if ($s.Signed) {
        $items = [Collections.Generic.List[object]]::new()
        if ($s.AuthMode -eq 'app') { $items.Add((New-MbtGuiSharedMailboxSection -Use import)) }
        else { $items.Add((New-GuiMessage -Kind info -Icon shield -Text "**Into your own mailbox, $($s.Mailbox).** A sign-in as you writes only into your mailbox: its Primary mailbox or its Archive. To write into another mailbox, sign in as the application (Setup).")) }
        $items.Add((New-GuiText 'Part of the mailbox' 'Label'))
        $locs = @($s.Locations | Where-Object { $_.Kind -in 'Primary', 'Archive' })
        if ($s.ImShared) { $locs = @([pscustomobject]@{ Kind = 'Primary'; Label = "Primary mailbox of $($s.ImShared)"; Hint = 'The primary mailbox. Its archive is not reached by its address.' }) }
        if (-not $locs.Count) { $locs = @([pscustomobject]@{ Kind = 'Primary'; Label = 'Primary mailbox'; Hint = '' }, [pscustomobject]@{ Kind = 'Archive'; Label = 'Archive'; Hint = '' }) }
        $over = $false
        foreach ($l in $locs) {
            $kind = $l.Kind; $hint = [string]$l.Hint
            if ($l.PSObject.Properties['Quota'] -and $l.Quota -and $s.ImInfo -and $kind -eq $s.TgtLoc) {
                $free = [double]$l.Quota - [double]$l.Used
                if ($s.ImInfo.Bytes -gt $free) { $hint = "$(Format-MbtGuiBytes $l.Used) of $(Format-MbtGuiBytes $l.Quota) used. Importing $(Format-MbtGuiBytes $s.ImInfo.Bytes) would exceed the quota."; $over = $true }
            }
            $items.Add((New-GuiOption -Type Radio -Group loc -Title $l.Label -Hint $hint -Checked ($s.TgtLoc -eq $kind) -OnChange { $s.TgtLoc = $kind; Show-MbtGuiPage }.GetNewClosure()))
        }
        if ($over) { $items.Add((New-GuiMessage -Kind warn -Icon warn -Text '**This will probably not fit.** Microsoft refuses items once a mailbox is full; the import then ends with refused items. Choose another place, or free some room first.')) }
        $items.Add((New-GuiText 'Folder' 'Label'))
        if (-not $s.TgtName) { $stem = if ($s.ImPath) { [IO.Path]::GetFileNameWithoutExtension($s.ImPath) } else { 'export' }; $s.TgtName = "Imported $stem $((Get-Date).ToString('yyyy-MM-dd'))" }
        $newOpt = New-GuiOption -Type Radio -Group tm -Title 'A new folder (recommended)' -Hint 'Everything goes under one new folder. Easy to find, easy to move or delete later.' -Checked ($s.TgtMode -eq 'new') -Margin '0' -OnChange { $s.TgtMode = 'new'; Show-MbtGuiPage }
        $items.Add($newOpt)
        if ($s.TgtMode -eq 'new') {
            $name = New-GuiTextBox -Text $s.TgtName -Automation 'Folder name' -OnChange { param($x, $e) $s.TgtName = $x.Text }
            $name.Margin = [Windows.Thickness]::new(30, 6, 0, 0)
            $items.Add($name)
            $nameProblem = Test-MbtGuiFolderName $s.TgtName
            if ($nameProblem) { $items.Add((New-GuiMessage -Kind warn -Icon warn -Text $nameProblem)) }
        }
        $items.Add((New-GuiOption -Type Radio -Group tm -Title 'The same folder names' -Hint $(if ($s.ImSrc -eq 'eml') { 'Not for EML files: they always go under one folder.' } else { 'Items are added to folders with the same names. Existing items are left alone. Duplicates are possible.' }) -Checked ($s.TgtMode -eq 'same') -Disabled:($s.ImSrc -eq 'eml') -OnChange { $s.TgtMode = 'same'; Show-MbtGuiPage }))
        $items.Add((New-GuiMessage -Kind info -Icon shield -Text '**Nothing is deleted. Nothing is overwritten.** The import only adds items. A folder is never renamed.'))
        if (Test-MbtGuiSourceHasDeletedItems) {
            $items.Add((New-GuiMessage -Kind info -Icon shield -Text '**This source holds deleted items (Recoverable Items).** They are put in a normal folder named **Restored deleted items** (inside your new folder if you chose one), never back into Recoverable Items. They come back as ordinary items of that folder.'))
        }
        $mapBtn = New-GuiButton -Text 'Use a mapping file...' -Small -Icon open -OnClick { $f = Select-MbtGuiFile -Title 'Mapping file' -Filter 'Mapping (*.psd1;*.json)|*.psd1;*.json'; if ($f) { $s.Mapping = $f; Show-MbtGuiPage } }
        $mapLine = if ($s.Mapping) { New-GuiText "Mapping file: $($s.Mapping)" 'Small' } else { New-GuiText 'For renaming folders while importing (source path to target path). Optional.' 'Hint' }
        $items.Add((New-GuiMore -Summary 'More options' -Content @($mapBtn, $mapLine)))
        New-GuiCard -Title 'Where to import' -Info $s.Mailbox -Content $items.ToArray()
    } else {
        $t = New-GuiText 'Sign in to choose the mailbox.' 'Muted'; $t.TextAlignment = 'Center'; $t.Margin = [Windows.Thickness]::new(10, 40, 10, 40)
        New-GuiCard -Title 'Where to import' -Content @($t)
    }
    Set-MbtGuiActions -Footer $(if ($s.Signed) { 'Choose the mailbox and where in it the items go.' } else { 'Sign in first.' }) -Label 'Continue to Verify' -Enabled $s.Signed -OnClick { if ($s.TgtMode -eq 'new' -and (Test-MbtGuiFolderName $s.TgtName)) { Show-MbtGuiPage } else { Move-MbtGuiStep 1 } }
    return (New-GuiStack @((New-GuiText 'Into which mailbox?' 'H2'), (New-GuiTwoColumns -Left $left -Right $right)))
}

function New-MbtGuiImportConfirmCard {
    <# "Before anything is written": the target, the volume, what the import does, and ONE box that says exactly what and where. Start stays grey until it is ticked. #>
    param([long]$Items, [long]$Bytes, [string]$ItemsQuantity, [bool]$SizeUnknown)
    $g = $script:Gui; $s = $g.S
    $mid = [char]0x00B7
    $mailbox = if ($s.ImShared) { $s.ImShared } else { $s.Mailbox }
    $where = if ($s.TgtLoc -eq 'Archive' -and -not $s.ImShared) { 'Archive' } else { 'Primary' }
    $folderText = if ($s.TgtMode -eq 'new') { "new folder $($s.TgtName)" } else { 'folders with the same names' }
    $target = [Windows.Controls.Border]::new(); $target.CornerRadius = [Windows.CornerRadius]::new(8); $target.Padding = [Windows.Thickness]::new(14, 10, 14, 12); $target.Margin = [Windows.Thickness]::new(0, 4, 0, 0)
    Set-MbtGuiResource $target ([Windows.Controls.Border]::BackgroundProperty) 'MbtAccentSoft'
    $eyebrow = New-GuiText 'INTO' 'Small' -NoWrap; $eyebrow.FontWeight = [Windows.FontWeights]::SemiBold; Set-MbtGuiResource $eyebrow ([Windows.Controls.TextBlock]::ForegroundProperty) 'MbtBrandText'
    $big = New-GuiText $mailbox 'Big' -NoWrap; $big.FontSize = 22
    $line = [Windows.Controls.StackPanel]::new(); $line.Orientation = 'Horizontal'; $line.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
    $pill = New-GuiPill $where 'run'; $pill.Margin = [Windows.Thickness]::new(0, 0, 10, 0)
    $fi = New-GuiIcon folder 14; $fi.Margin = [Windows.Thickness]::new(0, 0, 6, 0)
    $ft = New-GuiText $(if ($s.TgtMode -eq 'new') { "in a new folder named **$($s.TgtName)**" } else { 'in the folders that have the same names' }) 'Body'; $ft.VerticalAlignment = 'Center'
    $line.Children.Add($pill) | Out-Null; $line.Children.Add($fi) | Out-Null; $line.Children.Add($ft) | Out-Null
    $target.Child = (New-GuiStack @($eyebrow, $big, $line))
    $volume = if ($SizeUnknown) { "$ItemsQuantity" } else { "$ItemsQuantity $mid $(Format-MbtGuiBytes $Bytes)" }
    $vol = New-GuiText $volume 'Big'; $vol.Margin = [Windows.Thickness]::new(0, 12, 0, 0)
    $sub = New-GuiText 'will be added to this mailbox.' 'Hint'; $sub.Margin = [Windows.Thickness]::new(0)
    $warn = New-GuiMessage -Kind warn -Icon shield -Text '**Adds only.** Nothing is deleted or overwritten. This tool cannot undo an import: you would delete the items yourself.'
    $claim = "I checked: add $ItemsQuantity to $mailbox $mid $where $mid $folderText"
    $box = New-GuiOption -Type Check -Title $claim -Checked $s.Confirmed -Margin '0,12,0,0' -OnChange {
        param($x, $e)
        $s.Confirmed = [bool]$x.IsChecked
        $s.ConfirmKey = $(if ($s.Confirmed) { Get-MbtGuiPlanKey } else { '' })
        Update-MbtGuiPrimary -Enabled ($s.Confirmed -and $s.PlanState -eq 'ok') -Footer (Get-MbtGuiPlanFooter)
    }.GetNewClosure()
    $box.FontWeight = [Windows.FontWeights]::SemiBold
    $frame = [Windows.Controls.Border]::new(); $frame.BorderThickness = [Windows.Thickness]::new(1); $frame.CornerRadius = [Windows.CornerRadius]::new(8); $frame.Padding = [Windows.Thickness]::new(12, 0, 12, 12); $frame.Margin = [Windows.Thickness]::new(0, 14, 0, 0)
    Set-MbtGuiResource $frame ([Windows.Controls.Border]::BorderBrushProperty) 'AccentFillColorDefaultBrush'
    $frame.Child = $box
    return (New-GuiCard -Title 'Before anything is written' -Content @($target, $vol, $sub, $warn, $frame))
}

function Build-MbtGuiImportPlanPage {
    $g = $script:Gui; $s = $g.S
    $info = $s.ImInfo
    $browser = Get-MbtGuiBrowser 'import'
    $sel = Get-MbtGuiTaskSelection
    $items = if ($sel.Items -gt 0) { $sel.Items } else { [long]$info.Items }
    $bytes = if ($sel.Items -gt 0) { $sel.Bytes } else { [long]$info.Bytes }
    $sizeUnknown = ($sel.Items -gt 0 -and $sel.UnknownSize -gt 0 -and $bytes -le 0)
    $itemsQuantity = if ($sel.Items -gt 0) { Format-MbtGuiSelectionItems $sel -WithNoun } else { Format-MbtGuiQuantity $items 'item' }
    $g.ImportTotals = @{ Items = $items; Bytes = $bytes }
    $target = if ($s.TgtMode -eq 'new') { $s.TgtName } else { 'the same folder names' }
    $steps = @('Check the PST files again.', "Create the folders under **$target**.", "Send $itemsQuantity, one by one, into the mailbox.", 'For any item Microsoft does not answer for, look in the folder and decide by itself.', 'Final check of the whole job.')
    if ($s.ImSrc -eq 'store') { $steps[0] = 'Check the export store again.' }
    elseif ($s.ImSrc -eq 'eml') { $steps[0] = 'Read the .eml files.' }
    $ol = [Windows.Controls.StackPanel]::new(); $i = 1
    foreach ($t in $steps) { $r = New-GuiText ("$i.  $t") 'Body'; $r.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $ol.Children.Add($r) | Out-Null; $i++ }
    $eta = Get-MbtGuiPlanEta -Task import -Bytes $bytes
    $sizeText = if ($sizeUnknown) { 'size unknown before the check' } elseif ($sel.Items -gt 0) { Format-MbtGuiSelectionSize $sel } else { Format-MbtGuiBytes $bytes }
    $srcText = "$([IO.Path]::GetFileName($s.ImPath)) ($(if ($sel.Items -gt 0) { "$itemsQuantity chosen, " })$sizeText$(if ($info -and $info.PSObject.Properties['Parts'] -and $info.Parts) { ", $(Format-MbtGuiQuantity $info.Parts 'part')" }))"
    $etaText = if ($sizeUnknown) { 'depends on the size of the folders chosen; the check below counts it' } else { "**$($eta.Text)**" }
    $facts = New-GuiFacts @(@('From', $srcText), @('Into', "$(if ($s.ImShared) { $s.ImShared } else { $s.Mailbox }) $([char]0x00B7) $($s.TgtLoc) mailbox"), @('Time', (New-GuiText $etaText 'Body')))
    $warn = New-GuiMessage -Kind warn -Icon clock -Text '**This is a long job.** Microsoft accepts only about 4 GiB per hour into one mailbox, whatever the speed of your PC or network. You can stop and continue as often as you like; items already done are skipped.'
    $g.PlanHolder = [Windows.Controls.ContentControl]::new(); $g.PlanHolder.Content = New-MbtGuiPlanStateMessage
    $sourceRow = if ($info -and $info.Ok) { New-MbtGuiCheckRow -State ok -Text $(if ($s.ImSrc -eq 'pst') { 'The PST file is clean' } else { 'The source is clean' }) -Detail $(if ($info.PSObject.Properties['Parts'] -and $info.Parts) { "$(Format-MbtGuiQuantity $info.Parts 'part'), $(Format-MbtGuiBytes $info.Bytes)" } else { Format-MbtGuiBytes $info.Bytes }) } else { New-MbtGuiCheckRow -State bad -Text 'The source is not checked' -Detail 'go back to From' }
    $roomRow = $null
    $loc = @($s.Locations | Where-Object { $_.Kind -eq $s.TgtLoc -and $_.PSObject.Properties['Quota'] -and $_.Quota })
    if ($loc.Count -and -not $s.ImShared) {
        $fits = ($bytes -le ([double]$loc[0].Quota - [double]$loc[0].Used))
        $roomRow = New-MbtGuiCheckRow -State $(if ($fits) { 'ok' } else { 'bad' }) -Text $(if ($fits) { 'The mailbox has room' } else { 'The mailbox is probably too small' }) -Detail "$($s.TgtLoc): $(Format-MbtGuiBytes $loc[0].Used) of $(Format-MbtGuiBytes $loc[0].Quota) used"
    }
    $checks = New-GuiCard -Title 'Checks' -Content @($g.PlanHolder, (New-MbtGuiSignInCheckRow), $sourceRow, $roomRow)
    $left = New-GuiStack @($checks, (New-GuiCard -Title 'What will be done' -Content @($ol, $facts, $warn)))
    $before = New-GuiCard -Title 'Before you start' -Content @((New-GuiText ([char]0x2022 + '  Keep the PC plugged in and awake (the console does this for you).') 'Body'), (New-GuiText ([char]0x2022 + '  The source files must stay where they are until the end.') 'Body'), (New-MbtGuiSleepInfo))
    $right = New-GuiStack @((New-MbtGuiImportConfirmCard -Items $items -Bytes $bytes -ItemsQuantity $itemsQuantity -SizeUnknown $sizeUnknown), $before)
    $key = Get-MbtGuiPlanKey
    if ($s.PlanKey -ne $key) { $s.PlanState = 'idle'; $s.PlanKey = $key; Start-MbtGuiPlanCheck }
    if ($s.ConfirmKey -ne $key) { $s.Confirmed = $false; $s.ConfirmKey = '' }
    $mailbox = if ($s.ImShared) { $s.ImShared } else { $s.Mailbox }
    Set-MbtGuiActions -Footer (Get-MbtGuiPlanFooter) -Label "Start the import into $mailbox" -Icon play -Enabled ($s.Confirmed -and $s.PlanState -eq 'ok') -OnClick { Start-MbtGuiRun }
    return (New-GuiStack @((New-MbtGuiVerifyBottom -Columns (New-GuiTwoColumns -Left $left -Right $right -LeftWidth 520) -Task import -Include @('conc', 'retry', 'reconcile', 'tries', 'awake'))))
}

#endregion

#region Check and repair a PST ---------------------------------------------------------------------------------------

function Start-MbtGuiPstCheck {
    <# As soon as a file is chosen: the check runs, read only; its result shows on this page. #>
    $g = $script:Gui; $s = $g.S
    if ([string]::IsNullOrWhiteSpace($s.RpPath)) { return }
    $s.RpChecking = $true; $s.RpCheck = $null
    if (-not $s.RpDest) {
        $dir = Split-Path $s.RpPath -Parent; $stem = [IO.Path]::GetFileNameWithoutExtension($s.RpPath)
        $s.RpDest = [IO.Path]::Combine($dir, "$stem.repaired.pst")
    }
    Show-MbtGuiPage
    & $g.Engine.CheckPst $s.RpPath {
        param($r) $s.RpCheck = $r; $s.RpChecking = $false; Show-MbtGuiPage
    }.GetNewClosure() {
        param($msg) $m = ConvertTo-MbtGuiMessage $msg; $s.RpCheck = [pscustomobject]@{ Failed = $true; Message = $m.Text; Tech = $m.Tech }; $s.RpChecking = $false; Show-MbtGuiPage
    }.GetNewClosure()
}

function Build-MbtGuiRepairFromPage {
    $g = $script:Gui; $s = $g.S
    $chk = $s.RpCheck
    $fileRow = New-MbtGuiBrowseFolderRow -Value $s.RpPath -Kind File -Title 'PST file' -Filter 'PST files (*.pst)|*.pst|All files (*.*)|*.*' -Automation 'PST file' -OnChange { param($v) if ($v -and $v -ne $s.RpPath) { $s.RpPath = $v; $s.RpDest = ''; Start-MbtGuiPstCheck } }
    $content = @((New-GuiText 'PST file' 'Label'), $fileRow)
    if ($chk -and -not $chk.PSObject.Properties['Failed']) {
        $content += New-GuiFacts @(@('Size', (Format-MbtGuiBytes $chk.SizeBytes)), @('Format', 'Unicode PST'), @('Checked', (New-GuiText "just now, in $($chk.Elapsed) (read only)" 'Body')))
    }
    $content += New-GuiText 'The check starts as soon as you pick a file. It never changes the file.' 'Hint'
    if ($s.RpPath) { $again = New-GuiButton -Text 'Check again' -Small -Icon refresh -Disabled:$s.RpChecking -OnClick { Start-MbtGuiPstCheck }; $again.Margin = [Windows.Thickness]::new(0, 8, 0, 0); $again.HorizontalAlignment = 'Left'; $content += $again }
    $left = New-GuiCard -Title 'File' -Content $content
    $bad = $false
    $resBody = [Collections.Generic.List[object]]::new()
    if ($s.RpChecking) {
        $resBody.Add((New-GuiMessage -Kind info -Icon clock -Text 'Checking the file. This only reads. A big file takes a few minutes.'))
        $info = ''
    } elseif ($chk -and $chk.PSObject.Properties['Failed']) {
        $resBody.Add((New-GuiMessage -Kind bad -Icon x -Text $chk.Message -After @((New-GuiMore -Summary 'Technical message' -Content @((New-GuiText $chk.Tech 'Small'))))))
        $info = ''
    } elseif ($chk) {
        $bad = ($chk.Errors -gt 0 -or $chk.Warnings -gt 0)
        if ($bad) {
            $pills = New-GuiStack @($(if ($chk.Errors) { $p = New-GuiPill (Format-MbtGuiQuantity $chk.Errors 'error') 'bad'; $p.Margin = [Windows.Thickness]::new(0, 0, 8, 0); $p }), $(if ($chk.Warnings) { New-GuiPill (Format-MbtGuiQuantity $chk.Warnings 'warning') 'warn' })) -Orientation Horizontal
            $resBody.Add($pills)
            $resBody.Add((New-GuiMessage -Kind warn -Icon warn -Text $(if ($chk.Errors) { '**The file has damage, but it can probably be repaired.** Outlook may refuse to open it, or show folders that are empty or missing items.' } else { '**The file has small problems.** Outlook can probably open it. A repaired copy puts the counters right.' })))
            $rows = @(foreach ($f in @($chk.Findings)) {
                $label = New-GuiStack @((New-GuiStack @((New-GuiPill $(if ($f.Severity -eq 'Error') { 'Error' } else { 'Warning' }) $(if ($f.Severity -eq 'Error') { 'bad' } else { 'warn' }) -Light), (New-GuiText $f.Text 'Body')) -Orientation Horizontal), (New-GuiText $f.Code 'Small'))
                $label.Children[0].Children[1].Margin = [Windows.Thickness]::new(8, 0, 0, 0)
                , @($label, [string]$f.Count, $f.Level)
            })
            $resBody.Add((New-GuiTable -Headers @('What was found', 'Count', 'Repaired by level') -Rows @($rows) -Widths @('*', '70', '130') -Right @(1)))
            $resBody.Add((New-GuiText "Recommended: **$($chk.Recommended)**. It also does everything Recompute does." 'Hint'))
        } else {
            $resBody.Add((New-GuiPill 'Clean' 'ok'))
            $resBody.Add((New-GuiMessage -Kind ok -Icon check -Text '**No error and no warning.** The file is healthy. There is nothing to repair.'))
        }
        $info = "Exit code $($chk.ExitCode) ($(if ($chk.ExitCode -eq 0) { 'clean' } elseif ($chk.ExitCode -eq 2) { 'warnings' } else { 'errors' }))"
    } else {
        $resBody.Add((New-GuiText 'Pick a PST file to check it.' 'Muted')); $info = ''
    }
    $right = New-GuiCard -Title 'Result of the check' -Info $info -Content $resBody.ToArray()
    $rightStack = @($right)
    if ($bad) { $rightStack += New-GuiCard -Title 'What to do next' -Content @((New-GuiText 'The file is damaged but can probably be repaired. A repaired copy is a new file; the original is never changed.' 'Body'), (New-GuiButton -Text 'Repair into a new file' -Icon wrench -Accent -OnClick { Move-MbtGuiStep 1 }.GetNewClosure() | ForEach-Object { $_.Margin = [Windows.Thickness]::new(0, 10, 0, 0); $_.HorizontalAlignment = 'Left'; $_ })) }
    if ($bad) { Set-MbtGuiActions -Footer 'The file has problems. You can write a repaired copy.' -Label 'Repair into a new file' -Icon wrench -OnClick { Move-MbtGuiStep 1 } }
    elseif ($chk -and -not $chk.PSObject.Properties['Failed']) { Set-MbtGuiActions -Footer 'The file is clean. Nothing to repair.' -Label 'Start a new task' -OnClick { $s.Task = $null; $s.Sel = $null; $s.Step = 1; Show-MbtGuiPage } }
    else { Set-MbtGuiActions -Footer 'Pick the PST file.' -Label 'Repair into a new file' -Icon wrench -Enabled $false -OnClick { } }
    return (New-GuiStack @((New-GuiText 'Which PST file?' 'H2'), (New-GuiTwoColumns -Left (New-GuiStack @($left)) -Right (New-GuiStack $rightStack))))
}

function Test-MbtGuiRepairDestination {
    param([hashtable]$S)
    if ([string]::IsNullOrWhiteSpace($S.RpDest)) { return 'Choose a name for the new file.' }
    $bad = Test-MbtGuiPstName ([IO.Path]::GetFileName($S.RpDest))
    if ($bad) { return $bad }
    if ($S.RpDest -ieq $S.RpPath) { return 'This is the original file. The repair never touches it. Choose another name.' }
    if (Test-Path -LiteralPath $S.RpDest) { return 'This file already exists. The tool never overwrites a file. Choose another name.' }
    return $null
}

function Build-MbtGuiRepairToPage {
    $g = $script:Gui; $s = $g.S
    $size = if ($s.RpCheck -and $s.RpCheck.PSObject.Properties['SizeBytes']) { [double]$s.RpCheck.SizeBytes } else { 0 }
    $free = $null; try { $free = [double](& $g.Engine.FreeSpace (Split-Path $s.RpDest -Parent)) } catch { }
    $holder = [Windows.Controls.ContentControl]::new(); $g.RepairHolder = $holder
    $update = {
        $p = Test-MbtGuiRepairDestination $s
        $g.RepairHolder.Content = if ($p) { New-GuiMessage -Kind bad -Icon x -Text $p } else { $null }
        $g.C.Primary.IsEnabled = -not $p
    }.GetNewClosure()
    $row = New-MbtGuiBrowseFolderRow -Value $s.RpDest -Kind SaveFile -Title 'New PST file' -Filter 'PST files (*.pst)|*.pst' -Automation 'New file' -OnChange { param($v) $s.RpDest = $v; & $update }.GetNewClosure()
    $hint = "A new file. The tool never overwrites a file and never touches the original. You need about as much free disk space as the PST ($(Format-MbtGuiBytes $size)$(if ($null -ne $free) { "; $(Format-MbtGuiBytes $free) free" }))."
    $newFile = New-GuiCard -Title 'New file' -Content @((New-GuiText 'Write the repaired copy as' 'Label'), $row, $holder, (New-GuiText $hint 'Hint'))
    $rec = [string]$(if ($s.RpCheck -and $s.RpCheck.PSObject.Properties['Recommended']) { $s.RpCheck.Recommended } else { 'Rebuild' })
    $levelRows = @(
        @('Automatic', 'Automatic', "The tool picks the lightest level that fixes what it found (here: $rec) and goes up by itself if needed.", $true),
        @('Recompute', 'Recompute only', 'The structure is trusted. Fixes counters, sizes and indexes.', $false),
        @('Rebuild', 'Rebuild', 'Finds folders and messages again from the file''s nodes. Needed when messages are not listed.', $false),
        @('Salvage', 'Salvage', 'Scans the whole file for readable blocks. For files that do not open at all, or are cut short. Slowest.', $false))
    $levels = foreach ($row in $levelRows) {
        $key = $row[0]
        New-GuiOption -Type Radio -Group lvl -Title $(if ($row[3]) { "$($row[1]) (recommended)" } else { $row[1] }) -Hint $row[2] -Checked ($s.RpLevel -eq $key) -OnChange { $s.RpLevel = $key }.GetNewClosure()
    }
    $keep = New-GuiMore -Summary 'More options' -Content @((New-GuiOption -Type Check -Title 'Keep damaged attachments' -Hint 'An attachment with unreadable blocks is kept at its full size, with zeros in place of what could not be read, and renamed **name (damaged).ext**. Off by default: such a file is damaged, so it says so.' -Checked $s.RpKeep -OnChange { param($x, $e) $s.RpKeep = [bool]$x.IsChecked }))
    $levelCard = New-GuiCard -Title 'Repair level' -Content (@($levels) + @($keep))
    $what = New-GuiCard -Title 'What a repair does' -Content @((@('Reads the damaged file as far as it can be read.', 'Writes a **new, clean PST** with everything it found.', 'Checks both files and shows a before and after.', 'Messages it could not read are listed with the reason. Nothing is guessed.') | ForEach-Object { $t = New-GuiText ([char]0x2022 + '  ' + $_) 'Body'; $t.Margin = [Windows.Thickness]::new(0, 0, 0, 6); $t }))
    Set-MbtGuiActions -Footer 'The original file is never changed.' -Label 'Continue to Verify' -Enabled (-not (Test-MbtGuiRepairDestination $s)) -OnClick { Move-MbtGuiStep 1 }
    & $update
    return (New-GuiStack @((New-GuiText 'How to repair it?' 'H2'), (New-GuiTwoColumns -Left (New-GuiStack @($newFile, $levelCard)) -Right (New-GuiStack @($what)))))
}

function Build-MbtGuiRepairPlanPage {
    $g = $script:Gui; $s = $g.S
    $size = if ($s.RpCheck -and $s.RpCheck.PSObject.Properties['SizeBytes']) { [double]$s.RpCheck.SizeBytes } else { 0 }
    $rec = [string]$(if ($s.RpCheck -and $s.RpCheck.PSObject.Properties['Recommended']) { $s.RpCheck.Recommended } else { 'Rebuild' })
    $lvl = if ($s.RpLevel -eq 'Automatic') { "Automatic ($rec expected)" } else { $s.RpLevel }
    $steps = @('Check the original once more (read only).', "Read what can be read at level **$lvl**.", "Write **$([IO.Path]::GetFileName($s.RpDest))**.", 'Check the new file and compare.')
    $ol = [Windows.Controls.StackPanel]::new(); $i = 1
    foreach ($t in $steps) { $r = New-GuiText ("$i.  $t") 'Body'; $r.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $ol.Children.Add($r) | Out-Null; $i++ }
    $mins = [math]::Max(1, [math]::Round($size / 1GB * 1.2))
    $facts = New-GuiFacts @(@('Original', (New-GuiText "$([IO.Path]::GetFileName($s.RpPath)) ($(Format-MbtGuiBytes $size)) $([char]0x00B7) **not changed**" 'Body')), @('New file', $s.RpDest), @('Time', (New-GuiText "measured once it runs (it depends on your disk)" 'Body')), @('Keep damaged attachments', $(if ($s.RpKeep) { 'yes' } else { 'no' })))
    $left = New-GuiStack @((New-GuiCard -Title 'What will be done' -Content @($ol, $facts)))
    $safe = New-GuiCard -Title 'Safe by design' -Content @((New-GuiText 'A cancelled or failed repair leaves no half-written file: the new file appears only when it is complete.' 'Body'))
    Set-MbtGuiActions -Footer 'Nothing starts before you press Start.' -Label 'Start the repair' -Icon play -OnClick { Start-MbtGuiRun }
    return (New-GuiStack @((New-GuiText 'Check the plan' 'H2'), (New-MbtGuiVerifyBottom -Columns (New-GuiTwoColumns -Left $left -Right (New-GuiStack @($safe)) -LeftWidth 520) -Task repair -Include @('awake'))))
}

#endregion

#region Open ----------------------------------------------------------------------------------------------------------

function Build-MbtGuiOpenPage {
    $g = $script:Gui; $s = $g.S
    $browser = Get-MbtGuiBrowser 'open'
    $info = $s.OpInfo
    $left = New-GuiStack @((New-MbtGuiSourceCard -Use open))
    if ($info -and $info.Ok) {
        $facts = New-GuiFacts @(@('Size', (Format-MbtGuiBytes $info.Bytes)), @('Items', "$(Format-MbtGuiCount $info.Items) in $(Format-MbtGuiQuantity $info.Folders 'folder')"))
        $left.Children[0].Child.Children.Add($facts) | Out-Null
    }
    $sel = Get-MbtGuiSelection -Roots @($browser.Roots)
    $left.Children.Add((New-MbtGuiSelectionCard $sel)) | Out-Null
    $right = [Windows.Controls.Grid]::new()
    if ($info -and $info.Ok) { Place-MbtGuiBrowser $browser $right }
    else {
        $ico = New-GuiIcon open 22 'MbtBrandText'; $ico.HorizontalAlignment = 'Center'
        $t1 = New-GuiText 'Pick a PST file or an export folder to look inside it.' 'Muted'; $t1.TextAlignment = 'Center'; $t1.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
        $body = New-GuiStack @($ico, $t1); $body.Margin = [Windows.Thickness]::new(10, 60, 10, 60)
        $right = New-GuiCard -Title 'Browse' -Content @($body)
    }
    $canImport = [bool]($info -and $info.Ok -and ($s.OpSrc -eq 'store' -or (Get-MbtGuiCapability 'PstImport')))
    # the ticks of an open PST can be written into new PST files (PST to PST, our own writer)
    $copyOn = Get-MbtGuiCapability 'PstCopy'
    $copyReady = { $x = Get-MbtGuiSelection -Roots @((Get-MbtGuiBrowser 'open').Roots); [bool]($script:Gui.S.OpInfo -and $script:Gui.S.OpInfo.Ok -and $script:Gui.S.OpSrc -eq 'pst' -and $x.Items -gt 0) }
    $export = New-GuiButton -Text $(if ($copyOn) { 'Export my ticks to a PST...' } else { 'Export my ticks to a PST... (coming in this version)' }) -Icon down -Disabled:(-not ($copyOn -and (& $copyReady))) -OnClick {
        $s.Task = 'copy'; $s.Sel = 'copy'; $s.Step = $script:Flows.copy.IndexOf('cpTo')
        Initialize-MbtGuiCopy
        Show-MbtGuiPage
    }
    $g.OpenCopyButton = $export
    $g.OnSelectionChanged = { if ($g.OpenCopyButton) { $g.OpenCopyButton.IsEnabled = [bool]((Get-MbtGuiCapability 'PstCopy') -and (& $copyReady)) } }.GetNewClosure()
    $footerText = if ($copyOn -and $s.OpSrc -eq 'pst') { 'Read only. Tick folders or messages to write them into a new PST.' } else { 'Read only. Nothing is written.' }
    Set-MbtGuiActions -Footer $footerText -Extra $export -Label 'Import into a mailbox...' -Enabled $canImport -OnClick {
        $s.ImSrc = $s.OpSrc; $s.ImPath = $s.OpPath; $s.ImInfo = $s.OpInfo
        $s.Task = 'import'; $s.Sel = 'import'; $s.Step = $script:Flows.import.IndexOf('imFrom')
        Show-MbtGuiPage
    }
    return (New-GuiFillPage -Columns (New-GuiTwoColumns -Left $left -Right $right -Fill))
}

#endregion

#region Copy part of a PST (the ticks of an open PST, written into new PST files) ------------------------------------------

function Initialize-MbtGuiCopy {
    <# The task starts from the Open page: the new files go next to the original by default, under another name. #>
    $s = $script:Gui.S
    $s.Outcome = $null; $s.PlanState = 'idle'; $s.PlanMsg = ''; $s.PlanTotals = $null; $s.RStep = $null; $s.RObj = $null
    $script:Gui.Run = $null; $script:Gui.Job = $null
    if ($s.CpFor -ne $s.OpPath) {
        $s.CpFor = $s.OpPath
        $s.OutFolder = Split-Path $s.OpPath -Parent
        $s.OutName = [IO.Path]::GetFileNameWithoutExtension($s.OpPath) -replace '-\d{3}$', ''
        $s.OutName += '-copy.pst'
    }
}

function Test-MbtGuiCopyDestination {
    <# The problem in plain words, or $null. #>
    param([hashtable]$S)
    if ([string]::IsNullOrWhiteSpace($S.OutFolder)) { return 'Choose a folder on this computer.' }
    $bad = Test-MbtGuiPstName $S.OutName
    if ($bad) { return $bad }
    $full = [IO.Path]::Combine($S.OutFolder, $S.OutName)
    if ($full -ieq $S.OpPath) { return 'This is the original file. The copy never touches it. Choose another name.' }
    if (Test-Path -LiteralPath $full) { return 'This file already exists. The tool never overwrites a file. Choose another name.' }
    return $null
}

function Build-MbtGuiCopyToPage {
    $g = $script:Gui; $s = $g.S
    $holder = [Windows.Controls.ContentControl]::new()
    $update = {
        $p = Test-MbtGuiCopyDestination $s
        $holder.Content = if ($p) { New-GuiMessage -Kind bad -Icon x -Text $p } else { $null }
        $g.C.Primary.IsEnabled = -not $p
    }.GetNewClosure()
    $folder = New-MbtGuiBrowseFolderRow -Value $s.OutFolder -Kind Folder -Title 'Folder for the new PST' -Automation 'Folder' -OnChange { param($v) $s.OutFolder = $v; & $update }.GetNewClosure()
    $name = New-GuiTextBox -Text $s.OutName -Automation 'File name' -OnChange { param($x, $e) $s.OutName = $x.Text; & $update }.GetNewClosure()
    $stem = [IO.Path]::GetFileNameWithoutExtension($s.OutName)
    $split = [Windows.Controls.ComboBox]::new(); $split.HorizontalAlignment = 'Stretch'
    foreach ($t in '50 GB (Outlook''s default limit)', '20 GB', '10 GB', '5 GB') { [void]$split.Items.Add($t) }
    $split.SelectedIndex = switch ([int]$s.SplitGB) { 50 { 0 } 20 { 1 } 10 { 2 } 5 { 3 } default { 0 } }
    Add-MbtGuiHandler $split 'SelectionChanged' { param($sender, $e) $s.SplitGB = @(50, 20, 10, 5)[$sender.SelectedIndex] }
    $more = New-GuiMore -Summary 'More options' -Content @(
        (New-GuiText 'Largest PST file' 'Label'), $split,
        (New-GuiText 'A message is never cut in two. The files are named after the first one, then -002, -003 and so on.' 'Hint'),
        (New-GuiOption -Type Check -Title 'Check every PST file when it is written' -Hint 'Recommended.' -Checked $s.CheckEach -OnChange { param($x, $e) $s.CheckEach = [bool]$x.IsChecked }))
    $content = @((New-GuiText 'Folder on this computer' 'Label'), $folder, (New-GuiText 'File name' 'Label'), $name,
        (New-GuiText "Parts are named $stem.pst, $stem-002.pst, ... Names such as AUX, CON or NUL are refused (Windows reserves them)." 'Hint'), $more, $holder)
    $left = New-GuiCard -Title 'Save as' -Content $content
    $sel = Get-MbtGuiSelection -Roots @((Get-MbtGuiBrowser 'open').Roots)
    $right = New-GuiStack @((New-MbtGuiSelectionCard $sel), (New-GuiCard -Title 'Good to know' -Content @((New-GuiText 'The original PST is only read. The new files hold the folders and messages you ticked, with the same folder names. A copy that is stopped leaves no file: start it again.' 'Body'))))
    Set-MbtGuiActions -Footer 'Choose where to save the new PST. You can still change everything on the next page.' -Label 'Continue to Verify' -Enabled (-not (Test-MbtGuiCopyDestination $s)) -OnClick { Move-MbtGuiStep 1 }
    & $update
    return (New-GuiStack @((New-GuiText 'Where to save the new PST?' 'H2'), (New-GuiTwoColumns -Left (New-GuiStack @($left)) -Right $right)))
}

function Build-MbtGuiCopyPlanPage {
    $g = $script:Gui; $s = $g.S
    $sel = Get-MbtGuiTaskSelection
    $steps = @('Read the folders and messages you ticked in the PST (it is only read).', "Write the new PST files ($($s.SplitGB) GB at most each) with the same folder tree.", $(if ($s.CheckEach) { 'Check every new PST file.' } else { 'Skip the check of the new files (you turned it off).' }))
    $ol = [Windows.Controls.StackPanel]::new(); $i = 1
    foreach ($t in $steps) { $r = New-GuiText ("$i.  $t") 'Body'; $r.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $ol.Children.Add($r) | Out-Null; $i++ }
    $facts = New-GuiFacts @(@('From', "$([IO.Path]::GetFileName($s.OpPath)) $([char]0x00B7) **not changed**"), @('Saved to', [IO.Path]::Combine($s.OutFolder, $s.OutName)),
        @('Ticked', "$(Format-MbtGuiSelectionItems $sel -WithNoun) $([char]0x00B7) $(Format-MbtGuiSelectionSize $sel)"), @('Time', (New-GuiText 'measured once it runs (it depends on your disk)' 'Body')))
    $g.PlanHolder = [Windows.Controls.ContentControl]::new(); $g.PlanHolder.Content = New-MbtGuiPlanStateMessage
    $left = New-GuiStack @((New-GuiCard -Title 'Checks' -Content @($g.PlanHolder)), (New-GuiCard -Title 'What will be done' -Content @($ol, $facts)))
    $right = New-GuiStack @((New-GuiCard -Title 'What is in it' -Content @((New-MbtGuiPlanTable $sel))), (New-GuiCard -Title 'Safe by design' -Content @((New-GuiText 'The original is only read. The new file appears only when it is complete: a copy that is stopped or fails leaves no file behind.' 'Body'))))
    $key = Get-MbtGuiPlanKey
    if ($s.PlanKey -ne $key) { $s.PlanState = 'idle'; $s.PlanKey = $key; Start-MbtGuiPlanCheck }
    Set-MbtGuiActions -Footer (Get-MbtGuiPlanFooter) -Label 'Start the copy' -Icon play -Enabled ($s.PlanState -eq 'ok') -OnClick { Start-MbtGuiRun }
    return (New-GuiStack @((New-GuiText 'Check the plan' 'H2'), (New-MbtGuiVerifyBottom -Columns (New-GuiTwoColumns -Left $left -Right $right -LeftWidth 520) -Task copy -Include @('awake'))))
}

#endregion

#region Copy to another mailbox (a mailbox read into a store, then added to another mailbox) -----------------------------------

function Get-MbtGuiTargetArchive {
    <# What is known of the archive of the mailbox that receives the copy: State Known, None (the mailbox has none), Unknown (the lookup could not tell), Reading, or Unasked (no address yet). Id: the one typed or found. #>
    param([string]$Address, [bool]$App)
    $g = $script:Gui; $s = $g.S
    $key = ([string]$Address).Trim().ToLowerInvariant()
    if ($App) {
        $typed = if ($s.McArchive -is [Collections.IDictionary] -and $key -and $s.McArchive.Contains($key)) { [string]$s.McArchive[$key] } else { '' }
        if ($typed) { return [pscustomobject]@{ State = 'Known'; Id = $typed; Typed = $true; Detail = '' } }
        $st = Get-MbtGuiProp $g 'McTargetArchive' $null
        if ($st -and $key -and [string](Get-MbtGuiProp $st 'Key' '') -eq $key) { return [pscustomobject]@{ State = [string]$st.State; Id = [string](Get-MbtGuiProp $st 'Id' ''); Typed = $false; Detail = [string](Get-MbtGuiProp $st 'Detail' '') } }
        return [pscustomobject]@{ State = 'Unasked'; Id = ''; Typed = $false; Detail = '' }
    }
    # a delegated sign-in writes into its own mailbox: what Where to look found of it
    $has = @($s.Locations | Where-Object { $_.Kind -eq 'Archive' }).Count -gt 0
    $state = [string](Get-MbtGuiProp $s.ArchiveState 'State' '')
    return [pscustomobject]@{ State = $(if ($has) { 'Known' } elseif ($state -in 'None', 'Unknown') { $state } else { 'Known' }); Id = ''; Typed = $false; Detail = '' }
}

function Start-MbtGuiTargetArchiveLookup {
    <# The application is going to write into a mailbox that is typed: look up whether it has an archive (Microsoft Graph, User.Read.All), once for each address. #>
    $g = $script:Gui; $s = $g.S
    if ($s.AuthMode -ne 'app') { return }
    $address = ([string]$s.McTarget).Trim()
    if ($address -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return }
    $key = $address.ToLowerInvariant()
    $current = Get-MbtGuiProp $g 'McTargetArchive' $null
    if ($current -and [string](Get-MbtGuiProp $current 'Key' '') -eq $key) { return }
    $g.McTargetArchive = @{ Key = $key; State = 'Reading'; Id = ''; Detail = '' }
    $typed = if ($s.McArchive.Contains($key)) { [string]$s.McArchive[$key] } else { '' }
    $refresh = { if ((Get-MbtGuiScreen) -eq 'mcTo') { Show-MbtGuiPage } }
    & $g.Engine.Locations $address '' {
        param($locs, $archive)
        $found = @($locs | Where-Object { $_.Kind -eq 'Archive' })
        $state = if ($found.Count) { 'Known' } elseif ($archive -and [string]$archive.State -eq 'None') { 'None' } else { 'Unknown' }
        $g.McTargetArchive = @{ Key = $key; State = $state; Id = $(if ($found.Count) { [string]$found[0].MailboxId } else { '' }); Detail = [string](Get-MbtGuiProp $archive 'Detail' '') }
        & $refresh
    }.GetNewClosure() {
        param($msg)
        $g.McTargetArchive = @{ Key = $key; State = 'Unknown'; Id = ''; Detail = [string]$msg }
        & $refresh
    }.GetNewClosure() $null $null $typed $null
}

function Get-MbtGuiMailCopyLayout {
    <# Where the two parts of a copy of both places go: the choice of the page, except that a target without an archive (or whose archive is not known yet) takes both into the Primary. #>
    param($Archive)
    $want = [string]$script:Gui.S.McLayout
    if ($want -notin 'Mirror', 'Primary', 'Archive') { $want = 'Mirror' }
    if ($Archive.State -notin 'Known', 'Unasked', 'Reading' -and $want -ne 'Primary') { return 'Primary' }
    return $want
}

function Get-MbtGuiMailCopyEnds {
    <# The two ends of a copy to another mailbox as the pages say them: the mailbox that is read and its places (from what is ticked, or Everything), the mailbox that receives and where each part goes. Same: a part would write into the place it reads. #>
    $g = $script:Gui; $s = $g.S
    $found = [Collections.Generic.List[string]]::new()
    $sel = if ($g.ContainsKey('ExportSelection')) { $g.ExportSelection } else { $null }
    if ($sel -and $sel.PSObject.Properties['Everything'] -and $sel.Everything) {
        foreach ($p in @($sel.Places)) { if ([string]$p.Kind -in 'Primary', 'Archive' -and -not $found.Contains([string]$p.Kind)) { $found.Add([string]$p.Kind) } }
    }
    elseif ($sel -and @($sel.Entries).Count) {
        foreach ($e in @($sel.Entries)) {
            $k = switch (([string](Get-MbtGuiRootNode $e.Node).Kind).ToLowerInvariant()) { 'archive' { 'Archive' } 'primary' { 'Primary' } default { '' } }
            if ($k -and -not $found.Contains($k)) { $found.Add($k) }
        }
    }
    if (-not $found.Count) { $found.Add($(if ($s.McFromKind -in 'Primary', 'Archive') { [string]$s.McFromKind } else { 'Primary' })) }
    $fromKinds = @('Primary', 'Archive' | Where-Object { $found.Contains($_) })
    $both = ($fromKinds.Count -eq 2)
    $from = if ($s.Shared) { [string]$s.Shared } else { [string]$s.Mailbox }
    $app = ($s.AuthMode -eq 'app')
    $to = if ($app) { ([string]$s.McTarget).Trim() } elseif ($s.Account) { [string]$s.Account } else { [string]$s.Mailbox }
    $archive = Get-MbtGuiTargetArchive -Address $to -App $app
    $layout = if ($both) { Get-MbtGuiMailCopyLayout $archive } else { '' }
    # the parts: each reads one place of the source and writes into one place of the target
    $pairs = [Collections.Generic.List[object]]::new()
    if ($both) {
        $pairs.Add(@('Primary', $(if ($layout -eq 'Archive') { 'Archive' } else { 'Primary' })))
        $pairs.Add(@('Archive', $(if ($layout -eq 'Primary') { 'Primary' } else { 'Archive' })))
    } else { $pairs.Add(@($fromKinds[0], [string]$s.TgtLoc)) }
    $toKinds = @($pairs | ForEach-Object { $_[1] } | Select-Object -Unique)
    $sameMailbox = [bool]($from -and $to -and $from -ieq $to)
    [pscustomobject]@{
        From = $from; FromKinds = $fromKinds; FromKind = ($fromKinds -join ' + '); Both = $both; To = $to; ToKinds = $toKinds; ToKind = ($toKinds -join ' + '); ToPlaces = ($toKinds -join ' / ')
        Layout = $layout; Pairs = $pairs.ToArray(); App = $app; Archive = $archive; Same = [bool]($sameMailbox -and @($pairs | Where-Object { $_[0] -eq $_[1] }).Count)
    }
}

function Get-MbtGuiMailCopyFolderText {
    <# Where the items of a copy are created, in words: the folder of each part. #>
    param($Ends)
    $s = $script:Gui.S
    $one = { param($mode, $name) if ($mode -eq 'new') { "new folder $name" } else { 'folders with the same names' } }
    if ($Ends.Both) { return ('Primary part: {0}; Archive part: {1}' -f (& $one $s.TgtMode $s.TgtName), (& $one $s.TgtMode2 $s.TgtName2)) }
    return (& $one $s.TgtMode $s.TgtName)
}

function Test-MbtGuiMailCopyTarget {
    <# What stops the copy from going on, in plain words, or $null. #>
    param([hashtable]$S)
    $ends = Get-MbtGuiMailCopyEnds
    if ($ends.App) {
        if (-not $ends.To) { return 'Type the address of the mailbox that receives the copy.' }
        if ($ends.To -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { return 'Type the address of the mailbox like adele.vance@contoso.com.' }
        if ($ends.Archive.State -eq 'Reading') { return 'Looking up whether this mailbox has an archive...' }
    }
    if ($ends.Same) { return 'This is the mailbox you copy from. Choose the other part of the mailbox, or another mailbox: a copy writes into a different place.' }
    if (-not $ends.Both -and $S.TgtLoc -eq 'Archive' -and $ends.Archive.State -eq 'None') { return 'This mailbox has no archive. Choose its Primary mailbox.' }
    if (-not $ends.Both -and $S.TgtLoc -eq 'Archive' -and $ends.Archive.State -eq 'Unknown') { return 'The archive of this mailbox could not be looked up. Type its archive mailbox id, or choose the Primary mailbox.' }
    if ($S.TgtMode -eq 'new') { $p = Test-MbtGuiFolderName $S.TgtName; if ($p) { return $p } }
    if ($ends.Both -and $S.TgtMode2 -eq 'new') { $p = Test-MbtGuiFolderName $S.TgtName2; if ($p) { return $p } }
    return $null
}

function Update-MbtGuiMailCopyTo {
    <# The address of the target or the folder name changed: the button and the footer follow, without building the page again. #>
    if ((Get-MbtGuiScreen) -ne 'mcTo') { return }
    $problem = Test-MbtGuiMailCopyTarget $script:Gui.S
    Update-MbtGuiPrimary -Enabled (-not $problem) -Footer $(if ($problem) { $problem } else { 'Choose the mailbox and where in it the items go.' })
}

function New-MbtGuiMailCopyFolderOptions {
    <# The folder choice of one part: a new folder (named), or the same folder names. Which: 1 (the only part, or the Primary part) or 2 (the Archive part). #>
    param([ValidateSet(1, 2)][int]$Which, [string]$Title)
    $s = $script:Gui.S
    $modeKey = if ($Which -eq 1) { 'TgtMode' } else { 'TgtMode2' }; $nameKey = if ($Which -eq 1) { 'TgtName' } else { 'TgtName2' }; $group = if ($Which -eq 1) { 'tm' } else { 'tm2' }
    $out = [Collections.Generic.List[object]]::new()
    $out.Add((New-GuiText $Title 'Label'))
    $out.Add((New-GuiOption -Type Radio -Group $group -Title 'A new folder (recommended)' -Hint 'Everything goes under one folder of this name. If a folder of this name already exists, it is used as it is. Easy to find, easy to move or delete later.' -Checked ($s[$modeKey] -eq 'new') -Margin '0' -OnChange { $s[$modeKey] = 'new'; Show-MbtGuiPage }.GetNewClosure()))
    if ($s[$modeKey] -eq 'new') {
        $name = New-GuiTextBox -Text $s[$nameKey] -Automation $(if ($Which -eq 1) { 'Folder name' } else { 'Archive folder name' }) -OnChange { param($x, $e) $s[$nameKey] = $x.Text; Update-MbtGuiMailCopyTo }.GetNewClosure()
        $name.Margin = [Windows.Thickness]::new(30, 6, 0, 0)
        $out.Add($name)
    }
    $out.Add((New-GuiOption -Type Radio -Group $group -Title 'The same folder names' -Hint 'Items are added to folders with the same names. Existing items are left alone. Duplicates are possible.' -Checked ($s[$modeKey] -eq 'same') -OnChange { $s[$modeKey] = 'same'; Show-MbtGuiPage }.GetNewClosure()))
    return $out.ToArray()
}

function Build-MbtGuiMailCopyToPage {
    $g = $script:Gui; $s = $g.S
    if ($s.AuthMode -eq 'app') { Start-MbtGuiTargetArchiveLookup }
    $ends = Get-MbtGuiMailCopyEnds
    $s.McFromKind = if ($ends.Both) { 'All' } else { [string]$ends.FromKinds[0] }
    $sel = $g.ExportSelection
    $how = New-GuiCard -Title 'How the copy works' -Content @(
        (New-GuiText 'It has two parts, one after the other. First what you chose is read from the mailbox and saved in a store on this PC. Then the store is added to the mailbox that receives the copy.' 'Body'),
        (New-GuiText 'The mailbox you copy from is only read. When the copy is over, the store stays on this PC: it is the proof of what was read, and the point where a stopped copy continues.' 'Body'))
    $left = New-GuiStack @((New-MbtGuiSelectionCard $sel -Period (Get-MbtGuiPeriod $s)), $how)
    $items = [Collections.Generic.List[object]]::new()
    if ($ends.App) {
        $items.Add((New-GuiText 'Mailbox that receives the copy' 'Label'))
        $box = New-GuiTextBox -Text $s.McTarget -Automation 'Target mailbox address' -OnChange { param($x, $e) $s.McTarget = $x.Text; Update-MbtGuiMailCopyTo }
        Add-MbtGuiHandler $box 'LostFocus' { param($sender, $e) if ((Get-MbtGuiScreen) -eq 'mcTo') { $before = $g.McTargetArchive; Start-MbtGuiTargetArchiveLookup; if ($g.McTargetArchive -ne $before) { Show-MbtGuiPage } } }.GetNewClosure()
        $items.Add($box)
        $items.Add((New-GuiText 'Required: the application has no mailbox of its own. The mailbox must be in the scope of the application.' 'Hint'))
        foreach ($i in @(New-MbtGuiTargetArchiveItem $ends)) { $items.Add($i) }
    } else {
        $items.Add((New-GuiMessage -Kind info -Icon shield -Text "**The copy goes into your own mailbox, $($ends.To).** A delegated sign-in can write only into the mailbox of the account that signs in: Microsoft refuses a write into another person's or a shared mailbox. To write into another mailbox, sign in as the application (Setup)."))
    }
    $archiveOk = $ends.Archive.State -in 'Known', 'Unasked', 'Reading'
    if ($ends.Both) {
        $items.Add((New-GuiText 'Where does each part go?' 'Label'))
        $target = if ($ends.To) { $ends.To } else { 'the mailbox' }
        $choices = @(
            @('Mirror', 'Mirror (recommended)', "The Primary part goes into a folder of the Primary mailbox of $target, the Archive part into a folder of its Archive.", $archiveOk),
            @('Primary', 'All into the Primary mailbox', "Both parts go into folders of the Primary mailbox of $target.", $true),
            @('Archive', 'All into the Archive', "Both parts go into folders of the Archive of $target.", $archiveOk))
        foreach ($c in $choices) {
            $layout = $c[0]
            $items.Add((New-GuiOption -Type Radio -Group layout -Title $c[1] -Hint $(if ($c[3]) { $c[2] } else { 'This mailbox has no archive that can be used.' }) -Checked ($ends.Layout -eq $layout) -Disabled:(-not $c[3]) -OnChange { $s.McLayout = $layout; Show-MbtGuiPage }.GetNewClosure()))
        }
        if (-not $archiveOk) { $items.Add((New-GuiMessage -Kind info -Icon info -Text '**The Archive part goes into the Primary mailbox.** This mailbox has no archive that can be used, so the two parts go into folders of its Primary mailbox.')) }
        if (-not $s.TgtName) { $s.TgtName = "Copy of $(($ends.From -split '@')[0]) $((Get-Date).ToString('yyyy-MM-dd'))" }
        if (-not $s.TgtName2) { $s.TgtName2 = "$($s.TgtName) (Archive)" }
        foreach ($i in @(New-MbtGuiMailCopyFolderOptions -Which 1 -Title 'Folder of the Primary part')) { $items.Add($i) }
        foreach ($i in @(New-MbtGuiMailCopyFolderOptions -Which 2 -Title 'Folder of the Archive part')) { $items.Add($i) }
    } else {
        $items.Add((New-GuiText 'Part of the mailbox' 'Label'))
        foreach ($place in @(@('Primary', 'Primary mailbox', ''), @('Archive', 'Archive', $(switch ($ends.Archive.State) { 'None' { 'This mailbox has no archive.' } 'Unknown' { 'The archive could not be looked up: type its archive mailbox id.' } default { '' } })))) {
            $kind = $place[0]
            $disabled = ($kind -eq 'Archive' -and -not $archiveOk)
            $items.Add((New-GuiOption -Type Radio -Group loc -Title $place[1] -Hint $place[2] -Checked ($s.TgtLoc -eq $kind) -Disabled:$disabled -OnChange { $s.TgtLoc = $kind; Show-MbtGuiPage }.GetNewClosure()))
        }
        if (-not $s.TgtName) { $s.TgtName = "Copy of $(($ends.From -split '@')[0]) $((Get-Date).ToString('yyyy-MM-dd'))" }
        foreach ($i in @(New-MbtGuiMailCopyFolderOptions -Which 1 -Title 'Folder')) { $items.Add($i) }
    }
    $items.Add((New-GuiMessage -Kind info -Icon shield -Text '**Nothing is deleted. Nothing is overwritten.** The copy only adds items. A folder is never renamed.'))
    if (Test-MbtGuiSelectionHasDeletedItems $sel) {
        $items.Add((New-GuiMessage -Kind info -Icon shield -Text '**Your selection holds deleted items (Recoverable Items).** They are put in a normal folder named **Restored deleted items** (inside your new folder if you chose one), never back into Recoverable Items.'))
    }
    $mapBtn = New-GuiButton -Text 'Use a mapping file...' -Small -Icon open -OnClick { $f = Select-MbtGuiFile -Title 'Mapping file' -Filter 'Mapping (*.psd1;*.json)|*.psd1;*.json'; if ($f) { $s.Mapping = $f; Show-MbtGuiPage } }
    $mapLine = if ($s.Mapping) { New-GuiText "Mapping file: $($s.Mapping)" 'Small' } else { New-GuiText 'For renaming folders while copying (source path to target path). Optional.' 'Hint' }
    $items.Add((New-GuiMore -Summary 'More options' -Content @($mapBtn, $mapLine)))
    $problem = Test-MbtGuiMailCopyTarget $s
    if ($problem) { $items.Add((New-GuiMessage -Kind warn -Icon warn -Text $problem)) }
    $right = New-GuiCard -Title 'Where to copy' -Info "$($ends.From) ($($ends.FromKind)) $([char]0x2192)" -Content $items.ToArray()
    Set-MbtGuiActions -Footer $(if ($problem) { $problem } else { 'Choose the mailbox and where in it the items go.' }) -Label 'Continue to Verify' -Enabled (-not $problem) -OnClick { Move-MbtGuiStep 1 }
    return (New-GuiStack @((New-GuiText 'Into which mailbox?' 'H2'), (New-GuiTwoColumns -Left $left -Right $right)))
}

function New-MbtGuiMailCopyConfirmCard {
    <# "Before anything is written": the target, the volume, what the copy does, and ONE box that says exactly what, from where and into where. Start stays grey until it is ticked. #>
    param([string]$ItemsQuantity, [long]$Bytes, [bool]$SizeUnknown, [bool]$AtMost, [bool]$Everything)
    $g = $script:Gui; $s = $g.S
    $mid = [char]0x00B7
    $ends = Get-MbtGuiMailCopyEnds
    $folderText = Get-MbtGuiMailCopyFolderText $ends
    $target = [Windows.Controls.Border]::new(); $target.CornerRadius = [Windows.CornerRadius]::new(8); $target.Padding = [Windows.Thickness]::new(14, 10, 14, 12); $target.Margin = [Windows.Thickness]::new(0, 4, 0, 0)
    Set-MbtGuiResource $target ([Windows.Controls.Border]::BackgroundProperty) 'MbtAccentSoft'
    $eyebrow = New-GuiText 'INTO' 'Small' -NoWrap; $eyebrow.FontWeight = [Windows.FontWeights]::SemiBold; Set-MbtGuiResource $eyebrow ([Windows.Controls.TextBlock]::ForegroundProperty) 'MbtBrandText'
    $big = New-GuiText $ends.To 'Big' -NoWrap; $big.FontSize = 22
    $lines = [Collections.Generic.List[object]]::new()
    $parts = if ($ends.Both) { @(@('Primary', $s.TgtMode, $s.TgtName), @('Archive', $s.TgtMode2, $s.TgtName2)) } else { @(, @([string]$s.TgtLoc, $s.TgtMode, $s.TgtName)) }
    $index = 0
    foreach ($part in $parts) {
        $place = if ($ends.Both) { [string]$ends.Pairs[$index][1] } else { [string]$part[0] }
        $line = [Windows.Controls.StackPanel]::new(); $line.Orientation = 'Horizontal'; $line.Margin = [Windows.Thickness]::new(0, 6, 0, 0)
        $pill = New-GuiPill $place 'run'; $pill.Margin = [Windows.Thickness]::new(0, 0, 10, 0)
        $fi = New-GuiIcon folder 14; $fi.Margin = [Windows.Thickness]::new(0, 0, 6, 0)
        $prefix = if ($ends.Both) { "$($part[0]) part " + [char]0x2192 + ' ' } else { '' }
        $ft = New-GuiText $(if ($part[1] -eq 'new') { "$prefix" + "in a new folder named **$($part[2])**" } else { "$prefix" + 'in the folders that have the same names' }) 'Body'; $ft.VerticalAlignment = 'Center'
        $line.Children.Add($pill) | Out-Null; $line.Children.Add($fi) | Out-Null; $line.Children.Add($ft) | Out-Null
        $lines.Add($line); $index++
    }
    $from = New-GuiText "copied from **$($ends.From)** $mid $($ends.FromKind)" 'Body'; $from.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
    $target.Child = (New-GuiStack (@($eyebrow, $big) + $lines.ToArray() + @($from)))
    $volume = if ($SizeUnknown) { "$ItemsQuantity" } else { "$ItemsQuantity $mid $(if ($AtMost) { 'at most ' })$(Format-MbtGuiBytes $Bytes)" }
    $vol = New-GuiText $volume 'Big'; $vol.Margin = [Windows.Thickness]::new(0, 12, 0, 0)
    $sub = New-GuiText 'will be added to this mailbox.' 'Hint'; $sub.Margin = [Windows.Thickness]::new(0)
    $warn = New-GuiMessage -Kind warn -Icon shield -Text '**Adds only.** Nothing is deleted or overwritten. This tool cannot undo a copy: you would delete the items yourself.'
    $what = if ($Everything) { 'everything' } else { $ItemsQuantity }
    $claim = "I checked: copy $what from $($ends.FromKind) of $($ends.From) into $($ends.ToPlaces) of $($ends.To) $mid $folderText"
    $box = New-GuiOption -Type Check -Title $claim -Checked $s.Confirmed -Margin '0,12,0,0' -OnChange {
        param($x, $e)
        $s.Confirmed = [bool]$x.IsChecked
        $s.ConfirmKey = $(if ($s.Confirmed) { Get-MbtGuiPlanKey } else { '' })
        Update-MbtGuiPrimary -Enabled ($s.Confirmed -and $s.PlanState -eq 'ok') -Footer (Get-MbtGuiPlanFooter)
    }.GetNewClosure()
    $box.FontWeight = [Windows.FontWeights]::SemiBold
    $frame = [Windows.Controls.Border]::new(); $frame.BorderThickness = [Windows.Thickness]::new(1); $frame.CornerRadius = [Windows.CornerRadius]::new(8); $frame.Padding = [Windows.Thickness]::new(12, 0, 12, 12); $frame.Margin = [Windows.Thickness]::new(0, 14, 0, 0)
    Set-MbtGuiResource $frame ([Windows.Controls.Border]::BorderBrushProperty) 'AccentFillColorDefaultBrush'
    $frame.Child = $box
    return (New-GuiCard -Title 'Before anything is written' -Content @($target, $vol, $sub, $warn, $frame))
}

function Build-MbtGuiMailCopyPlanPage {
    $g = $script:Gui; $s = $g.S
    $sel = Get-MbtGuiPlanSelection
    $ends = Get-MbtGuiMailCopyEnds
    $period = Get-MbtGuiPeriod $s
    $mid = [char]0x00B7
    $everything = [bool]($sel.PSObject.Properties['Everything'] -and $sel.Everything)
    $bytes = [long]$sel.Bytes
    $sizeUnknown = ($sel.UnknownSize -gt 0 -and $bytes -le 0)
    $itemsQuantity = Format-MbtGuiSelectionItems $sel -WithNoun -AtMost:($period.Active)
    $store = (New-MbtGuiRequest -Task mailcopy).StorePath
    $target = if ($s.TgtMode -eq 'new') { $s.TgtName } else { 'the same folder names' }
    $targetArchive = if ($s.TgtMode2 -eq 'new') { $s.TgtName2 } else { 'the same folder names' }
    $readText = if ($everything) { "Read the folder tree of $($ends.FromKind) and list everything in it." } else { 'Read the folder tree and list the items you ticked.' }
    $steps = [Collections.Generic.List[string]]::new()
    $steps.Add($(if ($period.Active) { "Read the folder tree and list $(if ($everything) { 'everything' } else { 'the items you ticked' }), keeping the items $($period.Text)." } else { $readText }))
    $steps.Add('Download them into the store on this PC, checking each one.')
    $steps.Add('Open the store again and check it. Here the second part begins.')
    if ($ends.Both) {
        $steps.Add("Primary part: create the folders under **$target** in the $($ends.Pairs[0][1]) mailbox of $($ends.To), then send the items, one by one.")
        $steps.Add("Archive part: create the folders under **$targetArchive** in the $($ends.Pairs[1][1]) mailbox of $($ends.To), then send the items, one by one.")
    } else {
        $steps.Add("Create the folders under **$target** in the mailbox that receives the copy.")
        $steps.Add("Send $itemsQuantity, one by one, into that mailbox.")
    }
    $steps.Add('For any item Microsoft does not answer for, look in the folder and decide by itself. Then a final check of the whole job.')
    $ol = [Windows.Controls.StackPanel]::new(); $i = 1
    foreach ($t in $steps) { $r = New-GuiText ("$i.  $t") 'Body'; $r.Margin = [Windows.Thickness]::new(0, 0, 0, 4); $ol.Children.Add($r) | Out-Null; $i++ }
    $eta = Get-MbtGuiPlanEta -Task import -Bytes $bytes
    $etaText = if ($sizeUnknown) { 'cannot be estimated yet: Microsoft does not give the size of every folder in advance. It shows once the copy runs.' } else { "**$($eta.Text)**, mostly the second part: Microsoft accepts only about 4 GiB per hour into one mailbox" }
    $sizeText = if ($sizeUnknown) { 'size unknown before the check' } else { "$(if ($period.Active) { 'at most ' })$(Format-MbtGuiSelectionSize $sel)" }
    $intoText = if ($ends.Both) { "$($ends.To) $mid Primary part into the $($ends.Pairs[0][1]), Archive part into the $($ends.Pairs[1][1])" } else { "$($ends.To) $mid $($ends.ToKind) mailbox" }
    $factRows = @(@('From', "$($ends.From) $mid $($ends.FromKind) $mid **only read**"), @('Into', $intoText), @('Store', "$store (stays on this PC)"),
        @('Size', "$(if ($everything) { 'Everything: ' })$itemsQuantity $mid $sizeText"), @('Time', (New-GuiText $etaText 'Body')))
    $facts = New-GuiFacts $factRows
    $warn = New-GuiMessage -Kind warn -Icon clock -Text '**This is a long job.** The first part is quick, the second is not: Microsoft accepts only about 4 GiB per hour into one mailbox, whatever the speed of your PC or network. You can stop and continue as often as you like: the part that is finished is not done again, and no item is added twice.'
    $g.PlanHolder = [Windows.Controls.ContentControl]::new(); $g.PlanHolder.Content = New-MbtGuiPlanStateMessage
    $free = $null; try { $free = [double](& $g.Engine.FreeSpace $s.WorkFolder) } catch { $free = $null }
    $need = [double]$bytes * 1.1
    $drive = if ($s.WorkFolder -match '^([A-Za-z]:)') { $Matches[1].ToUpperInvariant() } else { 'the disk' }
    $enough = ($null -eq $free) -or ($need -le $free)
    $diskRow = New-MbtGuiCheckRow -State $(if ($enough) { 'ok' } else { 'bad' }) -Text $(if ($enough) { 'Enough disk space for the store' } else { 'Not enough disk space for the store' }) -Detail $(if ($bytes -le 0) { "$(if ($null -ne $free) { Format-MbtGuiBytes $free } else { 'unknown' }) free on $drive" } else { "about $(Format-MbtGuiBytes $need) needed, $(if ($null -ne $free) { Format-MbtGuiBytes $free } else { 'unknown' }) free on $drive" })
    $placeRow = New-MbtGuiCheckRow -State $(if ($ends.Same) { 'bad' } else { 'ok' }) -Text $(if ($ends.Same) { 'The mailbox you copy from is the target' } else { 'A different place' }) -Detail "from $($ends.From) ($($ends.FromKind)) into $($ends.To) ($($ends.ToKind))"
    $ruleRow = New-MbtGuiCheckRow -State ok -Text $(if ($ends.App) { 'The application writes into the target' } else { 'You write into your own mailbox' }) -Detail $(if ($ends.App) { 'the target must be in the scope of the application; the check above looks at it' } else { "$($ends.To), the account that is signed in (a delegated sign-in cannot write into another mailbox)" })
    $checkRows = @($g.PlanHolder, (New-MbtGuiSignInCheckRow), $placeRow, $ruleRow, $diskRow)
    $checks = New-GuiCard -Title 'Checks' -Content $checkRows
    $left = New-GuiStack @($checks, (New-GuiCard -Title 'What will be done' -Content @($ol, $facts, $warn)))
    $before = New-GuiCard -Title 'Before you start' -Content @((New-GuiText ([char]0x2022 + '  Keep the PC plugged in and awake (the console does this for you).') 'Body'), (New-GuiText ([char]0x2022 + '  The mailbox you copy from is only read. The store stays on this PC after the copy: it is the proof of what was read.') 'Body'), (New-MbtGuiSleepInfo))
    $right = New-GuiStack @((New-MbtGuiMailCopyConfirmCard -ItemsQuantity $itemsQuantity -Bytes $bytes -SizeUnknown $sizeUnknown -AtMost $period.Active -Everything $everything), $before)
    $key = Get-MbtGuiPlanKey
    if ($s.PlanKey -ne $key) { $s.PlanState = 'idle'; $s.PlanKey = $key; Start-MbtGuiPlanCheck }
    if ($s.ConfirmKey -ne $key) { $s.Confirmed = $false; $s.ConfirmKey = '' }
    Set-MbtGuiActions -Footer (Get-MbtGuiPlanFooter) -Label "Start the copy into $($ends.To)" -Icon play -Enabled ($s.Confirmed -and $s.PlanState -eq 'ok' -and -not $ends.Same) -OnClick { Start-MbtGuiRun }
    return (New-GuiStack @((New-MbtGuiVerifyBottom -Columns (New-GuiTwoColumns -Left $left -Right $right -LeftWidth 520) -Task mailcopy -Include @('conc', 'retry', 'reconcile', 'tries', 'awake'))))
}

#endregion

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBY3D7bsmhyO89w
# DqHGd9/rWon3GDy0xe4mxWwE6Ci48qCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCubtYijcN+
# IjSZyE5sDV9nUCkDl+Gb32WY0aripgNjRjANBgkqhkiG9w0BAQEFAASCAgAEPsYQ
# DoqQIAAGvszhE5LGG9zJu0FkBB2dSUvyJNSBWJCJWseB190UJO63vhFt2Yv6v8mz
# UnPwJQxdjlciQSApJORdYkD9b/tR0azOnK6JIW2W+SB9H2IuPoEnUwxDXJn94BS1
# tWb4S8NavkpKQXHi/+bl2yNfadIgv8VeDv1S+w1nKAzMgNmXPy38vpTSpJhZaq2p
# TQ/wdUuyWwVxi+Ozjem+eiTVpDfnMIsflodTRPGjgQJDzPjtl2Bj1KjxE0bSasCN
# 9grjqtZVhIYTYn0wVlGck+qAGGIu4Mh8I/fEf5e1y1Y50VchiPG+++AoMDOEXd1/
# n52nbPAeEIqireDNmadTIws9htiyPXr1kNHe1zZXyeklUN1/DilKvapMFn5Qgjbd
# VXCaOnyZW/zhm1v81MGAySjQdIBSKWRtPYEOj42mgbdSQiCZtxKAzfASBjoQGGXd
# N7CYr9P39DkAmNboljd2dm+yLnDNVI5OkI1GZ5qgCBKrA/8F6j/zDEqvvw3H9FZ7
# VRxYzOt7ekNCZL7SiUHrjPQNQAs/yaWVsd++7gszIcG4WydCEaplinLkzxaxK5DW
# 57v8FMObe9+JC8dt/R12rwXBr8AkCOL3aKFl9n2UyF1Ft4YSr6jRvN+8uAiuIOox
# psLJmG1WdJzIPtamJA6b/66hWOHTT2Uq7Q7Je6GCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIK2LrwCxxDelpl+ff7F8E+trzKqzL4TfurykBZd5qnKlAghUipcywBIn
# DhgPMjAyNjExMjUwODAzMTVaMAMCAQECEAdWpnchs9l7jjWMPDbW5BSgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTEyNTA4MDMxNVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIJXy
# 6tZ1ECilju9Z/nF0bSsYucCSgEFpZ1X0jFZKUJ6eMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiEA2XoOVtpqNsC30y5JOk2J/TnOIg/+45UHvsbeZ/tf+uECIBp4
# iRJRpAE4OOggPQ2IQwluH06ej5axibMiecOFczm/
# SIG # End signature block
