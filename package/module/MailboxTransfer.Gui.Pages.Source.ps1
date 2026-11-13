#Requires -Version 7.4
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
            if ($Use -eq 'import') { $browser.TrayEmpty = "Nothing ticked: **everything will be imported** ($(Format-MbtGuiQuantity $info.Items 'item'), $(Format-MbtGuiBytes $info.Bytes))."; Update-MbtGuiTray $browser }
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
        $left.Children.Add((New-GuiMessage -Kind info -Icon info -Text "**Tick folders to import only those.** A ticked folder takes the folders below it. Nothing ticked: the whole of $what is imported. Single messages are not chosen here; the folders the tick leaves out are not created.")) | Out-Null
    }
    $right = [Windows.Controls.Grid]::new()
    if ($s.ImInfo -and $s.ImInfo.Ok) { Place-MbtGuiBrowser $browser $right }
    else {
        $ico = New-GuiIcon pst 22 'MbtBrandText'; $ico.HorizontalAlignment = 'Center'
        $t1 = New-GuiText 'Pick the PST file or the export. It is checked at once, and you can look inside.' 'Muted'; $t1.TextAlignment = 'Center'; $t1.Margin = [Windows.Thickness]::new(0, 8, 0, 0)
        $body = New-GuiStack @($ico, $t1); $body.Margin = [Windows.Thickness]::new(10, 60, 10, 60)
        $right = New-GuiCard -Title 'What is inside' -Content @($body)
    }
    $ok = [bool]($s.ImInfo -and $s.ImInfo.Ok)
    Set-MbtGuiActions -Footer 'Pick what to import. The file is checked first.' -Label 'Continue to the mailbox' -Enabled $ok -OnClick { Move-MbtGuiStep 1 }
    $stack = New-GuiStack @((New-GuiText 'What do you want to import?' 'H2'), (New-GuiTwoColumns -Left $left -Right $right))
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
        $items.Add((New-MbtGuiSharedMailboxSection -Use import))
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
    $sel = Get-MbtGuiSelection -Roots @($browser.Roots)
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
    return (New-GuiTwoColumns -Left $left -Right $right)
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

