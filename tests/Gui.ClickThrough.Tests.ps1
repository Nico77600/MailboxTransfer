#Requires -Version 7.6
# Console: every button, check box, radio button, list, text box, fold, folder and message of every page of every scenario is used in a real window (fake engine), and no
# handler may raise an exception. A handler that uses a variable of the function that built the page fails at the first click (the function is gone when the control
# raises its event); this is the test that finds such a handler. Also: no handler is subscribed to a control outside the ways that cover it.
BeforeAll {
    $script:Package = Join-Path (Split-Path $PSScriptRoot -Parent) 'package'
    $script:Script = Join-Path $PSScriptRoot 'support\gui-clickthrough.ps1'
    $script:Pwsh = (Get-Process -Id $PID).Path
    # the scenarios are shared between several processes (WPF needs its own single-threaded apartment in each); the last one also walks the extra states
    $script:Parts = 4
    $script:Runs = foreach ($part in 0..($script:Parts - 1)) {
        $work = Join-Path ([IO.Path]::GetTempPath()) ('mbt-gui-click-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
        $stdout = Join-Path $work 'stdout.json'
        $null = New-Item -ItemType Directory -Force -Path $work
        # Start-Process does not quote the elements of -ArgumentList: each path is quoted (a path may contain a space)
        $process = Start-Process -FilePath $script:Pwsh -ArgumentList @('-NoProfile', '-STA', '-File', "`"$($script:Script)`"", '-Package', "`"$($script:Package)`"", '-Work', "`"$(Join-Path $work 'w')`"", '-Part', $part, '-Parts', $script:Parts) -RedirectStandardOutput $stdout -RedirectStandardError (Join-Path $work 'stderr.txt') -WindowStyle Hidden -PassThru
        [pscustomobject]@{ Part = $part; Work = $work; Stdout = $stdout; Process = $process }
    }
    foreach ($run in $script:Runs) {
        $null = $run.Process.WaitForExit(1500000)
        $raw = if (Test-Path -LiteralPath $run.Stdout) { (Get-Content -LiteralPath $run.Stdout -Raw) } else { '' }
        $run | Add-Member -NotePropertyName Raw -NotePropertyValue ([string]$raw).Trim()
        $run | Add-Member -NotePropertyName Result -NotePropertyValue $null
        try { $run.Result = $run.Raw | ConvertFrom-Json } catch { }
        $run | Add-Member -NotePropertyName Err -NotePropertyValue $(if (Test-Path -LiteralPath (Join-Path $run.Work 'stderr.txt')) { (Get-Content -LiteralPath (Join-Path $run.Work 'stderr.txt') -Raw) } else { '' })
    }
}
AfterAll { foreach ($run in $script:Runs) { Remove-Item -LiteralPath $run.Work -Recurse -Force -ErrorAction SilentlyContinue } }

Describe 'Console click-through (headless WPF, fake engine)' {
    It 'every part ran and reported' {
        foreach ($run in $script:Runs) { $run.Result | Should -Not -BeNullOrEmpty -Because "part $($run.Part): $($run.Raw) $($run.Err)" }
    }
    It 'walked every scenario and used every kind of control' {
        $scenarios = ($script:Runs | ForEach-Object { $_.Result.Scenarios } | Measure-Object -Sum).Sum
        $scenarios | Should -BeGreaterThan 50
        $controls = ($script:Runs | ForEach-Object { $_.Result.Controls } | Measure-Object -Sum).Sum
        $controls | Should -BeGreaterThan 800
        $kinds = @($script:Runs | ForEach-Object { $_.Result.Invoked.PSObject.Properties.Name } | Sort-Object -Unique)
        foreach ($kind in 'Button', 'ToggleButton', 'CheckBox', 'RadioButton', 'ComboBox', 'TextBox', 'Expander', 'TreeViewItem', 'ListBoxItem') { $kinds | Should -Contain $kind }
        $dialogs = ($script:Runs | ForEach-Object { $_.Result.Dialogs } | Measure-Object -Sum).Sum
        $dialogs | Should -BeGreaterThan 10 -Because 'the answers of the dialogs are used too'
    }
    It 'walked the states with exactly one of something, or none (one folder ticked alone, one message, one location, one finding, one step, one refused item, one incomplete message, one job)' {
        $edge = @($script:Runs | ForEach-Object { @($_.Result.Edge) } | Where-Object { $_ } | Sort-Object -Unique)
        $walked = @($script:Runs | ForEach-Object { @($_.Result.Walked) } | Where-Object { $_ })
        $edge.Count | Should -BeGreaterThan 12
        foreach ($name in $edge) { $walked | Should -Contain $name }
        foreach ($name in 'ExportPlanOneFolder', 'ExportPlanOneMessage', 'ExportFromOneLocation', 'ExportFromNoLocation', 'RepairFromOneFinding', 'RunOneStep', 'DoneImportOneRefused', 'DoneRepairOneIncomplete', 'DoneImportOneDecision', 'StartOneJob', 'ImportFromManyFolders') { $edge | Should -Contain $name }
    }
    It 'used the Pause button of the running step (and answered its question) and the Resume button of a stopped job' {
        $pause = ($script:Runs | ForEach-Object { $_.Result.StepControls.RunStepPause } | Measure-Object -Sum).Sum
        $resume = ($script:Runs | ForEach-Object { $_.Result.StepControls.RunStepResume } | Measure-Object -Sum).Sum
        $pause | Should -BeGreaterThan 0
        $resume | Should -BeGreaterThan 0
    }
    It 'no handler raised an exception, and the tick of the console raised nothing' {
        $errors = @($script:Runs | ForEach-Object { @($_.Result.Errors) } | Where-Object { $_ })
        $errors | Should -BeNullOrEmpty
    }
}

Describe 'Console handlers: how they are subscribed' {
    It 'a handler written in the function that builds a page is subscribed only through Add-MbtGuiHandler, New-GuiButton, New-GuiLink, New-GuiTextBox, New-GuiOption or New-GuiDatePicker' {
        $allowed = @{
            'MailboxTransfer.Gui.Core.ps1' = @('New-GuiButton', 'New-GuiLink', 'New-GuiTextBox', 'New-GuiOption', 'New-GuiDatePicker')
            'MailboxTransfer.Gui.Browse.ps1' = @('Initialize-MbtGuiBrowseEvents')
        }
        $found = [Collections.Generic.List[string]]::new()
        foreach ($file in Get-ChildItem -LiteralPath (Join-Path $script:Package 'module') -Filter 'MailboxTransfer.Gui*.ps1') {
            $tokens = $null; $errors = $null
            $ast = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$errors)
            foreach ($call in $ast.FindAll({ param($n) $n -is [Management.Automation.Language.InvokeMemberExpressionAst] -and $n.Member.Extent.Text -match '^(?i)add_' }, $true)) {
                $function = $call.Parent; while ($function -and $function -isnot [Management.Automation.Language.FunctionDefinitionAst]) { $function = $function.Parent }
                $name = if ($function) { $function.Name } else { '(script)' }
                if (-not ($allowed.ContainsKey($file.Name) -and $name -in $allowed[$file.Name])) { $found.Add("$($call.Member.Extent.Text) in $name ($($file.Name), line $($call.Extent.StartLineNumber))") }
            }
        }
        $found | Should -BeNullOrEmpty -Because 'a block that is not subscribed through them has no access to the variables of the function that wrote it'
    }
    It 'the functions that subscribe a block to a control by themselves hand it to ConvertTo-MbtGuiHandler' {
        $core = Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Core.ps1') -Raw
        foreach ($function in 'New-GuiButton', 'New-GuiLink', 'New-GuiTextBox', 'New-GuiOption', 'New-GuiDatePicker') {
            $body = [regex]::Match($core, "(?s)function $function \{.*?\r?\n\}").Value
            $body | Should -Match 'ConvertTo-MbtGuiHandler' -Because "$function subscribes a handler"
        }
        $core | Should -Match '(?s)function Add-MbtGuiHandler \{.*ConvertTo-MbtGuiHandler'
    }
    It 'the browse events of the tree and of the list raise blocks that use no variable of their own: functions and script variables only' {
        $browse = Get-Content -LiteralPath (Join-Path $script:Package 'module\MailboxTransfer.Gui.Browse.ps1') -Raw
        $browse | Should -Match "(?m)add_ExpandRequested\(\[Action\[MailboxTransferGui\.FolderNode\]\] \{ param\(\`$node\) Read-MbtGuiChildren \`$node \}\)"
        $browse | Should -Match "(?m)add_SelectionChanged\(\[Action\] \{ Update-MbtGuiBrowseTrays \}\)"
    }
}
