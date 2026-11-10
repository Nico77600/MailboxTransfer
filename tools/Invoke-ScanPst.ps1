#Requires -Version 7.4
<#
.SYNOPSIS
    OPTIONAL DEVELOPER COMPARISON TOOL. Runs Microsoft's Inbox Repair Tool (SCANPST.EXE) on a COPY of a PST file, without any click
    from you, and reports what it found. The scan only reads (the Repair button is never pressed); the original file is not touched.

    The product does not use this script: writing, reading and checking a PST (PstWriter, PstReader, Test-MbtPst) never need
    Outlook, SCANPST or any external PST tool, and no test of the default suite requires them. This script exists only to compare
    the verdict of our own checker with Microsoft's tool while developing the writer; run it by hand when SCANPST is installed.

.DESCRIPTION
    SCANPST has no command line and no documented automation. This script drives its dialog (UI Automation + window messages, in a
    Windows PowerShell 5.1 child process, because UI Automation needs the .NET Framework): it copies the PST to a temporary folder,
    types the path of the copy in the "file to scan" box (never the default one, which is your OST), presses Start, waits for the
    verdict, reads it and reads the log file the tool writes next to the copy, then closes the tool.

    The verdict is the text of SCANPST ("No errors detected", "Only minor inconsistencies were found", "This file contains errors"),
    with the number of folders and items it found. The lines of the log that start with "!!" are the errors, those that start with
    "??" are warnings: they are returned and printed.

    It needs a desktop session (the dialog is real, it is visible for a few seconds) and Outlook closed. Do not run it against a
    PST that is open in Outlook.

    Exit code: 0 when SCANPST reports no error or only minor inconsistencies and its log has no "!!" line, 1 otherwise.

.PARAMETER Path
    The PST file to check (a copy is scanned).

.PARAMETER ScanPstPath
    SCANPST.EXE (default: Office16 under Program Files, then Program Files (x86)).

.PARAMETER TimeoutSec
    How long to wait for the verdict (default 900 s; the scan reads the whole file).

.PARAMETER KeepCopy
    Keep the copy and its log (default: removed, the log text is returned in the result).

.EXAMPLE
    .\tools\Invoke-ScanPst.ps1 -Path .\out\pst-validation\corpus2.pst

.NOTES
    Author  : Nicolas Fabert
    Version : 0.1.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$Path,
    [string]$ScanPstPath,
    [int]$TimeoutSec = 900,
    [switch]$KeepCopy
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "PST not found: $Path" }
if (-not $ScanPstPath) {
    $ScanPstPath = @('C:\Program Files\Microsoft Office\root\Office16\SCANPST.EXE', 'C:\Program Files (x86)\Microsoft Office\root\Office16\SCANPST.EXE',
        'C:\Program Files\Microsoft Office\Office16\SCANPST.EXE', 'C:\Program Files (x86)\Microsoft Office\Office16\SCANPST.EXE') | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}
if (-not $ScanPstPath -or -not (Test-Path -LiteralPath $ScanPstPath)) { throw 'SCANPST.EXE not found (use -ScanPstPath).' }

$work = Join-Path ([IO.Path]::GetTempPath()) ('scanpst-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$null = New-Item -ItemType Directory -Path $work
$copy = Join-Path $work 'scan.pst'
Copy-Item -LiteralPath $Path -Destination $copy
$resultFile = Join-Path $work 'result.json'

# The UI Automation part (Windows PowerShell 5.1). Dialog controls of SCANPST: 1004 file name, 1012 Start, 1043 verdict text,
# 1048 Repair (never pressed), 1049 Cancel / Close.
$helper = @'
param([string]$Exe, [string]$Pst, [string]$Out, [int]$TimeoutSec)
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type @"
using System; using System.Runtime.InteropServices; using System.Text;
public static class W32 {
  [DllImport("user32.dll")] public static extern IntPtr GetDlgItem(IntPtr h, int id);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, string l);
  [DllImport("user32.dll")] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, IntPtr l);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr SendMessage(IntPtr h, int msg, IntPtr w, StringBuilder l);
}
"@
$p = Start-Process -FilePath $Exe -PassThru
try {
    $h = [IntPtr]::Zero; $edit = [IntPtr]::Zero
    for ($i = 0; $i -lt 60 -and $edit -eq [IntPtr]::Zero; $i++) { Start-Sleep -Milliseconds 500; $p.Refresh(); $h = $p.MainWindowHandle; if ($h -ne [IntPtr]::Zero) { $edit = [W32]::GetDlgItem($h, 1004) } }
    if ($edit -eq [IntPtr]::Zero) { throw 'The SCANPST window did not appear.' }
    [void][W32]::SendMessage($edit, 0x000C, [IntPtr]::Zero, $Pst)                 # WM_SETTEXT
    $sb = [Text.StringBuilder]::new(2048); [void][W32]::SendMessage($edit, 0x000D, [IntPtr]2048, $sb)
    if ($sb.ToString() -ne $Pst) { throw 'Could not set the file name in SCANPST.' }
    [void][W32]::SendMessage([W32]::GetDlgItem($h, 1012), 0x00F5, [IntPtr]::Zero, [IntPtr]::Zero)   # BM_CLICK Start
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $cond = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, [int]$p.Id)
    $end = (Get-Date).AddSeconds($TimeoutSec); $verdict = $null; $stable = 0; $last = ''
    while ((Get-Date) -lt $end -and -not $p.HasExited) {
        Start-Sleep -Milliseconds 1500
        $texts = @{}
        foreach ($w in $root.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)) {
            foreach ($e in $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
                if ($e.Current.Name) { $texts[$e.Current.AutomationId] = $e.Current.Name }
            }
        }
        $now = ($texts.Keys | Sort-Object | ForEach-Object { "$_=$($texts[$_])" }) -join "`n"
        if ($now -eq $last) { $stable++ } else { $stable = 0; $last = $now }
        if ($texts.ContainsKey('1043') -and $stable -ge 2) { $verdict = $texts['1043']; break }
    }
    $res = [ordered]@{ Verdict = $verdict; TimedOut = ($null -eq $verdict) }
    $res | ConvertTo-Json | Set-Content -LiteralPath $Out -Encoding UTF8
} finally {
    if (-not $p.HasExited) { Stop-Process -Id $p.Id -Force }
}
'@
$helperFile = Join-Path $work 'helper.ps1'
Set-Content -LiteralPath $helperFile -Value $helper -Encoding UTF8

try {
    Write-Host "SCANPST on a copy of $Path ..."
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $helperFile -Exe $ScanPstPath -Pst $copy -Out $resultFile -TimeoutSec $TimeoutSec
    if (-not (Test-Path -LiteralPath $resultFile)) { throw 'SCANPST could not be driven (no result).' }
    $res = Get-Content -LiteralPath $resultFile -Raw | ConvertFrom-Json
    $logFile = Join-Path $work 'scan.log'
    $log = ''
    for ($try = 0; $try -lt 20 -and (Test-Path -LiteralPath $logFile) -and $log -eq ''; $try++) {      # the tool may still hold the log for a moment
        try { $log = [IO.File]::ReadAllText($logFile, [Text.Encoding]::GetEncoding(1252)) -replace "`r", '' } catch { Start-Sleep -Milliseconds 500 }
    }
    $errors = @($log -split "`n" | Where-Object { $_ -match '^\s*!!' } | ForEach-Object { $_.Trim() })
    $warnings = @($log -split "`n" | Where-Object { $_ -match '^\s*\?\?' } | ForEach-Object { $_.Trim() })
    $counts = [regex]::Matches([string]$res.Verdict, '\d+') | ForEach-Object { [int]$_.Value }
    $out = [pscustomobject]@{
        Path = (Resolve-Path -LiteralPath $Path).Path; Verdict = $res.Verdict; TimedOut = $res.TimedOut
        Folders = if (@($counts).Count -ge 1) { @($counts)[0] } else { $null }; Items = if (@($counts).Count -ge 2) { @($counts)[1] } else { $null }
        ErrorLines = $errors; WarningLines = $warnings; Log = $log; Ok = (-not $res.TimedOut -and $errors.Count -eq 0)
    }
    Write-Host ($(if ($res.TimedOut) { 'No verdict within the timeout.' } else { ($res.Verdict -replace '\s+', ' ').Trim() }))
    foreach ($l in $errors) { Write-Host "  $l" -ForegroundColor Red }
    foreach ($l in $warnings) { Write-Host "  $l" -ForegroundColor Yellow }
    if ($KeepCopy) { Write-Host "Copy and log kept in $work" }
    $out
    if (-not $out.Ok) { exit 1 }
}
finally {
    if (-not $KeepCopy -and (Test-Path -LiteralPath $work)) { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}
