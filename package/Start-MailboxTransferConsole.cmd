@echo off
setlocal
set POWERSHELL_TELEMETRY_OPTOUT=1
where pwsh.exe >nul 2>nul
if errorlevel 1 (
  echo Mailbox Transfer needs PowerShell 7.5 or later ^(pwsh.exe^). Install it with:  winget install Microsoft.PowerShell
  pause
  exit /b 1
)
pwsh.exe -NoLogo -NoProfile -Command "if ($PSVersionTable.PSVersion -lt [version]'7.5') { exit 1 }"
if errorlevel 1 (
  echo Mailbox Transfer needs PowerShell 7.5 or later. Update it with:  winget upgrade Microsoft.PowerShell
  pause
  exit /b 1
)
rem The script checks the package (signatures and hashes) before it runs anything else, and starts itself again in a single-threaded apartment.
pwsh.exe -NoLogo -NoProfile -STA -File "%~dp0Start-MailboxTransferConsole.ps1" %*
set "result=%errorlevel%"
if not "%result%"=="0" pause
exit /b %result%
