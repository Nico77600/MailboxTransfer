# Mailbox Transfer

Exports Exchange Online mailboxes (primary, archive and auxiliary archives) to a verified local store or to PST files, imports PST files, stores and `.eml` files into a mailbox, and checks and repairs PST files. A guided Windows console and a PowerShell 7 command line. Author: Nicolas Fabert. Version 1.0.1.

This folder contains everything needed to run the tool: the command line (`Invoke-MailboxTransfer.ps1`), the console (`Start-MailboxTransferConsole.ps1` and `.cmd`, `module\`), the C# engine (`src\`), SQLite (`lib\sqlite\`), the configuration and the two guides (user guide and reference guide, HTML with their pictures). Tests and build tools stay outside it, in the repository.

> [!IMPORTANT]
> Windows marks files downloaded from the Internet. Once the ZIP is extracted, check the publisher first (right-click a `.ps1`, for example `Invoke-MailboxTransfer.ps1`, Properties, Digital Signatures: `Nicolas Fabert`), then unblock the extracted folder once:
>
> ```powershell
> Get-ChildItem "C:\Tools\MailboxTransfer" -Recurse -File -Force | Unblock-File
> ```
>
> Replace the example path with the folder where you extracted the ZIP.

## Requirements

- Windows 10 or 11 (x64 or arm64).
- PowerShell 7.5 or later (`pwsh.exe`) for the console, 7.4 for the command line. Windows PowerShell 5.1 is refused. PowerShell 7 is not part of this folder: `winget install Microsoft.PowerShell`.
- The module `Microsoft.Graph.Authentication` (the tool borrows its MSAL library to sign you in): `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser`.
- For a mailbox: one app registration for your company with the minimum permission of each task, listed in the [user guide, chapter 2](docs/MailboxTransfer-UserGuide.html). Checking and repairing a PST, and opening an export, need no right.
- Free disk: about twice the size of the data.
- Not needed: Outlook, Exchange Online PowerShell, any third-party PST tool.

## Quick start

```powershell
cd C:\Tools\MailboxTransfer
Get-ChildItem . -Recurse -File -Force | Unblock-File
.\Start-MailboxTransferConsole.cmd                  # the console (or: .\Start-MailboxTransferConsole.ps1)
.\Start-MailboxTransferConsole.ps1 -Demo            # made-up data: no network, no mailbox
.\Invoke-MailboxTransfer.ps1 -Mode CheckPst -Path D:\pst\mail.pst     # the command line: read only
```

Your own tenant id and application id go in `config\MailboxTransfer.config.local.psd1` (you create it; it is read after the shipped file). Keep exports, PST files and logs outside this folder, or in its `out\` and `logs\` folders.

## Package integrity

The tool runs only from a package signed by its owner. `src\Test-PackageIntegrity.ps1` and `PackageIntegrity.psd1` must carry a valid, timestamped signature of `0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6`; every PowerShell file must be signed by the same owner; every other file must match the signed manifest. The entry scripts and the module check this first. A changed, missing, unlisted or reparse-point file stops the tool before anything runs. There is no unsigned mode; do not disable your execution policy or this check to clear an error: download the release again.

Edit only `config\MailboxTransfer.config.psd1` and `config\MailboxTransfer.config.local.psd1`: they are data, outside the check, like the generated `out\` and `logs\` folders.

## Content

| Item | Role |
|---|---|
| `Invoke-MailboxTransfer.ps1` | The command line: `Export`, `ExportPst`, `Import`, `ImportPst`, `ExportEml`, `ImportEml`, `Browse`, `Status`, `CheckPst`, `RepairPst`. |
| `Start-MailboxTransferConsole.ps1`, `Start-MailboxTransferConsole.cmd` | The console (WPF) and its double-click launcher. |
| `MailboxTransfer.psd1`, `MailboxTransfer.psm1` | The module: sign-in, configuration, mailbox tree, export, import, store, PST, EML. |
| `module\` | The console (pages, state, jobs), the events file of a job, the EML commands. |
| `src\` | The C# engine (`Engine.*.cs`, compiled in memory at the first use, no compiled copy is stored), the console view models and the integrity verifier. |
| `lib\sqlite\` | SQLite for the export store (see `THIRD-PARTY-NOTICES.md`). |
| `config\` | `MailboxTransfer.config.psd1` (every key optional, no private value). |
| `docs\` | User guide and reference guide (self-contained HTML), the events contract. |
| `PackageIntegrity.psd1`, `product.json` | Owner-signed manifest of the files; identity of the product. |
| `README.md`, `LICENSE`, `THIRD-PARTY-NOTICES.md`, `CHANGELOG.md`, `SECURITY.md` | Documents. |

```powershell
Import-Module .\MailboxTransfer.psd1
$s = Connect-MbtGraph -TenantId '<tenant id>' -ClientId '<application id>' -UserPrincipalName 'adele.vance@contoso.com'
$tree = Get-MbtMailboxTree -MailboxId 'MBX:<mailbox guid>@<tenant id>'
```

## Documentation

- [User guide](docs/MailboxTransfer-UserGuide.html): for a beginner, one step after the other, with a glossary and the answers to what goes wrong.
- [Reference guide](docs/MailboxTransfer-Guide.html): parameters, modes, exit codes, the exact rights, the store and PST formats, EML fidelity, limits.
- [Events contract](docs/MailboxTransfer-Events.md): the events file, the stop file and the job folder.

## Disclaimer

This Script is a Personal project.
It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.

An import writes to a real mailbox: try it on a test mailbox first.
