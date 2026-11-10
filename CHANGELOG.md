# Changelog

All notable changes of Mailbox Transfer. Author: Nicolas Fabert. Versions follow `MAJOR.MINOR.PATCH`: a MAJOR version changes the command line or the configuration file, a MINOR version adds a feature, a PATCH version fixes a defect.

## 1.0.1 — 2026-11-10

Documentation only: no change to the engine, the console or the command line.

### Changed
- README and guide pictures redrawn to the house standard (white or dark background, 2160 px wide, framed banner); a new Get started picture; the real console pages shown in the README.
- The unblock note says to check the publisher first, with the example folder `C:\Tools\MailboxTransfer`.

### Fixed
- The permission picture of the user guide no longer says that writing into a shared mailbox with an app-only sign-in is "coming": it has been available since 1.0.0 (reference guide, 4.6).

## 1.0.0 — 2026-11-10

First public version.

### Added
- **Guided console** (`-Gui`, English): Task, From, To, Plan, Run, Done. Browse a mailbox before exporting it, tick what to keep, settings for the current run only, real step progress, and an honest time range for long jobs.
- **Export** of the primary mailbox, the archive and the auxiliary archives of auto-expanding archives to a local store (verified, resumable) or to **PST files** (real folder tree, split at 50 GB by default, each file checked), or to EML files.
- **Import** from a store, a PST or EML files into a mailbox: create only, resumable, uncertain items looked for by the tool; a person is asked only when two items could be the one.
- **PST check and repair** (own implementation, no Outlook needed): a repaired copy with a before and after report; the original is never changed.
- **Open** an export or a PST and browse it like a mailbox.
- Shared mailboxes by address: browse and export with the FullAccess right.
- **Sign-in with a certificate** (app-only, no person at the keyboard): `-CertificateThumbprint` / `-CertificatePath` on every mode and "The application (certificate)" in the console, for scheduled or bulk jobs and for writing into shared mailboxes. Permissions per task, with Exchange RBAC for Applications to limit the application to some mailboxes.
- **Choose folders** of a PST or of `.eml` files to import (`-IncludeFolder`, `-ExcludeFolder`, `-FolderFile`, ticks in the console); a PST is read in the reading pane, and the ticks can be copied to new PST files (`Copy-MbtPst`, `-Mode ExportPst -PstPath`).
- The plan shows the item counts and sizes counted by the service; the Done page explains items refused for good, with a CSV.
- **Deleted items**: export Recoverable Items (`-IncludeRecoverableItems`, a box in the console), kept apart in the store and the PST; an import puts them in a normal folder, *Restored deleted items*. A soft-deleted or inactive mailbox now stops with a clear explanation.
- A user guide for beginners (glossary, troubleshooting) and a reference guide with the minimum permissions for each task.
- **Signed release**: every PowerShell file carries the author's Authenticode signature and the package a signed manifest of the SHA-256 of every other file; the scripts and the module verify them before anything runs, and say so when a source checkout is started with `-DevelopmentMode`.
- Output in the style of the author's other tools: a banner, numbered-looking step headers, one icon per result line, a summary card with the exit code, the counts and the duration, and a daily log file (`logs\MailboxTransfer_<date>.log`). `MBT_ICONS` and `NO_COLOR` control icons and colours.
- The console has a theme switch (System, Light, Dark) and links to the user guide and to the latest release.

### Changed
- One exit-code rule for every mode: 3 (uncertain items), then 2 (resumable), then 1 (error or items refused for good), then 0.

### Fixed
- Long jobs: stop honoured while waiting to retry, no stalled download, upload time that follows the size, secrets of pre-authenticated URLs never logged.
- Safer files: a PST export never deletes a file that is not one of its own parts; the import journal and the token cache survive a power cut and two processes; spreadsheet formulas in `report.csv` are neutralised.
- A moved message is exported once; an uncertain `.eml` import can be settled (`Resolve-MbtEmlUnknown`).
- Exporting one folder to a PST no longer writes the empty folder tree of the whole mailbox: the store keeps the folders exported and the path to them.
- A mailbox id written in the configuration is used only for its own user; a sign-in without tenant or application stops with a clear message.

Validated offline and on one Microsoft 365 test tenant, including writing into a shared mailbox with a certificate.
