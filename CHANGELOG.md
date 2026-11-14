# Changelog

All notable changes of Mailbox Transfer. Author: Nicolas Fabert. Versions follow `MAJOR.MINOR.PATCH`: a MAJOR version changes the command line or the configuration file, a MINOR version adds a feature, a PATCH version fixes a defect.

## 1.2.0 — unreleased

### Changed
- **A simpler console, in five steps**: Setup, Task (What, From, To), Verify, Run, Next steps. The new Setup page shows the lights of the computer, the sign-in (**Me** or **The application**, with the values read from the configuration file; a value changed in the window applies to that run only) and the work folder.
- **Import confirmation**: instead of typing a word, you tick a box that names the target mailbox, the folder and the number of items. Start stays grey until you do, and the tick clears if the plan changes.
- **The command line is shown** on the Verify and Next steps pages, with a Copy button, and the Next steps page offers the logical next buttons.
- Tries per item are 3 in the console, as on the command line.
- The configuration section an operator fills in is now **`Tenant`**; a file that still says `Lab` keeps working.
- **Sign-in in a Windows window (no browser tab)**: a Microsoft sign-in window opens over the console and you type the account. The code stays as the alternative. Your administrator adds one redirect address to the application (user guide, chapter 2).
- **Keep this PC awake** is off by default and moved to the Expert mode of the Verify pages; Setup no longer has it.
- **An export prepares faster and says what it does**: the folder tree of a large mailbox is read by up to 4 requests at once, 1,000 folders to the page, and once per run. The Verify page shows what it is doing ("1,240 folders read", with the time), Start says why it waits, then shows "Starting...", and the Run page shows "Preparing" with the same progress before the first item. The disk check says "There is enough disk space."
- **The plan check is quicker and more reliable**: a folder with subfolders is counted from the folder down, not from the tree of the whole mailbox (about 6 s instead of a minute on an archive of 4,443 folders); calendar, contact and task folders show their real size when they are small; a temporary Microsoft error (HTTP 502) is tried again, and a check that still fails says so in plain words with a **Check again** button.
- **An export of a few folders no longer reads the whole mailbox tree**: the folders chosen, those below them and the ones above them are read (9 s instead of 77 s for 3 folders of an archive of 4,443 folders). The store gets the same folders and items. Recoverable Items and auxiliary archives still read the whole tree.

### Added
- **An HTML report with a chronology** at the end of every job (`report.html` in the job folder, path on the `Report` line of the summary card): the result, the numbers, each step with its start, end and duration, the waits for Microsoft, the sleeps of the PC, the warnings and the errors in time order, one row for each folder and the next step. One file, no network. **Open the report** opens it; `report.csv` stays, and a job of an older version still opens its CSV.
- **The report of a very large job stays small and quick**: the folders and the chronology are capped at 2,000 rows each (the folders that are not done come first), the page says how many rows were left out and where they are, and a job of 50,000 folders and 195,000 events is reported in under a minute in a page of about 1 MB.
- **One page for all your jobs**: `jobs\overview.html` has one row for each job (started, task, mailbox, destination, result, items, data, duration, link to its report), totals, the jobs that failed, stopped or wait for a decision first, filters by result, task and period, a search and sortable columns; it opens in thousands of jobs at once. It is rewritten at the end of every job (the summary card gives its path on the `All jobs` line), and the new **All jobs** button of the console opens it.
- **One command for a multi-location export**: the primary mailbox, the archive and the auxiliary archives ticked in the console are exported by a single `-SelectionFile` command. New `-RetryWaitSeconds`.

### Fixed
- Works with **PowerShell 7.6** (the console of 1.1.0 closed at the first click).
- The Setup page shows the work folder and its free space with the real engine. The certificate sign-in uses the tenant of the configuration when it has none of its own.
- Every number is written the same way (en-US), whatever the language of Windows.
- Names, addresses, paths and job names are **not case sensitive**: resuming a job, or importing again, no longer creates duplicates because of a capital letter.

## 1.1.0 — 2026-11-10

### Changed
- **PST export**: the folder of an auxiliary archive (`-SeparateAuxArchives`) now has an English name, "Auxiliary archive 1 (xxxxxxxx)". An export started with 1.0.x and continued keeps the name it already has, and a PST made by 1.0.x is read and imported as before.
- README banner redrawn to the same template as the other projects.

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
