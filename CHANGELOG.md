# Changelog

All notable changes of Mailbox Transfer. Author: Nicolas Fabert. Versions follow `MAJOR.MINOR.PATCH`: a MAJOR version changes the command line or the configuration file, a MINOR version adds a feature, a PATCH version fixes a defect.

## Unreleased

### Added
- **Export only some dates**: `-ArrivedFrom` and `-ArrivedTo` keep the items that arrived in the mailbox between two days (the console: the **Period** line under the browser of the From page). One period for the whole export, a list of mailboxes included. Messages you tick one by one are always exported; an item with no date is kept. The items left out are skipped, not an error; the summary, the report and the page of all the jobs say how many.

### Changed
- **Runs from a OneDrive-synced folder**: the package check judges a link by its type, not by the link attribute. Symbolic links, junctions and unknown types are still refused; OneDrive, deduplicated and CompactOS files are accepted and still hashed, and a OneDrive file that cannot be downloaded offline is named with the way out.
- **The browser fits the window**: the folder tree and the list of messages are no longer cut off when the window is small.
- **Pause and Resume on the running step**: a small button in the place of its icon pauses the job (stop and keep progress) and resumes it where it stopped.

## 1.2.0 — 2026-11-14

### Added
- **Export many mailboxes from a list**: `-Mode Export -MailboxCsv list.csv` exports the mailboxes of a CSV file, one ordinary job for each, one after the other or up to 4 at a time (`-ParallelMailboxes`). The file is checked first and every problem is printed together; sign-in is done once; the same command again (or `-ResumeBatch`) continues what is not complete. A mailbox that needs a sign-in counts as not complete, not failed. `batches\<name>\report.html` is the report of the batch. Export only.
- **An HTML report for every job**, with a chronology: the result, the numbers, each step with its duration, the waits for Microsoft, the warnings and the errors in time order. One file, no network; `report.csv` stays. A job of 50,000 folders is reported in under a minute in a page of about 1 MB.
- **One page for all your jobs**: `jobs\overview.html` lists every job with its result, items, data and duration, the ones that need a look first, filters and a search. It is rewritten at the end of every job, and the **All jobs** button of the console opens it. A job links to its report, else its CSV, else its status.
- **Old job folders are tidied away** after 90 days (`Jobs.RetentionDays`, 0 keeps everything); a job that can still be continued, your stores, selections, PST and EML files are never touched.
- **One command for a multi-location export**: the primary mailbox, the archive and the auxiliary archives ticked in the console are exported by one `-SelectionFile` command. New `-RetryWaitSeconds`.

### Changed
- **A simpler console, in five steps**: Setup, Task (What, From, To), Verify, Run, Next steps. Setup shows the lights of the computer, the sign-in (**Me** or **The application**) and the work folder.
- **Import confirmation**: you tick a box that names the target mailbox, the folder and the number of items, instead of typing a word.
- **The command line is shown** on the Verify and Next steps pages, with a Copy button.
- **Sign-in in a Windows window** over the console, no browser tab; the code stays as the alternative. Your administrator adds one redirect address to the application (user guide, chapter 2).
- **An export prepares faster and says what it does**: the folder tree of a large mailbox is read by several requests at once, and the console shows the progress while it waits. An export of a few folders reads only those folders.
- **The plan check is quicker and more reliable**: a folder is counted from the folder down, a temporary Microsoft error is tried again, and a check that fails says so with a **Check again** button.
- The configuration section an operator fills in is now **`Tenant`**; a file that still says `Lab` keeps working. **Keep this PC awake** moved to the Expert mode of the Verify pages. Tries per item are 3 in the console, as on the command line.

### Fixed
- The console no longer freezes at the end of a very large job.
- Works with **PowerShell 7.6** (the console of 1.1.0 closed at the first click).
- Names, addresses, paths and job names are **not case sensitive**: resuming or importing again no longer creates duplicates because of a capital letter. Every number is written the same way (en-US), whatever the language of Windows.

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
