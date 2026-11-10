#Requires -Version 7.4

<#
.SYNOPSIS
    Mailbox Transfer - exports an Exchange Online mailbox (primary or In-Place archive) into a local store and imports the store into
    a mailbox, with the Microsoft Graph mailbox import and export API.

.DESCRIPTION
    This is the owner's own tooling to migrate or back up mailboxes the operator is authorised to access (delegated consent of the
    mailbox user). Throttling is obeyed, not worked around: Retry-After is honoured exactly, the concurrency is halved when the service
    answers 429 / 503 and grows back slowly, it never exceeds the 4 concurrent requests per mailbox the service documents, and one
    application and one identity are used.

    Mode Export   Folder tree, inventory of the items, then the items in batches of 20 items / 200 MiB by -Concurrency workers (4 at
                  most), streamed into the packs of the store. Running it again does only what is left: Pending items, Failed items
                  tried fewer than -MaxAttempts times, items whose change key moved. -Mailbox Archive exports the In-Place archive
                  (redirections to an auxiliary archive of an auto-expanding archive are followed). -IncludeRecoverableItems also exports the
                  deleted items kept in Recoverable Items (Deletions, Purges, Versions, DiscoveryHolds; -RecoverableItemsFolders names others), stored
                  under one top folder "Recoverable Items" in the store and in the PST, never mixed with the normal folders. An import of such an export
                  puts them in a normal folder named "Restored deleted items", never back into Recoverable Items.
    Mode ExportPst  Converts the store into PST files with our own writer (no Outlook, no third-party component): the real folder tree of each
                  location under one top folder (Primary, Archive, "Archive auxiliaire n (guid8)"), the standard folders first, split at -PstMaxSizeGB
                  (default 50 = 51,200 MiB) on item boundaries, <name>.pst, <name>-002.pst... and <name>.manifest.json. The auxiliary archives are merged
                  into the Archive tree (same folder entry key) unless -SeparateAuxArchives. Resumable: run it again after an interruption, the finished
                  parts (verified by size and PST header) are kept. Each part is checked (Test-MbtPst) unless -SkipCheck. The store is only read.
                  -Mode Export -OutputPst <file.pst> does the export, then the same conversion.
    Mode Import   Creates the folder tree in the target mailbox (below -TargetRoot, or at the same paths, or following -MappingFile;
                  a folder is found by its name under its parent, never renamed or deleted) and imports the Exported items of the
                  store (create only). One import session per mailbox is reused until it is about to expire. Items are Pending ->
                  InFlight -> Done / Failed (refused: nothing created) / Unknown (timeout, 5xx: the item may exist). An Unknown row does
                  not stop the run: after -ReconcileDelaySeconds the target folder is listed and the item looked for (same size, not
                  known as imported, created after the attempt began): found once = Done, not found = imported again in the same run
                  (-MaxAttempts respected), several candidates = left Unknown and reported. Unknown rows left by an earlier run are
                  looked up at the start. -NoAutoReconcile keeps them Unknown. Running it again with the same -JobName resumes.
                  -WhatIf shows the plan and creates nothing.
    Auto-expanding archive (Mode Export and Import): the archive of a user may be made of the main archive and auxiliary archives. -Mailbox
                  Archive exports the main archive AND the items of its expanded folders, which live in the auxiliary archives: the
                  folder is recorded once (main archive tree), each item with the mailbox that really holds it. -Mailbox AuxArchive
                  -MailboxId MBX:... exports one auxiliary archive by itself, -Mailbox All exports the primary, the archive and every
                  auxiliary archive known. The import follows the service when a folder lives in an auxiliary archive (see -MailboxId).
                  The auxiliary archives are LEARNED from the redirects of the service and recorded in the store (shown by Mode Status);
                  -LocationsFile gives them up front.
    Mode ImportPst Imports one or more PST files into a mailbox (create only), without Outlook: -PstPath is a PST file, several files or a folder of PST files (the parts
                  of a split PST, name.pst, name-0002.pst, ..., are taken in order as one tree). The folder tree of the PST is recreated below -TargetRoot (container classes
                  kept; the Inbox, Sent Items ... of the PST become subfolders of it and are never merged into the real Inbox unless -MergeWellKnown), -MappingFile maps
                  PST path prefixes to target path prefixes, -PerFileSubfolder gives each file its own subfolder. Each file is checked first with the integrity checker:
                  a damaged file is refused, unless -Repair (then it is repaired into a temporary copy, the file itself is never modified). Each message is read
                  from the PST, converted to a FastTransfer stream and sent with importItem; the journal (<TempPath>\pstimport.db, one row per message keyed by file, node
                  id and content fingerprint) lets the same command resume and never create an item twice; an item whose answer is missing is reconciled like in Mode Import.
                  An item that cannot be read, cannot be expressed or is refused by the service is reported with the reason and the run goes on. -WhatIf prints the plan
                  (files, folders, items, size, estimated duration at about 4 GiB per hour) and creates nothing. Exit codes as Mode Import.
    Mode Browse   Live view of a mailbox BEFORE any export (read only: nothing is written to the store or to the mailbox). Without -Folder: the folder
                  tree with the item count and size of each folder. With -Folder: the rows of that folder, newest first, -Top at a time (subject,
                  sender, To/Cc counts, dates, size, attachments, importance, sensitivity, item class, categories, flag). With -Folder -ItemId: the
                  preview of an item (header, start of the text body, attachments, attached items). -Mailbox Primary and Archive are read with the
                  Microsoft Graph Mail API (delegated Mail.Read; the archive is addressed as a mailbox of its own, /users/MBX:<archive guid>@<tenant>).
                  -Mailbox AuxArchive is not reachable by the Mail API: the tree and the size and dates of the items come from the import-export API
                  and the details come after the export (previews of the store). Folders the Mail API does not list (Calendar, Contacts, Tasks...) are
                  merged from the import-export API; -RecoverableItems adds Recoverable Items (exportable). -SaveSelection writes what was chosen (the
                  folder, with -IncludeSubfolders, or the -ItemId items) as a selection file, checked against the import-export API and printed as
                  the export plan; -Mode Export -SelectionFile exports exactly that.
    Mode Status   The state of the store: trees with the counts and sizes by state, packs, import jobs. -Unknown lists the import rows
                  in the state Unknown (-Check also looks for the item in the target folder); -ResolveUnknown <importKey> -As Imported
                  | NotImported settles one. -Unknown -Check -ResolveAbsent settles at once every row for which the check found no
                  matching item in the target folder (puts it back to Pending; -WhatIf lists them without settling; -JobName limits it to
                  one job); the rows with candidates stay Unknown. Without a store: version, configuration and whether the engine compiles.

    While a transfer runs the script writes a progress line every -IntervalSeconds (items, bytes, MiB/s overall and for the last 30 s,
    GiB/h, number of 429, concurrency) and a line per window of 5 minutes (sustained rate, items, 429). At the end every 429 / 503 is
    listed with the time waited. Ctrl+C stops after the requests in flight; run again to resume.

    The end of an import tells what THIS run did and what the WHOLE job looks like (done / failed / unknown / pending of all the
    runs): rows left Unknown by an earlier run (a killed process turns its in-flight rows into Unknown at the next opening) count
    against the job, and the job is complete only when every row is Done. It also counts what the automatic reconciliation did
    (found created, imported again after a lookup, left ambiguous).

    Mode CheckPst  Checks a PST file (-Path), read only (Test-MbtPst): exit 0 clean, 2 warnings (or stopped), 1 errors.
    Mode RepairPst Repairs a PST file (-Path) into a NEW file (-Destination; never overwritten, the original is never changed), at -Level
                  (default Automatic = the level the check recommends, with escalation), see Repair-MbtPst. Exit 0 repaired and clean, 2 stopped or errors remain, 1 error.

    Console contract: -EventsFile writes the events of the job as JSON lines (docs\MailboxTransfer-Events.md) and status.json beside it; -StopFile is
    a file whose appearance stops the job like Ctrl+C (exit 2, run again to continue); -KeepAwake stops Windows from sleeping by idle while it runs.

.OUTPUTS
    Exit code (Mode Export): 0 the export is complete; 1 an error stopped it; 2 it is not complete but resumable, nothing was lost: a folder was DEFERRED
    (the service cannot read its content yet, typically a folder being moved to an auxiliary archive of an auto-expanding archive: it is listed in the
    summary with the message of the service, never exported as empty), a folder could not be listed, an item failed, or the run was stopped. Run the same
    command again (later, for a deferred folder): it does only what is left.
    Exit code (Mode Import and ImportPst): 0 the job is complete (every row Done); 1 an error stopped the run; 2 the job is not complete but only
    Pending / Failed rows remain (run the same command again); 3 the job holds Unknown rows (with the automatic reconciliation: several
    items could be the one, or -ReconcileRounds lookups were not enough; they are never replayed by themselves: settle them with
    -Mode Status -Unknown -Check, then run again). Other modes: 0, or 1 on error.
    Mode ImportPst has one more case of exit code 1: the job is FINISHED (every row is Done or Refused) but some items were refused for good, and the result says
    reason 'refused-items' (not an error: there is nothing to run again). A Refused item is final: the service refuses its content (HTTP 400, 413, 415, 422), it
    cannot be read from the PST or expressed as a FastTransfer stream, or it failed with nothing created on every attempt allowed (-MaxAttempts). They are listed one
    per line, with the reason, in <TempPath>\<JobName>-not-imported.csv. Order of the codes: 3, then 2, then 1, then 0.

.PARAMETER Mode
    Export, ExportPst, Import, Browse, CheckPst, RepairPst or Status (default).

.PARAMETER DeviceCode
    Sign in with a device code instead of the browser. The line "DEVICE CODE: <code> URL: <url>" is written at once to the console and
    to out\device-code.txt. A cached token (DPAPI) makes it silent.

.PARAMETER AppOnly
    Sign in as the application with a certificate (client credentials, no user, no prompt) using Auth.CertificateThumbprint or Auth.CertificatePath of the configuration. Implied by -CertificateThumbprint or -CertificatePath. See docs\MailboxTransfer-Guide.md, chapter 4.6.

.PARAMETER CertificateThumbprint
    App-only: the thumbprint of the application's certificate (CurrentUser\My, then LocalMachine\My; it needs its private key). Works for every mode. -TenantId and -ClientId are then those of the certificate application (Auth.AppTenantId / Auth.AppClientId of the configuration); the lab application is never used.

.PARAMETER CertificatePath
    App-only: a .pfx file with the certificate and its private key (-CertificatePassword: SecureString).

.PARAMETER CertificatePassword
    App-only: the password of the .pfx, as a SecureString.

.PARAMETER UserPrincipalName
    The mailbox user (default: Lab.UserPrincipalName of the configuration). Sign-in is delegated as this user. With -CertificateThumbprint / -CertificatePath it is the OWNER OF THE TARGET MAILBOX (there is no signed-in user and no /me), it is required, and the lab default is never used.

.PARAMETER Mailbox
    Primary (default), Archive (the main archive; with Export its expanded folders are followed into the auxiliary archives), AuxArchive
    (one auxiliary archive of an auto-expanding archive: needs -MailboxId, or a single one known) or All (Export only: Primary, Archive,
    then every auxiliary archive known): the mailbox exported from (Export) or imported into (Import). The ids come from -MailboxId,
    -PrimaryMailboxId / -ArchiveMailboxId, the Lab section of the configuration (for the user of the configuration only), -LocationsFile, the locations recorded in the store,
    or (Primary; Archive too with User.Read.All and Graph beta) Graph.

.PARAMETER MailboxId
    One explicit location, MBX:<mailbox guid>@<tenant id>, the bare guid, or the SMTP address / UPN of a mailbox (a SHARED mailbox or another user's: it is resolved
    to its id through Graph settings/exchange when the token allows it (User.Read.All), else through the import-export API; the sign-in stays -UserPrincipalName, who
    needs FullAccess on that mailbox, see docs/MailboxTransfer-Guide.md, chapter 4.5; -Mailbox then defaults to Primary): the mailbox to export from or import into. Without -Mailbox its
    kind is the one the store or -LocationsFile knows for it. For an import into an auxiliary archive the folders are looked up (and, if
    missing below -TargetRoot, created) in that mailbox's own tree.

.PARAMETER LocationsFile
    The output of Exchange Online PowerShell, which this tool does not need installed: Get-MailboxLocation -User <upn> |
    Select-Object MailboxGuid, MailboxLocationType | Export-Csv (or ConvertTo-Json, or Format-Table into a text file). Gives the primary,
    the main archive and the auxiliary archives up front; they are recorded in the store.

.PARAMETER Locations
    Status: signs in and lists the locations of the user (Graph when the token allows it, -LocationsFile, parameters, the store), with how
    each one was learned. Without it Status shows the locations recorded in the store.

.PARAMETER StorePath
    Folder of the store (catalog.db + packs): default <Output.Root>\transfer\store. Created by Export, reopened (with its crash
    recovery) later. Only one process at a time may write to it; Status can read it meanwhile.

.PARAMETER PstPath
    ImportPst: the PST file(s) to import, or a folder holding them (split parts in order). ExportEml: the one PST to read the messages from. ExportPst: the PST file(s) to copy from
    (instead of a store): the chosen folders and messages (-IncludeFolder, -ExcludeFolder, -IncludeItemFile; none: everything) are written into new PST files at -OutputPst, split at
    -PstMaxSizeGB, each file checked unless -SkipCheck; the source is only read, a stopped copy leaves no file. Exit 0 finished, 1 error or a file with check errors or messages that
    could not be copied (the copy is finished, reason 'refused-items'), 2 stopped.

.PARAMETER MergeWellKnown
    ImportPst: the well-known folders of the PST (Inbox, Sent Items, Deleted Items, Drafts, Outbox, Junk, Calendar, Contacts, Tasks, Notes, Journal) go into the matching folders of the mailbox instead of subfolders of -TargetRoot.

.PARAMETER PerFileSubfolder
    ImportPst: each PST file gets its own subfolder (named after its store) below -TargetRoot.

.PARAMETER Repair
    ImportPst: a PST file that has errors is repaired (-RepairLevel Recompute, Rebuild or Salvage; -KeepDamagedAttachments) into a temporary copy in <TempPath>\pstimport-work and the copy is imported; without it the file is refused. -NoCheck skips the integrity check.

.PARAMETER SourceMailbox
    Import: which mailbox of the store to import from (Primary or Archive); needed when the store holds both.

.PARAMETER TargetRoot
    Import: path in the target below which everything is created, for instance MBT-Phase1-A. Without it the same paths as the source.

.PARAMETER MappingFile
    Import: .psd1 or .json, source path prefix -> target path prefix (relative to -TargetRoot). Longest prefix wins.

.PARAMETER NoAutoReconcile
    Import: do not look up the Unknown rows by itself (today's behaviour: they stay Unknown, exit code 3, settle with -Mode Status -Unknown -Check).

.PARAMETER ReconcileDelaySeconds
    Import: seconds between the end of the request that left a row Unknown and the lookup in the target folder (default 300: the service
    may still be creating the item). For the Unknown rows of an earlier run the delay counts from the last change of the row.

.PARAMETER ReconcileRounds
    Import: how many times an item may be looked up and imported again before it is left Unknown (default 3, 1 to 20).

.PARAMETER Concurrency
    Workers, 1 to 4 (default 4). The engine lowers it by itself while the service throttles.

.PARAMETER NotReadyTrip
    Export: after this many folders in a row whose content is not available yet (an auxiliary archive being filled), the next folders are asked once instead of 3 times with a back-off, and deferred if they still refuse (default 3; 0 = always 3 asks). A folder that lists again puts the full asks back.

.PARAMETER Folder
    Export, Import: only this folder and the folders below it (path as in the store). Browse: the folder to list (id or path, or the end of a path
    such as Inbox/Reports).

.PARAMETER IncludeFolder
    ImportPst, ImportEml, ExportPst (from -PstPath): only some folders of the PST or of the folder of .eml files. Paths below the top folder of the PST (Top of Personal Folders)
    or below -EmlPath ("Inbox", "Inbox/Projects"), wildcards * (any text, slashes included) and ? (one character), case does not matter. A path chooses the folder and the folders
    below it; "Inbox/*" is what is below the Inbox without the Inbox itself; "*/2019" every folder called 2019. Several paths: comma separated. A path that matches no folder is
    reported, and a run in which no folder is chosen stops before anything is written.

.PARAMETER IncludeItemFile
    ExportPst from -PstPath: a text file, one PST node id (a number) per line: messages to copy one by one, in addition to the folders of -IncludeFolder. With this alone only these messages are copied.

.PARAMETER ResolveEml
    ImportEml: settles the .eml files an earlier import left Unknown (the request may have created the item), after a look in the mailbox: Imported (it is there: Done) or NotImported (it is
    not: the next run imports it). Without -EmlFile every Unknown file; -EmlFile: relative paths as the state lists them. No sign-in, nothing is sent.

.PARAMETER EmlFile
    ImportEml with -ResolveEml: the files to settle.

.PARAMETER FolderFile
    A text file of folder choices, UTF-8, one per line: "+path" is an -IncludeFolder, "-path" an -ExcludeFolder (added to the ones given on the command line). The console writes it.

.PARAMETER ExcludeFolder
    The same patterns, to leave folders out (and what is below them). An exclusion wins over an inclusion. To take the Inbox without its subfolders: -IncludeFolder Inbox -ExcludeFolder 'Inbox/*'.

.PARAMETER Top
    Browse: how many rows of the folder to print (default 20; 0 = all of them).

.PARAMETER ItemId
    Browse: with -Folder, the item(s) to preview: a row number of the listing (1 = the first row) or an item id. With -SaveSelection, the items
    to select instead of the whole folder.

.PARAMETER Sort
    Browse: order of the rows (Mail API): ReceivedDescending (default), ReceivedAscending, SentDescending, SentAscending, SubjectAscending,
    SubjectDescending. The import-export API (auxiliary archive) has no sort.

.PARAMETER NoRecipients
    Browse: do not read To and Cc (Exchange reads the recipients of each message for it: slower on folders of meeting messages).

.PARAMETER RecoverableItems
    Browse: add Recoverable Items and its folders to the tree (the tree and counts come from the import-export API, the messages from the Mail API). They can be selected and exported
    like any folder. -RecoverableItemsFolders limits the folders shown; without it the six folders of the service are listed.

.PARAMETER IncludeRecoverableItems
    Export: also export Recoverable Items, the hidden folders where Exchange keeps what a user deleted from Deleted Items (Deletions), what was purged (Purges), the versions of changed
    items (Versions) and what a hold keeps (DiscoveryHolds). Read through the import-export API, full fidelity. Stored under one top folder "Recoverable Items/<name>" in the store
    and in the PST, never mixed with the normal folders. Primary, Archive and AuxArchive. Needs the whole mailbox (or -Folder "Recoverable Items/<name>"): with -SelectionFile, select the folders
    of Recoverable Items in the browse view instead. An import never puts them back into Recoverable Items: see "Restored deleted items" in the guide.

.PARAMETER RecoverableItemsFolders
    Export (with -IncludeRecoverableItems, or with -SelectionFile to choose the folders under a selected "Recoverable Items"): the folders of Recoverable Items to take, by name. Default: Deletions, Purges,
    Versions, DiscoveryHolds. The names are Calendar Logging, Deletions, DiscoveryHolds, Purges, SubstrateHolds and Versions; Calendar Logging and SubstrateHolds are Microsoft's own copies
    and are only exported when named here.

.PARAMETER BodyChars
    Browse: characters of the text body shown in a preview (default 600; 0 = the short excerpt the service gives, bodyPreview, and no body request).

.PARAMETER SaveSelection
    Browse: writes the selection (the -Folder, with -IncludeSubfolders, or the -ItemId items of it) to this JSON file and prints the export plan.

.PARAMETER IncludeSubfolders
    Browse -SaveSelection: the folder and every folder below it.

.PARAMETER ShowMoves
    Browse (tree): also reads the other side of an auto-expanding archive and marks the folders that exist in the main archive and in an auxiliary archive:
    "being moved to aux archive 1" (the items are still in the main archive, the auxiliary archive holds the skeleton), "moved to ..." or "items on both
    sides". The folder is recognised by the folder entry key of its id. -Mailbox Archive does it by itself when an auxiliary archive is known (it is a small
    tree); for -Mailbox AuxArchive it needs the tree of the main archive (the Mail API: slow on thousands of folders), hence this switch.

.PARAMETER OutputPst
    Export, ExportPst: the first PST file (D:\out\mail.pst), or a folder (the file is named after the user of the store, "<user>-AuxArchive1" for an auxiliary archive alone; a Windows device name such as aux.pst is refused). The next files are
    mail-002.pst, mail-003.pst... and the manifest is mail.manifest.json. With Export the store is converted after the export; ExportPst converts
    the store of -StorePath (a store that was exported earlier). -Mailbox / -MailboxId / -SelectionFile restrict what is converted (without them:
    everything in the store).

.PARAMETER PstMaxSizeGB
    Export, ExportPst: largest PST in GB (default 50 = 51,200 MiB, Outlook's default; at least 1). Items are never split: an item larger than the
    limit is written alone in its own file and reported.

.PARAMETER PstMaxSizeMiB
    Export, ExportPst: the same limit in MiB (smaller than 1 GB: tests, small drives). Wins over -PstMaxSizeGB.

.PARAMETER SeparateAuxArchives
    Export, ExportPst: one top folder per auxiliary archive ("Archive auxiliaire 1 (guid8)") instead of merging them into the Archive tree.

.PARAMETER SkipCheck
    Export, ExportPst: do not run the integrity checker (Test-MbtPst) on each finished PST.

.PARAMETER Force
    Export, ExportPst: start the PST conversion again from zero (deletes the manifest and the PST files of an earlier run of the same destination).

.PARAMETER SelectionFile
    Export: a selection written by -Mode Browse -SaveSelection (or by the console). The mailbox is the one of the file; the selection is checked
    against the import-export API (a folder is found there by the same id, else by its path; selected items by id, else by size and date) and only
    the selected folders / items are inventoried and exported. An entry that cannot be exported (a folder or item not found)
    stops the run before anything is exported. With -WhatIf only the plan is printed.

.PARAMETER EventsFile
    Writes the events of the job (plan, step, object, item, wait, rate, warn, decision, error, result, slept, info) to this file, one JSON object per line, and
    status.json (state of the job for the Continue list) in the same folder. Format: docs\MailboxTransfer-Events.md. The text on the console is unchanged.

.PARAMETER StopFile
    A stop request: when this file exists the job stops as Ctrl+C stops it (the requests in flight finish, the state is saved), exit code 2, result next 'continue'.

.PARAMETER JobId
    The identifier of the job in the events (default: the name of the folder of -EventsFile).

.PARAMETER LogFolder
    Where the log file of the run is written: MailboxTransfer_yyyyMMdd.log (one line per message, no colour, no secret; the runs of a day append to it; files older than 14 days are deleted).
    Default: the logs folder of Output.Root (.\out\logs), or the folder named by the environment variable MAILBOXTRANSFER_LOG_FOLDER.

.PARAMETER Gui
    Opens the console (Start-MailboxTransferConsole.ps1): the window that guides browse and export, import, check and repair of a PST, and open.
    Every other parameter is ignored; the window starts the jobs itself, with this script.

.PARAMETER KeepAwake
    Windows does not go to sleep by idle while the script runs (it cannot stop a closed lid, Sleep chosen by hand or a restart). Released when the script ends.

.PARAMETER NoPrompt
    Silent sign-in only (no browser, no device code): when no valid token is cached the script stops with exit code 2 and the result reason 'signin-required'.

.PARAMETER Path
    CheckPst, RepairPst: the PST file.

.PARAMETER Destination
    RepairPst: the new file (never overwritten; a Windows device name such as aux.pst is refused).

.PARAMETER Level
    RepairPst: Automatic (default), Recompute, Rebuild or Salvage.

.PARAMETER KeepDamagedAttachments
    RepairPst: keep an attachment with unreadable blocks (zero-filled, renamed 'name (damaged).ext') instead of dropping it.

.EXAMPLE
    .\Invoke-MailboxTransfer.ps1 -Mode CheckPst -Path D:\pst\mail.pst
    .\Invoke-MailboxTransfer.ps1 -Mode RepairPst -Path D:\pst\mail.pst -Destination D:\pst\mail.repaired.pst -Level Automatic
    .\Invoke-MailboxTransfer.ps1 -Mode Export -DeviceCode -StorePath D:\store
    .\Invoke-MailboxTransfer.ps1 -Mode Export -Mailbox Archive -StorePath D:\store
    .\Invoke-MailboxTransfer.ps1 -Mode Import -Mailbox Archive -SourceMailbox Primary -TargetRoot MBT-Restore -StorePath D:\store -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode ImportPst -PstPath D:\old\archive.pst -TargetRoot Imported/archive -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode ImportPst -PstPath D:\old\parts -TargetRoot Imported/archive -Mailbox Archive
    .\Invoke-MailboxTransfer.ps1 -Mode Browse -Mailbox Archive
    .\Invoke-MailboxTransfer.ps1 -Mode Browse -Folder Inbox -Top 10 -Sort SubjectAscending
    .\Invoke-MailboxTransfer.ps1 -Mode Browse -Folder Inbox -ItemId 2
    .\Invoke-MailboxTransfer.ps1 -Mode Browse -Folder Inbox/Reports -IncludeSubfolders -SaveSelection D:\sel.json
    .\Invoke-MailboxTransfer.ps1 -Mode Export -SelectionFile D:\sel.json -StorePath D:\store
    .\Invoke-MailboxTransfer.ps1 -Mode Export -Mailbox Archive -StorePath D:\store -OutputPst D:\pst\adele.pst -PstMaxSizeGB 10
    .\Invoke-MailboxTransfer.ps1 -Mode ExportPst -StorePath D:\store -OutputPst D:\pst\adele.pst
    .\Invoke-MailboxTransfer.ps1 -Mode ExportPst -StorePath D:\store -OutputPst D:\pst\adele.pst -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode Export -Mailbox All -StorePath D:\store -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode Status -StorePath D:\store
    .\Invoke-MailboxTransfer.ps1 -Mode Status -StorePath D:\store -Unknown -Check
    .\Invoke-MailboxTransfer.ps1 -Mode Status -StorePath D:\store -Unknown -Check -ResolveAbsent -WhatIf

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.1
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

    As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Export', 'ExportPst', 'Import', 'ImportPst', 'Browse', 'Status', 'ExportEml', 'ImportEml', 'CheckPst', 'RepairPst')][string]$Mode = 'Status',
    [string]$ConfigPath,
    [switch]$DeviceCode,
    [switch]$AppOnly,
    [string]$CertificateThumbprint,
    [string]$CertificatePath,
    [securestring]$CertificatePassword,
    [string]$TenantId,
    [string]$ClientId,
    [string]$UserPrincipalName,
    [ValidateSet('Primary', 'Archive', 'AuxArchive', 'All')][string]$Mailbox = 'Primary',
    [string]$MailboxId,
    [string]$LocationsFile,
    [switch]$Locations,
    [string]$PrimaryMailboxId,
    [string]$ArchiveMailboxId,
    [string]$StorePath,
    [string]$TempPath,
    [ValidateRange(1, 4)][int]$Concurrency = 4,
    [int]$IntervalSeconds = 5,

    [string]$Folder,
    [string]$SelectionFile,
    [switch]$SkipInventory,
    [int]$MaxAttempts = 3,
    [ValidateRange(0, 100000)][int]$NotReadyTrip = 3,
    [ValidateRange(1, 20)][int]$BatchItems = 20,
    [ValidateRange(1, 4096)][int]$BatchMiB = 200,

    [ValidateSet('Primary', 'Archive')][string]$SourceMailbox,
    [string]$TargetRoot,
    [string[]]$PstPath,
    [switch]$MergeWellKnown,
    [switch]$PerFileSubfolder,
    [string[]]$IncludeFolder,
    [string[]]$ExcludeFolder,
    [string]$IncludeItemFile,
    [string]$FolderFile,
    [ValidateSet('Imported', 'NotImported')][string]$ResolveEml,
    [string[]]$EmlFile,
    [switch]$Repair,
    [ValidateSet('Recompute', 'Rebuild', 'Salvage')][string]$RepairLevel = 'Rebuild',
    [switch]$KeepDamagedAttachments,
    [switch]$NoCheck,
    [string]$MappingFile,
    [string]$JobName = 'import',
    [switch]$NoAutoReconcile,
    [ValidateRange(0, 86400)][double]$ReconcileDelaySeconds = 300,
    [ValidateRange(1, 20)][int]$ReconcileRounds = 3,

    [ValidateRange(0, 100000)][int]$Top = 20,
    [string[]]$ItemId,
    [ValidateSet('ReceivedDescending', 'ReceivedAscending', 'SentDescending', 'SentAscending', 'SubjectAscending', 'SubjectDescending')][string]$Sort = 'ReceivedDescending',
    [switch]$NoRecipients,
    [switch]$RecoverableItems,
    [switch]$IncludeRecoverableItems,
    [string[]]$RecoverableItemsFolders,
    [ValidateRange(0, 20000)][int]$BodyChars = 600,
    [string]$SaveSelection,
    [switch]$IncludeSubfolders,
    [switch]$ShowMoves,

    [string]$OutputPst,
    [ValidateRange(1, 1024)][double]$PstMaxSizeGB = 50,
    [ValidateRange(1, 1048576)][long]$PstMaxSizeMiB,
    [switch]$SeparateAuxArchives,
    [switch]$SkipCheck,
    [switch]$Force,
    [string]$EmlPath,
    [ValidateSet('Live', 'Store', 'Pst')][string]$EmlSource,
    [ValidateSet('Own', 'Exchange')][string]$MimeConversion = 'Own',
    [switch]$Unread,
    [switch]$NoResume,
    [switch]$NoTransportHeaders,

    [switch]$Unknown,
    [switch]$Check,
    [switch]$ResolveAbsent,
    [long]$ResolveUnknown,
    [ValidateSet('Imported', 'NotImported')][string]$As,
    [string]$TargetItemId,

    # console contract: events, stop request, keep awake, silent sign-in; PST check and repair
    [string]$EventsFile,
    [string]$StopFile,
    [string]$JobId,
    [switch]$Gui,
    [switch]$DevelopmentMode,
    [string]$LogFolder,
    [switch]$KeepAwake,
    [switch]$NoPrompt,
    [string]$Path,
    [string]$Destination,
    [ValidateSet('Automatic', 'Recompute', 'Rebuild', 'Salvage')][string]$Level = 'Automatic'
)

if ($Gui) { & (Join-Path $PSScriptRoot 'Start-MailboxTransferConsole.ps1') -DevelopmentMode:$DevelopmentMode; exit $LASTEXITCODE }

# The package is verified before any of its code runs: the verifier must carry the owner signature, then it checks the signed manifest.
# A tree without a manifest (a source checkout) is accepted only with -DevelopmentMode.
$integrityManifest = Join-Path $PSScriptRoot 'PackageIntegrity.psd1'
if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $integrityManifest) {
    $integrityVerifier = Join-Path $PSScriptRoot 'src\Test-PackageIntegrity.ps1'
    $integrityCursor = [IO.Path]::GetFullPath($integrityVerifier)
    while ($integrityCursor) {
        $integrityItem = Microsoft.PowerShell.Management\Get-Item -LiteralPath $integrityCursor -Force -ErrorAction Stop
        if ($integrityItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Package integrity: reparse point: $integrityCursor" }
        $integrityCursor = [IO.Path]::GetDirectoryName($integrityCursor)
    }
    $integritySignature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $integrityVerifier -ErrorAction Stop
    if ([string]$integritySignature.Status -cne 'Valid' -or $null -eq $integritySignature.SignerCertificate -or
        $integritySignature.SignerCertificate.Thumbprint -cne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
        $null -eq $integritySignature.TimeStamperCertificate) {
        throw 'Package integrity: verifier must have the trusted, timestamped owner signature. Use the complete release ZIP (https://github.com/Nico77600/MailboxTransfer/releases/latest): extract it again and unblock it. A source checkout is started with -DevelopmentMode.'
    }
    if ($DevelopmentMode) { & $integrityVerifier -Root $PSScriptRoot -DevelopmentMode | Out-Null } else { & $integrityVerifier -Root $PSScriptRoot -Distribution | Out-Null }
} elseif (-not $DevelopmentMode) {
    throw 'Package integrity: PackageIntegrity.psd1 is missing. Use the complete release ZIP (https://github.com/Nico77600/MailboxTransfer/releases/latest): extract it again and unblock it. A source checkout is started with -DevelopmentMode.'
} else {
    Write-Warning 'DevelopmentMode: no package manifest; package byte integrity and publisher authentication are not established.'
    # The module does not repeat this warning in the same process.
    [AppDomain]::CurrentDomain.SetData('MailboxTransfer.UnverifiedTreeWarned', $true)
}


# -FolderFile: the folders chosen, one per line, "+path" to take and "-path" to leave out (what the console writes; paths with commas or spaces need no quoting)
if ($FolderFile) {
    if (-not (Test-Path -LiteralPath $FolderFile -PathType Leaf)) { throw "-FolderFile not found: $FolderFile" }
    foreach ($line in Get-Content -LiteralPath $FolderFile -Encoding utf8) {
        if ($line.Length -lt 2) { continue }
        if ($line[0] -eq '+') { $IncludeFolder = @($IncludeFolder | Where-Object { $_ }) + $line.Substring(1) }
        elseif ($line[0] -eq '-') { $ExcludeFolder = @($ExcludeFolder | Where-Object { $_ }) + $line.Substring(1) }
    }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# -WhatIf only means "plan the import": the files of the run (store, transcript) are always written.
$mailboxGiven = $PSBoundParameters.ContainsKey('Mailbox')
$dryRun = [bool]$WhatIfPreference
$WhatIfPreference = $false
Import-Module (Join-Path $PSScriptRoot 'MailboxTransfer.psd1') -Force

$config = if ($ConfigPath) { Import-MbtConfiguration -Path $ConfigPath } else { Import-MbtConfiguration }
if (-not $TenantId) { $TenantId = $config.Lab.TenantId }
if (-not $ClientId) { $ClientId = $config.Lab.ClientId }
if (-not $UserPrincipalName) { $UserPrincipalName = $config.Lab.UserPrincipalName }
# the mailbox ids written in the configuration belong to the user of the configuration: for another user they are discovered, never borrowed
$configured = Get-MbtConfiguredMailboxId -Configuration $config -UserPrincipalName $UserPrincipalName
if (-not $PrimaryMailboxId) { $PrimaryMailboxId = $configured.Primary }
if (-not $ArchiveMailboxId) { $ArchiveMailboxId = $configured.Archive }
# App-only (certificate): the application, the tenant and the target mailbox are always given (parameters or the Auth.App* keys), never the lab defaults of a delegated sign-in.
if (-not $CertificateThumbprint -and -not $CertificatePath -and $AppOnly) { $CertificateThumbprint = $config.Auth.CertificateThumbprint; $CertificatePath = $config.Auth.CertificatePath }
$useCertificate = [bool]($CertificateThumbprint -or $CertificatePath)
if ($AppOnly -and -not $useCertificate) { throw '-AppOnly needs -CertificateThumbprint or -CertificatePath (or Auth.CertificateThumbprint / Auth.CertificatePath in the configuration).' }
if ($useCertificate) {
    if (-not $PSBoundParameters.ContainsKey('TenantId')) { $TenantId = $config.Auth.AppTenantId }
    if (-not $PSBoundParameters.ContainsKey('ClientId')) { $ClientId = $config.Auth.AppClientId }
    if (-not $ClientId) { throw 'App-only: give -ClientId (the application with the certificate) or set Auth.AppClientId in the configuration.' }
    if (-not $TenantId) { throw 'App-only: give -TenantId (tenant id or domain) or set Auth.AppTenantId in the configuration.' }
    if (-not $PSBoundParameters.ContainsKey('UserPrincipalName')) { $UserPrincipalName = $null }
    if (-not $PSBoundParameters.ContainsKey('PrimaryMailboxId')) { $PrimaryMailboxId = $null }
    if (-not $PSBoundParameters.ContainsKey('ArchiveMailboxId')) { $ArchiveMailboxId = $null }
}
$outRoot = if ([IO.Path]::IsPathRooted($config.Output.Root)) { $config.Output.Root } else { Join-Path (Get-Location).Path $config.Output.Root }
if (-not $StorePath) { $StorePath = Join-Path $outRoot 'transfer\store' }
if (-not $TempPath) { $TempPath = Join-Path $outRoot 'transfer' }
if (-not [IO.Path]::IsPathRooted($StorePath)) { $StorePath = Join-Path (Get-Location).Path $StorePath }

$pstMaxBytes = if ($PSBoundParameters.ContainsKey('PstMaxSizeMiB')) { $PstMaxSizeMiB * 1MB } else { [long]($PstMaxSizeGB * 1024) * 1MB }
if ($OutputPst -and $Mode -notin 'Export', 'ExportPst') { throw '-OutputPst belongs to -Mode Export (export, then convert) and -Mode ExportPst (convert a store).' }
if (($IncludeRecoverableItems -or $RecoverableItemsFolders) -and $Mode -notin 'Export', 'Browse') { throw '-IncludeRecoverableItems and -RecoverableItemsFolders belong to -Mode Export (and -Mode Browse, which has -RecoverableItems).' }
if ($RecoverableItemsFolders) {
    Initialize-MbtEngine
    [MailboxTransfer.RecoverableItems]::CheckNames([string[]]$RecoverableItemsFolders)
    if ($Mode -eq 'Export' -and -not $IncludeRecoverableItems -and -not $SelectionFile -and -not [MailboxTransfer.RecoverableItems]::IsPath($Folder)) { throw '-RecoverableItemsFolders needs -IncludeRecoverableItems (or -SelectionFile that selects Recoverable Items).' }
}
if ($Mode -eq 'Export' -and $IncludeRecoverableItems) {
    if ($SelectionFile) { throw '-IncludeRecoverableItems adds Recoverable Items to the export of the whole mailbox. With -SelectionFile, select the folders of Recoverable Items in the browse view (Browse -RecoverableItems, or the console) instead.' }
    Initialize-MbtEngine
    if ($Folder -and -not [MailboxTransfer.RecoverableItems]::IsPath($Folder)) { throw "-IncludeRecoverableItems adds Recoverable Items to the export of the whole mailbox; it cannot be combined with -Folder '$Folder'. Use -Folder 'Recoverable Items/Deletions' for the deleted items alone." }
}

# The frames and the icons of the console are Unicode; a redirected output (a job of the console, a scheduled task) is UTF-8 without BOM.
try { [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false) } catch { Write-Verbose "Console output encoding unchanged: $($_.Exception.Message)" }
if (-not $LogFolder) { $LogFolder = if ($env:MAILBOXTRANSFER_LOG_FOLDER) { $env:MAILBOXTRANSFER_LOG_FOLDER } else { Join-Path $outRoot 'logs' } }
$logPath = $null
try { $logPath = Start-MbtLog -Directory $LogFolder } catch { Write-Warning "No log file ($LogFolder): $($_.Exception.Message)" }
$runClock = [Diagnostics.Stopwatch]::StartNew()
$bannerSubtitle = switch ($Mode) {
    'Export' { 'Exports a mailbox to a verified local store, then to PST files if asked' }
    'ExportPst' { 'Writes PST files from an export store, or from chosen folders of a PST' }
    'Import' { 'Puts the items of an export store into a mailbox (items are added, never deleted)' }
    'ImportPst' { 'Puts the items of a PST file into a mailbox (items are added, never deleted)' }
    'ExportEml' { 'Writes the items of a store as .eml files' }
    'ImportEml' { 'Puts .eml files into a mailbox (items are added, never deleted)' }
    'Browse' { 'Looks inside a mailbox, read only' }
    'Status' { 'Shows the state of an export store and of its import jobs' }
    'CheckPst' { 'Checks a PST file, read only' }
    'RepairPst' { 'Writes a repaired copy of a PST file; the original is never changed' }
}
$bannerDetails = [ordered]@{ Mode = @('Mode', $Mode + $(if ($dryRun) { ' - WhatIf (nothing is written)' })) }
if ($Mode -in 'Export', 'ExportPst', 'Import', 'Status', 'ExportEml') { $bannerDetails.Store = @('Store', $StorePath) }
$bannerSource = @($PstPath, $EmlPath, $Path | Where-Object { $_ } | ForEach-Object { [string]$_ }) -join ', '
if ($bannerSource) { $bannerDetails.Source = @('Folder', $bannerSource) }
if ($OutputPst) { $bannerDetails.Output = @('Folder', $OutputPst) }
if ($Destination) { $bannerDetails.Output = @('Folder', $Destination) }
if ($logPath) { $bannerDetails.Log = @('Log', $logPath) }
Write-MbtBanner -Title 'Mailbox Transfer' -Subtitle $bannerSubtitle -Details $bannerDetails
Initialize-MbtEngine
$script:counts = @{}
$script:reason = $null
$script:next = $null
$script:stateOverride = $null
$script:sawResult = $false
$script:resultComplete = $false
# The card that ends every run: what happened (exit code rule: 3 uncertain items, 2 resumable, 1 error or items refused for good, 0 complete), the counts, the log, what to do next.
function Write-MbtRunSummary {
    param([int]$ExitCode, [string]$Failure)
    $meaning = switch ($ExitCode) {
        0 { 'complete' }
        1 { 'failed or some items were refused for good' }
        2 { 'not complete: nothing was lost, run the same command again to continue' }
        3 { 'uncertain items: settle them with -Mode Status -Unknown -Check' }
        default { 'see the messages above' }
    }
    $status = switch ($ExitCode) { 0 { 'Ok' } 1 { 'Fail' } default { 'Warn' } }
    $title = switch ($ExitCode) { 0 { "$Mode finished" } 1 { "$Mode failed" } default { "$Mode stopped before the end" } }
    if ($dryRun -and $ExitCode -eq 0) { $title = "$Mode planned (WhatIf, nothing was written)" }
    $values = [ordered]@{ Result = @($(if ($ExitCode -eq 0) { 'Ok' } elseif ($ExitCode -eq 1) { 'Fail' } else { 'Warn' }), "exit code $ExitCode - $meaning") }
    if ($Failure) { $values.Error = @('Fail', $Failure) }
    $parts = @(foreach ($k in ($script:counts.Keys | Sort-Object)) { $v = $script:counts[$k]; if ($v -is [ValueType] -or $v -is [string]) { '{0} {1}' -f $k, $(if ($v -is [long] -or $v -is [int]) { Format-MbtNumber $v } else { $v }) } })
    if ($parts) { $values.Counts = @('Store', ($parts -join ', ')) }
    $values.Duration = @('Clock', (Format-MbtDuration $runClock.Elapsed.TotalSeconds))
    if ($script:next -and $script:next -notin 'done') { $values.Next = @('Plan', [string]$script:next) }
    if ($logPath) { $values.Log = @('Log', $logPath) }
    Write-MbtSummary -Title $title -Values $values -Status $status
}

if ($StopFile) { Start-MbtStopWatch -Path $StopFile }
if ($KeepAwake) { [void](Start-MbtKeepAwake) }
if ($EventsFile) {
    $taskName = switch ($Mode) { 'Export' { 'export' } 'ExportEml' { 'export' } 'ExportPst' { 'exportpst' } 'Import' { 'import' } 'ImportPst' { 'import' } 'ImportEml' { 'import' } 'CheckPst' { 'check' } 'RepairPst' { 'repair' } 'Browse' { 'browse' } default { 'status' } }
    $title = switch ($Mode) {
        'Export' { "Export - $UserPrincipalName" } 'ExportPst' { "Write PST files - $OutputPst" } 'Import' { "Import - $UserPrincipalName" }
        'ImportPst' { "Import from PST - $UserPrincipalName" } 'ExportEml' { "Export EML files - $EmlPath" } 'ImportEml' { "Import EML files - $UserPrincipalName" }
        'CheckPst' { "Check - $Path" } 'RepairPst' { "Repair - $Path" } default { $Mode }
    }
    Open-MbtEventLog -Path $EventsFile -JobId $JobId -Task $taskName -Title $title -Mode $Mode -Store $(if ($Mode -in 'Export', 'ExportPst', 'Import') { $StorePath })
    $steps = switch ($Mode) {
        'Export' {
            @(@{ Id = 'tree'; Title = 'Read the folder tree' }, @{ Id = 'list'; Title = 'List the items you chose' }, @{ Id = 'download'; Title = 'Download the items' }, @{ Id = 'verify'; Title = 'Check the downloads' }) +
            $(if ($OutputPst) { @(@{ Id = 'pst'; Title = 'Write the PST files' }) + $(if (-not $SkipCheck) { @(@{ Id = 'pstcheck'; Title = 'Check every PST file' }) }) })
        }
        'ExportPst' { @(@{ Id = 'pst'; Title = 'Write the PST files' }) + $(if (-not $SkipCheck) { @(@{ Id = 'pstcheck'; Title = 'Check every PST file' }) }) }
        'Import' { @(@{ Id = 'check'; Title = 'Verify the export store' }, @{ Id = 'folders'; Title = 'Create the folders' }, @{ Id = 'send'; Title = 'Send the items' }, @{ Id = 'reconcile'; Title = 'Check the uncertain items' }, @{ Id = 'final'; Title = 'Final check' }) }
        'ImportPst' { @(@{ Id = 'check'; Title = 'Check the PST files' }, @{ Id = 'folders'; Title = 'Create the folders' }, @{ Id = 'send'; Title = 'Send the items' }, @{ Id = 'reconcile'; Title = 'Check the uncertain items' }, @{ Id = 'final'; Title = 'Final check' }) }
        'ExportEml' { @(@{ Id = 'export'; Title = 'Write the .eml files' }) }
        'ImportEml' { @(@{ Id = 'send'; Title = 'Import the messages' }) }
        'CheckPst' { @(@{ Id = 'check'; Title = 'Check the file' }) }
        'RepairPst' { @(@{ Id = 'before'; Title = 'Check the file (before)' }, @{ Id = 'scan'; Title = 'Scan and read what can be read' }, @{ Id = 'write'; Title = 'Write the new file' }, @{ Id = 'after'; Title = 'Check the new file (after)' }) }
        default { @(@{ Id = 'run'; Title = $title }) }
    }
    Write-MbtJobPlan -Steps $steps
}
$storeExists = Test-Path -LiteralPath (Join-Path $StorePath 'catalog.db')
$store = $null
$exitCode = 0

# App-only: the application signs in with its certificate (no user, no prompt). Nothing here goes through /me: -UserPrincipalName is the owner of the target mailbox.
function Connect-AppOnlyForTransfer {
    if ($DeviceCode) { throw '-DeviceCode is a delegated sign-in: it cannot be combined with a certificate (app-only).' }
    if (-not $UserPrincipalName) { throw 'App-only sign-in has no signed-in user: give -UserPrincipalName, the address of the target mailbox (its owner, or the shared mailbox itself). The lab user is never assumed.' }
    $certName = if ($CertificateThumbprint) { "certificate $CertificateThumbprint" } else { "certificate file $CertificatePath" }
    Write-MbtItem INFO "signing in as the application $ClientId with its $certName (app-only; target mailbox $UserPrincipalName)"
    $session = if ($CertificateThumbprint) { Connect-MbtGraph -TenantId $TenantId -ClientId $ClientId -UserPrincipalName $UserPrincipalName -CertificateThumbprint $CertificateThumbprint -Configuration $config }
    else { Connect-MbtGraph -TenantId $TenantId -ClientId $ClientId -UserPrincipalName $UserPrincipalName -CertificatePath $CertificatePath -CertificatePassword $CertificatePassword -Configuration $config }
    $t = $session.Token
    Write-MbtItem OK "signed in as the application $($t.AppId), tenant $($t.TenantId), no user"
    Write-MbtDetail ('application permissions in the token: ' + (@($t.Roles) -join ' '))
    $kind = if ($Mailbox) { $Mailbox } else { 'Primary' }
    $missing = @(Get-MbtMissingAppPermission -Token $t -Mode $Mode -Mailbox $kind -ExchangeConversion:($MimeConversion -eq 'Exchange'))
    # A token without roles is normal when the permissions are granted only through Exchange RBAC for Applications (Exchange decides, nothing to see in the token): not an error.
    if (-not @($t.Roles).Count) { Write-MbtItem INFO 'the token holds no application permission (no roles): fine if they are granted only through Exchange RBAC for Applications (check: Test-ServicePrincipalAuthorization -Identity <application id> -Resource <mailbox>); otherwise admin consent has not been granted yet.' }
    elseif ($missing) { Write-MbtItem WARN ('missing in the token: ' + ($missing -join ', ') + ' (ignore it if Exchange RBAC for Applications grants it). Otherwise an administrator adds the application permission to the app registration and grants admin consent (tools\New-MbtAppOnlyLab.ps1 shows the steps); then sign in again.') }
    # Primary mailbox of -UserPrincipalName with no id given: found by its address like a -MailboxId address (settings/exchange needs User.Read.All, which RBAC for Applications cannot grant; the import-export API answers for a mailbox in the scope)
    if (-not $MailboxId -and -not $PrimaryMailboxId -and -not $LocationsFile -and $Mailbox -eq 'Primary') { $script:MailboxId = $UserPrincipalName }
    # -MailboxId given as an address: resolved to its physical id (User.Read.All, else the import-export API); the Exchange RBAC for Applications scope must include the mailbox
    if ($MailboxId -and (Test-MbtMailboxAddress $MailboxId)) {
        $address = $MailboxId.Trim()
        $script:MailboxId = Resolve-MbtMailboxAddress -Address $address -TenantId $TenantId
        Write-MbtItem OK "mailbox $address is $script:MailboxId"
        if (-not $mailboxGiven) { $script:Mailbox = 'Primary'; $script:mailboxGiven = $true }
    }
    return $session
}

function Connect-ForTransfer {
    if ($useCertificate) { return (Connect-AppOnlyForTransfer) }
    Assert-MbtSignInSettings -TenantId $TenantId -ClientId $ClientId
    Write-MbtItem INFO ('signing in as {0} ({1}); a line "DEVICE CODE: ..." follows only when no valid token is cached' -f $UserPrincipalName, $(if ($DeviceCode) { 'device code' } else { 'browser' }))
    try {
        $session = Connect-MbtGraph -TenantId $TenantId -ClientId $ClientId -UserPrincipalName $UserPrincipalName -DeviceCode:$DeviceCode -DeviceCodeFile (Join-Path $outRoot 'device-code.txt') -NoPrompt:$NoPrompt -Configuration $config
    } catch {
        if ($NoPrompt) { throw 'SIGNIN-REQUIRED: no valid sign-in is cached and -NoPrompt forbids asking for one. ' + $_.Exception.Message }
        throw
    }
    $t = $session.Token
    # Browse of Primary / Archive reads with the Mail API: the scope is Mail.Read. The default scope (.default) carries every permission granted to the
    # application for this user, so no other scope is asked; a token cached before the grant is renewed once with the refresh token to pick it up.
    $browseMail = ($Mode -eq 'Browse' -or $Mode -eq 'ExportEml') -and $Mailbox -ne 'AuxArchive'
    if ($browseMail -and -not ('Mail.Read' -in $t.Scopes -or 'Mail.ReadWrite' -in $t.Scopes)) {
        Write-MbtItem INFO 'Mail.Read is not in the cached token: asking for a new one with the refresh token'
        Disconnect-MbtGraph
        $session = Connect-MbtGraph -TenantId $TenantId -ClientId $ClientId -UserPrincipalName $UserPrincipalName -DeviceCode:$DeviceCode -DeviceCodeFile (Join-Path $outRoot 'device-code.txt') -ForceRefresh -NoPrompt:$NoPrompt -Configuration $config
        $t = $session.Token
    }
    Write-MbtItem OK "signed in as $($t.Account), tenant $($t.TenantId), application $($t.AppId)"
    if ($t.Account -and $t.Account -ine $UserPrincipalName) { Write-MbtItem WARN "signed in as $($t.Account), not as $UserPrincipalName" }
    Write-MbtDetail ('permissions in the token: ' + ($t.Scopes -join ' '))
    $missing = @()
    if (-not ('MailboxFolder.Read' -in $t.Scopes -or 'MailboxFolder.ReadWrite' -in $t.Scopes)) { $missing += 'MailboxFolder.Read[Write]' }
    if (-not ('MailboxItem.Read' -in $t.Scopes -or 'MailboxItem.ReadWrite' -in $t.Scopes)) { $missing += 'MailboxItem.Read[Write]' }
    if ($Mode -notin 'Browse', 'ExportEml' -and -not ('MailboxItem.ImportExport' -in $t.Scopes -or ('MailboxItem.Export' -in $t.Scopes -and $Mode -eq 'Export'))) { $missing += 'MailboxItem.ImportExport' }
    if ($browseMail -and -not ('Mail.Read' -in $t.Scopes -or 'Mail.ReadWrite' -in $t.Scopes)) { $missing += 'Mail.Read (without it the browse shows the import-export listing: no subject, sender or body)' }
    if ($Mode -in 'Import', 'ImportEml' -and -not ('MailboxFolder.ReadWrite' -in $t.Scopes)) { $missing += 'MailboxFolder.ReadWrite (creates the folders)' }
    if ($missing) { Write-MbtItem WARN ('missing in the token: ' + ($missing -join ', ') + '. Run tools\Grant-LabConsent.ps1 (as an administrator) then sign in again.') }
    # -MailboxId given as an SMTP address / UPN (shared mailbox, another user's mailbox): resolved once to its physical id, the kind is Primary unless -Mailbox says otherwise
    if ($MailboxId -and (Test-MbtMailboxAddress $MailboxId)) {
        $address = $MailboxId.Trim()
        $script:MailboxId = Resolve-MbtMailboxAddress -Address $address -TenantId $TenantId
        Write-MbtItem OK "mailbox $address is $script:MailboxId"
        if (-not $mailboxGiven) { $script:Mailbox = 'Primary'; $script:mailboxGiven = $true }
    }
    if ($MailboxId -and (Test-OtherMailbox $script:MailboxId)) {
        Write-MbtItem INFO "$script:MailboxId is not a mailbox of $UserPrincipalName (shared or another user's): $UserPrincipalName needs FullAccess on it"
        if ($Mode -eq 'Browse' -and $Mailbox -ne 'AuxArchive' -and -not ('Mail.Read.Shared' -in $t.Scopes -or 'Mail.ReadWrite.Shared' -in $t.Scopes)) {
            Write-MbtItem INFO 'Mail.Read.Shared is not in the token: the Mail API opens only the signed-in user''s own mailbox, so the browse falls back to the import-export listing (no subject, sender or body). Grant Mail.Read.Shared to the application for the full view.'
        }
    }
    return $session
}

# A location that is not one of the signed-in user's known mailboxes (the configuration, -LocationsFile, the store, the parameters): a shared mailbox or another user's.
function Test-OtherMailbox([string]$Id) {
    if (-not $Id) { return $false }
    try { $id = ConvertTo-MbtMailboxId -Value $Id -TenantId $TenantId } catch { return $false }
    $own = @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $store -LocationsFile $LocationsFile -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -SkipGraph -NoSave -Verbose:$false)
    return -not @($own | Where-Object { $_.MailboxId -ieq $id }).Count
}

# What to check when the run ends on an access denial for a mailbox that is not the signed-in user's ('' otherwise).
function Get-ErrorHint([string]$Message) {
    $current = try { Get-MbtSession } catch { $null }
    $scopes = if ($current -and $current.Token) { @($current.Token.Scopes) } else { @() }
    $which = if ($Mode -in 'Browse', 'Export', 'Import') { $Mode } else { 'Resolve' }
    return Get-MbtAccessHint -Message $Message -Mode $which -Scopes $scopes -OtherMailbox:(Test-OtherMailbox $MailboxId) -AppOnly:$useCertificate
}

# The PST files of -PstPath (a file, several files or a folder of PST files, split parts in order).
function PstFilesOf([string[]]$Paths) {
    return [MailboxTransfer.PstImportFiles]::Resolve([string[]]$Paths)
}

function Get-StoreMailboxId([string]$Which) {
    $found = @($store.GetMailboxes() | Where-Object { $_.Kind.ToString() -eq $Which })
    if ($found.Count -ne 1) { throw "The store holds $($found.Count) $Which mailbox(es): cannot choose." }
    return $found[0].Id
}

# The arguments Export-MbtPst takes from the entry script (-OutputPst and its options).
function Get-PstExportArgument {
    $a = @{ Destination = $OutputPst; MaxSizeMiB = [long]($pstMaxBytes / 1MB); SeparateAuxArchives = $SeparateAuxArchives; SkipCheck = $SkipCheck; Force = $Force; Confirm = $false; CancellationToken = (Get-MbtStopToken) }
    if ($mailboxGiven -and $Mailbox -ne 'All') { $a.Mailbox = $Mailbox }
    if ($MailboxId -and -not $SelectionFile) { $a.MailboxId = @(ConvertTo-MbtMailboxId -Value $MailboxId -TenantId $TenantId) }
    if ($SelectionFile) { $a.SelectionFile = $SelectionFile }
    return $a
}

# The kind of mailbox asked: -Mailbox, else (only -MailboxId given) the kind the store or -LocationsFile knows for that id.
function Get-AskedKind {
    if ($mailboxGiven -or -not $MailboxId) { return $Mailbox }
    $id = ConvertTo-MbtMailboxId -Value $MailboxId -TenantId $TenantId
    $locs = @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $store -LocationsFile $LocationsFile -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -SkipGraph -NoSave)
    $kind = @($locs | Where-Object { $_.MailboxId -ieq $id } | ForEach-Object Kind) | Select-Object -First 1
    if (-not $kind) { throw "-MailboxId $id is not a known location (not in the store, -LocationsFile, -PrimaryMailboxId or -ArchiveMailboxId): say what it is with -Mailbox Primary|Archive|AuxArchive." }
    return $kind
}

# The Archive / AuxArchive locations other than $ExceptId that are known (store, -LocationsFile, parameters; no Graph call), as browse locations.
function Get-OtherArchiveLocation([string]$ExceptId) {
    $readOnly = $null
    $st = $store
    if (-not $st -and $storeExists) { $readOnly = Open-MbtExportStore -Path $StorePath -ReadOnly; $st = $readOnly }
    try {
        $locs = @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $st -LocationsFile $LocationsFile -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -SkipGraph -NoSave -Verbose:$false)
    } finally { if ($readOnly) { $readOnly.Dispose() } }
    foreach ($l in $locs) {
        if ($l.Kind -notin 'Archive', 'AuxArchive' -or $l.MailboxId -ieq $ExceptId) { continue }
        New-MbtBrowseLocation -Kind $l.Kind -MailboxId $l.MailboxId -Source $l.Source
    }
}

try {
    switch ($Mode) {
        'Status' {
            Write-MbtItem OK 'the C# engine compiles'
            Write-MbtItem INFO "Graph: $($config.Graph.BaseUrl), retry up to $($config.Graph.MaxAttempts) attempts, token cache: $($config.Auth.CacheFile)"
            if (-not $storeExists) {
                Write-MbtItem INFO "no store in $StorePath (Export creates it; -StorePath says where it is)."
                break
            }
            if ($ResolveUnknown) {
                if (-not $As) { throw '-ResolveUnknown needs -As Imported or -As NotImported.' }
                if (-not $PSCmdlet.ShouldProcess("import row $ResolveUnknown of the store $StorePath", "Settle it as $As")) { Write-MbtItem WARN 'WhatIf: the row was not settled.'; break }
                $store = Open-MbtExportStore -Path $StorePath
                if ($As -eq 'Imported') { Resolve-MbtImportUnknown -Store $store -ImportKey $ResolveUnknown -Imported -TargetItemId $TargetItemId -Confirm:$false }
                else { Resolve-MbtImportUnknown -Store $store -ImportKey $ResolveUnknown -NotImported -Confirm:$false }
                Write-MbtItem OK "import row $ResolveUnknown settled: $As"
                break
            }
            if ($ResolveAbsent) {
                if (-not ($Unknown -and $Check)) { throw '-ResolveAbsent needs -Unknown -Check: it settles the rows for which the check found no matching item in the target folder.' }
                $null = Connect-ForTransfer
                $store = Open-MbtExportStore -Path $StorePath
                $settledRows = @(Resolve-MbtImportUnknown -Store $store -Absent -JobName $(if ($PSBoundParameters.ContainsKey('JobName')) { $JobName }) -WhatIf:$dryRun -Confirm:$false)
                if ($dryRun) { Write-MbtItem WARN 'WhatIf: no row was settled.' }
                elseif ($settledRows | Where-Object { -not $_.Settled }) { Write-MbtItem WARN 'rows with candidates stay Unknown: settle them one by one with -ResolveUnknown <importKey> -As Imported|NotImported [-TargetItemId <id>]' }
                break
            }
            $store = Open-MbtExportStore -Path $StorePath -ReadOnly
            if ($Unknown) {
                if ($Check) { $null = Connect-ForTransfer }
                $rows = @(Get-MbtImportUnknown -Store $store -Check:$Check)
                if (-not $rows) { Write-MbtItem OK 'no import row is Unknown' }
                foreach ($u in $rows) {
                    Write-MbtItem WARN ("import {0} (job '{1}'): item {2} '{3}' ({4}) -> folder {5}" -f $u.ImportKey, $u.JobName, $u.ItemKey, $u.Subject, (Format-MbtBytes $u.Size), $u.TargetFolder)
                    Write-MbtDetail "error: $($u.Error)"
                    if ($Check) { Write-MbtDetail $(if ($u.Candidates.Count) { "may exist in the target: $($u.Candidates -join ', ')" } else { 'no matching item found in the target folder' }) }
                }
                if ($rows) { Write-MbtDetail 'settle one with: -Mode Status -ResolveUnknown <importKey> -As Imported|NotImported [-TargetItemId <id>]' }
                if ($rows -and $Check) { Write-MbtDetail 'settle every row with no match at once with: -Mode Status -Unknown -Check -ResolveAbsent [-JobName <name>] [-WhatIf] (stop the import run first)' }
                break
            }
            if ($Locations) {
                $null = Connect-ForTransfer
                $locs = @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $store -LocationsFile $LocationsFile -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -NoSave -Verbose:$false)
                if (-not $locs) { Write-MbtItem WARN 'no location known: give -LocationsFile (Exchange Online Get-MailboxLocation output), or sign in with User.Read.All' }
                foreach ($l in $locs) { Write-MbtItem INFO ("{0,-10} {1}  [{2}{3}]" -f $l.Kind, $l.MailboxId, $l.Source, $(if ($l.Learned) { ', learned from a redirect' } else { '' })) }
                break
            }
            Show-MbtStoreStatus -Status (Get-MbtStoreStatus -Store $store)
        }
        'Export' {
            $null = Connect-ForTransfer
            $plan = $null
            if ($SelectionFile) {
                # what the browse view selected: the mailbox is the one of the file; the entries are checked against the import-export API first
                if ($Folder) { throw '-SelectionFile and -Folder are exclusive: the selection names its folders.' }
                $selection = Read-MbtBrowseSelection -Path $SelectionFile
                $selectedId = ConvertTo-MbtMailboxId -Value $selection.MailboxId -TenantId $TenantId
                if ($mailboxGiven -and $Mailbox -ne $selection.Kind) { throw "-Mailbox $Mailbox does not match the selection, which is of the $($selection.Kind) mailbox." }
                if ($MailboxId -and (ConvertTo-MbtMailboxId -Value $MailboxId -TenantId $TenantId) -ine $selectedId) { throw "-MailboxId $MailboxId is not the mailbox of the selection ($selectedId)." }
                $Mailbox = $selection.Kind; $MailboxId = $selectedId; $mailboxGiven = $true
                Write-MbtItem INFO ("selection {0}: {1} entr{2} for the {3} mailbox {4}" -f $SelectionFile, $selection.Entries.Count, $(if ($selection.Entries.Count -eq 1) { 'y' } else { 'ies' }), $selection.Kind, $selectedId)
                # a selection made on an archive still resolves when a folder moved to an auxiliary archive since: the other archives are searched by folder entry key
                $also = if ($selection.Kind -in 'Archive', 'AuxArchive') { @(Get-OtherArchiveLocation $selectedId) } else { @() }
                $plan = Resolve-MbtBrowsePlan -Selection $selection -AlsoSearch $also -RecoverableItemsFolders $RecoverableItemsFolders
                Show-MbtExportPlan -Plan $plan
                if (-not $plan.IsComplete) { throw "-SelectionFile: $($plan.Problems.Count) problem(s), nothing was exported. Remove those entries from the selection (or select them again) and run again." }
                if ($dryRun) {
                    # the counts of the plan (the service now, folders below included), not the ones the browse view wrote in the file
                    $selTotals = $plan.Totals()
                    Write-MbtItem INFO ("selection: {0}{1}" -f (Format-MbtBrowseTotals $selTotals), $(if (-not $selTotals.SizeUnknown) { "; about {0} PST file(s) of at most {1} if converted (-OutputPst)" -f [math]::Max(1, [math]::Ceiling($selTotals.KnownSizeBytes / $pstMaxBytes)), (Format-MbtBytes $pstMaxBytes) }))
                    Write-MbtItem WARN 'WhatIf: the selection is valid, nothing was exported and no store was created.'
                    break
                }
            }
            if ($dryRun) {
                # -WhatIf: the plan and nothing else. No store is created and no file written; the sign-in and the reading of the folder trees are the only requests (read only).
                if ($Mailbox -eq 'All' -and $MailboxId) { throw '-MailboxId is one location: it cannot be combined with -Mailbox All.' }
                if ($storeExists) { $store = Open-MbtExportStore -Path $StorePath -ReadOnly }
                $asked = Get-AskedKind
                $targets = [System.Collections.Generic.List[object]]::new()
                foreach ($kind in $(if ($asked -eq 'All') { @('Primary', 'Archive') } else { @($asked) })) {
                    try {
                        $id = Resolve-MbtMailboxId -Mailbox $kind -TenantId $TenantId -UserPrincipalName $UserPrincipalName -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId `
                            -MailboxId $(if ($asked -ne 'All') { $MailboxId }) -Store $store -LocationsFile $LocationsFile
                        $targets.Add([pscustomobject]@{ Kind = $kind; Id = $id })
                    } catch {
                        if ($asked -ne 'All') { throw }
                        Write-MbtItem WARN "the $kind mailbox cannot be planned: $($_.Exception.Message)"
                    }
                }
                if ($asked -eq 'All') {
                    foreach ($a in @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $store -LocationsFile $LocationsFile -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -SkipGraph -NoSave -Verbose:$false | Where-Object Kind -EQ 'AuxArchive')) {
                        $targets.Add([pscustomobject]@{ Kind = 'AuxArchive'; Id = $a.MailboxId })
                    }
                }
                Write-MbtItem INFO ("plan: {0} location(s) to export into {1}" -f $targets.Count, $StorePath)
                $sumItems = [long]0; $sumBytes = [long]0; $sumUnknown = $false
                foreach ($tg in $targets) {
                    try {
                        $withRecoverable = [bool]($IncludeRecoverableItems -or ($Folder -and [MailboxTransfer.RecoverableItems]::IsPath($Folder)))
                        $recoverableNames = if ($IncludeRecoverableItems -or ($Folder -and $Folder.Trim('/') -ieq [MailboxTransfer.RecoverableItems]::TopName)) { [MailboxTransfer.RecoverableItems]::Chosen([string[]]@($RecoverableItemsFolders | Where-Object { $_ })) }
                        $tree = Get-MbtBrowseFolder -Location (New-MbtBrowseLocation -Kind $tg.Kind -MailboxId $tg.Id) -AsTree -RecoverableItems:$withRecoverable -RecoverableItemsFolders $recoverableNames
                        $totals = Show-MbtLiveExportPlan -Tree $tree -Kind $tg.Kind -Folder $Folder -PstMaxBytes $pstMaxBytes -IncludeRecoverableItems:$withRecoverable
                        $sumItems += $totals.Items; $sumBytes += $totals.KnownSizeBytes; if (-not $totals.SizeComplete) { $sumUnknown = $true }
                    } catch { Write-MbtItem WARN "the folders of the $($tg.Kind) mailbox $($tg.Id) could not be read: $($_.Exception.Message)" }
                }
                if ($asked -in 'Archive', 'All') { Write-MbtDetail 'an auto-expanding archive: the items of its expanded folders live in auxiliary archives; the export follows them and records them under the folder of the main archive' }
                Write-MbtItem INFO ("total: {0} item(s), {1}{2}; about {3} PST file(s) of at most {4} if converted ({5})" -f $sumItems, (Format-MbtBytes $sumBytes), $(if ($sumUnknown) { ' announced (some folders give no size)' } else { ' announced' }),
                    [math]::Max(1, [math]::Ceiling($sumBytes / $pstMaxBytes)), (Format-MbtBytes $pstMaxBytes), $(if ($OutputPst) { $OutputPst } else { 'give -OutputPst to convert' }))
                if ($OutputPst) { Write-MbtDetail "after the export the store would be converted: $OutputPst, then -002.pst... and a manifest" }
                if (-not $PSCmdlet.ShouldProcess($StorePath, "Export $($targets.Count) location(s) into the store")) { Write-MbtItem WARN 'WhatIf: nothing was exported and no store was created.' }
                break
            }
            $store = Open-MbtExportStore -Path $StorePath -Name 'export' -Source "$UserPrincipalName ($TenantId)"
            if ($Mailbox -eq 'All' -and $MailboxId) { throw '-MailboxId is one location: it cannot be combined with -Mailbox All.' }
            $asked = Get-AskedKind
            # the locations given up front (operator file, parameters, Graph) are recorded in the store, with how they were learned
            $known = @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $store -LocationsFile $LocationsFile -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -SkipGraph -Verbose:$false)
            $auxKnown = @($known | Where-Object Kind -EQ 'AuxArchive')
            if ($auxKnown) { Write-MbtItem INFO ("{0} auxiliary archive(s) known: {1}" -f $auxKnown.Count, (($auxKnown | ForEach-Object { "$($_.MailboxId) [$($_.Source)]" }) -join ', ')) }
            $kinds = if ($asked -eq 'All') { @('Primary', 'Archive') } else { @($asked) }
            $failures = 0
            $exported = @{}
            $script:resumable = $false
            function Invoke-OneExport([string]$Kind, [string]$Id) {
                # a plan names folders of several mailboxes when a folder moved to an auxiliary archive: each mailbox exports its own part
                $subPlan = if ($plan) { $plan.ForMailbox($Id) }
                if ($plan -and $subPlan.Folders.Count -eq 0) { Write-MbtItem INFO "nothing of the selection is in the $Kind mailbox $Id"; return }
                Write-MbtItem INFO "exporting the $Kind mailbox $Id into $StorePath"
                $exported[$Id.ToLowerInvariant()] = $true
                $r = Export-MbtMailbox -Store $store -MailboxId $Id -Kind $Kind -UserPrincipalName $UserPrincipalName -Concurrency $Concurrency `
                    -BatchItems $BatchItems -BatchBytes ([long]$BatchMiB * 1MB) -MaxAttempts $MaxAttempts -NotReadyTrip $NotReadyTrip -SkipInventory:$SkipInventory -Folder $Folder -Plan $subPlan -IncludeRecoverableItems:$IncludeRecoverableItems -RecoverableItemsFolders $RecoverableItemsFolders -TempPath $TempPath -IntervalSeconds $IntervalSeconds
                if ($r.Error) { throw $r.Error.Message }
                $script:counts['done'] = [long]$script:counts['done'] + [long]$r.Result.Exported
                $script:counts['failed'] = [long]$script:counts['failed'] + [long]$r.Result.Failed
                $script:counts['deferred'] = [long]$script:counts['deferred'] + [long]$r.Result.FoldersDeferred.Count
                $script:counts['bytes'] = [long]$script:counts['bytes'] + [long]$r.Result.Bytes
                foreach ($d in $r.Result.FoldersDeferred) { Write-MbtJobObject -Step download -Key ([string]$d.Path) -State notready -Msg ([string]$d.ToString()) }
                if ($r.Result.NeedsAnotherRun -or $r.Job.Cancelled) { $script:resumable = $true }
            }
            foreach ($kind in $kinds) {
                try {
                    $id = Resolve-MbtMailboxId -Mailbox $kind -TenantId $TenantId -UserPrincipalName $UserPrincipalName -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId `
                        -MailboxId $(if ($asked -ne 'All') { $MailboxId }) -Store $store -LocationsFile $LocationsFile
                    Invoke-OneExport $kind $id
                } catch {
                    if ($asked -ne 'All') { throw }
                    $failures++
                    Write-MbtItem FAIL "export of the $kind mailbox failed: $($_.Exception.Message)"
                }
            }
            if ($asked -eq 'All') {
                # the auxiliary archives learned while the main archive was exported are exported by themselves too (their own tree; folders the
                # main archive already records are not recorded twice); one that cannot be listed is reported and does not stop the others
                $aux = @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $store -SkipGraph -NoSave -Verbose:$false | Where-Object Kind -EQ 'AuxArchive')
                if (-not $aux) { Write-MbtItem INFO 'no auxiliary archive is known: the archive is not auto-expanded, or none of its folders has been expanded yet' }
                foreach ($a in $aux) {
                    if ($exported.ContainsKey($a.MailboxId.ToLowerInvariant())) { continue }
                    try { Invoke-OneExport 'AuxArchive' $a.MailboxId }
                    catch { $failures++; Write-MbtItem FAIL "export of the auxiliary archive $($a.MailboxId) failed: $($_.Exception.Message)" }
                }
            }
            if ($plan) {
                # the other mailboxes of the plan (folders and items that moved to an auxiliary archive)
                foreach ($m in @($plan.Mailboxes)) {
                    if ($exported.ContainsKey($m.ToLowerInvariant())) { continue }
                    $otherKind = @(Get-OtherArchiveLocation '' | Where-Object { $_.MailboxId -ieq $m } | ForEach-Object Kind) | Select-Object -First 1
                    Invoke-OneExport $(if ($otherKind) { $otherKind } else { 'AuxArchive' }) $m
                }
            }
            if ($failures) { $exitCode = 1; Write-MbtItem WARN "$failures export(s) of -Mailbox All failed (the others are done)" }
            elseif ($script:resumable) {
                $exitCode = 2
                Write-MbtItem WARN 'exit code 2: the export is not complete (deferred folder(s), folder(s) not listed, failed item(s) or an interrupted run): nothing was lost, run the same command again (later for a deferred folder); it does only what is left'
            }
            if ($OutputPst) {
                if ($exitCode -eq 1) { Write-MbtItem WARN '-OutputPst: the export failed, the store was not converted' }
                elseif (Test-MbtStopRequested) { Write-MbtItem WARN '-OutputPst: stopped on request, the store was not converted: run the same command again to continue' }
                else {
                    Write-MbtStep "Export to PST: $OutputPst"
                    $pstArgs = Get-PstExportArgument
                    $pst = Export-MbtPst -Store $store @pstArgs
                    if ($pst.Result) { $script:counts.parts = [int]$pst.Result.Parts.Count; $script:counts.itemsWritten = [long]$pst.Result.ItemsWritten }
                    if ($pst.ExitCode -eq 1) { $exitCode = 1 } elseif ($pst.ExitCode -eq 2 -and $exitCode -eq 0) { $exitCode = 2 }
                }
            }
        }
        'ExportPst' {
            if ($PstPath) {
                # a PST instead of a store: chosen folders and messages are copied into new PST files (PST to PST, our own writer, the source is only read)
                if (-not $OutputPst) { throw '-Mode ExportPst needs -OutputPst <file.pst or folder>.' }
                $source = @(PstFilesOf $PstPath)
                $target = if (Test-Path -LiteralPath $OutputPst -PathType Container) { Join-Path $OutputPst ('{0}-copy.pst' -f [IO.Path]::GetFileNameWithoutExtension($source[0])) } else { $OutputPst }
                $nids = @()
                if ($IncludeItemFile) {
                    if (-not (Test-Path -LiteralPath $IncludeItemFile -PathType Leaf)) { throw "-IncludeItemFile not found: $IncludeItemFile" }
                    $nids = @(Get-Content -LiteralPath $IncludeItemFile | Where-Object { $_.Trim() } | ForEach-Object { [uint32]($_.Trim()) })
                }
                Write-MbtItem INFO ("copying {0} into PST files at {1} (the source is only read)" -f ($source -join ', '), $target)
                $copyArgs = @{ Path = $source; Destination = $target; MaxSizeMiB = [long]($pstMaxBytes / 1MB); SkipCheck = $SkipCheck; Force = $Force; CancellationToken = (Get-MbtStopToken) }
                if ($IncludeFolder) { $copyArgs.IncludeFolder = $IncludeFolder }
                if ($ExcludeFolder) { $copyArgs.ExcludeFolder = $ExcludeFolder }
                if ($nids.Count) { $copyArgs.ItemNid = [uint32[]]$nids }
                $pst = Copy-MbtPst @copyArgs -WhatIf:$dryRun -Confirm:$false
                if ($dryRun) { Write-MbtItem WARN 'WhatIf: no PST was written.'; break }
                $exitCode = $pst.ExitCode
                if ($pst.Result) {
                    $script:counts = @{ parts = [int]$pst.Result.Parts.Count; itemsWritten = [long]$pst.Result.ItemsWritten; failed = [int]$pst.Result.Failed.Count; bytes = [long]$pst.Result.BytesWritten }
                    foreach ($f in @($pst.Result.Failed | Select-Object -First 200)) { Write-MbtEvent item ([ordered]@{ step = 'pst'; key = $(if ($f.Subject) { [string]$f.Subject } else { '(no subject)' }); state = 'refused'; msg = [string]$f.Reason; detail = [string]$f.Folder }) }
                    # 1 also when the copy is finished and some messages could not be copied (nothing to run again): the same reason as an import with refused items
                    if ($exitCode -eq 1 -and -not $pst.Result.Error -and -not $pst.Result.PartsWithErrors -and $pst.Result.Failed.Count) { $script:reason = 'refused-items'; $script:next = 'done'; $script:resultComplete = $true }
                }
                break
            }
            if (-not $storeExists) { throw "No store in ${StorePath}: export first (-Mode Export) or give -StorePath." }
            if (-not $OutputPst) { throw '-Mode ExportPst needs -OutputPst <file.pst or folder>.' }
            Write-MbtItem INFO "converting the store $StorePath into PST files (the store is only read)"
            $pstArgs = Get-PstExportArgument
            $pst = Export-MbtPst -Store $StorePath @pstArgs -WhatIf:$dryRun
            if ($dryRun) { Write-MbtItem WARN 'WhatIf: no PST and no manifest was written.' }
            else {
                $exitCode = $pst.ExitCode
                if ($pst.Result) { $script:counts.parts = [int]$pst.Result.Parts.Count; $script:counts.itemsWritten = [long]$pst.Result.ItemsWritten; $script:counts.failed = [int]$pst.Result.Failed.Count }
            }
        }
        'Browse' {
            if ($Mailbox -eq 'All') { throw '-Mailbox All is for Export. Browse one mailbox: Primary, Archive or AuxArchive (-MailboxId).' }
            if ($SaveSelection -and -not $Folder) { throw '-SaveSelection needs -Folder (the folder to select; with -ItemId, the items of it).' }
            if ($ItemId -and -not $Folder) { throw '-ItemId needs -Folder.' }
            $null = Connect-ForTransfer
            if ($storeExists) { $store = Open-MbtExportStore -Path $StorePath -ReadOnly }
            $kind = Get-AskedKind
            $id = Resolve-MbtMailboxId -Mailbox $kind -TenantId $TenantId -UserPrincipalName $UserPrincipalName -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId `
                -MailboxId $MailboxId -Store $store -LocationsFile $LocationsFile
            $loc = New-MbtBrowseLocation -Kind $kind -MailboxId $id
            Write-MbtItem INFO ("browsing the {0} mailbox {1} ({2})" -f $kind, $id, $(if ($loc.HasMailApi) { 'Microsoft Graph Mail API' } else { 'import-export API only: the Mail API does not reach an auxiliary archive' }))
            $watch = [Diagnostics.Stopwatch]::StartNew()
            if (-not $Folder) {
                $tree = Get-MbtBrowseFolder -Location $loc -AsTree -RecoverableItems:($RecoverableItems -or $IncludeRecoverableItems) -RecoverableItemsFolders $RecoverableItemsFolders
                # a folder in the main archive and in an auxiliary archive: where its items are (the other side's tree is read for it)
                if ($kind -in 'Archive', 'AuxArchive' -and ($ShowMoves -or $kind -eq 'Archive')) {
                    try {
                        $sides = @(Get-OtherArchiveLocation $id | Where-Object { ($kind -eq 'Archive' -and $_.Kind -eq 'AuxArchive') -or ($kind -eq 'AuxArchive' -and $_.Kind -eq 'Archive') })
                        $trees = @($tree) + @($sides | ForEach-Object { Get-MbtBrowseFolder -Location $_ -AsTree })
                        $groups = @(Add-MbtBrowseMoveInfo -Tree $trees)
                        $inMove = @($groups | Where-Object { $_.State -ne 'None' })
                        if ($sides.Count -eq 0 -and $ShowMoves) { Write-MbtItem WARN 'no other archive location is known to compare with (give -LocationsFile, or run an export first)' }
                        elseif ($sides.Count -gt 0) { Write-MbtItem INFO ("{0} folder(s) exist in the main archive and in an auxiliary archive, {1} of them with a move in progress or done" -f $groups.Count, $inMove.Count) }
                    } catch { Write-MbtItem WARN "the other archive could not be compared: $($_.Exception.Message)" }
                }
                Show-MbtBrowseTree -Tree $tree
                Write-MbtDetail ("tree read in {0}. List a folder with -Folder <path>; preview an item with -Folder <path> -ItemId <row number>." -f (Format-MbtDuration $watch.Elapsed.TotalSeconds))
                break
            }
            # the folder is found by walking down its path (one request a level), not by reading the whole tree
            $folderObj = Find-MbtBrowseFolder -Location $loc -Path $Folder -RecoverableItems:($RecoverableItems -or $IncludeRecoverableItems -or [MailboxTransfer.RecoverableItems]::IsPath($Folder)) -RecoverableItemsFolders $RecoverableItemsFolders
            Write-MbtItem INFO ("folder {0}: {1} item(s){2}{3}" -f $folderObj.Path, $folderObj.TotalItemCount, $(if ($folderObj.SizeBytes -ge 0) { ", $(Format-MbtBytes $folderObj.SizeBytes)" }), $(if ($folderObj.ChildFolderCount) { ", $($folderObj.ChildFolderCount) subfolder(s)" }))
            if ($SaveSelection -and -not $ItemId) {
                $selection = New-MbtBrowseSelection -Location $loc
                $null = Add-MbtBrowseSelection -Selection $selection -Folder $folderObj -IncludeSubfolders:$IncludeSubfolders
            } else {
                $rows = @(Get-MbtBrowseItem -Location $loc -Folder $folderObj -Top $Top -Sort $Sort -NoRecipients:$NoRecipients)
                if ($folderObj.State -eq 'ContentNotAvailable') { Write-MbtItem WARN ("folder {0}: content not available yet ({1}); the folder is not empty or lost, ask again later" -f $folderObj.Path, $folderObj.StateMessage) }
                if (-not $ItemId) {
                    $rows | Show-MbtBrowseItems
                    if ($folderObj.TotalItemCount -gt $rows.Count -and $folderObj.State -ne 'ContentNotAvailable') { Write-MbtDetail ("{0} of {1} item(s) shown (-Top 0 for all of them)" -f $rows.Count, $folderObj.TotalItemCount) }
                    break
                }
                # -ItemId: a row number of this listing (same -Top and -Sort), or an item id
                $picked = @(foreach ($key in $ItemId) {
                    if ($key -match '^\d{1,6}$') {
                        if ([int]$key -lt 1 -or [int]$key -gt $rows.Count) { throw "-ItemId ${key}: the listing has $($rows.Count) row(s) (-Top $Top)." }
                        $rows[[int]$key - 1]
                    } else {
                        $r = @($rows | Where-Object { $_.Id -ceq $key })
                        if ($r.Count -ne 1) { throw "-ItemId '$key' is not an item of the listing of $($folderObj.Path) (-Top $Top)." }
                        $r[0]
                    }
                })
                if ($SaveSelection) {
                    $selection = New-MbtBrowseSelection -Location $loc
                    $null = Add-MbtBrowseSelection -Selection $selection -Folder $folderObj -Item $picked
                }
            }
            if ($SaveSelection) {
                $written = $null
                if ($PSCmdlet.ShouldProcess($SaveSelection, 'Write the selection file')) {
                    $written = Save-MbtBrowseSelection -Selection $selection -Path $SaveSelection
                    Write-MbtItem OK "selection written to $written"
                } else { Write-MbtItem WARN 'WhatIf: the selection file was not written.' }
                $plan = Resolve-MbtBrowsePlan -Selection $selection -RecoverableItemsFolders $RecoverableItemsFolders
                Show-MbtExportPlan -Plan $plan
                if ($written) { Write-MbtDetail "export it with: -Mode Export -SelectionFile $written" }
                if (-not $plan.IsComplete) { $exitCode = 1 }
                break
            }
            foreach ($row in $picked) {
                Write-MbtItem INFO ('preview of row {0}' -f ([array]::IndexOf($rows, $row) + 1))
                Get-MbtBrowsePreview -Location $loc -Item $row -BodyChars $BodyChars | Show-MbtBrowsePreview
            }
        }
        'ImportPst' {
            if (-not $PstPath) { throw '-PstPath is required: a PST file, several files or a folder of PST files.' }
            if ($Mailbox -eq 'All') { throw '-Mailbox All is for Export. An import goes into one mailbox: Primary, Archive or AuxArchive (-MailboxId).' }
            if (-not $TargetRoot -and -not $MergeWellKnown) { throw '-TargetRoot is required: the folder below which the folders of the PST are created (for instance Imported/2019), or -MergeWellKnown.' }
            $null = Connect-ForTransfer
            $targetKind = Get-AskedKind
            $pstJob = if ($PSBoundParameters.ContainsKey('JobName')) { $JobName } else { 'pstimport' }
            $r = Import-MbtPst -Path $PstPath -Mailbox $targetKind -MailboxId $MailboxId -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -LocationsFile $LocationsFile `
                -TargetFolder $TargetRoot -FolderMap $MappingFile -MergeWellKnown:$MergeWellKnown -PerFileSubfolder:$PerFileSubfolder -NoCheck:$NoCheck -Repair:$Repair -RepairLevel $RepairLevel `
                -IncludeFolder $IncludeFolder -ExcludeFolder $ExcludeFolder -KeepDamagedAttachments:$KeepDamagedAttachments -JournalPath (Join-Path $TempPath 'pstimport.db') -WorkPath (Join-Path $TempPath 'pstimport-work') -JobName $pstJob `
                -Concurrency $Concurrency -MaxAttempts $MaxAttempts -NoAutoReconcile:$NoAutoReconcile -ReconcileDelaySeconds $ReconcileDelaySeconds -ReconcileRounds $ReconcileRounds `
                -IntervalSeconds $IntervalSeconds -WhatIf:$dryRun -Confirm:$false
            if ($dryRun) { Write-MbtItem WARN 'WhatIf: nothing was created and nothing was sent.' }
            elseif ($r.Error) { throw $r.Error.Message }
            else {
                $jr = $r.Result
                $script:counts = @{ done = [long]$jr.JobDone; failed = [long]$jr.JobFailed; unknown = [long]$jr.JobUnknown; pending = [long]$jr.JobPending; refused = [long]$jr.JobRefused; total = [long]$jr.JobTotal; foldersCreated = [long]$jr.FoldersCreated }
                if ($r.ReportPath) { $script:counts.report = $r.ReportPath }
                # 0 every row Done, 3 Unknown rows, 2 Pending / Failed rows (run again), 1 finished but some items are refused for good (nothing to run again: the report lists them)
                $exitCode = $jr.ExitCode
                if ($exitCode -ne 0) { Write-MbtItem WARN ("exit code ${exitCode}: " + $jr.JobSummary()) }
                if ($exitCode -eq 1) { $script:reason = 'refused-items'; $script:next = 'done'; $script:resultComplete = $true }
                # one item event per refused item (the first 200; the CSV has them all): what the Done page lists
                foreach ($i in @($jr.NotImported | Where-Object { $_.State -eq 'Refused' } | Select-Object -First 200)) {
                    Write-MbtEvent item ([ordered]@{ step = 'send'; key = $(if ($i.Subject) { [string]$i.Subject } else { '(no subject)' }); state = 'refused'; msg = [string]$i.Reason; detail = [string]$i.Folder })
                }
                Write-MbtJobStep -Id final -State $(if ($jr.JobComplete) { 'done' } else { 'running' }) -Done $jr.JobDone -Total $jr.JobTotal -Msg $jr.JobSummary()
            }
        }
        'ExportEml' {
            if (-not $EmlPath) { throw '-EmlPath (the folder the .eml files are written to) is required.' }
            if ($Mailbox -eq 'All') { throw '-Mailbox All is not available for ExportEml: export one mailbox at a time.' }
            $emlSource = if ($EmlSource) { $EmlSource } elseif ($PstPath) { 'Pst' } elseif ($PSBoundParameters.ContainsKey('StorePath')) { 'Store' } else { 'Live' }
            $emlArgs = @{ OutputPath = $EmlPath; Concurrency = $Concurrency; NoResume = $NoResume; NoTransportHeaders = $NoTransportHeaders; WhatIf = $dryRun; Confirm = $false; CancellationToken = (Get-MbtStopToken) }
            if ($Folder) { $emlArgs.Folder = $Folder }
            if ($emlSource -eq 'Store' -and $MailboxId) { $emlArgs.MailboxId = $MailboxId }
            switch ($emlSource) {
                'Pst' {
                    if (-not $PstPath -or @($PstPath).Count -ne 1) { throw '-EmlSource Pst needs exactly one -PstPath.' }
                    $r = Export-MbtEml @emlArgs -PstPath ([string]@($PstPath)[0])
                }
                'Store' {
                    if (-not $storeExists) { throw "No store in ${StorePath}." }
                    $store = Open-MbtExportStore -Path $StorePath -ReadOnly
                    $r = Export-MbtEml @emlArgs -Store $store
                }
                'Live' {
                    $null = Connect-ForTransfer
                    $kind = Get-AskedKind
                    $id = Resolve-MbtMailboxId -Mailbox $kind -TenantId $TenantId -UserPrincipalName $UserPrincipalName -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -MailboxId $MailboxId -LocationsFile $LocationsFile
                    $loc = New-MbtBrowseLocation -Kind $kind -MailboxId $id
                    $r = Export-MbtEml @emlArgs -Location $loc
                }
            }
            Write-MbtItem OK ("exported {0}, already there {1}, skipped {2}, failed {3}; {4}; manifest {5}" -f $r.Exported, $r.Resumed, $r.Skipped, $r.Failed, (Format-MbtBytes $r.Bytes), $r.ManifestPath)
            foreach ($n in $r.FolderNotes) { Write-MbtItem WARN "folder not exported: $n" }
            foreach ($p in @($r.Problems | Select-Object -First 20)) { Write-MbtItem WARN ("{0}: {1} - {2} ({3})" -f $p.State, $p.Folder, $p.Subject, $p.Reason) }
            $script:counts = @{ done = [long]$r.Exported; resumed = [long]$r.Resumed; skipped = [long]$r.Skipped; failed = [long]$r.Failed; bytes = [long]$r.Bytes; folders = [long]$r.Folders }
            if ($r.Failed -gt 0 -or $r.Cancelled) { $exitCode = 2 }
            Write-MbtJobStep -Id export -State $(if ($exitCode -eq 0) { 'done' } else { 'running' }) -Done ($r.Exported + $r.Resumed + $r.Skipped + $r.Failed) -Bytes $r.Bytes -Msg $(if ($r.Cancelled) { 'stopped' })
        }
        'ImportEml' {
            if (-not $EmlPath) { throw '-EmlPath (a folder of .eml files, or one .eml file) is required.' }
            if (-not $TargetRoot -and -not $ResolveEml) { throw '-TargetRoot (the top-level folder everything is created below) is required.' }
            if ($Mailbox -eq 'All') { throw '-Mailbox All is for Export. An import goes into one mailbox: Primary, Archive or AuxArchive (-MailboxId).' }
            if ($ResolveEml) {
                # the files an earlier run left Unknown, after a look in the mailbox (no sign-in needed)
                $settled = @(Resolve-MbtEmlUnknown -Path $EmlPath -As $ResolveEml -File $EmlFile -Confirm:$false)
                Write-MbtItem OK ("{0} file(s) settled as {1}: {2}" -f $settled.Count, $ResolveEml, ($settled -join ', '))
                break
            }
            $null = Connect-ForTransfer
            $targetKind = Get-AskedKind
            $targetId = Resolve-MbtMailboxId -Mailbox $targetKind -TenantId $TenantId -UserPrincipalName $UserPrincipalName -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -MailboxId $MailboxId -LocationsFile $LocationsFile
            $r = Import-MbtEml -Path $EmlPath -MailboxId $targetId -TargetRoot $TargetRoot -MimeConversion $MimeConversion -Unread:$Unread -NoResume:$NoResume -IncludeFolder $IncludeFolder -ExcludeFolder $ExcludeFolder -CancellationToken (Get-MbtStopToken) -WhatIf:$dryRun -Confirm:$false
            Write-MbtItem OK ("imported {0}, already done {1}, failed {2}, unknown {3}; {4} folder(s) created; state {5}" -f $r.Done, $r.Resumed, $r.Failed, $r.Unknown, $r.FoldersCreated, $r.StatePath)
            if ($r.Filtered) { Write-MbtItem INFO ("{0} file(s) left out: their folder is not chosen (-IncludeFolder / -ExcludeFolder)" -f $r.Filtered) }
            foreach ($u in @($r.UnmatchedPatterns)) { Write-MbtItem WARN "$u matches no folder of $EmlPath" }
            if (($IncludeFolder -or $ExcludeFolder) -and $r.Done + $r.Resumed + $r.Failed + $r.Unknown -eq 0) { throw "No .eml file is in a chosen folder (-IncludeFolder / -ExcludeFolder): nothing was imported. The paths are relative to $EmlPath, for instance Inbox/Projects." }
            foreach ($e in @($r.Entries | Where-Object { $_.State -ne 'Done' } | Select-Object -First 20)) { Write-MbtItem WARN ("{0}: {1} - {2}" -f $e.State, $e.File, $e.Error) }
            $script:counts = @{ done = [long]$r.Done; resumed = [long]$r.Resumed; failed = [long]$r.Failed; unknown = [long]$r.Unknown; foldersCreated = [long]$r.FoldersCreated }
            if ($r.Unknown -gt 0) { $exitCode = 3 } elseif ($r.Failed -gt 0 -or $r.Cancelled) { $exitCode = 2 }
            Write-MbtJobStep -Id send -State $(if ($exitCode -eq 0) { 'done' } else { 'running' }) -Done ($r.Done + $r.Resumed + $r.Failed + $r.Unknown) -Msg $(if ($r.Cancelled) { 'stopped' })
        }
        'CheckPst' {
            if (-not $Path) { throw '-Mode CheckPst needs -Path <file.pst>.' }
            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "PST file not found: $Path" }
            Write-MbtItem INFO "checking $Path (read only; the file is never changed)"
            if ($dryRun) { Write-MbtItem WARN 'WhatIf: the file would be read and checked; nothing is written.'; break }
            Write-MbtJobStep -Id check -State running
            $chk = Test-MbtPst -Path $Path -CancellationToken (Get-MbtStopToken)
            Write-Host $chk.Report.Summary()
            foreach ($f in $chk.Findings | Select-Object -First 100) { Write-MbtDetail $f.ToString() }
            $script:counts = @{ errors = [int]$chk.Errors; warnings = [int]$chk.Warnings; infos = [int]$chk.Infos; fileBytes = [long]$chk.Report.FileLength }
            # one item event per finding code: the first finding of the code and how many there are
            $seen = @{}
            foreach ($f in $chk.Findings) {
                if ($seen.ContainsKey($f.Code)) { continue }
                $seen[$f.Code] = $true
                if ($seen.Count -le 200) { Write-MbtEvent item ([ordered]@{ step = 'check'; key = [string]$f.Code; state = ([string]$f.Severity).ToLowerInvariant(); msg = [string]$f.Message; detail = [string]$f.Where; count = [int]$chk.Report.CountByCode[$f.Code]; layer = [string]$f.Layer }) }
            }
            if ($chk.Cancelled) {
                $exitCode = 2; $script:reason = 'stopped'
                Write-MbtJobStep -Id check -State running -Msg 'stopped'
                Write-MbtItem WARN 'check stopped on request: the findings so far are listed'
            } else {
                $exitCode = $chk.ExitCode
                $script:next = if ($exitCode -eq 0) { 'done' } else { 'repair' }
                $script:reason = switch ($exitCode) { 0 { 'clean' } 2 { 'warnings' } default { 'errors' } }
                $script:stateOverride = 'complete'
                Write-MbtJobStep -Id check -State done -Msg $chk.Report.Summary().Split("`n")[-1].Trim()
            }
        }
        'RepairPst' {
            if (-not $Path -or -not $Destination) { throw '-Mode RepairPst needs -Path <file.pst> and -Destination <new file.pst>.' }
            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "PST file not found: $Path" }
            Write-MbtItem INFO "repairing $Path into a new file $Destination (level $Level; the original is never changed)"
            $rp = Repair-MbtPst -Path $Path -Destination $Destination -Level $Level -KeepDamagedAttachments:$KeepDamagedAttachments -CancellationToken (Get-MbtStopToken) -WhatIf:$dryRun -Confirm:$false
            if ($dryRun) { Write-MbtItem WARN 'WhatIf: nothing was written.'; break }
            if (-not $rp) { throw 'The repair did not run.' }
            if ($rp.Error) { throw $rp.Error }
            $script:counts = [ordered]@{ recovered = [long]$rp.ItemsRecovered; partial = [long]$rp.ItemsPartial; lost = [long]$rp.ItemsLost; orphan = [long]$rp.OrphanItems; levelUsed = [string]$rp.LevelUsed; recommended = [string]$rp.Recommended; destination = [string]$rp.Destination }
            if ($rp.Before) { $script:counts.before = @{ errors = [int]$rp.Before.Errors; warnings = [int]$rp.Before.Warnings } }
            if ($rp.After) { $script:counts.after = @{ errors = [int]$rp.After.Errors; warnings = [int]$rp.After.Warnings } }
            foreach ($i in @($rp.Items | Where-Object { [string]$_.Status -ne 'Recovered' } | Select-Object -First 200)) {
                Write-MbtEvent item ([ordered]@{ step = 'write'; key = [string]$i.Subject; state = ([string]$i.Status).ToLowerInvariant(); msg = ($i.Lost -join '; '); detail = [string]$i.SourceFolder })
            }
            if ($rp.Cancelled) {
                $exitCode = 2; $script:reason = 'stopped'
                Write-MbtItem WARN 'repair stopped on request: no file was written'
            } else {
                Write-MbtJobStep -Id after -State done
                if (-not $rp.Success) { $exitCode = 2; $script:reason = 'errors-remain'; $script:next = 'done'; Write-MbtItem WARN 'the new file still has errors (see the report above)' }
            }
        }
        'Import' {
            if (-not $storeExists) { throw "No store in ${StorePath}: export first (-Mode Export) or give -StorePath." }
            if ($Mailbox -eq 'All') { throw '-Mailbox All is for Export. An import goes into one mailbox: Primary, Archive or AuxArchive (-MailboxId).' }
            $store = Open-MbtExportStore -Path $StorePath
            $sourceKind = if ($SourceMailbox) { $SourceMailbox } else { $null }
            if (-not $sourceKind) {
                # the items of an auxiliary archive belong to the folders of the main archive: only Primary / Archive are sources
                $kinds = @($store.GetMailboxes() | Where-Object { $_.Kind.ToString() -ne 'AuxArchive' } | ForEach-Object { $_.Kind.ToString() })
                if ($kinds.Count -ne 1) { throw "The store holds $($kinds.Count) mailboxes ($($kinds -join ', ')): choose with -SourceMailbox Primary|Archive." }
                $sourceKind = $kinds[0]
            }
            $sourceId = Get-StoreMailboxId $sourceKind
            Write-MbtJobStep -Id check -State done -Msg 'The export store is open.'
            $null = Connect-ForTransfer
            $targetKind = Get-AskedKind
            $targetId = Resolve-MbtMailboxId -Mailbox $targetKind -TenantId $TenantId -UserPrincipalName $UserPrincipalName -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId `
                -MailboxId $MailboxId -Store $store -LocationsFile $LocationsFile
            if ($targetKind -eq 'AuxArchive') {
                # an auxiliary archive named by the operator is recorded as such (the import itself learns the others from the redirects)
                $null = Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $store -AuxArchiveMailboxId $targetId -SkipGraph -Verbose:$false
            }
            Write-MbtItem INFO "importing the $sourceKind items of the store into the $targetKind mailbox $targetId$(if ($TargetRoot) { " below '$TargetRoot'" })"
            $r = Import-MbtMailbox -Store $store -SourceMailboxId $sourceId -TargetMailboxId $targetId -TargetRoot $TargetRoot -MappingFile $MappingFile -SourceFolder $Folder `
                -JobName $JobName -Concurrency $Concurrency -MaxAttempts $MaxAttempts -NoAutoReconcile:$NoAutoReconcile -ReconcileDelaySeconds $ReconcileDelaySeconds -ReconcileRounds $ReconcileRounds -IntervalSeconds $IntervalSeconds -WhatIf:$dryRun -Confirm:$false
            if ($dryRun) { Write-MbtItem WARN 'WhatIf: nothing was created and nothing was sent.' }
            elseif ($r.Error) { throw $r.Error.Message }
            else {
                $jr = $r.Result
                $script:counts = @{ done = [long]$jr.JobDone; failed = [long]$jr.JobFailed; unknown = [long]$jr.JobUnknown; pending = [long]$jr.JobPending; total = [long]$jr.JobTotal; foldersCreated = [long]$jr.FoldersCreated }
                if (-not $jr.JobComplete) {
                    $exitCode = if ($jr.JobUnknown -gt 0) { 3 } else { 2 }
                    Write-MbtItem WARN ("exit code ${exitCode}: " + $jr.JobSummary())
                }
                if ($jr.JobUnknown -gt 0) {
                    # the rows the automatic reconciliation could not settle: one decision each (with the items of the folder that could be the one)
                    try {
                        $rows = @(Get-MbtImportUnknown -Store $store -Check)
                        foreach ($u in $rows | Select-Object -First 200) {
                            $n = @($u.Candidates).Count
                            Write-MbtEvent decision ([ordered]@{
                                    key = [string]$u.ImportKey; msg = $(if ($n -gt 1) { "Two or more items in the folder could be '$($u.Subject)'. We do not guess. Is it already in the mailbox?" } elseif ($n -eq 1) { "One item in the folder could be '$($u.Subject)'. Is it already in the mailbox?" } else { "'$($u.Subject)' was sent but no answer came and it was not found in the folder." })
                                    candidates = @($u.Candidates | ForEach-Object { [string]$_ }); item = [string]$u.ItemKey; folder = [string]$u.TargetFolder; size = [long]$u.Size
                                })
                        }
                    } catch { Write-MbtEvent warn ([ordered]@{ code = 'DECISION_LIST'; msg = "The uncertain items could not be listed: $($_.Exception.Message)" }) }
                }
                Write-MbtJobStep -Id final -State $(if ($jr.JobComplete) { 'done' } else { 'running' }) -Done $jr.JobDone -Total $jr.JobTotal -Msg $jr.JobSummary()
            }
        }
    }
    if ($exitCode -ne 0 -and -not $script:reason -and (Test-MbtStopRequested)) { $script:reason = 'stopped' }
    if ($script:reason -eq 'stopped' -and -not $script:next) { $script:next = 'continue' }
    $resultArgs = @{ ExitCode = $exitCode; Counts = $script:counts }
    if ($script:resultComplete) { $resultArgs.Complete = $true }
    if ($script:next) { $resultArgs.Next = $script:next }
    if ($script:reason) { $resultArgs.Reason = $script:reason }
    Write-MbtJobResult @resultArgs
    $script:sawResult = $true
    if ($Mode -ne 'Browse') { Write-MbtRunSummary -ExitCode $exitCode }
} catch {
    $hint = try { Get-ErrorHint $_.Exception.Message } catch { '' }
    $failure = Write-MbtJobError -Message ($_.Exception.Message + $hint)
    $needsSignIn = $failure.Code -eq 'signin-required'
    $exitCode = if ($needsSignIn) { 2 } else { 1 }
    if ($EventsFile) { Write-MbtJobResult -ExitCode $exitCode -Counts $script:counts -Next $(if ($needsSignIn) { 'signin' } else { 'retry' }) -Reason $(if ($needsSignIn) { 'signin-required' }) }
    $script:sawResult = $true
    Write-MbtRunSummary -ExitCode $exitCode -Failure $failure.Message
    # a sign-in that is missing is a normal stop (exit 2, resumable); any other error is thrown as before (exit 1)
    if ($needsSignIn) { Write-MbtItem FAIL $failure.Message } elseif ($hint) { throw ($_.Exception.Message + $hint) } else { throw }
} finally {
    if ($store) { $store.Dispose() }
    Disconnect-MbtGraph
    Stop-MbtStopWatch
    Stop-MbtKeepAwake
    Close-MbtEventLog
}
exit $exitCode
