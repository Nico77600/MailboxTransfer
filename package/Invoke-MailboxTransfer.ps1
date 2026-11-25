#Requires -Version 7.6

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
                  puts them in a normal folder named "Restored deleted items", never back into Recoverable Items. -ArrivedFrom / -ArrivedTo keep only the
                  items that arrived in the mailbox inside a period (see them below).
    Mode ExportPst  Converts the store into PST files with our own writer (no Outlook, no third-party component): the real folder tree of each
                  location under one top folder (Primary, Archive, "Auxiliary archive n (guid8)"), the standard folders first, split at -PstMaxSizeGB
                  (default 50 = 51,200 MiB) on item boundaries, <name>.pst, <name>-002.pst... and <name>.manifest.json. The auxiliary archives are merged
                  into the Archive tree (same folder entry key) unless -SeparateAuxArchives. Resumable: run it again after an interruption, the finished
                  parts (verified by size and PST header) are kept. Each part is checked (Test-MbtPst) unless -SkipCheck. The store is only read.
                  -Mode Export -OutputPst <file.pst> does the export, then the same conversion.
    Batch         -Mode Export -MailboxCsv <file.csv> exports the mailboxes of a list, one ordinary export job for each (its own job folder, events, status.json, report.html), 1 to 4 mailboxes at a time
                  (-ParallelMailboxes). The columns of the CSV file (headers not case sensitive, delimiter comma or semicolon, UTF-8): UserPrincipalName (required), Location (Primary, Archive or Both;
                  empty: -Mailbox, Primary unless given), OutputPst (optional: a .pst file or a folder), SelectionFile (optional), Notes (ignored). The whole file is checked first and every problem is
                  printed together; nothing starts while one remains. -OutputPst is then a folder (one <address>.pst for each mailbox) or a .pst name that gets the address added; -StorePath is the folder
                  that holds one export store for each mailbox (default <work folder>\stores\<batch>). The batch is a folder, <work folder>\batches\<batch>\ (batch.json, batch.log, report.html);
                  the same command again, or -ResumeBatch <batch>, skips the complete mailboxes, continues the ones that are not complete (their own job folder: only what is left) and tries the failed ones
                  again. Sign-in is done once, here; a mailbox never asks for it (-NoPrompt): one that needs it fails with that reason and the batch goes on. A stop request (stop.request in the batch folder,
                  or -StopFile) or Ctrl+C stops the running mailboxes the way the console does; run the batch again to continue. Exit code: 0 every mailbox complete, 1 the file has problems or a mailbox failed,
                  2 not complete (stopped, or mailboxes left): nothing was lost.
    Import batch  -Mode Import (export stores), ImportPst (PST files) or ImportEml (folders of .eml files) with -MailboxCsv imports a list: one ordinary import job for each line, up to -ParallelMailboxes at a
                  time. Columns (headers not case sensitive): Mailbox (the target address, required), Source (required: the store folder, the .pst file or folder, the .eml folder; relative to the CSV file),
                  Location (Primary or Archive; empty: -Mailbox), TargetRoot (empty: -TargetRoot; a store may keep the folder names, a PST needs one or -MergeWellKnown, .eml files need one), SourceMailbox (-Mode Import:
                  Primary or Archive of a store that holds both; empty: -SourceMailbox), Notes (ignored); any other column is a problem. The whole file is checked first, every problem printed together: a source that is
                  missing or of the wrong kind, the same target, Location and Source twice, and with a delegated sign-in a target that is not the signed-in account (the service refuses a delegated write into another
                  mailbox: use -AppOnly). Lines that write into the same mailbox location, or read the same source, run one after the other, never together. Every import is create only and is never done twice:
                  the same line again, in this batch or another, skips what was imported. -WhatIf lists every target and writes nothing. Exit codes as above.
    Mode Copy     Copies ONE mailbox location into ONE other in a single job: the export of the source into an export store, then the import of that store into the target, chained by this one
                  command (it is not a direct copy: nothing is streamed from one mailbox to the other). The store stays on disk: it is the resume point, the proof of what was read, and it can be
                  imported again (-Mode Import) or opened in the console. The source is given as for an export (-UserPrincipalName, -Mailbox Primary or Archive, -MailboxId, -SelectionFile or -Folder, the
                  period -ArrivedFrom / -ArrivedTo, -IncludeRecoverableItems), the target as for an import (-TargetMailbox Primary or Archive, -TargetMailboxId, -TargetRoot, or the same folder names, -MappingFile).
                  -WhatIf prints ONE plan for both halves (what is read, the store, where it goes, the time at about 4 GiB per hour per target mailbox, dominated by the import) and creates nothing.
                  The same command again continues where it stopped: an export half that is complete is not read again (-Force reads the source again to take what arrived since), the import resumes
                  its job (named after the target, the store and the folder, whatever the case) and never creates an item twice. A source and a target that are the same location are refused. Sign-in:
                  delegated, it reads what the signed-in account can read and writes into the mailbox of that account only (the service refuses a delegated write into another person's or a shared mailbox:
                  use -AppOnly with -TargetMailboxId); with a certificate (app-only) it reads and writes the mailboxes of its scope. An export half that left items (a failed item, a deferred folder)
                  does not stop the import of what was read: the copy ends with exit code 2 and the next run does only what is left.
    Copy batch    -Mode Copy with -MailboxCsv copies a list: one ordinary copy job for each line, up to -ParallelMailboxes at a time. Columns (headers not case sensitive): SourceAddress (the mailbox read, required),
                  SourceLocation (Primary or Archive; empty: -Mailbox), TargetAddress (the mailbox written into, required), TargetLocation (Primary or Archive; empty: -TargetMailbox), TargetRoot (empty: -TargetRoot; empty
                  everywhere: the folder names of the source), SelectionFile (optional: what to copy, of one mailbox location), Notes (ignored); any other column is a problem. The whole file is checked first, every problem
                  printed together: an address, a Location, a selection, a source that is the target, the same copy twice, and with a delegated sign-in a target that is not the signed-in account (the service refuses a
                  delegated write into another mailbox: use -AppOnly) and the archive of another mailbox as a source. Lines that write into the same target mailbox location, or read the same source mailbox location, run
                  one after the other, never together; a source copied into two targets is read once (one store). -StorePath is the folder that holds the stores of the batch (default <work folder>\stores\<batch>);
                  -ArrivedFrom, -ArrivedTo, -Folder, -IncludeRecoverableItems, -Force and the tuning options apply to every line. Every copy is create only and is never done twice: the same line again, in this batch or
                  another, skips what was copied. -WhatIf lists every line and writes nothing. Exit codes as above.
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
    Exit code (Mode Copy): the codes of the import, for the whole job: 0 the export half is complete and every row of the import is Done; 1 an error stopped it, or the CSV file / the request was refused (nothing was read);
    2 not complete but nothing was lost: a stop, an export half that left items (a failed item, a deferred folder, which does not stop the import of what was read) or Pending / Failed rows: run the same command again;
    3 the import holds Unknown rows (settle them with -Mode Status -Unknown -Check, then run again).

.PARAMETER Mode
    Export, ExportPst, Import, ImportPst, ExportEml, ImportEml, Copy (a mailbox into another: the export, then the import of its store), Browse, CheckPst, RepairPst or Status (default).

.PARAMETER DeviceCode
    Sign in with a device code instead of the browser. The line "DEVICE CODE: <code> URL: <url>" is written at once to the console and
    to out\device-code.txt. A cached token (DPAPI) makes it silent.

.PARAMETER AppOnly
    Sign in as the application with a certificate (client credentials, no user, no prompt) using Auth.CertificateThumbprint or Auth.CertificatePath of the configuration. Implied by -CertificateThumbprint or -CertificatePath. See docs\MailboxTransfer-Guide.md, chapter 4.6.

.PARAMETER CertificateThumbprint
    App-only: the thumbprint of the application's certificate (CurrentUser\My, then LocalMachine\My; it needs its private key). Works for every mode. -TenantId and -ClientId are then those of the certificate application (Auth.AppTenantId / Auth.AppClientId of the configuration); the application of the Tenant section is never used.

.PARAMETER CertificatePath
    App-only: a .pfx file with the certificate and its private key (-CertificatePassword: SecureString).

.PARAMETER CertificatePassword
    App-only: the password of the .pfx, as a SecureString.

.PARAMETER UserPrincipalName
    The mailbox user (default: Tenant.UserPrincipalName of the configuration). Sign-in is delegated as this user. With -CertificateThumbprint / -CertificatePath it is the OWNER OF THE TARGET MAILBOX (there is no signed-in user and no /me), it is required, and the configuration default is never used.

.PARAMETER Mailbox
    Copy: the mailbox location READ (Primary or Archive; the target is -TargetMailbox). Primary (default), Archive (the main archive; with Export its expanded folders are followed into the auxiliary archives), AuxArchive
    (one auxiliary archive of an auto-expanding archive: needs -MailboxId, or a single one known) or All (Export only: Primary, Archive,
    then every auxiliary archive known): the mailbox exported from (Export) or imported into (Import). The ids come from -MailboxId,
    -PrimaryMailboxId / -ArchiveMailboxId, the Tenant section of the configuration (for the user of the configuration only), -LocationsFile, the locations recorded in the store,
    or (Primary; Archive too with User.Read.All and Graph beta) Graph.

.PARAMETER MailboxId
    Copy: the mailbox READ (the target is -TargetMailboxId). One explicit location, MBX:<mailbox guid>@<tenant id>, the bare guid, or the SMTP address / UPN of a mailbox (a SHARED mailbox or another user's: it is resolved
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
    recovery) later. Only one process at a time may write to it; Status can read it meanwhile. Copy: default <work folder>\stores\copy-<source>-<location>-<8 digits that come from the choice: selection, folder, period>.
    A store holds one copy of one source: another choice never lands in it (it is refused), and the same source copied into two targets shares the one store.

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
    Import, Copy: path in the target below which everything is created, for instance MBT-Phase1-A. Without it the same paths as the source.

.PARAMETER TargetMailbox
    Copy: the location of the target mailbox written into, Primary (default) or Archive. The mailbox is the signed-in account's own, or -TargetMailboxId. The source and the target must not be the same location.

.PARAMETER TargetMailboxId
    Copy: the mailbox written into: its SMTP address, or its id (MBX:<mailbox guid>@<tenant id>, or the bare guid). Without it the signed-in account's own mailbox. A delegated sign-in writes into its own mailbox only: another
    address is refused before anything is read. The application sign-in (-AppOnly, a certificate) needs it, since the application has no mailbox of its own (then -UserPrincipalName or -MailboxId is the mailbox read).

.PARAMETER MappingFile
    Import, Copy: .psd1 or .json, source path prefix -> target path prefix (relative to -TargetRoot). Longest prefix wins.

.PARAMETER NoAutoReconcile
    Import: do not look up the Unknown rows by itself (today's behaviour: they stay Unknown, exit code 3, settle with -Mode Status -Unknown -Check).

.PARAMETER ReconcileDelaySeconds
    Import: seconds between the end of the request that left a row Unknown and the lookup in the target folder (default 300: the service
    may still be creating the item). For the Unknown rows of an earlier run the delay counts from the last change of the row.

.PARAMETER JobName
    Import, Copy: the name of the import job in the store (default import). Running again with the same name resumes it. A copy derives copy-<12 digits> from the target, its Location, the store, the folder and the source, whatever the case,
    so that the same copy never creates an item twice.

.PARAMETER ReconcileRounds
    Import: how many times an item may be looked up and imported again before it is left Unknown (default 3, 1 to 20).

.PARAMETER Concurrency
    Workers, 1 to 4 (default 4). The engine lowers it by itself while the service throttles.

.PARAMETER NotReadyTrip
    Export: after this many folders in a row whose content is not available yet (an auxiliary archive being filled), the next folders are asked once instead of 3 times with a back-off, and deferred if they still refuse (default 3; 0 = always 3 asks). A folder that lists again puts the full asks back.

.PARAMETER Folder
    Export, Import: only this folder and the folders below it (path as in the store). Browse: the folder to list (id or path, or the end of a path
    such as Inbox/Reports).

.PARAMETER ArrivedFrom
    Export, Copy, ExportEml (live mailbox): the first day, or the first instant, of the period of the export. Only the items that arrived in the mailbox on or after it are listed and exported (the bound is
    included). Written year first, whatever the language of Windows: 2026-03-01 (the whole day, from midnight of the local time zone) or 2026-03-01T08:30 (used as it is; seconds, a fraction and
    Z or +hh:mm are accepted). Not case sensitive. Left out: no start. "Arrived" is the date the item was created in the mailbox (createdDateTime of the import-export listing). For mail that was migrated
    or imported into the mailbox it is the date of that import, not the original date of the message. One period for the whole export (every mailbox of a -MailboxCsv batch, every location of a selection).
    Items named one by one in a -SelectionFile are always taken, whatever their date: the period applies to the items listed from the folders asked. An item that has no date is kept, and counted.
    Items outside the period are skipped, not an error: the exit code does not change. A store keeps one period for a mailbox: continuing an export with another period is refused. -WhatIf shows the
    folder totals as the most the export can take: the exact numbers come at the end of the run (kept, outside the period, without a date). Not for the import modes, nor for -Mode ExportPst
    and -Mode ExportEml from a store or a PST (they hold no date of arrival): give the period to the -Mode Export that fills the store.

.PARAMETER ArrivedTo
    Export, ExportEml (live mailbox): the last day, or the last instant, of the period (the bound is included). 2026-03-31 is the whole day, up to the last instant of it in the local time zone;
    2026-03-31T18:30 is used as it is. Left out: no end. -ArrivedFrom must not be after -ArrivedTo: the run is refused before anything starts. See -ArrivedFrom.

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
    Export, ExportPst: one top folder per auxiliary archive ("Auxiliary archive 1 (guid8)") instead of merging them into the Archive tree.

.PARAMETER SkipCheck
    Export, ExportPst: do not run the integrity checker (Test-MbtPst) on each finished PST.

.PARAMETER Force
    Export, ExportPst: start the PST conversion again from zero (deletes the manifest and the PST files of an earlier run of the same destination). Copy: read the source again although the export half is complete (what changed or arrived since is taken, and only that is imported).

.PARAMETER SelectionFile
    Export: a selection written by -Mode Browse -SaveSelection (or by the console). The mailbox is the one of the file; the selection is checked
    against the import-export API (a folder is found there by the same id, else by its path; selected items by id, else by size and date) and only
    the selected folders / items are inventoried and exported. An entry that cannot be exported (a folder or item not found)
    stops the run before anything is exported. With -WhatIf only the plan is printed. A file written by the console for several locations
    (the primary mailbox, the archive, an auxiliary archive) holds one selection per mailbox: one command exports all of them into the same store,
    and with -OutputPst the PST files hold what was selected in all of them. ExportPst takes the same file to convert only what it selects.

.PARAMETER RetryWaitSeconds
    The wait before a failed request is tried again, in seconds (default: Graph.BaseBackoffSeconds of the configuration, 2; it doubles each time up to Graph.MaxBackoffSeconds).
    For this run only: nothing is written to the configuration. Retry-After of the service is always obeyed first.

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

.PARAMETER MailboxCsv
    Export: the CSV file of a batch, the mailboxes to export one after the other (see Batch above and the user guide, "Exporting many mailboxes"). The first line names the columns: UserPrincipalName
    (required), Location (Primary, Archive or Both), OutputPst, SelectionFile, Notes. A relative OutputPst or SelectionFile of a row is relative to the folder of the CSV file.
    Import, ImportPst, ImportEml: the CSV file of an import batch (see Import batch above and the user guide, "Importing many mailboxes"): Mailbox, Source, Location, TargetRoot, SourceMailbox, Notes.
    Copy: the CSV file of a batch of copies (see Copy batch above and the user guide, "Copying many mailboxes"): SourceAddress, SourceLocation, TargetAddress, TargetLocation, TargetRoot, SelectionFile, Notes.

.PARAMETER BatchName
    With -MailboxCsv: the name of the batch (letters, digits, space, dot, dash, underscore; it is the name of its folder and the name in the reports). Without it a name is made of the date and the time.
    The same name again continues the batch; a batch is made of the same mailboxes for ever (another list needs another name).

.PARAMETER ResumeBatch
    Export, Import, ImportPst, ImportEml, Copy: continues the batch of that name (with the same -Mode) without the CSV file: complete mailboxes are skipped, the others are done again from where they were. The options of the first run (sign-in, output, tuning) are kept
    unless you give them again.

.PARAMETER ParallelMailboxes
    With -MailboxCsv: how many mailboxes are exported, imported into or copied at the same time, 1 to 4 (default 1; an import or a copy never runs two lines of the same target mailbox location together, nor two copies of the same source). Each one is a process of its own with its own requests; with more than one at a time, each uses 2 parallel
    requests (-Concurrency 4 is the default of one mailbox) unless -Concurrency is given.

.PARAMETER WorkFolder
    With -MailboxCsv or -ResumeBatch, and for the default store of a copy: the work folder that holds jobs, batches, stores and logs (default %LOCALAPPDATA%\MailboxTransfer, the folder of the console).

.PARAMETER BatchId
    The batch a job belongs to (status.json: batch). A batch sets it for the jobs it starts; leave it out.

.PARAMETER ResumeJob
    Runs a job again from its job folder: the name of the folder (the first column of -Mode Status) or the job id, whatever the case. The command that was recorded runs again in the folder it first ran in, in the same job folder, so only what is
    left is done (what each mode does again is in the guide). A single job started from the command line (Export, ExportPst, Import, ImportPst, ImportEml, ExportEml, Copy) has such a folder in the work folder, like the jobs of the console;
    a job that is not finished says this command in its Next steps. -WorkFolder, -StopFile, -LogFolder, -NoPrompt, -KeepAwake and -WhatIf of this run are passed on; a certificate password is never saved, so a job started with one is run again by its own command.

.PARAMETER RemoveJob
    Deletes the folder of a job (its events, status, request, log and report): the name of the folder or the job id, whatever the case. Asks first (-Confirm:$false does not ask, -WhatIf only shows). Never deleted: an export store you chose, PST files, source files,
    mailboxes. Refused: a job that is running, a job of a batch that is not complete. The jobs are listed by -Mode Status.

.PARAMETER RemoveCopyStore
    With -RemoveJob, for a copy that is complete: the copy store (the folder of items read from the source mailbox) goes with the job, when it is in the stores folder of the work folder and no other job that is not finished uses it. Refused when the copy is not complete.

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
    .\Invoke-MailboxTransfer.ps1 -Mode Copy -MailboxId shared@contoso.com -TargetRoot 'From the shared mailbox' -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode Copy -MailboxId shared@contoso.com -Folder Inbox/Projects -TargetRoot 'From the shared mailbox' -ArrivedFrom 2026-01-01
    .\Invoke-MailboxTransfer.ps1 -Mode Copy -SelectionFile D:\sel.json -TargetMailbox Archive -TargetRoot Restored -StorePath D:\copy-store
    .\Invoke-MailboxTransfer.ps1 -Mode Copy -AppOnly -UserPrincipalName alex@contoso.com -TargetMailboxId adele@contoso.com -TargetRoot 'Mail of Alex'
    .\Invoke-MailboxTransfer.ps1 -Mode Copy -MailboxCsv D:\batch\copies.csv -TargetRoot Copied -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode Copy -MailboxCsv D:\batch\copies.csv -BatchName Move1 -ParallelMailboxes 2 -AppOnly
    .\Invoke-MailboxTransfer.ps1 -Mode Copy -ResumeBatch Move1
    .\Invoke-MailboxTransfer.ps1 -Mode ImportPst -PstPath D:\old\archive.pst -TargetRoot Imported/archive -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode ImportPst -PstPath D:\old\parts -TargetRoot Imported/archive -Mailbox Archive
    .\Invoke-MailboxTransfer.ps1 -Mode Browse -Mailbox Archive
    .\Invoke-MailboxTransfer.ps1 -Mode Browse -Folder Inbox -Top 10 -Sort SubjectAscending
    .\Invoke-MailboxTransfer.ps1 -Mode Browse -Folder Inbox -ItemId 2
    .\Invoke-MailboxTransfer.ps1 -Mode Browse -Folder Inbox/Reports -IncludeSubfolders -SaveSelection D:\sel.json
    .\Invoke-MailboxTransfer.ps1 -Mode Export -SelectionFile D:\sel.json -StorePath D:\store
    .\Invoke-MailboxTransfer.ps1 -Mode Export -SelectionFile D:\selections\all-locations.json -StorePath D:\store -OutputPst D:\pst\adele.pst -RetryWaitSeconds 5
    .\Invoke-MailboxTransfer.ps1 -Mode Export -Mailbox Archive -StorePath D:\store -OutputPst D:\pst\adele.pst -PstMaxSizeGB 10
    .\Invoke-MailboxTransfer.ps1 -Mode Export -StorePath D:\store -ArrivedFrom 2026-01-01 -ArrivedTo 2026-03-31
    .\Invoke-MailboxTransfer.ps1 -Mode Export -MailboxCsv D:\batch\mailboxes.csv -OutputPst D:\pst -ArrivedFrom 2025-07-01
    .\Invoke-MailboxTransfer.ps1 -Mode ExportPst -StorePath D:\store -OutputPst D:\pst\adele.pst
    .\Invoke-MailboxTransfer.ps1 -Mode ExportPst -StorePath D:\store -OutputPst D:\pst\adele.pst -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode Export -Mailbox All -StorePath D:\store -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode Export -MailboxCsv D:\batch\mailboxes.csv -OutputPst D:\pst -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode Export -MailboxCsv D:\batch\mailboxes.csv -BatchName Wave1 -OutputPst D:\pst -ParallelMailboxes 2 -AppOnly
    .\Invoke-MailboxTransfer.ps1 -Mode Export -ResumeBatch Wave1
    .\Invoke-MailboxTransfer.ps1 -Mode ImportPst -MailboxCsv D:\batch\imports.csv -TargetRoot Imported -WhatIf
    .\Invoke-MailboxTransfer.ps1 -Mode ImportPst -MailboxCsv D:\batch\imports.csv -BatchName Restore1 -TargetRoot Imported -ParallelMailboxes 2 -AppOnly
    .\Invoke-MailboxTransfer.ps1 -Mode ImportPst -ResumeBatch Restore1
    .\Invoke-MailboxTransfer.ps1 -Mode Status
    .\Invoke-MailboxTransfer.ps1 -ResumeJob 20261124-2241-export-3fa2
    .\Invoke-MailboxTransfer.ps1 -RemoveJob 20261124-2241-export-3fa2 -WhatIf
    .\Invoke-MailboxTransfer.ps1 -RemoveJob 20261110-0915-mailcopy-9b1c -RemoveCopyStore -Confirm:$false
    .\Invoke-MailboxTransfer.ps1 -Mode Status -StorePath D:\store
    .\Invoke-MailboxTransfer.ps1 -Mode Status -StorePath D:\store -Unknown -Check
    .\Invoke-MailboxTransfer.ps1 -Mode Status -StorePath D:\store -Unknown -Check -ResolveAbsent -WhatIf

.NOTES
    Author  : Nicolas Fabert
    Version : 1.3.0
    History : see CHANGELOG.md

    This Script is a Personal project.
    It's provided "AS-IS". It's not an official Microsoft product so no support can be expected from Microsoft.

    As any scripts you must read carefully the documentation and test it first in a Test environment before any run in Production.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet('Export', 'ExportPst', 'Import', 'ImportPst', 'Browse', 'Status', 'ExportEml', 'ImportEml', 'CheckPst', 'RepairPst', 'Copy')][string]$Mode = 'Status',
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
    [object]$ArrivedFrom,
    [object]$ArrivedTo,
    [string]$SelectionFile,
    [switch]$SkipInventory,
    [int]$MaxAttempts = 3,
    [ValidateRange(0.1, 3600)][double]$RetryWaitSeconds,
    [ValidateRange(0, 100000)][int]$NotReadyTrip = 3,
    [ValidateRange(1, 20)][int]$BatchItems = 20,
    [ValidateRange(1, 4096)][int]$BatchMiB = 200,

    [ValidateSet('Primary', 'Archive')][string]$SourceMailbox,
    [string]$TargetRoot,
    [ValidateSet('Primary', 'Archive')][string]$TargetMailbox = 'Primary',
    [string]$TargetMailboxId,
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
    [string]$MailboxCsv,
    [string]$BatchName,
    [string]$ResumeBatch,
    [ValidateRange(1, 4)][int]$ParallelMailboxes = 1,
    [string]$WorkFolder,
    [string]$BatchId,
    [string]$ResumeJob,
    [string]$RemoveJob,
    [switch]$RemoveCopyStore,
    [string]$Path,
    [string]$Destination,
    [ValidateSet('Automatic', 'Recompute', 'Rebuild', 'Salvage')][string]$Level = 'Automatic'
)

function Get-PathReparseTag {
    <#
        The tag of the reparse point at Path, or $null when Path is not one. WIN32_FIND_DATAW.dwReserved0 holds the tag when FILE_ATTRIBUTE_REPARSE_POINT
        is set; FindFirstFileW reads it from the directory entry without opening the item, so a cloud placeholder is not downloaded. The two calls are
        declared with Reflection.Emit: no C# is compiled, because this runs before the engine exists. Layout of the 592-byte structure: dwFileAttributes
        at offset 0, dwReserved0 at offset 36. The same text is in the verifier, the module, both entry scripts and the two tools of the repository.
    #>
    param([Parameter(Mandatory)][string]$Path)
    $api = Get-Variable -Name FindFileApi -Scope Script -ValueOnly -ErrorAction Ignore
    if ($null -eq $api) {
        $module = [Reflection.Emit.AssemblyBuilder]::DefineDynamicAssembly([Reflection.AssemblyName]::new('ReparseTag'), [Reflection.Emit.AssemblyBuilderAccess]::Run).DefineDynamicModule('ReparseTag')
        $native = $module.DefineType('Native', 'Public, Abstract, Sealed')
        foreach ($call in @(@('FindFirstFileW', [IntPtr], @([string], [IntPtr])), @('FindClose', [bool], @([IntPtr])))) {
            $method = $native.DefinePInvokeMethod($call[0], 'kernel32.dll', $call[0], 'Public, Static, PinvokeImpl', 'Standard', $call[1], [type[]]$call[2], 'Winapi', 'Unicode')
            $method.SetImplementationFlags('PreserveSig')
        }
        $api = $native.CreateType()
        Set-Variable -Name FindFileApi -Scope Script -Value $api
    }
    $full = [IO.Path]::GetFullPath($Path)
    $full = if ($full.StartsWith('\\?\', [StringComparison]::Ordinal)) { $full } elseif ($full.StartsWith('\\', [StringComparison]::Ordinal)) { '\\?\UNC\' + $full.Substring(2) } else { '\\?\' + $full }
    $data = [Runtime.InteropServices.Marshal]::AllocHGlobal(592)
    try {
        $handle = $api::FindFirstFileW($full, $data)
        if ($handle -eq [IntPtr]::new(-1)) { throw "Package integrity: the reparse tag of this path cannot be read (Windows error $([Runtime.InteropServices.Marshal]::GetLastSystemError())): $Path" }
        [void]$api::FindClose($handle)
        if (([Runtime.InteropServices.Marshal]::ReadInt32($data, 0) -band 0x400) -eq 0) { return $null }
        return [uint32]([long][Runtime.InteropServices.Marshal]::ReadInt32($data, 36) -band 0xFFFFFFFFL)
    }
    finally { [Runtime.InteropServices.Marshal]::FreeHGlobal($data) }
}
function Test-ReparseTagAllowed {
    <#
        The rule for the tag of a reparse point, not for the bare attribute. Refused: every tag that redirects to another path, that is every tag with the
        name-surrogate bit 0x20000000 (symbolic link 0xA000000C, junction and mount point 0xA0000003, and the others), and every tag that is not in the closed
        list below (fail closed). Accepted, the Microsoft types that do not redirect: the cloud files of OneDrive (0x9000001A and its variants 0x9000101A to
        0x9000F01A, that is the tag with the bits 0x0000F000 cleared), data deduplication 0x80000013 and WOF (Windows Overlay Filter, CompactOS) 0x80000017.
        The SHA256 of the signed manifest still detects any change of content, including one made by a synchronisation.
    #>
    param([Parameter(Mandatory)][uint32]$Tag)
    if ($Tag -band 0x20000000u) { return $false }
    return (($Tag -band 0xFFFF0FFFu) -eq 0x9000001Au) -or $Tag -eq 0x80000013u -or $Tag -eq 0x80000017u
}
function Get-RefusedReparseTag {
    <# $null when the item (from Get-Item -Force) is not a reparse point or its tag is accepted by Test-ReparseTagAllowed; otherwise the refused tag as eight hexadecimal digits. #>
    param([Parameter(Mandatory)]$Item)
    if (-not ($Item.Attributes -band [IO.FileAttributes]::ReparsePoint)) { return $null }
    $tag = [uint32](Get-PathReparseTag $Item.FullName)
    if (Test-ReparseTagAllowed $tag) { return $null }
    return $tag.ToString('X8')
}

if ($Gui) { & (Join-Path $PSScriptRoot 'Start-MailboxTransferConsole.ps1') -DevelopmentMode:$DevelopmentMode; exit $LASTEXITCODE }

# The package is verified before any of its code runs: the verifier must carry the owner signature, then it checks the signed manifest.
# A tree without a manifest (a source checkout) is accepted only with -DevelopmentMode.
$integrityManifest = Join-Path $PSScriptRoot 'PackageIntegrity.psd1'
if (Microsoft.PowerShell.Management\Test-Path -LiteralPath $integrityManifest) {
    $integrityVerifier = Join-Path $PSScriptRoot 'src\Test-PackageIntegrity.ps1'
    $integrityCursor = [IO.Path]::GetFullPath($integrityVerifier)
    while ($integrityCursor) {
        $integrityItem = Microsoft.PowerShell.Management\Get-Item -LiteralPath $integrityCursor -Force -ErrorAction Stop
        $integrityRefused = Get-RefusedReparseTag $integrityItem
        if ($integrityRefused) { throw "Package integrity: reparse point that is not accepted (tag 0x$integrityRefused): $integrityCursor" }
        $integrityCursor = [IO.Path]::GetDirectoryName($integrityCursor)
    }
    $integritySignature = Microsoft.PowerShell.Security\Get-AuthenticodeSignature -LiteralPath $integrityVerifier -ErrorAction Stop
    if ([string]$integritySignature.Status -cne 'Valid' -or $null -eq $integritySignature.SignerCertificate -or
        $integritySignature.SignerCertificate.Thumbprint -ne '0F28E442889F17B5F4D251E8F96E7CC6E2C9C6D6' -or
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

# Every number and date of the tool is written the same way, whatever the language of Windows: en-US for the threads of this process and the ones that start later.
$mbtCulture = [Globalization.CultureInfo]::GetCultureInfo('en-US')
[Globalization.CultureInfo]::DefaultThreadCurrentCulture = $mbtCulture
[Globalization.CultureInfo]::DefaultThreadCurrentUICulture = $mbtCulture
[Threading.Thread]::CurrentThread.CurrentCulture = $mbtCulture
[Threading.Thread]::CurrentThread.CurrentUICulture = $mbtCulture
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# -WhatIf only means "plan the import": the files of the run (store, transcript) are always written.
$mailboxGiven = $PSBoundParameters.ContainsKey('Mailbox')
$boundParameters = $PSBoundParameters
$entryScript = $PSCommandPath
$dryRun = [bool]$WhatIfPreference
$WhatIfPreference = $false
Import-Module (Join-Path $PSScriptRoot 'MailboxTransfer.psd1') -Force

# -ResumeJob and -RemoveJob work on one job folder of the work folder (the folders the console lists) and do nothing else.
if ($ResumeJob -or $RemoveJob) {
    if ($ResumeJob -and $RemoveJob) { throw '-ResumeJob and -RemoveJob each name one job: give one of them.' }
    $jobWork = if ($WorkFolder) { [IO.Path]::GetFullPath($WorkFolder) } else { Get-MbtDefaultWorkFolder }
    $jobAsked = if ($ResumeJob) { $ResumeJob } else { $RemoveJob }
    $foundJob = Find-MbtJob -WorkFolder $jobWork -Name $jobAsked
    if (-not $foundJob) { throw "No job is named '$jobAsked' in $(Join-Path $jobWork 'jobs'). -Mode Status lists the jobs (add -WorkFolder when they are in another work folder)." }
    try { [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false) } catch { Write-Verbose "Console output encoding unchanged: $($_.Exception.Message)" }
    $jobClock = [Diagnostics.Stopwatch]::StartNew()
    $jobLog = $null
    try { $jobLog = Start-MbtLog -Directory $(if ($LogFolder) { $LogFolder } else { Join-Path $jobWork 'logs' }) } catch { Write-Warning "No log file: $($_.Exception.Message)" }
    if ($RemoveJob) {
        $details = [ordered]@{ Mode = @('Mode', ('RemoveJob' + $(if ($dryRun) { ' - WhatIf (nothing is deleted)' }))); Job = @('Folder', $foundJob.Name) }
        if ($jobLog) { $details.Log = @('Log', $jobLog) }
        Write-MbtBanner -Title 'Mailbox Transfer' -Subtitle 'Deletes the folder of a job: its events, log and report' -Details $details
        Write-MbtStep "Delete the job $($foundJob.Name)"
        $removal = Get-MbtJobRemovalPlan -WorkFolder $jobWork -Job $foundJob -WithSize:$RemoveCopyStore
        $values = [ordered]@{}
        $code = 0
        try {
            Write-MbtItem INFO $removal.Title
            $removalWhy = Get-MbtJobDescription -Job $foundJob
            if ($removalWhy.Why) { Write-MbtDetail $removalWhy.Why }
            if (-not $removal.Allowed) { throw $removal.Reason }
            foreach ($warning in $removal.Warnings) { Write-MbtItem WARN $warning }
            Write-MbtItem INFO "deleted: the folder $($removal.Folder) (events, status, request, log and report of the job)"
            Write-MbtItem INFO 'never deleted: an export store you chose, PST files, source files, mailboxes'
            $confirmation = @{}
            if ($boundParameters.ContainsKey('Confirm')) { $confirmation.Confirm = [bool]$boundParameters['Confirm'] }
            $removed = Remove-MbtJob -WorkFolder $jobWork -Name $foundJob.Name -IncludeStore:$RemoveCopyStore -WhatIf:$dryRun @confirmation
            if ($removed.StoreRemoved) { Write-MbtItem OK ('the copy store was deleted ({0})' -f (Format-MbtBytes ([long]$removed.StoreBytes))); $values['Copy store'] = @('Store', 'deleted') }
            if ($removed.Removed) { Write-MbtItem OK 'the job folder was deleted'; $values.Job = @('Folder', "$($foundJob.Name) deleted") }
            elseif ($dryRun) { Write-MbtItem WARN 'WhatIf: nothing was deleted.'; $values.Job = @('Folder', "$($foundJob.Name) - WhatIf, nothing was deleted") }
            else { Write-MbtItem WARN 'nothing was deleted.'; $values.Job = @('Folder', "$($foundJob.Name) kept") }
        } catch {
            $code = 1
            Write-MbtItem FAIL $_.Exception.Message
            $values.Error = @('Fail', $_.Exception.Message)
            $values.Job = @('Folder', "$($foundJob.Name) kept")
        }
        $values = [ordered]@{ Result = @($(if ($code) { 'Fail' } else { 'Ok' }), "exit code $code") } + $values
        $values.Duration = @('Clock', (Format-MbtDuration $jobClock.Elapsed.TotalSeconds))
        if ($code) { $values['Next steps'] = @('Plan', @('-Mode Status lists the jobs and says why a job cannot be deleted.')) }
        if ($jobLog) { $values.Log = @('Log', $jobLog) }
        Write-MbtSummary -Title $(if ($code) { 'RemoveJob failed' } elseif ($dryRun) { 'RemoveJob planned (WhatIf, nothing was deleted)' } else { 'RemoveJob finished' }) -Values $values -Status $(if ($code) { 'Fail' } else { 'Ok' })
        exit $code
    }
    # ResumeJob: the recorded command, again, in the folder it first ran in
    $jobRequest = Read-MbtJobRequest -Folder $foundJob.Folder
    if ($null -eq $jobRequest) { throw "The job '$($foundJob.Name)' has no request.json: nothing says what it ran, so it cannot be run again. Run its command again." }
    if ([string](Get-MbtJobField $jobRequest 'Source' '') -eq 'cli') {
        if ([bool](Get-MbtJobField $jobRequest 'NeedsSecret' $false)) { throw "The job '$($foundJob.Name)' was started with a certificate password, which is never saved. Run its command again with -CertificatePassword." }
        $jobFolder = [string](Get-MbtJobField $jobRequest 'WorkingDirectory' '')
        if (-not $jobFolder -or -not (Test-Path -LiteralPath $jobFolder -PathType Container)) { throw "The job '$($foundJob.Name)' ran in the folder '$jobFolder', which is not there any more: its journals and its relative paths are relative to it. Create the folder again, or run its command again: $(Get-MbtJobField $jobRequest 'CommandLine' '')" }
        $replay = Get-MbtCliJobParameter -Request $jobRequest
        $replay['EventsFile'] = Join-Path $foundJob.Folder 'events.jsonl'
        $replay['JobId'] = if ($foundJob.JobId) { [string]$foundJob.JobId } else { $foundJob.Name }
        foreach ($forward in 'StopFile', 'LogFolder', 'NoPrompt', 'KeepAwake', 'DevelopmentMode') { if ($boundParameters.ContainsKey($forward)) { $replay[$forward] = $boundParameters[$forward] } }
        if ($NoPrompt -and $replay.Contains('DeviceCode')) { $replay.Remove('DeviceCode') }
        if ($WorkFolder) { $replay['WorkFolder'] = $WorkFolder }
        if ($dryRun) { $replay['WhatIf'] = $true }
        Write-MbtStep "Resume the job $($foundJob.Name)"
        Write-MbtItem INFO "the recorded command: $(Get-MbtJobField $jobRequest 'CommandLine' '')"
        Write-MbtItem INFO "runs in: $jobFolder"
        Push-Location -LiteralPath $jobFolder
        try { & $PSCommandPath @replay } finally { Pop-Location }
        exit $LASTEXITCODE
    }
    # a job of the console: its command lines are in request.json, they run as the console runs them
    $jobCommands = @(Get-MbtJobField $jobRequest 'Commands' @())
    if (-not $jobCommands.Count) { throw "The job '$($foundJob.Name)' has no recorded command: it cannot be run again." }
    Write-MbtStep "Resume the job $($foundJob.Name)"
    Write-MbtItem INFO "the job was started by the console: its $($jobCommands.Count) recorded command line(s) run one after the other"
    $jobExit = 0
    foreach ($command in $jobCommands) {
        $childArgs = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath) + $(if ($DevelopmentMode) { @('-DevelopmentMode') } else { @() }) + @('-LogFolder', $(if ($LogFolder) { $LogFolder } else { Join-Path $jobWork 'logs' })) + @($command | ForEach-Object { [string]$_ })
        & (Get-Process -Id $PID).Path @childArgs
        $jobExit = $LASTEXITCODE
        if ($jobExit -notin 0, 1) { break }
    }
    exit $jobExit
}

$config = if ($ConfigPath) { Import-MbtConfiguration -Path $ConfigPath } else { Import-MbtConfiguration }
if ($PSBoundParameters.ContainsKey('RetryWaitSeconds')) { $config.Graph.BaseBackoffSeconds = $RetryWaitSeconds }
if (-not $TenantId) { $TenantId = $config.Tenant.TenantId }
if (-not $ClientId) { $ClientId = $config.Tenant.ClientId }
if (-not $UserPrincipalName) { $UserPrincipalName = $config.Tenant.UserPrincipalName }
# the mailbox ids written in the configuration belong to the user of the configuration: for another user they are discovered, never borrowed
$configured = Get-MbtConfiguredMailboxId -Configuration $config -UserPrincipalName $UserPrincipalName
if (-not $PrimaryMailboxId) { $PrimaryMailboxId = $configured.Primary }
if (-not $ArchiveMailboxId) { $ArchiveMailboxId = $configured.Archive }
# App-only (certificate): the application, the tenant and the target mailbox are always given (parameters or the Auth.App* keys), never the user defaults of a delegated sign-in.
if (-not $CertificateThumbprint -and -not $CertificatePath -and $AppOnly) { $CertificateThumbprint = $config.Auth.CertificateThumbprint; $CertificatePath = $config.Auth.CertificatePath }
$useCertificate = [bool]($CertificateThumbprint -or $CertificatePath)
if ($AppOnly -and -not $useCertificate) { throw '-AppOnly needs -CertificateThumbprint or -CertificatePath (or Auth.CertificateThumbprint / Auth.CertificatePath in the configuration).' }
if ($useCertificate) {
    if (-not $PSBoundParameters.ContainsKey('TenantId')) { $TenantId = if ($config.Auth.AppTenantId) { $config.Auth.AppTenantId } else { $config.Tenant.TenantId } }
    if (-not $PSBoundParameters.ContainsKey('ClientId')) { $ClientId = $config.Auth.AppClientId }
    if (-not $ClientId) { throw 'App-only: give -ClientId (the application with the certificate) or set Auth.AppClientId in the configuration.' }
    if (-not $TenantId) { throw 'App-only: give -TenantId (tenant id or domain) or set Auth.AppTenantId (or Tenant.TenantId) in the configuration.' }
    if (-not $PSBoundParameters.ContainsKey('UserPrincipalName')) { $UserPrincipalName = $null }
    if (-not $PSBoundParameters.ContainsKey('PrimaryMailboxId')) { $PrimaryMailboxId = $null }
    if (-not $PSBoundParameters.ContainsKey('ArchiveMailboxId')) { $ArchiveMailboxId = $null }
}
$outRoot = if ([IO.Path]::IsPathRooted($config.Output.Root)) { $config.Output.Root } else { Join-Path (Get-Location).Path $config.Output.Root }
if (-not $StorePath) { $StorePath = Join-Path $outRoot 'transfer\store' }
if (-not $TempPath) { $TempPath = Join-Path $outRoot 'transfer' }
if (-not [IO.Path]::IsPathRooted($StorePath)) { $StorePath = Join-Path (Get-Location).Path $StorePath }

$pstMaxBytes = if ($PSBoundParameters.ContainsKey('PstMaxSizeMiB')) { $PstMaxSizeMiB * 1MB } else { [long]($PstMaxSizeGB * 1024) * 1MB }
$batchMode = [bool]($MailboxCsv -or $ResumeBatch)
if ($batchMode) {
    if ($Mode -notin 'Export', 'Import', 'ImportPst', 'ImportEml', 'Copy') { throw '-MailboxCsv and -ResumeBatch belong to -Mode Export, Import, ImportPst, ImportEml and Copy.' }
    if ($BatchName -and $ResumeBatch -and $BatchName.Trim() -ine $ResumeBatch.Trim()) { throw '-BatchName and -ResumeBatch name two batches: give one of them.' }
    foreach ($single in 'EventsFile', 'JobId', 'SelectionFile', 'MailboxId', 'TargetMailboxId', 'LocationsFile', 'PrimaryMailboxId', 'ArchiveMailboxId', 'RecoverableItemsFolders', 'TempPath', 'BatchId') {
        if ($PSBoundParameters.ContainsKey($single)) { throw "-$single is for one mailbox: with -MailboxCsv give it in the CSV file (SelectionFile, OutputPst), or leave it out. Each mailbox of a batch has its own job folder, store and job." }
    }
    if ($Mode -eq 'Copy') {
        foreach ($single in 'PstPath', 'EmlPath', 'JobName', 'IncludeFolder', 'ExcludeFolder', 'FolderFile', 'IncludeItemFile', 'ResolveEml', 'EmlFile', 'SourceMailbox') {
            if ($PSBoundParameters.ContainsKey($single)) { throw "-$single is for one copy: with -MailboxCsv give the mailboxes in the CSV file (the import job of each line is named after its target, its store and its folder), or leave it out." }
        }
        if ($Mailbox -in 'AuxArchive', 'All') { throw "-Mailbox $Mailbox is not for a batch of copies: a copy reads Primary or Archive (the SourceLocation column; -Mailbox is the SourceLocation of the lines that have none)." }
        if ($PSBoundParameters.ContainsKey('MappingFile') -and -not (Test-Path -LiteralPath $MappingFile -PathType Leaf)) { throw "-MappingFile not found: $MappingFile" }
    }
    elseif ($Mode -ne 'Export') {
        foreach ($single in 'StorePath', 'PstPath', 'EmlPath', 'JobName', 'IncludeFolder', 'ExcludeFolder', 'FolderFile', 'IncludeItemFile', 'ResolveEml', 'EmlFile') {
            if ($PSBoundParameters.ContainsKey($single)) { throw "-$single is for one import: with -MailboxCsv give the source in the Source column of the CSV file (the journal and the job name of each line are made from the line), or leave it out." }
        }
        if ($Mailbox -in 'AuxArchive', 'All') { throw "-Mailbox $Mailbox is not for a batch of imports: an import goes into Primary or Archive (the Location column; -Mailbox is the Location of the lines that have none)." }
        if ($PSBoundParameters.ContainsKey('SourceMailbox') -and $Mode -ne 'Import') { throw '-SourceMailbox chooses the Primary or the Archive of an export store: it belongs to -Mode Import.' }
        if ($PSBoundParameters.ContainsKey('Folder') -and $Mode -ne 'Import') { throw '-Folder chooses a folder of an export store: it belongs to -Mode Import.' }
        if ($PSBoundParameters.ContainsKey('MappingFile') -and -not (Test-Path -LiteralPath $MappingFile -PathType Leaf)) { throw "-MappingFile not found: $MappingFile" }
    }
    elseif ($Mailbox -eq 'AuxArchive') { throw '-Mailbox AuxArchive names one auxiliary archive: in a batch use Primary, Archive or All (Both in the CSV file) for the mailboxes.' }
    if ($CertificatePassword) { throw 'A batch starts a process for each mailbox and cannot hand a password to them: put the certificate in the store of the user and give -CertificateThumbprint.' }
    if ($DeviceCode -and $NoPrompt) { throw '-DeviceCode asks for a sign-in and -NoPrompt forbids it: give one of them.' }
}
elseif ($BatchName -or $PSBoundParameters.ContainsKey('ParallelMailboxes')) { throw '-BatchName and -ParallelMailboxes belong to -MailboxCsv (a batch of mailboxes).' }
if ($OutputPst -and $Mode -notin 'Export', 'ExportPst') { throw '-OutputPst belongs to -Mode Export (export, then convert) and -Mode ExportPst (convert a store).' }
if (($IncludeRecoverableItems -or $RecoverableItemsFolders) -and $Mode -notin 'Export', 'Browse', 'Copy') { throw '-IncludeRecoverableItems and -RecoverableItemsFolders belong to -Mode Export and -Mode Copy (and -Mode Browse, which has -RecoverableItems).' }
if ($RecoverableItemsFolders) {
    Initialize-MbtEngine
    [MailboxTransfer.RecoverableItems]::CheckNames([string[]]$RecoverableItemsFolders)
    if ($Mode -in 'Export', 'Copy' -and -not $IncludeRecoverableItems -and -not $SelectionFile -and -not [MailboxTransfer.RecoverableItems]::IsPath($Folder)) { throw '-RecoverableItemsFolders needs -IncludeRecoverableItems (or -SelectionFile that selects Recoverable Items).' }
}
if ($Mode -in 'Export', 'Copy' -and $IncludeRecoverableItems) {
    if ($SelectionFile) { throw '-IncludeRecoverableItems adds Recoverable Items to the export of the whole mailbox. With -SelectionFile, select the folders of Recoverable Items in the browse view (Browse -RecoverableItems, or the console) instead.' }
    Initialize-MbtEngine
    if ($Folder -and -not [MailboxTransfer.RecoverableItems]::IsPath($Folder)) { throw "-IncludeRecoverableItems adds Recoverable Items to the export of the whole mailbox; it cannot be combined with -Folder '$Folder'. Use -Folder 'Recoverable Items/Deletions' for the deleted items alone." }
}

# The period of an export (-ArrivedFrom, -ArrivedTo): read here, before the sign-in, so that a date that cannot be read or a start after the end stops everything at once.
$arrivalPeriod = $null
if ($PSBoundParameters.ContainsKey('ArrivedFrom') -or $PSBoundParameters.ContainsKey('ArrivedTo')) {
    if ($Mode -eq 'ExportPst') { throw '-ArrivedFrom and -ArrivedTo choose the items an export lists in a mailbox. -Mode ExportPst only converts what a store or a PST already holds: give the period to the -Mode Export that fills the store (with -OutputPst it writes the PST files in the same command).' }
    if ($Mode -notin 'Export', 'ExportEml', 'Copy') { throw '-ArrivedFrom and -ArrivedTo belong to the export modes (-Mode Export, -Mode Copy, and -Mode ExportEml from the live mailbox): an import takes every item its source holds. Choose the period when the items are exported.' }
    $arrivalPeriod = ConvertTo-MbtArrivalPeriod -From $ArrivedFrom -To $ArrivedTo
    if ($arrivalPeriod -and $Mode -eq 'ExportEml') {
        $emlSourceAsked = if ($EmlSource) { $EmlSource } elseif ($PstPath) { 'Pst' } elseif ($PSBoundParameters.ContainsKey('StorePath')) { 'Store' } else { 'Live' }
        if ($emlSourceAsked -ne 'Live') { throw '-ArrivedFrom and -ArrivedTo apply to -Mode ExportEml from the live mailbox: a store or a PST holds no date of arrival. Give the period to the -Mode Export that fills the store.' }
    }
}

# A copy (-Mode Copy) reads ONE mailbox location, an export into a store, and writes that store into ONE other mailbox location, an import, in one job.
# The source side is given as for an export (-UserPrincipalName, -Mailbox, -MailboxId, -SelectionFile, -Folder, the period), the target side as for an import (-TargetMailbox,
# -TargetMailboxId, -TargetRoot, -MappingFile). Everything that can be refused without signing in is refused here, before a banner or a request.
$copyWork = if ($WorkFolder) { [IO.Path]::GetFullPath($WorkFolder) } else { Get-MbtDefaultWorkFolder }
$copy = $null
if ($Mode -eq 'Copy' -and -not $batchMode) {
    foreach ($notCopy in 'SourceMailbox', 'PstPath', 'EmlPath', 'IncludeFolder', 'ExcludeFolder', 'FolderFile', 'IncludeItemFile') {
        if (-not $PSBoundParameters.ContainsKey($notCopy)) { continue }
        if ($notCopy -eq 'SourceMailbox') { throw '-SourceMailbox chooses the Primary or the Archive of an export store: it belongs to -Mode Import. A copy takes its source from -Mailbox (and -MailboxId): the mailbox it reads.' }
        throw "-$notCopy is not for -Mode Copy: a copy reads a mailbox. Choose what to copy with -Folder or -SelectionFile."
    }
    if ($Mailbox -in 'All', 'AuxArchive') { throw "-Mailbox $Mailbox is not for a copy: a copy reads one mailbox location, Primary or Archive (the auxiliary archives of an auto-expanding archive follow the Archive by themselves). Copy the other locations one after the other." }
    if ($SelectionFile -and $Folder) { throw '-SelectionFile and -Folder are exclusive: the selection names its folders.' }
    Initialize-MbtEngine
    $copy = [ordered]@{ SourceKind = [string]$Mailbox; SourceId = ''; SourceAddress = ''; TargetKind = [string]$TargetMailbox; TargetId = ''; TargetAddress = '' }
    if ($SelectionFile) {
        $selections = @(Read-MbtBrowseSelectionSet -Path $SelectionFile)
        if ($selections.Count -ne 1) { throw "The selection $SelectionFile names $($selections.Count) mailbox locations: a copy takes one (Primary or Archive). Select what to copy in one mailbox location, or copy the locations one after the other." }
        if ([string]$selections[0].Kind -notin 'Primary', 'Archive') { throw "The selection $SelectionFile is of the $($selections[0].Kind) mailbox: a copy reads the Primary or the Archive mailbox (the auxiliary archives follow the Archive)." }
        if ($mailboxGiven -and $Mailbox -ne [string]$selections[0].Kind) { throw "-Mailbox $Mailbox does not match the selection, which is of the $($selections[0].Kind) mailbox." }
        $copy.SourceKind = [string]$selections[0].Kind
        $copy.SourceId = ConvertTo-MbtMailboxId -Value $selections[0].MailboxId -TenantId $TenantId
    }
    elseif ($MailboxId -and -not (Test-MbtMailboxAddress $MailboxId)) { $copy.SourceId = ConvertTo-MbtMailboxId -Value $MailboxId -TenantId $TenantId }
    $copy.SourceAddress = if ($MailboxId -and (Test-MbtMailboxAddress $MailboxId)) { $MailboxId.Trim() } elseif ($MailboxId -or $SelectionFile) { '' } else { [string]$UserPrincipalName }
    if ($TargetMailboxId -and (Test-MbtMailboxAddress $TargetMailboxId)) { $copy.TargetAddress = $TargetMailboxId.Trim() }
    elseif ($TargetMailboxId) { $copy.TargetId = ConvertTo-MbtMailboxId -Value $TargetMailboxId -TenantId $TenantId }
    elseif (-not $useCertificate) { $copy.TargetAddress = [string]$UserPrincipalName }
    if ($useCertificate) {
        if (-not $TargetMailboxId) { throw 'The application has no mailbox of its own: give the mailbox to write into with -TargetMailboxId (its address).' }
        if (-not $copy.SourceAddress -and -not $UserPrincipalName) { throw 'Give the mailbox to read: -UserPrincipalName (its address), or -MailboxId.' }
        if (-not $UserPrincipalName) { $UserPrincipalName = $copy.SourceAddress }
    }
    elseif ($copy.TargetAddress -and $UserPrincipalName -and $copy.TargetAddress -ine $UserPrincipalName) {
        throw "A delegated sign-in writes into the mailbox of the signed-in account only ($UserPrincipalName): $($copy.TargetAddress) is another mailbox, and the service refuses a delegated write into another person's or a shared mailbox. Nothing was started, so nothing was read: sign in as the application (-AppOnly) to write into $($copy.TargetAddress), or copy into $UserPrincipalName."
    }
    if ($copy.SourceKind -ceq $copy.TargetKind -and (($copy.SourceAddress -and $copy.SourceAddress -ieq $copy.TargetAddress) -or ($copy.SourceId -and $copy.SourceId -ieq $copy.TargetId))) {
        throw "The source and the target are the same mailbox location (the $($copy.SourceKind) mailbox of $(if ($copy.SourceAddress) { $copy.SourceAddress } else { $copy.SourceId })): a copy writes into another location. Choose another -TargetMailbox, another -TargetMailboxId, or another source."
    }
    if ($MappingFile -and -not (Test-Path -LiteralPath $MappingFile -PathType Leaf)) { throw "-MappingFile not found: $MappingFile" }
    # the store of a copy is named after what it holds, so that another choice never lands in it
    $copy.ChoiceKey = Get-MbtCopyChoiceKey -Kind $copy.SourceKind -SelectionFile $SelectionFile -Folder $Folder -Period $arrivalPeriod -IncludeRecoverableItems:$IncludeRecoverableItems -RecoverableItemsFolders $RecoverableItemsFolders
    if (-not $PSBoundParameters.ContainsKey('StorePath')) {
        $StorePath = Get-MbtCopyStorePath -WorkFolder $copyWork -Source $(if ($copy.SourceAddress) { $copy.SourceAddress } elseif ($copy.SourceId) { $copy.SourceId } else { 'mailbox' }) -Kind $copy.SourceKind -Key $copy.ChoiceKey
    }
}
elseif ($PSBoundParameters.ContainsKey('TargetMailbox') -or $PSBoundParameters.ContainsKey('TargetMailboxId')) { throw '-TargetMailbox and -TargetMailboxId belong to -Mode Copy (the mailbox a copy writes into). The mailbox of an import is -Mailbox and -MailboxId.' }

# The frames and the icons of the console are Unicode; a redirected output (a job of the console, a scheduled task) is UTF-8 without BOM.
try { [Console]::OutputEncoding = [Text.UTF8Encoding]::new($false) } catch { Write-Verbose "Console output encoding unchanged: $($_.Exception.Message)" }
if (-not $LogFolder) { $LogFolder = if ($env:MAILBOXTRANSFER_LOG_FOLDER) { $env:MAILBOXTRANSFER_LOG_FOLDER } else { Join-Path $outRoot 'logs' } }
$logPath = $null
try { $logPath = Start-MbtLog -Directory $LogFolder } catch { Write-Warning "No log file ($LogFolder): $($_.Exception.Message)" }
$runClock = [Diagnostics.Stopwatch]::StartNew()
$bannerSubtitle = switch ($Mode) {
    'Export' { if ($batchMode) { 'Exports the mailboxes of a list, one job each' } else { 'Exports a mailbox to a verified local store, then to PST files if asked' } }
    'ExportPst' { 'Writes PST files from an export store, or from chosen folders of a PST' }
    'Import' { if ($batchMode) { 'Imports export stores into the mailboxes of a list, one job each (items are added, never deleted)' } else { 'Puts the items of an export store into a mailbox (items are added, never deleted)' } }
    'ImportPst' { if ($batchMode) { 'Imports PST files into the mailboxes of a list, one job each (items are added, never deleted)' } else { 'Puts the items of a PST file into a mailbox (items are added, never deleted)' } }
    'Copy' { if ($batchMode) { 'Copies mailboxes into other mailboxes from a list, one job each (items are added, never deleted)' } else { 'Reads a mailbox into a store, then imports it into another mailbox' } }
    'ExportEml' { 'Writes the items of a store as .eml files' }
    'ImportEml' { if ($batchMode) { 'Imports folders of .eml files into the mailboxes of a list, one job each (items are added, never deleted)' } else { 'Puts .eml files into a mailbox (items are added, never deleted)' } }
    'Browse' { 'Looks inside a mailbox, read only' }
    'Status' { 'Shows the state of an export store and of its import jobs' }
    'CheckPst' { 'Checks a PST file, read only' }
    'RepairPst' { 'Writes a repaired copy of a PST file; the original is never changed' }
}
# A single job started from the command line is a job like the ones of the console: its folder (events, status.json, request.json with this command and the folder it runs in) is in the work folder,
# so that the console lists it and -ResumeJob continues it. The same command again continues the job of that command while it is not complete.
$cliJob = $null
if (-not $EventsFile -and -not $batchMode -and -not $dryRun -and $Mode -in 'Export', 'ExportPst', 'Import', 'ImportPst', 'ImportEml', 'ExportEml', 'Copy') {
    try {
        $cliWork = if ($WorkFolder) { [IO.Path]::GetFullPath($WorkFolder) } else { Get-MbtDefaultWorkFolder }
        $cliTask = switch ($Mode) { 'Copy' { 'mailcopy' } 'Export' { 'export' } 'ExportEml' { 'export' } 'ExportPst' { 'exportpst' } default { 'import' } }
        $cliParameters = @{}
        foreach ($k in $boundParameters.Keys) { $cliParameters[$k] = $boundParameters[$k] }
        $cliResolved = @{ StorePath = [IO.Path]::GetFullPath($StorePath); TempPath = [IO.Path]::GetFullPath($TempPath) }
        if ($Mode -eq 'ImportPst') { $cliResolved.JournalPath = Join-Path $cliResolved.TempPath 'pstimport.db' }
        $cliTitle = '{0} - {1}' -f $Mode, @($UserPrincipalName, $PstPath, $EmlPath, $StorePath | Where-Object { $_ } | ForEach-Object { [string]@($_)[0] })[0]
        $cliJob = New-MbtCliJob -WorkFolder $cliWork -Mode $Mode -Task $cliTask -Parameter $cliParameters -WorkingDirectory (Get-Location).ProviderPath -ScriptPath $PSCommandPath -Resolved $cliResolved -Title $cliTitle
        $EventsFile = $cliJob.EventsFile
        $JobId = $cliJob.JobId
    } catch { Write-Warning "The job folder could not be written ($($_.Exception.Message)): the job runs without it, and the console will not list it."; $cliJob = $null }
}
$bannerDetails = [ordered]@{ Mode = @('Mode', ($Mode + $(if ($batchMode) { ' - a list of mailboxes' }) + $(if ($dryRun) { ' - WhatIf (nothing is written)' }))) }
if ($batchMode) { $bannerDetails.Mailboxes = @('Folder', $(if ($MailboxCsv) { $MailboxCsv } else { "batch $ResumeBatch" })) }
elseif ($Mode -in 'Export', 'ExportPst', 'Import', 'Status', 'ExportEml', 'Copy') { $bannerDetails.Store = @('Store', $StorePath) }
if ($Mode -eq 'Copy' -and -not $batchMode) {
    $bannerDetails.From = @('Folder', ('{0} mailbox {1}' -f $copy.SourceKind, $(if ($copy.SourceAddress) { $copy.SourceAddress } elseif ($copy.SourceId) { $copy.SourceId } else { 'of the selection' })))
    $bannerDetails.To = @('Folder', ('{0} mailbox {1}{2}' -f $copy.TargetKind, $(if ($copy.TargetAddress) { $copy.TargetAddress } else { $copy.TargetId }), $(if ($TargetRoot) { ", folder $TargetRoot" } elseif ($MappingFile) { ', mapping file' } else { ', same folder names' })))
}
$bannerSource = @($PstPath, $EmlPath, $Path | Where-Object { $_ } | ForEach-Object { [string]$_ }) -join ', '
if ($bannerSource) { $bannerDetails.Source = @('Folder', $bannerSource) }
if ($OutputPst) { $bannerDetails.Output = @('Folder', $OutputPst) }
if ($arrivalPeriod) { $bannerDetails.Period = @('Clock', $arrivalPeriod.Describe()) }
if ($Destination) { $bannerDetails.Output = @('Folder', $Destination) }
if ($cliJob) { $bannerDetails.Job = @('Folder', ($cliJob.Name + $(if ($cliJob.Reused) { ' (the same command again: this job continues)' }))) }
if ($logPath) { $bannerDetails.Log = @('Log', $logPath) }
Write-MbtBanner -Title 'Mailbox Transfer' -Subtitle $bannerSubtitle -Details $bannerDetails
Initialize-MbtEngine
$script:counts = @{}
$script:reason = $null
$script:next = $null
$script:stateOverride = $null
$overviewBudgetSeconds = 20
$script:sawResult = $false
$script:resultComplete = $false
$script:phase = $Mode
$script:resumable = $false
$script:copyExportStopped = $false
$script:copyNext = $null
# The name of the job folder of a run that has one made by the command line: its Next steps say how to continue it.
function Get-CliJobName {
    if ($cliJob) { return [string]$cliJob.Name }
    if ($EventsFile) {
        $jobFolder = Split-Path ([IO.Path]::GetFullPath($EventsFile)) -Parent
        $request = Read-MbtJobRequest -Folder $jobFolder
        if ($request -and [string](Get-MbtJobField $request 'Source' '') -eq 'cli') { return (Split-Path $jobFolder -Leaf) }
    }
    return ''
}
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
    $parts = @(foreach ($k in ($script:counts.Keys | Sort-Object)) { if ($arrivalPeriod -and $k -in 'kept', 'outsidePeriod', 'withoutDate') { continue }; $v = $script:counts[$k]; if ($v -is [ValueType] -or $v -is [string]) { '{0} {1}' -f $k, $(if ($v -is [long] -or $v -is [int]) { Format-MbtNumber $v } else { $v }) } })
    if ($parts) { $values.Counts = @('Store', ($parts -join ', ')) }
    if ($arrivalPeriod -and $script:counts.ContainsKey('outsidePeriod')) {
        $periodNote = "$($arrivalPeriod.Describe()): $(Format-MbtNumber $script:counts['kept']) kept, $(Format-MbtNumber $script:counts['outsidePeriod']) outside the period (skipped, not an error)"
        if ([long]$script:counts['withoutDate'] -gt 0) { $periodNote += "; $(Format-MbtNumber $script:counts['withoutDate']) without a date were kept" }
        $values.Period = @('Clock', $periodNote)
    }
    $values.Duration = @('Clock', (Format-MbtDuration $runClock.Elapsed.TotalSeconds))
    if ($script:next -and $script:next -notin 'done') { $values.Next = @('Plan', [string]$script:next) }
    $cliName = Get-CliJobName
    $nextSteps = @($script:copyNext | Where-Object { $_ })
    if ($cliName -and -not $dryRun -and $ExitCode -ne 0 -and -not $script:resultComplete) {
        $nextSteps += @("Continue this job: .\Invoke-MailboxTransfer.ps1 -ResumeJob $cliName", 'Running the same command again continues it as well. The console lists it under Continue a job.')
    }
    if ($nextSteps.Count) { $values['Next steps'] = @('Plan', $nextSteps) }
    if ($cliName) { $values.Job = @('Folder', $cliName) }
    if ($logPath) { $values.Log = @('Log', $logPath) }
    # the report of the job (steps and chronology, waits, warnings, folders): one HTML file in the job folder, written from the events just written
    if ($EventsFile -and -not $dryRun -and $Mode -in 'Export', 'ExportPst', 'Import', 'ImportPst', 'ExportEml', 'ImportEml', 'CheckPst', 'RepairPst', 'Copy') {
        try { $values.Report = @('Report', [string](Write-MbtJobHtmlReport -Folder (Split-Path ([IO.Path]::GetFullPath($EventsFile)) -Parent))) }
        catch { Write-MbtItem WARN "the HTML report could not be written: $($_.Exception.Message)" }
        # the overview of all the jobs of the work folder (<work folder>\jobs\overview.html): this job is one more row of it
        if (-not $BatchId) {
            try {
                $overviewInfo = @{}
                $overviewPath = Write-MbtJobOverview -JobFolder (Split-Path ([IO.Path]::GetFullPath($EventsFile)) -Parent) -TimeBudgetSeconds $overviewBudgetSeconds -Info $overviewInfo
                if ($overviewPath) { $values['All jobs'] = @('Report', [string]$overviewPath) }
                elseif ($overviewInfo.Contains('Complete') -and -not $overviewInfo.Complete) { Write-MbtItem WARN "the overview of all the jobs was not finished in $overviewBudgetSeconds s ($(Format-MbtNumber $overviewInfo.Unread) of $(Format-MbtNumber $overviewInfo.Folders) job folders are still to read): the next job or the console finishes it" }
            }
            catch { Write-MbtItem WARN "the overview of all the jobs could not be written: $($_.Exception.Message)" }
        }
    }
    Write-MbtSummary -Title $title -Values $values -Status $status
}

# The folders of the jobs that are finished and older than Jobs.RetentionDays (90 days; 0 keeps everything) are removed before this job starts; a job that can still be continued never is.
if ($EventsFile -and -not $dryRun -and -not $BatchId) {
    try {
        $purge = Invoke-MbtJobRetention -JobFolder (Split-Path ([IO.Path]::GetFullPath($EventsFile)) -Parent) -RetentionDays $config.Jobs.RetentionDays -OverviewBudgetSeconds $overviewBudgetSeconds
        if ($purge -and @($purge.Removed).Count) { Write-MbtItem INFO "removed $(Format-MbtNumber @($purge.Removed).Count) old job $(if (@($purge.Removed).Count -eq 1) { 'folder' } else { 'folders' }) (finished more than $($purge.RetentionDays) days ago); the names are in the log" }
        if ($purge -and @($purge.Failed).Count) { Write-MbtItem WARN "$(Format-MbtNumber @($purge.Failed).Count) old job $(if (@($purge.Failed).Count -eq 1) { 'folder' } else { 'folders' }) could not be removed (something holds a file); the next job tries again" }
    } catch { Write-MbtItem WARN "the old job folders could not be removed: $($_.Exception.Message)" }
}
if ($StopFile -and -not $batchMode) { Start-MbtStopWatch -Path $StopFile }
if ($KeepAwake) { [void](Start-MbtKeepAwake) }
$copyText = if ($copy) { @{ Source = $(if ($copy.SourceAddress) { $copy.SourceAddress } elseif ($copy.SourceId) { $copy.SourceId } else { 'the selection' }); Target = $(if ($copy.TargetAddress) { $copy.TargetAddress } else { $copy.TargetId }) } }
if ($EventsFile) {
    $taskName = switch ($Mode) { 'Copy' { 'mailcopy' } 'Export' { 'export' } 'ExportEml' { 'export' } 'ExportPst' { 'exportpst' } 'Import' { 'import' } 'ImportPst' { 'import' } 'ImportEml' { 'import' } 'CheckPst' { 'check' } 'RepairPst' { 'repair' } 'Browse' { 'browse' } default { 'status' } }
    $title = switch ($Mode) {
        'Export' { "Export - $UserPrincipalName" } 'ExportPst' { "Write PST files - $OutputPst" } 'Import' { "Import - $UserPrincipalName" }
        'ImportPst' { "Import from PST - $UserPrincipalName" } 'Copy' { "Copy - $($copyText.Source) to $($copyText.Target)" } 'ExportEml' { "Export EML files - $EmlPath" } 'ImportEml' { "Import EML files - $UserPrincipalName" }
        'CheckPst' { "Check - $Path" } 'RepairPst' { "Repair - $Path" } default { $Mode }
    }
    $mailboxText = if ($MailboxId) { $MailboxId } else { $UserPrincipalName }
    $mailboxText = ($mailboxText + $(if ($Mailbox -and $Mailbox -ne 'Primary') { " ($Mailbox)" } else { '' })).Trim()
    $sourceText = switch ($Mode) {
        'Export' { $mailboxText } 'ExportPst' { $(if ($PstPath) { $PstPath } else { $StorePath }) } 'Import' { $StorePath } 'ImportPst' { $PstPath } 'ExportEml' { $(if ($PstPath) { $PstPath } elseif ($StorePath) { $StorePath } else { $mailboxText }) }
        'ImportEml' { $EmlPath } 'CheckPst' { $Path } 'RepairPst' { $Path } 'Copy' { "$($copyText.Source) ($($copy.SourceKind))" } default { '' }
    }
    $destinationText = switch ($Mode) {
        'Export' { $StorePath + $(if ($OutputPst) { ", PST files: $OutputPst" } else { '' }) } 'ExportPst' { $OutputPst } 'Import' { $mailboxText + $(if ($TargetRoot) { ", folder $TargetRoot" } else { '' }) }
        'ImportPst' { $mailboxText + $(if ($TargetRoot) { ", folder $TargetRoot" } else { '' }) } 'ExportEml' { $EmlPath } 'ImportEml' { $mailboxText + $(if ($TargetRoot) { ", folder $TargetRoot" } else { '' }) } 'RepairPst' { $Destination }
        'Copy' { "$($copyText.Target) ($($copy.TargetKind))" + $(if ($TargetRoot) { ", folder $TargetRoot" } else { '' }) } default { '' }
    }
    Open-MbtEventLog -Path $EventsFile -JobId $JobId -Task $taskName -Title $title -Mode $Mode -Store $(if ($Mode -in 'Export', 'ExportPst', 'Import', 'Copy') { $StorePath }) -Source ([string]$sourceText) -Destination ([string]$destinationText) -Batch $BatchId -Period $arrivalPeriod
    $steps = switch ($Mode) {
        'Export' {
            @(@{ Id = 'tree'; Title = 'Read the folder tree' }, @{ Id = 'list'; Title = 'List the items you chose' }, @{ Id = 'download'; Title = 'Download the items' }, @{ Id = 'verify'; Title = 'Check the downloads' }) +
            $(if ($OutputPst) { @(@{ Id = 'pst'; Title = 'Write the PST files' }) + $(if (-not $SkipCheck) { @(@{ Id = 'pstcheck'; Title = 'Check every PST file' }) }) })
        }
        'ExportPst' { @(@{ Id = 'pst'; Title = 'Write the PST files' }) + $(if (-not $SkipCheck) { @(@{ Id = 'pstcheck'; Title = 'Check every PST file' }) }) }
        'Copy' {
            @(@{ Id = 'tree'; Title = 'Read the folder tree' }, @{ Id = 'list'; Title = 'List the items you chose' }, @{ Id = 'download'; Title = 'Download the items' }, @{ Id = 'verify'; Title = 'Check the downloads' },
                @{ Id = 'check'; Title = 'Verify the export store' }, @{ Id = 'folders'; Title = 'Create the folders' }, @{ Id = 'send'; Title = 'Send the items' }, @{ Id = 'reconcile'; Title = 'Check the uncertain items' }, @{ Id = 'final'; Title = 'Final check' })
        }
        'Import' { @(@{ Id = 'check'; Title = 'Verify the export store' }, @{ Id = 'folders'; Title = 'Create the folders' }, @{ Id = 'send'; Title = 'Send the items' }, @{ Id = 'reconcile'; Title = 'Check the uncertain items' }, @{ Id = 'final'; Title = 'Final check' }) }
        'ImportPst' { @(@{ Id = 'check'; Title = 'Check the PST files' }, @{ Id = 'folders'; Title = 'Create the folders' }, @{ Id = 'send'; Title = 'Send the items' }, @{ Id = 'reconcile'; Title = 'Check the uncertain items' }, @{ Id = 'final'; Title = 'Final check' }) }
        'ExportEml' { @(@{ Id = 'export'; Title = 'Write the .eml files' }) }
        'ImportEml' { @(@{ Id = 'send'; Title = 'Import the messages' }) }
        'CheckPst' { @(@{ Id = 'check'; Title = 'Check the file' }) }
        'RepairPst' { @(@{ Id = 'before'; Title = 'Check the file (before)' }, @{ Id = 'scan'; Title = 'Scan and read what can be read' }, @{ Id = 'write'; Title = 'Write the new file' }, @{ Id = 'after'; Title = 'Check the new file (after)' }) }
        default { @(@{ Id = 'run'; Title = $title }) }
    }
    Write-MbtJobPlan -Steps $steps
    # the first step is under way from the start: the sign-in and the reading of the selection are part of it, so the console never shows a job that waits
    if ($Mode -in 'Export', 'Copy') { Write-MbtJobStep -Id tree -State running -Msg 'Signing in to Microsoft' }
}
$storeExists = Test-Path -LiteralPath (Join-Path $StorePath 'catalog.db')
$store = $null
$exitCode = 0

# App-only: the application signs in with its certificate (no user, no prompt). Nothing here goes through /me: -UserPrincipalName is the owner of the target mailbox.
function Connect-AppOnlyForTransfer {
    if ($DeviceCode) { throw '-DeviceCode is a delegated sign-in: it cannot be combined with a certificate (app-only).' }
    if (-not $UserPrincipalName) { throw 'App-only sign-in has no signed-in user: give -UserPrincipalName, the address of the target mailbox (its owner, or the shared mailbox itself). The user of the configuration is never assumed.' }
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
    Write-MbtItem INFO ('signing in as {0} ({1}); nothing is asked when a valid sign-in is cached' -f $UserPrincipalName, $(if ($DeviceCode) { 'device code: a line "DEVICE CODE: ..." follows' } else { 'Windows sign-in window' }))
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
    if ($Mode -in 'Import', 'ImportEml', 'Copy' -and -not ('MailboxFolder.ReadWrite' -in $t.Scopes)) { $missing += 'MailboxFolder.ReadWrite (creates the folders)' }
    if ($missing) { Write-MbtItem WARN ('missing in the token: ' + ($missing -join ', ') + '. An administrator grants these permissions to the application (admin consent); then sign in again.') }
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
    $step = if ($Mode -eq 'Copy') { $script:phase } else { $Mode }
    $which = if ($step -in 'Browse', 'Export', 'Import') { $step } else { 'Resolve' }
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

# A copy, after the sign-in and before anything is read: the id of the target mailbox location, whether it can be read, and that it is not the source. A copy that would be refused at the end
# (the target cannot be written, the target is the source) must be refused before hours of reading.
function Resolve-CopyTarget {
    $kind = $copy.TargetKind
    $id = $copy.TargetId
    if ($id) {
        if (-not $useCertificate -and (Test-OtherMailbox $id)) {
            throw "$id is not known as a mailbox of $UserPrincipalName. A delegated sign-in writes into the mailbox of the signed-in account only: give -TargetMailbox Primary or Archive (and -ArchiveMailboxId for the archive) instead of an id, or sign in as the application (-AppOnly). Nothing was read."
        }
    }
    else {
        $address = $copy.TargetAddress
        $arguments = @{ Mailbox = $kind; TenantId = $TenantId; UserPrincipalName = $address }
        if ($address -ieq $UserPrincipalName) { $arguments.PrimaryMailboxId = $PrimaryMailboxId; $arguments.ArchiveMailboxId = $ArchiveMailboxId; $arguments.LocationsFile = $LocationsFile }
        elseif ($kind -eq 'Primary') { $arguments.MailboxId = $address }
        $id = Resolve-MbtMailboxId @arguments
    }
    $copy.TargetId = $id
    try { $folders = @((Get-MbtSession).Api.GetFolderTree($id)) }
    catch { throw "The $kind mailbox $id cannot be read, so nothing can be written into it: $($_.Exception.Message). Nothing was read from the source." }
    Write-MbtItem OK "target: the $kind mailbox $(if ($copy.TargetAddress) { $copy.TargetAddress } else { $id }) ($(Format-MbtNumber $folders.Count) folders) can be reached"
}

# The source of a copy is never the target.
function Assert-CopySource([string]$SourceId) {
    if ($SourceId -ieq $copy.TargetId) {
        throw "The source and the target are the same mailbox location ($($copy.TargetKind) mailbox $SourceId): a copy writes into another location. Choose another -TargetMailbox, another -TargetMailboxId, or another source. Nothing was read."
    }
}

# -WhatIf of a copy: one plan for both halves, from the totals of what the export half would read.
function Write-CopyPlan {
    param([long]$Items, [double]$Bytes, [bool]$SizeUnknown)
    $what = if ($SelectionFile) { "your selection ($SelectionFile)" } elseif ($Folder) { "the folder $Folder and the folders below it" } else { 'the whole mailbox' }
    $arguments = @{
        Source = $(if ($copy.SourceAddress) { $copy.SourceAddress } else { $copy.SourceId }); SourceKind = $copy.SourceKind; What = $what
        Target = $(if ($copy.TargetAddress) { $copy.TargetAddress } else { $copy.TargetId }); TargetKind = $copy.TargetKind; TargetRoot = $TargetRoot; MappingFile = $MappingFile
        StorePath = $StorePath; Items = $Items; Bytes = $Bytes; SizeUnknown = $SizeUnknown; AtMost = [bool]$arrivalPeriod; AppOnly = [bool]$useCertificate; SignedIn = [string]$UserPrincipalName
    }
    if ($arrivalPeriod) { $arguments.Period = $arrivalPeriod.Describe() }
    Show-MbtCopyPlan @arguments
    Write-MbtPeriodPlanNote
    Write-MbtItem WARN 'WhatIf: nothing was exported, no store was created and nothing was written into the target mailbox.'
}

# -Mode Copy, before the export half reads anything: the source is not the target, the store is the store of this copy, and an export half that was finished earlier is not done again
# (-Force reads the source again: what changed or arrived since is taken, and only that is imported). Returns $true when the export half is already complete.
function Test-CopyExportDone {
    if ($copy.SourceId) { $sourceId = $copy.SourceId }
    else {
        $sourceId = Resolve-MbtMailboxId -Mailbox $copy.SourceKind -TenantId $TenantId -UserPrincipalName $UserPrincipalName -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -MailboxId $MailboxId -LocationsFile $LocationsFile
        $copy.SourceId = $sourceId
    }
    Assert-CopySource $sourceId
    if (-not $script:store) { $script:store = Open-MbtExportStore -Path $StorePath -Name 'export' -Source "$UserPrincipalName ($TenantId)" }
    Set-MbtCopyStoreKey -Store $store -Key (Get-CopyStoreKey) -StorePath $StorePath
    $done = $store.GetExportOption('copy.export')
    if (-not $done -or $Force) { return $false }
    $exported = [long](@($store.GetItemStats($sourceId) | Where-Object { [string]$_.State -eq 'Exported' } | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum)
    Write-MbtItem OK "the export half is already complete (finished $done, $(Format-MbtNumber $exported) items in the store): it is not read again. -Force reads the source again."
    foreach ($id in 'tree', 'list', 'download', 'verify') { Write-MbtJobStep -Id $id -State done -Msg 'Already done: the export store of this copy is complete.' }
    $script:counts = @{ done = $exported; failed = 0L; deferred = 0L; bytes = 0L }
    return $true
}

# The key that says a store is the store of this copy: what is read (the choice) and from which mailbox location.
function Get-CopyStoreKey {
    return ([Convert]::ToHexString([Security.Cryptography.SHA256]::HashData([Text.Encoding]::UTF8.GetBytes(($copy.ChoiceKey + '|' + $copy.SourceId.ToLowerInvariant())))).Substring(0, 16).ToLowerInvariant())
}

# After the export half of a copy: its counts under their own names, so that the import half and the end of the job add theirs.
function Convert-CopyExportCounts {
    $c = $script:counts
    $kept = @{ exported = [long]$c['done']; exportFailed = [long]$c['failed']; failed = [long]$c['failed']; deferred = [long]$c['deferred']; bytes = [long]$c['bytes']; done = 0L }
    foreach ($k in 'kept', 'outsidePeriod', 'withoutDate') { if ($c.ContainsKey($k)) { $kept[$k] = [long]$c[$k] } }
    $script:counts = $kept
}

# What to do next after a copy, by its exit code.
function Get-CopyNext([int]$ExitCode) {
    $target = "$($copy.TargetKind) mailbox $(if ($copy.TargetAddress) { $copy.TargetAddress } else { $copy.TargetId })"
    if ($dryRun) { return @('Run the same command without -WhatIf to start the copy. Nothing was read or written by this plan.') }
    switch ($ExitCode) {
        0 {
            return @("Open the $target in Outlook and look in $(if ($TargetRoot) { "the folder '$TargetRoot'" } else { 'the folders of the same names' }): everything that was there before is unchanged.",
                "The export store $StorePath is kept: it is the proof of what was read and the resume point. Open it in the console, or import it again with -Mode Import (nothing is created twice). Delete the folder when you no longer need it.",
                'Running the same command again adds nothing: an item that was copied is never created twice. -Force reads the source again to take what arrived since.')
        }
        3 { return @("Some items may already exist in the target: settle them with -Mode Status -StorePath '$StorePath' -Unknown -Check, then run the same command again.") }
        2 { return @('Run the same command again to continue: an export half that is complete is not read again, and an item that was copied is never created twice.') }
        default { return @('Read the error above, then run the same command again: what was read is kept in the store and what was copied is never created twice.') }
    }
}

# -Mode Copy, between the two halves: from here the script speaks as an import. The export store is the source (its Primary or Archive), the target location is the mailbox.
function Enter-CopyImport {
    $script:storeExists = $true
    $script:SourceMailbox = [string]$copy.SourceKind
    $script:Mailbox = [string]$copy.TargetKind; $script:mailboxGiven = $true
    $script:MailboxId = [string]$copy.TargetId
    if (-not $boundParameters.ContainsKey('JobName')) { $script:JobName = Get-MbtCopyJobName -Target $copy.TargetId -TargetKind $copy.TargetKind -StorePath $StorePath -TargetRoot $TargetRoot -SourceKind $copy.SourceKind }
}

# What a plan (-WhatIf) of an export with a period says about its numbers: the folder totals are the most the export can take; the period is applied while the items are listed.
function Write-MbtPeriodPlanNote {
    if (-not $arrivalPeriod) { return }
    $note = "period ($($arrivalPeriod.Describe())): the numbers above are the totals of the folders, the most this export can take. The period is applied while the items are listed during the export; the end of the export gives the exact numbers (items kept, items outside the period)."
    Write-MbtItem INFO $note
    Write-MbtEvent info ([ordered]@{ msg = $note })
}

# -Mode Status: the jobs of the work folder (the ones the console lists), the ones that are not finished first, with why they stopped and the command that continues them.
function Show-JobList {
    $listWork = if ($WorkFolder) { [IO.Path]::GetFullPath($WorkFolder) } else { Get-MbtDefaultWorkFolder }
    Write-MbtStep "Jobs in $(Join-Path $listWork 'jobs')"
    $listing = Get-MbtJobListing -WorkFolder $listWork -FinishedLimit 10
    if (-not @($listing.Unfinished).Count -and -not @($listing.Finished).Count) { Write-MbtItem INFO 'no job yet: a job started from the command line or from the console is listed here.'; return }
    foreach ($j in @($listing.Unfinished) + @($listing.Finished)) {
        $icon = switch ($j.Kind) { 'complete' { 'OK' } 'running' { 'INFO' } 'error' { 'FAIL' } default { 'WARN' } }
        Write-MbtItem $icon ('{0}  {1}' -f $j.Name, $(if ($j.Title) { $j.Title } else { $j.JobId }))
        if ($j.Why) { Write-MbtDetail $j.Why }
        if ($j.Progress) { Write-MbtDetail $j.Progress }
        if ($j.Kind -notin 'complete', 'running') { Write-MbtDetail "continue: -ResumeJob $($j.Name)" }
    }
    if ($listing.FinishedTotal -gt @($listing.Finished).Count) { Write-MbtItem INFO ('{0} more finished jobs: the overview of all the jobs (jobs\overview.html) lists them.' -f ($listing.FinishedTotal - @($listing.Finished).Count)) }
    Write-MbtDetail 'continue one with: -ResumeJob <name>; delete one with: -RemoveJob <name> (-WhatIf shows what goes)'
}

# The export of a list of mailboxes, the import into a list of mailboxes, or the copy of a list of mailboxes into others (-MailboxCsv, or -ResumeBatch for a batch that exists). The whole CSV file is checked first and every problem is printed together;
# then each mailbox (each line of an import) is an ordinary job in a process of its own (Invoke-MbtBatch). The sign-in is done once here, and never asked for again in the middle of a run.
function Invoke-BatchRun {
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $work = if ($WorkFolder) { [IO.Path]::GetFullPath($WorkFolder) } else { Get-MbtDefaultWorkFolder }
    $stop = { param([string[]]$Problems)
        Write-MbtStep 'The batch cannot start'
        foreach ($p in $Problems) { Write-MbtItem FAIL $p }
        $count = $Problems.Count
        Write-MbtBatchSummary -ExitCode 1 -Failure @("$count $(if ($count -eq 1) { 'problem' } else { 'problems' }), listed above: nothing was started") -Log $logPath -Mode $Mode
        $script:exitCode = 1
    }
    $name = if ($ResumeBatch) { $ResumeBatch.Trim() } elseif ($BatchName) { $BatchName.Trim() } else { '' }
    if ($name) { $why = Test-MbtBatchName $name; if ($why) { & $stop @($why); return } }
    $folder = if ($name) { Find-MbtBatchFolder -WorkFolder $work -Name $name } else { $null }
    $batch = if ($folder) { Read-MbtBatch -Folder $folder } else { $null }
    if ($folder -and -not $batch) { & $stop @("The folder $folder has no readable batch.json: it is not a batch of this tool, or it is damaged. Use another -BatchName."); return }
    if ($batch -and (Get-MbtBatchMode $batch) -ne $Mode) { & $stop @("The batch $($batch.id) was made for -Mode $(Get-MbtBatchMode $batch), this command says -Mode $Mode. Continue it with -Mode $(Get-MbtBatchMode $batch), or use another -BatchName."); return }
    if ($ResumeBatch -and -not $batch) { & $stop @("There is no batch named $name in $(Join-Path $work 'batches'). Check the name, or -WorkFolder."); return }
    if (-not $MailboxCsv -and -not $batch) { & $stop @('Give -MailboxCsv (the list of mailboxes) or -ResumeBatch (a batch that exists).'); return }
    $id = if ($batch) { [string]$batch.id } elseif ($name) { $name } else { (Get-Date).ToString('yyyyMMdd-HHmm') + '-' + (-join ((1..4) | ForEach-Object { '0123456789abcdef'[(Get-Random -Maximum 16)] })) }

    # the options of the run: what the first run kept, over which what is given now
    $stored = if ($batch -and $batch.Contains('options') -and $batch.options -is [System.Collections.IDictionary]) { $batch.options } else { @{} }
    $storedPass = if ($stored.Contains('Pass') -and $stored.Pass -is [System.Collections.IDictionary]) { $stored.Pass } else { @{} }
    $pass = @{}
    foreach ($k in $storedPass.Keys) { $pass[[string]$k] = $storedPass[$k] }
    $valueOptions = switch ($Mode) {
        'Export' { 'Concurrency', 'MaxAttempts', 'RetryWaitSeconds', 'NotReadyTrip', 'BatchItems', 'BatchMiB', 'IntervalSeconds', 'Folder', 'PstMaxSizeGB', 'PstMaxSizeMiB', 'ConfigPath' }
        'Import' { 'Concurrency', 'MaxAttempts', 'RetryWaitSeconds', 'IntervalSeconds', 'Folder', 'ConfigPath', 'ReconcileDelaySeconds', 'ReconcileRounds' }
        'Copy' { 'Concurrency', 'MaxAttempts', 'RetryWaitSeconds', 'NotReadyTrip', 'BatchItems', 'BatchMiB', 'IntervalSeconds', 'Folder', 'ConfigPath', 'ReconcileDelaySeconds', 'ReconcileRounds' }
        'ImportPst' { 'Concurrency', 'MaxAttempts', 'RetryWaitSeconds', 'IntervalSeconds', 'ConfigPath', 'ReconcileDelaySeconds', 'ReconcileRounds', 'RepairLevel' }
        default { 'Concurrency', 'MaxAttempts', 'RetryWaitSeconds', 'IntervalSeconds', 'ConfigPath', 'MimeConversion' }
    }
    foreach ($n in $valueOptions) {
        if ($boundParameters.ContainsKey($n)) { $pass[$n] = (Get-Variable -Name $n -ValueOnly) }
    }
    $switchOptions = switch ($Mode) {
        'Export' { 'SkipInventory', 'IncludeRecoverableItems', 'SeparateAuxArchives', 'SkipCheck', 'Force' }
        'Import' { 'NoAutoReconcile' }
        'Copy' { 'SkipInventory', 'IncludeRecoverableItems', 'NoAutoReconcile', 'Force' }
        'ImportPst' { 'NoAutoReconcile', 'MergeWellKnown', 'PerFileSubfolder', 'NoCheck', 'Repair', 'KeepDamagedAttachments' }
        default { 'Unread', 'NoResume' }
    }
    foreach ($n in $switchOptions) {
        if ($boundParameters.ContainsKey($n)) { $pass[$n] = [bool](Get-Variable -Name $n -ValueOnly) }
    }
    if ($Mode -ne 'Export' -and $boundParameters.ContainsKey('MappingFile') -and $Mode -ne 'ImportEml') { $pass['MappingFile'] = [IO.Path]::GetFullPath($MappingFile) }
    # the period applies to every mailbox of the batch: the children get the bounds as written (ISO text), the same ones again when the batch is continued
    foreach ($n in 'ArrivedFrom', 'ArrivedTo') {
        if ($boundParameters.ContainsKey($n)) {
            $v = Get-Variable -Name $n -ValueOnly
            $pass[$n] = if ($v -is [datetime]) { $v.ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff", [Globalization.CultureInfo]::InvariantCulture) } elseif ($null -eq $v) { '' } else { ([string]$v).Trim() }
        }
    }
    if ($batch -and ($boundParameters.ContainsKey('ArrivedFrom') -or $boundParameters.ContainsKey('ArrivedTo'))) {
        $wasPeriod = ConvertTo-MbtArrivalPeriod -From $storedPass['ArrivedFrom'] -To $storedPass['ArrivedTo']
        $nowPeriod = ConvertTo-MbtArrivalPeriod -From $pass['ArrivedFrom'] -To $pass['ArrivedTo']
        if ([MailboxTransfer.ArrivalPeriod]::KeyOf($wasPeriod) -ne [MailboxTransfer.ArrivalPeriod]::KeyOf($nowPeriod)) {
            & $stop @("The batch $id was started with the period '$(if ($wasPeriod) { $wasPeriod.Describe() } else { 'all dates' })', this run asks '$(if ($nowPeriod) { $nowPeriod.Describe() } else { 'all dates' })'. Continuing it would mix two periods in the stores of the batch. Give the same period, or another -BatchName."); return
        }
    }
    $givenAuth = $useCertificate -or $boundParameters.ContainsKey('UserPrincipalName') -or $boundParameters.ContainsKey('TenantId') -or $boundParameters.ContainsKey('ClientId')
    $appOnly = [bool]$useCertificate
    $signInUser = if ($appOnly) { '' } else { [string]$UserPrincipalName }
    if ($useCertificate) {
        foreach ($n in 'CertificateThumbprint', 'CertificatePath', 'TenantId', 'ClientId') { $v = [string](Get-Variable -Name $n -ValueOnly); if ($v) { $pass[$n] = $v } }
    } elseif ($boundParameters.ContainsKey('TenantId') -or $boundParameters.ContainsKey('ClientId')) {
        foreach ($n in 'TenantId', 'ClientId') { if ($boundParameters.ContainsKey($n)) { $pass[$n] = [string](Get-Variable -Name $n -ValueOnly) } }
    }
    if ($batch -and -not $givenAuth -and $stored.Contains('AppOnly')) {
        $appOnly = [bool]$stored.AppOnly
        $signInUser = [string]$stored.SignInUser
        if ($appOnly) {
            $script:useCertificate = $true
            foreach ($n in 'CertificateThumbprint', 'CertificatePath', 'TenantId', 'ClientId') { if ($pass.Contains($n)) { Set-Variable -Name $n -Value ([string]$pass[$n]) -Scope Script } }
        }
    }
    $globalOutput = if ($OutputPst) { $OutputPst } elseif ($batch -and $stored.Contains('OutputPst')) { [string]$stored.OutputPst } else { '' }
    $storeBase = if ($boundParameters.ContainsKey('StorePath')) { $StorePath } elseif ($batch -and $stored.Contains('StoreBase') -and $stored.StoreBase) { [string]$stored.StoreBase } else { Join-Path (Join-Path $work 'stores') $id }
    $parallel = if ($boundParameters.ContainsKey('ParallelMailboxes')) { $ParallelMailboxes } elseif ($batch -and $batch.Contains('parallel')) { [int]$batch.parallel } else { 1 }
    if (-not $appOnly -and -not $signInUser) { & $stop @('Delegated sign-in needs the account that signs in: give -UserPrincipalName (or set Tenant.UserPrincipalName in the configuration), or sign in as the application (-AppOnly).'); return }
    $option = @{ AppOnly = $appOnly; SignInUser = $signInUser; OutputPst = $globalOutput; StoreBase = $storeBase; Pass = $pass }

    # the whole CSV file, before anything starts
    if ($MailboxCsv) {
        $csv = if ($Mode -eq 'Export') { Read-MbtBatchCsv -Path $MailboxCsv -DefaultLocation $Mailbox -OutputPst $globalOutput -StoreBase $storeBase -SignedInUser $signInUser -AppOnly:$appOnly }
        elseif ($Mode -eq 'Copy') {
            $choice = @{ IncludeRecoverableItems = [bool]$IncludeRecoverableItems }
            if ($Folder) { $choice.Folder = $Folder }
            if ($arrivalPeriod) { $choice.Period = $arrivalPeriod }
            if ($RecoverableItemsFolders) { $choice.RecoverableItemsFolders = $RecoverableItemsFolders }
            Read-MbtCopyBatchCsv -Path $MailboxCsv -DefaultSourceLocation $Mailbox -DefaultTargetLocation $TargetMailbox -DefaultTargetRoot $TargetRoot -StoreBase $storeBase -Choice $choice -SignedInUser $signInUser -AppOnly:$appOnly
        }
        else { Read-MbtImportBatchCsv -Path $MailboxCsv -Mode $Mode -DefaultLocation $Mailbox -DefaultTargetRoot $TargetRoot -DefaultSourceMailbox $SourceMailbox -MergeWellKnown:$MergeWellKnown -SignedInUser $signInUser -AppOnly:$appOnly -WorkFolder $work }
        if ($csv.Problems.Count) { & $stop $csv.Problems; return }
        if ($batch) {
            $diff = @(if ($Mode -eq 'Export') { Compare-MbtBatchRow -Batch $batch -Row $csv.Rows } elseif ($Mode -eq 'Copy') { Compare-MbtCopyBatchRow -Batch $batch -Row $csv.Rows } else { Compare-MbtImportBatchRow -Batch $batch -Row $csv.Rows })
            if ($diff.Count) {
                & $stop (@($diff | Select-Object -First 10 | ForEach-Object { "The CSV file is not the list of the batch $id`: $_." }) + @('A batch keeps the mailboxes it was made with: use -ResumeBatch to continue it, or another -BatchName for this list.'))
                return
            }
            $batch.csv = [IO.Path]::GetFullPath($MailboxCsv)
        } else {
            $batch = New-MbtBatch -Id $id -Csv ([IO.Path]::GetFullPath($MailboxCsv)) -Row $csv.Rows -Option $option -Parallel $parallel -Mode $Mode
        }
    }
    if ($Mode -ne 'Export' -and -not $appOnly) {
        $foreign = @(@($batch.rows) | Where-Object { [string]$_.upn -ine $signInUser })
        if ($foreign.Count) {
            & $stop @($foreign | Select-Object -First 10 | ForEach-Object { "$($_.upn) is not the signed-in account ($signInUser). A delegated sign-in writes into its own mailbox only: sign in as the application (-AppOnly) to continue this batch." })
            return
        }
    }
    $batch.options = $option
    $batch.parallel = $parallel
    $counts = Get-MbtBatchCount $batch
    Write-MbtStep ("Batch $id")
    Write-MbtItem INFO ("{0} {1}: {2} complete, {3} to do; {4} at a time; batch folder {5}" -f (Format-MbtNumber $counts.Total), $(if ($Mode -eq 'Export') { 'mailbox' + $(if ($counts.Total -eq 1) { '' } else { 'es' }) } elseif ($Mode -eq 'Copy') { 'cop' + $(if ($counts.Total -eq 1) { 'y' } else { 'ies' }) } else { 'import' + $(if ($counts.Total -eq 1) { '' } else { 's' }) }), (Format-MbtNumber $counts.Complete), (Format-MbtNumber ($counts.Total - $counts.Complete)), $parallel, (Join-Path (Join-Path $work 'batches') $id))
    Show-MbtBatchPlan -Batch $batch -Limit $(if ($Mode -eq 'Export') { 25 } else { 200 })
    if ($Mode -ne 'Export') {
        $shared = @(@($batch.rows) | Group-Object { ([string]$_.upn).ToLowerInvariant() + '|' + ([string]$_.location).ToLowerInvariant() } | Where-Object { $_.Count -gt 1 })
        $sameSource = @(@($batch.rows) | Group-Object { if ($Mode -eq 'Copy') { ([string]$_.sourceUpn + '|' + [string]$_.sourceLocation).ToLowerInvariant() } else { ([string]$_.source).ToLowerInvariant() } } | Where-Object { $_.Count -gt 1 })
        if ($shared.Count) { Write-MbtDetail ('{0} target mailbox location{1} {2} written by several lines: those lines run one after the other (the service throttles each mailbox); different mailboxes run side by side' -f (Format-MbtNumber $shared.Count), $(if ($shared.Count -eq 1) { '' } else { 's' }), $(if ($shared.Count -eq 1) { 'is' } else { 'are' })) }
        if ($sameSource.Count) { Write-MbtDetail ('{0} source{1} {2} read by several lines: those lines run one after the other' -f (Format-MbtNumber $sameSource.Count), $(if ($sameSource.Count -eq 1) { '' } else { 's' }), $(if ($sameSource.Count -eq 1) { 'is' } else { 'are' })) }
        $targets = [Collections.Generic.List[string]]::new()
        $seenTargets = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($r in @($batch.rows)) { if ($seenTargets.Add(([string]$r.upn + '|' + [string]$r.location))) { $targets.Add(('{0} ({1})' -f $r.upn, $r.location)) } }
        $targetText = (($targets | Select-Object -First 50) -join ', ') + $(if ($targets.Count -gt 50) { " and $(Format-MbtNumber ($targets.Count - 50)) more (the list above, and batch.json)" })
        Write-MbtItem INFO ('{0} will be written to: {1}. Items are created only; an item that was {2} is never created twice.' -f $(if ($targets.Count -eq 1) { 'this mailbox location' } else { ('these ' + $targets.Count + ' mailbox locations') }), $targetText, $(if ($Mode -eq 'Copy') { 'copied' } else { 'imported' }))
    }
    $command = ('.\Invoke-MailboxTransfer.ps1 -Mode ' + $Mode + ' -ResumeBatch ') + (Format-MbtBatchToken $id) + $(if ($WorkFolder) { ' -WorkFolder ' + (Format-MbtBatchToken $work) }) + $(if ($parallel -gt 1) { " -ParallelMailboxes $parallel" })
    if ($dryRun) {
        Write-MbtItem WARN $(if ($Mode -eq 'Export') { 'WhatIf: the file is valid, nothing was started and no folder was written.' } elseif ($Mode -eq 'Copy') { 'WhatIf: the file is valid; nothing was started, nothing was read and nothing was written to any mailbox or store. To see the plan of one line (what is read, the time), run that copy alone with -WhatIf.' } else { 'WhatIf: the file is valid and every source was found; nothing was started and nothing was written to any mailbox. To see the folders and the item counts of one line, run that import alone with -WhatIf.' })
        Write-MbtBatchSummary -ExitCode 0 -Batch $batch -WhatIf -Log $logPath -Mode $Mode
        return
    }
    if ($counts.Complete -eq $counts.Total) { Write-MbtItem OK 'every mailbox of this batch is already complete: nothing to do' }

    # one sign-in for the whole batch: the account or the application is checked here, and what is cached serves every mailbox silently
    if ($counts.Complete -lt $counts.Total) {
        if ($appOnly) { $script:UserPrincipalName = if ($Mode -eq 'Copy') { [string]@($batch.rows)[0].sourceUpn } else { [string]@($batch.rows)[0].upn } }
        $null = Connect-ForTransfer
        Disconnect-MbtGraph
    }
    try {
        $purge = Invoke-MbtJobRetention -WorkFolder $work -RetentionDays $config.Jobs.RetentionDays -OverviewBudgetSeconds $overviewBudgetSeconds
        if ($purge -and @($purge.Removed).Count) { Write-MbtItem INFO "removed $(Format-MbtNumber @($purge.Removed).Count) old job $(if (@($purge.Removed).Count -eq 1) { 'folder' } else { 'folders' }); the names are in the log" }
    } catch { Write-MbtItem WARN "the old job folders could not be removed: $($_.Exception.Message)" }

    $run = $option.Clone()
    $runPass = $pass.Clone()
    if ($parallel -gt 1 -and -not $runPass.Contains('Concurrency')) { $runPass.Concurrency = 2 }
    $run.Pass = $runPass
    $result = Invoke-MbtBatch -WorkFolder $work -Batch $batch -EntryScript $entryScript -Option $run -Parallel $parallel -StopFile $StopFile -LogFolder $LogFolder -DevelopmentMode:$DevelopmentMode
    $failedRows = @($result.Batch.rows | Where-Object { $_.state -eq 'failed' })
    if ($failedRows.Count) {
        Write-MbtStep ('Mailboxes that failed ({0})' -f (Format-MbtNumber $failedRows.Count))
        foreach ($r in ($failedRows | Select-Object -First 25)) { Write-MbtItem FAIL ('{0}: {1}' -f $r.upn, $r.message) }
        if ($failedRows.Count -gt 25) { Write-MbtDetail ('and {0} more: see the report and batch.log' -f (Format-MbtNumber ($failedRows.Count - 25))) }
    }
    $allJobs = $null
    try {
        $overviewInfo = @{}
        $allJobs = Write-MbtJobOverview -WorkFolder $work -TimeBudgetSeconds $overviewBudgetSeconds -Info $overviewInfo
        if (-not $allJobs -and $overviewInfo.Contains('Complete') -and -not $overviewInfo.Complete) { Write-MbtItem WARN "the overview of all the jobs was not finished in $overviewBudgetSeconds s: the next job or the console finishes it" }
    } catch { Write-MbtItem WARN "the overview of all the jobs could not be written: $($_.Exception.Message)" }
    Write-MbtBatchSummary -ExitCode $result.ExitCode -Batch $result.Batch -Folder $result.Folder -Seconds $clock.Elapsed.TotalSeconds -Report $result.Report -AllJobs $allJobs -Log $logPath -Command $command -Stopped:$result.Stopped -Mode $Mode
    $script:exitCode = $result.ExitCode
}

try {
    # -Mode Copy is two jobs in one process: the export of the source into the store, then the import of that store into the target
    $phases = if ($Mode -eq 'Copy' -and -not $batchMode) { @('Export', 'Import') } elseif ($Mode -eq 'Copy') { @('Export') } else { @($Mode) }
    foreach ($phase in $phases) {
        $script:phase = $phase
        if ($Mode -eq 'Copy' -and -not $batchMode) { Set-MbtJobPhase -Phase $phase.ToLowerInvariant(); if ($phase -eq 'Import') { Enter-CopyImport } }
        switch ($phase) {
            'Status' {
                Write-MbtItem OK 'the C# engine compiles'
                Write-MbtItem INFO "Graph: $($config.Graph.BaseUrl), retry up to $($config.Graph.MaxAttempts) attempts, token cache: $($config.Auth.CacheFile)"
                if (-not $PSBoundParameters.ContainsKey('StorePath') -and -not ($Unknown -or $Locations -or $ResolveUnknown -or $ResolveAbsent)) { Show-JobList }
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
                if ($batchMode) { Invoke-BatchRun; break }
                $null = Connect-ForTransfer
                # a copy checks its target and its store before it reads anything; an export half that is already complete is not done again
                if ($Mode -eq 'Copy') {
                    Resolve-CopyTarget
                    if (-not $dryRun -and (Test-CopyExportDone)) { break }
                }
                # read-only: the folder trees are read once for the whole run, and the reading says how far it is
                Write-MbtJobStep -Id tree -State running -Msg $(if ($SelectionFile) { 'Reading your selection' } else { 'Reading the folder tree' })
                Start-MbtPhaseProgress -Step tree
                $plan = $null
                # what the run exports: everything asked by -Mailbox / -MailboxId / -Folder (one unit without a plan), or one unit per mailbox of -SelectionFile
                $units = @([pscustomobject]@{ Plan = $null; Kind = $null; Id = $null })
                if ($SelectionFile) {
                    # what the browse view selected: the mailbox is the one of the file (a file may hold several mailboxes); the entries are checked against the import-export API first
                    if ($Folder) { throw '-SelectionFile and -Folder are exclusive: the selection names its folders.' }
                    $units = @()
                    foreach ($selection in @(Read-MbtBrowseSelectionSet -Path $SelectionFile)) {
                        $selectedId = ConvertTo-MbtMailboxId -Value $selection.MailboxId -TenantId $TenantId
                        if ($mailboxGiven -and $Mailbox -ne $selection.Kind) { throw "-Mailbox $Mailbox does not match the selection, which is of the $($selection.Kind) mailbox." }
                        if ($MailboxId -and (ConvertTo-MbtMailboxId -Value $MailboxId -TenantId $TenantId) -ine $selectedId) { throw "-MailboxId $MailboxId is not the mailbox of the selection ($selectedId)." }
                        Write-MbtItem INFO ("selection {0}: {1} entr{2} for the {3} mailbox {4}" -f $SelectionFile, $selection.Entries.Count, $(if ($selection.Entries.Count -eq 1) { 'y' } else { 'ies' }), $selection.Kind, $selectedId)
                        # a selection made on an archive still resolves when a folder moved to an auxiliary archive since: the other archives are searched by folder entry key
                        $also = if ($selection.Kind -in 'Archive', 'AuxArchive') { @(Get-OtherArchiveLocation $selectedId) } else { @() }
                        $selectionPlan = Resolve-MbtBrowsePlan -Selection $selection -AlsoSearch $also -RecoverableItemsFolders $RecoverableItemsFolders
                        Show-MbtExportPlan -Plan $selectionPlan -AtMost:([bool]$arrivalPeriod)
                        if (-not $selectionPlan.IsComplete) { throw "-SelectionFile: $($selectionPlan.Problems.Count) problem(s), nothing was exported. Remove those entries from the selection (or select them again) and run again." }
                        if ($Mode -eq 'Copy') { Assert-CopySource $selectedId; $copy.SourceId = $selectedId }
                        $units += [pscustomobject]@{ Plan = $selectionPlan; Kind = [string]$selection.Kind; Id = $selectedId }
                    }
                    Stop-MbtPhaseProgress
                    if ($dryRun) {
                        # the counts of the plan (counted by the service, folders below included), not the ones the browse view wrote in the file
                        foreach ($unit in $units) {
                            $selTotals = $unit.Plan.Totals()
                            Write-MbtItem INFO ("selection ({0} mailbox): {1}{2}" -f $unit.Kind, (Format-MbtBrowseTotals $selTotals -AtMost:([bool]$arrivalPeriod)), $(if ($Mode -ne 'Copy' -and -not $selTotals.SizeUnknown) { "; about {0} PST file(s) of at most {1} if converted (-OutputPst)" -f [math]::Max(1, [math]::Ceiling($selTotals.KnownSizeBytes / $pstMaxBytes)), (Format-MbtBytes $pstMaxBytes) }))
                        }
                        Write-MbtJobStep -Id tree -State done
                        if ($Mode -eq 'Copy') {
                            $planTotals = @($units | ForEach-Object { $_.Plan.Totals() })
                            Write-CopyPlan -Items ([long]($planTotals | Measure-Object -Property Items -Sum).Sum) -Bytes ([double]($planTotals | Measure-Object -Property KnownSizeBytes -Sum).Sum) -SizeUnknown ([bool]@($planTotals | Where-Object SizeUnknown).Count)
                            break
                        }
                        Write-MbtPeriodPlanNote
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
                            if ($Mode -eq 'Copy') { Assert-CopySource $id; $copy.SourceId = $id }
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
                            $totals = Show-MbtLiveExportPlan -Tree $tree -Kind $tg.Kind -Folder $Folder -PstMaxBytes $pstMaxBytes -IncludeRecoverableItems:$withRecoverable -NoPst:($Mode -eq 'Copy')
                            $sumItems += $totals.Items; $sumBytes += $totals.KnownSizeBytes; if (-not $totals.SizeComplete) { $sumUnknown = $true }
                        } catch { Write-MbtItem WARN "the folders of the $($tg.Kind) mailbox $($tg.Id) could not be read: $($_.Exception.Message)" }
                    }
                    if ($asked -in 'Archive', 'All') { Write-MbtDetail 'an auto-expanding archive: the items of its expanded folders live in auxiliary archives; the export follows them and records them under the folder of the main archive' }
                    if ($Mode -eq 'Copy') { Stop-MbtPhaseProgress; Write-MbtJobStep -Id tree -State done; Write-CopyPlan -Items $sumItems -Bytes $sumBytes -SizeUnknown $sumUnknown; break }
                    Write-MbtItem INFO ("total: {0}{1} item(s), {2}{3}; about {4} PST file(s) of at most {5} if converted ({6})" -f $(if ($arrivalPeriod) { 'at most ' }), $sumItems, (Format-MbtBytes $sumBytes), $(if ($sumUnknown) { ' announced (some folders give no size)' } else { ' announced' }),
                        [math]::Max(1, [math]::Ceiling($sumBytes / $pstMaxBytes)), (Format-MbtBytes $pstMaxBytes), $(if ($OutputPst) { $OutputPst } else { 'give -OutputPst to convert' }))
                    if ($OutputPst) { Write-MbtDetail "after the export the store would be converted: $OutputPst, then -002.pst... and a manifest" }
                    Write-MbtPeriodPlanNote
                    Stop-MbtPhaseProgress; Write-MbtJobStep -Id tree -State done
                    if (-not $PSCmdlet.ShouldProcess($StorePath, "Export $($targets.Count) location(s) into the store")) { Write-MbtItem WARN 'WhatIf: nothing was exported and no store was created.' }
                    break
                }
                if (-not $store) { $store = Open-MbtExportStore -Path $StorePath -Name 'export' -Source "$UserPrincipalName ($TenantId)" }
                if ($Mailbox -eq 'All' -and $MailboxId) { throw '-MailboxId is one location: it cannot be combined with -Mailbox All.' }
                # the locations given up front (operator file, parameters, Graph) are recorded in the store, with how they were learned
                $known = @(Get-MbtMailboxLocation -UserPrincipalName $UserPrincipalName -TenantId $TenantId -Store $store -LocationsFile $LocationsFile -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId -SkipGraph -Verbose:$false)
                $auxKnown = @($known | Where-Object Kind -EQ 'AuxArchive')
                if ($auxKnown) { Write-MbtItem INFO ("{0} auxiliary archive(s) known: {1}" -f $auxKnown.Count, (($auxKnown | ForEach-Object { "$($_.MailboxId) [$($_.Source)]" }) -join ', ')) }
                $failures = 0
                $exported = @{}
                $script:resumable = $false
                function Invoke-OneExport([string]$Kind, [string]$Id) {
                    Stop-MbtPhaseProgress
                    # a plan names folders of several mailboxes when a folder moved to an auxiliary archive: each mailbox exports its own part
                    $subPlan = if ($plan) { $plan.ForMailbox($Id) }
                    if ($plan -and $subPlan.Folders.Count -eq 0) { Write-MbtItem INFO "nothing of the selection is in the $Kind mailbox $Id"; return }
                    Write-MbtItem INFO "exporting the $Kind mailbox $Id into $StorePath"
                    $exported[$Id.ToLowerInvariant()] = $true
                    $r = Export-MbtMailbox -Store $store -MailboxId $Id -Kind $Kind -UserPrincipalName $UserPrincipalName -Concurrency $Concurrency `
                        -BatchItems $BatchItems -BatchBytes ([long]$BatchMiB * 1MB) -MaxAttempts $MaxAttempts -NotReadyTrip $NotReadyTrip -SkipInventory:$SkipInventory -Folder $Folder -Plan $subPlan -IncludeRecoverableItems:$IncludeRecoverableItems -RecoverableItemsFolders $RecoverableItemsFolders -Period $arrivalPeriod -TempPath $TempPath -IntervalSeconds $IntervalSeconds
                    if ($r.Error) { throw $r.Error.Message }
                    if ($arrivalPeriod) {
                        $script:counts['kept'] = [long]$script:counts['kept'] + [long]$r.Result.Inventoried
                        $script:counts['outsidePeriod'] = [long]$script:counts['outsidePeriod'] + [long]$r.Result.ItemsOutsidePeriod
                        $script:counts['withoutDate'] = [long]$script:counts['withoutDate'] + [long]$r.Result.ItemsWithoutDate
                    }
                    $script:counts['done'] = [long]$script:counts['done'] + [long]$r.Result.Exported
                    $script:counts['failed'] = [long]$script:counts['failed'] + [long]$r.Result.Failed
                    $script:counts['deferred'] = [long]$script:counts['deferred'] + [long]$r.Result.FoldersDeferred.Count
                    $script:counts['bytes'] = [long]$script:counts['bytes'] + [long]$r.Result.Bytes
                    foreach ($d in $r.Result.FoldersDeferred) { Write-MbtJobObject -Step download -Key ([string]$d.Path) -State notready -Msg ([string]$d.ToString()) }
                    if ($r.Result.NeedsAnotherRun -or $r.Job.Cancelled) { $script:resumable = $true }
                    if ($Mode -eq 'Copy' -and ($r.Job.Cancelled -or -not $r.Result.Completed)) { $script:copyExportStopped = $true }
                }
                $askedMailbox = $Mailbox; $askedMailboxId = $MailboxId; $askedGiven = $mailboxGiven
                foreach ($unit in $units) {
                $plan = $unit.Plan
                if ($plan) { $Mailbox = $unit.Kind; $MailboxId = $unit.Id; $mailboxGiven = $true }
                $asked = Get-AskedKind
                $kinds = if ($asked -eq 'All') { @('Primary', 'Archive') } else { @($asked) }
                foreach ($kind in $kinds) {
                    try {
                        $id = Resolve-MbtMailboxId -Mailbox $kind -TenantId $TenantId -UserPrincipalName $UserPrincipalName -PrimaryMailboxId $PrimaryMailboxId -ArchiveMailboxId $ArchiveMailboxId `
                            -MailboxId $(if ($asked -ne 'All') { $MailboxId }) -Store $store -LocationsFile $LocationsFile
                        if ($Mode -eq 'Copy') { Assert-CopySource $id }
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
                }
                # several mailboxes in the selection: the PST holds what the selection names, whatever the mailbox
                if ($units.Count -gt 1) { $Mailbox = $askedMailbox; $MailboxId = $askedMailboxId; $mailboxGiven = $askedGiven }
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
                if ($batchMode) { Invoke-BatchRun; break }
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
                        if ($arrivalPeriod) { $emlArgs.Period = $arrivalPeriod }
                        $r = Export-MbtEml @emlArgs -Location $loc
                    }
                }
                Write-MbtItem OK ("exported {0}, already there {1}, skipped {2}, failed {3}; {4}; manifest {5}" -f $r.Exported, $r.Resumed, $r.Skipped, $r.Failed, (Format-MbtBytes $r.Bytes), $r.ManifestPath)
                foreach ($n in $r.FolderNotes) { Write-MbtItem WARN "folder not exported: $n" }
                foreach ($p in @($r.Problems | Select-Object -First 20)) { Write-MbtItem WARN ("{0}: {1} - {2} ({3})" -f $p.State, $p.Folder, $p.Subject, $p.Reason) }
                $script:counts = @{ done = [long]$r.Exported; resumed = [long]$r.Resumed; skipped = [long]$r.Skipped; failed = [long]$r.Failed; bytes = [long]$r.Bytes; folders = [long]$r.Folders }
                if ($arrivalPeriod) {
                    $script:counts.kept = [long]($r.Exported + $r.Resumed + $r.Skipped + $r.Failed)
                    $script:counts.outsidePeriod = [long]$r.OutsidePeriod
                    $script:counts.withoutDate = [long]$r.WithoutDate
                    Write-MbtItem INFO ("period ({0}): {1} message(s) outside the period left out (not an error){2}" -f $arrivalPeriod.Describe(), (Format-MbtNumber $r.OutsidePeriod), $(if ($r.WithoutDate) { "; $(Format-MbtNumber $r.WithoutDate) without a date were kept" }))
                }
                if ($r.Failed -gt 0 -or $r.Cancelled) { $exitCode = 2 }
                Write-MbtJobStep -Id export -State $(if ($exitCode -eq 0) { 'done' } else { 'running' }) -Done ($r.Exported + $r.Resumed + $r.Skipped + $r.Failed) -Bytes $r.Bytes -Msg $(if ($r.Cancelled) { 'stopped' })
            }
            'ImportEml' {
                if ($batchMode) { Invoke-BatchRun; break }
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
                if ($batchMode) { Invoke-BatchRun; break }
                if (-not $storeExists) { throw "No store in ${StorePath}: export first (-Mode Export) or give -StorePath." }
                if ($Mailbox -eq 'All') { throw '-Mailbox All is for Export. An import goes into one mailbox: Primary, Archive or AuxArchive (-MailboxId).' }
                if (-not $store) { $store = Open-MbtExportStore -Path $StorePath }
                $sourceKind = if ($SourceMailbox) { $SourceMailbox } else { $null }
                if (-not $sourceKind) {
                    # the items of an auxiliary archive belong to the folders of the main archive: only Primary / Archive are sources
                    $kinds = @($store.GetMailboxes() | Where-Object { $_.Kind.ToString() -ne 'AuxArchive' } | ForEach-Object { $_.Kind.ToString() })
                    if ($kinds.Count -ne 1) { throw "The store holds $($kinds.Count) mailboxes ($($kinds -join ', ')): choose with -SourceMailbox Primary|Archive." }
                    $sourceKind = $kinds[0]
                }
                $sourceId = Get-StoreMailboxId $sourceKind
                Write-MbtJobStep -Id check -State done -Msg 'The export store is open.'
                if ($Mode -ne 'Copy') { $null = Connect-ForTransfer }
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
                    $importCounts = @{ done = [long]$jr.JobDone; failed = [long]$jr.JobFailed; unknown = [long]$jr.JobUnknown; pending = [long]$jr.JobPending; total = [long]$jr.JobTotal; foldersCreated = [long]$jr.FoldersCreated }
                    if ($Mode -eq 'Copy') {
                        # both halves in the counts of the job: what the export half did stays, its failed items count as failed
                        foreach ($k in $script:counts.Keys) { if (-not $importCounts.ContainsKey($k)) { $importCounts[$k] = $script:counts[$k] } }
                        $importCounts.failed = [long]$jr.JobFailed + [long]$script:counts['exportFailed']
                    }
                    $script:counts = $importCounts
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
        if ($Mode -eq 'Copy' -and -not $batchMode -and $phase -eq 'Export' -and -not $dryRun) {
            # the export half is complete only when nothing is left of it: the next run goes straight to the import
            if ($exitCode -eq 0 -and -not $script:resumable -and -not $script:copyExportStopped -and -not (Test-MbtStopRequested)) { [void]$store.SetExportOption('copy.export', [DateTime]::UtcNow.ToString('yyyy-MM-dd HH:mm') + ' UTC', $true) }
            Convert-CopyExportCounts
        }
        # the import half of a copy follows an export half that finished its work, not a plan, an error or a stop
        if ($Mode -eq 'Copy' -and -not $batchMode -and $phase -eq 'Export' -and ($dryRun -or $exitCode -eq 1 -or $script:copyExportStopped -or (Test-MbtStopRequested))) { break }
    }
    if ($exitCode -ne 0 -and -not $script:reason -and (Test-MbtStopRequested)) { $script:reason = 'stopped' }
    if ($script:reason -eq 'stopped' -and -not $script:next) { $script:next = 'continue' }
    if ($Mode -eq 'Copy' -and -not $batchMode) { $script:copyNext = Get-CopyNext $exitCode }
    $resultArgs = @{ ExitCode = $exitCode; Counts = $script:counts }
    if ($script:resultComplete) { $resultArgs.Complete = $true }
    if ($script:next) { $resultArgs.Next = $script:next }
    if ($script:reason) { $resultArgs.Reason = $script:reason }
    Write-MbtJobResult @resultArgs
    $script:sawResult = $true
    if ($Mode -ne 'Browse' -and -not $batchMode) { Write-MbtRunSummary -ExitCode $exitCode }
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
    Stop-MbtPhaseProgress
    Disconnect-MbtGraph
    Stop-MbtStopWatch
    Stop-MbtKeepAwake
    Close-MbtEventLog
}
exit $exitCode

# SIG # Begin signature block
# MIImKgYJKoZIhvcNAQcCoIImGzCCJhcCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCJuuJ8LPSv1Tx7
# M3Aqd7waAGqqM7JTO4wbWnoTsFu0faCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
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
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBrr1s0mLt9
# GQVAPPZVNV3zgJ0FSv9cRPnSMBYeu8dapzANBgkqhkiG9w0BAQEFAASCAgAPHKMq
# bRbU6FInDlsUWpTxkych9EmogQW9TTeqzpl2EImg5KWn0kft3MZgcT2W7lmlU7XP
# GDujc6iBQBIzZfU16u+ym4brG33yzY0xuOTQgtY+XA0kpCyqwNPmLM7Lx/MvjiSt
# gIU9v9iTNpaT7jsDtvUjoXgbJNLgphAgkYVQGZeNcar+mwd5RMb9l4gwc90f6l4t
# xML6kvgzudcZSGEU2CtrOUHH0fNOO949/k5TfCUVT/A5FUtAlrKCVKBpJzfAehRe
# ZcnPKSiLiSjyASAHd5J6I1c7j0MEebk0KrT+e8tcokzYyrGH7aXnozLfFc2FRysU
# dyIbnJB+fS2HOa/C9ukxbW2EjZ5MkANUyh6f5ykZ7usXy7GiA4E/UkqCoUcxE40M
# v05EI2M4rvAbfFOQ7CfXLB2hDPduVsqpNX23RjkdZMwrnMhDS7FChq6c9omjdzE8
# AMUdwEfZ7lBa5xxgmdHt6ZdzLcNuWX4+x9V9vLOoFr3ExDDiW37e+RkKTNu2v+FG
# PldXzfdj1XWvdQPaqmkmymXb4FsvgYaBXuZ8DrOTEGpIK3ftJ01pFFvncGtglrQe
# C0wfv4w0NdoGUjrAZ5qgSF/wz9uaeSFnqbTlmCsn5vIHEeJHD41t1/xbrmkSs00z
# w03Gi2WYe0PCImyeLjoH3re1oZ4VoOxzA6SApaGCFM4wghTKBgorBgEEAYI3AwMB
# MYIUujCCFLYGCSqGSIb3DQEHAqCCFKcwghSjAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEIOYbrIJPZemS4aaa3W01nunjFCIXaz8wP5LkLw2Q+fCUAggjmj9FoICt
# LBgPMjAyNjExMTUxNDMzNThaMAMCAQECEAMFjZDfKELJkbkF2CHqC9WgghGnMIIE
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
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExNTE0MzM1OFowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIFyK
# vEf3iSidP9DUvdYn+3L5OUNZollDsDGbsopNSlh2MIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERzBFAiAPdkW0UmFqBWUJHj63k0/eoUMamlSEwttz/H6fICjeuQIhAJEd
# yPc+o7KAI5y27RLnGp5MJerzQUcdBJRhhDR0XMrl
# SIG # End signature block
