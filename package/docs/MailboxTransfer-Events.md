# Events of a job (the contract of the console)

A long job (export, PST write, import, check, repair) is run by `Invoke-MailboxTransfer.ps1` in a child process. The console never parses its text: it reads **events.jsonl**, written by the script, and **status.json**, and it stops the job with a **stop file**. This page is the contract, field by field. The text the script writes on the console is unchanged.

## Command line

| Parameter | Meaning |
|---|---|
| `-EventsFile <path>` | Appends the events of the job to this file (one JSON object per line, UTF-8 without BOM, flushed at each line). The folder of the file is the job folder: `status.json` is written beside it. |
| `-JobId <text>` | Identifier of the job in the events and in `status.json` (default: the name of the job folder). |
| `-StopFile <path>` | When this file exists the job stops like Ctrl+C: the requests in flight finish, the state is saved, exit code 2, `result.next = "continue"`, `result.reason = "stopped"`. The file is looked at once when the script starts, then every second. |
| `-KeepAwake` | Windows does not go to sleep by idle while the script runs (`SetThreadExecutionState`, held by a thread of its own, released when the script ends). A closed lid, Sleep chosen by hand and a restart still stop the PC. |
| `-NoPrompt` | Silent sign-in only (no browser, no device code). Without a valid cached token the script ends with exit code 2, an `error` event with code `signin-required` and `result.reason = "signin-required"`, `result.next = "signin"`. |
| `-ConfigPath <psd1>` | The configuration file, now also used by the sign-in (token cache, Graph settings): the console writes an *effective* configuration here (run-only settings, never written to `config\`). |

Modes added for the console (the other modes keep their parameters and exit codes):

| Mode | Arguments | Exit code |
|---|---|---|
| `CheckPst` | `-Path <pst>` (read only) | 0 clean, 2 warnings, 1 errors; **2 with reason `stopped`** when stopped |
| `RepairPst` | `-Path <pst> -Destination <new pst> [-Level Automatic|Recompute|Rebuild|Salvage] [-KeepDamagedAttachments]` (never overwrites, never touches the original; `aux.pst` and the other Windows device names are refused) | 0 repaired and clean, 2 stopped (`stopped`) or errors remain (`errors-remain`), 1 error |
| `ImportPst` | `-PstPath <pst or folder> -TargetRoot <folder> [-MergeWellKnown] [-IncludeFolder ...] [-ExcludeFolder ...] [-FolderFile <file>] [-Mailbox ...] [-MailboxId <address>]` (the journal is in `-TempPath`) | 0 complete, 3 Unknown rows, 2 resumable (or stopped), 1 error **or finished with items refused for good** (`reason` `refused-items`), see "Exit codes" |
| `ExportEml` / `ImportEml` | `-EmlPath <folder>` with `-StorePath`, `-PstPath` or the live mailbox / `-EmlPath <folder> -TargetRoot <folder>` | 0, 2 failed items or stopped, 3 (ImportEml: Unknown), 1 error |
| `ExportPst` from a PST | `-PstPath <pst> -OutputPst <file or folder> [-IncludeFolder ...] [-ExcludeFolder ...] [-FolderFile <file>] [-IncludeItemFile <file>] [-PstMaxSizeGB n] [-SkipCheck]` (chosen folders and messages into new PST files; the source is only read; a stopped copy leaves no file) | 0 finished, 2 stopped, 1 error, or finished with messages not copied (`refused-items`) |
| `Export` / `ExportPst` (from a store) / `Import` / `Status` / `Browse` | as before | as before (Import: 0, 1, 2, 3); a store with items not exported yet: `ExportPst` 2, `Import` 2 |

### Exit codes: one rule for every mode

| Code | Meaning | `result.next` |
|---|---|---|
| 0 | Complete: nothing is left to do and nothing was lost. | `done` |
| 1 | An error stopped the run, **or** (ImportPst, ExportPst from a PST) the job is finished but some items are refused for good: `reason` `refused-items`, `complete` true, status.json state `complete`. There is nothing to continue; the list `<TempPath>\<JobName>-not-imported.csv` (`counts.report`, one line per item with its reason) and one `item` event per refused item (the first 200) say which. | `retry` (error), `done` (refused) |
| 2 | Not complete and resumable: run the same command again, it does only what is left (Pending or Failed rows, a deferred folder, a stopped run, items of the store not exported yet, a folder of the target that could not be created, a sign-in to renew). | `continue`, `signin` |
| 3 | Rows are Unknown (the item may exist). They are never replayed by themselves: settle them (`-Mode Status -Unknown -Check`; `Resolve-MbtEmlUnknown` for .eml files), then run again. | `decide` |

When several apply the order is 3, 2, 1, 0. A row is **Refused** (final) when trying again cannot change the answer: the service refuses the content (HTTP 400, 413, 415, 422), it cannot be read from the PST or expressed as a FastTransfer stream, or it failed with nothing created on every attempt allowed (`-MaxAttempts`); a job is complete when every row is Done or Refused. A 2xx answer of `importItem` that cannot be read means the item exists: Unknown, never sent again. A console chain of commands ends with the worst code of its commands.

Every mode that runs as a job of the console takes `-EventsFile`, `-StopFile`, `-KeepAwake`, `-NoPrompt`, `-WhatIf`: Export, ExportPst, Import, ImportPst, ExportEml, ImportEml, CheckPst, RepairPst. A mailbox that is not the signed-in user's (a shared mailbox) is given by its address in `-MailboxId`.

`-Level Automatic` (default) checks the file once, then repairs at the level the check recommends (`Repair-MbtPst -Level Automatic`). All modes take `-WhatIf`.

`Get-MbtCapability` tells what an engine can do (modes and parameters of the entry script and commands of the module), so that the console shows the missing features as "coming" and lights them up when they exist: `Events`, `StopFile`, `KeepAwake`, `PstExport` (mode ExportPst), `PstCheck`, `PstRepair`, `PstImport` (mode **ImportPst** and parameter `-PstPath`), `EmlExport` (mode ExportEml), `EmlImport` (mode ImportEml), `SharedMailbox` (parameter `-MailboxId` and command `Resolve-MbtMailboxAddress`: a mailbox is given by its address).

## The line

Every line has `v` (format version, 1), `t` (UTC, `2026-11-09T14:02:11.123Z`) and `e` (the kind). Fields that are unknown are **left out** (unknown is not zero). A reader ignores the kinds and the fields it does not know.

| Kind | Fields | Example |
|---|---|---|
| `plan` | `job`, `task` (`export`, `exportpst`, `import`, `check`, `repair`, `status`, `browse`), `mode`, `steps` [`{id,title}`], all known before the start | `{"v":1,"t":"...","e":"plan","job":"imp-7c1e","task":"import","mode":"Import","steps":[{"id":"check","title":"Verify the export store"},{"id":"folders","title":"Create the folders"},{"id":"send","title":"Send the items"},{"id":"reconcile","title":"Check the uncertain items"},{"id":"final","title":"Final check"}]}` |
| `step` | `id`, `state` (`waiting`, `running`, `done`, `failed`), `done`, `total` (items), `bytes`, `bytesTotal`, `msg` | `{"v":1,"t":"...","e":"step","id":"send","state":"running","done":5120,"total":9860,"bytes":26391000000,"bytesTotal":50784000000}` |
| `object` | `step`, `key` (folder path or group), `state` (`planned`, `refused`, `running`, `done`, `check`, `notready`...), `done`, `total`, `bytes`, `msg`. With `-WhatIf` the plan is `object` events with `step = "plan"`: `key` the folder, `total` its items, `bytes` its size (-1: unknown). | `{"v":1,"t":"...","e":"object","step":"plan","key":"R/Inbox","state":"planned","total":6,"bytes":18233}` |
| `item` | `step`, `key`, `state`, `msg`, `detail` (+ `count`, `layer` for a check finding). CheckPst: one per finding code (`state` `error` or `warning`, `count` = how many, `msg` the first finding, `detail` where). RepairPst: one per message not fully recovered (`state` `partial` or `lost`, `msg` what was lost, `detail` its folder). | `{"v":1,"t":"...","e":"item","step":"check","key":"MSG-NOT-LISTED","state":"error","msg":"...","count":3}` |
| `wait` | `reason` (`throttle`), `seconds` (this wait), `count` (waits so far), `waited` (seconds waited so far) | `{"v":1,"t":"...","e":"wait","reason":"throttle","seconds":38,"count":7,"waited":252}` |
| `rate` | `window` (`5m`), `gibPerHour` of that 5-minute window | `{"v":1,"t":"...","e":"rate","window":"5m","gibPerHour":3.9}` |
| `warn` | `code`, `step`, `key`, `msg` (an item refused, an unknown row, ...) | `{"v":1,"t":"...","e":"warn","code":"engine","step":"send","msg":"..."}` |
| `decision` | `key` (the import row), `msg` (plain sentence), `candidates` (ids of the items that could be it), `item`, `folder`, `size` | `{"v":1,"t":"...","e":"decision","key":"42","msg":"Two or more items in the folder could be 'Signed contract'. We do not guess. Is it already in the mailbox?","candidates":["AAMk1","AAMk2"]}` |
| `error` | `code` (`signin-required`, `disk-full`, `reserved-name`, `destination-exists`, `mailbox-not-found`, `access-denied`, `store-busy`, `error`), `msg` (plain sentence: what happened, what is safe), `tech` (the raw message, for "Technical message") | `{"v":1,"t":"...","e":"error","code":"disk-full","msg":"The disk is full. Free some space and press Try again. What was downloaded is kept.","tech":"..."}` |
| `slept` | `from`, `to`, `seconds`: the process was frozen (sleep, hibernation); a heartbeat every 5 s sees a jump of the clock over 120 s | `{"v":1,"t":"...","e":"slept","from":"2026-11-09T02:14:00.000Z","to":"2026-11-09T06:48:00.000Z","seconds":16440}` |
| `info` | `msg` | |
| `result` | `exit` (the exit code of the script), `complete` (`exit == 0`), `counts` (object), `next` (`done`, `continue`, `decide`, `retry`, `signin`, `repair`), `reason` | `{"v":1,"t":"...","e":"result","exit":2,"complete":false,"counts":{"done":5120,"failed":2},"next":"continue","reason":"stopped"}` |

A job ends with exactly one `result`, also after an error (`error` then `result` with `exit` 1, `next` `retry`).

### Steps by mode

| Mode | Step ids (titles are the UI text) |
|---|---|
| Export | `tree` Read the folder tree, `list` List the items you chose, `download` Download the items, `verify` Check the downloads; with `-OutputPst`: `pst` Write the PST files, `pstcheck` Check every PST file (not with `-SkipCheck`) |
| ExportPst | `pst`, `pstcheck` |
| Import | `check` Verify the export store, `folders` Create the folders, `send` Send the items, `reconcile` Check the uncertain items, `final` Final check |
| ImportPst | the same five steps (`check` is Check the PST files) |
| ExportEml | `export` Write the .eml files (step events carry the items done and to do, and the bytes) |
| ImportEml | `send` Import the messages |
| CheckPst | `check` Check the file (step events carry `bytes` / `bytesTotal`: blocks verified of the file, and `done` / `total`: messages checked; `msg` is the phase) |
| RepairPst | `before` Check the file (before), `scan` Scan and read what can be read, `write` Write the new file, `after` Check the new file (after) (`scan` and `write` run together: they carry the same `done` / `total` messages) |
| Status, Browse | `run` |

### The folders of a run

Every job reports the progress of each folder as `object` events of the step that works on items: `download` (Export), `send` (Import, ImportPst, ImportEml), `pst` (ExportPst and Export with `-OutputPst`), `export` (ExportEml). `key` is the folder (for an import into a mailbox: the folder of the source or the folder of the target, see below), `state` `planned` (to do), `running`, `done`, `partial` (some items failed or were not exported), `check` (some are Unknown), `failed`, `refused` (a folder that is not exported, with the reason in `msg`) or `notready`; `done` and `total` are the items; `msg` says what failed when the folder is not simply done. An event is written when a folder changes (at most one per second for a job that runs in the calling thread, at each poll for the transfer jobs). The folders are the keys: Export and Import use the path of the folder in the export store (`Inbox/Projects`, with `Archive/` or `Auxiliary archive/` in front outside the primary mailbox), ImportPst the path of the folder in the target (below `-TargetRoot`), ExportPst the path in the PST (`Primary/Inbox`), ExportEml the path of the folder in the source, ImportEml the target folder and the path below it. With `-WhatIf` the same folders are `object` events of the step `plan`.

### Counts of the result

Export `done`, `failed`, `deferred`, `bytes` (+ `parts`, `itemsWritten` with PST). ExportPst `parts`, `itemsWritten`, `failed`. Import and ImportPst (the whole job, earlier runs included) `done`, `failed`, `unknown`, `pending`, `total`, `foldersCreated` (+ `refused` and `report`, the path of the list of the items not imported, for ImportPst). ExportEml `done`, `resumed`, `skipped`, `failed`, `bytes`, `folders`. ImportEml `done`, `resumed`, `failed`, `unknown`, `foldersCreated`. CheckPst `errors`, `warnings`, `infos`, `fileBytes`; `reason` `clean`, `warnings` or `errors`, `next` `done` or `repair`. RepairPst `recovered`, `partial`, `lost`, `orphan`, `levelUsed`, `recommended`, `destination`, `before` and `after` (`errors`, `warnings`).

### Reading the file

`[MailboxTransfer.EventTail]::Read(path, offset)` returns the complete new lines and the offset for the next call (a last line still being written is left for the next call; a file shorter than the offset is read again from zero). The file is opened with `FileShare.ReadWrite`: the console reads while the script writes.

## status.json

Written in the job folder at the start (`running`), refreshed while the job runs (at most every 15 seconds, when a step or folder event is written: `updated`, `lastEvent` and, when the counting step gives them, `counts.done` and `counts.total`), and updated at the end (the result closes it: nothing changes it afterwards); what the **Continue** list reads (`Get-MbtJob -WorkFolder <dir>` lists `<dir>\jobs\*\status.json`, newest first).

```json
{ "v": 1, "jobId": "imp-7c1e", "task": "import", "title": "Import - adele@contoso.com", "mode": "Import", "store": "D:\\MailboxTransfer\\stores\\adele",
  "state": "resumable", "exitCode": 2, "counts": { "done": 5120, "pending": 4738 }, "next": "continue", "reason": "stopped",
  "started": "2026-11-09T14:02:11.000Z", "updated": "2026-11-09T22:41:03.000Z", "lastEvent": "2026-11-09T22:41:03.000Z", "pid": 10432 }
```

`state`: `running`, `stopped` (stop request), `resumable` (exit 2), `needs-decision` (exit 3), `complete` (exit 0, and a check that ended, whatever it found), `error` (exit 1). `Get-MbtJob` adds `EffectiveState`: a job `running` whose process is gone (crash, power loss, killed) is `resumable` with `Interrupted = $true`; `Alive` is true when the process is still there ("running elsewhere").

## Stop, sleep, errors

- **Stop**: create the stop file. The script cancels as Ctrl+C does; Export, Import and the PST jobs finish the requests in flight and save their state; a check or a repair stops between two blocks or messages (a stopped repair writes no file). Exit 2, `result.reason = "stopped"`, `next = "continue"`.
- **Sleep**: a `slept` event after the wake-up. The engine retries the requests that timed out; the items whose answer was lost become Unknown and are looked up (`item` / `decision` events).
- **Sign-in**: with `-NoPrompt` a missing or expired sign-in ends the job resumable (exit 2, `signin-required`); the same error in the middle of a run has the same result.
- **Other errors**: an `error` event with the plain sentence and the raw message in `tech`, then `result` with `exit` 1; the exception is thrown as before (the script's exit code is 1).
