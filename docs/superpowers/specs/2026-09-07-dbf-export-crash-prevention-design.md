# Prevent the ACE-Driver Crash Class During Export

## Context

CLAUDE.md already documents a known reliability risk: a native `AccessViolationException` from
the ACE OleDb dBASE driver's COM interop bypasses every managed `try`/`catch` in
`DbfExportService`/`DbfReaderService` and kills the app process outright — previously seen only
after a bad per-row `INSERT`.

This session investigated a live incident where Confirm & Export crashed on every attempt (7/7),
silently, with no exception logged. Root-cause investigation (via an isolated harness process,
never touching the live EMAS folder after an initial restore) found the *actual* trigger was
broader than previously known: `Microsoft.ACE.OLEDB.12.0` on that machine resolved to Office
2024 Click-to-Run's own copy of `ACEOLEDB.DLL` (under `...\root\VFS\...`), not the standalone
Access Database Engine redistributable this project depends on
(`prerequisites\accessdatabaseengine_X64.exe`). The Click-to-Run copy crashes natively on
`CREATE TABLE`/write, reproducibly, regardless of folder contents, row data, or file count —
confirmed by reproducing the crash against a completely empty scratch folder, and by the project's
own `Export_IcmasteAndIctrane_ProducesValidDbfFiles` test crashing its test host on the same
machine. Installing the standalone redistributable fixed it immediately (verified: the test now
passes, and the harness round-trips real invoice data successfully).

While tracing why this machine ended up with the wrong driver bound despite the installer already
being designed to install the correct one, a second, independent bug was found:
`installer/eneBridge.iss`'s `IsAccessDatabaseEngineInstalled` (lines 59-65) only checks that the
`Microsoft.ACE.OLEDB.12.0\CLSID` registry key *exists* — which Office's Click-to-Run install also
creates, pointing at its own DLL. On any machine with Office already installed, this check
silently reports "already installed" and the installer skips running the standalone redistributable
entirely. This is almost certainly how this machine ended up in the broken state, and it means any
user's machine with Office installed is exposed to the same silent gap today.

A third, unrelated gap surfaced during the incident: a crashed export had renamed `icmaste.dbf` to
a timestamped backup (per `DbfExportService.BackupExistingFile`'s existing "never delete" policy)
but never got far enough to recreate it, leaving EMAS's Inventory Control module unable to see the
table at all until it was manually restored from that backup. Nothing in the app today would have
told the user this happened short of noticing EMAS itself breaking.

`icmaste.dbf`/`ictrane.dbf` are EMAS's live Inventory Control tables — real accounting data, not
disposable output. `icmaste.dbf` surviving this incident at all was only because
`BackupExistingFile`'s same-folder rename happened to complete before the crash hit. That backup
lives in the exact same production folder as everything else, though: it's a *rename* (there's a
window where neither the old nor new file exists at that path), and it depends on that folder/drive
being fully healthy, same as the live file it's protecting.

## Goal

Reduce the chance of a repeat of this incident — and make it recoverable/diagnosable in-app if it
does happen — via four independent, complementary layers. None of them depend on each other; each
closes a different gap:

1. **Safety backup**: before touching either table, copy the current on-disk file to a separate,
   app-controlled location — independent of the live EMAS folder, and independent of whether the
   ACE driver even works, since it's a plain filesystem copy.
2. **Installer**: stop silently skipping the correct engine install on Office machines.
3. **Runtime self-test**: catch it if the driver breaks anyway (e.g. a later Office update changes
   the shared DLL while eneBridge is already installed), without ever risking the main app process.
4. **Pre-export table check**: catch folder/path/table-state problems unrelated to the driver
   itself (wrong folder, a table locked by EMAS, or state left over from a previous interrupted
   export — exactly what happened to `icmaste.dbf` in this incident) before the irreversible
   backup-rename step runs.

## Design

### 1. Safety backup to a separate app-controlled folder

At the very start of `ConfirmExportAsync`, before anything else runs (before the pre-export table
check in part 4, before either `ExportStageAsync` call) — for each of `icmaste`/`ictrane`, if
`<dbfFolder>\<tableName>.dbf` currently exists, copy it, plus `<dbfFolder>\<tableName>.FPT` if that
also exists (the memo companion file the ACE dBASE driver creates alongside a table; present today
for `icmaste`), to `%AppData%\eneBridge\Backups\<tableName>\<tableName>_<yyyyMMddHHmmss>.dbf`
(and `.FPT`) — same timestamp format `BackupExistingFile` already uses elsewhere. This is a plain
`File.Copy`, deliberately independent of `System.Data.OleDb`/the ACE driver entirely, so it works
even in the exact scenario that caused this incident (the driver itself being broken) and even if
`AceEngineHealthy` (part 3) is `false`. Only files with the table's exact base name are copied —
not EMAS's own same-prefixed index files (e.g. `icmastei.CDX`, `icmasteie.DBF`), which this app
doesn't own or manage.

This runs in addition to, not instead of, the existing same-folder rename in
`DbfExportService.BackupExistingFile` (still required — `CREATE TABLE` needs the filename free).
The two are intentionally redundant and independent: one lives next to the live data for
convenience, the other lives in a location with no dependency on the EMAS folder, drive, or the ACE
driver being healthy at all. Matches the project's existing "never delete, no retention limit"
backup policy (`DbfExportService`, `ExcelSourceStagingService`) — these files are small, so
unbounded retention isn't a practical disk-space concern.

If the source file can't be copied (e.g. genuinely locked), log the failure via `FileLogger` and
continue with the rest of `ConfirmExportAsync` as today — this is an extra safety net, not a new
way for a normal export to be blocked.

### 2. Installer detection fix

`IsAccessDatabaseEngineInstalled` changes from a key-existence check to resolving what DLL the
provider actually points at:

1. Read `Microsoft.ACE.OLEDB.12.0\CLSID`'s default value to get the CLSID GUID (check both the
   64-bit and `Wow6432Node` views, matching the existing function's pattern).
2. Read that CLSID's `InprocServer32` default value (again both hive views) to get the DLL path.
3. If no CLSID/InprocServer32 value is found at all, or the resolved path contains `\root\VFS\`
   (Office Click-to-Run's virtualization signature — confirmed present in the actual broken path
   seen on the investigated machine: `...\Microsoft Office\root\VFS\...\Office16\ACEOLEDB.DLL`),
   treat the engine as **not** properly installed and let `[Run]` install it.

`[Files]` and `[Run]` change so the redistributable survives past setup instead of being deleted:

- `[Files]`: `accessdatabaseengine_X64.exe` moves from `DestDir: "{tmp}"` with
  `Flags: deleteafterinstall` to `DestDir: "{app}\prerequisites"` with no delete flag — installed
  permanently alongside the app.
- `[Run]`: its `Filename` updates from `{tmp}\accessdatabaseengine_X64.exe` to
  `{app}\prerequisites\accessdatabaseengine_X64.exe` (Inno Setup copies `[Files]` before running
  `[Run]`, so the permanent path already exists by the time it's invoked).

This also gives the app a stable, known local path to re-invoke later for the "Fix it now" repair
flow in part 3.

### 3. Startup self-test (child process, once per session)

**Core, testable logic** — a new `AceEngineSelfTestService` in `eneBridge.Wpf.Core`: builds a
throwaway single-column `DataTable` (one row, one string column — not the real 151/80-column
schema, which isn't needed to prove the driver itself works), writes it via the real
`DbfExportService.Export` and reads it back via the real `DbfReaderService.Read` against a fresh
temp folder (`Path.Combine(Path.GetTempPath(), "eneBridge-selftest-" + Guid.NewGuid())`), confirms
the row round-tripped, then deletes the temp folder in a `finally`. Returns a simple success/failure
result. This exercises the exact same code path a real export does, so it fails exactly when a real
export would.

**Wpf, child-process invocation**: `App.xaml.cs`'s `OnStartup` checks `e.Args` for a hidden
`--selftest-ace` flag *before* constructing `FileLogger`/services/`MainWindow`. If present: run
`AceEngineSelfTestService`, call `Environment.Exit(0)` on success or a distinct non-zero code on
failure or any exception, and return — no window, no dispatcher, no normal startup path at all.

On a normal launch (no self-test flag), after showing `MainWindow`, spawn
`Process.Start(Environment.ProcessPath!, "--selftest-ace")` and await its exit (with a bounded
timeout, e.g. 15s, to guard against a genuine hang rather than a crash) on a background thread —
this does not block the UI from appearing. A crash in that child process (the native
`AccessViolationException`) only kills the child; the parent just observes a non-zero/abnormal exit
code, exactly the isolation this design is for.

The result is exposed as a new bindable `MainViewModel.AceEngineHealthy` (default `true` optimistically
while the check runs, so the UI isn't blocked; `CanExport()` gains `&& AceEngineHealthy`). On
failure, show a `MessageBox` (same direct-from-ViewModel pattern `BrowseExcel`/`BrowseDbfFolder`
already use) explaining the Access Database Engine isn't working correctly, with **Fix it now** /
**Not now** buttons:

- **Fix it now**: launch `{app}\prerequisites\accessdatabaseengine_X64.exe /quiet /norestart` via
  `ProcessStartInfo { Verb = "runas", UseShellExecute = true }` (triggers a UAC prompt — the
  installer requires admin, and this repair does too), wait for exit, then re-run the child-process
  self-test. If it now passes, set `AceEngineHealthy = true` and show a brief confirmation. If it
  still fails, say so and suggest checking the Office installation or contacting support — don't
  loop automatically.
- **Not now**: dismiss; `AceEngineHealthy` stays `false`, Confirm & Export stays disabled with a
  status message explaining why, for the rest of the session (matches the "once per session" choice
  — no repeated interruption).

All of this is logged via `FileLogger` (self-test start/result, repair attempt/result) for the same
reason the existing per-row/per-connection checkpoints are: if something still goes wrong, the last
logged line should say what.

### 4. Pre-export table check

Runs immediately after the safety backup (part 1), still before either `ExportStageAsync` call
(i.e. before `BackupExistingFile` ever renames anything): call `_dbfReaderService.Read` for both
`IcmasteSchema.TableName` and `IctraneSchema.TableName` against the target `dbfFolder`. This only
runs once the driver itself is already known-good (`AceEngineHealthy` gates `CanExport()`), so it's
safe to do in-process.

If either read fails (`Success == false` — covers "file doesn't exist", "table locked by EMAS", and
"folder unreachable" alike, since `DbfReaderService.Read` already catches `OleDbException`/
`IOException`/`UnauthorizedAccessException` uniformly), show a `MessageBox` naming the table and the
underlying error, e.g.:

> *"icmaste.dbf could not be found or read in 'C:\emas\ch2026\acc\data':
> \<error message>.
>
> This is expected on a first-ever export to this folder, but if you expect this table to already
> exist, something may be wrong (wrong folder selected, the table is locked by EMAS, or a previous
> export was interrupted). Continue anyway?"*

with **Continue** / **Cancel** buttons. Continue proceeds into the export exactly as today; Cancel
aborts immediately — `IsRunning = false`, commands re-enabled, no files touched. This check runs for
both tables regardless of whether the first one fails, so a user who continues sees both warnings
(if both apply) rather than only the first.

## Error handling

Follows the project's existing "nothing crashes or fails silently" principle throughout: the
self-test's own child process can crash freely (that's the point — it's disposable), but the parent
process never calls anything that could trigger the native fault directly; every new failure path
ends in either a logged, actionable dialog or a safe abort, never an unhandled state.

## Testing

- The safety-backup copy logic (part 1) is plain file I/O against a scratch folder — fully
  unit-testable with no ACE provider involved: cover the table-exists (copies `.dbf` and, if
  present, `.FPT`, with the expected timestamped name), table-missing (no-op, no error), and
  companion-file-absent cases.
- `AceEngineSelfTestService`'s round-trip logic is unit-testable the same way
  `DbfExportServiceTests`' existing round-trip test is (real ACE provider required, same as all
  DBF-touching tests today) — add a passing-case test.
- The child-process spawning, exit-code interpretation, elevated repair flow, and the Inno Setup
  Pascal script changes are **not** meaningfully unit-testable (cross-process/OS-level, same
  reasoning already given in this project for why the native-crash risk itself has no in-process
  test). These need manual verification, called out explicitly rather than left implicit:
  1. Installer fix: inspect the resolved `InprocServer32` path before/after on a machine with only
     Office installed (as done live during this session's investigation) and confirm the new check
     correctly identifies the Click-to-Run path as "not installed" where the old one didn't.
  2. Self-test: already effectively verified live this session — the same `DbfExportService`/
     `DbfReaderService` calls reliably distinguished the broken vs. fixed driver in an isolated
     process.
  3. Pre-export check: manually create a scratch DBF folder missing one table, run Confirm &
     Export, confirm the dialog appears; Cancel touches nothing, Continue proceeds and creates the
     table normally.

## Out of scope

- Running *every* real export in a child process (a much larger architectural change than what's
  needed here — the self-test only needs to prove the driver works once per session; real exports
  stay in-process as today).
- Auto-retry/polling anywhere in this design.
- The separate, already-documented "a failed connection poisons the ACE engine for every later
  OleDb connection in the same process" reliability risk from CLAUDE.md's Current Status section —
  orthogonal, pre-existing, not addressed by this work.
