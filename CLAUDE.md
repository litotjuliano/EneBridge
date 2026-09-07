# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project is

eneBridge is a Windows desktop rebuild of a lost-source .NET console app. The original app is a
data bridge for **EMAS**, a legacy Malaysian accounting system: it reads transaction data out of an
Excel workbook and writes it into two DBF files (`icmaste.dbf`, `ictrane.dbf`) that EMAS's
Inventory Control module imports. `icmaste` is the one-row-per-document master table (151 columns);
`ictrane` is the many-rows-per-document transaction detail table (80 columns).

The original console app's source was gone; its behavior was fully reverse-engineered from an
`ildasm` disassembly of the compiled DLL (the PDB preserved method/variable names) and validated
against the real sample workbook (preserved today as `tests/eneBridge.Wpf.Core.Tests/Fixtures/data.xlsx`).
This rebuild is a WPF GUI with the same core pipeline plus quality-of-life features (path pickers,
data preview, run history, real error reporting instead of the original's silent failures/crashes).

- `prerequisites/` — installers required on any machine running the app: the Access Database
  Engine (OleDb DBF/Excel provider) and the .NET 8 Desktop Runtime.

## Commands

```
# Build everything
dotnet build "src/eneBridge.Wpf.slnx"

# Run all tests (56 tests; requires the Access Database Engine OleDb provider to be
# installed/registered on the machine, since several tests round-trip through real .dbf files)
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj"

# Run a single test
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~ReadIcmaste_DedupesByRef_OneRowPerDocument"

# Run the app
dotnet run --project "src/eneBridge.Wpf/eneBridge.Wpf.csproj"

# Build the installer (requires Inno Setup 6's iscc.exe on PATH or its install location;
# see docs/superpowers/specs/2026-08-16-deployment-installer-design.md for the full design)
dotnet publish "src/eneBridge.Wpf/eneBridge.Wpf.csproj" -c Release -o "installer/publish"
iscc "installer/eneBridge.iss"
# Produces installer/Output/eneBridge-Setup.exe
```

Both `eneBridge.Wpf.Core` and the WPF app target `net8.0-windows` (Core is windows-targeted too,
since `System.Data.OleDb` is Windows-only and this app has no cross-platform use case).

## Architecture

**Two-project split**: `src/eneBridge.Wpf.Core` (schema/mapping/export logic, no WPF dependency,
independently testable) and `src/eneBridge.Wpf` (MVVM UI, thin — orchestration only). No DI
container; `App.xaml.cs OnStartup` is the composition root that constructs services by hand and
passes them into `MainViewModel`.

**Data model is `DataTable`-based, not POCOs** (`Models/IcmasteSchema.cs`, `Models/IctraneSchema.cs`).
With 151/80 mechanical columns, a `DataTable`'s `DataColumnCollection` is already the natural input
the CREATE TABLE/INSERT SQL generators need, and it binds directly to the WPF preview `DataGrid`
with no extra glue. Column names actually touched by mapping code are exposed as `const string`
members on the schema classes (e.g. `IcmasteSchema.Ref`) so mapping code never uses raw string
literals for column names — the full column order/type list must stay byte-for-byte exact, since
that's what the downstream EMAS import expects.

**Excel reading** (`Services/ExcelReaderService.cs`) uses ClosedXML, reading explicit physical
row/column positions rather than treating the workbook like a database table (which is what the
original OleDb-based app did, and which silently assumed physical row 1 was a header). The
row-skip offset is centralized in `Models/ExcelColumnMap.cs`'s `ExcelLayout` (first row actually
processed = physical row 2 for the current invoice template), and the raw Excel column indices used by each field are centralized
in `IcmasteExcelCol`/`IctraneExcelCol` in the same file — these are the single source of truth for
"col[N]" mapping rules, reverse-engineered from the original IL and confirmed against the real
sample workbook (`tests/eneBridge.Wpf.Core.Tests/Fixtures/data.xlsx`).

The current invoice template has only 9 columns (`IctraneExcelCol.MinColumnCount = 9`,
`IcmasteExcelCol.MinColumnCount = 8`); 3 columns an earlier template carried (an unused one,
`FcRate`, `TaxCode`) are absent from it. When `FcRate`/`TaxCode` come back blank — the column is
missing entirely, or the specific cell is empty — `ExcelReaderService` defaults them to `0` /
`"SST0"` rather than leaving them null/blank, matching the literal values every row of the earlier
template always carried; a blank tax code isn't valid for the downstream EMAS/SST import.

`OpenWorkbook` opens the file via its own `FileStream` requesting `FileShare.ReadWrite` (rather
than ClosedXML's default, more restrictive sharing mode) and reads it into memory before handing
it to ClosedXML, so it succeeds even while the source file is already open in Microsoft Excel — a
common real workflow (checking the invoice before running the tool). A genuine exclusive lock is
still detected (via `HResult == 0x80070020`, `ERROR_SHARING_VIOLATION`) and surfaces a clear,
actionable error message rather than a raw exception.

Row validation differs deliberately between the two tables:
- `icmaste`: skips rows with a blank REF/DATE/CODE/NAME, and **dedupes by REF** — correct here
  because REF is a document number and icmaste is one-row-per-document.
- `ictrane`: skips rows with blank required fields but does **not** dedupe by REF — REF legitimately
  repeats across many transaction lines per document in this table. (Confirmed against real data:
  naively copying icmaste's REF-dedup onto ictrane would silently collapse ~97% of real
  transaction rows — this was caught during planning by directly inspecting the sample workbook's
  raw XML, not assumed.)

**Known limitation**: `ictrane.desc1` mirrors `item_no`'s source Excel column
(`IctraneExcelCol.ItemNoAndDesc1`) rather than a distinct description field. This reproduces a bug
in the original app; the correct source column could not be determined from the one available
sample workbook, so it's kept as-is and flagged both in code (doc comment on
`IctraneExcelCol.ItemNoAndDesc1`) and in the UI (a persistent banner on the ictrane preview tab).
Don't silently "fix" this without confirming the correct source column with whoever maintains the
Excel template.

**Most schema columns are intentionally always blank**: `ExcelReaderService.ReadIcmaste` populates
only 20 of icmaste's 151 columns (13 with real/derived values, 7 explicitly set to `DBNull`);
`ReadIctrane` populates only 12 of ictrane's 80 columns. The remaining ~131/68 columns (including
all 15 `ictrane.dfieldN` columns) are part of the required byte-for-byte DBF schema but are never
assigned a value, so they always render blank in the preview grid and in the exported `.dbf`
files. This was directly verified against the original compiled app's IL (recovered from git
history and re-disassembled with `ildasm`) — the original populated the exact same set of columns
and left the rest schema-only too, so this is not a gap introduced by the rebuild.

**DBF export** (`Services/DbfExportService.cs`, `Services/DbfSchemaBuilder.cs`,
`Services/SqlValueFormatter.cs`) writes via OleDb through the Access Database Engine's dBASE IV
driver (`Provider=Microsoft.ACE.OLEDB.12.0`) — there's no good managed alternative for writing DBF
files, so this piece intentionally keeps the original's approach. If a `<table>.dbf` already exists
in the target folder, it's renamed with a timestamp suffix as a backup (never deleted, no retention
limit) before a fresh file is created. DBF field names are limited to 10 chars and `[A-Za-z0-9_]`;
`DbfSchemaBuilder.SanitizeColumnName` enforces this for both the generated `CREATE TABLE` and
`INSERT` column lists (the original only sanitized names for `CREATE TABLE`, not `INSERT`, which
was a latent bug — harmless today since no current column name needs sanitizing, but fixed here to
guard against future schema edits).

**DBF export crash prevention**: four independent layers guard against the ACE OleDb driver's
native-crash risk described above and the incident that surfaced it (`Microsoft.ACE.OLEDB.12.0`
silently resolving to Office Click-to-Run's own sandboxed `ACEOLEDB.DLL` instead of the standalone
redistributable — see `docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md`
for the full incident writeup). `DbfSafetyBackupService` copies `icmaste.dbf`/`ictrane.dbf` (plus
`.FPT` companion) to `%AppData%\eneBridge\Backups\<table>\` before anything else runs in
`ConfirmExportAsync` — plain `File.Copy`, independent of OleDb, so it works even if the driver
itself is broken. `AceEngineGuardService` spawns `eneBridge.Wpf.exe --selftest-ace` as a disposable
child process once per session (via `AceEngineSelfTestService`, a real DBF round-trip through a
throwaway table) so a native crash from a broken driver only kills that child, never the app;
`MainViewModel.CanExport` is gated on the result, with a "Fix it now" repair flow that re-runs the
bundled `accessdatabaseengine_X64.exe` elevated. `MainViewModel.ConfirmTablesReadableAsync` warns
(Continue/Cancel) if either table can't be read before the export's own backup-rename runs.
`installer/eneBridge.iss`'s `IsAccessDatabaseEngineInstalled` now resolves the ACE provider's
actual bound DLL path (not just registry-key existence) so it correctly detects and fixes the
Office Click-to-Run case instead of silently skipping the real fix.

**Error handling**: nothing should ever crash the app or fail silently (the original did both,
depending on which stage failed — see the git history / design notes for details if needed). Excel
open failures and column-count validation failures are fatal for a stage; per-row parse/blank-field
issues are recorded as a `RowSkipReason` and the row is skipped, not fatal; DBF export failures are
fatal for that stage only, with per-row INSERT failures caught individually so one bad row doesn't
lose the rest of the file. `MainViewModel` splits reading and writing into two explicit,
user-triggered steps rather than one combined action: `PreviewAsync` (bound to the "Preview"
button) reads both tables, populates the preview grids, and logs skip reasons without writing
anything, giving a pre-verify checkpoint before any DBF file is touched; `ConfirmExportAsync`
(bound to "Confirm & Export", only enabled once `PreviewAsync` has produced read results) writes
the DBF files from those already-read results. Both always run icmaste and ictrane together (one
doesn't block the other except when the workbook itself can't be opened at all), and both wrap
their work in an outer catch as a backstop. Changing either path field after a Preview invalidates
the cached read results, disabling Confirm & Export until a fresh Preview runs.

**Persistence**: `SettingsService` reads `appsettings.json` (initial default paths, shipped with
the app) and a separate `%AppData%\eneBridge\settings.json` (last-used paths, user-writable,
overrides the defaults). `RunHistoryService` persists run history to
`%AppData%\eneBridge\runHistory.json`. `FileLogger` writes full exception detail to a rolling log
file under `%AppData%\eneBridge\logs\`, kept separate from the concise on-screen run history.
`ExcelSourceStagingService` copies every Excel file picked via Browse into
`%AppData%\eneBridge\ExcelSource\` before `ExcelFilePath` is set, so later runs don't depend on
the original (possibly removable) media staying available. `StageFile` takes the target file name
as a parameter (not a hardcoded constant) — each workflow stages under its own fixed name in the
same shared folder: the Invoice tab always stages to `invoice.xlsx`
(`MainViewModel.InvoiceStagedFileName`); a future Stock Received workflow is expected to use
`stockreceived.xlsx`. This rename-and-replace happens only at import time (a new Browse), never
during Run — a same-named existing staged file is renamed with a timestamp suffix and kept as a
permanent backup (mirroring `DbfExportService`'s backup-on-collision behavior, same "never
deleted" retention policy, same folder rather than a separate backup subfolder), then the new file
takes its place immediately, so the live file is never left missing. Run itself only ever reads
the current staged file — it never renames, moves, or deletes it, so re-running repeatedly is
always safe. Staging failures are logged via `FileLogger` and fall back to using the originally
picked path directly rather than blocking the user.

## Current status

Core services and the WPF UI are implemented; all 56 tests pass, including a real DBF round-trip
through the actual OleDb/ACE provider (`DbfExportServiceTests`) and full-pipeline verification
against the real sample workbook (`tests/eneBridge.Wpf.Core.Tests/Fixtures/data.xlsx`) (icmaste: 110
rows read → 3 written; ictrane: 110 rows read → 110 written). The running app has been manually driven end-to-end (path selection, Preview, Confirm &
Export, run history) with correct results. The error-path testing (nonexistent Excel path, read-only DBF
folder, corrupted cell values) called for in the original verification plan is now covered in
`ExcelReaderServiceTests`, `DbfSchemaBuilderTests`, and `DbfExportServiceTests`.

**Known reliability risk, not yet fixed**: `DbfExportService.Export` catches a per-row
`OleDbException` during INSERT and keeps reusing the same `OleDbConnection` for the remaining rows
(by design — see the DBF export section above). While adding tests for this path, a per-row INSERT
failure (e.g. a value too wide for its DBF field) was found to reliably crash the process afterward
with a native `AccessViolationException` in `ComObject.Finalize` — reproduced both in the test host
and in a standalone STA console harness mirroring `MainViewModel.ConfirmExportAsync`'s shape (one `Export`
call that hits a row error, followed by a second, separate `Export` call in the same process — i.e.
icmaste then ictrane in one run). This appears to be an instability in the ACE OleDb dBASE driver's
COM interop under .NET 8 when a connection that absorbed a row-level error is later followed by
another OleDb connection in-process, not something caused by test code. No test exercises this path
in-process (it isn't safe to), and no production fix has been attempted yet — flagged here so it
isn't lost. If a real row fails during a run, the *other* stage's export in the same run could be at
risk of crashing the app outright instead of failing gracefully, which would undermine this
project's "nothing should ever crash the app" goal. Separately, a distinct and more common trigger
for this same crash class — the ACE OleDb provider resolving to Office Click-to-Run's sandboxed
DLL instead of the standalone redistributable — was found and mitigated (see "DBF export crash
prevention" above); this per-row/reused-connection risk remains open.
