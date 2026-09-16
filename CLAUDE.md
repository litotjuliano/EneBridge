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

# Run all tests (64 tests; requires the Access Database Engine OleDb provider to be
# installed/registered on the machine, since several tests round-trip through real .dbf files).
# An unfiltered run can abort partway through with a native AccessViolationException in
# ComObject.Finalize -- a pre-existing ACE driver instability (see "Known reliability risk" in
# Current status below), not a real test failure. Filtering to a single class/test reduces how
# often this happens but doesn't eliminate it: DbfReaderServiceTests.Read_TableDoesNotExist_
# ReturnsFailureWithErrorMessage alone still triggers it intermittently (roughly half the time).
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
files, so this piece intentionally keeps the original's approach. `DbfExportService.Export` backs
up any existing `<table>.dbf` (renamed to a timestamped copy in the same folder,
`<table>_yyyyMMddHHmmss.dbf`, matching the original console app's exact convention — confirmed by
decompiling it) and always creates the table fresh via `CREATE TABLE`, then inserts this run's
rows — nothing from a previous run survives in the live file. This is a separate, in-folder backup
mechanism from `DbfSafetyBackupService`'s AppData copy (described in the next paragraph); both run
on every export. DBF field names are limited to 10 chars and `[A-Za-z0-9_]`;
`DbfSchemaBuilder.SanitizeColumnName` enforces this for both the generated `CREATE TABLE` and
`INSERT` column lists (the original only sanitized names for `CREATE TABLE`, not `INSERT`, which
was a latent bug — harmless today since no current column name needs sanitizing, but fixed here to
guard against future schema edits).

**DBF export crash prevention**: four independent layers guard against the ACE OleDb driver's
native-crash risk (see "Known reliability risk" under Current status, below) and the incident that surfaced it (`Microsoft.ACE.OLEDB.12.0`
silently resolving to Office Click-to-Run's own sandboxed `ACEOLEDB.DLL` instead of the standalone
redistributable — see `docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md`
for the full incident writeup). `DbfSafetyBackupService` copies `icmaste.dbf`/`ictrane.dbf` (plus
`.FPT` companion) to `%AppData%\eneBridge\Backups\<table>\` before anything else runs in
`ConfirmExportAsync` — plain `File.Copy`, independent of OleDb, so it works even if the driver
itself is broken. `AceEngineGuardService` spawns `eneBridge.Wpf.exe --selftest-ace` as a disposable
child process once per session (via `AceEngineSelfTestService`, a real DBF round-trip through a
throwaway table) so a native crash from a broken driver only kills that child, never the app;
`MainViewModel.CanExport` is gated on the result, with a "Fix it now" repair flow that re-runs the
bundled `accessdatabaseengine_X64.exe` elevated. `DbfWorkflowHelper.ConfirmTablesReadable` warns
(Continue/Cancel) if either table can't be read, checked before `DbfExportService.Export` is
called — deliberately checked via plain file I/O (`Directory.Exists`/`File.Exists`/a shared-read
`FileStream` probe) rather than `DbfReaderService.Read`, since a real OleDb `SELECT` against a
nonexistent table was found to intermittently crash the process with this same native-crash
signature even on an otherwise-healthy driver (see "Known reliability risk" below) — the check is
OleDb-free for the same reason `DbfSafetyBackupService` is.
`installer/eneBridge.iss`'s `IsAccessDatabaseEngineInstalled` now resolves the ACE provider's
actual bound DLL path (not just registry-key existence) so it correctly detects and fixes the
Office Click-to-Run case instead of silently skipping the real fix.

**Stock Received workflow**: a second, independent workflow (`StockReceivedViewModel`,
`MainWindow.xaml`'s "Stock Received" tab) reads `Self-Billed_Format-ORI.xlsx`-shaped workbooks
(`ExcelReaderService.ReadStockReceivedIcmaste`/`ReadStockReceivedIctrane`, driven by
`StockReceivedExcelCol`) and writes into the SAME `icmaste.dbf`/`ictrane.dbf` Invoice writes to,
distinguished by `TYPE` (`"RE"` vs Invoice's `"IN"`). See
`docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md` for the full field mapping
and rationale. `DbfExportService.Export`'s write behavior itself went through two iterations: it
briefly changed from "back up and recreate the table on every run" to "create the table only if
missing and always append" (so Invoice's and Stock Received's rows could coexist in the shared
files without one wiping the other), but this was reverted after real testing against an EMAS
installation found a worse problem — see `DbfExportService.Export`'s own doc comment for the full
story. It's back to backing up and recreating on every run now, exactly matching the original
console app's behavior (confirmed by decompiling it), with the accepted tradeoff that exporting one
workflow now wipes the other's not-yet-imported rows from the staging file if an EMAS import hasn't
run in between — safe only when each export is immediately followed by an EMAS import, which
matches how the app was always used before Stock Received existed. `MainViewModel`'s and
`StockReceivedViewModel`'s post-export verification (`DbfVerificationHelper.VerifyDbfAsync`)
correspondingly compares the table's current `TYPE`-filtered row count directly against
`RowsWritten`, not a before/after delta (that was only needed while rows accumulated indefinitely).
The ACE-driver-crash-avoidance table check lives in the Wpf-project `Services/DbfWorkflowHelper.cs`
(UI-coupled — shows a `MessageBox`); the verification/row-counting logic it's paired with lives in
the Core-project `Services/DbfVerificationHelper.cs` (pure, unit-tested). Both are shared between
the two ViewModels rather than duplicated, since they're safety-critical; everything else in
`StockReceivedViewModel` (read/export/preview plumbing) deliberately mirrors `MainViewModel`'s
Invoice-related surface as a separate, parallel implementation rather than a shared base class,
matching this project's existing philosophy of only sharing code across the two workflows when it's
safety-critical. A separate Core-project `Services/ExportGateService.cs` enforces mutual exclusion
between the two workflows' Confirm & Export commands: since both write into the same physical DBF
files, and the ACE OleDb driver has documented native-crash fragility under far milder stress (see
"Known reliability risk" below), letting both workflows export concurrently was judged a real,
newly-introduced crash risk once Stock Received started sharing Invoice's tables — a single shared
gate instance, injected into both ViewModels, blocks one workflow's export while the other's is in
progress and disables its Confirm & Export button (`CanExport`) accordingly. `StockReceivedViewModel`
has its own Excel path (`SettingsService.ResolveEffectiveStockReceivedPaths`, staged to
`stockreceived.xlsx`) and its own run history file (`%AppData%\eneBridge\stockReceivedRunHistory.json`),
but shares Invoice's `DbfFolderPath` (the two must point at the same folder for this to mean
anything) and the single app-wide ACE-driver health check (`MainViewModel` propagates its result to
`StockReceivedViewModel` via `SetAceEngineHealth` whenever it changes, rather than running a second
self-test).

**Wrong-tab file detection**: `ExcelReaderService.ValidateFileFormat` sanity-checks a workbook's
header row (row 1) against the tab it's being previewed on, before either format's row-by-row
column mapping runs. Both formats read by fixed physical column position (see `ExcelLayout`) with
no header-name dependency for the actual mapping, which means Browsing a Stock Received-shaped file
into the Invoice tab (or vice versa) previously misaligned every column silently instead of failing
loudly — confirmed against real client data: Invoice's DATE column position landed on Stock
Received's Item Code column, producing a confusing "DATE parse failed: 'sub-con'" per-row skip
reason instead of a clear error. `ValidateFileFormat` looks for each format's own distinguishing
header text anywhere in row 1 (`"Invoice Number"` / `"Stock Received No"`, taken from the real
sample workbooks) and throws `ExcelValidationException` before the mismatched read begins. Both
`MainViewModel.PreviewAsync` and `StockReceivedViewModel.PreviewAsync` call it right after opening
the workbook and catch that exception locally (rather than letting it fall through to the generic
"Outer backstop" catch) to show a `MessageBox` naming the mismatch, with no override — a wrong file
type is never a legitimate reason to proceed.

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
overrides the defaults) — but only for the DBF folder path. `ResolveEffectivePaths`/
`ResolveEffectiveStockReceivedPaths` deliberately never restore a remembered Excel file path, even
though `UserSettingsModel.ExcelFilePath`/`StockReceivedExcelFilePath` are still written on every
Browse (`SaveUserPaths`): a remembered Excel path looks "ready to export" the moment the app opens,
inviting an accidental re-export of a stale file instead of the user actively picking today's file.
Both `ExcelFilePath` properties always start blank; `Preview` against a blank path fails
gracefully (a clear "Cannot open Excel file" message, no crash — the button itself isn't gated on
the path being set) but produces no read results, so `Confirm & Export` stays disabled until a
fresh Browse and a successful Preview. The DBF folder path has no such risk (it rarely changes,
and pointing at the wrong DBF folder is caught by `DbfWorkflowHelper.ConfirmTablesReadable`
anyway), so it's still remembered normally. `RunHistoryService` persists run history to
`%AppData%\eneBridge\runHistory.json`. `FileLogger` writes full exception detail to a rolling log
file under `%AppData%\eneBridge\logs\`, kept separate from the concise on-screen run history.
`ExcelSourceStagingService` copies every Excel file picked via Browse into
`%AppData%\eneBridge\ExcelSource\` before `ExcelFilePath` is set, so later runs don't depend on
the original (possibly removable) media staying available. `StageFile` takes the target file name
as a parameter (not a hardcoded constant) — each workflow stages under its own fixed name in the
same shared folder: the Invoice tab always stages to `invoice.xlsx`
(`MainViewModel.InvoiceStagedFileName`); the Stock Received workflow stages to `stockreceived.xlsx`
(`StockReceivedViewModel.StageExcelSource`/`StockReceivedStagedFileName`). This rename-and-replace
happens only at import time (a new Browse), never during Run — a same-named existing staged file
is renamed with a timestamp suffix and kept as a permanent backup right there in the same
`ExcelSource` folder rather than a separate backup subfolder, never deleted, no retention limit
(unlike `DbfSafetyBackupService`, which backs up into its own separate `Backups\<table>\` folder —
the two don't mirror each other's layout, just the same never-delete philosophy), then the new file
takes its place immediately, so the live file is never left missing. Run itself only ever reads
the current staged file — it never renames, moves, or deletes it, so re-running repeatedly is
always safe. Staging failures are logged via `FileLogger` and fall back to using the originally
picked path directly rather than blocking the user.

## Current status

Core services and the WPF UI are implemented; 63 of 64 tests pass reliably (an unfiltered
`dotnet test` run can additionally abort partway through on the pre-existing ACE driver
instability described just below — not a real regression). The one exception,
`DbfReaderServiceTests.Read_TableDoesNotExist_ReturnsFailureWithErrorMessage`, triggers that same
instability intermittently even run alone (see "Known reliability risk" below) — its assertion is
correct when the process survives, it just doesn't always survive. Also included: a real DBF round-trip
through the actual OleDb/ACE provider (`DbfExportServiceTests`) and full-pipeline verification
against the real sample workbook (`tests/eneBridge.Wpf.Core.Tests/Fixtures/data.xlsx`) (icmaste: 110
rows read → 3 written; ictrane: 110 rows read → 110 written). The running app has been manually driven end-to-end (path selection, Preview, Confirm &
Export, run history) with correct results. The error-path testing (nonexistent Excel path, read-only DBF
folder, corrupted cell values) called for in the original verification plan is now covered in
`ExcelReaderServiceTests`, `DbfSchemaBuilderTests`, and `DbfExportServiceTests`. Whether EMAS's
Inventory Control module correctly handles `icmaste.dbf`/`ictrane.dbf` accumulating rows
indefinitely was tested against a real EMAS installation and found not to work safely — see "Stock
Received workflow" above for why `DbfExportService.Export` went back to backing up and recreating
the table on every run instead.

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
prevention" above). The underlying driver instability itself turned out to be broader than
"per-row INSERT + reused connection", though: a single failed `SELECT` against a table that
doesn't exist yet (`DbfReaderService.Read`, no INSERT, no connection reuse involved) was
independently confirmed to trigger the same native crash on this machine, intermittently, even
with an otherwise-healthy driver — `DbfReaderServiceTests.Read_TableDoesNotExist_ReturnsFailureWithErrorMessage`
reproduces it in isolation (crashed 2 of 3 runs). Because of this, `DbfWorkflowHelper.ConfirmTablesReadable`
(see "DBF export crash prevention" above) deliberately avoids `DbfReaderService.Read` entirely,
using plain file I/O instead. `VerifyDbfAsync`'s own post-export read-back still calls
`DbfReaderService.Read` and carries this same pre-existing exposure — untouched by this work,
flagged here for whoever picks it up next. A third trigger was confirmed against a real EMAS
installation during Stock Received testing: `DbfReaderService.Read` against `icmast.dbf` — a
Visual FoxPro table (version byte `0x30`), a materially different binary format from the plain
dBASE III files (`0x03`) this app writes, which `Provider=Microsoft.ACE.OLEDB.12.0` with
`Extended Properties="dBASE IV"` cannot parse at all — reliably threw `OleDbException: "External
table is not in the expected format"` (100% reproducible, not intermittent), and the app hung and
then closed on Confirm & Export reproducibly right after this failed read, on a machine that had
exported successfully moments earlier in the same session. See "Duplicate check reads EMAS's live
table" below for how this is now avoided entirely (a hand-written reader with no OleDb involved,
rather than a fix to `DbfReaderService`/the ACE driver itself).

**Duplicate check reads EMAS's live table via a hand-written Visual FoxPro parser, not OleDb.**
`icmast.dbf`/`ictran.dbf` (no trailing "e") are EMAS's own live/master tables, separate physical
files from `icmaste.dbf`/`ictrane.dbf` (`IcmasteSchema.TableName`/`IctraneSchema.TableName`) — the
staging files eneBridge writes and EMAS's Inventory Control module imports from — confirmed against
a real installation, along with the client confirming `icmast.dbf` shares `icmaste.dbf`'s column
layout (same `REF`/`CODE`/`TYPE` names). Checking `icmaste.dbf` instead (as an earlier iteration of
this feature briefly did, back when `Export` appended forever) would be wrong for two different
reasons depending on which `Export` behavior is current: while `Export` appended forever,
`icmaste.dbf` would keep reporting a document as a duplicate even after the user deletes it from
EMAS's live data (confirmed by the client hitting exactly this); now that `Export` is back to
backing up and recreating on every run (see "Stock Received workflow" above), `icmaste.dbf` only
ever holds the most recent run's rows, so it can't reliably tell you whether a document was ever
imported previously at all — only `icmast.dbf`, EMAS's actual live data, can answer that. Reading
`icmast.dbf` via `DbfReaderService`/OleDb was tried first and
reverted: its first byte (`0x30`) marks it as a **Visual FoxPro** table, a materially different
binary format from the plain dBASE III files (`0x03`) this app writes, which
`Provider=Microsoft.ACE.OLEDB.12.0` with `Extended Properties="dBASE IV"` cannot parse at all —
every attempt failed with `OleDbException: "External table is not in the expected format"`, 100%
reproducibly, and was strongly correlated with the native-crash class described two paragraphs
above (the app hung and closed on Confirm & Export, reproducibly, right after this failed read, on
a machine that had exported successfully moments earlier in the same session). The real fix:
`Services/FoxProDbfReader.cs` (Core project) hand-parses the classic dBASE/FoxPro free-table binary
layout directly — a 32-byte header, one 32-byte field descriptor per column, a single `0x0D`
terminator, then fixed-width records each prefixed by a 1-byte deletion flag — confirmed
byte-for-byte against a real `icmast.dbf` (every classic dBASE/FoxPro field type, numeric and date
included, stores its value as plain padded text, so every column is decoded uniformly as trimmed
text with no per-type logic needed). This is pure managed byte parsing with no COM/OleDb involved
at all, so it carries none of the ACE driver's native-crash risk. The "textbook correct" fix, the
legacy `VFPOLEDB` provider, was considered and rejected: VFPOLEDB is 32-bit only and this app
already depends on the 64-bit Access Database Engine, so using it would need either switching the
whole app to 32-bit or spawning a separate 32-bit helper process — real complexity `FoxProDbfReader`
avoids entirely, at the cost of implementing (a narrow slice of) the binary format by hand.
`DuplicateDocumentChecker.FindDuplicateRefsAsync` takes a `FoxProDbfReader` (not a
`DbfReaderService`) and reads `IcmasteSchema.LiveTableName` (`icmast.dbf`) — verified against the
real file (all active `RE`-type rows read correctly, the one deleted record correctly excluded) and
against a full Confirm & Export run with no crash, before being wired in.

**A REF only counts as a duplicate if BOTH `icmast.dbf` AND `ictran.dbf` have a match — not
`icmast.dbf` alone.** Real testing found EMAS's own document delete is a "soft delete" (per the
client) that removes a document's `ictran.dbf` line items but leaves its `icmast.dbf` header
behind, orphaned — confirmed directly on a real installation: a deleted `SB-000005` was fully
active in `icmast.dbf` (deletion flag clear, no void/cancel field set, every field still populated)
while `ictran.dbf` had zero matching rows for it, and EMAS's own Stock Received screen could no
longer find the document at all. An `icmast.dbf`-only check saw the leftover header and wrongly
blocked re-export of a document EMAS itself no longer considered to exist.
`DuplicateDocumentChecker.FindDuplicateRefsAsync` now also reads `IctraneSchema.LiveTableName`
(`ictran.dbf`) and only reports a REF as a duplicate when both tables agree — verified against the
real `SB-000005` case (correctly no longer blocked) via a throwaway harness before being wired in.
This incidentally also resolves the previously-documented gap where the check only read the header
table and could be fooled by a partial prior export (icmaste's row failed to write while ictrane's
succeeded, or vice versa): requiring both tables to agree catches that case too, in either
direction.

**Fixed: a table read failure (e.g. EMAS holding the file open) is no longer silently treated as
"no duplicates".** `FindDuplicateRefsAsync` used to return an empty duplicate list whenever
`FoxProDbfReader.Read`'s `Success` was `false` for either table, with no way for the caller to tell
"the table doesn't exist yet" (a genuine first-ever-export case, safe to proceed) apart from "the
table exists but couldn't be read for some other reason" (NOT safe to treat as "no duplicates" —
the check simply didn't run). This was found to be a real, confirmed production bug, not a rare
edge case: EMAS itself being open — an ordinary, frequent state, since the whole eneBridge workflow
revolves around switching between it and EMAS's own Import Transaction screen — can hold
`icmast.dbf`/`ictran.dbf` locked enough that `FoxProDbfReader.Read` fails, and a real duplicate
document (confirmed directly: a document already imported into EMAS) was exported a second time
with zero warning as a result. `DuplicateDocumentChecker.FindDuplicateRefsAsync` now returns a
`DuplicateCheckResult` (`Verified`, `DuplicateRefs`, `UnverifiableReason`) instead of a bare list:
`Verified` is `true` only when both tables were successfully read (or genuinely don't exist yet —
checked via `File.Exists` before ever calling `FoxProDbfReader.Read`, so "doesn't exist" and "exists
but unreadable" are distinguished up front rather than inferred from the read result).
`DbfWorkflowHelper.ConfirmNoDuplicateDocuments` now hard-blocks (no override, same as a real
duplicate) whenever `Verified` is `false`, telling the user the check couldn't run and to close
EMAS and retry — silently proceeding is no longer an option. `DbfWorkflowHelper.ConfirmTablesReadable`,
which runs earlier in the same `ConfirmExportAsync` flow, doesn't cover this case either — it
deliberately only does a plain-file-I/O readability probe (`File.Exists`/a shared `FileStream`
open) against `icmaste.dbf`/`ictrane.dbf` (the staging files), not `icmast.dbf`/`ictran.dbf` (the
live tables `FindDuplicateRefsAsync` reads), for the native-crash reasons explained above.
