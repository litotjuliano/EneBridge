# Stock Received Workflow

## Context

The Stock Received tab has been a placeholder since it was first added ("Coming soon — awaiting
requirements from client."). The client has now supplied the requirements as two files:

- `Self-Billed_Format-ORI.xlsx` — the source workbook shape: one header row, then one row per
  self-billed stock-received document (`No, Stock Received No, Date, Supplier Code, Supplier Name,
  Item Code, Item Description, Qty, Unit Price, Total Amount`). Sample rows are subcontractor wage
  self-billings, each with a distinct `Stock Received No`.
- `Reference.xlsx` — an explicit column → EMAS field → table mapping, including which values are
  read from Excel versus fixed constants.

Working through the mapping surfaced two decisions with real consequences beyond just "add a new
tab":

1. Every Excel row is its own standalone document — confirmed no grouping/dedup like Invoice's
   REF-dedup-in-icmaste behavior is needed here.
2. Stock Received needs to write into the **same** `icmaste.dbf`/`ictrane.dbf` Invoice already
   writes to (distinguished by `TYPE`), not a separate file. `DbfExportService.Export` today always
   backs up and fully recreates the table on every run — if Stock Received simply used that as-is,
   every Invoice run afterward would wipe out whatever Stock Received had written. Rather than
   design around that footgun, this work changes the export write path itself: **both workflows
   now append** into the shared tables instead of recreating them. This is a real behavior change
   to Invoice's already-shipped export path, not just new code for Stock Received — confirmed and
   accepted as part of this design.

## Goal

Ship a working Stock Received tab that reads `Self-Billed_Format-ORI.xlsx`-shaped workbooks and
writes into `icmaste.dbf`/`ictrane.dbf` correctly and safely alongside Invoice, on the same
append-based write path both workflows now share.

## Design

### 1. Field mapping

Column indices are 0-based (`A` = 0), header row 1, data from row 2 — same layout convention as
Invoice's `ExcelLayout`.

| Excel column | Index | EMAS field | Table |
|---|---|---|---|
| No | 0 | *(unused — not mapped to any field)* | — |
| Stock Received No | 1 | REF | icmaste, ictrane |
| Date | 2 | DATE | icmaste |
| Supplier Code | 3 | CODE | icmaste |
| Supplier Name | 4 | NAME | icmaste |
| Item Code | 5 | ITEM_NO | ictrane |
| Item Description | 6 | DESC1 | ictrane |
| Qty | 7 | DESC2 (as literal text, not the numeric qty field) | ictrane |
| Unit Price | 8 | N_AMT | icmaste |
| Total Amount | 9 | T_AMT | icmaste |

Fixed constants (not read from Excel — `Self-Billed_Format-ORI.xlsx` has no source column for
these at all):

| Field | Value | Table |
|---|---|---|
| TYPE | `"RE"` | icmaste, ictrane |
| ACCNO | `"3020/000"` | icmaste |
| POSTACCNO | `"6010/000"` | icmaste |
| TAXCODE | `"SST0"` | icmaste |
| VENDNO | *(= Supplier Code)* | icmaste |
| CUST2 | *(= Supplier Code)* | icmaste |

Invoice's TYPE stays `"IN"` — this, alongside REF, is exactly what now distinguishes the two
workflows' rows inside the shared tables.

Matching Invoice's existing defaults for fields Reference.xlsx doesn't mention (confirmed):
`ENTRY`/`entry` = truncated REF, `USER`/`userid` = `"admin"`, `CURRCODE` = `"MYR"`.

`ictrane.price`/`ictrane.amount` are left blank (confirmed) — Reference.xlsx maps Unit
Price/Total Amount only to icmaste's N_AMT/T_AMT, not to ictrane's own numeric fields.

A row is skipped (from both tables together — there is no per-table divergence here, unlike
Invoice) if REF, DATE, CODE, NAME, or ITEM_NO is blank. Since every valid row produces exactly one
icmaste row and one ictrane row with no deduping, `RowsRead`/`SkipReasons` are identical between
the two tables for a given run.

### 2. Export write path: append-always, both workflows

`DbfExportService.Export` changes from "always back up and recreate" to a single unconditional
behavior: if `<dbfFolder>\<tableName>.dbf` doesn't exist yet, `CREATE TABLE` first (same DDL as
today); either way, `INSERT` this run's rows into it. No mode flag, no configuration — this is now
just how `Export` works, for every caller.

`BackupExistingFile` and its rename-to-timestamped-backup step are removed entirely (nothing calls
them anymore). `DbfSafetyBackupService`'s copy-based backup to `%AppData%\eneBridge\Backups\` —
already independent of OleDb, already running before anything else touches the tables — becomes
the sole automatic backup mechanism for both workflows. Since it's a plain `File.Copy` taken before
every export run regardless of outcome, it still provides a recovery point even though the live
file is never fully replaced anymore.

**Consequence to be explicit about**: there is no longer a "start fresh" export. Every Confirm &
Export run, on either tab, adds to what's already live in `icmaste.dbf`/`ictrane.dbf`. Resetting
the live tables (if ever needed) is a manual operation outside the app — renaming or deleting the
`.dbf` files before the next run.

This needs real-world verification against actual EMAS behavior once built: does EMAS's Inventory
Control module handle rows accumulating between its own reads/imports correctly? That can't be
confirmed from this codebase alone and is flagged as a required post-ship check, not a guarantee.

### 3. Verification: before/after delta, TYPE-filtered

`MainViewModel.VerifyDbfAsync`'s post-export check changes from "total rows in file == rows
written" (which breaks once rows accumulate across runs) to a **before/after delta**: read the
row count filtered to this workflow's `TYPE` value immediately before the export's `INSERT` loop
runs, read it again immediately after, and confirm the count grew by exactly `RowsWritten`. This
replaces today's check for both Invoice and Stock Received — accumulation is now universal, not
Stock-Received-specific.

The preview grid populated by this read-back (`IcmasteDbfPreview`/`IctraneDbfPreview`) is also
filtered to the current workflow's `TYPE`, so a user reviewing what Confirm & Export just wrote
sees only their own rows, not the other workflow's accumulated history mixed in.

### 4. `DbfReaderService.Read` hardening

Two changes to `DbfReaderService.Read`, both motivated by the before/after delta needing a "before"
read that can legitimately run against a table that doesn't exist yet (a genuine first-ever run):

- **File-existence guard**: check `File.Exists` (plain I/O) before opening any OleDb connection.
  If the table's `.dbf` file isn't there, return a clean `Success = false` result immediately,
  never touching OleDb. This is the exact same defensive pattern `MainViewModel.ConfirmTablesReadable`
  already uses (see `docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md`) to
  avoid the native-crash risk of an OleDb `SELECT` against a nonexistent table — without this
  guard, reading the "before" count on a first-ever run would hit the identical crash class that
  work was written to close. Folding it into `Read` itself (rather than requiring every caller to
  remember a `File.Exists` check first) also hardens the pre-existing Invoice verification call
  site for free.
- **Optional TYPE filter**: `Read` gains an optional parameter that, when supplied, adds
  `WHERE TYPE = '<value>'` to the generated `SELECT` (matching each schema's actual column casing
  — `TYPE` for icmaste, `type` for ictrane). Used by both the before/after delta counts and the
  filtered preview grid described above.

### 5. Settings persistence: two Excel paths, one shared DBF path

`UserSettingsModel` currently has one `ExcelFilePath` field — Stock Received needs its own,
independent from Invoice's, while both share the single `DbfFolderPath` (see §2). This surfaces a
real bug risk in the current save pattern: `MainViewModel.SaveUserPaths` builds a fresh
`UserSettingsModel` from only the fields it knows about and calls `SaveUserSettings`, which
serializes and fully overwrites `settings.json`. If `StockReceivedViewModel` did the same with its
own fresh object, saving Stock Received's path would silently blank out Invoice's already-saved
path the next time settings load, and vice versa.

Fix: `SettingsService.SaveUserSettings` changes from taking a complete `UserSettingsModel` to
taking a mutation delegate — it loads whatever is currently persisted (or starts fresh if nothing
is saved yet), applies the caller's changes, and writes the merged result back. Each ViewModel's
save call only touches the fields it owns:

```csharp
public void SaveUserSettings(Action<UserSettingsModel> update)
{
    var current = LoadUserSettings() ?? new UserSettingsModel();
    update(current);
    // ... serialize `current` as today
}
```

`UserSettingsModel` gains `string? StockReceivedExcelFilePath`. `AppSettingsModel.FilePathsSection`
gains a matching `StockReceivedExcelFilePath` default (empty string if `appsettings.json` doesn't
define one — no app-shipped default is required for this to work, it just means the field starts
blank until the user Browses once, same as any other first run). `SettingsService` gains
`ResolveEffectiveStockReceivedPaths`, mirroring `ResolveEffectivePaths` but resolving
`StockReceivedExcelFilePath` independently while resolving `DbfFolderPath` identically to Invoice's
(same user-saved value, same fallback).

### 6. New/changed files

- `Models/StockReceivedExcelCol.cs` — 0-based column indices per the mapping table above;
  `MinColumnCount = 10` (through Total Amount at index 9).
- `Services/ExcelReaderService.cs` — new `ReadStockReceivedIcmaste`/`ReadStockReceivedIctrane`,
  sharing the row-validation logic described in §1 (extracted to a private helper so the five-field
  blank check isn't duplicated between the two methods).
- `Services/DbfExportService.cs` — `Export` rewritten per §2; `BackupExistingFile` removed.
- `Services/DbfReaderService.cs` — `Read` changed per §4.
- `Models/AppSettingsModel.cs` / `SettingsService.cs` — `UserSettingsModel.StockReceivedExcelFilePath`,
  `AppSettingsModel.FilePathsSection.StockReceivedExcelFilePath`, `SaveUserSettings`'s merge-based
  signature, and `ResolveEffectiveStockReceivedPaths`, all per §5.
- `ViewModels/StockReceivedViewModel.cs` (new) — mirrors `MainViewModel`'s Invoice-related surface
  (Excel/DBF path fields, Preview/Confirm & Export commands, progress state, preview grids),
  constructed in `App.xaml.cs`'s composition root alongside `MainViewModel` and reusing the same
  service instances (`ExcelReaderService`, `DbfExportService`, `DbfReaderService`,
  `DbfSafetyBackupService`, `AceEngineGuardService`, `FileLogger`, `SettingsService`).
  `DbfFolderPath` is the same value Invoice uses (not a separate per-tab path — see §2, they must
  point at the same files for append to be meaningful); `ExcelFilePath` stays independent per tab
  (see §5), staged to its own `stockreceived.xlsx` (already anticipated by
  `ExcelSourceStagingService`'s design, per `CLAUDE.md`).
- `ViewModels/MainViewModel.cs` — two changes: gains a
  `public StockReceivedViewModel StockReceivedViewModel { get; }` property, set in its constructor,
  so the window's single root `DataContext` (still `MainViewModel`, unchanged) can reach it
  declaratively (no `MainWindow.xaml.cs` code-behind changes needed); and `SaveUserPaths` updated to
  call `SaveUserSettings`'s new delegate-based signature (§5). No other Invoice-side behavior
  changes.
- `MainWindow.xaml` — Stock Received tab gets real content (file-paths panel, run panel with its
  own Preview/Confirm & Export buttons, progress panel, preview grids), replacing today's
  placeholder text. The tab content's root element sets
  `DataContext="{Binding StockReceivedViewModel}"` so its bindings resolve against the nested
  ViewModel instead of `MainViewModel`.
- `tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs` — existing backup-rename assertions
  removed (the behavior they tested no longer exists); new cases added for create-if-missing and
  append-accumulates-rows-across-calls.
- `tests/eneBridge.Wpf.Core.Tests/DbfReaderServiceTests.cs` — new cases for the file-existence
  guard (already partially covered by the existing "table does not exist" test — confirm it still
  passes under the new guard, which changes *how* that result is produced but not what it asserts)
  and the TYPE filter.
- New `ExcelReaderServiceTests` cases for the Stock Received read methods, following the same
  fixture-based pattern already used for `ReadIcmaste`/`ReadIctrane`.

## Testing

- `ExcelReaderService`'s new methods: unit-testable the same way `ReadIcmaste`/`ReadIctrane`
  already are — a fixture workbook shaped like `Self-Billed_Format-ORI.xlsx`, asserting correct
  field mapping, the fixed constants, and the shared skip logic.
- `DbfExportService.Export`'s new behavior: real round-trip tests (same ACE-provider requirement as
  today) — create-if-missing on an empty folder, then a second `Export` call against the same
  folder asserting the row count grew by exactly the second call's `RowsWritten` rather than the
  file being replaced.
- `DbfReaderService.Read`'s file-existence guard and TYPE filter: real round-trip tests — read
  against a missing file (asserts clean failure, no crash risk to reason about since this is a
  pure `File.Exists` check, no OleDb involved in that path), and read with/without a TYPE filter
  against a table containing rows of more than one TYPE value.
- `StockReceivedViewModel` and the `VerifyDbfAsync` delta logic: not unit-tested, consistent with
  `MainViewModel` having no existing test file (already-accepted reasoning from the crash-prevention
  work — real service dependencies, WPF-bound state, disproportionate scaffolding for what's being
  tested). Verified manually end-to-end instead.

## Out of scope

- Any change to how EMAS's Inventory Control module itself consumes the now-accumulating tables —
  outside this codebase, flagged in §2 as needing real-world confirmation post-ship.
- A UI affordance for resetting/clearing the live tables — not requested; today's manual
  file-rename-outside-the-app approach is accepted as sufficient for now.
- The `No` column (Excel column A) — confirmed unused, not mapped to anything, safe to ignore.
- Any retention/cleanup policy for `DbfSafetyBackupService`'s accumulating backup copies — already
  explicitly out of scope per the crash-prevention design, unchanged here.
