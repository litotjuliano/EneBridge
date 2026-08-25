# Support the New Invoice Excel Template

## Context

The client has switched to a new source Excel template for the Invoice workflow
(`25-08-2026-data.xlsx`, a sample provided for this change). Compared to the template the pipeline
was originally built and tested against (`tests/eneBridge.Wpf.Core.Tests/Fixtures/data.xlsx`, and
the reference copy at `E:\Irene Projects\testdata.xlsx`), the new template differs in two ways:

1. **Column count**: 9 columns (`No, Item Description, Qty, Unit Price, Total Amount, Date,
   Debtor Code, Debtor Name, Invoice Number`) instead of 12. The 3 missing trailing columns are:
   - raw column index 9 — never read by any mapping code in either table (confirmed by grep across
     `ExcelReaderService.cs`); the `"5001/000"` values `icmaste.accno`/`icmaste.postaccno` end up
     with come from hardcoded string literals in `ExcelReaderService.ReadIcmaste`, not from this
     column. No handling needed for it.
   - raw column index 10 — `IcmasteExcelCol.FcRate`.
   - raw column index 11 — `IcmasteExcelCol.TaxCode` / `IctraneExcelCol.TaxCode`.

   Today, `IctraneExcelCol.MinColumnCount = 12` makes a 9-column file fail
   `ExcelReaderService.ValidateColumnCount` with a fatal `ExcelValidationException` for the ictrane
   stage. (`IcmasteExcelCol.MinColumnCount` is already only 8, so icmaste already tolerates a
   9-column file today — reading `FcRate`/`TaxCode` as blank via ClosedXML's out-of-range cell
   handling, which doesn't throw.)

2. **Row layout**: the new template has a real single header row (row 1: column names), with data
   starting immediately at row 2. The current `ExcelLayout` (`HeaderRows=1, SkipDataRows=4`, i.e.
   `FirstDataRow=6`) assumes 4 extra rows to skip past the header before data starts — a hardcoded
   assumption reverse-engineered from the original app's `HDR=YES` OleDb behavior, and confirmed
   against the *old* fixture (whose physical row 1 is actually already real data, not a header —
   `FirstDataRow=6` on that fixture silently skips real records 01–05 by design, which is where the
   documented "105 rows read" figure in `CLAUDE.md` comes from on a 110-row fixture). The new
   template has no such preamble.

The new template replaces the old one going forward — the old 12-column/row-6 layout does not need
to keep working.

## Goal

Make the Invoice pipeline (`ExcelReaderService.ReadIcmaste` / `ReadIctrane`) read the new
9-column, header-row-2 template correctly, with `FcRate` and `TaxCode` defaulted to the same
values every row of the old template always carried, and update the golden test fixture and its
dependent tests to match.

## Design

### Row layout (`ExcelColumnMap.cs`, `ExcelLayout`)

Change:
```csharp
public const int HeaderRows = 1;
public const int SkipDataRows = 4;
```
to:
```csharp
public const int HeaderRows = 1;
public const int SkipDataRows = 0;
```
`FirstDataRow` (`HeaderRows + SkipDataRows + 1`) becomes `2` — data starts right after the single
header row, matching the new template exactly.

### Column count validation (`ExcelColumnMap.cs`, `IctraneExcelCol`)

Lower `IctraneExcelCol.MinColumnCount` from `12` to `9`. This matches the highest column index
either table actually requires to identify a row (`Ref` at index 8, i.e. 9 columns needed). A file
with fewer than 9 columns is genuinely incomplete and should keep failing validation; a 9–11 column
file (the new template) should proceed. `IcmasteExcelCol.MinColumnCount` needs no change (already
8, which was already below the columns it reads).

### FcRate default (`ExcelReaderService.ReadIcmaste`)

Currently, a blank `FcRate` cell (column missing entirely, or present but empty) results in
`DBNull`. Change the default to `0`:
```csharp
decimal? fcRate = 0;
if (!IsBlank(fcRateRaw))
{
    if (!decimal.TryParse(fcRateRaw, out var parsedFcRate)) { ... }
    fcRate = parsedFcRate;
}
...
row[IcmasteSchema.FcRate] = fcRate;
```
This matches the literal value (`0.00000000`) every row carries in the old template, and applies
uniformly whether the column is absent from the workbook or just empty for a given row — no need
to distinguish those two cases.

### TaxCode default (`ExcelReaderService.ReadIcmaste` and `ReadIctrane`)

Currently, a blank `TaxCode` cell results in an empty string. Change the default to `"SST0"` in
both methods:
```csharp
string taxCode = GetString(worksheet, excelRow, IcmasteExcelCol.TaxCode); // or IctraneExcelCol
if (IsBlank(taxCode)) { taxCode = "SST0"; }
```
Same reasoning as `FcRate`: every row in the old template carries this exact value, and a blank tax
code is not valid for the downstream EMAS/SST import. Applies uniformly to "column absent" and
"column present but empty."

Both defaults are applied silently — no UI or run-history indication that a value was defaulted
rather than read, matching how blank `FcRate` is already handled today with no visibility.

### Golden fixture replacement

Replace `tests/eneBridge.Wpf.Core.Tests/Fixtures/data.xlsx` with a copy of
`E:\Irene Projects\data\25-08-2026-data.xlsx` as-is (header row + 9 columns, no manual column
padding — the reader now handles it natively via the changes above).

The new fixture has 111 total rows (1 header + 110 data rows) and the same 3 unique invoice
numbers as the old fixture (`2501001`, `2501002`, `2501003`) — this appears to be the same
underlying dataset re-exported through the new template, not new/different data.

Update `ExcelReaderServiceTests.cs` tests that hardcode counts against the fixture:

- `ReadIcmaste_DedupesByRef_OneRowPerDocument`: `RowsRead` 105→110, `SkipReasons.Count` 102→107;
  `RowsWritten` stays 3 and the 3 expected refs (`2501001`, `2501002`, `2501003`) stay the same.
- `ReadIctrane_KeepsEveryLineItem_NoRefDedup`: `RowsRead`/`RowsWritten` 105→110.
- `ReadIctrane_SpotCheck_RowSixValues` (and any other test asserting literal values at a specific
  row number): re-point at whichever row now holds that same record under the row-2 start.
- Add assertions (new or extended) confirming a row with blank `FcRate`/`TaxCode` cells in the
  fixture actually resolves to `0` / `"SST0"` in the resulting table, so the default path itself is
  covered, not just inherited from the old passing values.

Tests that already reference `ExcelLayout.FirstDataRow` symbolically (rather than a literal row
number) need no changes — they adapt automatically to the new value.

`DbfExportServiceTests.cs` also uses this fixture and hardcodes the same figure twice
(`Assert.Equal(105, ictraneExport.RowsWritten)` and `AssertRowCount("ictrane", 105)`) — both need
updating to 110 alongside the `ExcelReaderServiceTests.cs` changes.

### Out of scope

- Raw column index 9 (unused by any mapping code) — no change.
- `DbfExportService`, `DbfSchemaBuilder`, `SqlValueFormatter`, and the WPF layer — untouched; this
  is purely a change to how `ExcelReaderService` interprets the source workbook.
- Supporting both the old and new template simultaneously — the new template replaces the old one,
  so no format-detection logic is introduced.
