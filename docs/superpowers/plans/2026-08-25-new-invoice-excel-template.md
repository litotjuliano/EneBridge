# New Invoice Excel Template Support Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the Invoice pipeline (`ExcelReaderService.ReadIcmaste` / `ReadIctrane`) correctly read the client's new 9-column, single-header-row Excel template, defaulting the now-absent `FcRate`/`TaxCode` columns to the values every row of the old template always carried, and replace the golden test fixture with a sample of the new template.

**Architecture:** No new files, no new abstractions. Four small, targeted edits to
`src/eneBridge.Wpf.Core/Models/ExcelColumnMap.cs` and
`src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs`, plus a binary fixture swap and matching
test-literal updates in `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs` and
`tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs`.

**Tech Stack:** .NET 8, ClosedXML (Excel reading), xUnit, System.Data.OleDb (DBF round-trip tests).

**Spec:** `docs/superpowers/specs/2026-08-25-new-invoice-excel-template-design.md`

---

### Task 1: Lower `IctraneExcelCol.MinColumnCount` from 12 to 9

The new template has only 9 columns. `IctraneExcelCol.MinColumnCount = 12` currently makes a
9-column file fail validation with a fatal `ExcelValidationException` before any row is even read.
`IcmasteExcelCol.MinColumnCount` is already 8 and needs no change.

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Models/ExcelColumnMap.cs:38`
- Test: `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs:193-206`

- [x] **Step 1: Update the existing test to expect the new minimum**
- [x] **Step 2: Run the test to verify it fails**
- [x] **Step 3: Change the constant**
- [x] **Step 4: Run the test to verify it passes**
- [x] **Step 5: Commit**

Change `IctraneExcelCol.MinColumnCount` from `12` to `9`; update
`ReadIctrane_WorksheetHasTooFewColumns_ThrowsExcelValidationException`'s expected message from
"need at least 12" to "need at least 9". Committed as `1e45c0c`.

---

### Task 2: Default `FcRate` and `TaxCode` in `ReadIcmaste` when blank

Every row of the old template carried literal `FcRate=0.00000000` and `TaxCode=SST0`. The new
template has no columns for either, so today they'd come back as `DBNull` and `""` respectively —
`""` is not a valid EMAS/SST tax code. Default both instead, uniformly whether the column is
missing from the workbook entirely or just empty for a given row.

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs:79-106`
- Test: `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs`

- [x] **Step 1: Write the failing test** (`ReadIcmaste_FcRateAndTaxCodeBlank_DefaultsApplied`)
- [x] **Step 2: Run test to verify it fails**
- [x] **Step 3: Write minimal implementation**
- [x] **Step 4: Run test to verify it passes**
- [x] **Step 5: Run the full icmaste test group to check for regressions**
- [x] **Step 6: Commit**

`decimal? fcRate = null;` → `decimal? fcRate = 0;`; added
`if (IsBlank(taxCode)) { taxCode = "SST0"; }` after reading `taxCode`. Committed as `f5a5a4c`,
with a follow-up cleanup (`decimal? fcRate` → non-nullable `decimal fcRate`, since it can no
longer be null) committed as `8e5c546`.

---

### Task 3: Default `TaxCode` in `ReadIctrane` when blank

Same reasoning as Task 2's `TaxCode` change, applied to the `ictrane` table's own `TaxCode` read.

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs:157`
- Test: `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs`

- [x] **Step 1: Write the failing test** (`ReadIctrane_TaxCodeBlank_DefaultsToSst0`)
- [x] **Step 2: Run test to verify it fails**
- [x] **Step 3: Write minimal implementation**
- [x] **Step 4: Run test to verify it passes**
- [x] **Step 5: Commit**

Committed as `e030300`.

---

### Task 4: Change row-start layout and replace the golden fixture

The new template has a real header on row 1 and data starting row 2 — no extra rows to skip. This
requires changing `ExcelLayout` and, in the same step, replacing the golden fixture (the old
fixture has no header row at all; its row 1 is already real data, and it only worked because
`FirstDataRow=6` happened to skip past it). Because the fixture and the layout constant are
coupled, this task changes both together along with every test literal that depends on the
fixture's exact row/column layout.

Verified ground truth for the new fixture (`25-08-2026-data.xlsx`, 111 total rows: 1 header + 110
data rows):
- icmaste: `RowsRead=110`, `SkipReasons.Count=107`, `RowsWritten=3`, unique refs
  `2501001, 2501002, 2501003`.
- ictrane: `RowsRead=110`, `RowsWritten=110`, zero skips.
- Row 2 (first processed row, table index 0): `Qty=0.629, Price=900, Amount=566.10, Ref=2501001`.
- Row 7 (6th processed row, table index 5): `Qty=0.362, Price=850, Amount=307.70, Ref=2501001`.

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Models/ExcelColumnMap.cs:50-51`
- Replace (binary): `tests/eneBridge.Wpf.Core.Tests/Fixtures/data.xlsx`
- Modify: `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs`
- Modify: `tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs`

- [x] **Step 1: Replace the fixture file**
- [x] **Step 2: Change `ExcelLayout`** (`SkipDataRows` 4→0)
- [x] **Step 3: Update `ReadIcmaste_DedupesByRef_OneRowPerDocument`** (105→110, 102→107)
- [x] **Step 4: Update `ReadIctrane_KeepsEveryLineItem_NoRefDedup`** (105→110)
- [x] **Step 5: Rewrite the two ictrane spot-check tests** as `ReadIctrane_SpotCheck_FirstProcessedRowValues` / `ReadIctrane_SpotCheck_SixthProcessedRowValues`
- [x] **Step 6: Update `ReadIcmaste_SpotCheck_KnownDocumentValues`** (added `FcRate` assertion)
- [x] **Step 7: Update `DbfExportServiceTests.cs`** (105→110)
- [x] **Step 8: Run the full ExcelReaderService test file**
- [x] **Step 9: Run the full test suite**
- [x] **Step 10: Commit**

Committed as `f0dbd35`, with a follow-up doc-comment fix (the `ExcelLayout` XML doc still
referenced "first physical row processed = 6") committed as `acf41c9`.

---

### Task 5: Update CLAUDE.md's stale figures

`CLAUDE.md`'s "Current status" section documents the old fixture's row counts. Update to 110/3/110.

**Files:**
- Modify: `CLAUDE.md`

- [x] **Step 1: Update the row-count figures**
- [x] **Step 2: Commit**

Committed as `b258954` (also corrected a stale test-count figure found along the way: 47→49
tests, verified by actually running `dotnet test`). A further follow-up fixing two more stale
"row 6"/"columns A-L" references discovered during final review was committed as `c1431a9`.

---

## Self-Review Notes

- **Spec coverage:** Row layout change (spec Change 1) → Task 4. Column count validation
  (Change 2) → Task 1. FcRate default (Change 3) → Task 2. TaxCode default (Change 4) → Tasks 2
  & 3. Fixture replacement and dependent test updates → Task 4. Column J requiring no handling is
  inherently satisfied — no task references it.
- **Placeholder scan:** No TBD/TODO.
- **Type consistency:** `IcmasteSchema.FcRate`/`TaxCode`, `IctraneSchema.TaxCode`/`Qty`/`Price`/
  `Amount`/`Ref` all verified against the actual schema files before use. `ExcelLayout.FirstDataRow`
  is computed (`HeaderRows + SkipDataRows + 1`), not hardcoded.

All 5 tasks completed via subagent-driven-development: each implemented, spec-reviewed, and
code-quality-reviewed individually, plus a final whole-implementation review across the full
range (`9bd3147..b258954`) before this branch moved on to the next spec.
