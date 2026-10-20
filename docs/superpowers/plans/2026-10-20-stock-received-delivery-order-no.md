# Stock Received Delivery Order No -> HFIELD1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Read the optional "Delivery Order No" column (Excel column L) of the Stock Received template and write it to `HFIELD1` in `icmaste.dbf`.

**Architecture:** Add one column constant to `StockReceivedExcelCol`, one name constant to `IcmasteSchema`, read the cell in `ReadStockReceivedRowFields`, and assign it in `ReadStockReceivedIcmaste`. The column is optional (`MinColumnCount` stays 11), so a missing column or blank cell yields an empty `HFIELD1` and never skips a row. The preview grid's "populated columns" display filter gains `HFIELD1`. Spec: `docs/superpowers/specs/2026-10-20-stock-received-delivery-order-no-design.md`.

**Tech Stack:** C# / .NET 8 (`net8.0-windows`), ClosedXML, xUnit, OleDb (ACE) for the DBF round trip.

---

## File Structure

- Modify `src/eneBridge.Wpf.Core/Models/StockReceivedExcelCol.cs`: add `DeliveryOrderNo`.
- Modify `src/eneBridge.Wpf.Core/Models/IcmasteSchema.cs`: add `HField1` constant.
- Modify `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs`: read and map the value.
- Modify `src/eneBridge.Wpf/MainWindow.xaml.cs`: add `"HFIELD1"` to `StockReceivedIcmastePopulatedColumns`.
- Modify `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs`: reader tests.
- Modify `tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs`: round-trip test.
- Modify `CLAUDE.md`: short note.

Run commands from `E:\Irene Projects\eneBridge 2.0`. Note: an unfiltered `dotnet test` can abort on the known ACE driver instability (see CLAUDE.md). Use `--filter` as shown.

---

### Task 1: Reader mapping (TDD)

**Files:**
- Modify: `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs` (helpers near line 413-438, new tests after line 508)
- Modify: `src/eneBridge.Wpf.Core/Models/StockReceivedExcelCol.cs`
- Modify: `src/eneBridge.Wpf.Core/Models/IcmasteSchema.cs`
- Modify: `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs`

- [ ] **Step 1: Extend the test helpers**

In `SetStockReceivedRow`, add a trailing optional parameter and set the cell. Replace the method (lines 413-428) with:

```csharp
    private static void SetStockReceivedRow(
        IXLWorksheet worksheet, int row, string? refValue, string? dateValue, string? codeValue, string? nameValue,
        string? itemNoValue, string? desc1Value = "ITEM DESC", string? qtyValue = "1", string? nAmtValue = "100",
        string? tAmtValue = "100", string? desc2Value = null, string? deliveryOrderNoValue = null)
    {
        SetCell(worksheet, row, StockReceivedExcelCol.Ref, refValue);
        SetCell(worksheet, row, StockReceivedExcelCol.Date, dateValue);
        SetCell(worksheet, row, StockReceivedExcelCol.Code, codeValue);
        SetCell(worksheet, row, StockReceivedExcelCol.Name, nameValue);
        SetCell(worksheet, row, StockReceivedExcelCol.ItemNo, itemNoValue);
        SetCell(worksheet, row, StockReceivedExcelCol.Desc1, desc1Value);
        SetCell(worksheet, row, StockReceivedExcelCol.Desc2, desc2Value);
        SetCell(worksheet, row, StockReceivedExcelCol.Qty, qtyValue);
        SetCell(worksheet, row, StockReceivedExcelCol.NAmt, nAmtValue);
        SetCell(worksheet, row, StockReceivedExcelCol.TAmt, tAmtValue);
        SetCell(worksheet, row, StockReceivedExcelCol.DeliveryOrderNo, deliveryOrderNoValue);
    }
```

Replace `BuildStockReceivedWorksheet` (lines 430-438) with:

```csharp
    private static IXLWorksheet BuildStockReceivedWorksheet(
        string? refValue = "SB-000001", string? dateValue = "2026-09-01", string? codeValue = "4000T001",
        string? nameValue = "TAN KANG KAE", string? itemNoValue = "sub-con", string? desc1Value = "Sub Con Wages",
        string? qtyValue = "1", string? nAmtValue = "1000", string? tAmtValue = "1000", string? desc2Value = null,
        string? deliveryOrderNoValue = null)
    {
        var worksheet = NewWorksheet();
        SetStockReceivedRow(worksheet, ExcelLayout.FirstDataRow, refValue, dateValue, codeValue, nameValue, itemNoValue, desc1Value, qtyValue, nAmtValue, tAmtValue, desc2Value, deliveryOrderNoValue);
        return worksheet;
    }
```

- [ ] **Step 2: Add the failing tests**

Insert after `ReadStockReceivedIctrane_Desc2Blank_MapsToEmptyString` (after line 508):

```csharp
    [Fact]
    public void ReadStockReceivedIcmaste_DeliveryOrderNoPopulated_MapsToHField1()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(deliveryOrderNoValue: "DO1234");

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal("DO1234", result.Table.Rows[0][IcmasteSchema.HField1]);
    }

    [Fact]
    public void ReadStockReceivedIcmaste_DeliveryOrderNoBlank_MapsToEmptyStringAndRowIsKept()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet();

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Empty(result.SkipReasons);
        Assert.Equal(string.Empty, result.Table.Rows[0][IcmasteSchema.HField1]);
    }

    [Fact]
    public void ReadStockReceivedIcmaste_DeliveryOrderNoLongerThan40_IsTruncated()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(deliveryOrderNoValue: new string('D', 50));

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal(new string('D', 40), result.Table.Rows[0][IcmasteSchema.HField1]);
    }

    [Fact]
    public void ReadStockReceivedIcmaste_ElevenColumnWorkbookWithoutDeliveryOrderColumn_StillReads()
    {
        // BuildStockReceivedWorksheet with no DO value never touches column L, so the worksheet's
        // used range ends at column K (11 columns) -- the pre-existing template shape.
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet();

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Empty(result.SkipReasons);
        Assert.Equal(string.Empty, result.Table.Rows[0][IcmasteSchema.HField1]);
    }

    [Fact]
    public void ReadStockReceivedIctrane_DeliveryOrderNoPopulated_DoesNotAffectIctrane()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(deliveryOrderNoValue: "DO1234");

        var result = service.ReadStockReceivedIctrane(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal("sub-con", result.Table.Rows[0][IctraneSchema.ItemNo]);
        Assert.Equal(1000m, result.Table.Rows[0][IctraneSchema.Amount]);
    }
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~DeliveryOrderNo"`
Expected: build FAIL (`StockReceivedExcelCol.DeliveryOrderNo` and `IcmasteSchema.HField1` do not exist).

- [ ] **Step 4: Add the constants**

In `src/eneBridge.Wpf.Core/Models/StockReceivedExcelCol.cs`, after `public const int TAmt = 10;` and before the `MinColumnCount` summary, insert:

```csharp

    /// <summary>
    /// "Delivery Order No" -- added to the template after the 11-column layout (client-provided
    /// sample-19-10-2026.xlsx, 2026-10-19), as column L. Maps to <see cref="Models.IcmasteSchema.HField1"/>.
    /// Deliberately OPTIONAL: <see cref="MinColumnCount"/> is not raised, so older 11-column
    /// workbooks still validate, and a missing column or blank cell just leaves HFIELD1 blank
    /// (never a row-skip condition).
    /// </summary>
    public const int DeliveryOrderNo = 11;
```

Also update the `MinColumnCount` summary to: `/// <summary>Minimum raw column count required (through Total Amount at index 10). Delivery Order No (index 11) is optional.</summary>`

In `src/eneBridge.Wpf.Core/Models/IcmasteSchema.cs`, after `public const string TAmt = "T_AMT";` add:

```csharp
    public const string HField1 = "HFIELD1";
```

- [ ] **Step 5: Implement the reader change**

In `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs`:

1. Change the `StockReceivedRowFields` record (line 191-192) to:

```csharp
    private readonly record struct StockReceivedRowFields(
        string Ref, DateTime Date, string Code, string Name, string ItemNo, string Desc1, string Desc2, string DeliveryOrderNo, decimal Qty, decimal NAmt, decimal TAmt);
```

2. In `ReadStockReceivedRowFields`, after the `desc2Value` line (225) add:

```csharp
        // Optional column: GetString returns empty for an empty cell, and ClosedXML returns an empty
        // cell for any position beyond the used range, so an 11-column workbook is safe here.
        string deliveryOrderNoValue = GetString(worksheet, excelRow, StockReceivedExcelCol.DeliveryOrderNo);
```

3. Change the return (line 236) to:

```csharp
        return new StockReceivedRowFields(refValue, date, codeValue, nameValue, itemNoValue, desc1Value, desc2Value, deliveryOrderNoValue, qtyOpt ?? 0m, nAmtOpt ?? 0m, tAmtOpt ?? 0m);
```

4. In `ReadStockReceivedIcmaste`, after the `row[IcmasteSchema.TAmt] = fields.Value.TAmt;` line add:

```csharp
            row[IcmasteSchema.HField1] = TextTruncation.Truncate(fields.Value.DeliveryOrderNo, 40);
```

- [ ] **Step 6: Run tests to verify they pass**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~ReadStockReceived"`
Expected: PASS (all existing Stock Received reader tests plus the 5 new ones).

- [ ] **Step 7: Commit**

```bash
git add src/eneBridge.Wpf.Core/Models/StockReceivedExcelCol.cs src/eneBridge.Wpf.Core/Models/IcmasteSchema.cs src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs
git commit -m "feat: map Stock Received Delivery Order No to icmaste HFIELD1

Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Preview grid shows HFIELD1

**Files:**
- Modify: `src/eneBridge.Wpf/MainWindow.xaml.cs:40-44`

The preview grids auto-generate every column and hide those not in a "populated" set unless "Show all columns" is ticked. Without this change the new value would be hidden by default.

- [ ] **Step 1: Add the column to the Stock Received icmaste display set**

Replace lines 40-44:

```csharp
    private static readonly HashSet<string> StockReceivedIcmastePopulatedColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "TYPE", "ENTRY", "REF", "DATE", "CODE", "NAME", "ACCNO", "USER", "POSTACCNO",
        "CUST2", "CURRCODE", "TAXCODE", "VENDNO", "N_AMT", "T_AMT", "HFIELD1"
    };
```

Do not change `IcmastePopulatedColumns` (Invoice never populates HFIELD1).

- [ ] **Step 2: Build**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: Build succeeded, 0 errors.

- [ ] **Step 3: Commit**

```bash
git add src/eneBridge.Wpf/MainWindow.xaml.cs
git commit -m "feat: show HFIELD1 in the Stock Received icmaste preview by default

Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 3: DBF round-trip test

**Files:**
- Modify: `tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs` (add after `Export_IcmasteAndIctrane_ProducesValidDbfFiles`, before the next `[Fact]` at line 70)

- [ ] **Step 1: Add the test**

```csharp
    [Fact]
    public void Export_StockReceivedIcmaste_DeliveryOrderNoSurvivesRoundTripInHField1()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();

        using var workbook = new ClosedXML.Excel.XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");
        int row = ExcelLayout.FirstDataRow;
        worksheet.Cell(row, StockReceivedExcelCol.Ref + 1).SetValue("SB-000001");
        worksheet.Cell(row, StockReceivedExcelCol.Date + 1).SetValue("2026-09-01");
        worksheet.Cell(row, StockReceivedExcelCol.Code + 1).SetValue("4000T001");
        worksheet.Cell(row, StockReceivedExcelCol.Name + 1).SetValue("TAN KANG KAE");
        worksheet.Cell(row, StockReceivedExcelCol.ItemNo + 1).SetValue("sub-con");
        worksheet.Cell(row, StockReceivedExcelCol.Qty + 1).SetValue("1");
        worksheet.Cell(row, StockReceivedExcelCol.NAmt + 1).SetValue("1000");
        worksheet.Cell(row, StockReceivedExcelCol.TAmt + 1).SetValue("1000");
        worksheet.Cell(row, StockReceivedExcelCol.DeliveryOrderNo + 1).SetValue("DO1234");

        var read = excelReader.ReadStockReceivedIcmaste(worksheet);
        var export = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, read.Table, IcmasteSchema.Columns);

        Assert.True(export.Success, export.ErrorMessage);
        Assert.Empty(export.RowErrors);
        Assert.Equal(1, export.RowsWritten);

        var dbf = new DbfReaderService().Read(_scratchFolder, IcmasteSchema.TableName);
        Assert.True(dbf.Success, dbf.ErrorMessage);
        Assert.Single(dbf.Table.Rows);
        Assert.Equal("DO1234", dbf.Table.Rows[0][IcmasteSchema.HField1].ToString()?.Trim());
    }
```

If `ExcelLayout` or `ClosedXML` are not already imported in this file, add `using eneBridge.Wpf.Core.Models;` (already present) and use the fully qualified `ClosedXML.Excel.XLWorkbook` as written. `ExcelLayout` lives in `eneBridge.Wpf.Core.Models` (ExcelColumnMap.cs).

- [ ] **Step 2: Run the test**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~Export_StockReceivedIcmaste_DeliveryOrderNoSurvivesRoundTripInHField1"`
Expected: PASS. If the process aborts with `AccessViolationException` in `ComObject.Finalize`, that is the known ACE instability (CLAUDE.md "Known reliability risk"), not a failure of this change: re-run once. Requires the Access Database Engine installed.

- [ ] **Step 3: Commit**

```bash
git add tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs
git commit -m "test: DBF round trip for Stock Received HFIELD1

Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Documentation and manual verification

**Files:**
- Modify: `CLAUDE.md` (add a paragraph after the "Stock Received template gained an 'Item Description 2' column" paragraph)

- [ ] **Step 1: Add the CLAUDE.md note**

Append this paragraph after the Item Description 2 paragraph:

```markdown
**Stock Received template gained an optional "Delivery Order No" column.** A client-provided
`sample-19-10-2026.xlsx` (2026-10-19) added a 12th column L, "Delivery Order No".
`StockReceivedExcelCol.DeliveryOrderNo` (index 11) maps it to `IcmasteSchema.HField1`
(`HFIELD1`, varchar 40, already part of the 151-column schema) in `icmaste.dbf` only; `ictrane` is
unchanged and Invoice is untouched. Unlike Item Description 2, it is **optional**:
`MinColumnCount` stays at 11, so older 11-column workbooks still validate, and a missing column or
blank cell leaves `HFIELD1` blank without skipping the row. Values are truncated to 40 characters.
`MainWindow.xaml.cs`'s `StockReceivedIcmastePopulatedColumns` includes `HFIELD1` so the preview grid
shows it without "Show all columns".
```

- [ ] **Step 2: Run the Stock Received and export test groups**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~ExcelReaderServiceTests"`
Expected: PASS, no failures.

- [ ] **Step 3: Manual check against the real sample**

Run the app (`dotnet run --project "src/eneBridge.Wpf/eneBridge.Wpf.csproj"`), go to the Stock Received tab, Browse to `E:\Irene Projects\sample-19-10-2026.xlsx`, click Preview. Expected: the icmaste preview shows an `HFIELD1` column with `DO1234` (SB-000001) and `DO3456` (SB-000002). Do NOT click Confirm & Export against the live EMAS folder unless you intend to overwrite the staging files.

- [ ] **Step 4: Commit**

```bash
git add CLAUDE.md
git commit -m "docs: note optional Delivery Order No column in CLAUDE.md

Co-Authored-By: Claude Sonnet 5.5 <noreply@anthropic.com>"
```

---

## Self-Review

- **Spec coverage:** constant on `StockReceivedExcelCol` (Task 1), `HField1` constant (Task 1), reader mapping and optional/blank/missing handling and 40-char truncation (Task 1), preview grid (Task 2), round-trip test (Task 3), CLAUDE.md note (Task 4). Invoice and `ictrane` untouched (asserted for ictrane in Task 1's last test).
- **Placeholders:** none.
- **Type consistency:** `DeliveryOrderNo` (const int and record field and test parameter `deliveryOrderNoValue`), `IcmasteSchema.HField1` are used identically in every task.
