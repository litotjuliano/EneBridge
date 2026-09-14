# Stock Received Workflow Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ship a working Stock Received tab that reads `Self-Billed_Format-ORI.xlsx`-shaped workbooks and writes into the same `icmaste.dbf`/`ictrane.dbf` Invoice writes to (distinguished by `TYPE`), on a shared append-based export path both workflows now use.

**Architecture:** New Core-layer read methods (`ExcelReaderService`) and a new Excel column map feed the existing `DbfExportService`/`DbfReaderService`, both rewritten to append instead of recreate and to support TYPE-filtered reads. A new shared `DbfWorkflowHelper` (Wpf project) holds the ACE-driver-crash-avoidance table check and the new before/after-delta verification, used by both a rewired `MainViewModel` (Invoice) and a new parallel `StockReceivedViewModel`, each bound to its own `TabItem` in `MainWindow.xaml`.

**Tech Stack:** .NET 8, WPF, ClosedXML, `System.Data.OleDb` (ACE provider), CommunityToolkit.Mvvm, xUnit.

---

### Task 1: Excel column map + schema constants for the fields Stock Received newly touches

Per the project's convention ("column names actually touched by mapping code are exposed as
`const string` members"), `VENDNO`/`N_AMT`/`T_AMT` (icmaste) and `desc2` (ictrane) are about to be
written from mapping code for the first time — they need named constants, matching how `ACCNO`/
`POSTACCNO`/etc. already work. Pure rename-to-constant refactor of existing literals; the resulting
DBF column names are unchanged, so no existing test is affected.

**Files:**
- Create: `src/eneBridge.Wpf.Core/Models/StockReceivedExcelCol.cs`
- Modify: `src/eneBridge.Wpf.Core/Models/IcmasteSchema.cs`
- Modify: `src/eneBridge.Wpf.Core/Models/IctraneSchema.cs`

- [ ] **Step 1: Create the new column map**

```csharp
namespace eneBridge.Wpf.Core.Models;

/// <summary>
/// 0-based raw Excel column indices for the Stock Received workflow's source format
/// (Self-Billed_Format-ORI.xlsx), reverse-engineered from Reference.xlsx's explicit column
/// mapping. Unlike Invoice, every column here feeds either icmaste or ictrane from the exact
/// same physical row -- there is no separate icmaste/ictrane column set. See
/// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md.
/// </summary>
public static class StockReceivedExcelCol
{
    // Column A ("No") is a sequential row number in the source file -- not mapped to any EMAS field.
    public const int Ref = 1;
    public const int Date = 2;
    public const int Code = 3;
    public const int Name = 4;
    public const int ItemNo = 5;
    public const int Desc1 = 6;
    public const int Qty = 7;
    public const int NAmt = 8;
    public const int TAmt = 9;

    /// <summary>Minimum raw column count required (through Total Amount at index 9).</summary>
    public const int MinColumnCount = 10;
}
```

- [ ] **Step 2: Add the three new constants to `IcmasteSchema` and use them in `BuildColumns`**

In `src/eneBridge.Wpf.Core/Models/IcmasteSchema.cs`, add three new constants next to the existing
ones (after `public const string Erefund = "EREFUND";`):

```csharp
    public const string Erefund = "EREFUND";
    public const string VendNo = "VENDNO";
    public const string NAmt = "N_AMT";
    public const string TAmt = "T_AMT";
```

Then replace the two literal-string column definitions:

```csharp
            new("T_AMT", m),
            new("N_AMT", m),
```

with:

```csharp
            new(TAmt, m),
            new(NAmt, m),
```

and replace:

```csharp
            new DbfColumnDefinition("VENDNO", s, 8),
```

with:

```csharp
            new DbfColumnDefinition(VendNo, s, 8),
```

- [ ] **Step 3: Add the `Desc2` constant to `IctraneSchema` and use it in `BuildColumns`**

In `src/eneBridge.Wpf.Core/Models/IctraneSchema.cs`, add one new constant next to `Desc1`:

```csharp
    public const string Desc1 = "desc1";
    public const string Desc2 = "desc2";
```

Then replace:

```csharp
            new("desc2", s, 40),
```

with:

```csharp
            new(Desc2, s, 40),
```

- [ ] **Step 4: Build and run the existing schema tests to confirm nothing changed behaviorally**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~IcmasteSchema_HasExpectedColumnCount|FullyQualifiedName~IctraneSchema_HasExpectedColumnCount"`
Expected: `Passed! - Failed: 0, Passed: 2, Skipped: 0, Total: 2` (151/80 columns respectively — unchanged by a pure rename)

- [ ] **Step 5: Commit**

```bash
git add src/eneBridge.Wpf.Core/Models/StockReceivedExcelCol.cs src/eneBridge.Wpf.Core/Models/IcmasteSchema.cs src/eneBridge.Wpf.Core/Models/IctraneSchema.cs
git commit -m "Add Stock Received Excel column map and schema constants it needs

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 2: `ExcelReaderService` — read Stock Received's icmaste/ictrane rows (TDD)

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs`
- Modify: `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs`

- [ ] **Step 1: Write the failing tests**

Add these helper methods and test cases to `tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs`,
right after the existing `BuildIctraneWorksheet` method (before the closing brace of the class):

```csharp
    private static void SetStockReceivedRow(
        IXLWorksheet worksheet, int row, string? refValue, string? dateValue, string? codeValue, string? nameValue,
        string? itemNoValue, string? desc1Value = "ITEM DESC", string? qtyValue = "1", string? nAmtValue = "100",
        string? tAmtValue = "100")
    {
        SetCell(worksheet, row, StockReceivedExcelCol.Ref, refValue);
        SetCell(worksheet, row, StockReceivedExcelCol.Date, dateValue);
        SetCell(worksheet, row, StockReceivedExcelCol.Code, codeValue);
        SetCell(worksheet, row, StockReceivedExcelCol.Name, nameValue);
        SetCell(worksheet, row, StockReceivedExcelCol.ItemNo, itemNoValue);
        SetCell(worksheet, row, StockReceivedExcelCol.Desc1, desc1Value);
        SetCell(worksheet, row, StockReceivedExcelCol.Qty, qtyValue);
        SetCell(worksheet, row, StockReceivedExcelCol.NAmt, nAmtValue);
        SetCell(worksheet, row, StockReceivedExcelCol.TAmt, tAmtValue);
    }

    private static IXLWorksheet BuildStockReceivedWorksheet(
        string? refValue = "SB-000001", string? dateValue = "2026-09-01", string? codeValue = "4000T001",
        string? nameValue = "TAN KANG KAE", string? itemNoValue = "sub-con", string? desc1Value = "Sub Con Wages",
        string? qtyValue = "1", string? nAmtValue = "1000", string? tAmtValue = "1000")
    {
        var worksheet = NewWorksheet();
        SetStockReceivedRow(worksheet, ExcelLayout.FirstDataRow, refValue, dateValue, codeValue, nameValue, itemNoValue, desc1Value, qtyValue, nAmtValue, tAmtValue);
        return worksheet;
    }

    [Fact]
    public void ReadStockReceivedIcmaste_ValidRow_MapsFieldsCorrectly()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet();

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.Table.Rows);
        var row = result.Table.Rows[0];
        Assert.Equal("RE", row[IcmasteSchema.Type]);
        Assert.Equal("SB-000001", row[IcmasteSchema.Ref]);
        Assert.Equal("4000T001", row[IcmasteSchema.Code]);
        Assert.Equal("TAN KANG KAE", row[IcmasteSchema.Name]);
        Assert.Equal("3020/000", row[IcmasteSchema.Accno]);
        Assert.Equal("6010/000", row[IcmasteSchema.PostAccno]);
        Assert.Equal("SST0", row[IcmasteSchema.TaxCode]);
        Assert.Equal("4000T001", row[IcmasteSchema.VendNo]);
        Assert.Equal("4000T001", row[IcmasteSchema.Cust2]);
        Assert.Equal(1000m, row[IcmasteSchema.NAmt]);
        Assert.Equal(1000m, row[IcmasteSchema.TAmt]);
        Assert.Equal("MYR", row[IcmasteSchema.CurrCode]);
        Assert.Equal("admin", row[IcmasteSchema.User]);
    }

    [Fact]
    public void ReadStockReceivedIctrane_ValidRow_MapsFieldsCorrectly()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet();

        var result = service.ReadStockReceivedIctrane(worksheet);

        Assert.Single(result.Table.Rows);
        var row = result.Table.Rows[0];
        Assert.Equal("RE", row[IctraneSchema.Type]);
        Assert.Equal("SB-000001", row[IctraneSchema.Ref]);
        Assert.Equal("sub-con", row[IctraneSchema.ItemNo]);
        Assert.Equal("Sub Con Wages", row[IctraneSchema.Desc1]);
        Assert.Equal("1", row[IctraneSchema.Desc2]);
        Assert.Equal("SST0", row[IctraneSchema.TaxCode]);
        Assert.Equal("admin", row[IctraneSchema.UserId]);
        Assert.Equal(DBNull.Value, row[IctraneSchema.Price]);
        Assert.Equal(DBNull.Value, row[IctraneSchema.Amount]);
        Assert.Equal(DBNull.Value, row[IctraneSchema.Qty]);
    }

    [Fact]
    public void ReadStockReceivedIcmasteAndIctrane_SameRefAcrossMultipleRows_NoDedup()
    {
        var service = new ExcelReaderService();
        var worksheet = NewWorksheet();
        SetStockReceivedRow(worksheet, ExcelLayout.FirstDataRow, "SB-000001", "2026-09-01", "4000T001", "TAN KANG KAE", "sub-con");
        SetStockReceivedRow(worksheet, ExcelLayout.FirstDataRow + 1, "SB-000001", "2026-09-02", "4000T002", "TAN KANG SHYONG", "item-2");

        var icmasteResult = service.ReadStockReceivedIcmaste(worksheet);
        var ictraneResult = service.ReadStockReceivedIctrane(worksheet);

        // Every row is its own standalone document -- unlike Invoice, icmaste does NOT dedupe by REF.
        Assert.Equal(2, icmasteResult.Table.Rows.Count);
        Assert.Equal(2, ictraneResult.Table.Rows.Count);
    }

    [Theory]
    [InlineData("ref", "REF is blank")]
    [InlineData("date", "DATE is blank")]
    [InlineData("code", "CODE is blank")]
    [InlineData("name", "NAME is blank")]
    [InlineData("itemNo", "ITEM_NO is blank")]
    public void ReadStockReceivedIcmaste_BlankRequiredField_SkipsRowWithReason(string blankField, string expectedReason)
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(
            refValue: blankField == "ref" ? null : "SB-000001",
            dateValue: blankField == "date" ? null : "2026-09-01",
            codeValue: blankField == "code" ? null : "4000T001",
            nameValue: blankField == "name" ? null : "TAN KANG KAE",
            itemNoValue: blankField == "itemNo" ? null : "sub-con");

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal(expectedReason, result.SkipReasons[0].Reason);
        Assert.Empty(result.Table.Rows);
    }

    [Theory]
    [InlineData("ref", "REF is blank")]
    [InlineData("date", "DATE is blank")]
    [InlineData("code", "CODE is blank")]
    [InlineData("name", "NAME is blank")]
    [InlineData("itemNo", "ITEM_NO is blank")]
    public void ReadStockReceivedIctrane_BlankRequiredField_SkipsRowWithReason(string blankField, string expectedReason)
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(
            refValue: blankField == "ref" ? null : "SB-000001",
            dateValue: blankField == "date" ? null : "2026-09-01",
            codeValue: blankField == "code" ? null : "4000T001",
            nameValue: blankField == "name" ? null : "TAN KANG KAE",
            itemNoValue: blankField == "itemNo" ? null : "sub-con");

        var result = service.ReadStockReceivedIctrane(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal(expectedReason, result.SkipReasons[0].Reason);
        Assert.Empty(result.Table.Rows);
    }

    [Fact]
    public void ReadStockReceivedIcmaste_NAmtNotNumeric_SkipsRowWithReason()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(nAmtValue: "abc");

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal("Unit Price value 'abc' is not a valid number", result.SkipReasons[0].Reason);
    }

    [Fact]
    public void ReadStockReceivedIcmaste_TAmtNotNumeric_SkipsRowWithReason()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(tAmtValue: "abc");

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal("Total Amount value 'abc' is not a valid number", result.SkipReasons[0].Reason);
    }

    [Fact]
    public void ReadStockReceivedIcmaste_NAmtAndTAmtBlank_DefaultToZero()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(nAmtValue: null, tAmtValue: null);

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal(0m, result.Table.Rows[0][IcmasteSchema.NAmt]);
        Assert.Equal(0m, result.Table.Rows[0][IcmasteSchema.TAmt]);
    }

    [Fact]
    public void ReadStockReceivedIcmaste_WorksheetHasTooFewColumns_ThrowsExcelValidationException()
    {
        var service = new ExcelReaderService();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");
        worksheet.Cell(1, 5).SetValue("x");

        var ex = Assert.Throws<ExcelValidationException>(() => service.ReadStockReceivedIcmaste(worksheet));

        Assert.Equal(
            "The Excel file does not contain the required columns (found 5, need at least 10).",
            ex.Message);
    }

    [Fact]
    public void ReadStockReceivedIctrane_WorksheetHasTooFewColumns_ThrowsExcelValidationException()
    {
        var service = new ExcelReaderService();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");
        worksheet.Cell(1, 5).SetValue("x");

        var ex = Assert.Throws<ExcelValidationException>(() => service.ReadStockReceivedIctrane(worksheet));

        Assert.Equal(
            "The Excel file does not contain the required columns (found 5, need at least 10).",
            ex.Message);
    }
```

- [ ] **Step 2: Run tests to verify they fail (compile error)**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~ReadStockReceived"`
Expected: build FAILS — `ReadStockReceivedIcmaste`/`ReadStockReceivedIctrane` do not exist.

- [ ] **Step 3: Implement the shared row-field reader and the two new public methods**

Add this private nested record and helper method to `src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs`,
right before the existing `private static bool TryParseOptionalDecimal(...)` method:

```csharp
    private readonly record struct StockReceivedRowFields(
        string Ref, DateTime Date, string Code, string Name, string ItemNo, string Desc1, string Qty, decimal NAmt, decimal TAmt);

    /// <summary>
    /// Shared validation for both ReadStockReceivedIcmaste and ReadStockReceivedIctrane -- every
    /// valid row produces exactly one row in each table with identical skip conditions (see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md).
    /// </summary>
    private static StockReceivedRowFields? ReadStockReceivedRowFields(IXLWorksheet worksheet, int excelRow, List<RowSkipReason> skipReasons)
    {
        string refValue = GetString(worksheet, excelRow, StockReceivedExcelCol.Ref);
        string dateRaw = GetString(worksheet, excelRow, StockReceivedExcelCol.Date);
        string codeValue = GetString(worksheet, excelRow, StockReceivedExcelCol.Code);
        string nameValue = GetString(worksheet, excelRow, StockReceivedExcelCol.Name);
        string itemNoValue = GetString(worksheet, excelRow, StockReceivedExcelCol.ItemNo);

        if (IsBlank(refValue)) { skipReasons.Add(new RowSkipReason(excelRow, "REF is blank")); return null; }
        if (IsBlank(dateRaw)) { skipReasons.Add(new RowSkipReason(excelRow, "DATE is blank")); return null; }
        if (IsBlank(codeValue)) { skipReasons.Add(new RowSkipReason(excelRow, "CODE is blank")); return null; }
        if (IsBlank(nameValue)) { skipReasons.Add(new RowSkipReason(excelRow, "NAME is blank")); return null; }
        if (IsBlank(itemNoValue)) { skipReasons.Add(new RowSkipReason(excelRow, "ITEM_NO is blank")); return null; }

        DateTime date;
        try
        {
            date = ParseDate(worksheet, excelRow, StockReceivedExcelCol.Date, dateRaw);
        }
        catch (Exception ex)
        {
            skipReasons.Add(new RowSkipReason(excelRow, $"DATE parse failed: {ex.Message}"));
            return null;
        }

        string desc1Value = GetString(worksheet, excelRow, StockReceivedExcelCol.Desc1);
        string qtyValue = GetString(worksheet, excelRow, StockReceivedExcelCol.Qty);

        string nAmtRaw = GetString(worksheet, excelRow, StockReceivedExcelCol.NAmt);
        if (!TryParseOptionalDecimal(nAmtRaw, "Unit Price", excelRow, skipReasons, out var nAmtOpt)) { return null; }

        string tAmtRaw = GetString(worksheet, excelRow, StockReceivedExcelCol.TAmt);
        if (!TryParseOptionalDecimal(tAmtRaw, "Total Amount", excelRow, skipReasons, out var tAmtOpt)) { return null; }

        return new StockReceivedRowFields(refValue, date, codeValue, nameValue, itemNoValue, desc1Value, qtyValue, nAmtOpt ?? 0m, tAmtOpt ?? 0m);
    }
```

Then add these two public methods right after the existing `ReadIctrane` method (before `TryParseOptionalDecimal`):

```csharp
    /// <summary>
    /// Reads icmaste rows for the Stock Received workflow. Unlike Invoice, every valid row
    /// produces its own icmaste row -- no REF dedup, since each row is a standalone self-billed
    /// document (confirmed against the real sample data; see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md).
    /// </summary>
    public StageReadResult ReadStockReceivedIcmaste(IXLWorksheet worksheet)
    {
        ValidateColumnCount(worksheet, StockReceivedExcelCol.MinColumnCount);

        var table = IcmasteSchema.BuildEmptyTable();
        var skipReasons = new List<RowSkipReason>();

        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int rowsRead = 0;

        for (int excelRow = ExcelLayout.FirstDataRow; excelRow <= lastRow; excelRow++)
        {
            rowsRead++;

            var fields = ReadStockReceivedRowFields(worksheet, excelRow, skipReasons);
            if (fields is null)
            {
                continue;
            }

            var row = table.NewRow();
            row[IcmasteSchema.Type] = "RE";
            row[IcmasteSchema.Ref] = TextTruncation.Truncate(fields.Value.Ref, 11);
            row[IcmasteSchema.Date] = fields.Value.Date.Date;
            row[IcmasteSchema.Code] = TextTruncation.Truncate(fields.Value.Code, 8);
            row[IcmasteSchema.Name] = TextTruncation.Truncate(fields.Value.Name, 40);
            row[IcmasteSchema.Accno] = "3020/000";
            row[IcmasteSchema.User] = "admin";
            row[IcmasteSchema.PostAccno] = "6010/000";
            row[IcmasteSchema.Cust2] = TextTruncation.Truncate(fields.Value.Code, 8);
            row[IcmasteSchema.Entry] = TextTruncation.Truncate(fields.Value.Ref, 10);
            row[IcmasteSchema.CurrCode] = "MYR";
            row[IcmasteSchema.TaxCode] = "SST0";
            row[IcmasteSchema.VendNo] = TextTruncation.Truncate(fields.Value.Code, 8);
            row[IcmasteSchema.NAmt] = fields.Value.NAmt;
            row[IcmasteSchema.TAmt] = fields.Value.TAmt;
            table.Rows.Add(row);
        }

        return new StageReadResult { Table = table, RowsRead = rowsRead, SkipReasons = skipReasons };
    }

    /// <summary>
    /// Reads ictrane rows for the Stock Received workflow -- one row per valid Excel row, same
    /// skip conditions as ReadStockReceivedIcmaste (see that method's doc comment).
    /// </summary>
    public StageReadResult ReadStockReceivedIctrane(IXLWorksheet worksheet)
    {
        ValidateColumnCount(worksheet, StockReceivedExcelCol.MinColumnCount);

        var table = IctraneSchema.BuildEmptyTable();
        var skipReasons = new List<RowSkipReason>();

        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int rowsRead = 0;

        for (int excelRow = ExcelLayout.FirstDataRow; excelRow <= lastRow; excelRow++)
        {
            rowsRead++;

            var fields = ReadStockReceivedRowFields(worksheet, excelRow, skipReasons);
            if (fields is null)
            {
                continue;
            }

            var row = table.NewRow();
            row[IctraneSchema.Type] = "RE";
            row[IctraneSchema.Ref] = TextTruncation.Truncate(fields.Value.Ref, 11);
            row[IctraneSchema.ItemNo] = TextTruncation.Truncate(fields.Value.ItemNo, 24);
            row[IctraneSchema.Desc1] = TextTruncation.Truncate(fields.Value.Desc1, 60);
            row[IctraneSchema.Desc2] = TextTruncation.Truncate(fields.Value.Qty, 40);
            row[IctraneSchema.TaxCode] = "SST0";
            row[IctraneSchema.UserId] = "admin";
            row[IctraneSchema.Entry] = TextTruncation.Truncate(fields.Value.Ref, 10);
            table.Rows.Add(row);
        }

        return new StageReadResult { Table = table, RowsRead = rowsRead, SkipReasons = skipReasons };
    }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~ReadStockReceived"`
Expected: `Passed! - Failed: 0, Passed: 13, Skipped: 0, Total: 13`

- [ ] **Step 5: Commit**

```bash
git add src/eneBridge.Wpf.Core/Services/ExcelReaderService.cs tests/eneBridge.Wpf.Core.Tests/ExcelReaderServiceTests.cs
git commit -m "Add ExcelReaderService.ReadStockReceivedIcmaste/ReadStockReceivedIctrane

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 3: `DbfExportService` — append instead of recreate (TDD)

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Services/DbfExportService.cs`
- Modify: `tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs`

- [ ] **Step 1: Replace the backup-rename test with an append test, and write it to fail first**

In `tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs`, replace the entire
`Export_SecondRun_BacksUpPreviousFileInsteadOfOverwriting` test:

```csharp
    [Fact]
    public void Export_SecondRun_BacksUpPreviousFileInsteadOfOverwriting()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);

        var firstExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        Assert.True(firstExport.Success, firstExport.ErrorMessage);

        // Ensure the backup timestamp (second-resolution) differs from the first export.
        Thread.Sleep(1100);

        var secondExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        Assert.True(secondExport.Success, secondExport.ErrorMessage);

        var backups = Directory.GetFiles(_scratchFolder, "icmaste_*.dbf");
        Assert.Single(backups);
        Assert.True(File.Exists(Path.Combine(_scratchFolder, "icmaste.dbf")));
    }
```

with:

```csharp
    [Fact]
    public void Export_SecondRun_AppendsRowsInsteadOfReplacing()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);

        var firstExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        Assert.True(firstExport.Success, firstExport.ErrorMessage);
        Assert.Equal(3, firstExport.RowsWritten);

        var secondExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        Assert.True(secondExport.Success, secondExport.ErrorMessage);
        Assert.Equal(3, secondExport.RowsWritten);

        // No backup-rename file is ever created now -- rows accumulate in the live file instead.
        Assert.Empty(Directory.GetFiles(_scratchFolder, "icmaste_*.dbf"));
        Assert.True(File.Exists(Path.Combine(_scratchFolder, "icmaste.dbf")));

        AssertRowCount("icmaste", 6);
    }
```

- [ ] **Step 2: Run the test to verify it fails**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~Export_SecondRun_AppendsRowsInsteadOfReplacing"`
Expected: FAIL — `Assert.Empty(Directory.GetFiles(_scratchFolder, "icmaste_*.dbf"))` fails, because the
current code still creates a backup file (there will be 1, not 0).

- [ ] **Step 3: Rewrite `Export` to create-if-missing and always append; remove `BackupExistingFile`**

Replace the entire contents of `src/eneBridge.Wpf.Core/Services/DbfExportService.cs` with:

```csharp
using System.Data;
using System.Data.OleDb;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Exports a DataTable to a DBF file via the OleDb/Access-Database-Engine dBASE IV driver
/// (no good managed alternative for DBF exists, so this piece keeps the original's approach).
/// </summary>
public sealed class DbfExportService
{
    private readonly FileLogger? _fileLogger;

    public DbfExportService(FileLogger? fileLogger = null)
    {
        _fileLogger = fileLogger;
    }

    /// <summary>
    /// Creates the DBF file if it doesn't exist yet, then inserts every row -- rows always
    /// accumulate into the live file across calls, never replacing what's already there (see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md for why: Invoice and
    /// Stock Received both write into the same icmaste.dbf/ictrane.dbf, distinguished by TYPE, so
    /// neither can be allowed to wipe out the other's rows). A failure opening the connection /
    /// creating the table is fatal for this stage (Success=false); a failure inserting a single
    /// row is recorded in RowErrors and the remaining rows still run.
    /// </summary>
    public StageExportResult Export(string dbfFolder, string tableName, DataTable data, IReadOnlyList<DbfColumnDefinition> schema)
    {
        var rowErrors = new List<string>();

        // Checkpoint logging around every OleDb call: a native AccessViolationException from the
        // ACE dBASE driver's COM interop bypasses every catch block below and kills the process
        // outright (see DbfExportServiceTests' Export_DbfFolderCannotBeCreated comment for the
        // mechanism). FileLogger writes and closes the file on every call, so these lines survive
        // even that kind of crash and let the last line logged pin down exactly which step died.
        try
        {
            Directory.CreateDirectory(dbfFolder);

            var tableExists = File.Exists(Path.Combine(dbfFolder, tableName + ".dbf"));

            var connectionString =
                $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbfFolder};Extended Properties=""dBASE IV;CollatingSequence=1252"";";

            _fileLogger?.LogInfo($"[{tableName}] Opening OleDb connection to '{dbfFolder}'");
            using var connection = new OleDbConnection(connectionString);
            connection.Open();
            _fileLogger?.LogInfo($"[{tableName}] Connection opened");

            if (!tableExists)
            {
                _fileLogger?.LogInfo($"[{tableName}] Table does not exist yet, running CREATE TABLE");
                var createTableSql = DbfSchemaBuilder.BuildCreateTableSql(tableName, schema);
                using (var createCommand = new OleDbCommand(createTableSql, connection))
                {
                    createCommand.ExecuteNonQuery();
                }
                _fileLogger?.LogInfo($"[{tableName}] CREATE TABLE succeeded");
            }

            _fileLogger?.LogInfo($"[{tableName}] Inserting {data.Rows.Count} row(s)");

            int rowsWritten = 0;
            int rowIndex = 0;
            foreach (DataRow row in data.Rows)
            {
                rowIndex++;
                try
                {
                    var insertSql = BuildInsertSql(tableName, data, row);
                    using var insertCommand = new OleDbCommand(insertSql, connection);
                    insertCommand.ExecuteNonQuery();
                    rowsWritten++;
                }
                catch (OleDbException ex)
                {
                    rowErrors.Add($"Row {rowIndex}: {ex.Message}");
                    _fileLogger?.LogInfo($"[{tableName}] Row {rowIndex} insert failed: {ex.Message}");
                }
            }
            _fileLogger?.LogInfo($"[{tableName}] Export complete: {rowsWritten} written, {rowErrors.Count} row error(s)");

            return new StageExportResult { Success = true, RowsWritten = rowsWritten, RowErrors = rowErrors };
        }
        catch (Exception ex) when (ex is OleDbException or IOException or UnauthorizedAccessException)
        {
            _fileLogger?.LogException($"[{tableName}] Export failed", ex);
            return new StageExportResult { Success = false, ErrorMessage = ex.Message, RowErrors = rowErrors };
        }
    }

    private static string BuildInsertSql(string tableName, DataTable table, DataRow row)
    {
        var columnNames = new List<string>();
        var values = new List<string>();

        foreach (DataColumn column in table.Columns)
        {
            columnNames.Add($"[{DbfSchemaBuilder.SanitizeColumnName(column.ColumnName)}]");
            values.Add(SqlValueFormatter.Format(row[column], column.DataType));
        }

        return $"INSERT INTO {tableName} ({string.Join(", ", columnNames)}) VALUES ({string.Join(", ", values)})";
    }
}
```

(`BackupExistingFile` is gone — nothing calls it anymore.)

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~DbfExportServiceTests"`
Expected: `Passed! - Failed: 0, Passed: 4, Skipped: 0, Total: 4`

- [ ] **Step 5: Commit**

```bash
git add src/eneBridge.Wpf.Core/Services/DbfExportService.cs tests/eneBridge.Wpf.Core.Tests/DbfExportServiceTests.cs
git commit -m "Change DbfExportService to append instead of recreate on every run

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 4: `DbfReaderService` — file-existence guard + TYPE filter (TDD)

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Services/DbfReaderService.cs`
- Modify: `tests/eneBridge.Wpf.Core.Tests/DbfReaderServiceTests.cs`

- [ ] **Step 1: Write the failing test**

Add this test to `tests/eneBridge.Wpf.Core.Tests/DbfReaderServiceTests.cs`, right after
`Read_TableDoesNotExist_ReturnsFailureWithErrorMessage`:

```csharp
    [Fact]
    public void Read_WithTypeFilter_ReturnsOnlyMatchingRows()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();
        var dbfReader = new DbfReaderService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);
        dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);

        // The fixture's rows are all TYPE="IN" (Invoice). Filtering for a TYPE that doesn't exist
        // in the data should cleanly return zero rows, proving the filter actually filters rather
        // than just ignoring the extra arguments.
        var matchingResult = dbfReader.Read(_scratchFolder, IcmasteSchema.TableName, IcmasteSchema.Type, "IN");
        var nonMatchingResult = dbfReader.Read(_scratchFolder, IcmasteSchema.TableName, IcmasteSchema.Type, "RE");

        Assert.True(matchingResult.Success, matchingResult.ErrorMessage);
        Assert.Equal(3, matchingResult.RowCount);

        Assert.True(nonMatchingResult.Success, nonMatchingResult.ErrorMessage);
        Assert.Equal(0, nonMatchingResult.RowCount);
    }
```

- [ ] **Step 2: Run test to verify it fails (compile error)**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~Read_WithTypeFilter_ReturnsOnlyMatchingRows"`
Expected: build FAILS — `Read` has no 4-argument overload.

- [ ] **Step 3: Add the file-existence guard and optional TYPE filter**

Replace the entire contents of `src/eneBridge.Wpf.Core/Services/DbfReaderService.cs` with:

```csharp
using System.Data;
using System.Data.OleDb;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Reads back a table already written by <see cref="DbfExportService"/>, so the app can show
/// what actually ended up in the .dbf file rather than only what was intended to be written.
/// </summary>
public sealed class DbfReaderService
{
    private readonly FileLogger? _fileLogger;

    public DbfReaderService(FileLogger? fileLogger = null)
    {
        _fileLogger = fileLogger;
    }

    /// <summary>
    /// Reads a table, optionally filtered to rows where `typeColumnName` equals `typeValue` (used
    /// to separate Invoice's "IN" rows from Stock Received's "RE" rows in the shared tables -- see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md). Checks the table's
    /// .dbf file exists via plain file I/O before ever opening an OleDb connection: a real OleDb
    /// SELECT against a nonexistent table was found to intermittently crash the process with a
    /// native AccessViolationException/ComObject.Finalize signature, even on an otherwise-healthy
    /// driver, and callers that need a "before" row count (which can legitimately run against a
    /// table that doesn't exist yet on a genuine first-ever run) would hit that exact trigger
    /// without this guard.
    /// </summary>
    public DbfReadResult Read(string dbfFolder, string tableName, string? typeColumnName = null, string? typeValue = null)
    {
        var dbfPath = Path.Combine(dbfFolder, tableName + ".dbf");
        if (!File.Exists(dbfPath))
        {
            return new DbfReadResult { Success = false, ErrorMessage = $"'{tableName}.dbf' does not exist in '{dbfFolder}'." };
        }

        try
        {
            var connectionString =
                $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbfFolder};Extended Properties=""dBASE IV;CollatingSequence=1252"";";

            _fileLogger?.LogInfo($"[{tableName}] (verify) Opening OleDb connection to '{dbfFolder}'");
            using var connection = new OleDbConnection(connectionString);
            connection.Open();
            _fileLogger?.LogInfo($"[{tableName}] (verify) Connection opened, reading rows");

            var selectSql = $"SELECT * FROM {tableName}";
            if (!string.IsNullOrEmpty(typeColumnName))
            {
                selectSql += $" WHERE [{typeColumnName}] = {SqlValueFormatter.Format(typeValue, typeof(string))}";
            }

            var table = new DataTable();
            using (var adapter = new OleDbDataAdapter(selectSql, connection))
            {
                adapter.Fill(table);
            }
            _fileLogger?.LogInfo($"[{tableName}] (verify) Read {table.Rows.Count} row(s)");

            return new DbfReadResult { Success = true, Table = table };
        }
        catch (Exception ex) when (ex is OleDbException or IOException or UnauthorizedAccessException)
        {
            _fileLogger?.LogException($"[{tableName}] (verify) Read failed", ex);
            return new DbfReadResult { Success = false, ErrorMessage = ex.Message };
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~DbfReaderServiceTests"`
Expected: `Passed! - Failed: 0, Passed: 3, Skipped: 0, Total: 3`

- [ ] **Step 5: Manually verify the file-existence guard actually reduces the known intermittent crash**

The existing `Read_TableDoesNotExist_ReturnsFailureWithErrorMessage` test used to go through OleDb
(open a connection, then get an `OleDbException` from the missing table) and was found earlier to
crash the test host intermittently (roughly half the time) purely from that COM object churn. It
should no longer touch OleDb at all now. Run it alone, 5 times in a row, and confirm it passes
clean every time with no `AccessViolationException`:

Run (repeat 5 times):
```bash
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~Read_TableDoesNotExist_ReturnsFailureWithErrorMessage"
```
Expected each time: `Passed! - Failed: 0, Passed: 1, Skipped: 0, Total: 1` — no `Test host process crashed`.

This doesn't prove the broader, already-documented ACE driver instability is gone (it isn't — see
`CLAUDE.md`'s "Known reliability risk" note) — only that this specific trigger (a `SELECT` against
a table that was never there) no longer reaches OleDb at all.

- [ ] **Step 6: Commit**

```bash
git add src/eneBridge.Wpf.Core/Services/DbfReaderService.cs tests/eneBridge.Wpf.Core.Tests/DbfReaderServiceTests.cs
git commit -m "Add file-existence guard and TYPE filter to DbfReaderService.Read

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 5: `SettingsService` — merge-based save + Stock Received's own Excel path (TDD)

Stock Received needs its own persisted Excel path, independent of Invoice's, while sharing the
same DBF folder path. This requires `SaveUserSettings` to merge onto whatever's already saved
instead of fully replacing it — otherwise saving one tab's path would silently blank out the
other's. This task also fixes `MainViewModel`'s one call site so the whole solution keeps building.

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Models/AppSettingsModel.cs`
- Modify: `src/eneBridge.Wpf.Core/Services/SettingsService.cs`
- Create: `tests/eneBridge.Wpf.Core.Tests/SettingsServiceTests.cs`
- Modify: `src/eneBridge.Wpf/ViewModels/MainViewModel.cs`

- [ ] **Step 1: Write the failing tests**

Create `tests/eneBridge.Wpf.Core.Tests/SettingsServiceTests.cs`:

```csharp
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class SettingsServiceTests : IDisposable
{
    private readonly string _scratchFolder;
    private readonly string _appBaseDirectory;
    private readonly string _userDataDirectory;

    public SettingsServiceTests()
    {
        _scratchFolder = Path.Combine(Path.GetTempPath(), "eneBridge.Wpf.Tests." + Guid.NewGuid().ToString("N"));
        _appBaseDirectory = Path.Combine(_scratchFolder, "app");
        _userDataDirectory = Path.Combine(_scratchFolder, "appdata");
        Directory.CreateDirectory(_appBaseDirectory);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_scratchFolder, recursive: true);
        }
        catch (IOException)
        {
            // best-effort cleanup
        }
    }

    [Fact]
    public void SaveUserSettings_SavingInvoicePath_DoesNotClobberPreviouslySavedStockReceivedPath()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);

        service.SaveUserSettings(s => s.StockReceivedExcelFilePath = @"C:\stockreceived.xlsx");
        service.SaveUserSettings(s => s.ExcelFilePath = @"C:\invoice.xlsx");

        var saved = service.LoadUserSettings();
        Assert.Equal(@"C:\invoice.xlsx", saved!.ExcelFilePath);
        Assert.Equal(@"C:\stockreceived.xlsx", saved.StockReceivedExcelFilePath);
    }

    [Fact]
    public void SaveUserSettings_SavingStockReceivedPath_DoesNotClobberPreviouslySavedInvoicePath()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);

        service.SaveUserSettings(s => s.ExcelFilePath = @"C:\invoice.xlsx");
        service.SaveUserSettings(s => s.StockReceivedExcelFilePath = @"C:\stockreceived.xlsx");

        var saved = service.LoadUserSettings();
        Assert.Equal(@"C:\invoice.xlsx", saved!.ExcelFilePath);
        Assert.Equal(@"C:\stockreceived.xlsx", saved.StockReceivedExcelFilePath);
    }

    [Fact]
    public void ResolveEffectiveStockReceivedPaths_NoUserSettingsSaved_ExcelPathIsEmpty()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);

        var (excelPath, _) = service.ResolveEffectiveStockReceivedPaths(_appBaseDirectory);

        Assert.Equal(string.Empty, excelPath);
    }

    [Fact]
    public void ResolveEffectiveStockReceivedPaths_UsesSameDbfFolderAsInvoice()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);
        service.SaveUserSettings(s => s.DbfFolderPath = @"C:\emas\data");

        var (_, invoiceDbfFolder) = service.ResolveEffectivePaths(_appBaseDirectory);
        var (_, stockReceivedDbfFolder) = service.ResolveEffectiveStockReceivedPaths(_appBaseDirectory);

        Assert.Equal(@"C:\emas\data", invoiceDbfFolder);
        Assert.Equal(invoiceDbfFolder, stockReceivedDbfFolder);
    }

    [Fact]
    public void ResolveEffectiveStockReceivedPaths_UserSettingSaved_UsesUserValue()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);
        service.SaveUserSettings(s => s.StockReceivedExcelFilePath = @"C:\my-stock-received.xlsx");

        var (excelPath, _) = service.ResolveEffectiveStockReceivedPaths(_appBaseDirectory);

        Assert.Equal(@"C:\my-stock-received.xlsx", excelPath);
    }
}
```

- [ ] **Step 2: Run tests to verify they fail (compile error)**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~SettingsServiceTests"`
Expected: build FAILS — `StockReceivedExcelFilePath`/`ResolveEffectiveStockReceivedPaths` don't
exist, and `SaveUserSettings` doesn't accept a delegate.

- [ ] **Step 3: Add the new model field and default**

In `src/eneBridge.Wpf.Core/Models/AppSettingsModel.cs`, replace the whole file with:

```csharp
namespace eneBridge.Wpf.Core.Models;

/// <summary>Deserialized shape of appsettings.json's "FilePaths" section (initial defaults).</summary>
public sealed class AppSettingsModel
{
    public FilePathsSection FilePaths { get; init; } = new();

    public sealed class FilePathsSection
    {
        public string ExcelFilePath { get; init; } = string.Empty;
        public string DbfFilePath { get; init; } = string.Empty;
        public string StockReceivedExcelFilePath { get; init; } = string.Empty;
    }
}

/// <summary>Persisted last-used paths, saved to %AppData%\eneBridge\settings.json.</summary>
public sealed class UserSettingsModel
{
    public string? ExcelFilePath { get; set; }
    public string? DbfFolderPath { get; set; }
    public string? StockReceivedExcelFilePath { get; set; }
}
```

- [ ] **Step 4: Change `SaveUserSettings` to merge, and add `ResolveEffectiveStockReceivedPaths`**

Replace the whole `src/eneBridge.Wpf.Core/Services/SettingsService.cs` with:

```csharp
using System.Text.Json;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Reads appsettings.json (initial defaults) and reads/writes a persisted user settings file
/// (%AppData%\eneBridge\settings.json) that stores the last-used Excel/DBF paths.
/// </summary>
public sealed class SettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _appSettingsPath;
    private readonly string _userSettingsPath;

    public SettingsService(string appBaseDirectory, string? userDataDirectory = null)
    {
        _appSettingsPath = Path.Combine(appBaseDirectory, "appsettings.json");
        var dataDir = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "eneBridge");
        Directory.CreateDirectory(dataDir);
        _userSettingsPath = Path.Combine(dataDir, "settings.json");
    }

    public AppSettingsModel LoadAppDefaults()
    {
        if (!File.Exists(_appSettingsPath))
        {
            return new AppSettingsModel();
        }
        try
        {
            var json = File.ReadAllText(_appSettingsPath);
            return JsonSerializer.Deserialize<AppSettingsModel>(json, JsonOptions) ?? new AppSettingsModel();
        }
        catch (Exception)
        {
            return new AppSettingsModel();
        }
    }

    public UserSettingsModel? LoadUserSettings()
    {
        if (!File.Exists(_userSettingsPath))
        {
            return null;
        }
        try
        {
            var json = File.ReadAllText(_userSettingsPath);
            return JsonSerializer.Deserialize<UserSettingsModel>(json, JsonOptions);
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// Loads whatever is currently persisted (or starts fresh if nothing is saved yet), applies
    /// `update`, and writes the merged result back. Merge-based rather than replace-based so that
    /// Invoice and Stock Received -- which each save only the fields they own -- never clobber
    /// each other's saved paths.
    /// </summary>
    public void SaveUserSettings(Action<UserSettingsModel> update)
    {
        var current = LoadUserSettings() ?? new UserSettingsModel();
        update(current);

        var json = JsonSerializer.Serialize(current, JsonOptions);
        var tempPath = _userSettingsPath + ".tmp";
        File.WriteAllText(tempPath, json);
        File.Move(tempPath, _userSettingsPath, overwrite: true);
    }

    /// <summary>User-saved paths win over appsettings.json defaults (which are combined with the app's base directory, same as the original).</summary>
    public (string excelPath, string dbfFolder) ResolveEffectivePaths(string appBaseDirectory)
    {
        var defaults = LoadAppDefaults();
        var user = LoadUserSettings();

        string excelPath = !string.IsNullOrWhiteSpace(user?.ExcelFilePath)
            ? user!.ExcelFilePath!
            : Path.Combine(appBaseDirectory, defaults.FilePaths.ExcelFilePath);

        string dbfFolder = !string.IsNullOrWhiteSpace(user?.DbfFolderPath)
            ? user!.DbfFolderPath!
            : Path.Combine(appBaseDirectory, defaults.FilePaths.DbfFilePath);

        return (excelPath, dbfFolder);
    }

    /// <summary>
    /// Stock Received's own Excel path, resolved independently of Invoice's -- but the SAME DBF
    /// folder Invoice uses, since both workflows now write into the same live tables (see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md). Unlike
    /// ResolveEffectivePaths, an unset default yields an empty string rather than a combined path
    /// pointing at the app's own base directory -- no app-shipped default is required, the field
    /// just starts blank until the user Browses once.
    /// </summary>
    public (string excelPath, string dbfFolder) ResolveEffectiveStockReceivedPaths(string appBaseDirectory)
    {
        var defaults = LoadAppDefaults();
        var user = LoadUserSettings();

        string excelPath = !string.IsNullOrWhiteSpace(user?.StockReceivedExcelFilePath)
            ? user!.StockReceivedExcelFilePath!
            : (string.IsNullOrWhiteSpace(defaults.FilePaths.StockReceivedExcelFilePath)
                ? string.Empty
                : Path.Combine(appBaseDirectory, defaults.FilePaths.StockReceivedExcelFilePath));

        string dbfFolder = !string.IsNullOrWhiteSpace(user?.DbfFolderPath)
            ? user!.DbfFolderPath!
            : Path.Combine(appBaseDirectory, defaults.FilePaths.DbfFilePath);

        return (excelPath, dbfFolder);
    }
}
```

- [ ] **Step 5: Fix `MainViewModel`'s call site to match the new signature**

In `src/eneBridge.Wpf/ViewModels/MainViewModel.cs`, replace:

```csharp
    private void SaveUserPaths()
    {
        _settingsService.SaveUserSettings(new UserSettingsModel
        {
            ExcelFilePath = ExcelFilePath,
            DbfFolderPath = DbfFolderPath
        });
    }
```

with:

```csharp
    private void SaveUserPaths()
    {
        _settingsService.SaveUserSettings(s =>
        {
            s.ExcelFilePath = ExcelFilePath;
            s.DbfFolderPath = DbfFolderPath;
        });
    }
```

- [ ] **Step 6: Run tests and build to verify everything passes**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~SettingsServiceTests"`
Expected: `Passed! - Failed: 0, Passed: 5, Skipped: 0, Total: 5`

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 7: Commit**

```bash
git add src/eneBridge.Wpf.Core/Models/AppSettingsModel.cs src/eneBridge.Wpf.Core/Services/SettingsService.cs tests/eneBridge.Wpf.Core.Tests/SettingsServiceTests.cs src/eneBridge.Wpf/ViewModels/MainViewModel.cs
git commit -m "Add merge-based settings save and Stock Received's own Excel path

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 6: `RunHistoryService` — optional custom file name (TDD)

Stock Received gets its own run history, kept in a separate file from Invoice's so the two
workflows' histories aren't mixed together.

**Files:**
- Modify: `src/eneBridge.Wpf.Core/Services/RunHistoryService.cs`
- Create: `tests/eneBridge.Wpf.Core.Tests/RunHistoryServiceTests.cs`

- [ ] **Step 1: Write the failing tests**

Create `tests/eneBridge.Wpf.Core.Tests/RunHistoryServiceTests.cs`:

```csharp
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class RunHistoryServiceTests : IDisposable
{
    private readonly string _scratchFolder;

    public RunHistoryServiceTests()
    {
        _scratchFolder = Path.Combine(Path.GetTempPath(), "eneBridge.Wpf.Tests." + Guid.NewGuid().ToString("N"));
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_scratchFolder, recursive: true);
        }
        catch (IOException)
        {
            // best-effort cleanup
        }
    }

    private static RunHistoryEntry SampleEntry() => new() { ExcelPath = "x.xlsx", DbfFolder = @"C:\dbf" };

    [Fact]
    public void Append_DefaultFileName_WritesToRunHistoryJson()
    {
        var service = new RunHistoryService(_scratchFolder);

        service.Append(SampleEntry());

        Assert.True(File.Exists(Path.Combine(_scratchFolder, "runHistory.json")));
    }

    [Fact]
    public void Append_CustomFileName_WritesToThatFileNotTheDefault()
    {
        var service = new RunHistoryService(_scratchFolder, "stockReceivedRunHistory.json");

        service.Append(SampleEntry());

        Assert.True(File.Exists(Path.Combine(_scratchFolder, "stockReceivedRunHistory.json")));
        Assert.False(File.Exists(Path.Combine(_scratchFolder, "runHistory.json")));
    }

    [Fact]
    public void LoadAll_TwoServicesWithDifferentFileNames_DoNotSeeEachOthersEntries()
    {
        var invoiceService = new RunHistoryService(_scratchFolder);
        var stockReceivedService = new RunHistoryService(_scratchFolder, "stockReceivedRunHistory.json");

        invoiceService.Append(SampleEntry());

        Assert.Single(invoiceService.LoadAll());
        Assert.Empty(stockReceivedService.LoadAll());
    }
}
```

- [ ] **Step 2: Run tests to verify they fail (compile error)**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~RunHistoryServiceTests"`
Expected: build FAILS — no `RunHistoryService(string, string)` constructor overload.

- [ ] **Step 3: Add the optional `fileName` parameter**

In `src/eneBridge.Wpf.Core/Services/RunHistoryService.cs`, replace:

```csharp
    public RunHistoryService(string? userDataDirectory = null)
    {
        var dataDir = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "eneBridge");
        Directory.CreateDirectory(dataDir);
        _historyPath = Path.Combine(dataDir, "runHistory.json");
    }
```

with:

```csharp
    public RunHistoryService(string? userDataDirectory = null, string fileName = "runHistory.json")
    {
        var dataDir = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "eneBridge");
        Directory.CreateDirectory(dataDir);
        _historyPath = Path.Combine(dataDir, fileName);
    }
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~RunHistoryServiceTests"`
Expected: `Passed! - Failed: 0, Passed: 3, Skipped: 0, Total: 3`

- [ ] **Step 5: Commit**

```bash
git add src/eneBridge.Wpf.Core/Services/RunHistoryService.cs tests/eneBridge.Wpf.Core.Tests/RunHistoryServiceTests.cs
git commit -m "Add optional custom file name to RunHistoryService

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 7: `DbfWorkflowHelper` (new shared service) + rewire `MainViewModel` to use it

The ACE-driver-crash-avoidance table check (currently private methods on `MainViewModel`) and the
new append-aware before/after-delta verification are both about to be needed identically by two
ViewModels. Rather than duplicate ~90 lines of safety-critical logic, this task extracts them into
a shared static helper and proves it works by wiring `MainViewModel` through it first — Task 8
builds `StockReceivedViewModel` on top of this already-proven helper.

Not unit-tested: `ConfirmTablesReadable` calls `MessageBox.Show` (no UI test framework in this
project — same reasoning already accepted for `MainViewModel` having no test file), and the
delta-verification logic is a thin orchestration of already-tested `DbfReaderService` calls.
Verified via manual end-to-end testing in Task 12 instead.

**Files:**
- Create: `src/eneBridge.Wpf/Services/DbfWorkflowHelper.cs`
- Modify: `src/eneBridge.Wpf/ViewModels/MainViewModel.cs`

- [ ] **Step 1: Create the shared helper**

```csharp
using System.Data;
using System.Windows;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Services;

/// <summary>
/// Shared DBF safety/verification helpers used by both MainViewModel (Invoice) and
/// StockReceivedViewModel, so the ACE-driver-crash-avoidance table check and the append-aware
/// before/after-delta verification only exist in one place. See
/// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md.
/// </summary>
public static class DbfWorkflowHelper
{
    /// <summary>
    /// Warns (with a Continue/Cancel choice) if either table can't currently be read — covers
    /// "file doesn't exist", "table locked by EMAS", and "folder unreachable" alike. Deliberately
    /// checks via the filesystem only (File.Exists/Directory.Exists/a shared-read FileStream probe)
    /// rather than DbfReaderService.Read: a real OleDb SELECT against a nonexistent table was found
    /// to intermittently crash the process with the same native AccessViolationException/
    /// ComObject.Finalize signature this whole feature exists to guard against — reproduced even on
    /// a machine with an otherwise-healthy ACE driver. Using plain file I/O here removes OleDb from
    /// this code path entirely, matching DbfSafetyBackupService's same discipline. Checks both
    /// tables even if the first fails, unless the user cancels on the first warning.
    /// </summary>
    public static bool ConfirmTablesReadable(string dbfFolder)
    {
        foreach (var tableName in new[] { IcmasteSchema.TableName, IctraneSchema.TableName })
        {
            var (isReadable, errorMessage) = CheckTableFileReadable(dbfFolder, tableName);
            if (isReadable)
            {
                continue;
            }

            var proceed = MessageBox.Show(
                $"{tableName}.dbf could not be found or read in '{dbfFolder}':\n{errorMessage}\n\n" +
                "This is expected on a first-ever export to this folder, but if you expect this " +
                "table to already exist, something may be wrong (wrong folder selected, the table " +
                "is locked by EMAS, or a previous export was interrupted).\n\n" +
                "Continue anyway?",
                "eneBridge - Table Check",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (proceed != MessageBoxResult.Yes)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Pure filesystem readability probe for one table's .dbf file — no OleDb involved. Opens with
    /// FileShare.ReadWrite (the most permissive request our own read can make) so this only reports
    /// "locked" when another process genuinely holds an incompatible lock, not merely because
    /// something else has the file open for reading too.
    /// </summary>
    private static (bool IsReadable, string? ErrorMessage) CheckTableFileReadable(string dbfFolder, string tableName)
    {
        if (!Directory.Exists(dbfFolder))
        {
            return (false, $"The folder '{dbfFolder}' does not exist or is not reachable.");
        }

        var dbfPath = Path.Combine(dbfFolder, tableName + ".dbf");
        if (!File.Exists(dbfPath))
        {
            return (false, $"'{tableName}.dbf' does not exist in this folder.");
        }

        try
        {
            using var stream = new FileStream(dbfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return (true, null);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return (false, ex.Message);
        }
    }

    /// <summary>
    /// Counts how many rows currently match `typeValue` in `tableName` — used as the "before"
    /// count for the append-aware delta verification below. Returns 0 without the caller needing
    /// to special-case a missing table: DbfReaderService.Read already returns a clean failure (and
    /// RowCount 0) without touching OleDb when the table's .dbf file doesn't exist yet, which is
    /// exactly the case on a genuine first-ever run.
    /// </summary>
    public static async Task<int> CountRowsByTypeAsync(
        DbfReaderService dbfReaderService, string dbfFolder, string tableName, string typeColumnName, string typeValue)
    {
        var result = await Task.Run(() => dbfReaderService.Read(dbfFolder, tableName, typeColumnName, typeValue));
        return result.Success ? result.RowCount : 0;
    }

    /// <summary>
    /// Reads the table back (filtered to `typeValue`) after an export, and compares how many new
    /// matching rows appeared against what the export reported writing — an append-aware delta
    /// check, since rows now accumulate across runs instead of the file being fully replaced each
    /// time (a plain "total rows == rows written" comparison would show a false mismatch on every
    /// run after the first). Never throws.
    /// </summary>
    public static async Task VerifyDbfDeltaAsync(
        DbfReaderService dbfReaderService,
        string tableName,
        string dbfFolder,
        string typeColumnName,
        string typeValue,
        int beforeCount,
        StageExportResult? exportResult,
        Action<DataView?> setPreview,
        Action<string> setStatusText,
        Action<bool> setVerified)
    {
        var readResult = await Task.Run(() => dbfReaderService.Read(dbfFolder, tableName, typeColumnName, typeValue));
        setPreview(readResult.Success ? readResult.Table.DefaultView : null);

        if (!readResult.Success)
        {
            setStatusText($"Could not verify: {readResult.ErrorMessage}");
            setVerified(false);
            return;
        }

        var expectedNewRows = exportResult?.RowsWritten ?? 0;
        var actualNewRows = readResult.RowCount - beforeCount;

        if ((exportResult?.Success ?? false) && expectedNewRows == actualNewRows)
        {
            setStatusText($"{tableName}: {actualNewRows} row(s) written and confirmed in DBF ✓ ({readResult.RowCount} total)");
            setVerified(true);
        }
        else if (exportResult?.Success ?? false)
        {
            setStatusText($"{tableName}: wrote {expectedNewRows} row(s) but DBF shows {actualNewRows} new row(s) ⚠ ({readResult.RowCount} total)");
            setVerified(false);
        }
        else
        {
            setStatusText($"{tableName}: export failed — DBF currently has {readResult.RowCount} row(s) (may include previous runs)");
            setVerified(false);
        }
    }
}
```

- [ ] **Step 2: Remove `MainViewModel`'s now-redundant private methods**

In `src/eneBridge.Wpf/ViewModels/MainViewModel.cs`, delete these three private methods entirely
(they're being replaced by `DbfWorkflowHelper`):

- `ConfirmTablesReadable(string dbfFolder)` (the one with the `foreach` over `IcmasteSchema.TableName`/`IctraneSchema.TableName`)
- `CheckTableFileReadable(string dbfFolder, string tableName)`
- `VerifyDbfAsync(...)` (the one with `Action<DataView?> setPreview, Action<string> setStatusText, Action<bool> setVerified` parameters)

- [ ] **Step 3: Rewire `ConfirmExportAsync` to use the shared helper and the delta check**

Replace the entire `ConfirmExportAsync` method body with:

```csharp
    [RelayCommand(CanExecute = nameof(CanExport))]
    private async Task ConfirmExportAsync()
    {
        IsRunning = true;
        PreviewCommand.NotifyCanExecuteChanged();
        ConfirmExportCommand.NotifyCanExecuteChanged();
        IcmasteDbfPreview = null;
        IctraneDbfPreview = null;
        IcmasteDbfStatusText = "Verifying…";
        IctraneDbfStatusText = "Verifying…";
        IcmasteDbfVerified = false;
        IctraneDbfVerified = false;

        string excelPath = ExcelFilePath;
        string dbfFolder = DbfFolderPath;
        var stopwatch = Stopwatch.StartNew();

        StageExportResult? icmasteExport = null;
        StageExportResult? ictraneExport = null;
        string? fatalError = null;
        bool cancelledByUser = false;

        try
        {
            await Task.Run(() =>
            {
                _dbfSafetyBackupService.BackupIfExists(dbfFolder, IcmasteSchema.TableName);
                _dbfSafetyBackupService.BackupIfExists(dbfFolder, IctraneSchema.TableName);
            });

            if (!DbfWorkflowHelper.ConfirmTablesReadable(dbfFolder))
            {
                cancelledByUser = true;
                IcmasteDbfStatusText = "Run Confirm & Export to verify.";
                IctraneDbfStatusText = "Run Confirm & Export to verify.";
                AppendLog("Export cancelled by user after the table check.");
                return;
            }

            var icmasteBeforeCount = await DbfWorkflowHelper.CountRowsByTypeAsync(
                _dbfReaderService, dbfFolder, IcmasteSchema.TableName, IcmasteSchema.Type, "IN");
            icmasteExport = await ExportStageAsync(
                "icmaste",
                IcmasteStage,
                _icmasteReadResult!,
                result => _dbfExportService.Export(dbfFolder, IcmasteSchema.TableName, result.Table, IcmasteSchema.Columns));
            await DbfWorkflowHelper.VerifyDbfDeltaAsync(
                _dbfReaderService, IcmasteSchema.TableName, dbfFolder, IcmasteSchema.Type, "IN", icmasteBeforeCount, icmasteExport,
                v => IcmasteDbfPreview = v, s => IcmasteDbfStatusText = s, b => IcmasteDbfVerified = b);

            var ictraneBeforeCount = await DbfWorkflowHelper.CountRowsByTypeAsync(
                _dbfReaderService, dbfFolder, IctraneSchema.TableName, IctraneSchema.Type, "IN");
            ictraneExport = await ExportStageAsync(
                "ictrane",
                IctraneStage,
                _ictraneReadResult!,
                result => _dbfExportService.Export(dbfFolder, IctraneSchema.TableName, result.Table, IctraneSchema.Columns));
            await DbfWorkflowHelper.VerifyDbfDeltaAsync(
                _dbfReaderService, IctraneSchema.TableName, dbfFolder, IctraneSchema.Type, "IN", ictraneBeforeCount, ictraneExport,
                v => IctraneDbfPreview = v, s => IctraneDbfStatusText = s, b => IctraneDbfVerified = b);
        }
        catch (Exception ex)
        {
            // Outer backstop: nothing should ever crash the app.
            _fileLogger.LogException("Unhandled error during export", ex);
            AppendLog($"Unexpected error: {ex.Message}");
            fatalError = ex.Message;
        }
        finally
        {
            stopwatch.Stop();
            if (!cancelledByUser)
            {
                var historyEntry = new RunHistoryEntry
                {
                    ExcelPath = excelPath,
                    DbfFolder = dbfFolder,
                    IcmasteRowsRead = _icmasteReadResult?.RowsRead ?? 0,
                    IcmasteRowsSkipped = _icmasteReadResult?.SkipReasons.Count ?? 0,
                    IcmasteRowsWritten = icmasteExport?.RowsWritten ?? 0,
                    IcmasteSuccess = icmasteExport?.Success ?? false,
                    IctraneRowsRead = _ictraneReadResult?.RowsRead ?? 0,
                    IctraneRowsSkipped = _ictraneReadResult?.SkipReasons.Count ?? 0,
                    IctraneRowsWritten = ictraneExport?.RowsWritten ?? 0,
                    IctraneSuccess = ictraneExport?.Success ?? false,
                    ErrorSummary = fatalError,
                    DurationMs = stopwatch.ElapsedMilliseconds,
                };
                _runHistoryService.Append(historyEntry);
                RunHistory.Insert(0, historyEntry);
            }
            IsRunning = false;
            PreviewCommand.NotifyCanExecuteChanged();
            ConfirmExportCommand.NotifyCanExecuteChanged();
        }
    }
```

(`MainViewModel.cs` already has `using eneBridge.Wpf.Services;` from earlier work, so no new
`using` is needed to reach `DbfWorkflowHelper`.)

- [ ] **Step 4: Build to verify it compiles**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 5: Manually verify Invoice's export still works end-to-end against a scratch folder**

```bash
mkdir -p /tmp/enebridge-manual-verify-task7
```

Run the app (`dotnet run --project "src/eneBridge.Wpf/eneBridge.Wpf.csproj"`), point the DBF
folder field at that empty scratch folder, Preview against the real fixture/staged Excel file,
then Confirm & Export twice in a row.

Expected:
- First run: Table Check dialog appears (folder is empty, first-ever run) — Continue proceeds,
  export succeeds, both DBF preview tabs show a green "✓ ... confirmed" status.
- Second run: no Table Check dialog (tables now exist and are readable) — export succeeds, status
  text shows the new row count "written and confirmed", and the underlying `.dbf` files now
  contain **both** runs' rows (not just the second run's) — confirm by checking the row counts
  shown are cumulative, not reset.

- [ ] **Step 6: Commit**

```bash
git add src/eneBridge.Wpf/Services/DbfWorkflowHelper.cs src/eneBridge.Wpf/ViewModels/MainViewModel.cs
git commit -m "Extract DbfWorkflowHelper and rewire Invoice's export to the delta-aware check

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 8: `StockReceivedViewModel` (new)

Mirrors `MainViewModel`'s Invoice-related surface — same shape, same commands, same progress/
preview state — built on top of `DbfWorkflowHelper` (Task 7) and the Stock Received read methods
(Task 2). Deliberately duplicates `ReadStageAsync`/`ExportStageAsync`/`AppendLog`/`TryOpenWorkbook`
rather than sharing them with `MainViewModel` (approved as part of the design — extracting a full
shared base class was judged a bigger, riskier change than this task calls for); only the
safety-critical table-check and delta-verification logic was worth the extraction risk, and that
already happened in Task 7.

Not unit-tested — consistent with `MainViewModel` having no test file (real service dependencies,
WPF-bound state, disproportionate scaffolding for what's being tested). Verified manually in
Task 12.

**Files:**
- Create: `src/eneBridge.Wpf/ViewModels/StockReceivedViewModel.cs`

- [ ] **Step 1: Create the file**

```csharp
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;
using eneBridge.Wpf.Services;

namespace eneBridge.Wpf.ViewModels;

/// <summary>
/// Stock Received workflow — reads Self-Billed_Format-ORI.xlsx-shaped workbooks and writes into
/// the SAME icmaste.dbf/ictrane.dbf Invoice writes to (TYPE="RE" vs Invoice's "IN"), appending
/// rather than recreating. Mirrors MainViewModel's Invoice-related surface; see
/// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md for the full rationale.
/// </summary>
public partial class StockReceivedViewModel : ObservableObject
{
    private readonly ExcelReaderService _excelReaderService;
    private readonly DbfExportService _dbfExportService;
    private readonly DbfReaderService _dbfReaderService;
    private readonly DbfSafetyBackupService _dbfSafetyBackupService;
    private readonly SettingsService _settingsService;
    private readonly RunHistoryService _runHistoryService;
    private readonly FileLogger _fileLogger;
    private readonly ExcelSourceStagingService _excelSourceStagingService;
    private readonly string _appBaseDirectory;

    private StageReadResult? _icmasteReadResult;
    private StageReadResult? _ictraneReadResult;

    [ObservableProperty]
    private string _excelFilePath = string.Empty;

    [ObservableProperty]
    private string _dbfFolderPath = string.Empty;

    [ObservableProperty]
    private bool _isRunning;

    [ObservableProperty]
    private DataView? _icmastePreview;

    [ObservableProperty]
    private DataView? _ictranePreview;

    [ObservableProperty]
    private DataView? _icmasteDbfPreview;

    [ObservableProperty]
    private DataView? _ictraneDbfPreview;

    [ObservableProperty]
    private string _icmasteDbfStatusText = "Run Confirm & Export to verify.";

    [ObservableProperty]
    private string _ictraneDbfStatusText = "Run Confirm & Export to verify.";

    [ObservableProperty]
    private bool _icmasteDbfVerified;

    [ObservableProperty]
    private bool _ictraneDbfVerified;

    [ObservableProperty]
    private bool _showAllColumns;

    /// <summary>
    /// Reflects the SAME ACE-driver health check MainViewModel runs once per session (there's only
    /// one machine-level driver to check, not one per tab) — set via SetAceEngineHealth, called by
    /// MainViewModel whenever its own check completes.
    /// </summary>
    [ObservableProperty]
    private bool _aceEngineHealthy = true;

    [ObservableProperty]
    private string _aceEngineStatusText = string.Empty;

    public ObservableCollection<RunHistoryEntry> RunHistory { get; } = [];
    public ObservableCollection<string> LogLines { get; } = [];

    public StageProgressViewModel IcmasteStage { get; } = new("icmaste");
    public StageProgressViewModel IctraneStage { get; } = new("ictrane");

    public StockReceivedViewModel(
        ExcelReaderService excelReaderService,
        DbfExportService dbfExportService,
        DbfReaderService dbfReaderService,
        DbfSafetyBackupService dbfSafetyBackupService,
        SettingsService settingsService,
        RunHistoryService runHistoryService,
        FileLogger fileLogger,
        ExcelSourceStagingService excelSourceStagingService,
        string appBaseDirectory)
    {
        _excelReaderService = excelReaderService;
        _dbfExportService = dbfExportService;
        _dbfReaderService = dbfReaderService;
        _dbfSafetyBackupService = dbfSafetyBackupService;
        _settingsService = settingsService;
        _runHistoryService = runHistoryService;
        _fileLogger = fileLogger;
        _excelSourceStagingService = excelSourceStagingService;
        _appBaseDirectory = appBaseDirectory;
    }

    public void Initialize()
    {
        var (excelPath, dbfFolder) = _settingsService.ResolveEffectiveStockReceivedPaths(_appBaseDirectory);
        ExcelFilePath = excelPath;
        DbfFolderPath = dbfFolder;

        // Most recent run first.
        foreach (var entry in _runHistoryService.LoadAll().AsEnumerable().Reverse())
        {
            RunHistory.Add(entry);
        }
    }

    /// <summary>Called by MainViewModel whenever the shared ACE-driver health check result changes.</summary>
    public void SetAceEngineHealth(bool healthy)
    {
        AceEngineHealthy = healthy;
        AceEngineStatusText = healthy
            ? string.Empty
            : "Access Database Engine isn't working correctly — Confirm & Export is disabled until this is fixed.";
        ConfirmExportCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    private void BrowseExcel()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            FileName = ExcelFilePath
        };
        if (dialog.ShowDialog() == true)
        {
            ExcelFilePath = StageExcelSource(dialog.FileName);
            SaveUserPaths();
        }
    }

    private const string StockReceivedStagedFileName = "stockreceived.xlsx";

    /// <summary>
    /// Copies the picked Excel file into the local staging folder so later runs don't depend on
    /// removable/network media staying plugged in. Never blocks the user: if staging fails for any
    /// reason, the failure is logged and the original picked path is used as-is.
    /// </summary>
    private string StageExcelSource(string originalPath)
    {
        try
        {
            var stagedPath = _excelSourceStagingService.StageFile(originalPath, StockReceivedStagedFileName);
            _fileLogger.LogInfo($"Imported Stock Received Excel source '{originalPath}' -> '{stagedPath}'");
            return stagedPath;
        }
        catch (Exception ex)
        {
            _fileLogger.LogException("Failed to stage Stock Received Excel source file", ex);
            return originalPath;
        }
    }

    [RelayCommand]
    private void BrowseDbfFolder()
    {
        var dialog = new Microsoft.Win32.OpenFolderDialog
        {
            InitialDirectory = Directory.Exists(DbfFolderPath) ? DbfFolderPath : string.Empty
        };
        if (dialog.ShowDialog() == true)
        {
            DbfFolderPath = dialog.FolderName;
            SaveUserPaths();
        }
    }

    private void SaveUserPaths()
    {
        _settingsService.SaveUserSettings(s =>
        {
            s.StockReceivedExcelFilePath = ExcelFilePath;
            s.DbfFolderPath = DbfFolderPath;
        });
    }

    partial void OnExcelFilePathChanged(string value) => InvalidatePreview();
    partial void OnDbfFolderPathChanged(string value) => InvalidatePreview();

    private void InvalidatePreview()
    {
        _icmasteReadResult = null;
        _ictraneReadResult = null;
        ConfirmExportCommand.NotifyCanExecuteChanged();
    }

    private bool CanPreview() => !IsRunning;

    /// <summary>
    /// Reads the Excel file and populates the previews/skip-reason log. Writes nothing to disk —
    /// this is the pre-verify checkpoint before Confirm &amp; Export actually writes the DBF files.
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanPreview))]
    private async Task PreviewAsync()
    {
        IsRunning = true;
        PreviewCommand.NotifyCanExecuteChanged();
        ConfirmExportCommand.NotifyCanExecuteChanged();
        LogLines.Clear();
        IcmasteStage.Reset();
        IctraneStage.Reset();
        _icmasteReadResult = null;
        _ictraneReadResult = null;

        string excelPath = ExcelFilePath;

        try
        {
            var openResult = await Task.Run(() => TryOpenWorkbook(excelPath));

            if (openResult.Workbook is null)
            {
                var message = openResult.Error ?? "Failed to open the Excel file.";
                AppendLog($"Cannot open Excel file: {message}");
                IcmasteStage.Fail(message);
                IctraneStage.Fail(message);
            }
            else
            {
                using (openResult.Workbook)
                {
                    var workbook = openResult.Workbook;
                    var worksheet = workbook.Worksheets.First();

                    _icmasteReadResult = await ReadStageAsync(
                        IcmasteStage,
                        () => _excelReaderService.ReadStockReceivedIcmaste(worksheet),
                        result => IcmastePreview = result.Table.DefaultView);

                    _ictraneReadResult = await ReadStageAsync(
                        IctraneStage,
                        () => _excelReaderService.ReadStockReceivedIctrane(worksheet),
                        result => IctranePreview = result.Table.DefaultView);
                }
            }
        }
        catch (Exception ex)
        {
            // Outer backstop: nothing should ever crash the app.
            _fileLogger.LogException("Unhandled error during Stock Received preview", ex);
            AppendLog($"Unexpected error: {ex.Message}");
        }
        finally
        {
            IsRunning = false;
            PreviewCommand.NotifyCanExecuteChanged();
            ConfirmExportCommand.NotifyCanExecuteChanged();
        }
    }

    private bool CanExport() => !IsRunning && AceEngineHealthy && _icmasteReadResult is not null && _ictraneReadResult is not null;

    /// <summary>Writes the DBF files from the read results Preview already produced.</summary>
    [RelayCommand(CanExecute = nameof(CanExport))]
    private async Task ConfirmExportAsync()
    {
        IsRunning = true;
        PreviewCommand.NotifyCanExecuteChanged();
        ConfirmExportCommand.NotifyCanExecuteChanged();
        IcmasteDbfPreview = null;
        IctraneDbfPreview = null;
        IcmasteDbfStatusText = "Verifying…";
        IctraneDbfStatusText = "Verifying…";
        IcmasteDbfVerified = false;
        IctraneDbfVerified = false;

        string excelPath = ExcelFilePath;
        string dbfFolder = DbfFolderPath;
        var stopwatch = Stopwatch.StartNew();

        StageExportResult? icmasteExport = null;
        StageExportResult? ictraneExport = null;
        string? fatalError = null;
        bool cancelledByUser = false;

        try
        {
            await Task.Run(() =>
            {
                _dbfSafetyBackupService.BackupIfExists(dbfFolder, IcmasteSchema.TableName);
                _dbfSafetyBackupService.BackupIfExists(dbfFolder, IctraneSchema.TableName);
            });

            if (!DbfWorkflowHelper.ConfirmTablesReadable(dbfFolder))
            {
                cancelledByUser = true;
                IcmasteDbfStatusText = "Run Confirm & Export to verify.";
                IctraneDbfStatusText = "Run Confirm & Export to verify.";
                AppendLog("Export cancelled by user after the table check.");
                return;
            }

            var icmasteBeforeCount = await DbfWorkflowHelper.CountRowsByTypeAsync(
                _dbfReaderService, dbfFolder, IcmasteSchema.TableName, IcmasteSchema.Type, "RE");
            icmasteExport = await ExportStageAsync(
                "icmaste",
                IcmasteStage,
                _icmasteReadResult!,
                result => _dbfExportService.Export(dbfFolder, IcmasteSchema.TableName, result.Table, IcmasteSchema.Columns));
            await DbfWorkflowHelper.VerifyDbfDeltaAsync(
                _dbfReaderService, IcmasteSchema.TableName, dbfFolder, IcmasteSchema.Type, "RE", icmasteBeforeCount, icmasteExport,
                v => IcmasteDbfPreview = v, s => IcmasteDbfStatusText = s, b => IcmasteDbfVerified = b);

            var ictraneBeforeCount = await DbfWorkflowHelper.CountRowsByTypeAsync(
                _dbfReaderService, dbfFolder, IctraneSchema.TableName, IctraneSchema.Type, "RE");
            ictraneExport = await ExportStageAsync(
                "ictrane",
                IctraneStage,
                _ictraneReadResult!,
                result => _dbfExportService.Export(dbfFolder, IctraneSchema.TableName, result.Table, IctraneSchema.Columns));
            await DbfWorkflowHelper.VerifyDbfDeltaAsync(
                _dbfReaderService, IctraneSchema.TableName, dbfFolder, IctraneSchema.Type, "RE", ictraneBeforeCount, ictraneExport,
                v => IctraneDbfPreview = v, s => IctraneDbfStatusText = s, b => IctraneDbfVerified = b);
        }
        catch (Exception ex)
        {
            // Outer backstop: nothing should ever crash the app.
            _fileLogger.LogException("Unhandled error during Stock Received export", ex);
            AppendLog($"Unexpected error: {ex.Message}");
            fatalError = ex.Message;
        }
        finally
        {
            stopwatch.Stop();
            if (!cancelledByUser)
            {
                var historyEntry = new RunHistoryEntry
                {
                    ExcelPath = excelPath,
                    DbfFolder = dbfFolder,
                    IcmasteRowsRead = _icmasteReadResult?.RowsRead ?? 0,
                    IcmasteRowsSkipped = _icmasteReadResult?.SkipReasons.Count ?? 0,
                    IcmasteRowsWritten = icmasteExport?.RowsWritten ?? 0,
                    IcmasteSuccess = icmasteExport?.Success ?? false,
                    IctraneRowsRead = _ictraneReadResult?.RowsRead ?? 0,
                    IctraneRowsSkipped = _ictraneReadResult?.SkipReasons.Count ?? 0,
                    IctraneRowsWritten = ictraneExport?.RowsWritten ?? 0,
                    IctraneSuccess = ictraneExport?.Success ?? false,
                    ErrorSummary = fatalError,
                    DurationMs = stopwatch.ElapsedMilliseconds,
                };
                _runHistoryService.Append(historyEntry);
                RunHistory.Insert(0, historyEntry);
            }
            IsRunning = false;
            PreviewCommand.NotifyCanExecuteChanged();
            ConfirmExportCommand.NotifyCanExecuteChanged();
        }
    }

    /// <summary>Reads one table, updating the stage VM. Never throws.</summary>
    private async Task<StageReadResult> ReadStageAsync(
        StageProgressViewModel stage,
        Func<StageReadResult> read,
        Action<StageReadResult> onRead)
    {
        var readResult = await Task.Run(read);
        onRead(readResult);

        var readyCount = readResult.Table.Rows.Count;
        stage.Complete(true, $"Read {readResult.RowsRead}, skipped {readResult.SkipReasons.Count}, {readyCount} ready to export");

        return readResult;
    }

    /// <summary>Exports one already-read table, logging row errors and updating the stage VM. Never throws.</summary>
    private async Task<StageExportResult?> ExportStageAsync(
        string label,
        StageProgressViewModel stage,
        StageReadResult readResult,
        Func<StageReadResult, StageExportResult> export)
    {
        try
        {
            var exportResult = await Task.Run(() => export(readResult));
            foreach (var rowError in exportResult.RowErrors)
            {
                AppendLog($"[{label}] {rowError}");
            }

            if (exportResult.Success)
            {
                AppendLog($"[{label}] Done — written {exportResult.RowsWritten} row(s).");
                stage.Complete(true, $"Read {readResult.RowsRead}, skipped {readResult.SkipReasons.Count}, written {exportResult.RowsWritten}");
            }
            else
            {
                AppendLog($"[{label}] Export failed: {exportResult.ErrorMessage}");
                stage.Complete(false, $"Export failed: {exportResult.ErrorMessage}");
            }

            return exportResult;
        }
        catch (Exception ex)
        {
            _fileLogger.LogException($"{label} stage export failed", ex);
            AppendLog($"[{label}] Failed: {ex.Message}");
            stage.Complete(false, ex.Message);
            return null;
        }
    }

    private (XLWorkbook? Workbook, string? Error) TryOpenWorkbook(string excelPath)
    {
        try
        {
            return (_excelReaderService.OpenWorkbook(excelPath), null);
        }
        catch (Exception ex)
        {
            _fileLogger.LogException("Failed to open Stock Received workbook", ex);
            return (null, ex.Message);
        }
    }

    private void AppendLog(string message) => LogLines.Add($"{DateTime.Now:HH:mm:ss} {message}");
}
```

- [ ] **Step 2: Build to verify it compiles**

Every type `StockReceivedViewModel` depends on (`ExcelReaderService`, `DbfExportService`,
`DbfReaderService`, `DbfSafetyBackupService`, `SettingsService`, `RunHistoryService`, `FileLogger`,
`ExcelSourceStagingService`, `DbfWorkflowHelper`, `IcmasteSchema`/`IctraneSchema`) already exists
from earlier tasks, and nothing yet constructs it — so this file has no unmet dependency and
should compile cleanly on its own, even before Task 9 wires it into the composition root.

Run: `dotnet build "src/eneBridge.Wpf/eneBridge.Wpf.csproj"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`. If it fails, the error will be a real typo or
wrong member name in this new file — fix it now rather than carrying it into Task 9.

- [ ] **Step 3: Commit**

```bash
git add src/eneBridge.Wpf/ViewModels/StockReceivedViewModel.cs
git commit -m "Add StockReceivedViewModel

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 9: Wire `StockReceivedViewModel` into `MainViewModel` and `App.xaml.cs`

`MainViewModel` gains a `StockReceivedViewModel` property so `MainWindow.xaml` (still rooted at
`MainViewModel` as its single `DataContext`) can reach it declaratively — no code-behind changes
needed. The one shared ACE-driver health check propagates to both.

**Files:**
- Modify: `src/eneBridge.Wpf/ViewModels/MainViewModel.cs`
- Modify: `src/eneBridge.Wpf/App.xaml.cs`

- [ ] **Step 1: Add the field, constructor parameter, and public property to `MainViewModel`**

Replace:

```csharp
    private readonly AceEngineGuardService _aceEngineGuardService;
    private readonly string _appBaseDirectory;
```

with:

```csharp
    private readonly AceEngineGuardService _aceEngineGuardService;
    private readonly string _appBaseDirectory;

    public StockReceivedViewModel StockReceivedViewModel { get; }
```

Replace the constructor:

```csharp
    public MainViewModel(
        ExcelReaderService excelReaderService,
        DbfExportService dbfExportService,
        DbfReaderService dbfReaderService,
        DbfSafetyBackupService dbfSafetyBackupService,
        SettingsService settingsService,
        RunHistoryService runHistoryService,
        FileLogger fileLogger,
        ExcelSourceStagingService excelSourceStagingService,
        AceEngineGuardService aceEngineGuardService,
        string appBaseDirectory)
    {
        _excelReaderService = excelReaderService;
        _dbfExportService = dbfExportService;
        _dbfReaderService = dbfReaderService;
        _dbfSafetyBackupService = dbfSafetyBackupService;
        _settingsService = settingsService;
        _runHistoryService = runHistoryService;
        _fileLogger = fileLogger;
        _excelSourceStagingService = excelSourceStagingService;
        _aceEngineGuardService = aceEngineGuardService;
        _appBaseDirectory = appBaseDirectory;
    }
```

with:

```csharp
    public MainViewModel(
        ExcelReaderService excelReaderService,
        DbfExportService dbfExportService,
        DbfReaderService dbfReaderService,
        DbfSafetyBackupService dbfSafetyBackupService,
        SettingsService settingsService,
        RunHistoryService runHistoryService,
        FileLogger fileLogger,
        ExcelSourceStagingService excelSourceStagingService,
        AceEngineGuardService aceEngineGuardService,
        StockReceivedViewModel stockReceivedViewModel,
        string appBaseDirectory)
    {
        _excelReaderService = excelReaderService;
        _dbfExportService = dbfExportService;
        _dbfReaderService = dbfReaderService;
        _dbfSafetyBackupService = dbfSafetyBackupService;
        _settingsService = settingsService;
        _runHistoryService = runHistoryService;
        _fileLogger = fileLogger;
        _excelSourceStagingService = excelSourceStagingService;
        _aceEngineGuardService = aceEngineGuardService;
        StockReceivedViewModel = stockReceivedViewModel;
        _appBaseDirectory = appBaseDirectory;
    }
```

- [ ] **Step 2: Propagate the shared health check to `StockReceivedViewModel`**

Replace:

```csharp
    private void SetAceEngineHealth(bool healthy)
    {
        AceEngineHealthy = healthy;
        AceEngineStatusText = healthy
            ? string.Empty
            : "Access Database Engine isn't working correctly — Confirm & Export is disabled until this is fixed.";
        ConfirmExportCommand.NotifyCanExecuteChanged();
    }
```

with:

```csharp
    private void SetAceEngineHealth(bool healthy)
    {
        AceEngineHealthy = healthy;
        AceEngineStatusText = healthy
            ? string.Empty
            : "Access Database Engine isn't working correctly — Confirm & Export is disabled until this is fixed.";
        ConfirmExportCommand.NotifyCanExecuteChanged();
        StockReceivedViewModel.SetAceEngineHealth(healthy);
    }
```

- [ ] **Step 3: Construct `StockReceivedViewModel` in `App.xaml.cs` and pass it to `MainViewModel`**

Replace:

```csharp
        var fileLogger = new FileLogger();
        RegisterGlobalExceptionHandlers(fileLogger);

        var excelReaderService = new ExcelReaderService();
        var dbfExportService = new DbfExportService(fileLogger);
        var dbfReaderService = new DbfReaderService(fileLogger);
        var dbfSafetyBackupService = new DbfSafetyBackupService(fileLogger);
        var settingsService = new SettingsService(AppContext.BaseDirectory);
        var runHistoryService = new RunHistoryService();
        var excelSourceStagingService = new ExcelSourceStagingService();
        var aceEngineGuardService = new AceEngineGuardService(fileLogger, AppContext.BaseDirectory);

        var mainViewModel = new MainViewModel(
            excelReaderService,
            dbfExportService,
            dbfReaderService,
            dbfSafetyBackupService,
            settingsService,
            runHistoryService,
            fileLogger,
            excelSourceStagingService,
            aceEngineGuardService,
            AppContext.BaseDirectory);
        mainViewModel.Initialize();
```

with:

```csharp
        var fileLogger = new FileLogger();
        RegisterGlobalExceptionHandlers(fileLogger);

        var excelReaderService = new ExcelReaderService();
        var dbfExportService = new DbfExportService(fileLogger);
        var dbfReaderService = new DbfReaderService(fileLogger);
        var dbfSafetyBackupService = new DbfSafetyBackupService(fileLogger);
        var settingsService = new SettingsService(AppContext.BaseDirectory);
        var runHistoryService = new RunHistoryService();
        var stockReceivedRunHistoryService = new RunHistoryService(fileName: "stockReceivedRunHistory.json");
        var excelSourceStagingService = new ExcelSourceStagingService();
        var aceEngineGuardService = new AceEngineGuardService(fileLogger, AppContext.BaseDirectory);

        var stockReceivedViewModel = new StockReceivedViewModel(
            excelReaderService,
            dbfExportService,
            dbfReaderService,
            dbfSafetyBackupService,
            settingsService,
            stockReceivedRunHistoryService,
            fileLogger,
            excelSourceStagingService,
            AppContext.BaseDirectory);
        stockReceivedViewModel.Initialize();

        var mainViewModel = new MainViewModel(
            excelReaderService,
            dbfExportService,
            dbfReaderService,
            dbfSafetyBackupService,
            settingsService,
            runHistoryService,
            fileLogger,
            excelSourceStagingService,
            aceEngineGuardService,
            stockReceivedViewModel,
            AppContext.BaseDirectory);
        mainViewModel.Initialize();
```

(`App.xaml.cs` already has `using eneBridge.Wpf.ViewModels;`, so `StockReceivedViewModel` resolves
without a new `using`.)

- [ ] **Step 4: Build to verify it compiles**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 5: Commit**

```bash
git add src/eneBridge.Wpf/ViewModels/MainViewModel.cs src/eneBridge.Wpf/App.xaml.cs
git commit -m "Wire StockReceivedViewModel into the composition root

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 10: `MainWindow.xaml` — real Stock Received tab content

Replaces today's placeholder with content mirroring the Invoice tab, bound to the nested
`StockReceivedViewModel` via `DataContext="{Binding StockReceivedViewModel}"` on the tab's root
element (the outer `TabControl`/`Window` stays rooted at `MainViewModel` — see Task 9).

**Files:**
- Modify: `src/eneBridge.Wpf/MainWindow.xaml`

- [ ] **Step 1: Replace the placeholder `TabItem`**

Replace:

```xml
            <TabItem Header="Stock Received">
                <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                    <TextBlock Text="Stock Received" FontSize="18" FontWeight="Bold" HorizontalAlignment="Center"/>
                </StackPanel>
            </TabItem>
```

with:

```xml
            <TabItem Header="Stock Received">
                <Grid Margin="0,8,0,0" DataContext="{Binding StockReceivedViewModel}">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="2*"/>
                        <RowDefinition Height="1*"/>
                    </Grid.RowDefinitions>

                    <!-- Path panel -->
                    <GroupBox Grid.Row="0" Header="File Paths" Margin="0,0,0,8">
                        <Grid Margin="8">
                            <Grid.RowDefinitions>
                                <RowDefinition/>
                                <RowDefinition/>
                            </Grid.RowDefinitions>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="Auto"/>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>

                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Excel file:" VerticalAlignment="Center" Margin="0,0,8,4"/>
                            <TextBox Grid.Row="0" Grid.Column="1" Text="{Binding ExcelFilePath, UpdateSourceTrigger=PropertyChanged}" Margin="0,0,8,4" VerticalContentAlignment="Center"/>
                            <Button Grid.Row="0" Grid.Column="2" Content="Browse..." Command="{Binding BrowseExcelCommand}" Padding="10,2" Margin="0,0,0,4"/>

                            <TextBlock Grid.Row="1" Grid.Column="0" Text="DBF folder:" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBox Grid.Row="1" Grid.Column="1" Text="{Binding DbfFolderPath, UpdateSourceTrigger=PropertyChanged}" Margin="0,0,8,0" VerticalContentAlignment="Center"/>
                            <Button Grid.Row="1" Grid.Column="2" Content="Browse..." Command="{Binding BrowseDbfFolderCommand}" Padding="10,2"/>
                        </Grid>
                    </GroupBox>

                    <!-- Run panel -->
                    <StackPanel Grid.Row="1" Margin="0,0,0,8">
                        <StackPanel Orientation="Horizontal">
                            <Button Content="Preview" Command="{Binding PreviewCommand}" Padding="24,6" FontWeight="Bold" MinWidth="100" Margin="0,0,8,0"/>
                            <Button Content="Confirm &amp; Export" Command="{Binding ConfirmExportCommand}" Padding="24,6" FontWeight="Bold" MinWidth="140"/>
                        </StackPanel>
                        <TextBlock Text="{Binding AceEngineStatusText}" Foreground="DarkRed" TextWrapping="Wrap" Margin="0,4,0,0"/>
                    </StackPanel>

                    <!-- Progress panel -->
                    <GroupBox Grid.Row="2" Header="Progress" Margin="0,0,0,8">
                        <UniformGrid Columns="2" Margin="8">
                            <StackPanel Margin="0,0,8,0">
                                <TextBlock Text="{Binding IcmasteStage.Name}" FontWeight="Bold"/>
                                <ProgressBar IsIndeterminate="{Binding IcmasteStage.IsRunning}" Height="6" Margin="0,4"/>
                                <TextBlock Text="{Binding IcmasteStage.StatusText}" TextWrapping="Wrap"/>
                            </StackPanel>
                            <StackPanel>
                                <TextBlock Text="{Binding IctraneStage.Name}" FontWeight="Bold"/>
                                <ProgressBar IsIndeterminate="{Binding IctraneStage.IsRunning}" Height="6" Margin="0,4"/>
                                <TextBlock Text="{Binding IctraneStage.StatusText}" TextWrapping="Wrap"/>
                            </StackPanel>
                        </UniformGrid>
                    </GroupBox>

                    <!-- Preview panel -->
                    <GroupBox Grid.Row="3" Margin="0,0,0,8">
                        <GroupBox.Header>
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="Preview" VerticalAlignment="Center"/>
                                <CheckBox Content="Show all columns" IsChecked="{Binding ShowAllColumns}" Margin="16,0,0,0" VerticalAlignment="Center"/>
                            </StackPanel>
                        </GroupBox.Header>
                        <TabControl>
                            <TabItem Header="icmaste">
                                <DataGrid x:Name="StockReceivedIcmasteGrid" ItemsSource="{Binding IcmastePreview}" IsReadOnly="True" AutoGenerateColumns="True" FrozenColumnCount="1"
                                          AutoGeneratedColumns="StockReceivedIcmasteGrid_AutoGeneratedColumns"/>
                            </TabItem>
                            <TabItem Header="ictrane">
                                <DataGrid x:Name="StockReceivedIctraneGrid" ItemsSource="{Binding IctranePreview}" IsReadOnly="True" AutoGenerateColumns="True" FrozenColumnCount="1"
                                          AutoGeneratedColumns="StockReceivedIctraneGrid_AutoGeneratedColumns"/>
                            </TabItem>
                            <TabItem Header="icmaste (DBF)">
                                <DockPanel>
                                    <Border DockPanel.Dock="Top" Padding="6">
                                        <Border.Style>
                                            <Style TargetType="Border">
                                                <Setter Property="Background" Value="#FFF3CD"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding IcmasteDbfVerified}" Value="True">
                                                        <Setter Property="Background" Value="#D4EDDA"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </Border.Style>
                                        <TextBlock Text="{Binding IcmasteDbfStatusText}" TextWrapping="Wrap"/>
                                    </Border>
                                    <DataGrid x:Name="StockReceivedIcmasteDbfGrid" ItemsSource="{Binding IcmasteDbfPreview}" IsReadOnly="True" AutoGenerateColumns="True" FrozenColumnCount="1"
                                              AutoGeneratedColumns="StockReceivedIcmasteDbfGrid_AutoGeneratedColumns"/>
                                </DockPanel>
                            </TabItem>
                            <TabItem Header="ictrane (DBF)">
                                <DockPanel>
                                    <Border DockPanel.Dock="Top" Padding="6">
                                        <Border.Style>
                                            <Style TargetType="Border">
                                                <Setter Property="Background" Value="#FFF3CD"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding IctraneDbfVerified}" Value="True">
                                                        <Setter Property="Background" Value="#D4EDDA"/>
                                                    </DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </Border.Style>
                                        <TextBlock Text="{Binding IctraneDbfStatusText}" TextWrapping="Wrap"/>
                                    </Border>
                                    <DataGrid x:Name="StockReceivedIctraneDbfGrid" ItemsSource="{Binding IctraneDbfPreview}" IsReadOnly="True" AutoGenerateColumns="True" FrozenColumnCount="1"
                                              AutoGeneratedColumns="StockReceivedIctraneDbfGrid_AutoGeneratedColumns"/>
                                </DockPanel>
                            </TabItem>
                        </TabControl>
                    </GroupBox>

                    <!-- Run history / log panel -->
                    <GroupBox Grid.Row="4" Header="Run History &amp; Log">
                        <Grid Margin="4">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>

                            <DataGrid Grid.Column="0" ItemsSource="{Binding RunHistory}" IsReadOnly="True" AutoGenerateColumns="False" Margin="0,0,4,0">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Time" Binding="{Binding Timestamp, StringFormat='yyyy-MM-dd HH:mm:ss'}"/>
                                    <DataGridTextColumn Header="icmaste R/S/W" Binding="{Binding IcmasteRowsRead, StringFormat='{}{0}'}"/>
                                    <DataGridTextColumn Header="icmaste OK" Binding="{Binding IcmasteSuccess}"/>
                                    <DataGridTextColumn Header="ictrane R/S/W" Binding="{Binding IctraneRowsRead, StringFormat='{}{0}'}"/>
                                    <DataGridTextColumn Header="ictrane OK" Binding="{Binding IctraneSuccess}"/>
                                    <DataGridTextColumn Header="Error" Binding="{Binding ErrorSummary}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>

                            <ListBox Grid.Column="1" ItemsSource="{Binding LogLines}" Margin="4,0,0,0" FontFamily="Consolas" FontSize="11" ScrollViewer.VerticalScrollBarVisibility="Auto"/>
                        </Grid>
                    </GroupBox>
                </Grid>
            </TabItem>
```

- [ ] **Step 2: Build to verify it compiles**

Run: `dotnet build "src/eneBridge.Wpf/eneBridge.Wpf.csproj"`
Expected: build FAILS — the `AutoGeneratedColumns="StockReceived*_AutoGeneratedColumns"` XAML
event hookups reference code-behind handlers that don't exist yet (Task 11). Confirm the errors are
ONLY about these four missing handler methods, nothing else (no XAML syntax errors, no unresolved
bindings reported as build errors — WPF binding errors don't fail the build, only show at runtime,
so a clean compile here doesn't yet prove the bindings themselves are correct — that's confirmed in
Task 12's manual run).

- [ ] **Step 3: Commit**

```bash
git add src/eneBridge.Wpf/MainWindow.xaml
git commit -m "Add real Stock Received tab content to MainWindow.xaml

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 11: `MainWindow.xaml.cs` — column-visibility handlers for the Stock Received grids

The existing `UpdateColumnVisibility` reads the Window's own `DataContext` (`MainViewModel`) to
decide whether to show all columns — that's correct for Invoice's grids but wrong for Stock
Received's (whose `ShowAllColumns` lives on the nested `StockReceivedViewModel`). Rather than
change that method's contract (risking Invoice's already-working behavior), this task adds a
parallel set of populated-column lists, handlers, and a second `PropertyChanged` subscription.

**Files:**
- Modify: `src/eneBridge.Wpf/MainWindow.xaml.cs`

- [ ] **Step 1: Add the Stock Received populated-column lists**

Add these two `HashSet`s right after the existing `IctranePopulatedColumns`:

```csharp
    // Mirrors what ExcelReaderService.ReadStockReceivedIcmaste/ReadStockReceivedIctrane actually
    // populate (see StockReceivedExcelCol and the Stock Received design doc). Same purely-display
    // purpose as IcmastePopulatedColumns/IctranePopulatedColumns above.
    private static readonly HashSet<string> StockReceivedIcmastePopulatedColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "TYPE", "ENTRY", "REF", "DATE", "CODE", "NAME", "ACCNO", "USER", "POSTACCNO",
        "CUST2", "CURRCODE", "TAXCODE", "VENDNO", "N_AMT", "T_AMT"
    };

    private static readonly HashSet<string> StockReceivedIctranePopulatedColumns = new(StringComparer.OrdinalIgnoreCase)
    {
        "type", "ref", "item_no", "desc1", "desc2", "taxcode", "userid", "entry"
    };
```

- [ ] **Step 2: Subscribe to `StockReceivedViewModel.PropertyChanged` alongside the existing subscription**

Replace:

```csharp
    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is MainViewModel oldViewModel)
        {
            oldViewModel.PropertyChanged -= OnViewModelPropertyChanged;
        }
        if (e.NewValue is MainViewModel newViewModel)
        {
            newViewModel.PropertyChanged += OnViewModelPropertyChanged;
        }
    }
```

with:

```csharp
    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is MainViewModel oldViewModel)
        {
            oldViewModel.PropertyChanged -= OnViewModelPropertyChanged;
            oldViewModel.StockReceivedViewModel.PropertyChanged -= OnStockReceivedViewModelPropertyChanged;
        }
        if (e.NewValue is MainViewModel newViewModel)
        {
            newViewModel.PropertyChanged += OnViewModelPropertyChanged;
            newViewModel.StockReceivedViewModel.PropertyChanged += OnStockReceivedViewModelPropertyChanged;
        }
    }

    private void OnStockReceivedViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(StockReceivedViewModel.ShowAllColumns))
        {
            UpdateStockReceivedColumnVisibility(StockReceivedIcmasteGrid, StockReceivedIcmastePopulatedColumns);
            UpdateStockReceivedColumnVisibility(StockReceivedIctraneGrid, StockReceivedIctranePopulatedColumns);
            UpdateStockReceivedColumnVisibility(StockReceivedIcmasteDbfGrid, StockReceivedIcmastePopulatedColumns);
            UpdateStockReceivedColumnVisibility(StockReceivedIctraneDbfGrid, StockReceivedIctranePopulatedColumns);
        }
    }
```

- [ ] **Step 3: Add the four `AutoGeneratedColumns` handlers and the parallel visibility method**

Add these right after the existing `IctraneDbfGrid_AutoGeneratedColumns` method:

```csharp
    private void StockReceivedIcmasteGrid_AutoGeneratedColumns(object sender, EventArgs e) =>
        UpdateStockReceivedColumnVisibility(StockReceivedIcmasteGrid, StockReceivedIcmastePopulatedColumns);

    private void StockReceivedIctraneGrid_AutoGeneratedColumns(object sender, EventArgs e) =>
        UpdateStockReceivedColumnVisibility(StockReceivedIctraneGrid, StockReceivedIctranePopulatedColumns);

    private void StockReceivedIcmasteDbfGrid_AutoGeneratedColumns(object sender, EventArgs e) =>
        UpdateStockReceivedColumnVisibility(StockReceivedIcmasteDbfGrid, StockReceivedIcmastePopulatedColumns);

    private void StockReceivedIctraneDbfGrid_AutoGeneratedColumns(object sender, EventArgs e) =>
        UpdateStockReceivedColumnVisibility(StockReceivedIctraneDbfGrid, StockReceivedIctranePopulatedColumns);
```

Add this right after the existing `UpdateColumnVisibility` method (same class, same indentation):

```csharp
    private void UpdateStockReceivedColumnVisibility(DataGrid grid, HashSet<string> populatedColumns)
    {
        bool showAll = DataContext is MainViewModel viewModel && viewModel.StockReceivedViewModel.ShowAllColumns;
        foreach (var column in grid.Columns)
        {
            column.Visibility = showAll || populatedColumns.Contains(column.SortMemberPath)
                ? Visibility.Visible
                : Visibility.Collapsed;
        }
    }
```

- [ ] **Step 4: Build to verify it compiles**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 5: Commit**

```bash
git add src/eneBridge.Wpf/MainWindow.xaml.cs
git commit -m "Add column-visibility handlers for the Stock Received grids

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 12: `CLAUDE.md` update + full verification pass

**Files:**
- Modify: `CLAUDE.md`

- [ ] **Step 1: Add a "Stock Received workflow" paragraph under Architecture**

Find the paragraph beginning `**DBF export crash prevention**` and add a new paragraph immediately
after it (before `**Error handling**`):

```markdown
**Stock Received workflow**: a second, independent workflow (`StockReceivedViewModel`,
`MainWindow.xaml`'s "Stock Received" tab) reads `Self-Billed_Format-ORI.xlsx`-shaped workbooks
(`ExcelReaderService.ReadStockReceivedIcmaste`/`ReadStockReceivedIctrane`, driven by
`StockReceivedExcelCol`) and writes into the SAME `icmaste.dbf`/`ictrane.dbf` Invoice writes to,
distinguished by `TYPE` (`"RE"` vs Invoice's `"IN"`). See
`docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md` for the full field mapping
and rationale. This required changing `DbfExportService.Export` itself: it no longer backs up and
recreates the table on every run — it now creates the table only if missing and always appends,
for both workflows. `MainViewModel`'s and `StockReceivedViewModel`'s post-export verification
correspondingly compares a before/after row-count delta (filtered by `TYPE`) instead of an
absolute total, since rows now accumulate indefinitely rather than the file being replaced each
run — there is no "start fresh" export anymore; resetting the live tables is a manual operation
outside the app. The ACE-driver-crash-avoidance table check and this delta verification live in
the shared `Services/DbfWorkflowHelper.cs` (Wpf project) rather than being duplicated between the
two ViewModels, since they're safety-critical; everything else in `StockReceivedViewModel`
(read/export/preview plumbing) deliberately mirrors `MainViewModel`'s Invoice-related surface as a
separate, parallel implementation rather than a shared base class. `StockReceivedViewModel` has
its own Excel path (`SettingsService.ResolveEffectiveStockReceivedPaths`, staged to
`stockreceived.xlsx`) and its own run history file (`%AppData%\eneBridge\stockReceivedRunHistory.json`),
but shares Invoice's `DbfFolderPath` (the two must point at the same folder for append to mean
anything) and the single app-wide ACE-driver health check (`MainViewModel` propagates its result
to `StockReceivedViewModel` via `SetAceEngineHealth` whenever it changes, rather than running a
second self-test).
```

- [ ] **Step 2: Note the append behavior needs real-EMAS verification, in Current status**

Find the `## Current status` paragraph (the one beginning `Core services and the WPF UI are
implemented...`) and add one more sentence at its end:

Replace:

```markdown
The error-path testing (nonexistent Excel path, read-only DBF
folder, corrupted cell values) called for in the original verification plan is now covered in
`ExcelReaderServiceTests`, `DbfSchemaBuilderTests`, and `DbfExportServiceTests`.
```

with:

```markdown
The error-path testing (nonexistent Excel path, read-only DBF
folder, corrupted cell values) called for in the original verification plan is now covered in
`ExcelReaderServiceTests`, `DbfSchemaBuilderTests`, and `DbfExportServiceTests`. **Not yet
confirmed**: whether EMAS's Inventory Control module correctly handles `icmaste.dbf`/`ictrane.dbf`
accumulating rows indefinitely across both workflows' runs (see "Stock Received workflow" above)
— this needs to be watched the first few times both workflows are used against a real EMAS
installation.
```

- [ ] **Step 3: Commit the CLAUDE.md update**

```bash
git add CLAUDE.md
git commit -m "Document the Stock Received workflow in CLAUDE.md

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

- [ ] **Step 4: Full Core test suite, run by class (matches this project's established pattern of
avoiding the pre-existing full-suite ACE crash — see `CLAUDE.md`'s Commands section)**

Run each of these and confirm all pass:

```bash
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~ExcelReaderServiceTests"
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~DbfExportServiceTests"
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~DbfReaderServiceTests"
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~SettingsServiceTests"
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~RunHistoryServiceTests"
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~DbfSchemaBuilderTests"
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~SqlValueFormatterTests"
dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~ExcelSourceStagingServiceTests"
```

Expected: every one prints `Passed!` with 0 failures. (`ExcelReaderServiceTests` should now show
roughly 41 tests — the pre-existing ~28 plus the 13 new Stock Received ones from Task 2.)

- [ ] **Step 5: Full solution build**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 6: Manual end-to-end verification of the Stock Received tab itself, against a scratch folder**

```bash
mkdir -p /tmp/enebridge-manual-verify-stock-received
```

You'll need a small `.xlsx` file shaped like `Self-Billed_Format-ORI.xlsx` (columns: `No, Stock
Received No, Date, Supplier Code, Supplier Name, Item Code, Item Description, Qty, Unit Price,
Total Amount`, header row 1, data from row 2 — 2-3 sample rows is enough) to point the Stock
Received tab's Excel file field at. If `E:\Irene Projects\Self-Billed_Format-ORI.xlsx` (the file
this workflow was designed against) is still available, use it directly.

Run the app: `dotnet run --project "src/eneBridge.Wpf/eneBridge.Wpf.csproj"`

On the **Stock Received** tab:
1. Browse to the sample Excel file. Browse the DBF folder to the empty scratch folder.
2. Click **Preview**. Confirm both preview grids populate with the expected rows (matching the
   field mapping in the design doc), and the progress panel shows a sane read/skip count.
3. Click **Confirm & Export**. Confirm: the Table Check dialog appears (first-ever run, empty
   folder) and Continue proceeds; both DBF-verification tabs turn green with a "✓ ... confirmed"
   message; `icmaste.dbf`/`ictrane.dbf` now exist in the scratch folder.
4. Click **Confirm & Export** a second time (no changes needed). Confirm: no Table Check dialog
   this time (tables exist and are readable); the DBF-verification status text shows the newly
   written row count, not a doubled or reset total; the row count in the underlying `.dbf` files
   is now the sum of both runs (spot-check by switching to the **Invoice** tab, pointing it at the
   SAME scratch folder, and confirming its own verification reads back a table containing rows
   from both tabs' runs when unfiltered — the per-tab preview grids themselves only show each
   tab's own `TYPE`-filtered rows, by design).
5. Switch to the **Invoice** tab (same scratch folder). Run Preview + Confirm & Export there too
   (using the existing fixture workbook or any valid invoice-shaped source). Confirm Invoice's own
   rows land alongside Stock Received's in the same live files without disturbing them — check the
   Stock Received tab's own DBF-verification tabs still show their own prior rows correctly after
   switching back.
6. Confirm the "Show all columns" checkbox toggles visibility of the full 151/80-column schema
   independently on the Stock Received tab's four grids, without affecting the Invoice tab's grids.

- [ ] **Step 7: Confirm no production files were touched**

This task's manual verification only ever uses the scratch folder created in Step 6 — no
production EMAS folder is involved in this plan. Nothing further to check here.
