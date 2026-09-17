using System.IO.Compression;
using ClosedXML.Excel;
using eneBridge.Wpf.Core.Exceptions;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class ExcelReaderServiceTests
{
    private static string FixturePath => Path.Combine(AppContext.BaseDirectory, "Fixtures", "data.xlsx");

    [Fact]
    public void IcmasteSchema_HasExpectedColumnCount()
    {
        Assert.Equal(151, IcmasteSchema.Columns.Count);
    }

    [Fact]
    public void IctraneSchema_HasExpectedColumnCount()
    {
        Assert.Equal(80, IctraneSchema.Columns.Count);
    }

    [Fact]
    public void ReadIcmaste_DedupesByRef_OneRowPerDocument()
    {
        var service = new ExcelReaderService();
        using var workbook = service.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var result = service.ReadIcmaste(worksheet);

        Assert.Equal(110, result.RowsRead);
        Assert.Equal(3, result.RowsWritten);
        Assert.Equal(107, result.SkipReasons.Count);

        var refs = result.Table.Rows.Cast<System.Data.DataRow>()
            .Select(r => (string)r[IcmasteSchema.Ref])
            .OrderBy(r => r)
            .ToArray();
        Assert.Equal(["2501001", "2501002", "2501003"], refs);
    }

    [Fact]
    public void ReadIctrane_KeepsEveryLineItem_NoRefDedup()
    {
        var service = new ExcelReaderService();
        using var workbook = service.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var result = service.ReadIctrane(worksheet);

        Assert.Equal(110, result.RowsRead);
        Assert.Equal(110, result.RowsWritten);
        Assert.Empty(result.SkipReasons);
    }

    [Fact]
    public void ReadIctrane_SpotCheck_FirstProcessedRowValues()
    {
        var service = new ExcelReaderService();
        using var workbook = service.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var result = service.ReadIctrane(worksheet);
        // Physical row 2 (right after the header row) is the first processed row.
        var row = result.Table.Rows[0];

        Assert.Equal(0.629m, row[IctraneSchema.Qty]);
        Assert.Equal(900m, row[IctraneSchema.Price]);
        Assert.Equal(566.10m, row[IctraneSchema.Amount]);
        Assert.Equal("2501001", row[IctraneSchema.Ref]);
        Assert.Equal("SST0", row[IctraneSchema.TaxCode]);
    }

    [Fact]
    public void ReadIctrane_SpotCheck_SixthProcessedRowValues()
    {
        var service = new ExcelReaderService();
        using var workbook = service.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var result = service.ReadIctrane(worksheet);
        // Physical row 7 is the 6th processed row (rows 2,3,4,5,6,7) -> index 5.
        var row = result.Table.Rows[5];

        Assert.Equal(0.362m, row[IctraneSchema.Qty]);
        Assert.Equal(850m, row[IctraneSchema.Price]);
        Assert.Equal(307.7m, row[IctraneSchema.Amount]);
        Assert.Equal("2501001", row[IctraneSchema.Ref]);
    }

    [Fact]
    public void ReadIcmaste_SpotCheck_KnownDocumentValues()
    {
        var service = new ExcelReaderService();
        using var workbook = service.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var result = service.ReadIcmaste(worksheet);
        var row = result.Table.Rows.Cast<System.Data.DataRow>()
            .Single(r => (string)r[IcmasteSchema.Ref] == "2501001");

        Assert.Equal("3000L001", row[IcmasteSchema.Code]);
        Assert.Equal("LAGENDA SP TIMBER SDN BHD", row[IcmasteSchema.Name]);
        Assert.Equal("SST0", row[IcmasteSchema.TaxCode]);
        Assert.Equal(0m, row[IcmasteSchema.FcRate]);
        Assert.Equal("MYR", row[IcmasteSchema.CurrCode]);
        Assert.Equal("IN", row[IcmasteSchema.Type]);
    }

    [Fact]
    public void OpenWorkbook_FileDoesNotExist_ThrowsExcelReadException()
    {
        var service = new ExcelReaderService();
        var missingPath = Path.Combine(Path.GetTempPath(), "eneBridge-missing-" + Guid.NewGuid().ToString("N") + ".xlsx");

        var ex = Assert.Throws<ExcelReadException>(() => service.OpenWorkbook(missingPath));

        Assert.Contains("Excel file not found:", ex.Message);
        Assert.Contains(missingPath, ex.Message);
    }

    [Fact]
    public void OpenWorkbook_FileIsNotValidExcel_ThrowsExcelReadExceptionWithInnerException()
    {
        var service = new ExcelReaderService();
        var corruptPath = Path.Combine(Path.GetTempPath(), "eneBridge-corrupt-" + Guid.NewGuid().ToString("N") + ".xlsx");

        // A well-formed but empty zip archive: this is not a valid OOXML package (no xl/workbook.xml
        // part), so ClosedXML/OpenXml fails in managed code while parsing package parts. Deliberately
        // NOT using raw garbage bytes here — a too-small/malformed file can be misdetected as a legacy
        // OLE2 compound-file format and crash the process with a native AccessViolationException during
        // COM finalization (reproduced while writing this test).
        using (var stream = new FileStream(corruptPath, FileMode.Create))
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create))
        {
            var entry = archive.CreateEntry("not-a-workbook.txt");
            using var entryStream = entry.Open();
            using var writer = new StreamWriter(entryStream);
            writer.Write("not a real workbook");
        }

        try
        {
            var ex = Assert.Throws<ExcelReadException>(() => service.OpenWorkbook(corruptPath));

            Assert.NotNull(ex.InnerException);
            Assert.StartsWith("Failed to open Excel file", ex.Message);
            Assert.Contains(corruptPath, ex.Message);
        }
        finally
        {
            try { File.Delete(corruptPath); } catch (IOException) { }
        }
    }

    [Fact]
    public void OpenWorkbook_FileIsLockedByAnotherProcess_ThrowsExcelReadExceptionWithInnerException()
    {
        var service = new ExcelReaderService();
        var lockedPath = Path.Combine(Path.GetTempPath(), "eneBridge-locked-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(FixturePath, lockedPath);

        var lockStream = new FileStream(lockedPath, FileMode.Open, FileAccess.Read, FileShare.None);
        try
        {
            var ex = Assert.Throws<ExcelReadException>(() => service.OpenWorkbook(lockedPath));

            Assert.NotNull(ex.InnerException);
            Assert.Equal(
                "The Excel file is currently open in another program. Please close it and click Run again.",
                ex.Message);
        }
        finally
        {
            lockStream.Dispose();
            try { File.Delete(lockedPath); } catch (IOException) { }
        }
    }

    [Fact]
    public void OpenWorkbook_FileIsOpenInExcel_SucceedsAnyway()
    {
        var service = new ExcelReaderService();
        var sharedPath = Path.Combine(Path.GetTempPath(), "eneBridge-excelopen-" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(FixturePath, sharedPath);

        // Simulates how Excel holds a file open for editing: ReadWrite access, Read share.
        var excelLikeStream = new FileStream(sharedPath, FileMode.Open, FileAccess.ReadWrite, FileShare.Read);
        try
        {
            using var workbook = service.OpenWorkbook(sharedPath);

            Assert.NotNull(workbook);
            Assert.NotEmpty(workbook.Worksheets);
        }
        finally
        {
            excelLikeStream.Dispose();
            try { File.Delete(sharedPath); } catch (IOException) { }
        }
    }

    [Fact]
    public void ReadIcmaste_WorksheetHasTooFewColumns_ThrowsExcelValidationException()
    {
        var service = new ExcelReaderService();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");
        worksheet.Cell(1, 5).SetValue("x");

        var ex = Assert.Throws<ExcelValidationException>(() => service.ReadIcmaste(worksheet));

        Assert.Equal(
            "The Excel file does not contain the required columns (found 5, need at least 8).",
            ex.Message);
    }

    [Fact]
    public void ReadIctrane_WorksheetHasTooFewColumns_ThrowsExcelValidationException()
    {
        var service = new ExcelReaderService();
        using var workbook = new XLWorkbook();
        var worksheet = workbook.AddWorksheet("Sheet1");
        worksheet.Cell(1, 5).SetValue("x");

        var ex = Assert.Throws<ExcelValidationException>(() => service.ReadIctrane(worksheet));

        Assert.Equal(
            "The Excel file does not contain the required columns (found 5, need at least 9).",
            ex.Message);
    }

    [Theory]
    [InlineData("ref", "REF is blank")]
    [InlineData("date", "DATE is blank")]
    [InlineData("code", "CODE is blank")]
    [InlineData("name", "NAME is blank")]
    public void ReadIcmaste_BlankRequiredField_SkipsRowWithReason(string blankField, string expectedReason)
    {
        var service = new ExcelReaderService();
        var worksheet = BuildIcmasteWorksheet(
            refValue: blankField == "ref" ? null : "REF1",
            dateValue: blankField == "date" ? null : "2025-01-01",
            codeValue: blankField == "code" ? null : "CODE1",
            nameValue: blankField == "name" ? null : "NAME1");

        var result = service.ReadIcmaste(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal(expectedReason, result.SkipReasons[0].Reason);
        Assert.Empty(result.Table.Rows);
    }

    [Fact]
    public void ReadIcmaste_DuplicateRef_SkipsSecondRowWithDuplicateReason()
    {
        var service = new ExcelReaderService();
        var worksheet = NewWorksheet();
        SetIcmasteRow(worksheet, ExcelLayout.FirstDataRow, "REF1", "2025-01-01", "CODE1", "NAME1");
        SetIcmasteRow(worksheet, ExcelLayout.FirstDataRow + 1, "REF1", "2025-01-02", "CODE2", "NAME2");

        var result = service.ReadIcmaste(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal(ExcelLayout.FirstDataRow + 1, result.SkipReasons[0].ExcelRow);
        Assert.Contains("Duplicate REF", result.SkipReasons[0].Reason);
        Assert.Single(result.Table.Rows);
    }

    [Fact]
    public void ReadIcmaste_DateCellIsNotParsable_SkipsRowWithDateParseFailedReason()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildIcmasteWorksheet(dateValue: "not-a-date");

        var result = service.ReadIcmaste(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.StartsWith("DATE parse failed:", result.SkipReasons[0].Reason);
    }

    [Fact]
    public void ReadIcmaste_FcRateIsNotNumeric_SkipsRowWithFcRateReason()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildIcmasteWorksheet(fcRateValue: "abc");

        var result = service.ReadIcmaste(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal("FCRATE value 'abc' is not a valid number", result.SkipReasons[0].Reason);
    }

    [Fact]
    public void ReadIcmaste_FcRateAndTaxCodeBlank_DefaultsApplied()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildIcmasteWorksheet(fcRateValue: null, taxCodeValue: null);

        var result = service.ReadIcmaste(worksheet);

        Assert.Single(result.Table.Rows);
        var row = result.Table.Rows[0];
        Assert.Equal(0m, row[IcmasteSchema.FcRate]);
        Assert.Equal("SST0", row[IcmasteSchema.TaxCode]);
    }

    [Theory]
    [InlineData("ref", "ref is blank")]
    [InlineData("itemNo", "item_no is blank")]
    [InlineData("amounts", "qty, price and amount are all blank")]
    public void ReadIctrane_BlankRequiredField_SkipsRowWithReason(string blankField, string expectedReason)
    {
        var service = new ExcelReaderService();
        var worksheet = BuildIctraneWorksheet(
            refValue: blankField == "ref" ? null : "REF1",
            itemNoValue: blankField == "itemNo" ? null : "ITEM1",
            qtyValue: blankField == "amounts" ? null : "1",
            priceValue: blankField == "amounts" ? null : "10",
            amountValue: blankField == "amounts" ? null : "10");

        var result = service.ReadIctrane(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal(expectedReason, result.SkipReasons[0].Reason);
    }

    [Theory]
    [InlineData("qty")]
    [InlineData("price")]
    [InlineData("amount")]
    public void ReadIctrane_NonNumericAmount_SkipsRowWithReason(string field)
    {
        var service = new ExcelReaderService();
        var worksheet = BuildIctraneWorksheet(
            qtyValue: field == "qty" ? "abc" : "1",
            priceValue: field == "price" ? "abc" : "10",
            amountValue: field == "amount" ? "abc" : "10");

        var result = service.ReadIctrane(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal($"{field} value 'abc' is not a valid number", result.SkipReasons[0].Reason);
    }

    [Fact]
    public void ReadIctrane_TaxCodeBlank_DefaultsToSst0()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildIctraneWorksheet(taxCodeValue: null);

        var result = service.ReadIctrane(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal("SST0", result.Table.Rows[0][IctraneSchema.TaxCode]);
    }

    private static IXLWorksheet NewWorksheet() => new XLWorkbook().AddWorksheet("Sheet1");

    private static void SetCell(IXLWorksheet worksheet, int row, int zeroBasedColumn, string? value)
    {
        if (value is not null)
        {
            worksheet.Cell(row, zeroBasedColumn + 1).SetValue(value);
        }
    }

    private static void SetIcmasteRow(
        IXLWorksheet worksheet, int row, string? refValue, string? dateValue, string? codeValue, string? nameValue,
        string? fcRateValue = null, string? taxCodeValue = "SST0")
    {
        SetCell(worksheet, row, IcmasteExcelCol.Date, dateValue);
        SetCell(worksheet, row, IcmasteExcelCol.Code, codeValue);
        SetCell(worksheet, row, IcmasteExcelCol.Name, nameValue);
        SetCell(worksheet, row, IcmasteExcelCol.Ref, refValue);
        SetCell(worksheet, row, IcmasteExcelCol.FcRate, fcRateValue);
        SetCell(worksheet, row, IcmasteExcelCol.TaxCode, taxCodeValue);
    }

    private static IXLWorksheet BuildIcmasteWorksheet(
        string? refValue = "REF1", string? dateValue = "2025-01-01", string? codeValue = "CODE1", string? nameValue = "NAME1",
        string? fcRateValue = null, string? taxCodeValue = "SST0")
    {
        var worksheet = NewWorksheet();
        SetIcmasteRow(worksheet, ExcelLayout.FirstDataRow, refValue, dateValue, codeValue, nameValue, fcRateValue, taxCodeValue);
        return worksheet;
    }

    private static void SetIctraneRow(
        IXLWorksheet worksheet, int row, string? refValue, string? itemNoValue, string? qtyValue, string? priceValue,
        string? amountValue, string? taxCodeValue = "SST0")
    {
        SetCell(worksheet, row, IctraneExcelCol.ItemNoAndDesc1, itemNoValue);
        SetCell(worksheet, row, IctraneExcelCol.Qty, qtyValue);
        SetCell(worksheet, row, IctraneExcelCol.Price, priceValue);
        SetCell(worksheet, row, IctraneExcelCol.Amount, amountValue);
        SetCell(worksheet, row, IctraneExcelCol.Ref, refValue);
        SetCell(worksheet, row, IctraneExcelCol.TaxCode, taxCodeValue);
    }

    private static IXLWorksheet BuildIctraneWorksheet(
        string? refValue = "REF1", string? itemNoValue = "ITEM1", string? qtyValue = "1", string? priceValue = "10",
        string? amountValue = "10", string? taxCodeValue = "SST0")
    {
        var worksheet = NewWorksheet();
        SetIctraneRow(worksheet, ExcelLayout.FirstDataRow, refValue, itemNoValue, qtyValue, priceValue, amountValue, taxCodeValue);
        return worksheet;
    }

    private static void SetStockReceivedRow(
        IXLWorksheet worksheet, int row, string? refValue, string? dateValue, string? codeValue, string? nameValue,
        string? itemNoValue, string? desc1Value = "ITEM DESC", string? qtyValue = "1", string? nAmtValue = "100",
        string? tAmtValue = "100", string? desc2Value = null)
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
    }

    private static IXLWorksheet BuildStockReceivedWorksheet(
        string? refValue = "SB-000001", string? dateValue = "2026-09-01", string? codeValue = "4000T001",
        string? nameValue = "TAN KANG KAE", string? itemNoValue = "sub-con", string? desc1Value = "Sub Con Wages",
        string? qtyValue = "1", string? nAmtValue = "1000", string? tAmtValue = "1000", string? desc2Value = null)
    {
        var worksheet = NewWorksheet();
        SetStockReceivedRow(worksheet, ExcelLayout.FirstDataRow, refValue, dateValue, codeValue, nameValue, itemNoValue, desc1Value, qtyValue, nAmtValue, tAmtValue, desc2Value);
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
        Assert.Equal("SST0", row[IctraneSchema.TaxCode]);
        Assert.Equal("admin", row[IctraneSchema.UserId]);
        Assert.Equal(1m, row[IctraneSchema.Qty]);
        Assert.Equal(1000m, row[IctraneSchema.Price]);
        Assert.Equal(1000m, row[IctraneSchema.Amount]);
    }

    [Fact]
    public void ReadStockReceivedIctrane_Desc2Populated_MapsToDesc2()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(desc2Value: "PROJECT A");

        var result = service.ReadStockReceivedIctrane(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal("PROJECT A", result.Table.Rows[0][IctraneSchema.Desc2]);
    }

    [Fact]
    public void ReadStockReceivedIctrane_Desc2Blank_MapsToEmptyString()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet();

        var result = service.ReadStockReceivedIctrane(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal(string.Empty, result.Table.Rows[0][IctraneSchema.Desc2]);
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
    public void ReadStockReceivedIctrane_QtyNotNumeric_SkipsRowWithReason()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(qtyValue: "abc");

        var result = service.ReadStockReceivedIctrane(worksheet);

        Assert.Single(result.SkipReasons);
        Assert.Equal("Qty value 'abc' is not a valid number", result.SkipReasons[0].Reason);
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
        var worksheet = BuildStockReceivedWorksheet(nAmtValue: " ", tAmtValue: " ");

        var result = service.ReadStockReceivedIcmaste(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal(0m, result.Table.Rows[0][IcmasteSchema.NAmt]);
        Assert.Equal(0m, result.Table.Rows[0][IcmasteSchema.TAmt]);
    }

    [Fact]
    public void ReadStockReceivedIctrane_QtyNAmtAndTAmtBlank_DefaultToZero()
    {
        var service = new ExcelReaderService();
        var worksheet = BuildStockReceivedWorksheet(qtyValue: " ", nAmtValue: " ", tAmtValue: " ");

        var result = service.ReadStockReceivedIctrane(worksheet);

        Assert.Single(result.Table.Rows);
        Assert.Equal(0m, result.Table.Rows[0][IctraneSchema.Qty]);
        Assert.Equal(0m, result.Table.Rows[0][IctraneSchema.Price]);
        Assert.Equal(0m, result.Table.Rows[0][IctraneSchema.Amount]);
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
            "The Excel file does not contain the required columns (found 5, need at least 11).",
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
            "The Excel file does not contain the required columns (found 5, need at least 11).",
            ex.Message);
    }

    [Fact]
    public void ValidateFileFormat_InvoiceExpected_StockReceivedHeaderPresent_Throws()
    {
        var service = new ExcelReaderService();
        var worksheet = NewWorksheet();
        SetCell(worksheet, 1, 1, "Stock Received No");

        var ex = Assert.Throws<ExcelValidationException>(
            () => service.ValidateFileFormat(worksheet, ExcelFileFormat.Invoice));

        Assert.Equal(
            "This file looks like a Stock Received file (found a \"Stock Received No\" column), " +
            "not an Invoice file. Please use the Stock Received tab, or browse to the correct Invoice file.",
            ex.Message);
    }

    [Fact]
    public void ValidateFileFormat_InvoiceExpected_InvoiceHeaderPresent_DoesNotThrow()
    {
        var service = new ExcelReaderService();
        var worksheet = NewWorksheet();
        SetCell(worksheet, 1, 8, "Invoice Number");

        service.ValidateFileFormat(worksheet, ExcelFileFormat.Invoice);
    }

    [Fact]
    public void ValidateFileFormat_StockReceivedExpected_InvoiceHeaderPresent_Throws()
    {
        var service = new ExcelReaderService();
        var worksheet = NewWorksheet();
        SetCell(worksheet, 1, 8, "Invoice Number");

        var ex = Assert.Throws<ExcelValidationException>(
            () => service.ValidateFileFormat(worksheet, ExcelFileFormat.StockReceived));

        Assert.Equal(
            "This file looks like an Invoice file (found an \"Invoice Number\" column), " +
            "not a Stock Received file. Please use the Invoice tab, or browse to the correct Stock Received file.",
            ex.Message);
    }

    [Fact]
    public void ValidateFileFormat_StockReceivedExpected_StockReceivedHeaderPresent_DoesNotThrow()
    {
        var service = new ExcelReaderService();
        var worksheet = NewWorksheet();
        SetCell(worksheet, 1, 1, "Stock Received No");

        service.ValidateFileFormat(worksheet, ExcelFileFormat.StockReceived);
    }

    [Fact]
    public void ValidateFileFormat_NeitherHeaderPresent_DoesNotThrow()
    {
        var service = new ExcelReaderService();
        var worksheet = NewWorksheet();
        SetCell(worksheet, 1, 0, "No");
        SetCell(worksheet, 1, 1, "Item Description");

        service.ValidateFileFormat(worksheet, ExcelFileFormat.Invoice);
        service.ValidateFileFormat(worksheet, ExcelFileFormat.StockReceived);
    }
}
