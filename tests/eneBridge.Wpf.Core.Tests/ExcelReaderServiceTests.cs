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

        Assert.Equal(105, result.RowsRead);
        Assert.Equal(3, result.RowsWritten);
        Assert.Equal(102, result.SkipReasons.Count);

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

        Assert.Equal(105, result.RowsRead);
        Assert.Equal(105, result.RowsWritten);
        Assert.Empty(result.SkipReasons);
    }

    [Fact]
    public void ReadIctrane_SpotCheck_RowSixValues()
    {
        var service = new ExcelReaderService();
        using var workbook = service.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var result = service.ReadIctrane(worksheet);
        // Physical row 6 is the first processed row -> first row in the table.
        var row = result.Table.Rows[0];

        Assert.Equal(0.362m, row[IctraneSchema.Qty]);
        Assert.Equal(850m, row[IctraneSchema.Price]);
        Assert.Equal(307.7m, row[IctraneSchema.Amount]);
        Assert.Equal("2501001", row[IctraneSchema.Ref]);
    }

    [Fact]
    public void ReadIctrane_SpotCheck_RowElevenValues()
    {
        var service = new ExcelReaderService();
        using var workbook = service.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var result = service.ReadIctrane(worksheet);
        // Physical row 11 is the 6th processed row (rows 6,7,8,9,10,11) -> index 5.
        var row = result.Table.Rows[5];

        Assert.Equal(0.912m, row[IctraneSchema.Qty]);
        Assert.Equal(1030m, row[IctraneSchema.Price]);
        Assert.Equal(939.36m, row[IctraneSchema.Amount]);
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
            Assert.StartsWith("Failed to open Excel file", ex.Message);
        }
        finally
        {
            lockStream.Dispose();
            try { File.Delete(lockedPath); } catch (IOException) { }
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
}
