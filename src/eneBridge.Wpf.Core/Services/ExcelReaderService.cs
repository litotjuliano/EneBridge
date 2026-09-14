using ClosedXML.Excel;
using eneBridge.Wpf.Core.Exceptions;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Reads the source Excel workbook and builds the icmaste/ictrane DataTables, using ClosedXML
/// (a pure managed library) instead of the original's OleDb/Access-Database-Engine approach.
/// Reading explicit physical row/column positions removes the original's implicit "row 1 is a
/// header" ambiguity — see <see cref="ExcelLayout"/>.
/// </summary>
public sealed class ExcelReaderService
{
    public XLWorkbook OpenWorkbook(string path)
    {
        if (!File.Exists(path))
        {
            throw new ExcelReadException($"Excel file not found: {path}");
        }

        try
        {
            using var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            var memoryStream = new MemoryStream();
            fileStream.CopyTo(memoryStream);
            memoryStream.Position = 0;
            return new XLWorkbook(memoryStream);
        }
        catch (IOException ex) when (ex.HResult == unchecked((int)0x80070020))
        {
            throw new ExcelReadException(
                "The Excel file is currently open in another program. Please close it and click Run again.",
                ex);
        }
        catch (Exception ex) when (ex is not ExcelReadException)
        {
            throw new ExcelReadException($"Failed to open Excel file '{path}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Reads icmaste (inventory master, one row per document). Rows with a blank REF/DATE/CODE/NAME
    /// are skipped; rows whose REF repeats an already-seen document are skipped as duplicates
    /// (correct here — REF is a document number and icmaste is a one-row-per-document table).
    /// </summary>
    public StageReadResult ReadIcmaste(IXLWorksheet worksheet)
    {
        ValidateColumnCount(worksheet, IcmasteExcelCol.MinColumnCount);

        var table = IcmasteSchema.BuildEmptyTable();
        var skipReasons = new List<RowSkipReason>();
        var seenRefs = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int rowsRead = 0;

        for (int excelRow = ExcelLayout.FirstDataRow; excelRow <= lastRow; excelRow++)
        {
            rowsRead++;

            string refValue = GetString(worksheet, excelRow, IcmasteExcelCol.Ref);
            string dateRaw = GetString(worksheet, excelRow, IcmasteExcelCol.Date);
            string codeValue = GetString(worksheet, excelRow, IcmasteExcelCol.Code);
            string nameValue = GetString(worksheet, excelRow, IcmasteExcelCol.Name);

            if (IsBlank(refValue)) { skipReasons.Add(new RowSkipReason(excelRow, "REF is blank")); continue; }
            if (IsBlank(dateRaw)) { skipReasons.Add(new RowSkipReason(excelRow, "DATE is blank")); continue; }
            if (IsBlank(codeValue)) { skipReasons.Add(new RowSkipReason(excelRow, "CODE is blank")); continue; }
            if (IsBlank(nameValue)) { skipReasons.Add(new RowSkipReason(excelRow, "NAME is blank")); continue; }

            if (!seenRefs.Add(refValue))
            {
                skipReasons.Add(new RowSkipReason(excelRow, $"Duplicate REF '{refValue}' (icmaste keeps only the first row per document; this REF already appeared earlier in this file)"));
                continue;
            }

            DateTime date;
            try
            {
                date = ParseDate(worksheet, excelRow, IcmasteExcelCol.Date, dateRaw);
            }
            catch (Exception ex)
            {
                skipReasons.Add(new RowSkipReason(excelRow, $"DATE parse failed: {ex.Message}"));
                continue;
            }

            string fcRateRaw = GetString(worksheet, excelRow, IcmasteExcelCol.FcRate);
            decimal fcRate = 0;
            if (!IsBlank(fcRateRaw))
            {
                if (!decimal.TryParse(fcRateRaw, out var parsedFcRate))
                {
                    skipReasons.Add(new RowSkipReason(excelRow, $"FCRATE value '{fcRateRaw}' is not a valid number"));
                    continue;
                }
                fcRate = parsedFcRate;
            }

            string taxCode = GetString(worksheet, excelRow, IcmasteExcelCol.TaxCode);
            if (IsBlank(taxCode)) { taxCode = "SST0"; }

            var row = table.NewRow();
            row[IcmasteSchema.Type] = "IN";
            row[IcmasteSchema.Ref] = TextTruncation.Truncate(refValue, 11);
            row[IcmasteSchema.Date] = date.Date;
            row[IcmasteSchema.Code] = TextTruncation.Truncate(codeValue, 8);
            row[IcmasteSchema.Name] = TextTruncation.Truncate(nameValue, 40);
            row[IcmasteSchema.Accno] = "5001/000";
            row[IcmasteSchema.User] = "admin";
            row[IcmasteSchema.PostAccno] = "5001/000";
            row[IcmasteSchema.Cust2] = TextTruncation.Truncate(codeValue, 8);
            row[IcmasteSchema.Entry] = TextTruncation.Truncate(refValue, 10);
            row[IcmasteSchema.CurrCode] = "MYR";
            row[IcmasteSchema.TaxCode] = TextTruncation.Truncate(taxCode, 8);
            row[IcmasteSchema.FcRate] = fcRate;
            row[IcmasteSchema.AddCost] = DBNull.Value;
            row[IcmasteSchema.Tick] = DBNull.Value;
            row[IcmasteSchema.BillAge] = DBNull.Value;
            row[IcmasteSchema.DownAmt] = DBNull.Value;
            row[IcmasteSchema.Atms] = DBNull.Value;
            row[IcmasteSchema.DepPaid] = DBNull.Value;
            row[IcmasteSchema.Erefund] = DBNull.Value;
            table.Rows.Add(row);
        }

        return new StageReadResult { Table = table, RowsRead = rowsRead, SkipReasons = skipReasons };
    }

    /// <summary>
    /// Reads ictrane (transaction detail, many rows per document). Only blank required fields
    /// cause a row to be skipped — REF is intentionally NOT deduped here, since it legitimately
    /// repeats across many transaction lines per document (confirmed against real sample data).
    /// </summary>
    public StageReadResult ReadIctrane(IXLWorksheet worksheet)
    {
        ValidateColumnCount(worksheet, IctraneExcelCol.MinColumnCount);

        var table = IctraneSchema.BuildEmptyTable();
        var skipReasons = new List<RowSkipReason>();

        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 0;
        int rowsRead = 0;

        for (int excelRow = ExcelLayout.FirstDataRow; excelRow <= lastRow; excelRow++)
        {
            rowsRead++;

            string refValue = GetString(worksheet, excelRow, IctraneExcelCol.Ref);
            string itemNoValue = GetString(worksheet, excelRow, IctraneExcelCol.ItemNoAndDesc1);
            string qtyRaw = GetString(worksheet, excelRow, IctraneExcelCol.Qty);
            string priceRaw = GetString(worksheet, excelRow, IctraneExcelCol.Price);
            string amountRaw = GetString(worksheet, excelRow, IctraneExcelCol.Amount);

            if (IsBlank(refValue)) { skipReasons.Add(new RowSkipReason(excelRow, "ref is blank")); continue; }
            if (IsBlank(itemNoValue)) { skipReasons.Add(new RowSkipReason(excelRow, "item_no is blank")); continue; }
            if (IsBlank(qtyRaw) && IsBlank(priceRaw) && IsBlank(amountRaw))
            {
                skipReasons.Add(new RowSkipReason(excelRow, "qty, price and amount are all blank"));
                continue;
            }

            if (!TryParseOptionalDecimal(qtyRaw, "qty", excelRow, skipReasons, out var qty)) continue;
            if (!TryParseOptionalDecimal(priceRaw, "price", excelRow, skipReasons, out var price)) continue;
            if (!TryParseOptionalDecimal(amountRaw, "amount", excelRow, skipReasons, out var amount)) continue;

            string taxCode = GetString(worksheet, excelRow, IctraneExcelCol.TaxCode);
            if (IsBlank(taxCode)) { taxCode = "SST0"; }

            var row = table.NewRow();
            row[IctraneSchema.Type] = "IN";
            row[IctraneSchema.Ref] = TextTruncation.Truncate(refValue, 11);
            row[IctraneSchema.ItemNo] = TextTruncation.Truncate(itemNoValue, 24);
            // KNOWN LIMITATION: mirrors item_no's source column — see IctraneExcelCol.ItemNoAndDesc1.
            row[IctraneSchema.Desc1] = TextTruncation.Truncate(itemNoValue, 60);
            row[IctraneSchema.Qty] = (object?)qty ?? DBNull.Value;
            row[IctraneSchema.Price] = (object?)price ?? DBNull.Value;
            row[IctraneSchema.Amount] = (object?)amount ?? DBNull.Value;
            row[IctraneSchema.TaxCode] = TextTruncation.Truncate(taxCode, 8);
            row[IctraneSchema.UserId] = "admin";
            row[IctraneSchema.Entry] = TextTruncation.Truncate(refValue, 10);
            row[IctraneSchema.Accno] = "5002/000";
            row[IctraneSchema.Entry2] = RandomHexGenerator.Generate();
            table.Rows.Add(row);
        }

        return new StageReadResult { Table = table, RowsRead = rowsRead, SkipReasons = skipReasons };
    }

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

    private static bool TryParseOptionalDecimal(
        string raw, string fieldName, int excelRow, List<RowSkipReason> skipReasons, out decimal? value)
    {
        value = null;
        if (IsBlank(raw))
        {
            return true;
        }
        if (!decimal.TryParse(raw, out var parsed))
        {
            skipReasons.Add(new RowSkipReason(excelRow, $"{fieldName} value '{raw}' is not a valid number"));
            return false;
        }
        value = parsed;
        return true;
    }

    private static void ValidateColumnCount(IXLWorksheet worksheet, int minColumns)
    {
        int lastColumn = worksheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        if (lastColumn < minColumns)
        {
            throw new ExcelValidationException(
                $"The Excel file does not contain the required columns (found {lastColumn}, need at least {minColumns}).");
        }
    }

    private static string GetString(IXLWorksheet worksheet, int row, int zeroBasedColumn)
    {
        var cell = worksheet.Cell(row, zeroBasedColumn + 1);
        return cell.IsEmpty() ? string.Empty : cell.GetString().Trim();
    }

    private static bool IsBlank(string value) => string.IsNullOrWhiteSpace(value);

    private static DateTime ParseDate(IXLWorksheet worksheet, int row, int zeroBasedColumn, string rawString)
    {
        var cell = worksheet.Cell(row, zeroBasedColumn + 1);
        return cell.DataType == XLDataType.DateTime ? cell.GetDateTime() : DateTime.Parse(rawString);
    }
}
