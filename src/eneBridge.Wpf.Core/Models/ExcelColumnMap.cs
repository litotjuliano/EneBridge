namespace eneBridge.Wpf.Core.Models;

/// <summary>
/// 0-based raw Excel column indices used by the original eneBridge app's mapping logic,
/// reverse-engineered from IL and confirmed against the real sample data.xlsx (columns A-L).
/// </summary>
public static class IcmasteExcelCol
{
    public const int Date = 5;
    public const int Code = 6;
    public const int Name = 7;
    public const int Ref = 8;
    public const int FcRate = 10;
    public const int TaxCode = 11;

    /// <summary>Minimum raw column count required (original throws below this).</summary>
    public const int MinColumnCount = 8;
}

public static class IctraneExcelCol
{
    /// <summary>
    /// KNOWN LIMITATION: the original app reads both item_no and desc1 from this same column
    /// (raw index 1), which is almost certainly a bug (desc1 should be a real description field),
    /// but the correct source column could not be determined from the one available sample
    /// workbook. Kept as-is per product decision until the correct source is confirmed with
    /// whoever maintains the Excel template.
    /// </summary>
    public const int ItemNoAndDesc1 = 1;

    public const int Qty = 2;
    public const int Price = 3;
    public const int Amount = 4;
    public const int Ref = 8;
    public const int TaxCode = 11;

    /// <summary>Minimum raw column count required (original throws below this).</summary>
    public const int MinColumnCount = 9;
}

/// <summary>
/// Defines where real data starts in the source workbook. The current invoice Excel template has
/// a single header row (row 1) with real data starting immediately on row 2 — no extra rows to
/// skip. (An older template needed 4 additional rows skipped after the header, treating physical
/// row 6 as the first data row — reverse-engineered from the original app's OleDb HDR=YES
/// behavior. That template has been replaced and no longer applies.)
/// </summary>
public static class ExcelLayout
{
    public const int HeaderRows = 1;
    public const int SkipDataRows = 0;

    /// <summary>1-based physical row number of the first row that is actually processed.</summary>
    public static int FirstDataRow => HeaderRows + SkipDataRows + 1;
}
