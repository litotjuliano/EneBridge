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
