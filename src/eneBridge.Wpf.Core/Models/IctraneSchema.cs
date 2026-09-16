using System.Data;

namespace eneBridge.Wpf.Core.Models;

/// <summary>
/// EMAS Inventory Control transaction (detail) table schema (80 columns, lower-case names).
/// Column names, order, types and max lengths are reverse-engineered from the original eneBridge
/// console app's IL and must be preserved exactly, since this is what the legacy EMAS accounting
/// system expects on import.
/// </summary>
public static class IctraneSchema
{
    public const string TableName = "ictrane";

    /// <summary>
    /// EMAS's own live/master table for line items — the ictran.dbf counterpart to
    /// IcmasteSchema.LiveTableName's icmast.dbf (see that constant's doc comment for the format
    /// details: Visual FoxPro, read via FoxProDbfReader, not OleDb). Used by
    /// DuplicateDocumentChecker alongside icmast.dbf: a document is only treated as a genuine
    /// duplicate if BOTH its header (icmast) and at least one line item (ictran) still exist for
    /// that REF — confirmed against a real installation that EMAS's own delete can leave an
    /// orphaned icmast header behind after removing a document's ictran line items (a "soft
    /// delete", per the client), which an icmast-only check would incorrectly still block as a
    /// duplicate.
    /// </summary>
    public const string LiveTableName = "ictran";

    // Columns actually populated from Excel — see ExcelReaderService.
    public const string Type = "type";
    public const string Entry = "entry";
    public const string ItemNo = "item_no";
    public const string Desc1 = "desc1";
    public const string Desc2 = "desc2";
    public const string Qty = "qty";
    public const string Price = "price";
    public const string Amount = "amount";
    public const string TaxCode = "taxcode";
    public const string UserId = "userid";
    public const string Accno = "accno";
    public const string Ref = "ref";
    public const string Entry2 = "entry2";

    public static IReadOnlyList<DbfColumnDefinition> Columns { get; } = BuildColumns();

    public static DataTable BuildEmptyTable()
    {
        var table = new DataTable(TableName);
        foreach (var column in Columns)
        {
            var dataColumn = table.Columns.Add(column.Name, column.ClrType);
            if (column.MaxLength > 0)
            {
                dataColumn.MaxLength = column.MaxLength;
            }
        }
        return table;
    }

    private static List<DbfColumnDefinition> BuildColumns()
    {
        var s = typeof(string);
        var d = typeof(DateTime);
        var m = typeof(decimal);

        var columns = new List<DbfColumnDefinition>
        {
            new(Type, s, 2),
            new(Entry, s, 10),
            new("sequence", s, 1),
            new(ItemNo, s, 24),
            new(Desc1, s, 60),
            new(Desc2, s, 40),
            new("location", s, 10),
            new("bin", s, 10),
            new("toloc", s, 10),
            new("tobin", s, 10),
            new(Qty, m),
            new(Price, m),
            new(Amount, m),
            new("cvalue", m),
            new("date", d),
            new("dsource", s, 10),
            new("djob", s, 10),
            new("u_measure", s, 6),
            new("qty1", m),
            new("po_entry", s, 10),
        };

        for (int i = 1; i <= 5; i++)
        {
            columns.Add(new DbfColumnDefinition($"dfield{i}", s, 40));
        }
        for (int i = 6; i <= 15; i++)
        {
            columns.Add(new DbfColumnDefinition($"dfield{i}", m));
        }

        columns.AddRange(new[]
        {
            new DbfColumnDefinition("n3type", s, 1),
            new DbfColumnDefinition("billno", s, 11),
            new DbfColumnDefinition("billdate", d),
            new DbfColumnDefinition(Accno, s, 8),
            new DbfColumnDefinition("misccode", s, 3),
            new DbfColumnDefinition("do_entry", s, 10),
            new DbfColumnDefinition(Entry2, s, 10),
            new DbfColumnDefinition("reason", s, 30),
            new DbfColumnDefinition("fcamt", m),
            new DbfColumnDefinition("po_no", s, 20),
            new DbfColumnDefinition("po_date", d),
            new DbfColumnDefinition("po_entry2", s, 10),
            new DbfColumnDefinition("subtotal", s, 1),
            new DbfColumnDefinition(TaxCode, s, 8),
            new DbfColumnDefinition("line_no", m),
            new DbfColumnDefinition("newrec", s, 1),
            new DbfColumnDefinition("exported", s, 1),
            new DbfColumnDefinition("itemdisc", s, 20),
            new DbfColumnDefinition("u_measure1", s, 6),
            new DbfColumnDefinition("qty2", m),
            new DbfColumnDefinition("dvalue", m),
            new DbfColumnDefinition("measure", s, 10),
            new DbfColumnDefinition("qtyinfo", m),
            new DbfColumnDefinition("uqty", s, 20),
            new DbfColumnDefinition("u_cj5", s, 1),
            new DbfColumnDefinition("netprice", m),
            new DbfColumnDefinition("sobal", m),
            new DbfColumnDefinition("sobal2", m),
            new DbfColumnDefinition("overlimit", s, 1),
            new DbfColumnDefinition(Ref, s, 11),
            new DbfColumnDefinition(UserId, s, 10),
            new DbfColumnDefinition("compdate", d),
            new DbfColumnDefinition("comptime", s, 8),
            new DbfColumnDefinition("ref8", s, 14),
            new DbfColumnDefinition("sbatchno", s, 20),
            new DbfColumnDefinition("rbatchno", s, 20),
            new DbfColumnDefinition("editrec", s, 1),
            new DbfColumnDefinition("u_cost", m),
            new DbfColumnDefinition("monthend", s, 1),
            new DbfColumnDefinition("t_cvalue", m),
            new DbfColumnDefinition("shipcode", s, 8),
            new DbfColumnDefinition("fm_no", s, 11),
            new DbfColumnDefinition("discamt", m),
            new DbfColumnDefinition("taxamt", m),
            new DbfColumnDefinition("assembly", s, 1),
        });

        return columns;
    }
}
