using System.Data;

namespace eneBridge.Wpf.Core.Models;

/// <summary>
/// EMAS Inventory Control master table schema (151 columns). Column names, order, types and
/// max lengths are reverse-engineered from the original eneBridge console app's IL and must be
/// preserved exactly, since this is what the legacy EMAS accounting system expects on import.
/// </summary>
public static class IcmasteSchema
{
    public const string TableName = "icmaste";

    // Columns actually populated from Excel — see ExcelReaderService.
    public const string Void = "VOID";
    public const string Type = "TYPE";
    public const string Entry = "ENTRY";
    public const string Ref = "REF";
    public const string Date = "DATE";
    public const string Code = "CODE";
    public const string Name = "NAME";
    public const string Posted = "POSTED";
    public const string Accno = "ACCNO";
    public const string User = "USER";
    public const string PostAccno = "POSTACCNO";
    public const string Cust2 = "CUST2";
    public const string CurrCode = "CURRCODE";
    public const string TaxCode = "TAXCODE";
    public const string FcRate = "FCRATE";
    public const string AddCost = "ADD_COST";
    public const string Tick = "TICK";
    public const string BillAge = "BILLAGE";
    public const string DownAmt = "DOWNAMT";
    public const string Atms = "ATMS";
    public const string DepPaid = "DEPPAID";
    public const string Erefund = "EREFUND";

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
        var b = typeof(bool);

        var columns = new List<DbfColumnDefinition>
        {
            new(Void, s, 1),
            new(Type, s, 2),
            new(Entry, s, 10),
            new(Ref, s, 11),
            new(Date, d),
            new(Code, s, 8),
            new(Name, s, 40),
            new("NAME2", s, 40),
            new("DOC", s, 11),
            new("DOC1", s, 11),
            new("DO_NO", s, 11),
            new("T_AMT", m),
            new("N_AMT", m),
            new(AddCost, m),
            new("ADD_ACC", s, 8),
            new("ADD_ACC1", s, 8),
            new("ADD_DESC", s, 25),
            new("REMARK1", s, 25),
            new("REMARK2", s, 25),
            new("REASON", s, 30),
            new("HSOURCE", s, 10),
            new("HJOB", s, 10),
            new(Posted, s, 1),
            new(Accno, s, 8),
        };

        for (int i = 1; i <= 14; i++)
        {
            columns.Add(new DbfColumnDefinition($"HFIELD{i}", s, 40));
        }
        for (int i = 1; i <= 8; i++)
        {
            columns.Add(new DbfColumnDefinition($"MISCDESC{i}", s, 40));
        }
        for (int i = 1; i <= 8; i++)
        {
            columns.Add(new DbfColumnDefinition($"MISCAMT{i}", m));
        }
        for (int i = 1; i <= 8; i++)
        {
            columns.Add(new DbfColumnDefinition($"FMISCAMT{i}", m));
        }
        for (int i = 1; i <= 4; i++)
        {
            columns.Add(new DbfColumnDefinition($"ADD{i}", s, 35));
        }
        for (int i = 1; i <= 4; i++)
        {
            columns.Add(new DbfColumnDefinition($"DADDR{i}", s, 35));
        }

        columns.AddRange(new[]
        {
            new DbfColumnDefinition("REM1", s, 40),
            new DbfColumnDefinition("REM2", s, 40),
            new DbfColumnDefinition("AGENT", s, 12),
            new DbfColumnDefinition("TERM", s, 12),
            new DbfColumnDefinition("AREA", s, 12),
            new DbfColumnDefinition("BUSINESS", s, 15),
            new DbfColumnDefinition(Tick, b),
            new DbfColumnDefinition("IN_NO", s, 11),
        });

        for (int i = 1; i <= 10; i++)
        {
            columns.Add(new DbfColumnDefinition($"BILLNO{i}", s, 11));
        }
        for (int i = 1; i <= 10; i++)
        {
            columns.Add(new DbfColumnDefinition($"BILLDATE{i}", d));
        }

        columns.AddRange(new[]
        {
            new DbfColumnDefinition(BillAge, m),
            new DbfColumnDefinition("RETR", s, 1),
            new DbfColumnDefinition(User, s, 10),
            new DbfColumnDefinition(PostAccno, s, 8),
            new DbfColumnDefinition("VENDNO", s, 8),
            new DbfColumnDefinition(FcRate, m),
            new DbfColumnDefinition("PO_NO", s, 20),
            new DbfColumnDefinition("DOWNACC", s, 8),
            new DbfColumnDefinition(DownAmt, m),
            new DbfColumnDefinition("DOWNDESC", s, 20),
            new DbfColumnDefinition(TaxCode, s, 8),
            new DbfColumnDefinition("TAXAMT", m),
            new DbfColumnDefinition("REF2", s, 11),
            new DbfColumnDefinition("ASSEMBLY", s, 1),
            new DbfColumnDefinition("PRINT", s, 1),
            new DbfColumnDefinition("T_LINE", m),
            new DbfColumnDefinition("EXPORT", s, 1),
            new DbfColumnDefinition(CurrCode, s, 4),
            new DbfColumnDefinition("DUEDATE", d),
            new DbfColumnDefinition("INVDATE", d),
            new DbfColumnDefinition("TIME", s, 8),
            new DbfColumnDefinition("FM_NO", s, 11),
            new DbfColumnDefinition("CASHPUR", s, 1),
            new DbfColumnDefinition("STR_NO", s, 11),
            new DbfColumnDefinition("STR_ENTRY", s, 10),
            new DbfColumnDefinition("RE_NO", s, 11),
            new DbfColumnDefinition("HMAINACC", s, 8),
            new DbfColumnDefinition("OVERTERM", s, 1),
            new DbfColumnDefinition("ADTORE", s, 1),
            new DbfColumnDefinition("NEWPOST", s, 1),
            new DbfColumnDefinition(Cust2, s, 8),
            new DbfColumnDefinition("REF8", s, 14),
            new DbfColumnDefinition("SBATCHNO", s, 20),
            new DbfColumnDefinition("RBATCHNO", s, 20),
            new DbfColumnDefinition("EDITREC", s, 1),
            new DbfColumnDefinition("FOOTER", s, 4),
            new DbfColumnDefinition("MONTHEND", s, 1),
            new DbfColumnDefinition("SEQUENCE", s, 1),
            new DbfColumnDefinition("SHIPCODE", s, 8),
            new DbfColumnDefinition("DOWNTOSO", s, 1),
            new DbfColumnDefinition("SO_NO", s, 20),
            new DbfColumnDefinition("CPERMITNO", s, 20),
            new DbfColumnDefinition(Atms, b),
            new DbfColumnDefinition(DepPaid, b),
            new DbfColumnDefinition("COUNTRY", s, 50),
            new DbfColumnDefinition("IN_DATE", d),
            new DbfColumnDefinition("CA_NO", s, 11),
            new DbfColumnDefinition("CUSTDATE", d),
            new DbfColumnDefinition("OVERLIMIT", s, 1),
            new DbfColumnDefinition("PRTCOUNT", m),
            new DbfColumnDefinition("ETYPE", s, 2),
            new DbfColumnDefinition("EREFNO", s, 20),
            new DbfColumnDefinition(Erefund, b),
        });

        return columns;
    }
}
