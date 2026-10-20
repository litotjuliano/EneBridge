using System.Data;
using System.Data.OleDb;
using System.Security.AccessControl;
using System.Security.Principal;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// End-to-end round trip through the real OleDb/Access-Database-Engine DBF driver: reads the real
/// sample workbook, exports to actual .dbf files in a scratch folder, then reopens them via a
/// fresh OleDb connection to confirm structural validity. Requires the Access Database Engine to
/// be installed on the machine running the tests (same requirement the shipped app has).
/// </summary>
public class DbfExportServiceTests : IDisposable
{
    private readonly string _scratchFolder;

    public DbfExportServiceTests()
    {
        _scratchFolder = Path.Combine(Path.GetTempPath(), "eneBridge.Wpf.Tests." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_scratchFolder);
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

    private static string FixturePath => Path.Combine(AppContext.BaseDirectory, "Fixtures", "data.xlsx");

    [Fact]
    public void Export_IcmasteAndIctrane_ProducesValidDbfFiles()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var icmasteRead = excelReader.ReadIcmaste(worksheet);
        var icmasteExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);

        var ictraneRead = excelReader.ReadIctrane(worksheet);
        var ictraneExport = dbfExporter.Export(_scratchFolder, IctraneSchema.TableName, ictraneRead.Table, IctraneSchema.Columns);

        Assert.True(icmasteExport.Success, icmasteExport.ErrorMessage);
        Assert.Empty(icmasteExport.RowErrors);
        Assert.Equal(3, icmasteExport.RowsWritten);

        Assert.True(ictraneExport.Success, ictraneExport.ErrorMessage);
        Assert.Empty(ictraneExport.RowErrors);
        Assert.Equal(110, ictraneExport.RowsWritten);

        Assert.True(File.Exists(Path.Combine(_scratchFolder, "icmaste.dbf")));
        Assert.True(File.Exists(Path.Combine(_scratchFolder, "ictrane.dbf")));

        AssertRowCount("icmaste", 3);
        AssertRowCount("ictrane", 110);
    }

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

    [Fact]
    public void Export_SecondRun_BacksUpPreviousFileInsteadOfAccumulating()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);

        var firstExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        Assert.True(firstExport.Success, firstExport.ErrorMessage);
        Assert.Equal(3, firstExport.RowsWritten);

        // Ensure the backup timestamp (second-resolution) differs from the first export.
        Thread.Sleep(1100);

        var secondExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        Assert.True(secondExport.Success, secondExport.ErrorMessage);
        Assert.Equal(3, secondExport.RowsWritten);

        var backups = Directory.GetFiles(_scratchFolder, "icmaste_*.dbf");
        Assert.Single(backups);
        Assert.True(File.Exists(Path.Combine(_scratchFolder, "icmaste.dbf")));

        // The live file reflects only the second run's rows -- not accumulated with the first's.
        AssertRowCount("icmaste", 3);
    }

    [Fact]
    public void Export_ColumnNamesCollideAfterSanitization_PropagatesInvalidOperationException()
    {
        var dbfExporter = new DbfExportService();
        var schema = new List<DbfColumnDefinition>
        {
            new("MyField", typeof(string), 10),
            new("myfield", typeof(string), 10),
        };
        var table = new DataTable();
        table.Columns.Add("MyField", typeof(string));
        table.Columns.Add("myfield", typeof(string));
        table.Rows.Add("a", "b");

        Assert.Throws<InvalidOperationException>(() =>
            dbfExporter.Export(_scratchFolder, "collide", table, schema));
    }

    [Fact]
    public void Export_DbfFolderCannotBeCreated_ReturnsFailureWithErrorMessage()
    {
        // Deny CreateDirectories on the (existing) parent scratch folder so that Export's own
        // Directory.CreateDirectory(dbfFolder) call fails immediately with UnauthorizedAccessException
        // -- before any OleDb connection is ever opened. Deliberately NOT letting a real OleDb
        // connection/CREATE TABLE fail against a permission-denied folder: doing so was tried first,
        // and a failed CREATE TABLE mid-connection left the in-process ACE/Jet engine corrupted for
        // every OleDb connection opened afterward in the same test host, crashing the process with an
        // AccessViolationException in COM finalization on a later, unrelated test.
        var scratchInfo = new DirectoryInfo(_scratchFolder);
        var identity = WindowsIdentity.GetCurrent().User!;
        var denyRule = new FileSystemAccessRule(
            identity,
            FileSystemRights.CreateDirectories | FileSystemRights.Write,
            AccessControlType.Deny);

        var security = scratchInfo.GetAccessControl();
        security.AddAccessRule(denyRule);
        scratchInfo.SetAccessControl(security);

        try
        {
            var deniedFolder = Path.Combine(_scratchFolder, "denied-child");
            var dbfExporter = new DbfExportService();
            var table = new DataTable();
            table.Columns.Add("REF", typeof(string));
            table.Rows.Add("ROW1");
            var schema = new List<DbfColumnDefinition> { new("REF", typeof(string), 10) };

            var result = dbfExporter.Export(deniedFolder, "denied", table, schema);

            Assert.False(result.Success);
            Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
            Assert.Empty(result.RowErrors);
            Assert.Equal(0, result.RowsWritten);
        }
        finally
        {
            var cleanupSecurity = scratchInfo.GetAccessControl();
            cleanupSecurity.RemoveAccessRule(denyRule);
            scratchInfo.SetAccessControl(cleanupSecurity);
        }
    }

    private void AssertRowCount(string tableName, int expectedCount)
    {
        var connectionString =
            $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_scratchFolder};Extended Properties=""dBASE IV;CollatingSequence=1252"";";
        using var connection = new OleDbConnection(connectionString);
        connection.Open();
        using var command = new OleDbCommand($"SELECT COUNT(*) FROM {tableName}", connection);
        var count = Convert.ToInt32(command.ExecuteScalar());
        Assert.Equal(expectedCount, count);
    }
}
