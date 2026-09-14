using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// Round trip through the real OleDb/Access-Database-Engine DBF driver: exports the real sample
/// workbook via <see cref="DbfExportService"/>, then reads it back via <see cref="DbfReaderService"/>
/// to confirm what actually landed on disk. Requires the Access Database Engine to be installed
/// (same requirement as <see cref="DbfExportServiceTests"/>).
/// </summary>
public class DbfReaderServiceTests : IDisposable
{
    private readonly string _scratchFolder;

    public DbfReaderServiceTests()
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
    public void Read_AfterExport_ReturnsMatchingRowCountAndData()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();
        var dbfReader = new DbfReaderService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();

        var icmasteRead = excelReader.ReadIcmaste(worksheet);
        var icmasteExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        Assert.True(icmasteExport.Success, icmasteExport.ErrorMessage);

        var result = dbfReader.Read(_scratchFolder, IcmasteSchema.TableName);

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(icmasteExport.RowsWritten, result.RowCount);
        Assert.Equal(3, result.RowCount);

        var expectedFirstRef = icmasteRead.Table.Rows[0][IcmasteSchema.Ref];
        Assert.Equal(expectedFirstRef, result.Table.Rows[0][IcmasteSchema.Ref]);
    }

    [Fact]
    public void Read_TableDoesNotExist_ReturnsFailureWithErrorMessage()
    {
        var dbfReader = new DbfReaderService();

        var result = dbfReader.Read(_scratchFolder, "nosuchtable");

        Assert.False(result.Success);
        Assert.False(string.IsNullOrEmpty(result.ErrorMessage));
        Assert.Equal(0, result.RowCount);
    }

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
}
