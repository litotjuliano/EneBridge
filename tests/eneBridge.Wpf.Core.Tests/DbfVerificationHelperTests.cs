using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// Real round-trip through the actual OleDb/ACE provider (no mocks), matching this codebase's
/// established test style (see DbfExportServiceTests/DbfReaderServiceTests) -- verifies the
/// append-aware before/after-delta logic that MainViewModel.ConfirmExportAsync now relies on.
/// </summary>
public class DbfVerificationHelperTests : IDisposable
{
    private readonly string _scratchFolder;

    public DbfVerificationHelperTests()
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
    public async Task VerifyDbfDeltaAsync_TwoExportRunsInARow_ReportsCorrectDeltaEachTime()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();
        var dbfReader = new DbfReaderService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);

        // Run 1
        var before1 = await DbfVerificationHelper.CountRowsByTypeAsync(dbfReader, _scratchFolder, IcmasteSchema.TableName, IcmasteSchema.Type, "IN");
        var export1 = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        string? status1 = null;
        bool? verified1 = null;
        await DbfVerificationHelper.VerifyDbfDeltaAsync(
            dbfReader, IcmasteSchema.TableName, _scratchFolder, IcmasteSchema.Type, "IN", before1, export1,
            _ => { }, s => status1 = s, v => verified1 = v);

        Assert.Equal(0, before1);
        Assert.True(verified1);
        Assert.Contains("3 row(s) written and confirmed", status1);

        // Run 2 -- same data exported again, rows should accumulate
        var before2 = await DbfVerificationHelper.CountRowsByTypeAsync(dbfReader, _scratchFolder, IcmasteSchema.TableName, IcmasteSchema.Type, "IN");
        var export2 = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        string? status2 = null;
        bool? verified2 = null;
        await DbfVerificationHelper.VerifyDbfDeltaAsync(
            dbfReader, IcmasteSchema.TableName, _scratchFolder, IcmasteSchema.Type, "IN", before2, export2,
            _ => { }, s => status2 = s, v => verified2 = v);

        Assert.Equal(3, before2);
        Assert.True(verified2);
        Assert.Contains("3 row(s) written and confirmed", status2);
        Assert.Contains("6 total", status2);
    }

    [Fact]
    public async Task VerifyDbfDeltaAsync_ExportFailed_ReportsFailureNotFalseSuccess()
    {
        // The table must already exist and be readable for this branch to be reachable at all --
        // VerifyDbfDeltaAsync's own "table unreadable" branch takes priority over the "export
        // failed" branch, since a failed read means there's no reliable row count to report either
        // way. So seed a real, successful export first (creating icmaste.dbf), then simulate a
        // *subsequent* export call failing against that now-existing table.
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();
        var dbfReader = new DbfReaderService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);
        dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);

        var failedExport = new StageExportResult { Success = false, ErrorMessage = "boom" };

        string? status = null;
        bool? verified = null;
        await DbfVerificationHelper.VerifyDbfDeltaAsync(
            dbfReader, IcmasteSchema.TableName, _scratchFolder, IcmasteSchema.Type, "IN", beforeCount: 0, failedExport,
            _ => { }, s => status = s, v => verified = v);

        Assert.False(verified);
        Assert.Contains("export failed", status);
    }

    [Fact]
    public async Task VerifyDbfDeltaAsync_ActualNewRowsNegative_ReportsFewerRowsMessage()
    {
        // Drive the actualNewRows < 0 sub-branch (rows appear to have disappeared since the
        // "before" count was taken) without needing to actually delete real rows from a DBF: seed
        // a real export (3 rows, per the fixture), then pass a beforeCount deliberately higher
        // than the table's actual post-export row count, so readResult.RowCount - beforeCount
        // computes negative even though the export itself reported success.
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();
        var dbfReader = new DbfReaderService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);
        dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);

        var exportResult = new StageExportResult { Success = true, RowsWritten = 3 };

        string? status = null;
        bool? verified = null;
        await DbfVerificationHelper.VerifyDbfDeltaAsync(
            dbfReader, IcmasteSchema.TableName, _scratchFolder, IcmasteSchema.Type, "IN", beforeCount: 10, exportResult,
            _ => { }, s => status = s, v => verified = v);

        Assert.False(verified);
        Assert.Contains("FEWER", status);
    }
}
