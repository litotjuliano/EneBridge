using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// Real round-trip through the actual OleDb/ACE provider (no mocks), matching this codebase's
/// established test style (see DbfExportServiceTests/DbfReaderServiceTests) -- verifies the
/// current-count-vs-RowsWritten check that MainViewModel.ConfirmExportAsync relies on, now that
/// DbfExportService.Export backs up and recreates on every run (see its own doc comment).
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
    public async Task VerifyDbfAsync_SuccessfulExport_ReportsVerified()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();
        var dbfReader = new DbfReaderService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);
        var exportResult = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);

        string? status = null;
        bool? verified = null;
        await DbfVerificationHelper.VerifyDbfAsync(
            dbfReader, IcmasteSchema.TableName, _scratchFolder, IcmasteSchema.Type, "IN", exportResult,
            _ => { }, s => status = s, v => verified = v);

        Assert.True(verified);
        Assert.Contains("3 row(s) written and confirmed", status);
    }

    [Fact]
    public async Task VerifyDbfAsync_SecondExportRun_ReplacesNotAccumulates_StillReportsVerified()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();
        var dbfReader = new DbfReaderService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);

        dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);
        var secondExport = dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);

        string? status = null;
        bool? verified = null;
        await DbfVerificationHelper.VerifyDbfAsync(
            dbfReader, IcmasteSchema.TableName, _scratchFolder, IcmasteSchema.Type, "IN", secondExport,
            _ => { }, s => status = s, v => verified = v);

        // Backed up and recreated, not accumulated -- still exactly 3, not 6.
        Assert.True(verified);
        Assert.Contains("3 row(s) written and confirmed", status);
    }

    [Fact]
    public async Task VerifyDbfAsync_ExportFailed_ReportsFailureNotFalseSuccess()
    {
        // The table must already exist and be readable for this branch to be reachable at all --
        // VerifyDbfAsync's own "table unreadable" branch takes priority over the "export failed"
        // branch, since a failed read means there's no reliable row count to report either way. So
        // seed a real, successful export first (creating icmaste.dbf), then simulate a *subsequent*
        // export call failing against that now-existing table.
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
        await DbfVerificationHelper.VerifyDbfAsync(
            dbfReader, IcmasteSchema.TableName, _scratchFolder, IcmasteSchema.Type, "IN", failedExport,
            _ => { }, s => status = s, v => verified = v);

        Assert.False(verified);
        Assert.Contains("export failed", status);
    }

    [Fact]
    public async Task VerifyDbfAsync_RowCountMismatch_ReportsWarning()
    {
        var excelReader = new ExcelReaderService();
        var dbfExporter = new DbfExportService();
        var dbfReader = new DbfReaderService();

        using var workbook = excelReader.OpenWorkbook(FixturePath);
        var worksheet = workbook.Worksheets.First();
        var icmasteRead = excelReader.ReadIcmaste(worksheet);
        dbfExporter.Export(_scratchFolder, IcmasteSchema.TableName, icmasteRead.Table, IcmasteSchema.Columns);

        // The real export wrote 3 rows; claim it wrote 5 to force a mismatch.
        var mismatchedExportResult = new StageExportResult { Success = true, RowsWritten = 5 };

        string? status = null;
        bool? verified = null;
        await DbfVerificationHelper.VerifyDbfAsync(
            dbfReader, IcmasteSchema.TableName, _scratchFolder, IcmasteSchema.Type, "IN", mismatchedExportResult,
            _ => { }, s => status = s, v => verified = v);

        Assert.False(verified);
        Assert.Contains("wrote 5", status);
        Assert.Contains("shows 3", status);
    }
}
