using System.Data;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// Real round-trip through the actual OleDb/ACE provider (no mocks), matching this codebase's
/// established test style (see DbfVerificationHelperTests/DbfExportServiceTests) -- seeds rows
/// into icmast.dbf (IcmasteSchema.LiveTableName), the table FindDuplicateRefsAsync actually reads
/// (EMAS's own live table, not eneBridge's icmaste.dbf staging file -- see DuplicateDocumentChecker's
/// class-level doc comment), via a real DbfExportService.Export call rather than faking a
/// DbfReadResult. Reuses IcmasteSchema.Columns since icmast.dbf is assumed to share icmaste.dbf's
/// column layout.
/// </summary>
public class DuplicateDocumentCheckerTests : IDisposable
{
    private readonly string _scratchFolder;

    public DuplicateDocumentCheckerTests()
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

    private static void SeedIcmastRow(string scratchFolder, string type, string reference, string code)
    {
        var dbfExporter = new DbfExportService();
        var table = IcmasteSchema.BuildEmptyTable();
        var row = table.NewRow();
        row[IcmasteSchema.Type] = type;
        row[IcmasteSchema.Ref] = reference;
        row[IcmasteSchema.Code] = code;
        table.Rows.Add(row);

        var exportResult = dbfExporter.Export(scratchFolder, IcmasteSchema.LiveTableName, table, IcmasteSchema.Columns);
        Assert.True(exportResult.Success, exportResult.ErrorMessage);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_TableDoesNotExist_ReturnsEmpty()
    {
        var dbfReader = new DbfReaderService();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(dbfReader, _scratchFolder, "IN", candidates);

        Assert.Empty(result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_NoCandidates_ReturnsEmpty()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var dbfReader = new DbfReaderService();

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(
            dbfReader, _scratchFolder, "IN", Array.Empty<(string Ref, string Code)>());

        Assert.Empty(result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_ExactRefAndCodeMatch_ReturnsThatRef()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var dbfReader = new DbfReaderService();
        var candidates = new List<(string Ref, string Code)> { ("X", "A"), ("Y", "B") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(dbfReader, _scratchFolder, "IN", candidates);

        Assert.Equal(new[] { "X" }, result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_SameRefDifferentCode_ReturnsEmpty()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var dbfReader = new DbfReaderService();
        var candidates = new List<(string Ref, string Code)> { ("X", "B") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(dbfReader, _scratchFolder, "IN", candidates);

        Assert.Empty(result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_SameRefDifferentType_ReturnsEmpty()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var dbfReader = new DbfReaderService();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(dbfReader, _scratchFolder, "RE", candidates);

        Assert.Empty(result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_DuplicateCandidatesInInput_ReturnsDistinctRefs()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var dbfReader = new DbfReaderService();
        var candidates = new List<(string Ref, string Code)> { ("X", "A"), ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(dbfReader, _scratchFolder, "IN", candidates);

        Assert.Equal(new[] { "X" }, result);
    }
}
