using System.Data;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// Real round-trip (no mocks), matching this codebase's established test style (see
/// DbfVerificationHelperTests/DbfExportServiceTests) -- seeds "already exported" rows into
/// icmast.dbf (IcmasteSchema.LiveTableName, the table FindDuplicateRefsAsync actually reads) via a
/// real DbfExportService.Export call, then reads them back with the actual FoxProDbfReader rather
/// than faking a DbfReadResult. DbfExportService writes plain dBASE III bytes (not genuine Visual
/// FoxPro), but FoxProDbfReader doesn't validate the version byte -- the field/record layout is
/// identical across the dBASE/FoxPro family for character fields, which is all REF/CODE/TYPE are
/// -- so this is a faithful test of the reader's actual parsing logic, exercised through this
/// project's existing OleDb-based seeding helper instead of hand-built byte fixtures (that
/// coverage lives in FoxProDbfReaderTests instead, including the real 0x30 Visual FoxPro version
/// byte).
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
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.Empty(result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_NoCandidates_ReturnsEmpty()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var foxProDbfReader = new FoxProDbfReader();

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(
            foxProDbfReader, _scratchFolder, "IN", Array.Empty<(string Ref, string Code)>());

        Assert.Empty(result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_ExactRefAndCodeMatch_ReturnsThatRef()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A"), ("Y", "B") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.Equal(new[] { "X" }, result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_SameRefDifferentCode_ReturnsEmpty()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "B") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.Empty(result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_SameRefDifferentType_ReturnsEmpty()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "RE", candidates);

        Assert.Empty(result);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_DuplicateCandidatesInInput_ReturnsDistinctRefs()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A"), ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.Equal(new[] { "X" }, result);
    }
}
