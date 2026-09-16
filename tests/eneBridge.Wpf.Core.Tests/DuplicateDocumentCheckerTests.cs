using System.Data;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// Real round-trip (no mocks), matching this codebase's established test style (see
/// DbfVerificationHelperTests/DbfExportServiceTests) -- seeds "already imported" rows into
/// icmast.dbf/ictran.dbf (IcmasteSchema.LiveTableName/IctraneSchema.LiveTableName, the tables
/// FindDuplicateRefsAsync actually reads) via real DbfExportService.Export calls, then reads them
/// back with the actual FoxProDbfReader rather than faking a DbfReadResult. DbfExportService writes
/// plain dBASE III bytes (not genuine Visual FoxPro), but FoxProDbfReader doesn't validate the
/// version byte -- the field/record layout is identical across the dBASE/FoxPro family for
/// character fields, which is all REF/CODE/TYPE are -- so this is a faithful test of the reader's
/// actual parsing logic, exercised through this project's existing OleDb-based seeding helper
/// instead of hand-built byte fixtures (that coverage lives in FoxProDbfReaderTests instead,
/// including the real 0x30 Visual FoxPro version byte).
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

    private static void SeedIctranRow(string scratchFolder, string type, string reference)
    {
        var dbfExporter = new DbfExportService();
        var table = IctraneSchema.BuildEmptyTable();
        var row = table.NewRow();
        row[IctraneSchema.Type] = type;
        row[IctraneSchema.Ref] = reference;
        table.Rows.Add(row);

        var exportResult = dbfExporter.Export(scratchFolder, IctraneSchema.LiveTableName, table, IctraneSchema.Columns);
        Assert.True(exportResult.Success, exportResult.ErrorMessage);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_TableDoesNotExist_VerifiedWithNoDuplicates()
    {
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.True(result.Verified);
        Assert.Empty(result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_NoCandidates_VerifiedWithNoDuplicates()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        SeedIctranRow(_scratchFolder, "IN", "X");
        var foxProDbfReader = new FoxProDbfReader();

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(
            foxProDbfReader, _scratchFolder, "IN", Array.Empty<(string Ref, string Code)>());

        Assert.True(result.Verified);
        Assert.Empty(result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_IcmastAndIctranBothMatch_ReturnsThatRef()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        SeedIctranRow(_scratchFolder, "IN", "X");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A"), ("Y", "B") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.True(result.Verified);
        Assert.Equal(new[] { "X" }, result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_IcmastMatchButNoIctranLineItems_ReturnsEmpty()
    {
        // Confirmed against a real EMAS installation: EMAS's own delete can leave an orphaned
        // icmast.dbf header behind after removing a document's ictran.dbf line items (a "soft
        // delete", per the client) -- an icmast-only check would wrongly still block re-export of a
        // document EMAS's own screen could no longer even find.
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        SeedIctranRow(_scratchFolder, "IN", "Y"); // ictran.dbf exists and is readable, just has no match
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.True(result.Verified);
        Assert.Empty(result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_IctranMatchButNoIcmastHeader_VerifiedWithNoDuplicates()
    {
        SeedIcmastRow(_scratchFolder, "IN", "OTHER", "Z"); // icmast.dbf exists and is readable, just has no match
        SeedIctranRow(_scratchFolder, "IN", "X");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.True(result.Verified);
        Assert.Empty(result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_SameRefDifferentCode_ReturnsEmpty()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        SeedIctranRow(_scratchFolder, "IN", "X");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "B") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.True(result.Verified);
        Assert.Empty(result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_SameRefDifferentType_ReturnsEmpty()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        SeedIctranRow(_scratchFolder, "IN", "X");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "RE", candidates);

        Assert.True(result.Verified);
        Assert.Empty(result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_DuplicateCandidatesInInput_ReturnsDistinctRefs()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        SeedIctranRow(_scratchFolder, "IN", "X");
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A"), ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.True(result.Verified);
        Assert.Equal(new[] { "X" }, result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_IcmastExistsButUnreadable_ReturnsUnverified()
    {
        // Confirmed as a real production bug, not a hypothetical: EMAS holding icmast.dbf/ictran.dbf
        // open (an ordinary occurrence, since the whole workflow revolves around switching between
        // eneBridge and EMAS) used to make FindDuplicateRefsAsync silently report "no duplicates"
        // instead of "couldn't check" -- letting a real duplicate document through uncaught. A
        // too-short file triggers the same kind of read failure (EndOfStreamException) a lock would.
        var icmastPath = Path.Combine(_scratchFolder, IcmasteSchema.LiveTableName + ".dbf");
        var ictranPath = Path.Combine(_scratchFolder, IctraneSchema.LiveTableName + ".dbf");
        File.WriteAllBytes(icmastPath, new byte[] { 0x30 });
        File.WriteAllBytes(ictranPath, new byte[] { 0x30 });
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.False(result.Verified);
        Assert.False(string.IsNullOrEmpty(result.UnverifiableReason));
        Assert.Empty(result.DuplicateRefs);
    }

    [Fact]
    public async Task FindDuplicateRefsAsync_IctranExistsButUnreadable_ReturnsUnverified()
    {
        SeedIcmastRow(_scratchFolder, "IN", "X", "A");
        var ictranPath = Path.Combine(_scratchFolder, IctraneSchema.LiveTableName + ".dbf");
        File.WriteAllBytes(ictranPath, new byte[] { 0x30 });
        var foxProDbfReader = new FoxProDbfReader();
        var candidates = new List<(string Ref, string Code)> { ("X", "A") };

        var result = await DuplicateDocumentChecker.FindDuplicateRefsAsync(foxProDbfReader, _scratchFolder, "IN", candidates);

        Assert.False(result.Verified);
        Assert.False(string.IsNullOrEmpty(result.UnverifiableReason));
    }
}
