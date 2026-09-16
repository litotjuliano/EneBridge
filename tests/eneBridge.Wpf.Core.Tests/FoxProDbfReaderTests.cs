using System.Text;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// Tests against hand-built byte fixtures matching the classic dBASE/FoxPro free-table layout
/// confirmed byte-for-byte against a real icmast.dbf during Stock Received testing (see
/// FoxProDbfReader's class doc comment and CLAUDE.md's "Known reliability risk" section).
/// </summary>
public class FoxProDbfReaderTests : IDisposable
{
    private readonly string _scratchFolder;
    private static readonly Encoding TextEncoding = Encoding.Latin1;

    public FoxProDbfReaderTests()
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

    private readonly record struct FieldDef(string Name, byte Length);

    /// <summary>
    /// Builds a byte-accurate classic dBASE/FoxPro free-table file: 32-byte header, one 32-byte
    /// descriptor per field, a 0x0D terminator, then one fixed-width record per row (a leading
    /// deletion-flag byte -- 0x20 active, 0x2A deleted -- followed by each field's value
    /// space-padded to its declared length).
    /// </summary>
    private static void WriteFoxProDbf(
        string path, byte versionByte, IReadOnlyList<FieldDef> fields,
        IReadOnlyList<(bool Deleted, IReadOnlyDictionary<string, string> Values)> rows)
    {
        int recordLength = 1 + fields.Sum(f => f.Length);
        int headerLength = 32 + (fields.Count * 32) + 1;

        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
        using var writer = new BinaryWriter(stream);

        writer.Write(versionByte);
        writer.Write((byte)0); writer.Write((byte)0); writer.Write((byte)0); // last-update date, unused
        writer.Write((uint)rows.Count);
        writer.Write((ushort)headerLength);
        writer.Write((ushort)recordLength);
        writer.Write(new byte[20]); // reserved, pad header to 32 bytes

        foreach (var field in fields)
        {
            var nameBytes = new byte[11];
            var encodedName = TextEncoding.GetBytes(field.Name);
            Array.Copy(encodedName, nameBytes, Math.Min(encodedName.Length, 10));
            writer.Write(nameBytes);
            writer.Write((byte)'C'); // field type -- FoxProDbfReader treats every field as text
            writer.Write(new byte[4]); // field data address, unused
            writer.Write(field.Length);
            writer.Write((byte)0); // decimal count, unused for character fields
            writer.Write(new byte[14]); // reserved
        }
        writer.Write((byte)0x0D); // field descriptor terminator

        foreach (var (deleted, values) in rows)
        {
            writer.Write(deleted ? (byte)'*' : (byte)' ');
            foreach (var field in fields)
            {
                var raw = values.TryGetValue(field.Name, out var value) ? value : string.Empty;
                var fieldBytes = new byte[field.Length];
                Array.Fill(fieldBytes, (byte)' ');
                var encoded = TextEncoding.GetBytes(raw);
                Array.Copy(encoded, fieldBytes, Math.Min(encoded.Length, field.Length));
                writer.Write(fieldBytes);
            }
        }
    }

    private static IReadOnlyDictionary<string, string> Row(params (string Name, string Value)[] values) =>
        values.ToDictionary(v => v.Name, v => v.Value);

    [Fact]
    public void Read_ValidFile_ReturnsAllActiveRowsTrimmed()
    {
        var path = Path.Combine(_scratchFolder, "icmast.dbf");
        var fields = new[] { new FieldDef("TYPE", 2), new FieldDef("REF", 11), new FieldDef("CODE", 8) };
        WriteFoxProDbf(path, versionByte: 0x30, fields, new[]
        {
            (false, Row(("TYPE", "IN"), ("REF", "SB-000001"), ("CODE", "4000T001"))),
            (false, Row(("TYPE", "RE"), ("REF", "SB-000002"), ("CODE", "4000T002"))),
        });

        var result = new FoxProDbfReader().Read(_scratchFolder, "icmast");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(2, result.RowCount);
        Assert.Equal("IN", result.Table.Rows[0]["TYPE"]);
        Assert.Equal("SB-000001", result.Table.Rows[0]["REF"]);
        Assert.Equal("4000T001", result.Table.Rows[0]["CODE"]);
        Assert.Equal("RE", result.Table.Rows[1]["TYPE"]);
    }

    [Fact]
    public void Read_VisualFoxProVersionByte_ParsesCorrectly()
    {
        // The exact scenario this reader exists for: a real icmast.dbf's version byte is 0x30
        // (Visual FoxPro), which the ACE OleDb dBASE-IV driver cannot parse at all.
        var path = Path.Combine(_scratchFolder, "icmast.dbf");
        var fields = new[] { new FieldDef("REF", 11) };
        WriteFoxProDbf(path, versionByte: 0x30, fields, new[]
        {
            (false, Row(("REF", "SB-000001"))),
        });

        var result = new FoxProDbfReader().Read(_scratchFolder, "icmast");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Table.Rows);
    }

    [Fact]
    public void Read_DeletedRecordFlagSet_ExcludesThatRow()
    {
        var path = Path.Combine(_scratchFolder, "icmast.dbf");
        var fields = new[] { new FieldDef("REF", 11) };
        WriteFoxProDbf(path, versionByte: 0x30, fields, new[]
        {
            (false, Row(("REF", "SB-000001"))),
            (true, Row(("REF", "SB-000002"))),
        });

        var result = new FoxProDbfReader().Read(_scratchFolder, "icmast");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Single(result.Table.Rows);
        Assert.Equal("SB-000001", result.Table.Rows[0]["REF"]);
    }

    [Fact]
    public void Read_WithTypeFilter_ReturnsOnlyMatchingRows()
    {
        var path = Path.Combine(_scratchFolder, "icmast.dbf");
        var fields = new[] { new FieldDef("TYPE", 2), new FieldDef("REF", 11) };
        WriteFoxProDbf(path, versionByte: 0x30, fields, new[]
        {
            (false, Row(("TYPE", "IN"), ("REF", "SB-000001"))),
            (false, Row(("TYPE", "RE"), ("REF", "SB-000002"))),
        });

        var matching = new FoxProDbfReader().Read(_scratchFolder, "icmast", "TYPE", "RE");

        Assert.True(matching.Success, matching.ErrorMessage);
        Assert.Single(matching.Table.Rows);
        Assert.Equal("SB-000002", matching.Table.Rows[0]["REF"]);
    }

    [Fact]
    public void Read_EmptyTable_ReturnsSuccessWithNoRows()
    {
        var path = Path.Combine(_scratchFolder, "icmast.dbf");
        var fields = new[] { new FieldDef("REF", 11) };
        WriteFoxProDbf(path, versionByte: 0x30, fields, Array.Empty<(bool, IReadOnlyDictionary<string, string>)>());

        var result = new FoxProDbfReader().Read(_scratchFolder, "icmast");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.Empty(result.Table.Rows);
    }

    [Fact]
    public void Read_TableDoesNotExist_ReturnsFailureWithErrorMessage()
    {
        var result = new FoxProDbfReader().Read(_scratchFolder, "icmast");

        Assert.False(result.Success);
        Assert.Contains("icmast.dbf", result.ErrorMessage);
    }
}
