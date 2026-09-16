using System.Data;
using System.Text;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Read-only parser for Visual FoxPro .dbf tables (EMAS's own live icmast.dbf/ictran.dbf), which
/// use the classic dBASE free-table binary layout -- a 32-byte header, one 32-byte field
/// descriptor per column, a single 0x0D terminator, then fixed-width records each prefixed by a
/// 1-byte deletion flag ('*' deleted, ' ' active) -- confirmed byte-for-byte against a real
/// icmast.dbf during Stock Received testing. Only the file's version byte differs from the plain
/// dBASE III files (0x03) this app writes; icmast.dbf's is 0x30 (Visual FoxPro).
///
/// <c>Provider=Microsoft.ACE.OLEDB.12.0</c> with <c>Extended Properties="dBASE IV"</c>
/// (<see cref="DbfReaderService"/>, used for icmaste.dbf/ictrane.dbf) cannot open Visual FoxPro
/// tables at all -- every attempt threw <c>OleDbException: "External table is not in the expected
/// format"</c>, 100% reproducibly, and was strongly correlated with the native-crash class
/// CLAUDE.md's "Known reliability risk" section documents (confirmed against a real EMAS
/// installation: the app hung and closed on Confirm &amp; Export, twice in a row, immediately
/// after this read). The "correct" alternative, Microsoft's legacy VFPOLEDB provider, is 32-bit
/// only and cannot be loaded in-process by this app's 64-bit process without a bitness workaround
/// (a separate 32-bit helper process, or switching the whole app to 32-bit) -- not attempted here.
/// This reader sidesteps both problems: pure managed byte parsing, no COM/OleDb involved, so it
/// carries none of the ACE driver's native-crash risk and has no bitness constraint at all.
///
/// Every classic dBASE/FoxPro field type stores its value as plain text bytes -- numeric and date
/// fields included, right-aligned/padded with spaces -- so every column is decoded uniformly as
/// trimmed text; nothing here needs to know a field's declared type. (Memo/general/binary field
/// types store only a pointer into a companion .FPT file rather than inline text, but none of the
/// columns this app reads -- REF/CODE/TYPE -- are those types, so resolving .FPT content is out of
/// scope here.) Decoded with <see cref="Encoding.Latin1"/> rather than the Windows-1252 code page
/// the rest of this app's OleDb connection strings request (<c>CollatingSequence=1252</c>): the two
/// are identical across the plain-ASCII range these columns actually use, and Latin1 needs no extra
/// package (<c>System.Text.Encoding.CodePages</c>) or provider registration, keeping this reader's
/// whole point -- no new dependency, unlike VFPOLEDB -- fully intact.
/// </summary>
public sealed class FoxProDbfReader
{
    private const byte FieldDescriptorTerminator = 0x0D;
    private const byte DeletedRecordFlag = (byte)'*';
    private static readonly Encoding TextEncoding = Encoding.Latin1;

    private readonly FileLogger? _fileLogger;

    public FoxProDbfReader(FileLogger? fileLogger = null)
    {
        _fileLogger = fileLogger;
    }

    /// <summary>
    /// Reads every non-deleted record's columns into a DataTable of strings, optionally filtered
    /// to rows where <paramref name="typeColumnName"/> equals <paramref name="typeValue"/>. Mirrors
    /// <see cref="DbfReaderService.Read"/>'s <see cref="DbfReadResult"/> shape so callers can treat
    /// the two interchangeably. Never throws: any I/O or format problem is reported as
    /// Success=false rather than propagating, matching this project's "a duplicate check must never
    /// block an otherwise-valid export" requirement.
    /// </summary>
    public DbfReadResult Read(string dbfFolder, string tableName, string? typeColumnName = null, string? typeValue = null)
    {
        var dbfPath = Path.Combine(dbfFolder, tableName + ".dbf");
        if (!File.Exists(dbfPath))
        {
            return new DbfReadResult { Success = false, ErrorMessage = $"'{tableName}.dbf' does not exist in '{dbfFolder}'." };
        }

        try
        {
            using var stream = new FileStream(dbfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new BinaryReader(stream);

            stream.Seek(4, SeekOrigin.Begin); // skip version byte + 3-byte last-update date, neither needed
            uint recordCount = reader.ReadUInt32();
            ushort headerLength = reader.ReadUInt16();
            ushort recordLength = reader.ReadUInt16();

            stream.Seek(32, SeekOrigin.Begin);
            var fields = new List<(string Name, int Offset, int Length)>();
            int offset = 1; // every record starts with a 1-byte deletion flag
            while (stream.Position < headerLength - 1)
            {
                var descriptor = reader.ReadBytes(32);
                if (descriptor.Length < 32 || descriptor[0] == FieldDescriptorTerminator)
                {
                    break;
                }
                var name = TextEncoding.GetString(descriptor, 0, 11).TrimEnd('\0', ' ');
                int length = descriptor[16];
                fields.Add((name, offset, length));
                offset += length;
            }

            var table = new DataTable(tableName);
            foreach (var field in fields)
            {
                table.Columns.Add(field.Name, typeof(string));
            }

            stream.Seek(headerLength, SeekOrigin.Begin);
            for (uint i = 0; i < recordCount; i++)
            {
                var record = reader.ReadBytes(recordLength);
                if (record.Length < recordLength)
                {
                    break; // truncated/short final record -- stop rather than throw
                }
                if (record[0] == DeletedRecordFlag)
                {
                    continue;
                }

                var row = table.NewRow();
                foreach (var field in fields)
                {
                    row[field.Name] = TextEncoding.GetString(record, field.Offset, field.Length).Trim();
                }

                if (typeColumnName is null || (string)row[typeColumnName] == typeValue)
                {
                    table.Rows.Add(row);
                }
            }

            return new DbfReadResult { Success = true, Table = table };
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException or IndexOutOfRangeException)
        {
            _fileLogger?.LogException($"[{tableName}] FoxPro read failed", ex);
            return new DbfReadResult { Success = false, ErrorMessage = ex.Message };
        }
    }
}
