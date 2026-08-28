using System.Data;
using System.Data.OleDb;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Reads back a table already written by <see cref="DbfExportService"/>, so the app can show
/// what actually ended up in the .dbf file rather than only what was intended to be written.
/// </summary>
public sealed class DbfReaderService
{
    private readonly FileLogger? _fileLogger;

    public DbfReaderService(FileLogger? fileLogger = null)
    {
        _fileLogger = fileLogger;
    }

    public DbfReadResult Read(string dbfFolder, string tableName)
    {
        try
        {
            var connectionString =
                $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbfFolder};Extended Properties=""dBASE IV;CollatingSequence=1252"";";

            _fileLogger?.LogInfo($"[{tableName}] (verify) Opening OleDb connection to '{dbfFolder}'");
            using var connection = new OleDbConnection(connectionString);
            connection.Open();
            _fileLogger?.LogInfo($"[{tableName}] (verify) Connection opened, reading rows");

            var table = new DataTable();
            using (var adapter = new OleDbDataAdapter($"SELECT * FROM {tableName}", connection))
            {
                adapter.Fill(table);
            }
            _fileLogger?.LogInfo($"[{tableName}] (verify) Read {table.Rows.Count} row(s)");

            return new DbfReadResult { Success = true, Table = table };
        }
        catch (Exception ex) when (ex is OleDbException or IOException or UnauthorizedAccessException)
        {
            _fileLogger?.LogException($"[{tableName}] (verify) Read failed", ex);
            return new DbfReadResult { Success = false, ErrorMessage = ex.Message };
        }
    }
}
