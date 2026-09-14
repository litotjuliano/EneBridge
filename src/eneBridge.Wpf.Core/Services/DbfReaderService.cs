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

    /// <summary>
    /// Reads a table, optionally filtered to rows where `typeColumnName` equals `typeValue` (used
    /// to separate Invoice's "IN" rows from Stock Received's "RE" rows in the shared tables -- see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md). Checks the table's
    /// .dbf file exists via plain file I/O before ever opening an OleDb connection: a real OleDb
    /// SELECT against a nonexistent table was found to intermittently crash the process with a
    /// native AccessViolationException/ComObject.Finalize signature, even on an otherwise-healthy
    /// driver, and callers that need a "before" row count (which can legitimately run against a
    /// table that doesn't exist yet on a genuine first-ever run) would hit that exact trigger
    /// without this guard.
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
            var connectionString =
                $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbfFolder};Extended Properties=""dBASE IV;CollatingSequence=1252"";";

            _fileLogger?.LogInfo($"[{tableName}] (verify) Opening OleDb connection to '{dbfFolder}'");
            using var connection = new OleDbConnection(connectionString);
            connection.Open();
            _fileLogger?.LogInfo($"[{tableName}] (verify) Connection opened, reading rows");

            var selectSql = $"SELECT * FROM {tableName}";
            if (!string.IsNullOrEmpty(typeColumnName))
            {
                selectSql += $" WHERE [{typeColumnName}] = {SqlValueFormatter.Format(typeValue, typeof(string))}";
            }

            var table = new DataTable();
            using (var adapter = new OleDbDataAdapter(selectSql, connection))
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
