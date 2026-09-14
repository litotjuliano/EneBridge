using System.Data;
using System.Data.OleDb;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Exports a DataTable to a DBF file via the OleDb/Access-Database-Engine dBASE IV driver
/// (no good managed alternative for DBF exists, so this piece keeps the original's approach).
/// </summary>
public sealed class DbfExportService
{
    private readonly FileLogger? _fileLogger;

    public DbfExportService(FileLogger? fileLogger = null)
    {
        _fileLogger = fileLogger;
    }

    /// <summary>
    /// Creates the DBF file if it doesn't exist yet, then inserts every row -- rows always
    /// accumulate into the live file across calls, never replacing what's already there (see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md for why: Invoice and
    /// Stock Received both write into the same icmaste.dbf/ictrane.dbf, distinguished by TYPE, so
    /// neither can be allowed to wipe out the other's rows). A failure opening the connection /
    /// creating the table is fatal for this stage (Success=false); a failure inserting a single
    /// row is recorded in RowErrors and the remaining rows still run.
    /// </summary>
    public StageExportResult Export(string dbfFolder, string tableName, DataTable data, IReadOnlyList<DbfColumnDefinition> schema)
    {
        var rowErrors = new List<string>();

        // Checkpoint logging around every OleDb call: a native AccessViolationException from the
        // ACE dBASE driver's COM interop bypasses every catch block below and kills the process
        // outright (see DbfExportServiceTests' Export_DbfFolderCannotBeCreated comment for the
        // mechanism). FileLogger writes and closes the file on every call, so these lines survive
        // even that kind of crash and let the last line logged pin down exactly which step died.
        try
        {
            Directory.CreateDirectory(dbfFolder);

            var tableExists = File.Exists(Path.Combine(dbfFolder, tableName + ".dbf"));

            var connectionString =
                $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbfFolder};Extended Properties=""dBASE IV;CollatingSequence=1252"";";

            _fileLogger?.LogInfo($"[{tableName}] Opening OleDb connection to '{dbfFolder}'");
            using var connection = new OleDbConnection(connectionString);
            connection.Open();
            _fileLogger?.LogInfo($"[{tableName}] Connection opened");

            if (!tableExists)
            {
                _fileLogger?.LogInfo($"[{tableName}] Table does not exist yet, running CREATE TABLE");
                var createTableSql = DbfSchemaBuilder.BuildCreateTableSql(tableName, schema);
                using (var createCommand = new OleDbCommand(createTableSql, connection))
                {
                    createCommand.ExecuteNonQuery();
                }
                _fileLogger?.LogInfo($"[{tableName}] CREATE TABLE succeeded");
            }

            _fileLogger?.LogInfo($"[{tableName}] Inserting {data.Rows.Count} row(s)");

            int rowsWritten = 0;
            int rowIndex = 0;
            foreach (DataRow row in data.Rows)
            {
                rowIndex++;
                try
                {
                    var insertSql = BuildInsertSql(tableName, data, row);
                    using var insertCommand = new OleDbCommand(insertSql, connection);
                    insertCommand.ExecuteNonQuery();
                    rowsWritten++;
                }
                catch (OleDbException ex)
                {
                    rowErrors.Add($"Row {rowIndex}: {ex.Message}");
                    _fileLogger?.LogInfo($"[{tableName}] Row {rowIndex} insert failed: {ex.Message}");
                }
            }
            _fileLogger?.LogInfo($"[{tableName}] Export complete: {rowsWritten} written, {rowErrors.Count} row error(s)");

            return new StageExportResult { Success = true, RowsWritten = rowsWritten, RowErrors = rowErrors };
        }
        catch (Exception ex) when (ex is OleDbException or IOException or UnauthorizedAccessException)
        {
            _fileLogger?.LogException($"[{tableName}] Export failed", ex);
            return new StageExportResult { Success = false, ErrorMessage = ex.Message, RowErrors = rowErrors };
        }
    }

    private static string BuildInsertSql(string tableName, DataTable table, DataRow row)
    {
        var columnNames = new List<string>();
        var values = new List<string>();

        foreach (DataColumn column in table.Columns)
        {
            columnNames.Add($"[{DbfSchemaBuilder.SanitizeColumnName(column.ColumnName)}]");
            values.Add(SqlValueFormatter.Format(row[column], column.DataType));
        }

        return $"INSERT INTO {tableName} ({string.Join(", ", columnNames)}) VALUES ({string.Join(", ", values)})";
    }
}
