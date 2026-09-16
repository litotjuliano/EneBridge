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
    /// Backs up any existing &lt;tableName&gt;.dbf (rename to a timestamped copy in the same
    /// folder) and always creates the table fresh, then inserts this run's rows -- matching the
    /// original console app's proven behavior exactly (confirmed by decompiling it: it backed up
    /// and recreated on every run, and never accumulated rows across runs). An earlier iteration of
    /// this feature changed Export to always append instead, specifically so Invoice and Stock
    /// Received could share icmaste.dbf/ictrane.dbf without one wiping the other's rows -- but real
    /// testing found this caused a different, worse problem: since neither eneBridge nor (per
    /// observed behavior) EMAS's own import ever trims the staging file, every re-export of a
    /// document already sitting in ictrane.dbf left a second, duplicate set of line items once
    /// imported (confirmed against a real EMAS installation: deleting a document in EMAS, then
    /// re-exporting it, produced two identical line items instead of one, because the old line
    /// items were still physically sitting in ictrane.dbf and the new export only added to them).
    /// Reverted to backup-and-recreate on the client's explicit direction, accepting the known
    /// tradeoff: exporting one workflow (Invoice or Stock Received) now wipes the other's
    /// not-yet-imported rows from the staging file if the EMAS import hasn't run yet in between --
    /// safe only if each export is followed by an EMAS import before doing anything else, which
    /// matches how the single-workflow original app was always used. See CLAUDE.md's "Stock
    /// Received workflow" section for the full history. A failure opening the connection / creating
    /// the table is fatal for this stage (Success=false); a failure inserting a single row is
    /// recorded in RowErrors and the remaining rows still run.
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
            BackupExistingFile(dbfFolder, tableName);

            var connectionString =
                $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbfFolder};Extended Properties=""dBASE IV;CollatingSequence=1252"";";

            _fileLogger?.LogInfo($"[{tableName}] Opening OleDb connection to '{dbfFolder}'");
            using var connection = new OleDbConnection(connectionString);
            connection.Open();
            _fileLogger?.LogInfo($"[{tableName}] Connection opened");

            _fileLogger?.LogInfo($"[{tableName}] Running CREATE TABLE");
            var createTableSql = DbfSchemaBuilder.BuildCreateTableSql(tableName, schema);
            using (var createCommand = new OleDbCommand(createTableSql, connection))
            {
                createCommand.ExecuteNonQuery();
            }
            _fileLogger?.LogInfo($"[{tableName}] CREATE TABLE succeeded");

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

    /// <summary>
    /// Renames an existing &lt;tableName&gt;.dbf to a timestamped copy in the same folder, matching
    /// the original console app's exact naming convention (confirmed by decompiling it) --
    /// &lt;tableName&gt;_yyyyMMddHHmmss.dbf. No retention limit, same as the original: these backups
    /// are never deleted by this app. A no-op if the file doesn't exist yet (first-ever export).
    /// </summary>
    private static void BackupExistingFile(string dbfFolder, string tableName)
    {
        var dbfPath = Path.Combine(dbfFolder, tableName + ".dbf");
        if (!File.Exists(dbfPath))
        {
            return;
        }

        var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
        var backupPath = Path.Combine(dbfFolder, $"{tableName}_{timestamp}.dbf");
        File.Move(dbfPath, backupPath);
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
