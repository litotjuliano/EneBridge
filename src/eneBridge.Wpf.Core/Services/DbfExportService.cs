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
    public void BackupExistingFile(string dbfFolder, string tableName)
    {
        var path = Path.Combine(dbfFolder, tableName + ".dbf");
        if (File.Exists(path))
        {
            var backupPath = Path.Combine(dbfFolder, $"{tableName}_{DateTime.Now:yyyyMMddHHmmss}.dbf");
            File.Move(path, backupPath);
        }
    }

    /// <summary>
    /// Creates (or re-creates, via backup-rename) the DBF file and inserts every row. A failure
    /// opening the connection / creating the table is fatal for this stage (Success=false); a
    /// failure inserting a single row is recorded in RowErrors and the remaining rows still run.
    /// </summary>
    public StageExportResult Export(string dbfFolder, string tableName, DataTable data, IReadOnlyList<DbfColumnDefinition> schema)
    {
        var rowErrors = new List<string>();

        try
        {
            Directory.CreateDirectory(dbfFolder);
            BackupExistingFile(dbfFolder, tableName);

            var connectionString =
                $@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={dbfFolder};Extended Properties=""dBASE IV;CollatingSequence=1252"";";

            using var connection = new OleDbConnection(connectionString);
            connection.Open();

            var createTableSql = DbfSchemaBuilder.BuildCreateTableSql(tableName, schema);
            using (var createCommand = new OleDbCommand(createTableSql, connection))
            {
                createCommand.ExecuteNonQuery();
            }

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
                }
            }

            return new StageExportResult { Success = true, RowsWritten = rowsWritten, RowErrors = rowErrors };
        }
        catch (Exception ex) when (ex is OleDbException or IOException or UnauthorizedAccessException)
        {
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
