using System.Text.RegularExpressions;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Builds the DBF-flavored SQL (via the OleDb/ACE dBASE IV driver) used to create the DBF tables.
/// </summary>
public static class DbfSchemaBuilder
{
    public static string BuildCreateTableSql(string tableName, IReadOnlyList<DbfColumnDefinition> columns)
    {
        var seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var columnDefs = new List<string>();

        foreach (var column in columns)
        {
            var sanitized = SanitizeColumnName(column.Name);
            if (!seenNames.Add(sanitized))
            {
                throw new InvalidOperationException(
                    $"Column name collision after DBF sanitization: '{column.Name}' -> '{sanitized}'.");
            }
            columnDefs.Add($"[{sanitized}] {GetDbfColumnType(column.ClrType, column.MaxLength)}");
        }

        return $"CREATE TABLE {tableName} ({string.Join(", ", columnDefs)})";
    }

    /// <summary>DBF field names are limited to 10 chars and [A-Za-z0-9_].</summary>
    public static string SanitizeColumnName(string name)
    {
        var truncated = name.Length > 10 ? name[..10] : name;
        return Regex.Replace(truncated, "[^a-zA-Z0-9_]", "");
    }

    public static string GetDbfColumnType(Type dataType, int maxLength)
    {
        if (dataType == typeof(string)) return $"CHAR({maxLength})";
        if (dataType == typeof(DateTime)) return "DATE";
        if (dataType == typeof(decimal)) return "NUMERIC(20, 2)";
        if (dataType == typeof(bool)) return "LOGICAL";
        throw new NotSupportedException($"Data type '{dataType.Name}' is not supported.");
    }
}
