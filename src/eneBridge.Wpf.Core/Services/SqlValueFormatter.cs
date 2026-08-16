using System.Globalization;

namespace eneBridge.Wpf.Core.Services;

/// <summary>Formats a DataRow value as a literal for the generated INSERT INTO statements.</summary>
public static class SqlValueFormatter
{
    public static string Format(object? value, Type clrType)
    {
        if (value is null || value == DBNull.Value)
        {
            return "NULL";
        }
        if (clrType == typeof(string))
        {
            return "'" + ((string)value).Replace("'", "''") + "'";
        }
        if (clrType == typeof(DateTime))
        {
            return "'" + ((DateTime)value).ToString("yyyy-MM-dd", CultureInfo.InvariantCulture) + "'";
        }
        if (clrType == typeof(bool))
        {
            return (bool)value ? ".T." : ".F.";
        }
        return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "NULL";
    }
}
