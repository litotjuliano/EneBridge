using System.Data;

namespace eneBridge.Wpf.Core.Models;

public sealed class DbfReadResult
{
    public required bool Success { get; init; }
    public DataTable Table { get; init; } = new();
    public string? ErrorMessage { get; init; }
    public int RowCount => Table.Rows.Count;
}
