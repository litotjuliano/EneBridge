using System.Data;

namespace eneBridge.Wpf.Core.Models;

public sealed class StageReadResult
{
    public required DataTable Table { get; init; }

    /// <summary>Number of raw rows considered (candidates), including skipped ones.</summary>
    public required int RowsRead { get; init; }

    /// <summary>Number of rows actually added to <see cref="Table"/>.</summary>
    public int RowsWritten => Table.Rows.Count;

    public required IReadOnlyList<RowSkipReason> SkipReasons { get; init; }
}
