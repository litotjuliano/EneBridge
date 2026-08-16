namespace eneBridge.Wpf.Core.Models;

public sealed class StageExportResult
{
    public required bool Success { get; init; }
    public int RowsWritten { get; init; }
    public string? ErrorMessage { get; init; }
    public IReadOnlyList<string> RowErrors { get; init; } = [];
}
