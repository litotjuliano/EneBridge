namespace eneBridge.Wpf.Core.Models;

public sealed class RunHistoryEntry
{
    public DateTimeOffset Timestamp { get; init; } = DateTimeOffset.Now;
    public required string ExcelPath { get; init; }
    public required string DbfFolder { get; init; }

    public int IcmasteRowsRead { get; init; }
    public int IcmasteRowsSkipped { get; init; }
    public int IcmasteRowsWritten { get; init; }
    public bool IcmasteSuccess { get; init; }

    public int IctraneRowsRead { get; init; }
    public int IctraneRowsSkipped { get; init; }
    public int IctraneRowsWritten { get; init; }
    public bool IctraneSuccess { get; init; }

    public string? ErrorSummary { get; init; }
    public long DurationMs { get; init; }
}
