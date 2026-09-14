using System.Data;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Append-aware before/after-delta verification, shared by MainViewModel (Invoice) and
/// StockReceivedViewModel. Pure logic (no UI dependency) so it's directly unit-testable, unlike
/// the ACE-driver-crash-avoidance table check in eneBridge.Wpf.Services.DbfWorkflowHelper, which
/// needs MessageBox and stays in the Wpf project. See
/// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md.
/// </summary>
public static class DbfVerificationHelper
{
    /// <summary>
    /// Counts how many rows currently match `typeValue` in `tableName` — used as the "before"
    /// count for the append-aware delta verification below. Returns 0 without the caller needing
    /// to special-case a missing table: DbfReaderService.Read already returns a clean failure (and
    /// RowCount 0) without touching OleDb when the table's .dbf file doesn't exist yet, which is
    /// exactly the case on a genuine first-ever run.
    /// </summary>
    public static async Task<int> CountRowsByTypeAsync(
        DbfReaderService dbfReaderService, string dbfFolder, string tableName, string typeColumnName, string typeValue)
    {
        var result = await Task.Run(() => dbfReaderService.Read(dbfFolder, tableName, typeColumnName, typeValue));
        return result.Success ? result.RowCount : 0;
    }

    /// <summary>
    /// Reads the table back (filtered to `typeValue`) after an export, and compares how many new
    /// matching rows appeared against what the export reported writing — an append-aware delta
    /// check, since rows now accumulate across runs instead of the file being fully replaced each
    /// time (a plain "total rows == rows written" comparison would show a false mismatch on every
    /// run after the first). Never throws.
    /// </summary>
    public static async Task VerifyDbfDeltaAsync(
        DbfReaderService dbfReaderService,
        string tableName,
        string dbfFolder,
        string typeColumnName,
        string typeValue,
        int beforeCount,
        StageExportResult? exportResult,
        Action<DataView?> setPreview,
        Action<string> setStatusText,
        Action<bool> setVerified)
    {
        var readResult = await Task.Run(() => dbfReaderService.Read(dbfFolder, tableName, typeColumnName, typeValue));
        setPreview(readResult.Success ? readResult.Table.DefaultView : null);

        if (!readResult.Success)
        {
            setStatusText($"Could not verify: {readResult.ErrorMessage}");
            setVerified(false);
            return;
        }

        var expectedNewRows = exportResult?.RowsWritten ?? 0;
        var actualNewRows = readResult.RowCount - beforeCount;

        if ((exportResult?.Success ?? false) && expectedNewRows == actualNewRows)
        {
            setStatusText($"{tableName}: {actualNewRows} row(s) written and confirmed in DBF ✓ ({readResult.RowCount} total)");
            setVerified(true);
        }
        else if (exportResult?.Success ?? false)
        {
            var deltaDescription = actualNewRows < 0
                ? $"DBF shows {-actualNewRows} FEWER matching row(s) than before this export (rows may have been removed outside this tool)"
                : $"DBF shows {actualNewRows} new row(s)";
            setStatusText($"{tableName}: wrote {expectedNewRows} row(s) but {deltaDescription} ⚠ ({readResult.RowCount} total)");
            setVerified(false);
        }
        else
        {
            setStatusText($"{tableName}: export failed — DBF currently has {readResult.RowCount} row(s) (may include previous runs)");
            setVerified(false);
        }
    }
}
