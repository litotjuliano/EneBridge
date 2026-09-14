using System.Data;
using System.IO;
using System.Windows;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Services;

/// <summary>
/// Shared DBF safety/verification helpers used by both MainViewModel (Invoice) and
/// StockReceivedViewModel, so the ACE-driver-crash-avoidance table check and the append-aware
/// before/after-delta verification only exist in one place. See
/// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md.
/// </summary>
public static class DbfWorkflowHelper
{
    /// <summary>
    /// Warns (with a Continue/Cancel choice) if either table can't currently be read — covers
    /// "file doesn't exist", "table locked by EMAS", and "folder unreachable" alike. Deliberately
    /// checks via the filesystem only (File.Exists/Directory.Exists/a shared-read FileStream probe)
    /// rather than DbfReaderService.Read: a real OleDb SELECT against a nonexistent table was found
    /// to intermittently crash the process with the same native AccessViolationException/
    /// ComObject.Finalize signature this whole feature exists to guard against — reproduced even on
    /// a machine with an otherwise-healthy ACE driver. Using plain file I/O here removes OleDb from
    /// this code path entirely, matching DbfSafetyBackupService's same discipline. Checks both
    /// tables even if the first fails, unless the user cancels on the first warning.
    /// </summary>
    public static bool ConfirmTablesReadable(string dbfFolder)
    {
        foreach (var tableName in new[] { IcmasteSchema.TableName, IctraneSchema.TableName })
        {
            var (isReadable, errorMessage) = CheckTableFileReadable(dbfFolder, tableName);
            if (isReadable)
            {
                continue;
            }

            var proceed = MessageBox.Show(
                $"{tableName}.dbf could not be found or read in '{dbfFolder}':\n{errorMessage}\n\n" +
                "This is expected on a first-ever export to this folder, but if you expect this " +
                "table to already exist, something may be wrong (wrong folder selected, the table " +
                "is locked by EMAS, or a previous export was interrupted).\n\n" +
                "Continue anyway?",
                "eneBridge - Table Check",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (proceed != MessageBoxResult.Yes)
            {
                return false;
            }
        }

        return true;
    }

    /// <summary>
    /// Pure filesystem readability probe for one table's .dbf file — no OleDb involved. Opens with
    /// FileShare.ReadWrite (the most permissive request our own read can make) so this only reports
    /// "locked" when another process genuinely holds an incompatible lock, not merely because
    /// something else has the file open for reading too.
    /// </summary>
    private static (bool IsReadable, string? ErrorMessage) CheckTableFileReadable(string dbfFolder, string tableName)
    {
        if (!Directory.Exists(dbfFolder))
        {
            return (false, $"The folder '{dbfFolder}' does not exist or is not reachable.");
        }

        var dbfPath = Path.Combine(dbfFolder, tableName + ".dbf");
        if (!File.Exists(dbfPath))
        {
            return (false, $"'{tableName}.dbf' does not exist in this folder.");
        }

        try
        {
            using var stream = new FileStream(dbfPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return (true, null);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return (false, ex.Message);
        }
    }

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
            setStatusText($"{tableName}: wrote {expectedNewRows} row(s) but DBF shows {actualNewRows} new row(s) ⚠ ({readResult.RowCount} total)");
            setVerified(false);
        }
        else
        {
            setStatusText($"{tableName}: export failed — DBF currently has {readResult.RowCount} row(s) (may include previous runs)");
            setVerified(false);
        }
    }
}
