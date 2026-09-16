using System.Data;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Post-export verification, shared by MainViewModel (Invoice) and StockReceivedViewModel. Pure
/// logic (no UI dependency) so it's directly unit-testable, unlike the ACE-driver-crash-avoidance
/// table check in eneBridge.Wpf.Services.DbfWorkflowHelper, which needs MessageBox and stays in the
/// Wpf project.
///
/// Compares the table's current TYPE-filtered row count against what the export reported writing,
/// rather than a before/after delta: DbfExportService.Export backs up and recreates the table on
/// every run (matching the original console app's proven behavior -- see its own doc comment for
/// the full history of why this changed and changed back), so after a successful export the table
/// contains exactly this run's rows for that TYPE, nothing else. A before/after delta was needed
/// only while Export appended forever; that's no longer how Export works.
/// </summary>
public static class DbfVerificationHelper
{
    public static async Task VerifyDbfAsync(
        DbfReaderService dbfReaderService,
        string tableName,
        string dbfFolder,
        string typeColumnName,
        string typeValue,
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

        var expectedRows = exportResult?.RowsWritten ?? 0;

        if ((exportResult?.Success ?? false) && expectedRows == readResult.RowCount)
        {
            setStatusText($"{tableName}: {readResult.RowCount} row(s) written and confirmed in DBF ✓");
            setVerified(true);
        }
        else if (exportResult?.Success ?? false)
        {
            setStatusText($"{tableName}: wrote {expectedRows} row(s) but DBF shows {readResult.RowCount} ⚠");
            setVerified(false);
        }
        else
        {
            setStatusText($"{tableName}: export failed — DBF currently has {readResult.RowCount} row(s)");
            setVerified(false);
        }
    }
}
