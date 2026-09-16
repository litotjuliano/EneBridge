using System.IO;
using System.Windows;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Services;

/// <summary>
/// The ACE-driver-crash-avoidance table check, shared by MainViewModel (Invoice) and
/// StockReceivedViewModel, so it only exists in one place. Stays in the Wpf project (rather than
/// alongside eneBridge.Wpf.Core.Services.DbfVerificationHelper) because it needs
/// System.Windows.MessageBox to prompt the user; the append-aware before/after-delta
/// verification has no UI dependency and lives in Core instead, where it's directly
/// unit-testable. See docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md.
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
    /// Blocks the export — no override — unless <paramref name="result"/> confirms neither a
    /// duplicate was found nor the check itself failed to run. Deliberately offers no "continue
    /// anyway" choice in either failure case: a Yes/No prompt here would let a user click straight
    /// past the exact scenario this check exists to prevent, making the check pointless.
    ///
    /// Blocking on <see cref="DuplicateCheckResult.Verified"/> == false (not just on a populated
    /// duplicate list) is itself the fix for a confirmed real bug: icmast.dbf/ictran.dbf being
    /// briefly unreadable (e.g. locked because EMAS itself is open — an ordinary, common occurrence
    /// in this app's normal workflow, not a fluke) used to be silently treated the same as "checked,
    /// found nothing", letting a real duplicate document through with zero indication the check
    /// never actually ran. Now that case blocks too, telling the user why.
    /// </summary>
    public static bool ConfirmNoDuplicateDocuments(DuplicateCheckResult result)
    {
        if (!result.Verified)
        {
            MessageBox.Show(
                "Could not check EMAS's live data for duplicate documents (icmast.dbf/ictran.dbf " +
                $"exist but couldn't be read):\n{result.UnverifiableReason}\n\n" +
                "This usually means EMAS itself is currently open and holding these files. Close " +
                "EMAS and try again — export has been cancelled rather than proceeding without " +
                "this check.",
                "eneBridge - Duplicate Document Check",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);

            return false;
        }

        if (result.DuplicateRefs.Count == 0)
        {
            return true;
        }

        var refList = string.Join(", ", result.DuplicateRefs);
        MessageBox.Show(
            $"The following document number(s) already exist in EMAS for this " +
            $"supplier/customer: {refList}\n\n" +
            "This usually means this document was already imported into EMAS in an earlier " +
            "run. Export has been cancelled to prevent duplicating it.\n\n" +
            "If you believe this document was fully removed from EMAS, check EMAS directly and " +
            "remove any leftover record before exporting again.",
            "eneBridge - Duplicate Document Check",
            MessageBoxButton.OK,
            MessageBoxImage.Warning);

        return false;
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
}
