using System.Data;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Detects rows about to be exported into icmaste.dbf whose (REF, CODE) pair already exists there
/// for the same TYPE — i.e. the same document number for the same supplier/customer was already
/// exported in an earlier run. DbfExportService.Export always appends now rather than recreating
/// on collision, so nothing else in the pipeline prevents re-exporting the same source file twice
/// from silently duplicating every row. This check itself never blocks anything — it only reports
/// which REFs are duplicates; the caller decides what to do (see
/// eneBridge.Wpf.Services.DbfWorkflowHelper.ConfirmNoDuplicateDocuments, which turns a non-empty
/// result into a hard block with no override, since a Continue/Cancel choice here would let a
/// user click straight past the exact scenario this check exists to prevent).
///
/// Checks icmaste.dbf (<see cref="IcmasteSchema.TableName"/>), NOT EMAS's live icmast.dbf (<see
/// cref="IcmasteSchema.LiveTableName"/>), even though icmast.dbf would be the semantically correct
/// thing to check (icmaste.dbf keeps growing regardless of what happens inside EMAS, so a document
/// deleted from EMAS's live data still shows up as a "duplicate" here — confirmed against a real
/// installation). Reading icmast.dbf was tried and reverted: its first byte (0x30) marks it as a
/// Visual FoxPro table, a materially different binary format from the plain dBASE III files
/// (0x03) this app writes -- Provider=Microsoft.ACE.OLEDB.12.0 with Extended Properties="dBASE IV"
/// cannot parse it at all, failing every single time with "External table is not in the expected
/// format" (not intermittent -- 100% reproducible), and attempting the read was strongly correlated
/// with the exact native-crash class CLAUDE.md's "Known reliability risk" section documents,
/// confirmed against a real EMAS installation. Properly reading a Visual FoxPro table needs either
/// the legacy VFPOLEDB provider (a separate, unsupported, uncertain-to-redistribute component) or a
/// hand-written FoxPro DBF parser -- neither attempted here; flagged for whoever picks this up next.
///
/// Known gap, not yet fixed (see CLAUDE.md's "Current status" section for the matching
/// project-level writeup): this only reads icmaste.dbf, on the assumption that a REF+CODE match
/// there means the whole document — including its ictrane line items — was already exported. But
/// MainViewModel.ConfirmExportAsync/StockReceivedViewModel.ConfirmExportAsync run icmaste's and
/// ictrane's DbfExportService.Export calls as two independent stages (via the shared
/// ExportStageAsync helper) with no guard preventing the ictrane stage from running if the icmaste
/// stage failed. If a run writes ictrane's rows for a document but icmaste's row for that same
/// document fails to write, a later retry's check reads only icmaste, correctly finds no match,
/// proceeds unwarned, and re-writes ictrane's already-present rows for that document.
/// </summary>
public static class DuplicateDocumentChecker
{
    /// <summary>
    /// Returns the distinct REF values among <paramref name="candidates"/> whose (Ref, Code) pair
    /// already exists in icmaste.dbf for <paramref name="typeValue"/>, preserving first-seen order.
    /// Returns an empty list (not a failure) if icmaste.dbf doesn't exist yet or can't currently be
    /// read — a first-ever export has nothing to duplicate against, and this check must never
    /// throw or block an otherwise-valid export on its own. Note that this same empty-list fallback
    /// also fires when the table exists but the underlying OleDb SELECT fails for some other
    /// reason (e.g. a corrupt DBF, or the ACE-driver native-crash-adjacent instability CLAUDE.md's
    /// "Known reliability risk" section documents around DbfReaderService.Read) — that case is
    /// indistinguishable from "no duplicates" to callers today.
    /// </summary>
    public static async Task<IReadOnlyList<string>> FindDuplicateRefsAsync(
        DbfReaderService dbfReaderService,
        string dbfFolder,
        string typeValue,
        IReadOnlyList<(string Ref, string Code)> candidates)
    {
        if (candidates.Count == 0)
        {
            return Array.Empty<string>();
        }

        var readResult = await Task.Run(() =>
            dbfReaderService.Read(dbfFolder, IcmasteSchema.TableName, IcmasteSchema.Type, typeValue));

        if (!readResult.Success)
        {
            return Array.Empty<string>();
        }

        var existingKeys = new HashSet<(string Ref, string Code)>();
        foreach (DataRow row in readResult.Table.Rows)
        {
            var refValue = row[IcmasteSchema.Ref] as string ?? string.Empty;
            var codeValue = row[IcmasteSchema.Code] as string ?? string.Empty;
            existingKeys.Add((refValue, codeValue));
        }

        var duplicates = new List<string>();
        foreach (var candidate in candidates)
        {
            if (existingKeys.Contains((candidate.Ref, candidate.Code)) && !duplicates.Contains(candidate.Ref))
            {
                duplicates.Add(candidate.Ref);
            }
        }
        return duplicates;
    }
}
