using System.Data;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Detects rows about to be exported whose (REF, CODE) pair already exists in EMAS's live master
/// table (icmast.dbf, <see cref="IcmasteSchema.LiveTableName"/>) for the same TYPE — i.e. the same
/// document number for the same supplier/customer was already imported in an earlier run.
/// Deliberately checks icmast.dbf rather than icmaste.dbf (<see cref="IcmasteSchema.TableName"/>,
/// the staging file eneBridge itself writes and appends to forever): icmaste.dbf keeps growing
/// regardless of what happens inside EMAS afterward, so checking it would still report a
/// "duplicate" even after the user deletes that document from EMAS's own live data — confirmed by
/// the client hitting exactly this against a real installation. icmast.dbf reflects what EMAS
/// currently actually has, which is what re-export duplicate detection needs to answer.
/// DbfExportService.Export always appends now rather than recreating on collision, so nothing else
/// in the pipeline prevents re-exporting the same source file twice from silently duplicating every
/// row. This check itself never blocks anything — it only reports which REFs are duplicates; the
/// caller decides what to do (see eneBridge.Wpf.Services.DbfWorkflowHelper.ConfirmNoDuplicateDocuments,
/// which turns a non-empty result into a hard block with no override, since a Continue/Cancel
/// choice here would let a user click straight past the exact scenario this check exists to
/// prevent).
///
/// Two known gaps, not yet fixed (see CLAUDE.md's "Current status" section for the matching
/// project-level writeup):
/// 1. This only reads icmast.dbf, on the assumption that a REF+CODE match there means the whole
///    document — including its ictrane line items — was already imported. But
///    MainViewModel.ConfirmExportAsync/StockReceivedViewModel.ConfirmExportAsync run icmaste's and
///    ictrane's DbfExportService.Export calls as two independent stages (via the shared
///    ExportStageAsync helper) with no guard preventing the ictrane stage from running if the
///    icmaste stage failed. If a run writes ictrane's rows for a document but icmaste's row for
///    that same document fails to write, a later retry's check reads only icmast, correctly finds
///    no match, proceeds unwarned, and re-writes ictrane's already-present rows for that document.
/// 2. FindDuplicateRefsAsync below returns an empty list whenever DbfReaderService.Read's Success
///    is false, treating "icmast.dbf doesn't exist yet" (genuinely nothing to warn about) the
///    same as "the table exists but the OleDb SELECT itself failed" (e.g. a corrupt DBF, or the
///    ACE-driver native-crash-adjacent instability CLAUDE.md's "Known reliability risk" section
///    documents around DbfReaderService.Read). In the second case the check silently does nothing
///    and the export proceeds with no indication it didn't actually run.
///    DbfWorkflowHelper.ConfirmTablesReadable, which runs earlier in the same ConfirmExportAsync
///    flow, only does a plain-file-I/O readability probe and does not exercise the OleDb SELECT
///    path either, so it doesn't catch this failure mode.
/// </summary>
public static class DuplicateDocumentChecker
{
    /// <summary>
    /// Returns the distinct REF values among <paramref name="candidates"/> whose (Ref, Code) pair
    /// already exists in EMAS's live icmast.dbf for <paramref name="typeValue"/>, preserving
    /// first-seen order. Returns an empty list (not a failure) if icmast.dbf doesn't exist yet or
    /// can't currently be read — a first-ever export has nothing to duplicate against, and this
    /// check must never throw or block an otherwise-valid export on its own. Note that this same
    /// empty-list fallback also fires when the table exists but the underlying OleDb SELECT fails
    /// for some other reason (see the class-level doc comment's gap 2) — that case is
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
            dbfReaderService.Read(dbfFolder, IcmasteSchema.LiveTableName, IcmasteSchema.Type, typeValue));

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
