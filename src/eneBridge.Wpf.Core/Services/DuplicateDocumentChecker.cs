using System.Data;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Detects rows about to be exported into icmaste.dbf whose (REF, CODE) pair already exists there
/// for the same TYPE — i.e. the same document number for the same supplier/customer was already
/// imported in an earlier run. DbfExportService.Export always appends now rather than recreating
/// on collision, so nothing else in the pipeline prevents re-exporting the same source file twice
/// from silently duplicating every row. This check is display/warning-only: it never blocks
/// anything itself — callers decide what to do with the result (see
/// eneBridge.Wpf.Services.DbfWorkflowHelper.ConfirmNoDuplicateDocuments for the UI-coupled
/// Continue/Cancel prompt built on top of this).
/// </summary>
public static class DuplicateDocumentChecker
{
    /// <summary>
    /// Returns the distinct REF values among <paramref name="candidates"/> whose (Ref, Code) pair
    /// already exists in icmaste.dbf for <paramref name="typeValue"/>, preserving first-seen order.
    /// Returns an empty list (not a failure) if icmaste.dbf doesn't exist yet or can't currently be
    /// read — a first-ever export has nothing to duplicate against, and this check must never
    /// throw or block an otherwise-valid export on its own.
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
