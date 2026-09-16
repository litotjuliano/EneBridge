using System.Data;
using System.Linq;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Detects rows about to be exported whose (REF, CODE) pair already exists in EMAS's live master
/// table (icmast.dbf, <see cref="IcmasteSchema.LiveTableName"/>) for the same TYPE — i.e. the same
/// document number for the same supplier/customer was already imported in an earlier run. Checks
/// icmast.dbf rather than icmaste.dbf (<see cref="IcmasteSchema.TableName"/>, the staging file
/// eneBridge itself writes and appends to forever): icmaste.dbf keeps growing regardless of what
/// happens inside EMAS afterward, so checking it would still report a "duplicate" even after the
/// user deletes that document from EMAS's own live data — confirmed by the client hitting exactly
/// this against a real installation. icmast.dbf reflects what EMAS currently actually has, which is
/// what re-export duplicate detection needs to answer.
///
/// A REF is only treated as a genuine duplicate if BOTH its icmast.dbf header row AND at least one
/// matching ictran.dbf (<see cref="IctraneSchema.LiveTableName"/>) line item still exist. This was
/// added after real testing found EMAS's own delete leaves an orphaned icmast.dbf header behind
/// after removing a document's ictran.dbf line items (a "soft delete", per the client) — an
/// icmast-only check saw the leftover header and incorrectly blocked a legitimate re-export of a
/// document EMAS's own Stock Received screen could no longer even find.
///
/// Reads via <see cref="FoxProDbfReader"/>, NOT <see cref="DbfReaderService"/>/OleDb: icmast.dbf's
/// first byte (0x30) marks it as a Visual FoxPro table, a materially different binary format from
/// the plain dBASE III files (0x03) this app writes.
/// <c>Provider=Microsoft.ACE.OLEDB.12.0</c> with <c>Extended Properties="dBASE IV"</c> cannot parse
/// Visual FoxPro tables at all — every attempt failed with "External table is not in the expected
/// format" (100% reproducible, not intermittent) and was strongly correlated with the native-crash
/// class CLAUDE.md's "Known reliability risk" section documents, confirmed against a real EMAS
/// installation. <see cref="FoxProDbfReader"/> sidesteps this entirely with pure managed byte
/// parsing (no COM/OleDb), confirmed correct against the real icmast.dbf during Stock Received
/// testing — see its own class doc comment for the full story and the VFPOLEDB alternative that
/// was considered and rejected (32-bit only, incompatible with this app's 64-bit process).
///
/// This check itself never blocks anything — it only reports which REFs are duplicates; the caller
/// decides what to do (see eneBridge.Wpf.Services.DbfWorkflowHelper.ConfirmNoDuplicateDocuments,
/// which turns a non-empty result into a hard block with no override, since a Continue/Cancel
/// choice here would let a user click straight past the exact scenario this check exists to
/// prevent).
/// </summary>
public static class DuplicateDocumentChecker
{
    /// <summary>
    /// Returns the distinct REF values among <paramref name="candidates"/> whose (Ref, Code) pair
    /// already exists in EMAS's live icmast.dbf AND has at least one matching ictran.dbf line item,
    /// for <paramref name="typeValue"/>, preserving first-seen order. Returns an empty list (not a
    /// failure) if either table doesn't exist yet or can't currently be read — a first-ever export
    /// has nothing to duplicate against, and this check must never throw or block an otherwise-valid
    /// export on its own. Note that this same empty-list-for-that-table behavior also fires when a
    /// table exists but the underlying read fails for some other reason (e.g. a structurally corrupt
    /// or unexpectedly-shaped DBF) — that case is indistinguishable from "nothing there" to callers
    /// today.
    /// </summary>
    public static async Task<IReadOnlyList<string>> FindDuplicateRefsAsync(
        FoxProDbfReader foxProDbfReader,
        string dbfFolder,
        string typeValue,
        IReadOnlyList<(string Ref, string Code)> candidates)
    {
        if (candidates.Count == 0)
        {
            return Array.Empty<string>();
        }

        var icmastResult = await Task.Run(() =>
            foxProDbfReader.Read(dbfFolder, IcmasteSchema.LiveTableName, IcmasteSchema.Type, typeValue));

        if (!icmastResult.Success)
        {
            return Array.Empty<string>();
        }

        var icmastKeys = new HashSet<(string Ref, string Code)>();
        foreach (DataRow row in icmastResult.Table.Rows)
        {
            var refValue = row[IcmasteSchema.Ref] as string ?? string.Empty;
            var codeValue = row[IcmasteSchema.Code] as string ?? string.Empty;
            icmastKeys.Add((refValue, codeValue));
        }

        var ictranResult = await Task.Run(() =>
            foxProDbfReader.Read(dbfFolder, IctraneSchema.LiveTableName, IctraneSchema.Type, typeValue));

        var ictranRefs = ictranResult.Success
            ? ictranResult.Table.Rows.Cast<DataRow>()
                .Select(row => row[IctraneSchema.Ref] as string ?? string.Empty)
                .ToHashSet()
            : new HashSet<string>();

        var duplicates = new List<string>();
        foreach (var candidate in candidates)
        {
            if (icmastKeys.Contains((candidate.Ref, candidate.Code))
                && ictranRefs.Contains(candidate.Ref)
                && !duplicates.Contains(candidate.Ref))
            {
                duplicates.Add(candidate.Ref);
            }
        }
        return duplicates;
    }
}
