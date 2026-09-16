using System.Data;
using System.Linq;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Whether <see cref="DuplicateDocumentChecker.FindDuplicateRefsAsync"/> was actually able to check
/// EMAS's live data. <see cref="Verified"/> is false when icmast.dbf/ictran.dbf exist but couldn't
/// be read (e.g. locked by EMAS itself being open) -- confirmed against a real installation as a
/// genuine, common occurrence, not a rare edge case: the whole eneBridge workflow revolves around
/// switching between it and EMAS, so EMAS being open (and holding these files) while a Confirm &amp;
/// Export runs is an ordinary sequence, not a fluke. Earlier code treated a read failure the same as
/// "table doesn't exist yet, nothing to duplicate against" and silently let the export proceed --
/// confirmed directly to let a real duplicate document through uncaught. <see cref="DuplicateRefs"/>
/// is only meaningful when <see cref="Verified"/> is true.
/// </summary>
public readonly record struct DuplicateCheckResult(bool Verified, IReadOnlyList<string> DuplicateRefs, string? UnverifiableReason = null);

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
/// This check itself never blocks anything — it only reports the result; the caller decides what to
/// do (see eneBridge.Wpf.Services.DbfWorkflowHelper.ConfirmNoDuplicateDocuments, which turns either
/// a non-empty duplicate list OR an unverifiable result into a hard block with no override, since a
/// Continue/Cancel choice here would let a user click straight past the exact scenario this check
/// exists to prevent).
/// </summary>
public static class DuplicateDocumentChecker
{
    /// <summary>
    /// Checks whether any of <paramref name="candidates"/>' (Ref, Code) pairs already exist in
    /// EMAS's live icmast.dbf AND have at least one matching ictran.dbf line item, for
    /// <paramref name="typeValue"/>. <see cref="DuplicateCheckResult.Verified"/> is true (with an
    /// empty or populated <see cref="DuplicateCheckResult.DuplicateRefs"/>) when both tables were
    /// successfully read, OR when neither exists yet (a genuine first-ever export has nothing to
    /// duplicate against — this check must never throw or block an otherwise-valid export on its
    /// own in that case). <see cref="DuplicateCheckResult.Verified"/> is false when a table exists
    /// but its read failed for some other reason (e.g. locked by EMAS being open) — the caller must
    /// NOT treat that the same as "no duplicates found", since it means the check didn't actually
    /// run.
    /// </summary>
    public static async Task<DuplicateCheckResult> FindDuplicateRefsAsync(
        FoxProDbfReader foxProDbfReader,
        string dbfFolder,
        string typeValue,
        IReadOnlyList<(string Ref, string Code)> candidates)
    {
        if (candidates.Count == 0)
        {
            return new DuplicateCheckResult(true, Array.Empty<string>());
        }

        var icmastPath = Path.Combine(dbfFolder, IcmasteSchema.LiveTableName + ".dbf");
        var ictranPath = Path.Combine(dbfFolder, IctraneSchema.LiveTableName + ".dbf");
        if (!File.Exists(icmastPath) || !File.Exists(ictranPath))
        {
            // Genuinely nothing to duplicate against yet (e.g. a fresh test folder). In a real EMAS
            // installation these shared master tables essentially always already exist, so this
            // branch is mostly relevant to testing, not production.
            return new DuplicateCheckResult(true, Array.Empty<string>());
        }

        var icmastResult = await Task.Run(() =>
            foxProDbfReader.Read(dbfFolder, IcmasteSchema.LiveTableName, IcmasteSchema.Type, typeValue));
        if (!icmastResult.Success)
        {
            return new DuplicateCheckResult(false, Array.Empty<string>(), icmastResult.ErrorMessage);
        }

        var ictranResult = await Task.Run(() =>
            foxProDbfReader.Read(dbfFolder, IctraneSchema.LiveTableName, IctraneSchema.Type, typeValue));
        if (!ictranResult.Success)
        {
            return new DuplicateCheckResult(false, Array.Empty<string>(), ictranResult.ErrorMessage);
        }

        var icmastKeys = new HashSet<(string Ref, string Code)>();
        foreach (DataRow row in icmastResult.Table.Rows)
        {
            var refValue = row[IcmasteSchema.Ref] as string ?? string.Empty;
            var codeValue = row[IcmasteSchema.Code] as string ?? string.Empty;
            icmastKeys.Add((refValue, codeValue));
        }

        var ictranRefs = ictranResult.Table.Rows.Cast<DataRow>()
            .Select(row => row[IctraneSchema.Ref] as string ?? string.Empty)
            .ToHashSet();

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
        return new DuplicateCheckResult(true, duplicates);
    }
}
