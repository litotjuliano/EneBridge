using System.Data;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Proves the ACE OleDb dBASE driver actually works by round-tripping a throwaway table through
/// the exact same DbfExportService/DbfReaderService code paths a real export uses, against a
/// disposable temp folder. Intended to run inside a short-lived child process (see
/// eneBridge.Wpf's AceEngineGuardService) so a native AccessViolationException from a broken
/// driver only kills that child, never the main app — see
/// docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md.
/// </summary>
public sealed class AceEngineSelfTestService
{
    private const string TableName = "selftest";
    private const string ColumnName = "TESTCOL";

    public AceEngineSelfTestResult RunSelfTest()
    {
        var tempFolder = Path.Combine(Path.GetTempPath(), "eneBridge-selftest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);

        try
        {
            var table = new DataTable();
            table.Columns.Add(ColumnName, typeof(string));
            table.Rows.Add("OK");

            var columns = new List<DbfColumnDefinition> { new(ColumnName, typeof(string), 10) };

            var exportResult = new DbfExportService().Export(tempFolder, TableName, table, columns);
            if (!exportResult.Success)
            {
                return new AceEngineSelfTestResult { Success = false, ErrorMessage = exportResult.ErrorMessage };
            }

            var readResult = new DbfReaderService().Read(tempFolder, TableName);
            if (!readResult.Success)
            {
                return new AceEngineSelfTestResult { Success = false, ErrorMessage = readResult.ErrorMessage };
            }

            if (readResult.RowCount != 1)
            {
                return new AceEngineSelfTestResult
                {
                    Success = false,
                    ErrorMessage = $"Expected 1 row after round-trip, found {readResult.RowCount}."
                };
            }

            return new AceEngineSelfTestResult { Success = true };
        }
        catch (Exception ex)
        {
            return new AceEngineSelfTestResult { Success = false, ErrorMessage = ex.Message };
        }
        finally
        {
            try
            {
                Directory.Delete(tempFolder, recursive: true);
            }
            catch (IOException)
            {
                // best-effort cleanup
            }
        }
    }
}
