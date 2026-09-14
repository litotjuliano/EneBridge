namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Copies icmaste.dbf/ictrane.dbf (and their .FPT memo companion, if present) to a separate,
/// app-controlled backup folder (%AppData%\eneBridge\Backups\&lt;table&gt;\) before export
/// touches anything. Deliberately independent of System.Data.OleDb/the ACE driver — a plain
/// File.Copy — so it still works even if the driver itself is broken (the exact scenario that
/// caused the incident this exists to protect against; see
/// docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md). Runs independently
/// of DbfExportService.Export, which no longer does any same-folder rename-on-collision backup of
/// its own -- Export now appends into the existing table instead of recreating it (see
/// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md).
/// </summary>
public sealed class DbfSafetyBackupService
{
    private readonly string _backupRootDirectory;
    private readonly FileLogger? _fileLogger;

    public DbfSafetyBackupService(FileLogger? fileLogger = null, string? userDataDirectory = null)
    {
        var dataDir = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "eneBridge");
        _backupRootDirectory = Path.Combine(dataDir, "Backups");
        _fileLogger = fileLogger;
    }

    /// <summary>
    /// No-op if &lt;dbfFolder&gt;\&lt;tableName&gt;.dbf doesn't currently exist (nothing to back
    /// up — e.g. a first-ever export). Never throws: a copy failure is logged and swallowed so it
    /// can never block a real export.
    /// </summary>
    public void BackupIfExists(string dbfFolder, string tableName)
    {
        var sourceDbfPath = Path.Combine(dbfFolder, tableName + ".dbf");
        if (!File.Exists(sourceDbfPath))
        {
            return;
        }

        try
        {
            var tableBackupDir = Path.Combine(_backupRootDirectory, tableName);
            Directory.CreateDirectory(tableBackupDir);

            var timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            var destDbfPath = Path.Combine(tableBackupDir, $"{tableName}_{timestamp}.dbf");
            File.Copy(sourceDbfPath, destDbfPath, overwrite: false);

            var sourceFptPath = Path.Combine(dbfFolder, tableName + ".FPT");
            if (File.Exists(sourceFptPath))
            {
                var destFptPath = Path.Combine(tableBackupDir, $"{tableName}_{timestamp}.FPT");
                File.Copy(sourceFptPath, destFptPath, overwrite: false);
            }

            _fileLogger?.LogInfo($"[{tableName}] Safety-backed-up to '{destDbfPath}'");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            _fileLogger?.LogException($"[{tableName}] Safety backup failed (continuing export)", ex);
        }
    }
}
