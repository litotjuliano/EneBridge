using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class DbfSafetyBackupServiceTests : IDisposable
{
    private readonly string _scratchFolder;
    private readonly string _dbfFolder;
    private readonly string _userDataDirectory;

    public DbfSafetyBackupServiceTests()
    {
        _scratchFolder = Path.Combine(Path.GetTempPath(), "eneBridge.Wpf.Tests." + Guid.NewGuid().ToString("N"));
        _dbfFolder = Path.Combine(_scratchFolder, "dbf");
        _userDataDirectory = Path.Combine(_scratchFolder, "appdata");
        Directory.CreateDirectory(_dbfFolder);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_scratchFolder, recursive: true);
        }
        catch (IOException)
        {
            // best-effort cleanup
        }
    }

    [Fact]
    public void BackupIfExists_TableExists_CopiesToBackupFolderWithTimestampedName()
    {
        var sourcePath = Path.Combine(_dbfFolder, "icmaste.dbf");
        File.WriteAllText(sourcePath, "dbf content");
        var service = new DbfSafetyBackupService(userDataDirectory: _userDataDirectory);

        service.BackupIfExists(_dbfFolder, "icmaste");

        var backupFolder = Path.Combine(_userDataDirectory, "Backups", "icmaste");
        var backups = Directory.GetFiles(backupFolder, "icmaste_*.dbf");
        Assert.Single(backups);
        Assert.Equal("dbf content", File.ReadAllText(backups[0]));
    }

    [Fact]
    public void BackupIfExists_FptCompanionExists_CopiesItToo()
    {
        var sourceDbfPath = Path.Combine(_dbfFolder, "icmaste.dbf");
        File.WriteAllText(sourceDbfPath, "dbf content");
        var sourceFptPath = Path.Combine(_dbfFolder, "icmaste.FPT");
        File.WriteAllText(sourceFptPath, "memo content");
        var service = new DbfSafetyBackupService(userDataDirectory: _userDataDirectory);

        service.BackupIfExists(_dbfFolder, "icmaste");

        var backupFolder = Path.Combine(_userDataDirectory, "Backups", "icmaste");
        var fptBackups = Directory.GetFiles(backupFolder, "icmaste_*.FPT");
        Assert.Single(fptBackups);
        Assert.Equal("memo content", File.ReadAllText(fptBackups[0]));
    }

    [Fact]
    public void BackupIfExists_NoFptCompanion_OnlyCopiesDbf()
    {
        var sourceDbfPath = Path.Combine(_dbfFolder, "ictrane.dbf");
        File.WriteAllText(sourceDbfPath, "dbf content");
        var service = new DbfSafetyBackupService(userDataDirectory: _userDataDirectory);

        service.BackupIfExists(_dbfFolder, "ictrane");

        var backupFolder = Path.Combine(_userDataDirectory, "Backups", "ictrane");
        Assert.Single(Directory.GetFiles(backupFolder, "ictrane_*.dbf"));
        Assert.Empty(Directory.GetFiles(backupFolder, "ictrane_*.FPT"));
    }

    [Fact]
    public void BackupIfExists_TableDoesNotExist_DoesNothing()
    {
        var service = new DbfSafetyBackupService(userDataDirectory: _userDataDirectory);

        service.BackupIfExists(_dbfFolder, "icmaste");

        var backupFolder = Path.Combine(_userDataDirectory, "Backups", "icmaste");
        Assert.False(Directory.Exists(backupFolder));
    }

    [Fact]
    public void BackupIfExists_SourceFileLocked_DoesNotThrow()
    {
        var sourcePath = Path.Combine(_dbfFolder, "icmaste.dbf");
        File.WriteAllText(sourcePath, "dbf content");
        var service = new DbfSafetyBackupService(userDataDirectory: _userDataDirectory);

        using (var lockedStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            // Should not throw even though the file is locked
            service.BackupIfExists(_dbfFolder, "icmaste");
        }
    }
}
