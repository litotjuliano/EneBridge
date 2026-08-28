namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Copies Excel files picked via Browse into a local staging folder
/// (%AppData%\eneBridge\ExcelSource\) so later runs don't depend on the original file's
/// (possibly removable/network) media staying available. Each workflow (Invoice, Stock Received,
/// ...) stages under its own fixed <paramref name="targetFileName"/> in the same shared folder.
/// </summary>
public sealed class ExcelSourceStagingService
{
    private readonly string _stagingDirectory;

    public ExcelSourceStagingService(string? userDataDirectory = null)
    {
        var dataDir = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "eneBridge");
        _stagingDirectory = Path.Combine(dataDir, "ExcelSource");
        Directory.CreateDirectory(_stagingDirectory);
    }

    public string StageFile(string sourcePath, string targetFileName)
    {
        var destinationPath = Path.Combine(_stagingDirectory, targetFileName);

        if (File.Exists(destinationPath))
        {
            var nameWithoutExtension = Path.GetFileNameWithoutExtension(targetFileName);
            var extension = Path.GetExtension(targetFileName);
            var backupPath = Path.Combine(
                _stagingDirectory, $"{nameWithoutExtension}_{DateTime.Now:yyyyMMddHHmmss}{extension}");
            File.Move(destinationPath, backupPath);
        }

        File.Copy(sourcePath, destinationPath, overwrite: false);
        return destinationPath;
    }
}
