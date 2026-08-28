using System.Security.AccessControl;
using System.Security.Principal;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class ExcelSourceStagingServiceTests : IDisposable
{
    private readonly string _scratchFolder;
    private readonly string _sourceFolder;
    private readonly string _userDataDirectory;

    public ExcelSourceStagingServiceTests()
    {
        _scratchFolder = Path.Combine(Path.GetTempPath(), "eneBridge.Wpf.Tests." + Guid.NewGuid().ToString("N"));
        _sourceFolder = Path.Combine(_scratchFolder, "source");
        _userDataDirectory = Path.Combine(_scratchFolder, "appdata");
        Directory.CreateDirectory(_sourceFolder);
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
    public void StageFile_NewFile_CopiesIntoStagingFolderAndReturnsStagedPath()
    {
        var sourcePath = Path.Combine(_sourceFolder, "source.xlsx");
        File.WriteAllText(sourcePath, "sample content");
        var service = new ExcelSourceStagingService(_userDataDirectory);

        var stagedPath = service.StageFile(sourcePath, "invoice.xlsx");

        Assert.Equal(Path.Combine(_userDataDirectory, "ExcelSource", "invoice.xlsx"), stagedPath);
        Assert.True(File.Exists(stagedPath));
        Assert.Equal("sample content", File.ReadAllText(stagedPath));
    }

    [Fact]
    public void StageFile_SourceHasDifferentName_StillStagedAsTargetFileName()
    {
        var sourcePath = Path.Combine(_sourceFolder, "January_Transactions.xlsx");
        File.WriteAllText(sourcePath, "sample content");
        var service = new ExcelSourceStagingService(_userDataDirectory);

        var stagedPath = service.StageFile(sourcePath, "invoice.xlsx");

        Assert.Equal(Path.Combine(_userDataDirectory, "ExcelSource", "invoice.xlsx"), stagedPath);
    }

    [Fact]
    public void StageFile_DifferentTargetFileNames_StageIndependently()
    {
        // Proves the staging slot is keyed by target file name, not a single hardcoded name --
        // needed since each workflow (Invoice, and later Stock Received) stages under its own
        // fixed name in the same shared ExcelSource folder.
        var invoiceSourcePath = Path.Combine(_sourceFolder, "invoice-source.xlsx");
        File.WriteAllText(invoiceSourcePath, "invoice content");
        var stockSourcePath = Path.Combine(_sourceFolder, "stock-source.xlsx");
        File.WriteAllText(stockSourcePath, "stock received content");
        var service = new ExcelSourceStagingService(_userDataDirectory);

        var invoiceStagedPath = service.StageFile(invoiceSourcePath, "invoice.xlsx");
        var stockStagedPath = service.StageFile(stockSourcePath, "stockreceived.xlsx");

        Assert.Equal(Path.Combine(_userDataDirectory, "ExcelSource", "invoice.xlsx"), invoiceStagedPath);
        Assert.Equal(Path.Combine(_userDataDirectory, "ExcelSource", "stockreceived.xlsx"), stockStagedPath);
        Assert.Equal("invoice content", File.ReadAllText(invoiceStagedPath));
        Assert.Equal("stock received content", File.ReadAllText(stockStagedPath));
    }

    [Fact]
    public void StageFile_TargetFileNameAlreadyStaged_RenamesOldCopyWithTimestampSuffixBeforeCopying()
    {
        var sourcePath = Path.Combine(_sourceFolder, "source.xlsx");
        File.WriteAllText(sourcePath, "version 1");
        var service = new ExcelSourceStagingService(_userDataDirectory);
        var firstStagedPath = service.StageFile(sourcePath, "invoice.xlsx");

        Thread.Sleep(1100); // ensure the backup timestamp (second-resolution) differs
        File.WriteAllText(sourcePath, "version 2");
        var secondStagedPath = service.StageFile(sourcePath, "invoice.xlsx");

        Assert.Equal(firstStagedPath, secondStagedPath);
        Assert.Equal("version 2", File.ReadAllText(secondStagedPath));

        var stagingFolder = Path.Combine(_userDataDirectory, "ExcelSource");
        var backups = Directory.GetFiles(stagingFolder, "invoice_*.xlsx");
        Assert.Single(backups);
        Assert.Equal("version 1", File.ReadAllText(backups[0]));
    }

    [Fact]
    public void StageFile_SourceFileMissing_ThrowsFileNotFoundException()
    {
        var service = new ExcelSourceStagingService(_userDataDirectory);
        var missingSourcePath = Path.Combine(_sourceFolder, "does-not-exist.xlsx");

        Assert.Throws<FileNotFoundException>(() => service.StageFile(missingSourcePath, "invoice.xlsx"));
    }

    [Fact]
    public void StageFile_StagingFolderNotWritable_ThrowsUnauthorizedAccessException()
    {
        var sourcePath = Path.Combine(_sourceFolder, "source.xlsx");
        File.WriteAllText(sourcePath, "sample content");
        var service = new ExcelSourceStagingService(_userDataDirectory); // creates ExcelSource\ eagerly

        var stagingFolder = Path.Combine(_userDataDirectory, "ExcelSource");
        var folderInfo = new DirectoryInfo(stagingFolder);
        var identity = WindowsIdentity.GetCurrent().User!;
        var denyRule = new FileSystemAccessRule(identity, FileSystemRights.Write, AccessControlType.Deny);

        var security = folderInfo.GetAccessControl();
        security.AddAccessRule(denyRule);
        folderInfo.SetAccessControl(security);

        try
        {
            Assert.Throws<UnauthorizedAccessException>(() => service.StageFile(sourcePath, "invoice.xlsx"));
        }
        finally
        {
            var cleanupSecurity = folderInfo.GetAccessControl();
            cleanupSecurity.RemoveAccessRule(denyRule);
            folderInfo.SetAccessControl(cleanupSecurity);
        }
    }
}
