using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class SettingsServiceTests : IDisposable
{
    private readonly string _scratchFolder;
    private readonly string _appBaseDirectory;
    private readonly string _userDataDirectory;

    public SettingsServiceTests()
    {
        _scratchFolder = Path.Combine(Path.GetTempPath(), "eneBridge.Wpf.Tests." + Guid.NewGuid().ToString("N"));
        _appBaseDirectory = Path.Combine(_scratchFolder, "app");
        _userDataDirectory = Path.Combine(_scratchFolder, "appdata");
        Directory.CreateDirectory(_appBaseDirectory);
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
    public void SaveUserSettings_SavingInvoicePath_DoesNotClobberPreviouslySavedStockReceivedPath()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);

        service.SaveUserSettings(s => s.StockReceivedExcelFilePath = @"C:\stockreceived.xlsx");
        service.SaveUserSettings(s => s.ExcelFilePath = @"C:\invoice.xlsx");

        var saved = service.LoadUserSettings();
        Assert.Equal(@"C:\invoice.xlsx", saved!.ExcelFilePath);
        Assert.Equal(@"C:\stockreceived.xlsx", saved.StockReceivedExcelFilePath);
    }

    [Fact]
    public void SaveUserSettings_SavingStockReceivedPath_DoesNotClobberPreviouslySavedInvoicePath()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);

        service.SaveUserSettings(s => s.ExcelFilePath = @"C:\invoice.xlsx");
        service.SaveUserSettings(s => s.StockReceivedExcelFilePath = @"C:\stockreceived.xlsx");

        var saved = service.LoadUserSettings();
        Assert.Equal(@"C:\invoice.xlsx", saved!.ExcelFilePath);
        Assert.Equal(@"C:\stockreceived.xlsx", saved.StockReceivedExcelFilePath);
    }

    [Fact]
    public void ResolveEffectiveStockReceivedPaths_NoUserSettingsSaved_ExcelPathIsEmpty()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);

        var (excelPath, _) = service.ResolveEffectiveStockReceivedPaths(_appBaseDirectory);

        Assert.Equal(string.Empty, excelPath);
    }

    [Fact]
    public void ResolveEffectiveStockReceivedPaths_UsesSameDbfFolderAsInvoice()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);
        service.SaveUserSettings(s => s.DbfFolderPath = @"C:\emas\data");

        var (_, invoiceDbfFolder) = service.ResolveEffectivePaths(_appBaseDirectory);
        var (_, stockReceivedDbfFolder) = service.ResolveEffectiveStockReceivedPaths(_appBaseDirectory);

        Assert.Equal(@"C:\emas\data", invoiceDbfFolder);
        Assert.Equal(invoiceDbfFolder, stockReceivedDbfFolder);
    }

    [Fact]
    public void ResolveEffectiveStockReceivedPaths_UserSettingSaved_ExcelPathStillEmpty()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);
        service.SaveUserSettings(s => s.StockReceivedExcelFilePath = @"C:\my-stock-received.xlsx");

        var (excelPath, _) = service.ResolveEffectiveStockReceivedPaths(_appBaseDirectory);

        Assert.Equal(string.Empty, excelPath);
    }

    [Fact]
    public void ResolveEffectivePaths_NoUserSettingsSaved_ExcelPathIsEmpty()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);

        var (excelPath, _) = service.ResolveEffectivePaths(_appBaseDirectory);

        Assert.Equal(string.Empty, excelPath);
    }

    [Fact]
    public void ResolveEffectivePaths_UserSettingSaved_ExcelPathStillEmpty()
    {
        var service = new SettingsService(_appBaseDirectory, _userDataDirectory);
        service.SaveUserSettings(s => s.ExcelFilePath = @"C:\my-invoice.xlsx");

        var (excelPath, _) = service.ResolveEffectivePaths(_appBaseDirectory);

        Assert.Equal(string.Empty, excelPath);
    }
}
