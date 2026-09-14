using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class RunHistoryServiceTests : IDisposable
{
    private readonly string _scratchFolder;

    public RunHistoryServiceTests()
    {
        _scratchFolder = Path.Combine(Path.GetTempPath(), "eneBridge.Wpf.Tests." + Guid.NewGuid().ToString("N"));
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

    private static RunHistoryEntry SampleEntry() => new() { ExcelPath = "x.xlsx", DbfFolder = @"C:\dbf" };

    [Fact]
    public void Append_DefaultFileName_WritesToRunHistoryJson()
    {
        var service = new RunHistoryService(_scratchFolder);

        service.Append(SampleEntry());

        Assert.True(File.Exists(Path.Combine(_scratchFolder, "runHistory.json")));
    }

    [Fact]
    public void Append_CustomFileName_WritesToThatFileNotTheDefault()
    {
        var service = new RunHistoryService(_scratchFolder, "stockReceivedRunHistory.json");

        service.Append(SampleEntry());

        Assert.True(File.Exists(Path.Combine(_scratchFolder, "stockReceivedRunHistory.json")));
        Assert.False(File.Exists(Path.Combine(_scratchFolder, "runHistory.json")));
    }

    [Fact]
    public void LoadAll_TwoServicesWithDifferentFileNames_DoNotSeeEachOthersEntries()
    {
        var invoiceService = new RunHistoryService(_scratchFolder);
        var stockReceivedService = new RunHistoryService(_scratchFolder, "stockReceivedRunHistory.json");

        invoiceService.Append(SampleEntry());

        Assert.Single(invoiceService.LoadAll());
        Assert.Empty(stockReceivedService.LoadAll());
    }
}
