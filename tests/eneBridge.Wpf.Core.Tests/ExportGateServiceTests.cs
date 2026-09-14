using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class ExportGateServiceTests
{
    [Fact]
    public void TryBeginExport_WhenFree_ReturnsTrueAndClaimsGate()
    {
        var service = new ExportGateService();

        var result = service.TryBeginExport();

        Assert.True(result);
        Assert.True(service.IsExportInProgress);
    }

    [Fact]
    public void TryBeginExport_WhenAlreadyInProgress_ReturnsFalse()
    {
        var service = new ExportGateService();
        Assert.True(service.TryBeginExport());

        var result = service.TryBeginExport();

        Assert.False(result);
        Assert.True(service.IsExportInProgress);
    }

    [Fact]
    public void EndExport_ReleasesGate_AllowingNextTryBeginExport()
    {
        var service = new ExportGateService();
        Assert.True(service.TryBeginExport());

        service.EndExport();

        Assert.False(service.IsExportInProgress);
        Assert.True(service.TryBeginExport());
    }

    [Fact]
    public void TryBeginExport_RaisesStateChanged()
    {
        var service = new ExportGateService();
        var raised = false;
        service.StateChanged += () => raised = true;

        service.TryBeginExport();

        Assert.True(raised);
    }

    [Fact]
    public void EndExport_RaisesStateChanged()
    {
        var service = new ExportGateService();
        service.TryBeginExport();
        var raised = false;
        service.StateChanged += () => raised = true;

        service.EndExport();

        Assert.True(raised);
    }
}
