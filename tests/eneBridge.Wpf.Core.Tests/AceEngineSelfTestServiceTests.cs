using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

/// <summary>
/// Requires the ACE OleDb provider to be installed/registered on the machine running the tests —
/// same requirement as DbfExportServiceTests.
/// </summary>
public class AceEngineSelfTestServiceTests
{
    [Fact]
    public void RunSelfTest_AceProviderWorks_ReturnsSuccess()
    {
        var service = new AceEngineSelfTestService();

        var result = service.RunSelfTest();

        Assert.True(result.Success, result.ErrorMessage);
    }
}
