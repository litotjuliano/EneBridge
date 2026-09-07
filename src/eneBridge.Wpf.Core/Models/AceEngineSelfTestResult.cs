namespace eneBridge.Wpf.Core.Models;

/// <summary>Outcome of AceEngineSelfTestService.RunSelfTest.</summary>
public sealed class AceEngineSelfTestResult
{
    public required bool Success { get; init; }
    public string? ErrorMessage { get; init; }
}
