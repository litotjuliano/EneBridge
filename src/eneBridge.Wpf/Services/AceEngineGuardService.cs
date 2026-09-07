using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Services;

/// <summary>
/// Orchestrates the ACE-driver health check and repair flow around
/// eneBridge.Wpf.Core.Services.AceEngineSelfTestService, which does the actual round-trip check
/// but must run in a disposable child process — a broken driver can crash natively
/// (AccessViolationException), which bypasses every managed catch and would kill this app if run
/// in-process. See docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md.
/// </summary>
public sealed class AceEngineGuardService
{
    public const string SelfTestArg = "--selftest-ace";
    private static readonly TimeSpan SelfTestTimeout = TimeSpan.FromSeconds(15);

    private readonly FileLogger _fileLogger;
    private readonly string _appDirectory;

    public AceEngineGuardService(FileLogger fileLogger, string appDirectory)
    {
        _fileLogger = fileLogger;
        _appDirectory = appDirectory;
    }

    public static bool IsSelfTestRequest(string[] args) => args.Contains(SelfTestArg);

    /// <summary>
    /// Runs inside the child process itself: performs the self-test and exits the process with 0
    /// (healthy) or 1 (unhealthy/errored). Never returns normally.
    /// </summary>
    public static void RunSelfTestAndExit()
    {
        var result = new AceEngineSelfTestService().RunSelfTest();
        Environment.Exit(result.Success ? 0 : 1);
    }

    /// <summary>Runs in the main app process: launches the child process above and waits for it.</summary>
    public async Task<bool> CheckHealthAsync()
    {
        try
        {
            var exitCode = await RunSelfTestChildProcessAsync();
            _fileLogger.LogInfo($"ACE engine self-test child process exited with code {exitCode}");
            return exitCode == 0;
        }
        catch (Exception ex)
        {
            _fileLogger.LogException("ACE engine self-test failed to run", ex);
            return false;
        }
    }

    /// <summary>Re-runs the bundled Access Database Engine installer elevated, then re-checks health.</summary>
    public async Task<bool> TryRepairAsync()
    {
        var installerPath = Path.Combine(_appDirectory, "prerequisites", "accessdatabaseengine_X64.exe");
        if (!File.Exists(installerPath))
        {
            _fileLogger.LogInfo($"Repair skipped: installer not found at '{installerPath}'");
            return false;
        }

        try
        {
            var startInfo = new ProcessStartInfo
            {
                FileName = installerPath,
                Arguments = "/quiet /norestart",
                UseShellExecute = true,
                Verb = "runas"
            };
            using var process = Process.Start(startInfo);
            if (process is null)
            {
                return false;
            }
            await process.WaitForExitAsync();
            _fileLogger.LogInfo($"Repair installer exited with code {process.ExitCode}");
        }
        catch (Exception ex)
        {
            _fileLogger.LogException("Repair installer failed to run (user may have declined the UAC prompt)", ex);
            return false;
        }

        return await CheckHealthAsync();
    }

    private async Task<int> RunSelfTestChildProcessAsync()
    {
        var exePath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Could not determine the current process path.");

        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = SelfTestArg,
            UseShellExecute = false
        };

        using var process = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Failed to start the self-test child process.");

        using var cts = new CancellationTokenSource(SelfTestTimeout);
        try
        {
            await process.WaitForExitAsync(cts.Token);
        }
        catch (OperationCanceledException)
        {
            _fileLogger.LogInfo("ACE engine self-test child process timed out; treating as unhealthy.");
            try { process.Kill(); } catch (InvalidOperationException) { /* already exited */ }
            return -1;
        }

        return process.ExitCode;
    }
}
