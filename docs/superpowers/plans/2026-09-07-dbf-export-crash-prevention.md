# DBF Export Crash Prevention Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Implement the four-layer defense from
`docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md` against the ACE-driver
crash class that took down Confirm & Export: a safety backup of both tables to an app-controlled
folder, an installer fix so the correct driver is actually installed on Office machines, a
startup self-test that isolates the native-crash risk in a disposable child process, and a
pre-export check that warns if either table is missing/unreadable before anything irreversible
happens.

**Architecture:** Two new `eneBridge.Wpf.Core` services (`DbfSafetyBackupService`,
`AceEngineSelfTestService`) hold the testable logic; a new `eneBridge.Wpf`-side
`AceEngineGuardService` orchestrates the disposable child-process self-test and the elevated
repair flow; `MainViewModel`/`App.xaml.cs` wire it all together; `installer/eneBridge.iss` gets a
corrected detection function plus a permanent (not temp-deleted) copy of the redistributable.

**Tech Stack:** .NET 8, WPF, `System.Data.OleDb` (ACE provider), CommunityToolkit.Mvvm, xUnit,
Inno Setup 6 (Pascal scripting).

---

### Task 1: `DbfSafetyBackupService` (Core, TDD)

Plain-filesystem copy of `<table>.dbf` (+ `.FPT` companion if present) to
`%AppData%\eneBridge\Backups\<table>\<table>_<timestamp>.dbf`, independent of OleDb entirely.

**Files:**
- Create: `src/eneBridge.Wpf.Core/Services/DbfSafetyBackupService.cs`
- Test: `tests/eneBridge.Wpf.Core.Tests/DbfSafetyBackupServiceTests.cs`

- [ ] **Step 1: Write the failing tests**

```csharp
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

        using var lockedStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.None);

        var exception = Record.Exception(() => service.BackupIfExists(_dbfFolder, "icmaste"));

        Assert.Null(exception);
    }
}
```

- [ ] **Step 2: Run tests to verify they fail (compile error — class doesn't exist yet)**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~DbfSafetyBackupServiceTests"`
Expected: build FAILS — `DbfSafetyBackupService` does not exist.

- [ ] **Step 3: Write the implementation**

```csharp
namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Copies icmaste.dbf/ictrane.dbf (and their .FPT memo companion, if present) to a separate,
/// app-controlled backup folder (%AppData%\eneBridge\Backups\&lt;table&gt;\) before export
/// touches anything. Deliberately independent of System.Data.OleDb/the ACE driver — a plain
/// File.Copy — so it still works even if the driver itself is broken (the exact scenario that
/// caused the incident this exists to protect against; see
/// docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md). Runs in addition to,
/// not instead of, DbfExportService.BackupExistingFile's own same-folder rename-on-collision
/// backup, which is still required since CREATE TABLE needs the filename slot free.
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
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~DbfSafetyBackupServiceTests"`
Expected: `Passed! - Failed: 0, Passed: 5, Skipped: 0, Total: 5`

- [ ] **Step 5: Commit**

```bash
git add src/eneBridge.Wpf.Core/Services/DbfSafetyBackupService.cs tests/eneBridge.Wpf.Core.Tests/DbfSafetyBackupServiceTests.cs
git commit -m "Add DbfSafetyBackupService for app-controlled DBF table backups

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 2: `AceEngineSelfTestResult` + `AceEngineSelfTestService` (Core, TDD)

The round-trip logic that proves the ACE driver actually works, using the exact same
`DbfExportService`/`DbfReaderService` code paths a real export uses, against a disposable temp
folder. Designed to be run inside a short-lived child process (Task 3/4) so a native
`AccessViolationException` from a broken driver only kills that child.

**Files:**
- Create: `src/eneBridge.Wpf.Core/Models/AceEngineSelfTestResult.cs`
- Create: `src/eneBridge.Wpf.Core/Services/AceEngineSelfTestService.cs`
- Test: `tests/eneBridge.Wpf.Core.Tests/AceEngineSelfTestServiceTests.cs`

- [ ] **Step 1: Write the failing test**

```csharp
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
```

- [ ] **Step 2: Run test to verify it fails (compile error)**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~AceEngineSelfTestServiceTests"`
Expected: build FAILS — `AceEngineSelfTestService` does not exist.

- [ ] **Step 3: Write the result model**

```csharp
namespace eneBridge.Wpf.Core.Models;

/// <summary>Outcome of AceEngineSelfTestService.RunSelfTest.</summary>
public sealed class AceEngineSelfTestResult
{
    public required bool Success { get; init; }
    public string? ErrorMessage { get; init; }
}
```

- [ ] **Step 4: Write the self-test service**

```csharp
using System.Data;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Proves the ACE OleDb dBASE driver actually works by round-tripping a throwaway table through
/// the exact same DbfExportService/DbfReaderService code paths a real export uses, against a
/// disposable temp folder. Intended to run inside a short-lived child process (see
/// eneBridge.Wpf's AceEngineGuardService) so a native AccessViolationException from a broken
/// driver only kills that child, never the main app — see
/// docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md.
/// </summary>
public sealed class AceEngineSelfTestService
{
    private const string TableName = "selftest";
    private const string ColumnName = "TESTCOL";

    public AceEngineSelfTestResult RunSelfTest()
    {
        var tempFolder = Path.Combine(Path.GetTempPath(), "eneBridge-selftest-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(tempFolder);

        try
        {
            var table = new DataTable();
            table.Columns.Add(ColumnName, typeof(string));
            table.Rows.Add("OK");

            var columns = new List<DbfColumnDefinition> { new(ColumnName, typeof(string), 10) };

            var exportResult = new DbfExportService().Export(tempFolder, TableName, table, columns);
            if (!exportResult.Success)
            {
                return new AceEngineSelfTestResult { Success = false, ErrorMessage = exportResult.ErrorMessage };
            }

            var readResult = new DbfReaderService().Read(tempFolder, TableName);
            if (!readResult.Success)
            {
                return new AceEngineSelfTestResult { Success = false, ErrorMessage = readResult.ErrorMessage };
            }

            if (readResult.RowCount != 1)
            {
                return new AceEngineSelfTestResult
                {
                    Success = false,
                    ErrorMessage = $"Expected 1 row after round-trip, found {readResult.RowCount}."
                };
            }

            return new AceEngineSelfTestResult { Success = true };
        }
        catch (Exception ex)
        {
            return new AceEngineSelfTestResult { Success = false, ErrorMessage = ex.Message };
        }
        finally
        {
            try
            {
                Directory.Delete(tempFolder, recursive: true);
            }
            catch (IOException)
            {
                // best-effort cleanup
            }
        }
    }
}
```

- [ ] **Step 5: Run test to verify it passes**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj" --filter "FullyQualifiedName~AceEngineSelfTestServiceTests"`
Expected: `Passed! - Failed: 0, Passed: 1, Skipped: 0, Total: 1`

- [ ] **Step 6: Commit**

```bash
git add src/eneBridge.Wpf.Core/Models/AceEngineSelfTestResult.cs src/eneBridge.Wpf.Core/Services/AceEngineSelfTestService.cs tests/eneBridge.Wpf.Core.Tests/AceEngineSelfTestServiceTests.cs
git commit -m "Add AceEngineSelfTestService for a real DBF round-trip health check

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 3: `AceEngineGuardService` (Wpf, child-process orchestration)

Spawns `eneBridge.Wpf.exe --selftest-ace` as a child process and interprets its exit code (isolates
the native-crash risk); handles the elevated "Fix it now" repair flow. Not meaningfully
unit-testable (cross-process/OS-level) — verified manually in this task's last step.

**Files:**
- Create: `src/eneBridge.Wpf/Services/AceEngineGuardService.cs`

- [ ] **Step 1: Write the implementation**

```csharp
using System.Diagnostics;
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
```

- [ ] **Step 2: Build to verify it compiles**

Run: `dotnet build "src/eneBridge.Wpf/eneBridge.Wpf.csproj"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 3: Commit**

```bash
git add src/eneBridge.Wpf/Services/AceEngineGuardService.cs
git commit -m "Add AceEngineGuardService to isolate the ACE driver health check in a child process

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

(Manual verification of the child-process flow happens at the end of Task 5, once `App.xaml.cs`
and `MainViewModel` both wire up `--selftest-ace` and compile together.)

---

### Task 4: Wire the self-test into `App.xaml.cs`'s composition root

Before any window/service is created, check for `--selftest-ace` and short-circuit into the
disposable self-test path. On a normal launch, construct the new services and kick off the health
check after the window is shown.

**Files:**
- Modify: `src/eneBridge.Wpf/App.xaml.cs`

- [ ] **Step 1: Update `App.xaml.cs`**

Replace the full file with:

```csharp
using System.Windows;
using eneBridge.Wpf.Core.Services;
using eneBridge.Wpf.Services;
using eneBridge.Wpf.ViewModels;

namespace eneBridge.Wpf;

/// <summary>
/// Interaction logic for App.xaml. Composition root: constructs the services and the
/// MainViewModel by hand (no DI container needed at this app's size).
/// </summary>
public partial class App : Application
{
    protected override void OnStartup(StartupEventArgs e)
    {
        base.OnStartup(e);

        if (AceEngineGuardService.IsSelfTestRequest(e.Args))
        {
            // Disposable child-process mode (see AceEngineGuardService): perform the ACE-driver
            // round-trip check and exit immediately, with no window and no normal startup path.
            // A native AccessViolationException from a broken driver kills only this process.
            AceEngineGuardService.RunSelfTestAndExit();
            return;
        }

        var fileLogger = new FileLogger();
        RegisterGlobalExceptionHandlers(fileLogger);

        var excelReaderService = new ExcelReaderService();
        var dbfExportService = new DbfExportService(fileLogger);
        var dbfReaderService = new DbfReaderService(fileLogger);
        var dbfSafetyBackupService = new DbfSafetyBackupService(fileLogger);
        var settingsService = new SettingsService(AppContext.BaseDirectory);
        var runHistoryService = new RunHistoryService();
        var excelSourceStagingService = new ExcelSourceStagingService();
        var aceEngineGuardService = new AceEngineGuardService(fileLogger, AppContext.BaseDirectory);

        var mainViewModel = new MainViewModel(
            excelReaderService,
            dbfExportService,
            dbfReaderService,
            dbfSafetyBackupService,
            settingsService,
            runHistoryService,
            fileLogger,
            excelSourceStagingService,
            aceEngineGuardService,
            AppContext.BaseDirectory);
        mainViewModel.Initialize();

        var mainWindow = new MainWindow { DataContext = mainViewModel };
        mainWindow.Show();

        // Fire-and-forget: runs once per session, doesn't block the window from appearing.
        // AceEngineGuardService.CheckHealthAsync never throws, so this can't produce an
        // unobserved task exception.
        _ = mainViewModel.RunAceEngineHealthCheckAsync();
    }

    /// <summary>
    /// Best-effort safety net: logs anything catchable before it takes down the app, and keeps
    /// the app alive for exceptions on the UI thread (matches this project's "nothing should ever
    /// crash the app" goal). Cannot catch a native AccessViolationException from the ACE OleDb
    /// driver (see DbfExportService) — that class of fault is fatal by design in modern .NET and
    /// bypasses all of these handlers, which is why DbfExportService logs its own checkpoints and
    /// why the startup self-test (AceEngineGuardService) runs in its own disposable process.
    /// </summary>
    private static void RegisterGlobalExceptionHandlers(FileLogger fileLogger)
    {
        Current.DispatcherUnhandledException += (_, e) =>
        {
            fileLogger.LogException("Unhandled exception on UI thread", e.Exception);
            MessageBox.Show(
                $"An unexpected error occurred and was logged:\n\n{e.Exception.Message}",
                "eneBridge - Unexpected Error",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            e.Handled = true;
        };

        AppDomain.CurrentDomain.UnhandledException += (_, e) =>
        {
            if (e.ExceptionObject is Exception ex)
            {
                fileLogger.LogException($"Unhandled exception (IsTerminating={e.IsTerminating})", ex);
            }
        };

        TaskScheduler.UnobservedTaskException += (_, e) =>
        {
            fileLogger.LogException("Unobserved task exception", e.Exception);
            e.SetObserved();
        };
    }
}
```

This references `mainViewModel.RunAceEngineHealthCheckAsync()` and a new `MainViewModel`
constructor shape that don't exist yet — that's expected; Task 5 adds them. The build in this
task's Step 2 will fail until Task 5 is done, which is fine since Tasks 4 and 5 are verified
together at the end of Task 5.

- [ ] **Step 2: Commit (build will be verified at the end of Task 5)**

```bash
git add src/eneBridge.Wpf/App.xaml.cs
git commit -m "Wire the ACE-driver self-test into App.xaml.cs's startup

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 5: `MainViewModel` — health property, gating, and the repair flow

Adds the new constructor dependencies, the `AceEngineHealthy`/`AceEngineStatusText` bindable
properties, gates `ConfirmExportCommand` on health, and adds the health-check/repair orchestration
methods `App.xaml.cs` (Task 4) already calls.

**Files:**
- Modify: `src/eneBridge.Wpf/ViewModels/MainViewModel.cs`

- [ ] **Step 1: Add the `using` and new fields**

In `src/eneBridge.Wpf.csproj`'s `MainViewModel.cs`, add to the top `using` block:

```csharp
using System.Windows;
```

so the full using block reads:

```csharp
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;
using eneBridge.Wpf.Services;
```

- [ ] **Step 2: Add the new fields, next to the existing ones**

Replace:

```csharp
    private readonly ExcelReaderService _excelReaderService;
    private readonly DbfExportService _dbfExportService;
    private readonly DbfReaderService _dbfReaderService;
    private readonly SettingsService _settingsService;
    private readonly RunHistoryService _runHistoryService;
    private readonly FileLogger _fileLogger;
    private readonly ExcelSourceStagingService _excelSourceStagingService;
    private readonly string _appBaseDirectory;
```

with:

```csharp
    private readonly ExcelReaderService _excelReaderService;
    private readonly DbfExportService _dbfExportService;
    private readonly DbfReaderService _dbfReaderService;
    private readonly DbfSafetyBackupService _dbfSafetyBackupService;
    private readonly SettingsService _settingsService;
    private readonly RunHistoryService _runHistoryService;
    private readonly FileLogger _fileLogger;
    private readonly ExcelSourceStagingService _excelSourceStagingService;
    private readonly AceEngineGuardService _aceEngineGuardService;
    private readonly string _appBaseDirectory;
```

- [ ] **Step 3: Add the new observable properties, next to `_showAllColumns`**

Add after the existing `_showAllColumns` property block:

```csharp
    /// <summary>
    /// Set once per session by RunAceEngineHealthCheckAsync. Starts true (optimistic) so the app
    /// isn't blocked while the startup check runs; ConfirmExportCommand is gated on this so a
    /// broken ACE driver disables Confirm & Export instead of crashing it.
    /// </summary>
    [ObservableProperty]
    private bool _aceEngineHealthy = true;

    [ObservableProperty]
    private string _aceEngineStatusText = string.Empty;
```

- [ ] **Step 4: Update the constructor**

Replace:

```csharp
    public MainViewModel(
        ExcelReaderService excelReaderService,
        DbfExportService dbfExportService,
        DbfReaderService dbfReaderService,
        SettingsService settingsService,
        RunHistoryService runHistoryService,
        FileLogger fileLogger,
        ExcelSourceStagingService excelSourceStagingService,
        string appBaseDirectory)
    {
        _excelReaderService = excelReaderService;
        _dbfExportService = dbfExportService;
        _dbfReaderService = dbfReaderService;
        _settingsService = settingsService;
        _runHistoryService = runHistoryService;
        _fileLogger = fileLogger;
        _excelSourceStagingService = excelSourceStagingService;
        _appBaseDirectory = appBaseDirectory;
    }
```

with:

```csharp
    public MainViewModel(
        ExcelReaderService excelReaderService,
        DbfExportService dbfExportService,
        DbfReaderService dbfReaderService,
        DbfSafetyBackupService dbfSafetyBackupService,
        SettingsService settingsService,
        RunHistoryService runHistoryService,
        FileLogger fileLogger,
        ExcelSourceStagingService excelSourceStagingService,
        AceEngineGuardService aceEngineGuardService,
        string appBaseDirectory)
    {
        _excelReaderService = excelReaderService;
        _dbfExportService = dbfExportService;
        _dbfReaderService = dbfReaderService;
        _dbfSafetyBackupService = dbfSafetyBackupService;
        _settingsService = settingsService;
        _runHistoryService = runHistoryService;
        _fileLogger = fileLogger;
        _excelSourceStagingService = excelSourceStagingService;
        _aceEngineGuardService = aceEngineGuardService;
        _appBaseDirectory = appBaseDirectory;
    }
```

- [ ] **Step 5: Gate `CanExport` on `AceEngineHealthy`**

Replace:

```csharp
    private bool CanExport() => !IsRunning && _icmasteReadResult is not null && _ictraneReadResult is not null;
```

with:

```csharp
    private bool CanExport() => !IsRunning && AceEngineHealthy && _icmasteReadResult is not null && _ictraneReadResult is not null;
```

- [ ] **Step 6: Add the health-check and repair methods**

Add these new methods after `Initialize()` (right before the `BrowseExcel` command):

```csharp
    /// <summary>
    /// Runs once per app session (called once from App.xaml.cs after the window is shown). Spawns
    /// the disposable self-test child process; on failure, disables Confirm & Export and offers to
    /// repair. Never throws — AceEngineGuardService.CheckHealthAsync catches everything itself.
    /// </summary>
    public async Task RunAceEngineHealthCheckAsync()
    {
        var healthy = await _aceEngineGuardService.CheckHealthAsync();
        SetAceEngineHealth(healthy);

        if (!healthy)
        {
            PromptToRepairAceEngine();
        }
    }

    private void SetAceEngineHealth(bool healthy)
    {
        AceEngineHealthy = healthy;
        AceEngineStatusText = healthy
            ? string.Empty
            : "Access Database Engine isn't working correctly — Confirm & Export is disabled until this is fixed.";
        ConfirmExportCommand.NotifyCanExecuteChanged();
    }

    private void PromptToRepairAceEngine()
    {
        var result = MessageBox.Show(
            "eneBridge's Access Database Engine isn't working correctly, so Confirm & Export has " +
            "been disabled to avoid a crash.\n\n" +
            "This usually means the standalone Access Database Engine component isn't properly " +
            "installed. Would you like eneBridge to try fixing this now? " +
            "(This will prompt for administrator permission.)",
            "eneBridge - Access Database Engine Problem",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result != MessageBoxResult.Yes)
        {
            return;
        }

        _ = RepairAceEngineAsync();
    }

    private async Task RepairAceEngineAsync()
    {
        var fixedNow = await _aceEngineGuardService.TryRepairAsync();
        SetAceEngineHealth(fixedNow);

        MessageBox.Show(
            fixedNow
                ? "Fixed! Confirm & Export is now enabled."
                : "The repair attempt didn't fix the problem. Confirm & Export will stay disabled — " +
                  "please check your Office installation or contact support.",
            "eneBridge - Access Database Engine",
            MessageBoxButton.OK,
            fixedNow ? MessageBoxImage.Information : MessageBoxImage.Error);
    }
```

- [ ] **Step 7: Build to verify it compiles (this also completes Task 4's build verification)**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 8: Manually verify the self-test child process in isolation**

Run:
```bash
"src/eneBridge.Wpf/bin/Debug/net8.0-windows/eneBridge.Wpf.exe" --selftest-ace
echo "exit code: $?"
```
Expected: the process exits immediately with no window shown, `exit code: 0` (this machine's ACE
driver is already known-good from this session's earlier investigation).

- [ ] **Step 9: Commit**

```bash
git add src/eneBridge.Wpf/ViewModels/MainViewModel.cs
git commit -m "Add ACE engine health check, gating, and repair flow to MainViewModel

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 6: `MainWindow.xaml` — status banner

Shows `AceEngineStatusText` near the Confirm & Export button so the disabled state is visible, not
just inferred from a greyed-out button.

**Files:**
- Modify: `src/eneBridge.Wpf/MainWindow.xaml`

- [ ] **Step 1: Add the status TextBlock**

Find this block (the Run panel):

```xml
                    <!-- Run panel -->
                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button Content="Preview" Command="{Binding PreviewCommand}" Padding="24,6" FontWeight="Bold" MinWidth="100" Margin="0,0,8,0"/>
                        <Button Content="Confirm &amp; Export" Command="{Binding ConfirmExportCommand}" Padding="24,6" FontWeight="Bold" MinWidth="140"/>
                    </StackPanel>
```

Replace it with:

```xml
                    <!-- Run panel -->
                    <StackPanel Grid.Row="1" Margin="0,0,0,8">
                        <StackPanel Orientation="Horizontal">
                            <Button Content="Preview" Command="{Binding PreviewCommand}" Padding="24,6" FontWeight="Bold" MinWidth="100" Margin="0,0,8,0"/>
                            <Button Content="Confirm &amp; Export" Command="{Binding ConfirmExportCommand}" Padding="24,6" FontWeight="Bold" MinWidth="140"/>
                        </StackPanel>
                        <TextBlock Text="{Binding AceEngineStatusText}" Foreground="DarkRed" TextWrapping="Wrap" Margin="0,4,0,0"/>
                    </StackPanel>
```

(`AceEngineStatusText` is an empty string when healthy, so the `TextBlock` renders as a blank
line — matches this project's existing preference for simple, always-present status text over
introducing a visibility converter for one banner.)

- [ ] **Step 2: Build to verify it compiles**

Run: `dotnet build "src/eneBridge.Wpf/eneBridge.Wpf.csproj"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 3: Commit**

```bash
git add src/eneBridge.Wpf/MainWindow.xaml
git commit -m "Show ACE engine status banner near Confirm & Export

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 7: Wire the safety backup and pre-export table check into `ConfirmExportAsync`

Both run at the very start of `ConfirmExportAsync`, before either table's export/backup-rename.

**Files:**
- Modify: `src/eneBridge.Wpf/ViewModels/MainViewModel.cs`

- [ ] **Step 1: Add the pre-export table check helper**

Add this new method right after `RepairAceEngineAsync` (from Task 5):

```csharp
    /// <summary>
    /// Warns (with a Continue/Cancel choice) if either table can't currently be read — covers
    /// "file doesn't exist", "table locked by EMAS", and "folder unreachable" alike, since
    /// DbfReaderService.Read already catches all three uniformly. Runs only once AceEngineHealthy
    /// is known-good (CanExport already requires this), so it's safe to do in-process. Checks both
    /// tables even if the first fails, unless the user cancels on the first warning.
    /// </summary>
    private async Task<bool> ConfirmTablesReadableAsync(string dbfFolder)
    {
        foreach (var tableName in new[] { IcmasteSchema.TableName, IctraneSchema.TableName })
        {
            var readResult = await Task.Run(() => _dbfReaderService.Read(dbfFolder, tableName));
            if (readResult.Success)
            {
                continue;
            }

            var proceed = MessageBox.Show(
                $"{tableName}.dbf could not be found or read in '{dbfFolder}':\n{readResult.ErrorMessage}\n\n" +
                "This is expected on a first-ever export to this folder, but if you expect this " +
                "table to already exist, something may be wrong (wrong folder selected, the table " +
                "is locked by EMAS, or a previous export was interrupted).\n\n" +
                "Continue anyway?",
                "eneBridge - Table Check",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (proceed != MessageBoxResult.Yes)
            {
                return false;
            }
        }

        return true;
    }
```

- [ ] **Step 2: Wire both checks into the start of `ConfirmExportAsync`, and skip run-history logging on cancel**

Replace:

```csharp
    [RelayCommand(CanExecute = nameof(CanExport))]
    private async Task ConfirmExportAsync()
    {
        IsRunning = true;
        PreviewCommand.NotifyCanExecuteChanged();
        ConfirmExportCommand.NotifyCanExecuteChanged();
        IcmasteDbfPreview = null;
        IctraneDbfPreview = null;
        IcmasteDbfStatusText = "Verifying…";
        IctraneDbfStatusText = "Verifying…";
        IcmasteDbfVerified = false;
        IctraneDbfVerified = false;

        string excelPath = ExcelFilePath;
        string dbfFolder = DbfFolderPath;
        var stopwatch = Stopwatch.StartNew();

        StageExportResult? icmasteExport = null;
        StageExportResult? ictraneExport = null;
        string? fatalError = null;

        try
        {
            icmasteExport = await ExportStageAsync(
```

with:

```csharp
    [RelayCommand(CanExecute = nameof(CanExport))]
    private async Task ConfirmExportAsync()
    {
        IsRunning = true;
        PreviewCommand.NotifyCanExecuteChanged();
        ConfirmExportCommand.NotifyCanExecuteChanged();
        IcmasteDbfPreview = null;
        IctraneDbfPreview = null;
        IcmasteDbfStatusText = "Verifying…";
        IctraneDbfStatusText = "Verifying…";
        IcmasteDbfVerified = false;
        IctraneDbfVerified = false;

        string excelPath = ExcelFilePath;
        string dbfFolder = DbfFolderPath;
        var stopwatch = Stopwatch.StartNew();

        StageExportResult? icmasteExport = null;
        StageExportResult? ictraneExport = null;
        string? fatalError = null;
        bool cancelledByUser = false;

        try
        {
            await Task.Run(() =>
            {
                _dbfSafetyBackupService.BackupIfExists(dbfFolder, IcmasteSchema.TableName);
                _dbfSafetyBackupService.BackupIfExists(dbfFolder, IctraneSchema.TableName);
            });

            if (!await ConfirmTablesReadableAsync(dbfFolder))
            {
                cancelledByUser = true;
                AppendLog("Export cancelled by user after the table check.");
                return;
            }

            icmasteExport = await ExportStageAsync(
```

- [ ] **Step 3: Skip the run-history entry when the user cancelled**

Replace:

```csharp
        finally
        {
            stopwatch.Stop();
            var historyEntry = new RunHistoryEntry
            {
                ExcelPath = excelPath,
                DbfFolder = dbfFolder,
                IcmasteRowsRead = _icmasteReadResult?.RowsRead ?? 0,
                IcmasteRowsSkipped = _icmasteReadResult?.SkipReasons.Count ?? 0,
                IcmasteRowsWritten = icmasteExport?.RowsWritten ?? 0,
                IcmasteSuccess = icmasteExport?.Success ?? false,
                IctraneRowsRead = _ictraneReadResult?.RowsRead ?? 0,
                IctraneRowsSkipped = _ictraneReadResult?.SkipReasons.Count ?? 0,
                IctraneRowsWritten = ictraneExport?.RowsWritten ?? 0,
                IctraneSuccess = ictraneExport?.Success ?? false,
                ErrorSummary = fatalError,
                DurationMs = stopwatch.ElapsedMilliseconds,
            };
            _runHistoryService.Append(historyEntry);
            RunHistory.Insert(0, historyEntry);
            IsRunning = false;
            PreviewCommand.NotifyCanExecuteChanged();
            ConfirmExportCommand.NotifyCanExecuteChanged();
        }
```

with:

```csharp
        finally
        {
            stopwatch.Stop();
            if (!cancelledByUser)
            {
                var historyEntry = new RunHistoryEntry
                {
                    ExcelPath = excelPath,
                    DbfFolder = dbfFolder,
                    IcmasteRowsRead = _icmasteReadResult?.RowsRead ?? 0,
                    IcmasteRowsSkipped = _icmasteReadResult?.SkipReasons.Count ?? 0,
                    IcmasteRowsWritten = icmasteExport?.RowsWritten ?? 0,
                    IcmasteSuccess = icmasteExport?.Success ?? false,
                    IctraneRowsRead = _ictraneReadResult?.RowsRead ?? 0,
                    IctraneRowsSkipped = _ictraneReadResult?.SkipReasons.Count ?? 0,
                    IctraneRowsWritten = ictraneExport?.RowsWritten ?? 0,
                    IctraneSuccess = ictraneExport?.Success ?? false,
                    ErrorSummary = fatalError,
                    DurationMs = stopwatch.ElapsedMilliseconds,
                };
                _runHistoryService.Append(historyEntry);
                RunHistory.Insert(0, historyEntry);
            }
            IsRunning = false;
            PreviewCommand.NotifyCanExecuteChanged();
            ConfirmExportCommand.NotifyCanExecuteChanged();
        }
```

- [ ] **Step 4: Build to verify it compiles**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 5: Manually verify end-to-end against a scratch folder**

```bash
mkdir -p /tmp/enebridge-manual-verify
```

Run the app (`dotnet run --project "src/eneBridge.Wpf/eneBridge.Wpf.csproj"`), point the DBF
folder field at that empty scratch folder, run Preview against the real fixture/staged Excel file,
then Confirm & Export.

Expected:
- A "Table Check" dialog appears twice (icmaste then ictrane — neither exists yet in the empty
  folder); clicking **Continue** both times lets the export proceed normally to completion.
- `%AppData%\eneBridge\Backups\icmaste\` and `...\Backups\ictrane\` do **not** get a new entry this
  run (nothing existed yet to back up).
- Run the same Preview → Confirm & Export a second time against the same now-populated scratch
  folder: this time the Table Check dialogs don't appear (both tables now exist and read fine),
  and `%AppData%\eneBridge\Backups\icmaste\` and `...\Backups\ictrane\` each gain one new
  timestamped `.dbf` backup of what existed before this second run.

- [ ] **Step 6: Commit**

```bash
git add src/eneBridge.Wpf/ViewModels/MainViewModel.cs
git commit -m "Run safety backup and pre-export table check before Confirm & Export writes anything

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 8: Installer detection fix (`installer/eneBridge.iss`)

Fixes `IsAccessDatabaseEngineInstalled` to check which DLL the provider actually resolves to
(not just that the ProgID key exists), and keeps the redistributable installed permanently instead
of deleting it after setup.

**Files:**
- Modify: `installer/eneBridge.iss`

- [ ] **Step 1: Update the `[Files]` section**

Replace:

```
Source: "publish\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "..\prerequisites\accessdatabaseengine_X64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "..\prerequisites\dotnet-runtime-8.0.23-win-x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
```

with:

```
Source: "publish\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "..\prerequisites\accessdatabaseengine_X64.exe"; DestDir: "{app}\prerequisites"; Flags: ignoreversion
Source: "..\prerequisites\dotnet-runtime-8.0.23-win-x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
```

(Only the Access Database Engine installer moves to a permanent location — the app re-invokes this
one later for the in-app "Fix it now" repair flow. The .NET runtime installer has no such reuse
case and stays temp-deleted.)

- [ ] **Step 2: Update the `[Run]` section**

Replace:

```
[Run]
Filename: "{tmp}\accessdatabaseengine_X64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing Access Database Engine..."; Check: not IsAccessDatabaseEngineInstalled
Filename: "{tmp}\dotnet-runtime-8.0.23-win-x64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing .NET 8 Desktop Runtime..."; Check: not IsDotNet8DesktopRuntimeInstalled
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
```

with:

```
[Run]
Filename: "{app}\prerequisites\accessdatabaseengine_X64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing Access Database Engine..."; Check: not IsAccessDatabaseEngineInstalled
Filename: "{tmp}\dotnet-runtime-8.0.23-win-x64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing .NET 8 Desktop Runtime..."; Check: not IsDotNet8DesktopRuntimeInstalled
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
```

- [ ] **Step 3: Fix the detection function in `[Code]`**

Replace:

```
[Code]
function IsAccessDatabaseEngineInstalled(): Boolean;
begin
  { Checks the actual COM registration for the ACE OleDb provider that DbfExportService depends
    on (Provider=Microsoft.ACE.OLEDB.12.0) -- the most direct signal that it's already usable. }
  Result := RegKeyExists(HKLM, 'SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0\CLSID') or
            RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Classes\Microsoft.ACE.OLEDB.12.0\CLSID');
end;
```

with:

```
[Code]
function GetAceInprocServerPath(): String;
var
  ClsidStr: String;
  DllPath: String;
begin
  { Resolves what DLL Microsoft.ACE.OLEDB.12.0 actually points at: read the ProgID's CLSID, then
    that CLSID's InprocServer32 default value, checking both the native and Wow6432Node registry
    views since the provider can be registered under either depending on what else is installed. }
  Result := '';

  if not (RegQueryStringValue(HKLM, 'SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0\CLSID', '', ClsidStr) or
          RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Classes\Microsoft.ACE.OLEDB.12.0\CLSID', '', ClsidStr)) then
    exit;

  if RegQueryStringValue(HKLM, 'SOFTWARE\Classes\CLSID\' + ClsidStr + '\InprocServer32', '', DllPath) then
  begin
    Result := DllPath;
    exit;
  end;

  if RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Classes\CLSID\' + ClsidStr + '\InprocServer32', '', DllPath) then
    Result := DllPath;
end;

function IsAccessDatabaseEngineInstalled(): Boolean;
var
  DllPath: String;
begin
  { A key-existence check alone isn't enough: Office's Click-to-Run install also registers
    Microsoft.ACE.OLEDB.12.0, pointing at its own sandboxed copy of ACEOLEDB.DLL (path contains
    '\root\VFS\'), which crashes natively (AccessViolationException) on CREATE TABLE/write --
    confirmed live on a real machine, see
    docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md. Only the standalone
    redistributable's copy (this installer's own accessdatabaseengine_X64.exe) counts as properly
    installed here. }
  DllPath := GetAceInprocServerPath();
  Result := (DllPath <> '') and (Pos('\root\VFS\', DllPath) = 0);
end;
```

- [ ] **Step 4: Verify (best-effort — Inno Setup 6 is not installed in this dev environment)**

If `iscc` is available:
Run: `iscc "installer\eneBridge.iss"`
Expected: compiles with no errors, producing `installer\Output\eneBridge-Setup.exe`.

If `iscc` is not available (as in this session), skip compilation and instead re-read the diff for
Pascal syntax correctness (matched `begin`/`end`, semicolons, `RegQueryStringValue`'s
`var ResultStr: String` out-parameter usage) — this must be verified on a machine with Inno Setup
6 installed before the next real installer build.

- [ ] **Step 5: Commit**

```bash
git add installer/eneBridge.iss
git commit -m "Fix installer's ACE engine detection to check the actual bound DLL, not just key existence

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 9: Update `CLAUDE.md`

Documents the four new safety nets so future sessions (and future you) know they exist, and
narrows the scope of the still-open "Known reliability risk" note now that this session's actual
trigger (wrong driver bound) has a fix.

**Files:**
- Modify: `CLAUDE.md`

- [ ] **Step 1: Add a new subsection under "Architecture", after the DBF export paragraph**

Find the paragraph beginning `**DBF export**` (ends with `...to guard against future schema
edits).`) and add a new paragraph immediately after it:

```markdown
**DBF export crash prevention**: four independent layers guard against the ACE OleDb driver's
native-crash risk described above and the incident that surfaced it (`Microsoft.ACE.OLEDB.12.0`
silently resolving to Office Click-to-Run's own sandboxed `ACEOLEDB.DLL` instead of the standalone
redistributable — see `docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md`
for the full incident writeup). `DbfSafetyBackupService` copies `icmaste.dbf`/`ictrane.dbf` (plus
`.FPT` companion) to `%AppData%\eneBridge\Backups\<table>\` before anything else runs in
`ConfirmExportAsync` — plain `File.Copy`, independent of OleDb, so it works even if the driver
itself is broken. `AceEngineGuardService` spawns `eneBridge.Wpf.exe --selftest-ace` as a disposable
child process once per session (via `AceEngineSelfTestService`, a real DBF round-trip through a
throwaway table) so a native crash from a broken driver only kills that child, never the app;
`MainViewModel.CanExport` is gated on the result, with a "Fix it now" repair flow that re-runs the
bundled `accessdatabaseengine_X64.exe` elevated. `MainViewModel.ConfirmTablesReadableAsync` warns
(Continue/Cancel) if either table can't be read before the export's own backup-rename runs.
`installer/eneBridge.iss`'s `IsAccessDatabaseEngineInstalled` now resolves the ACE provider's
actual bound DLL path (not just registry-key existence) so it correctly detects and fixes the
Office Click-to-Run case instead of silently skipping the real fix.
```

- [ ] **Step 2: Narrow the "Known reliability risk" note in "Current status"**

Find the paragraph beginning `**Known reliability risk, not yet fixed**:` and add one sentence to
its end:

Replace the final sentence:
```markdown
If a real row fails during a run, the *other* stage's export in the same run could be at
risk of crashing the app outright instead of failing gracefully, which would undermine this
project's "nothing should ever crash the app" goal.
```

with:
```markdown
If a real row fails during a run, the *other* stage's export in the same run could be at
risk of crashing the app outright instead of failing gracefully, which would undermine this
project's "nothing should ever crash the app" goal. Separately, a distinct and more common trigger
for this same crash class — the ACE OleDb provider resolving to Office Click-to-Run's sandboxed
DLL instead of the standalone redistributable — was found and mitigated (see "DBF export crash
prevention" above); this per-row/reused-connection risk remains open.
```

- [ ] **Step 3: Commit**

```bash
git add CLAUDE.md
git commit -m "Document the DBF export crash-prevention safety nets in CLAUDE.md

Co-Authored-By: Claude Sonnet 5 <noreply@anthropic.com>"
```

---

### Task 10: Full verification pass

**Files:** none (verification only)

- [ ] **Step 1: Full test suite**

Run: `dotnet test "tests/eneBridge.Wpf.Core.Tests/eneBridge.Wpf.Core.Tests.csproj"`
Expected: all tests pass (56 existing + 6 new from Tasks 1–2 = 62), no crashed test host.

- [ ] **Step 2: Full solution build**

Run: `dotnet build "src/eneBridge.Wpf.slnx"`
Expected: `Build succeeded. 0 Warning(s) 0 Error(s)`

- [ ] **Step 3: Re-confirm the self-test CLI path still exits 0 after all changes**

Run:
```bash
"src/eneBridge.Wpf/bin/Debug/net8.0-windows/eneBridge.Wpf.exe" --selftest-ace
echo "exit code: $?"
```
Expected: `exit code: 0`, no window shown.

- [ ] **Step 4: Confirm no production files were touched by any of this session's verification runs**

Run (PowerShell):
```powershell
Test-Path "C:\emas\ch2026\acc\data\icmaste.dbf"
Test-Path "C:\emas\ch2026\acc\data\ictrane.dbf"
```
Expected: both `True` (matches the restored state from earlier in this session — this plan's
manual verification steps only ever touch scratch/temp folders).
