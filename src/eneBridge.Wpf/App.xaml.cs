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
        var foxProDbfReader = new FoxProDbfReader(fileLogger);
        var dbfSafetyBackupService = new DbfSafetyBackupService(fileLogger);
        var settingsService = new SettingsService(AppContext.BaseDirectory);
        var runHistoryService = new RunHistoryService();
        var stockReceivedRunHistoryService = new RunHistoryService(fileName: "stockReceivedRunHistory.json");
        var excelSourceStagingService = new ExcelSourceStagingService();
        var aceEngineGuardService = new AceEngineGuardService(fileLogger, AppContext.BaseDirectory);
        var exportGateService = new ExportGateService();

        var stockReceivedViewModel = new StockReceivedViewModel(
            excelReaderService,
            dbfExportService,
            dbfReaderService,
            foxProDbfReader,
            dbfSafetyBackupService,
            settingsService,
            stockReceivedRunHistoryService,
            fileLogger,
            excelSourceStagingService,
            exportGateService,
            AppContext.BaseDirectory);
        stockReceivedViewModel.Initialize();

        var mainViewModel = new MainViewModel(
            excelReaderService,
            dbfExportService,
            dbfReaderService,
            foxProDbfReader,
            dbfSafetyBackupService,
            settingsService,
            runHistoryService,
            fileLogger,
            excelSourceStagingService,
            aceEngineGuardService,
            stockReceivedViewModel,
            exportGateService,
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
