using System.Windows;
using eneBridge.Wpf.Core.Services;
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

        var fileLogger = new FileLogger();
        RegisterGlobalExceptionHandlers(fileLogger);

        var excelReaderService = new ExcelReaderService();
        var dbfExportService = new DbfExportService(fileLogger);
        var dbfReaderService = new DbfReaderService(fileLogger);
        var settingsService = new SettingsService(AppContext.BaseDirectory);
        var runHistoryService = new RunHistoryService();
        var excelSourceStagingService = new ExcelSourceStagingService();

        var mainViewModel = new MainViewModel(
            excelReaderService,
            dbfExportService,
            dbfReaderService,
            settingsService,
            runHistoryService,
            fileLogger,
            excelSourceStagingService,
            AppContext.BaseDirectory);
        mainViewModel.Initialize();

        var mainWindow = new MainWindow { DataContext = mainViewModel };
        mainWindow.Show();
    }

    /// <summary>
    /// Best-effort safety net: logs anything catchable before it takes down the app, and keeps
    /// the app alive for exceptions on the UI thread (matches this project's "nothing should ever
    /// crash the app" goal). Cannot catch a native AccessViolationException from the ACE OleDb
    /// driver (see DbfExportService) — that class of fault is fatal by design in modern .NET and
    /// bypasses all of these handlers, which is why DbfExportService logs its own checkpoints.
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
