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

        var excelReaderService = new ExcelReaderService();
        var dbfExportService = new DbfExportService();
        var settingsService = new SettingsService(AppContext.BaseDirectory);
        var runHistoryService = new RunHistoryService();
        var fileLogger = new FileLogger();

        var mainViewModel = new MainViewModel(
            excelReaderService,
            dbfExportService,
            settingsService,
            runHistoryService,
            fileLogger,
            AppContext.BaseDirectory);
        mainViewModel.Initialize();

        var mainWindow = new MainWindow { DataContext = mainViewModel };
        mainWindow.Show();
    }
}
