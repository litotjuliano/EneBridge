using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ExcelReaderService _excelReaderService;
    private readonly DbfExportService _dbfExportService;
    private readonly SettingsService _settingsService;
    private readonly RunHistoryService _runHistoryService;
    private readonly FileLogger _fileLogger;
    private readonly string _appBaseDirectory;

    [ObservableProperty]
    private string _excelFilePath = string.Empty;

    [ObservableProperty]
    private string _dbfFolderPath = string.Empty;

    [ObservableProperty]
    private bool _isRunning;

    [ObservableProperty]
    private DataView? _icmastePreview;

    [ObservableProperty]
    private DataView? _ictranePreview;

    public ObservableCollection<RunHistoryEntry> RunHistory { get; } = [];
    public ObservableCollection<string> LogLines { get; } = [];

    public StageProgressViewModel IcmasteStage { get; } = new("icmaste");
    public StageProgressViewModel IctraneStage { get; } = new("ictrane");

    public MainViewModel(
        ExcelReaderService excelReaderService,
        DbfExportService dbfExportService,
        SettingsService settingsService,
        RunHistoryService runHistoryService,
        FileLogger fileLogger,
        string appBaseDirectory)
    {
        _excelReaderService = excelReaderService;
        _dbfExportService = dbfExportService;
        _settingsService = settingsService;
        _runHistoryService = runHistoryService;
        _fileLogger = fileLogger;
        _appBaseDirectory = appBaseDirectory;
    }

    public void Initialize()
    {
        var (excelPath, dbfFolder) = _settingsService.ResolveEffectivePaths(_appBaseDirectory);
        ExcelFilePath = excelPath;
        DbfFolderPath = dbfFolder;

        // Most recent run first.
        foreach (var entry in _runHistoryService.LoadAll().AsEnumerable().Reverse())
        {
            RunHistory.Add(entry);
        }
    }

    [RelayCommand]
    private void BrowseExcel()
    {
        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Filter = "Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*",
            FileName = ExcelFilePath
        };
        if (dialog.ShowDialog() == true)
        {
            ExcelFilePath = dialog.FileName;
            SaveUserPaths();
        }
    }

    [RelayCommand]
    private void BrowseDbfFolder()
    {
        var dialog = new Microsoft.Win32.OpenFolderDialog
        {
            InitialDirectory = Directory.Exists(DbfFolderPath) ? DbfFolderPath : string.Empty
        };
        if (dialog.ShowDialog() == true)
        {
            DbfFolderPath = dialog.FolderName;
            SaveUserPaths();
        }
    }

    private void SaveUserPaths()
    {
        _settingsService.SaveUserSettings(new UserSettingsModel
        {
            ExcelFilePath = ExcelFilePath,
            DbfFolderPath = DbfFolderPath
        });
    }

    private bool CanRun() => !IsRunning;

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync()
    {
        IsRunning = true;
        RunCommand.NotifyCanExecuteChanged();
        LogLines.Clear();
        IcmasteStage.Reset();
        IctraneStage.Reset();

        string excelPath = ExcelFilePath;
        string dbfFolder = DbfFolderPath;
        var stopwatch = Stopwatch.StartNew();

        StageReadResult? icmasteRead = null;
        StageExportResult? icmasteExport = null;
        StageReadResult? ictraneRead = null;
        StageExportResult? ictraneExport = null;
        string? fatalError = null;

        try
        {
            var openResult = await Task.Run(() => TryOpenWorkbook(excelPath));

            if (openResult.Workbook is null)
            {
                var message = openResult.Error ?? "Failed to open the Excel file.";
                fatalError = message;
                AppendLog($"Cannot open Excel file: {message}");
                IcmasteStage.Fail(message);
                IctraneStage.Fail(message);
            }
            else
            {
                using (openResult.Workbook)
                {
                    var workbook = openResult.Workbook;
                    var worksheet = workbook.Worksheets.First();

                    (icmasteRead, icmasteExport) = await RunStageAsync(
                        "icmaste",
                        IcmasteStage,
                        () => _excelReaderService.ReadIcmaste(worksheet),
                        result => IcmastePreview = result.Table.DefaultView,
                        (result, columns) => _dbfExportService.Export(dbfFolder, IcmasteSchema.TableName, result.Table, columns),
                        IcmasteSchema.Columns);

                    (ictraneRead, ictraneExport) = await RunStageAsync(
                        "ictrane",
                        IctraneStage,
                        () => _excelReaderService.ReadIctrane(worksheet),
                        result => IctranePreview = result.Table.DefaultView,
                        (result, columns) => _dbfExportService.Export(dbfFolder, IctraneSchema.TableName, result.Table, columns),
                        IctraneSchema.Columns);
                }
            }
        }
        catch (Exception ex)
        {
            // Outer backstop: nothing should ever crash the app.
            _fileLogger.LogException("Unhandled error during run", ex);
            AppendLog($"Unexpected error: {ex.Message}");
            fatalError = ex.Message;
        }
        finally
        {
            stopwatch.Stop();
            var historyEntry = new RunHistoryEntry
            {
                ExcelPath = excelPath,
                DbfFolder = dbfFolder,
                IcmasteRowsRead = icmasteRead?.RowsRead ?? 0,
                IcmasteRowsSkipped = icmasteRead?.SkipReasons.Count ?? 0,
                IcmasteRowsWritten = icmasteExport?.RowsWritten ?? 0,
                IcmasteSuccess = icmasteExport?.Success ?? false,
                IctraneRowsRead = ictraneRead?.RowsRead ?? 0,
                IctraneRowsSkipped = ictraneRead?.SkipReasons.Count ?? 0,
                IctraneRowsWritten = ictraneExport?.RowsWritten ?? 0,
                IctraneSuccess = ictraneExport?.Success ?? false,
                ErrorSummary = fatalError,
                DurationMs = stopwatch.ElapsedMilliseconds,
            };
            _runHistoryService.Append(historyEntry);
            RunHistory.Insert(0, historyEntry);
            IsRunning = false;
            RunCommand.NotifyCanExecuteChanged();
        }
    }

    /// <summary>Reads then exports one table, logging skip reasons/row errors and updating the stage VM. Never throws — failures are caught and recorded on the stage.</summary>
    private async Task<(StageReadResult? read, StageExportResult? export)> RunStageAsync(
        string label,
        StageProgressViewModel stage,
        Func<StageReadResult> read,
        Action<StageReadResult> onRead,
        Func<StageReadResult, IReadOnlyList<DbfColumnDefinition>, StageExportResult> export,
        IReadOnlyList<DbfColumnDefinition> schema)
    {
        try
        {
            var readResult = await Task.Run(read);
            foreach (var reason in readResult.SkipReasons)
            {
                AppendLog($"[{label}] Row {reason.ExcelRow}: {reason.Reason}");
            }
            onRead(readResult);
            AppendLog($"[{label}] Read {readResult.RowsRead}, skipped {readResult.SkipReasons.Count}, parsed {readResult.RowsWritten}");

            var exportResult = await Task.Run(() => export(readResult, schema));
            foreach (var rowError in exportResult.RowErrors)
            {
                AppendLog($"[{label}] {rowError}");
            }

            if (exportResult.Success)
            {
                stage.Complete(true, $"Read {readResult.RowsRead}, skipped {readResult.SkipReasons.Count}, written {exportResult.RowsWritten}");
            }
            else
            {
                AppendLog($"[{label}] Export failed: {exportResult.ErrorMessage}");
                stage.Complete(false, $"Export failed: {exportResult.ErrorMessage}");
            }

            return (readResult, exportResult);
        }
        catch (Exception ex)
        {
            _fileLogger.LogException($"{label} stage failed", ex);
            AppendLog($"[{label}] Failed: {ex.Message}");
            stage.Complete(false, ex.Message);
            return (null, null);
        }
    }

    private (XLWorkbook? Workbook, string? Error) TryOpenWorkbook(string excelPath)
    {
        try
        {
            return (_excelReaderService.OpenWorkbook(excelPath), null);
        }
        catch (Exception ex)
        {
            _fileLogger.LogException("Failed to open workbook", ex);
            return (null, ex.Message);
        }
    }

    private void AppendLog(string message) => LogLines.Add($"{DateTime.Now:HH:mm:ss} {message}");
}
