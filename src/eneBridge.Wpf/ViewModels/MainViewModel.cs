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
    private readonly ExcelSourceStagingService _excelSourceStagingService;
    private readonly string _appBaseDirectory;

    private StageReadResult? _icmasteReadResult;
    private StageReadResult? _ictraneReadResult;

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

    /// <summary>
    /// When false (default), the preview grids show only the columns ExcelReaderService actually
    /// populates; the rest of the required 151/80-column DBF schema is hidden from view but still
    /// fully exported unchanged. Purely a display preference — never affects what gets written.
    /// </summary>
    [ObservableProperty]
    private bool _showAllColumns;

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
        ExcelSourceStagingService excelSourceStagingService,
        string appBaseDirectory)
    {
        _excelReaderService = excelReaderService;
        _dbfExportService = dbfExportService;
        _settingsService = settingsService;
        _runHistoryService = runHistoryService;
        _fileLogger = fileLogger;
        _excelSourceStagingService = excelSourceStagingService;
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
            ExcelFilePath = StageExcelSource(dialog.FileName);
            SaveUserPaths();
        }
    }

    private const string InvoiceStagedFileName = "invoice.xlsx";

    /// <summary>
    /// Copies the picked Excel file into the local staging folder so later runs don't depend on
    /// removable/network media staying plugged in. Never blocks the user: if staging fails for any
    /// reason, the failure is logged and the original picked path is used as-is.
    /// </summary>
    private string StageExcelSource(string originalPath)
    {
        try
        {
            var stagedPath = _excelSourceStagingService.StageFile(originalPath, InvoiceStagedFileName);
            _fileLogger.LogInfo($"Imported Excel source '{originalPath}' -> '{stagedPath}'");
            return stagedPath;
        }
        catch (Exception ex)
        {
            _fileLogger.LogException("Failed to stage Excel source file", ex);
            return originalPath;
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

    partial void OnExcelFilePathChanged(string value) => InvalidatePreview();
    partial void OnDbfFolderPathChanged(string value) => InvalidatePreview();

    private void InvalidatePreview()
    {
        _icmasteReadResult = null;
        _ictraneReadResult = null;
        ConfirmExportCommand.NotifyCanExecuteChanged();
    }

    private bool CanPreview() => !IsRunning;

    /// <summary>
    /// Reads the Excel file and populates the previews/skip-reason log. Writes nothing to disk —
    /// this is the pre-verify checkpoint before Confirm &amp; Export actually writes the DBF files.
    /// </summary>
    [RelayCommand(CanExecute = nameof(CanPreview))]
    private async Task PreviewAsync()
    {
        IsRunning = true;
        PreviewCommand.NotifyCanExecuteChanged();
        ConfirmExportCommand.NotifyCanExecuteChanged();
        LogLines.Clear();
        IcmasteStage.Reset();
        IctraneStage.Reset();
        _icmasteReadResult = null;
        _ictraneReadResult = null;

        string excelPath = ExcelFilePath;

        try
        {
            var openResult = await Task.Run(() => TryOpenWorkbook(excelPath));

            if (openResult.Workbook is null)
            {
                var message = openResult.Error ?? "Failed to open the Excel file.";
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

                    _icmasteReadResult = await ReadStageAsync(
                        "icmaste",
                        IcmasteStage,
                        () => _excelReaderService.ReadIcmaste(worksheet),
                        result => IcmastePreview = result.Table.DefaultView);

                    _ictraneReadResult = await ReadStageAsync(
                        "ictrane",
                        IctraneStage,
                        () => _excelReaderService.ReadIctrane(worksheet),
                        result => IctranePreview = result.Table.DefaultView);
                }
            }
        }
        catch (Exception ex)
        {
            // Outer backstop: nothing should ever crash the app.
            _fileLogger.LogException("Unhandled error during preview", ex);
            AppendLog($"Unexpected error: {ex.Message}");
        }
        finally
        {
            IsRunning = false;
            PreviewCommand.NotifyCanExecuteChanged();
            ConfirmExportCommand.NotifyCanExecuteChanged();
        }
    }

    private bool CanExport() => !IsRunning && _icmasteReadResult is not null && _ictraneReadResult is not null;

    /// <summary>Writes the DBF files from the read results Preview already produced.</summary>
    [RelayCommand(CanExecute = nameof(CanExport))]
    private async Task ConfirmExportAsync()
    {
        IsRunning = true;
        PreviewCommand.NotifyCanExecuteChanged();
        ConfirmExportCommand.NotifyCanExecuteChanged();

        string excelPath = ExcelFilePath;
        string dbfFolder = DbfFolderPath;
        var stopwatch = Stopwatch.StartNew();

        StageExportResult? icmasteExport = null;
        StageExportResult? ictraneExport = null;
        string? fatalError = null;

        try
        {
            icmasteExport = await ExportStageAsync(
                "icmaste",
                IcmasteStage,
                _icmasteReadResult!,
                result => _dbfExportService.Export(dbfFolder, IcmasteSchema.TableName, result.Table, IcmasteSchema.Columns));

            ictraneExport = await ExportStageAsync(
                "ictrane",
                IctraneStage,
                _ictraneReadResult!,
                result => _dbfExportService.Export(dbfFolder, IctraneSchema.TableName, result.Table, IctraneSchema.Columns));
        }
        catch (Exception ex)
        {
            // Outer backstop: nothing should ever crash the app.
            _fileLogger.LogException("Unhandled error during export", ex);
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
    }

    /// <summary>Reads one table, logging skip reasons and updating the stage VM. Never throws.</summary>
    private async Task<StageReadResult> ReadStageAsync(
        string label,
        StageProgressViewModel stage,
        Func<StageReadResult> read,
        Action<StageReadResult> onRead)
    {
        var readResult = await Task.Run(read);
        onRead(readResult);

        var readyCount = readResult.Table.Rows.Count;
        stage.Complete(true, $"Read {readResult.RowsRead}, skipped {readResult.SkipReasons.Count}, {readyCount} ready to export");

        return readResult;
    }

    /// <summary>Exports one already-read table, logging row errors and updating the stage VM. Never throws.</summary>
    private async Task<StageExportResult?> ExportStageAsync(
        string label,
        StageProgressViewModel stage,
        StageReadResult readResult,
        Func<StageReadResult, StageExportResult> export)
    {
        try
        {
            var exportResult = await Task.Run(() => export(readResult));
            foreach (var rowError in exportResult.RowErrors)
            {
                AppendLog($"[{label}] {rowError}");
            }

            if (exportResult.Success)
            {
                AppendLog($"[{label}] Done — written {exportResult.RowsWritten} row(s).");
                stage.Complete(true, $"Read {readResult.RowsRead}, skipped {readResult.SkipReasons.Count}, written {exportResult.RowsWritten}");
            }
            else
            {
                AppendLog($"[{label}] Export failed: {exportResult.ErrorMessage}");
                stage.Complete(false, $"Export failed: {exportResult.ErrorMessage}");
            }

            return exportResult;
        }
        catch (Exception ex)
        {
            _fileLogger.LogException($"{label} stage export failed", ex);
            AppendLog($"[{label}] Failed: {ex.Message}");
            stage.Complete(false, ex.Message);
            return null;
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
