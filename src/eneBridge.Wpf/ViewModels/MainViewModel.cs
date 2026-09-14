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

namespace eneBridge.Wpf.ViewModels;

public partial class MainViewModel : ObservableObject
{
    private readonly ExcelReaderService _excelReaderService;
    private readonly DbfExportService _dbfExportService;
    private readonly DbfReaderService _dbfReaderService;
    private readonly DbfSafetyBackupService _dbfSafetyBackupService;
    private readonly SettingsService _settingsService;
    private readonly RunHistoryService _runHistoryService;
    private readonly FileLogger _fileLogger;
    private readonly ExcelSourceStagingService _excelSourceStagingService;
    private readonly AceEngineGuardService _aceEngineGuardService;
    private readonly ExportGateService _exportGateService;
    private readonly string _appBaseDirectory;

    public StockReceivedViewModel StockReceivedViewModel { get; }

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

    [ObservableProperty]
    private DataView? _icmasteDbfPreview;

    [ObservableProperty]
    private DataView? _ictraneDbfPreview;

    [ObservableProperty]
    private string _icmasteDbfStatusText = "Run Confirm & Export to verify.";

    [ObservableProperty]
    private string _ictraneDbfStatusText = "Run Confirm & Export to verify.";

    [ObservableProperty]
    private bool _icmasteDbfVerified;

    [ObservableProperty]
    private bool _ictraneDbfVerified;

    /// <summary>
    /// When false (default), the preview grids show only the columns ExcelReaderService actually
    /// populates; the rest of the required 151/80-column DBF schema is hidden from view but still
    /// fully exported unchanged. Purely a display preference — never affects what gets written.
    /// </summary>
    [ObservableProperty]
    private bool _showAllColumns;

    /// <summary>
    /// Set once per session by RunAceEngineHealthCheckAsync. Starts true (optimistic) so the app
    /// isn't blocked while the startup check runs; ConfirmExportCommand is gated on this so a
    /// broken ACE driver disables Confirm & Export instead of crashing it.
    /// </summary>
    [ObservableProperty]
    private bool _aceEngineHealthy = true;

    [ObservableProperty]
    private string _aceEngineStatusText = string.Empty;

    public ObservableCollection<RunHistoryEntry> RunHistory { get; } = [];
    public ObservableCollection<string> LogLines { get; } = [];

    public StageProgressViewModel IcmasteStage { get; } = new("icmaste");
    public StageProgressViewModel IctraneStage { get; } = new("ictrane");

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
        StockReceivedViewModel stockReceivedViewModel,
        ExportGateService exportGateService,
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
        StockReceivedViewModel = stockReceivedViewModel;
        _exportGateService = exportGateService;
        _appBaseDirectory = appBaseDirectory;

        _exportGateService.StateChanged += () => ConfirmExportCommand.NotifyCanExecuteChanged();
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
        StockReceivedViewModel.SetAceEngineHealth(healthy);
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
        _settingsService.SaveUserSettings(s =>
        {
            s.ExcelFilePath = ExcelFilePath;
            s.DbfFolderPath = DbfFolderPath;
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

    private bool CanExport() => !IsRunning && AceEngineHealthy && _icmasteReadResult is not null && _ictraneReadResult is not null && !_exportGateService.IsExportInProgress;

    /// <summary>Writes the DBF files from the read results Preview already produced.</summary>
    [RelayCommand(CanExecute = nameof(CanExport))]
    private async Task ConfirmExportAsync()
    {
        if (!_exportGateService.TryBeginExport())
        {
            AppendLog("Cannot export: the Stock Received tab is currently exporting. Please wait for it to finish and try again.");
            return;
        }

        try
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

                if (!DbfWorkflowHelper.ConfirmTablesReadable(dbfFolder))
                {
                    cancelledByUser = true;
                    IcmasteDbfStatusText = "Run Confirm & Export to verify.";
                    IctraneDbfStatusText = "Run Confirm & Export to verify.";
                    AppendLog("Export cancelled by user after the table check.");
                    return;
                }

                var icmasteBeforeCount = await DbfVerificationHelper.CountRowsByTypeAsync(
                    _dbfReaderService, dbfFolder, IcmasteSchema.TableName, IcmasteSchema.Type, "IN");
                icmasteExport = await ExportStageAsync(
                    "icmaste",
                    IcmasteStage,
                    _icmasteReadResult!,
                    result => _dbfExportService.Export(dbfFolder, IcmasteSchema.TableName, result.Table, IcmasteSchema.Columns));
                await DbfVerificationHelper.VerifyDbfDeltaAsync(
                    _dbfReaderService, IcmasteSchema.TableName, dbfFolder, IcmasteSchema.Type, "IN", icmasteBeforeCount, icmasteExport,
                    v => IcmasteDbfPreview = v, s => IcmasteDbfStatusText = s, b => IcmasteDbfVerified = b);

                var ictraneBeforeCount = await DbfVerificationHelper.CountRowsByTypeAsync(
                    _dbfReaderService, dbfFolder, IctraneSchema.TableName, IctraneSchema.Type, "IN");
                ictraneExport = await ExportStageAsync(
                    "ictrane",
                    IctraneStage,
                    _ictraneReadResult!,
                    result => _dbfExportService.Export(dbfFolder, IctraneSchema.TableName, result.Table, IctraneSchema.Columns));
                await DbfVerificationHelper.VerifyDbfDeltaAsync(
                    _dbfReaderService, IctraneSchema.TableName, dbfFolder, IctraneSchema.Type, "IN", ictraneBeforeCount, ictraneExport,
                    v => IctraneDbfPreview = v, s => IctraneDbfStatusText = s, b => IctraneDbfVerified = b);
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
        }
        catch (Exception ex)
        {
            // Outermost backstop: covers any exception thrown by the setup statements between the
            // gate check and the inner try (e.g. a binding/PropertyChanged subscriber reacting
            // synchronously to one of the property setters above), so the gate below is always
            // released even if the inner try/finally was never reached.
            _fileLogger.LogException("Unhandled error during export", ex);
            AppendLog($"Unexpected error: {ex.Message}");
        }
        finally
        {
            _exportGateService.EndExport();
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
