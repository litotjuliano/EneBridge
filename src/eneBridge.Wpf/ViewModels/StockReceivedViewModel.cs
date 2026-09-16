using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using ClosedXML.Excel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using eneBridge.Wpf.Core.Exceptions;
using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;
using eneBridge.Wpf.Services;

namespace eneBridge.Wpf.ViewModels;

/// <summary>
/// Stock Received workflow — reads Self-Billed_Format-ORI.xlsx-shaped workbooks and writes into
/// the SAME icmaste.dbf/ictrane.dbf Invoice writes to (TYPE="RE" vs Invoice's "IN"), appending
/// rather than recreating. Mirrors MainViewModel's Invoice-related surface; see
/// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md for the full rationale.
/// </summary>
public partial class StockReceivedViewModel : ObservableObject
{
    private readonly ExcelReaderService _excelReaderService;
    private readonly DbfExportService _dbfExportService;
    private readonly DbfReaderService _dbfReaderService;
    private readonly FoxProDbfReader _foxProDbfReader;
    private readonly DbfSafetyBackupService _dbfSafetyBackupService;
    private readonly SettingsService _settingsService;
    private readonly RunHistoryService _runHistoryService;
    private readonly FileLogger _fileLogger;
    private readonly ExcelSourceStagingService _excelSourceStagingService;
    private readonly ExportGateService _exportGateService;
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

    [ObservableProperty]
    private bool _showAllColumns;

    /// <summary>
    /// Reflects the SAME ACE-driver health check MainViewModel runs once per session (there's only
    /// one machine-level driver to check, not one per tab) — set via SetAceEngineHealth, called by
    /// MainViewModel whenever its own check completes.
    /// </summary>
    [ObservableProperty]
    private bool _aceEngineHealthy = true;

    [ObservableProperty]
    private string _aceEngineStatusText = string.Empty;

    public ObservableCollection<RunHistoryEntry> RunHistory { get; } = [];
    public ObservableCollection<string> LogLines { get; } = [];

    public StageProgressViewModel IcmasteStage { get; } = new("icmaste");
    public StageProgressViewModel IctraneStage { get; } = new("ictrane");

    public StockReceivedViewModel(
        ExcelReaderService excelReaderService,
        DbfExportService dbfExportService,
        DbfReaderService dbfReaderService,
        FoxProDbfReader foxProDbfReader,
        DbfSafetyBackupService dbfSafetyBackupService,
        SettingsService settingsService,
        RunHistoryService runHistoryService,
        FileLogger fileLogger,
        ExcelSourceStagingService excelSourceStagingService,
        ExportGateService exportGateService,
        string appBaseDirectory)
    {
        _excelReaderService = excelReaderService;
        _dbfExportService = dbfExportService;
        _dbfReaderService = dbfReaderService;
        _foxProDbfReader = foxProDbfReader;
        _dbfSafetyBackupService = dbfSafetyBackupService;
        _settingsService = settingsService;
        _runHistoryService = runHistoryService;
        _fileLogger = fileLogger;
        _excelSourceStagingService = excelSourceStagingService;
        _exportGateService = exportGateService;
        _appBaseDirectory = appBaseDirectory;

        _exportGateService.StateChanged += () => ConfirmExportCommand.NotifyCanExecuteChanged();
    }

    public void Initialize()
    {
        var (excelPath, dbfFolder) = _settingsService.ResolveEffectiveStockReceivedPaths(_appBaseDirectory);
        ExcelFilePath = excelPath;
        DbfFolderPath = dbfFolder;

        // Most recent run first.
        foreach (var entry in _runHistoryService.LoadAll().AsEnumerable().Reverse())
        {
            RunHistory.Add(entry);
        }
    }

    /// <summary>Called by MainViewModel whenever the shared ACE-driver health check result changes.</summary>
    public void SetAceEngineHealth(bool healthy)
    {
        AceEngineHealthy = healthy;
        AceEngineStatusText = healthy
            ? string.Empty
            : "Access Database Engine isn't working correctly — Confirm & Export is disabled until this is fixed.";
        ConfirmExportCommand.NotifyCanExecuteChanged();
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

    private const string StockReceivedStagedFileName = "stockreceived.xlsx";

    /// <summary>
    /// Copies the picked Excel file into the local staging folder so later runs don't depend on
    /// removable/network media staying plugged in. Never blocks the user: if staging fails for any
    /// reason, the failure is logged and the original picked path is used as-is.
    /// </summary>
    private string StageExcelSource(string originalPath)
    {
        try
        {
            var stagedPath = _excelSourceStagingService.StageFile(originalPath, StockReceivedStagedFileName);
            _fileLogger.LogInfo($"Imported Stock Received Excel source '{originalPath}' -> '{stagedPath}'");
            return stagedPath;
        }
        catch (Exception ex)
        {
            _fileLogger.LogException("Failed to stage Stock Received Excel source file", ex);
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
            s.StockReceivedExcelFilePath = ExcelFilePath;
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

                    try
                    {
                        _excelReaderService.ValidateFileFormat(worksheet, ExcelFileFormat.StockReceived);
                    }
                    catch (ExcelValidationException ex)
                    {
                        MessageBox.Show(ex.Message, "eneBridge - Wrong File Type", MessageBoxButton.OK, MessageBoxImage.Warning);
                        AppendLog($"Cannot preview: {ex.Message}");
                        IcmasteStage.Fail(ex.Message);
                        IctraneStage.Fail(ex.Message);
                        return;
                    }

                    _icmasteReadResult = await ReadStageAsync(
                        "icmaste",
                        IcmasteStage,
                        () => _excelReaderService.ReadStockReceivedIcmaste(worksheet),
                        result => IcmastePreview = result.Table.DefaultView);

                    _ictraneReadResult = await ReadStageAsync(
                        "ictrane",
                        IctraneStage,
                        () => _excelReaderService.ReadStockReceivedIctrane(worksheet),
                        result => IctranePreview = result.Table.DefaultView);
                }
            }
        }
        catch (Exception ex)
        {
            // Outer backstop: nothing should ever crash the app.
            _fileLogger.LogException("Unhandled error during Stock Received preview", ex);
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
            AppendLog("Cannot export: the Invoice tab is currently exporting. Please wait for it to finish and try again.");
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

                var icmasteCandidates = _icmasteReadResult!.Table.AsEnumerable()
                    .Select(row => (Ref: row[IcmasteSchema.Ref] as string ?? string.Empty, Code: row[IcmasteSchema.Code] as string ?? string.Empty))
                    .ToList();
                var duplicateRefs = await DuplicateDocumentChecker.FindDuplicateRefsAsync(_foxProDbfReader, dbfFolder, "RE", icmasteCandidates);
                if (!DbfWorkflowHelper.ConfirmNoDuplicateDocuments(duplicateRefs))
                {
                    // Deliberately NOT cancelledByUser: this is the system refusing to proceed
                    // after detecting a real problem, not a user backing out of an ambiguous
                    // situation (contrast the table-check cancel above). It still gets a Run
                    // History entry, marked as a failure, so a blocked export leaves a record.
                    IcmasteDbfStatusText = "Run Confirm & Export to verify.";
                    IctraneDbfStatusText = "Run Confirm & Export to verify.";
                    fatalError = $"Blocked: {duplicateRefs.Count} duplicate document(s) already exist (see log)";
                    AppendLog($"Export blocked by the duplicate-document check — already exist for this supplier/customer: {string.Join(", ", duplicateRefs)}");
                    return;
                }

                icmasteExport = await ExportStageAsync(
                    "icmaste",
                    IcmasteStage,
                    _icmasteReadResult!,
                    result => _dbfExportService.Export(dbfFolder, IcmasteSchema.TableName, result.Table, IcmasteSchema.Columns));
                await DbfVerificationHelper.VerifyDbfAsync(
                    _dbfReaderService, IcmasteSchema.TableName, dbfFolder, IcmasteSchema.Type, "RE", icmasteExport,
                    v => IcmasteDbfPreview = v, s => IcmasteDbfStatusText = s, b => IcmasteDbfVerified = b);

                ictraneExport = await ExportStageAsync(
                    "ictrane",
                    IctraneStage,
                    _ictraneReadResult!,
                    result => _dbfExportService.Export(dbfFolder, IctraneSchema.TableName, result.Table, IctraneSchema.Columns));
                await DbfVerificationHelper.VerifyDbfAsync(
                    _dbfReaderService, IctraneSchema.TableName, dbfFolder, IctraneSchema.Type, "RE", ictraneExport,
                    v => IctraneDbfPreview = v, s => IctraneDbfStatusText = s, b => IctraneDbfVerified = b);
            }
            catch (Exception ex)
            {
                // Outer backstop: nothing should ever crash the app.
                _fileLogger.LogException("Unhandled error during Stock Received export", ex);
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
            _fileLogger.LogException("Unhandled error during Stock Received export", ex);
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

        foreach (var skipReason in readResult.SkipReasons)
        {
            AppendLog($"[{label}] Row {skipReason.ExcelRow}: {skipReason.Reason}");
        }

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
            _fileLogger.LogException("Failed to open Stock Received workbook", ex);
            return (null, ex.Message);
        }
    }

    private void AppendLog(string message) => LogLines.Add($"{DateTime.Now:HH:mm:ss} {message}");
}
