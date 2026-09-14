using System.Text.Json;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Reads appsettings.json (initial defaults) and reads/writes a persisted user settings file
/// (%AppData%\eneBridge\settings.json) that stores the last-used Excel/DBF paths.
/// </summary>
public sealed class SettingsService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _appSettingsPath;
    private readonly string _userSettingsPath;

    public SettingsService(string appBaseDirectory, string? userDataDirectory = null)
    {
        _appSettingsPath = Path.Combine(appBaseDirectory, "appsettings.json");
        var dataDir = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "eneBridge");
        Directory.CreateDirectory(dataDir);
        _userSettingsPath = Path.Combine(dataDir, "settings.json");
    }

    public AppSettingsModel LoadAppDefaults()
    {
        if (!File.Exists(_appSettingsPath))
        {
            return new AppSettingsModel();
        }
        try
        {
            var json = File.ReadAllText(_appSettingsPath);
            return JsonSerializer.Deserialize<AppSettingsModel>(json, JsonOptions) ?? new AppSettingsModel();
        }
        catch (Exception)
        {
            return new AppSettingsModel();
        }
    }

    public UserSettingsModel? LoadUserSettings()
    {
        if (!File.Exists(_userSettingsPath))
        {
            return null;
        }
        try
        {
            var json = File.ReadAllText(_userSettingsPath);
            return JsonSerializer.Deserialize<UserSettingsModel>(json, JsonOptions);
        }
        catch (Exception)
        {
            return null;
        }
    }

    /// <summary>
    /// Loads whatever is currently persisted (or starts fresh if nothing is saved yet), applies
    /// `update`, and writes the merged result back. Merge-based rather than replace-based so that
    /// Invoice and Stock Received -- which each save only the fields they own -- never clobber
    /// each other's saved paths.
    /// </summary>
    public void SaveUserSettings(Action<UserSettingsModel> update)
    {
        var current = LoadUserSettings() ?? new UserSettingsModel();
        update(current);

        var json = JsonSerializer.Serialize(current, JsonOptions);
        var tempPath = _userSettingsPath + ".tmp";
        File.WriteAllText(tempPath, json);
        File.Move(tempPath, _userSettingsPath, overwrite: true);
    }

    /// <summary>User-saved paths win over appsettings.json defaults (which are combined with the app's base directory, same as the original).</summary>
    public (string excelPath, string dbfFolder) ResolveEffectivePaths(string appBaseDirectory)
    {
        var defaults = LoadAppDefaults();
        var user = LoadUserSettings();

        string excelPath = !string.IsNullOrWhiteSpace(user?.ExcelFilePath)
            ? user!.ExcelFilePath!
            : Path.Combine(appBaseDirectory, defaults.FilePaths.ExcelFilePath);

        string dbfFolder = !string.IsNullOrWhiteSpace(user?.DbfFolderPath)
            ? user!.DbfFolderPath!
            : Path.Combine(appBaseDirectory, defaults.FilePaths.DbfFilePath);

        return (excelPath, dbfFolder);
    }

    /// <summary>
    /// Stock Received's own Excel path, resolved independently of Invoice's -- but the SAME DBF
    /// folder Invoice uses, since both workflows now write into the same live tables (see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md). Unlike
    /// ResolveEffectivePaths, an unset default yields an empty string rather than a combined path
    /// pointing at the app's own base directory -- no app-shipped default is required, the field
    /// just starts blank until the user Browses once.
    /// </summary>
    public (string excelPath, string dbfFolder) ResolveEffectiveStockReceivedPaths(string appBaseDirectory)
    {
        var defaults = LoadAppDefaults();
        var user = LoadUserSettings();

        string excelPath = !string.IsNullOrWhiteSpace(user?.StockReceivedExcelFilePath)
            ? user!.StockReceivedExcelFilePath!
            : (string.IsNullOrWhiteSpace(defaults.FilePaths.StockReceivedExcelFilePath)
                ? string.Empty
                : Path.Combine(appBaseDirectory, defaults.FilePaths.StockReceivedExcelFilePath));

        string dbfFolder = !string.IsNullOrWhiteSpace(user?.DbfFolderPath)
            ? user!.DbfFolderPath!
            : Path.Combine(appBaseDirectory, defaults.FilePaths.DbfFilePath);

        return (excelPath, dbfFolder);
    }
}
