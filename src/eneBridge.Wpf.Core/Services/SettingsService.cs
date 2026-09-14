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

    /// <summary>
    /// The DBF folder is remembered across restarts (user-saved wins over the appsettings.json
    /// default, combined with the app's base directory) -- it rarely changes and isn't a
    /// data-freshness risk. The Excel file path is deliberately NEVER restored, even though
    /// UserSettingsModel.ExcelFilePath is still written on every Browse (SaveUserPaths) -- a
    /// remembered Excel path looks "ready to export" on next launch, inviting an accidental
    /// re-export of a stale/previous file instead of the user actively picking today's file. The
    /// field always starts blank; Preview/Confirm &amp; Export stay disabled until a fresh Browse.
    /// </summary>
    public (string excelPath, string dbfFolder) ResolveEffectivePaths(string appBaseDirectory)
    {
        var defaults = LoadAppDefaults();
        var user = LoadUserSettings();

        string dbfFolder = !string.IsNullOrWhiteSpace(user?.DbfFolderPath)
            ? user!.DbfFolderPath!
            : Path.Combine(appBaseDirectory, defaults.FilePaths.DbfFilePath);

        return (string.Empty, dbfFolder);
    }

    /// <summary>
    /// Stock Received's own Excel path -- always blank on startup, for the same reason
    /// ResolveEffectivePaths never restores Invoice's (see its doc comment) -- but the SAME DBF
    /// folder Invoice uses, since both workflows now write into the same live tables (see
    /// docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md).
    /// </summary>
    public (string excelPath, string dbfFolder) ResolveEffectiveStockReceivedPaths(string appBaseDirectory)
    {
        var defaults = LoadAppDefaults();
        var user = LoadUserSettings();

        string dbfFolder = !string.IsNullOrWhiteSpace(user?.DbfFolderPath)
            ? user!.DbfFolderPath!
            : Path.Combine(appBaseDirectory, defaults.FilePaths.DbfFilePath);

        return (string.Empty, dbfFolder);
    }
}
