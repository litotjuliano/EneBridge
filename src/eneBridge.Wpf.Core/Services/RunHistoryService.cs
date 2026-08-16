using System.Text.Json;
using eneBridge.Wpf.Core.Models;

namespace eneBridge.Wpf.Core.Services;

/// <summary>Persists run history to %AppData%\eneBridge\runHistory.json so it survives restarts.</summary>
public sealed class RunHistoryService
{
    private static readonly JsonSerializerOptions JsonOptions = new() { WriteIndented = true };

    private readonly string _historyPath;

    public RunHistoryService(string? userDataDirectory = null)
    {
        var dataDir = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "eneBridge");
        Directory.CreateDirectory(dataDir);
        _historyPath = Path.Combine(dataDir, "runHistory.json");
    }

    public List<RunHistoryEntry> LoadAll()
    {
        if (!File.Exists(_historyPath))
        {
            return [];
        }
        try
        {
            var json = File.ReadAllText(_historyPath);
            return JsonSerializer.Deserialize<List<RunHistoryEntry>>(json, JsonOptions) ?? [];
        }
        catch (Exception)
        {
            return [];
        }
    }

    public void Append(RunHistoryEntry entry)
    {
        var all = LoadAll();
        all.Add(entry);
        var json = JsonSerializer.Serialize(all, JsonOptions);
        var tempPath = _historyPath + ".tmp";
        File.WriteAllText(tempPath, json);
        File.Move(tempPath, _historyPath, overwrite: true);
    }
}
