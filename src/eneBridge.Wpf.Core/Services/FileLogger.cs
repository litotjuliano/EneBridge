namespace eneBridge.Wpf.Core.Services;

/// <summary>Rolling text log for full exception detail, separate from the concise on-screen run history.</summary>
public sealed class FileLogger
{
    private readonly string _logDirectory;

    public FileLogger(string? userDataDirectory = null)
    {
        var dataDir = userDataDirectory ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "eneBridge");
        _logDirectory = Path.Combine(dataDir, "logs");
        Directory.CreateDirectory(_logDirectory);
    }

    public void LogException(string context, Exception ex) =>
        Append($"{context}{Environment.NewLine}{ex}");

    public void LogInfo(string message) => Append(message);

    private void Append(string message)
    {
        var logFile = Path.Combine(_logDirectory, $"eneBridge-{DateTime.Now:yyyyMMdd}.log");
        File.AppendAllText(logFile, $"[{DateTime.Now:O}] {message}{Environment.NewLine}");
    }
}
