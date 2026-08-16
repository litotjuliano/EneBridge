namespace eneBridge.Wpf.Core.Models;

/// <summary>Deserialized shape of appsettings.json's "FilePaths" section (initial defaults).</summary>
public sealed class AppSettingsModel
{
    public FilePathsSection FilePaths { get; init; } = new();

    public sealed class FilePathsSection
    {
        public string ExcelFilePath { get; init; } = string.Empty;
        public string DbfFilePath { get; init; } = string.Empty;
    }
}

/// <summary>Persisted last-used paths, saved to %AppData%\eneBridge\settings.json.</summary>
public sealed class UserSettingsModel
{
    public string? ExcelFilePath { get; set; }
    public string? DbfFolderPath { get; set; }
}
