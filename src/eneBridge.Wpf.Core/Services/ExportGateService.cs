namespace eneBridge.Wpf.Core.Services;

/// <summary>
/// Coordinates DBF export across the Invoice and Stock Received workflows, which append into the
/// SAME icmaste.dbf/ictrane.dbf files (see docs/superpowers/specs/2026-09-14-stock-received-workflow-design.md).
/// A single shared instance, injected into both ViewModels, prevents them from calling
/// DbfExportService.Export / DbfReaderService.Read against the same tables at the same time — a
/// real crash risk given this driver's documented native-crash fragility under far milder stress
/// (see CLAUDE.md's "Known reliability risk").
/// </summary>
public sealed class ExportGateService
{
    private readonly object _lock = new();
    private bool _exportInProgress;

    /// <summary>Raised whenever the gate is claimed or released, so callers can re-evaluate command CanExecute state.</summary>
    public event Action? StateChanged;

    public bool IsExportInProgress
    {
        get { lock (_lock) { return _exportInProgress; } }
    }

    /// <summary>Attempts to claim the gate for an export. Returns false if another export already holds it.</summary>
    public bool TryBeginExport()
    {
        lock (_lock)
        {
            if (_exportInProgress)
            {
                return false;
            }
            _exportInProgress = true;
        }
        StateChanged?.Invoke();
        return true;
    }

    /// <summary>Releases the gate. Must be called exactly once for every TryBeginExport() that returned true.</summary>
    public void EndExport()
    {
        lock (_lock)
        {
            _exportInProgress = false;
        }
        StateChanged?.Invoke();
    }
}
