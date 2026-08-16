using CommunityToolkit.Mvvm.ComponentModel;

namespace eneBridge.Wpf.ViewModels;

public partial class StageProgressViewModel : ObservableObject
{
    [ObservableProperty]
    private string _name;

    [ObservableProperty]
    private bool _isRunning;

    [ObservableProperty]
    private string _statusText = "Not run yet";

    [ObservableProperty]
    private bool? _success;

    public StageProgressViewModel(string name)
    {
        _name = name;
    }

    public void Reset()
    {
        IsRunning = true;
        StatusText = "Running...";
        Success = null;
    }

    public void Fail(string message)
    {
        IsRunning = false;
        Success = false;
        StatusText = message;
    }

    public void Complete(bool success, string statusText)
    {
        IsRunning = false;
        Success = success;
        StatusText = statusText;
    }
}
