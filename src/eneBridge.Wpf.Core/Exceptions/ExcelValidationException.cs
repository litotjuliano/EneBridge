namespace eneBridge.Wpf.Core.Exceptions;

/// <summary>Thrown when the worksheet does not have the minimum required column count.</summary>
public sealed class ExcelValidationException : Exception
{
    public ExcelValidationException(string message) : base(message)
    {
    }
}
