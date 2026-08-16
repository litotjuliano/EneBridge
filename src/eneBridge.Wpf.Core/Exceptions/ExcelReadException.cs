namespace eneBridge.Wpf.Core.Exceptions;

/// <summary>Thrown when the Excel workbook itself cannot be opened (missing, locked, corrupt).</summary>
public sealed class ExcelReadException : Exception
{
    public ExcelReadException(string message, Exception? innerException = null)
        : base(message, innerException)
    {
    }
}
