namespace eneBridge.Wpf.Core.Services;

internal static class TextTruncation
{
    public static string Truncate(string value, int maxLength)
    {
        if (maxLength <= 0 || value.Length <= maxLength)
        {
            return value;
        }
        return value[..maxLength];
    }
}
