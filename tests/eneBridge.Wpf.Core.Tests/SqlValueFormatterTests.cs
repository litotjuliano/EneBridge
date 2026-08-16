using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class SqlValueFormatterTests
{
    [Fact]
    public void Format_Null_ReturnsNull() =>
        Assert.Equal("NULL", SqlValueFormatter.Format(DBNull.Value, typeof(string)));

    [Fact]
    public void Format_String_EscapesSingleQuotes() =>
        Assert.Equal("'O''Brien'", SqlValueFormatter.Format("O'Brien", typeof(string)));

    [Fact]
    public void Format_DateTime_UsesIsoDate() =>
        Assert.Equal("'2025-12-15'", SqlValueFormatter.Format(new DateTime(2025, 12, 15), typeof(DateTime)));

    [Fact]
    public void Format_BoolTrue_ReturnsDotTDot() =>
        Assert.Equal(".T.", SqlValueFormatter.Format(true, typeof(bool)));

    [Fact]
    public void Format_BoolFalse_ReturnsDotFDot() =>
        Assert.Equal(".F.", SqlValueFormatter.Format(false, typeof(bool)));

    [Fact]
    public void Format_Decimal_ReturnsInvariantString() =>
        Assert.Equal("307.7", SqlValueFormatter.Format(307.7m, typeof(decimal)));
}
