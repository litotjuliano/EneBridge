using eneBridge.Wpf.Core.Models;
using eneBridge.Wpf.Core.Services;

namespace eneBridge.Wpf.Core.Tests;

public class DbfSchemaBuilderTests
{
    [Theory]
    [InlineData(typeof(string), 40, "CHAR(40)")]
    [InlineData(typeof(DateTime), 0, "DATE")]
    [InlineData(typeof(decimal), 0, "NUMERIC(20, 2)")]
    [InlineData(typeof(bool), 0, "LOGICAL")]
    public void GetDbfColumnType_MapsKnownTypes(Type clrType, int maxLength, string expected)
    {
        Assert.Equal(expected, DbfSchemaBuilder.GetDbfColumnType(clrType, maxLength));
    }

    [Fact]
    public void GetDbfColumnType_ThrowsForUnsupportedType()
    {
        Assert.Throws<NotSupportedException>(() => DbfSchemaBuilder.GetDbfColumnType(typeof(int), 0));
    }

    [Theory]
    [InlineData("REF", "REF")]
    [InlineData("BILLDATE10", "BILLDATE10")]
    [InlineData("this-has-more-than-ten-chars", "thishasm")]
    public void SanitizeColumnName_TruncatesAndStripsSpecialChars(string input, string expected)
    {
        Assert.Equal(expected, DbfSchemaBuilder.SanitizeColumnName(input));
    }

    [Fact]
    public void BuildCreateTableSql_IcmasteSchema_HasNoNameCollisions()
    {
        var sql = DbfSchemaBuilder.BuildCreateTableSql(IcmasteSchema.TableName, IcmasteSchema.Columns);
        Assert.StartsWith("CREATE TABLE icmaste (", sql);
        Assert.EndsWith(")", sql);
    }

    [Fact]
    public void BuildCreateTableSql_IctraneSchema_HasNoNameCollisions()
    {
        var sql = DbfSchemaBuilder.BuildCreateTableSql(IctraneSchema.TableName, IctraneSchema.Columns);
        Assert.StartsWith("CREATE TABLE ictrane (", sql);
        Assert.EndsWith(")", sql);
    }

    [Fact]
    public void BuildCreateTableSql_DuplicateNamesDifferingOnlyByCase_ThrowsInvalidOperationException()
    {
        var columns = new List<DbfColumnDefinition>
        {
            new("MyField", typeof(string), 10),
            new("myfield", typeof(string), 10),
        };

        var ex = Assert.Throws<InvalidOperationException>(() => DbfSchemaBuilder.BuildCreateTableSql("test", columns));

        Assert.Contains("Column name collision", ex.Message);
    }

    [Fact]
    public void BuildCreateTableSql_NamesCollideAfterTenCharTruncation_ThrowsInvalidOperationException()
    {
        var columns = new List<DbfColumnDefinition>
        {
            new("ABCDEFGHIJKLMNOP", typeof(string), 10),
            new("ABCDEFGHIJQRSTUV", typeof(string), 10),
        };

        var ex = Assert.Throws<InvalidOperationException>(() => DbfSchemaBuilder.BuildCreateTableSql("test", columns));

        Assert.Contains("Column name collision", ex.Message);
    }
}
