using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace SpreadsheetGen.Tests;

public class CellTests
{

    [Fact]
    public void CreateStringCell()
    {
        var dict = new Dictionary<string, int>();
        var list = new List<string>();
        var cell = CellGenerator.CreateStringCell("Hello", dict, list);

        var expected = new Cell()
        {
            DataType = CellValues.SharedString,
            CellValue = new CellValue("Hello")
        };

        Assert.Equal("Hello", list[0]);
        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreateIntegerCell()
    {
        var cell = CellGenerator.CreateIntegerCell(123);

        var expected = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue("123")
        };

        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreateDecimalCell()
    {
        var cell = CellGenerator.CreateIntegerCell(123.15m);

        var expected = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue("123.15")
        };

        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreatePercentageCell()
    {
        var cell = CellGenerator.CreatePercentageCell(0.25);

        var expected = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue("25")
        };

        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreateCurrencyCell_ThrowsNotImplemented()
    {
        Assert.Throws<NotImplementedException>(() => CellGenerator.CreateCurrencyCell(100));
    }

    [Fact]
    public void CreateDateTimeCell()
    {
        var datetime = new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc);
        var cell = CellGenerator.CreateDateTimeCell(datetime);

        var expected = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue("2020-12-24 23:59")
        };

        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreateDateCell_FromDateOnly()
    {
        var date = new DateOnly(2020, 12, 24);
        var cell = CellGenerator.CreateDateTimeCell(date);

        var expected = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue("2020-12-24")
        };

        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreateDateCell_FromDateTime()
    {
        var datetime = new DateTime(2020, 12, 24, 0, 0, 0, DateTimeKind.Utc);
        var cell = CellGenerator.CreateDateTimeCell(datetime);

        var expected = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue("2020-12-24")
        };

        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreateTimeCell()
    {
        var time = new TimeOnly(14, 30, 15);
        var cell = CellGenerator.CreateTimeCell(time);

        var expected = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue("14:30:15")
        };

        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreateBooleanCell_TrueValue()
    {
        var cell = CellGenerator.CreateBooleanCell(true);

        var expected = new Cell()
        {
            DataType = CellValues.Boolean,
            CellValue = new CellValue("TRUE")
        };

        Assert.Equal(expected, cell);
    }

    [Fact]
    public void CreateBooleanCell_FalseValue()
    {
        var cell = CellGenerator.CreateBooleanCell(false);

        var expected = new Cell()
        {
            DataType = CellValues.Boolean,
            CellValue = new CellValue("FALSE")
        };

        Assert.Equal(expected, cell);
    }
}
