using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Comparers;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Helpers;
using Xunit;

namespace SpreadsheetGen.Tests;

public class CellTests
{
    [Fact]
    public void CreateStringCell_CorrectIndexValue()
    {
        var dict = new Dictionary<string, int>();
        var list = new List<string>();
        var cell = CellHelper.CreateSharedStringCell("Hello", dict, list);

        var expected = new Cell()
        {
            DataType = CellValues.SharedString,
            CellValue = new CellValue("0")
        };

        Assert.Equal(expected, cell, new CellComparer());
    }

    [Fact]
    public void CreateStringCell_ValueExistsInSharedStrings()
    {
        var dict = new Dictionary<string, int>();
        var list = new List<string>();

        CellHelper.CreateSharedStringCell("Hello", dict, list);

        Assert.Equal("Hello", list[0]);
    }


    [Theory]
    [MemberData(nameof(CellTestCases))]
    public void CreateCell_CorrectCellValue(object value, ColumnType columnType, Cell expectedCell)
    {
        var cell = CellHelper.CreateCell(value, columnType);

        Assert.Equal(expectedCell?.CellValue.Text, cell.CellValue.Text);
    }

    [Theory]
    [MemberData(nameof(CellTestCases))]
    public void CreateCell_CorrectCellDataType(object value, ColumnType columnType, Cell expectedCell)
    {
        var cell = CellHelper.CreateCell(value, columnType);

        Assert.Equal(expectedCell?.DataType, cell.DataType);
    }

    // Date values in excel are index based on OLE Automation date unless formatted
    private static readonly DateTime _OADate = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc);
    private static readonly DateTime _testDateTime = new(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc);
    private static readonly DateOnly _testDateOnly = new(2020, 12, 24);

    public static TheoryData<object, ColumnType, Cell> CellTestCases => new()
    {
        {
            123,
            ColumnType.Integer,
            new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue("123")
            }
        },
        {
            1_000.15m,
            ColumnType.Decimal,
            new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue("1000.15")
            }
        },
        {
            0.25m,
            ColumnType.Percentage,
            new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue("0.25")
            }
        },
        {
            _testDateTime,
            ColumnType.DateTime,
            new Cell
            {
                DataType = null,
                CellValue = new CellValue((_testDateTime - _OADate).TotalDays.ToString(CultureInfo.InvariantCulture))
            }
        },
        {
            _testDateOnly,
            ColumnType.Date,
            new Cell
            {
                DataType = null,
                CellValue = new CellValue((_testDateOnly.ToDateTime(TimeOnly.MinValue) - _OADate).TotalDays.ToString(CultureInfo.InvariantCulture))
            }
        },
        {
            _testDateTime,
            ColumnType.Date,
            new Cell
            {
                DataType = null,
                CellValue = new CellValue((_testDateTime - _OADate).TotalDays.ToString(CultureInfo.InvariantCulture))
            }
        },
        {
            _testDateTime,
            ColumnType.Date,
            new Cell
            {
                DataType = null,
                CellValue = new CellValue((_testDateTime - _OADate).TotalDays.ToString(CultureInfo.InvariantCulture))
            }
        },
        {
            new TimeOnly(14, 30, 5),
            ColumnType.Time,
            new Cell
            {
                DataType = CellValues.String,
                CellValue = new CellValue("14:30:05")
            }
        },
        {
            true,
            ColumnType.Boolean,
            new Cell
            {
                DataType = CellValues.Boolean,
                CellValue = new CellValue(true)
            }
        },
        {
            false,
            ColumnType.Boolean,
            new Cell
            {
                DataType = CellValues.Boolean,
                CellValue = new CellValue(false)
            }
        },
    };
}
