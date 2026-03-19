using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Comparers;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Helpers;
using NUnit.Framework;

namespace SpreadsheetGen.Tests;

public class CellTests
{
    [Test]
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

        Assert.That(cell, Is.EqualTo(expected).Using(new CellComparer()));
    }

    [Test]
    public void CreateStringCell_ValueExistsInSharedStrings()
    {
        var dict = new Dictionary<string, int>();
        var list = new List<string>();

        CellHelper.CreateSharedStringCell("Hello", dict, list);

        Assert.That(list[0], Is.EqualTo("Hello"));
    }

    [TestCaseSource(nameof(CellTestCases))]
    public void CreateCell_CorrectCellValue(object value, ColumnType columnType, Cell expectedCell)
    {
        var cell = CellHelper.CreateCell(value, columnType);

        Assert.That(expectedCell?.CellValue.Text, Is.EqualTo(cell.CellValue.Text));
    }

    [TestCaseSource(nameof(CellTestCases))]
    public void CreateCell_CorrectCellDataType(object value, ColumnType columnType, Cell expectedCell)
    {
        var cell = CellHelper.CreateCell(value, columnType);

        Assert.That(expectedCell?.DataType, Is.EqualTo(cell.DataType));
    }

    // Date values in excel are index based on OLE Automation date unless formatted
    private static readonly DateTime _OADate = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc);
    private static readonly DateTime _testDateTime = new(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc);
    private static readonly DateOnly _testDateOnly = new(2020, 12, 24);
    private static readonly TimeOnly _testTime = new TimeOnly(14, 30, 5);

    public static IEnumerable<TestCaseData> CellTestCases
    {
        get
        {
            yield return new TestCaseData(123, ColumnType.Integer, new Cell { DataType = CellValues.Number, CellValue = new CellValue("123") });
            yield return new TestCaseData(1_000.15m, ColumnType.Decimal, new Cell { DataType = CellValues.Number, CellValue = new CellValue("1000.15") });
            yield return new TestCaseData(0.25m, ColumnType.Percentage, new Cell { DataType = CellValues.Number, CellValue = new CellValue("0.25") });
            yield return new TestCaseData(_testDateTime, ColumnType.DateTime, new Cell { DataType = null, CellValue = new CellValue((_testDateTime - _OADate).TotalDays.ToString(CultureInfo.InvariantCulture)) });
            yield return new TestCaseData(_testDateOnly, ColumnType.Date, new Cell { DataType = null, CellValue = new CellValue((_testDateOnly.ToDateTime(TimeOnly.MinValue) - _OADate).TotalDays.ToString(CultureInfo.InvariantCulture)) });
            yield return new TestCaseData(_testDateTime, ColumnType.Date, new Cell { DataType = null, CellValue = new CellValue((_testDateTime - _OADate).TotalDays.ToString(CultureInfo.InvariantCulture)) });
            yield return new TestCaseData(_testDateTime, ColumnType.Date, new Cell { DataType = null, CellValue = new CellValue((_testDateTime - _OADate).TotalDays.ToString(CultureInfo.InvariantCulture)) });
            var _testTimeFraction = ((decimal)_testTime.ToTimeSpan().TotalSeconds / 86400m).ToString(CultureInfo.InvariantCulture);
            yield return new TestCaseData(_testTime, ColumnType.Time, new Cell { DataType = null, CellValue = new CellValue(_testTimeFraction) });
            yield return new TestCaseData(true, ColumnType.Boolean, new Cell { DataType = CellValues.Boolean, CellValue = new CellValue(true) });
            yield return new TestCaseData(false, ColumnType.Boolean, new Cell { DataType = CellValues.Boolean, CellValue = new CellValue(false) });
        }
    }
}
