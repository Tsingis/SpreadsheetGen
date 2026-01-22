using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using NUnit.Framework;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Models;

namespace SpreadsheetGen.Tests;

public class SheetDataTests
{
    private readonly DateTime _OADate = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc);

    [Test]
    public void Spreadsheet_ExpectedRowCount()
    {
        var sheetData = GetSheetData();

        var rows = sheetData.Elements<Row>().ToList();

        Assert.That(rows, Has.Count.EqualTo(3));
    }

    [Test]
    public void Spreadsheet_ExpectedColumnCount()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().First();
        var cells = row.Elements<Cell>().ToList();

        Assert.That(cells, Has.Count.EqualTo(8));
    }

    [Test]
    public void Spreadsheet_HeaderRow_HasExpectedSharedStringIndices()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().First();
        var stringIndices = row.Elements<Cell>().Select(x => int.Parse(x.CellValue?.Text, CultureInfo.InvariantCulture)).ToList();

        var expected = Enumerable.Range(0, 8).ToList();

        Assert.That(stringIndices, Is.EqualTo(expected));
    }

    [Test]
    public void Spreadsheet_FirstDataRow_FirstCell_HasExpectedSharedStringIndex()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(0);
        var stringIndex = int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);

        Assert.That(stringIndex, Is.EqualTo(8));
    }

    [Test]
    public void Spreadsheet_FirstDataRow_SecondCell_HasNumber()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(1);
        var value = int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);

        Assert.That(value, Is.EqualTo(1));
    }

    [Test]
    public void Spreadsheet_FirstDataRow_ThirdCell_HasDecimal()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(2);
        var value = decimal.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);

        Assert.That(value, Is.EqualTo(50_000.01m));
    }

    [Test]
    public void Spreadsheet_FirstDataRow_FourthCell_HasDateTime()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(3);
        var expected = new DateTime(2020, 1, 15, 0, 0, 0, DateTimeKind.Utc);
        var actual = _OADate.AddDays(int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture));

        Assert.That(actual, Is.EqualTo(expected));
    }

    [Test]
    public void Spreadsheet_FirstDataRow_FifthCell_HasBoolean()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(4);
        var value = bool.Parse(cell.CellValue.Text);

        Assert.That(value);
    }

    [Test]
    public void Spreadsheet_FirstDataRow_SixthCell_HasPercentage()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(5);
        var value = decimal.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);

        Assert.That(value, Is.EqualTo(1.2345m));
    }

    [Test]
    public void Spreadsheet_FirstDataRow_SeventhCell_HasDateTime()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(6);
        var expected = new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc);
        var actual = _OADate.AddDays(double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture));

        Assert.That(expected, Is.EqualTo(actual).Within(TimeSpan.FromMicroseconds(1)));
    }

    [Test]
    public void Spreadsheet_FirstDataRow_EighthCell_HasTime()
    {
        var sheetData = GetSheetData();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(7);

        var expected = new TimeOnly(12, 15, 22);
        var actual = TimeOnly.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);

        Assert.That(actual, Is.EqualTo(expected));
    }

    private static SheetData GetSheetData()
    {
        var data = CreateSampleData();
        using var ms = new MemoryStream();
        using SpreadsheetDocument spreadsheet = data.ToSpreadsheetDocument(ms);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        return sheetData;
    }

    private static WorksheetData CreateSampleData()
    {
        List<Models.Column> columns =
        [
            new() { Name = "Text", Type = ColumnType.Text },
            new() { Name = "Integer", Type = ColumnType.Integer },
            new() { Name = "Decimal", Type = ColumnType.Decimal },
            new() { Name = "Date", Type = ColumnType.Date },
            new() { Name = "Boolean", Type = ColumnType.Boolean },
            new() { Name = "Percentage", Type = ColumnType.Percentage },
            new() { Name = "DateTime", Type = ColumnType.DateTime },
            new() { Name = "Time", Type = ColumnType.Time },
        ];

        List<object[]> rows =
        [
            [
                "John Doe",
                1,
                50_000.01m,
                new DateTime(2020, 1, 15, 0, 0, 0, DateTimeKind.Utc),
                true,
                1.2345m,
                new DateTime(2020, 12, 24, 23, 59, 59,
                DateTimeKind.Utc),
                new TimeOnly(12, 15, 22)
            ],
            [
                "Jane 'Very Long Name' Smith",
                2,
                60_500.50m,
                new DateTime(2021, 5, 10, 0, 5, 0, DateTimeKind.Utc),
                false,
                0.25m,
                new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc),
                new TimeOnly(15, 44, 0, 123)
            ],
        ];

        return new WorksheetData
        {
            Columns = columns,
            Rows = rows
        };
    }
}
