using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Globalization;
using Xunit;
using SpreadsheetGen.Models;
using SpreadsheetGen.Enums;

namespace SpreadsheetGen.Tests;

public class SheetContentTests
{
    private readonly DateTime _epochDate = new(1899, 12, 30, 0, 0, 0, DateTimeKind.Utc);

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_ExpectedRowCount(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var rows = sheetData.Elements<Row>().ToList();

        Assert.Equal(3, rows.Count);
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_ExpectedColumnCount(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        Row firstRow = sheetData.Elements<Row>().First();
        List<Cell> cells = firstRow.Elements<Cell>().ToList();

        Assert.Equal(8, cells.Count);
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_HeaderRow_Contains_Names(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().First();
        var sharedStringTable = spreadsheet.WorkbookPart.SharedStringTablePart.SharedStringTable;
        var cellValues = row.Elements<Cell>()
            .Select(cell =>
                cell.DataType != null && cell.DataType == CellValues.SharedString
                    ? sharedStringTable.ElementAt(int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture)).InnerText
                    : cell.CellValue?.Text
            )
            .ToList();

        List<string> expected = ["Integer", "Text", "Decimal", "Date", "Boolean", "Percentage", "DateTime", "Time"];

        Assert.Equal(expected, cellValues);
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_FirstDataRow_FirstCell_HasNumber(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(0);

        Assert.Equal(1, int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture));
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_FirstDataRow_SecondCell_HasString(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(1);
        var stringIndex = int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);
        var sharedStringTable = spreadsheet.WorkbookPart.SharedStringTablePart.SharedStringTable;
        var value = sharedStringTable.Elements<SharedStringItem>().ElementAt(stringIndex).InnerText;

        Assert.Equal("John Doe", value);
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_FirstDataRow_ThirdCell_HasDecimal(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(2);

        Assert.Equal(50_000.01m, decimal.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture));
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_FirstDataRow_FourthCell_HasDateTime(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(3);
        var expected = new DateTime(2020, 1, 15, 0, 0, 0, DateTimeKind.Utc);
        var actual = _epochDate.AddDays(int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture));

        Assert.Equal(expected, actual);
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_FirstDataRow_FifthCell_HasBoolean(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(4);

        Assert.True(bool.Parse(cell.CellValue.Text));
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_FirstDataRow_SixthCell_HasPercentage(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(5);

        Assert.Equal(1.2345m, decimal.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture));
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_FirstDataRow_SeventhCell_HasDateTime(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(6);
        var expected = new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc);
        var actual = _epochDate.AddDays(double.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture));

        Assert.Equal(expected, actual, precision: TimeSpan.FromMicroseconds(1));
    }

    [Theory]
    [ClassData(typeof(SampleData))]
    public async Task Spreadsheet_Contains_FirstDataRow_EighthCell_HasTime(WorksheetData data)
    {
        byte[] bytes = await data.ToByteArray();
        using var ms = new MemoryStream(bytes);
        using var spreadsheet = SpreadsheetDocument.Open(ms, false);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var row = sheetData.Elements<Row>().ElementAt(1);
        var cell = row.Elements<Cell>().ElementAt(7);

        var expected = new TimeOnly(12, 15, 22);
        var actual = TimeOnly.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture);

        Assert.Equal(expected, actual);
    }

    internal class SampleData : TheoryData<WorksheetData>
    {
        public SampleData()
        {
            Add(CreateSampleData());
        }

        private static WorksheetData CreateSampleData()
        {
            List<Models.Column> columns =
            [
                new() { Name = nameof(ColumnType.Integer), Type = ColumnType.Integer },
                new() { Name = nameof(ColumnType.Text), Type = ColumnType.Text },
                new() { Name = nameof(ColumnType.Decimal), Type = ColumnType.Decimal },
                new() { Name = nameof(ColumnType.Date), Type = ColumnType.Date },
                new() { Name = nameof(ColumnType.Boolean), Type = ColumnType.Boolean },
                new() { Name = nameof(ColumnType.Percentage), Type = ColumnType.Percentage },
                new() { Name = nameof(ColumnType.DateTime), Type = ColumnType.DateTime },
                new() { Name = nameof(ColumnType.Time), Type = ColumnType.Time },
            ];

            List<object[]> rows =
            [
                [
                    1, "John Doe", 50_000.01m, new DateTime(2020, 1, 15, 0, 0, 0, DateTimeKind.Utc), true,
                    1.2345m, new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc), new TimeOnly(12, 15, 22)
                ],
                [
                    2, "Jane 'Very Long Name' Smith", 60_500.50m, new DateTime(2021, 5, 10, 0, 5, 0, DateTimeKind.Utc), false,
                    0.25m, new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc), new TimeOnly(15, 44, 0, 123)
                ],
            ];

            return new WorksheetData
            {
                Columns = columns,
                Rows = rows
            };
        }
    }
}
