using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;
using NUnit.Framework;
using System.Globalization;
using SpreadsheetGen.Models;

namespace SpreadsheetGen.Tests;

public class SharedStringsTests
{
    [Test]
    public void Spreadsheet_ExpectedSharedStringValues()
    {
        (SheetData sheetData, SharedStringTable sharedStringTable) = GetSpreadsheetParts();

        var headerRow = sheetData.Elements<Row>().First();
        var cellValues = headerRow.Elements<Cell>()
            .Select(cell =>
                cell.DataType != null && cell.DataType == CellValues.SharedString
                    ? sharedStringTable.ElementAt(int.Parse(cell.CellValue.Text, CultureInfo.InvariantCulture)).InnerText
                    : cell.CellValue?.Text
            )
            .ToList();

        string[] expected = ["Text", "Integer", "Decimal", "Date", "Boolean", "Percentage", "DateTime", "Time"];

        Assert.That(expected, Is.EqualTo(cellValues));
    }

    private static (SheetData sheetData, SharedStringTable sharedStringTable) GetSpreadsheetParts()
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
            new() { Name = "Time", Type = ColumnType.Time }
        ];

        var data = new WorksheetData()
        {
            Columns = columns,
            Rows = [],
        };

        using var ms = new MemoryStream();
        using var spreadsheet = data.ToSpreadsheetDocument(ms);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        var sharedStringTable = spreadsheet.WorkbookPart.SharedStringTablePart.SharedStringTable;
        return (sheetData, sharedStringTable);
    }
}
