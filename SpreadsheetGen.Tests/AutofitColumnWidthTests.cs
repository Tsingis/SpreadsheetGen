using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Extensions;
using NUnit.Framework;

namespace SpreadsheetGen.Tests;

public class AutofitColumnWidthTests
{
    [Test]
    public void AutofitColumnWidth_SetsExpectedWidths()
    {
        var worksheet = new Worksheet(new List<OpenXmlElement> { new SheetData() });

        List<Models.Column> columns =
        [
            new() {
                Name = "Column1",
                Type = ColumnType.Text
            },
            new() {
                Name = "Column2",
                Type = ColumnType.Text
            }
        ];

        List<object[]> rows =
        [
            ["A", "This is a very long value"],
            ["B", "Short"],
        ];

        worksheet.AutoFitColumnWidths(columns, rows);

        var cols = worksheet.Elements<Columns>().FirstOrDefault();
        var colList = cols.Elements<Column>().ToList();

        Assert.That(colList[1].Width, Is.GreaterThan(colList[0].Width));
    }

    [Test]
    public void AutofitColumnWidth_NotCalled()
    {
        var worksheet = new Worksheet(new List<OpenXmlElement> { new SheetData() });
        var cols = worksheet.Elements<Columns>().FirstOrDefault();

        Assert.That(cols, Is.Null);
    }
}
