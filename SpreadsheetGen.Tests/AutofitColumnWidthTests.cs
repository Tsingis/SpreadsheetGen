using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Extensions;
using Xunit;

namespace SpreadsheetGen.Tests;

public class AutofitColumnWidthTests
{
    [Fact]
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

        Assert.True(colList[1].Width > colList[0].Width);
    }

    [Fact]
    public void AutofitColumnWidth_NotCalled()
    {
        var worksheet = new Worksheet(new List<OpenXmlElement> { new SheetData() });
        var cols = worksheet.Elements<Columns>().FirstOrDefault();

        Assert.Null(cols);
    }
}
