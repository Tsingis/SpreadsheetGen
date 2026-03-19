using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Models;

namespace SpreadsheetGen.Tests;

public class TotalsTests
{
    [Test]
    public void TotalsRow_ContainsExpectedFormulas()
    {
        var data = CreateTotalsSample();

        using var ms = new MemoryStream();
        using SpreadsheetDocument spreadsheet = data.ToSpreadsheetDocument(ms);
        var worksheetPart = spreadsheet.WorkbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        var rows = sheetData.Elements<Row>().ToList();
        var totalsRow = rows[rows.Count - 1];
        var cells = totalsRow.Elements<Cell>().ToList();

        Assert.That(cells[0].DataType?.Value, Is.EqualTo(CellValues.SharedString));

        var sumFormula = cells[2].CellFormula?.InnerText;
        Assert.That(sumFormula, Is.EqualTo("SUM(C2:C3)"));

        var countFormula = cells[3].CellFormula?.InnerText;
        Assert.That(countFormula, Is.EqualTo("COUNTA(D2:D3)"));

        var avgFormula = cells[4].CellFormula?.InnerText;
        Assert.That(avgFormula, Is.EqualTo("AVERAGE(E2:E3)"));

    }

    private static WorksheetData CreateTotalsSample()
    {
        var columns = new List<Models.Column>
        {
            new() { Name = "NoTotal", Type = ColumnType.Text },
            new() { Name = "Int", Type = ColumnType.Integer, TotalType = TotalType.Sum },
            new() { Name = "Text", Type = ColumnType.Text, TotalType = TotalType.Count },
            new() { Name = "Decimal", Type = ColumnType.Decimal, TotalType = TotalType.Average }
        };

        var rows = new List<object[]>
        {
            new object[] { "A", 1, "x", 10.5m },
            new object[] { "B", 2, "", 20.5m }
        };

        return new WorksheetData
        {
            Columns = columns,
            Rows = rows
        };
    }
}
