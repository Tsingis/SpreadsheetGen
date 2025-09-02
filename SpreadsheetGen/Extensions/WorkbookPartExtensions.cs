using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;

namespace SpreadsheetGen.Extensions;

internal static class WorkbookPartExtensions
{
    internal static void AddStyles(this WorkbookPart workbookPart, IReadOnlyCollection<Models.Column> columns)
    {
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStylesheet(columns);
    }

    internal static void SetupSheet(this WorkbookPart workbookPart, WorksheetPart worksheetPart)
    {
        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1,
            Name = "Sheet1"
        });
    }

    internal static Stylesheet CreateStylesheet(IReadOnlyCollection<Models.Column> columns)
    {
        var stylesheet = new Stylesheet();
        stylesheet.AddFonts();
        stylesheet.AddFills();
        stylesheet.AddBorders();
        stylesheet.AddCellFormatsAndStyles(columns);
        return stylesheet;
    }
}
