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

    private static Stylesheet CreateStylesheet(IReadOnlyCollection<Models.Column> columns)
    {
        var stylesheet = new Stylesheet();
        stylesheet.AppendChild(new Fonts(new List<OpenXmlElement>
        {
            new Font()
        })
        { Count = 1 });

        var fills = new Fills() { Count = 2 };
        fills.AppendChild(new Fill(new List<OpenXmlElement> { new PatternFill { PatternType = PatternValues.None } }));
        fills.AppendChild(new Fill(new List<OpenXmlElement> { new PatternFill { PatternType = PatternValues.Gray125 } }));
        stylesheet.AppendChild(fills);

        stylesheet.AppendChild(new Borders(new List<OpenXmlElement>
        {
            new Border()
        })
        { Count = 1 });

        stylesheet.AppendChild(new CellStyleFormats(new List<OpenXmlElement>
        {
            new CellFormat()
        })
        { Count = 1 });

        var cellFormats = new CellFormats();
        cellFormats.AppendChild(new CellFormat());

        for (var i = 0; i < columns.Count; i++)
        {
            uint numberFormatId = GetNumberFormatId(columns.ElementAt(i).Type);
            cellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = numberFormatId,
                ApplyNumberFormat = numberFormatId != 0
            });
        }

        cellFormats.Count = (uint)cellFormats.ChildElements.Count;
        stylesheet.AppendChild(cellFormats);

        stylesheet.AppendChild(new CellStyles(new List<OpenXmlElement>
        {
            new CellStyle {
                Name = "Normal",
                FormatId = 0,
                BuiltinId = 0
            } }
        )
        { Count = 1 });

        return stylesheet;
    }

    /// <summary>
    /// About formats: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1
    /// </summary>
    private static uint GetNumberFormatId(ColumnType type)
    {
        return type switch
        {
            ColumnType.Integer => 3,
            ColumnType.Decimal => 4,
            ColumnType.Currency => throw new NotImplementedException(),
            ColumnType.Percentage => 10,
            ColumnType.DateTime => 22,
            ColumnType.Date => 14,
            ColumnType.Time => 21,
            _ => 0
        };
    }
}
