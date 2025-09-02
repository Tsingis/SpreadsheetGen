using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;

namespace SpreadsheetGen.Extensions;

internal static class StylesheetExtensions
{
    internal static void AddFonts(this Stylesheet stylesheet)
    {
        var fontElements = new List<OpenXmlElement>();
        var font = new Font();
        fontElements.Add(font);

        var fonts = new Fonts() { Count = 1 };
        fonts.Append(fontElements);

        stylesheet.AppendChild(fonts);
    }

    internal static void AddFills(this Stylesheet stylesheet)
    {
        var nonePatternFill = new PatternFill { PatternType = PatternValues.None };
        var grayPatternFill = new PatternFill { PatternType = PatternValues.Gray125 };

        var nonePatternElements = new List<OpenXmlElement> { nonePatternFill };
        var grayPatternElements = new List<OpenXmlElement> { grayPatternFill };

        var noneFill = new Fill(nonePatternElements);
        var grayFill = new Fill(grayPatternElements);

        var fills = new Fills { Count = 2 };
        fills.AppendChild(noneFill);
        fills.AppendChild(grayFill);

        stylesheet.AppendChild(fills);
    }

    internal static void AddBorders(this Stylesheet stylesheet)
    {
        var border = new Border();
        var borders = new Borders(new List<OpenXmlElement> { border }) { Count = 1 };
        stylesheet.AppendChild(borders);
    }

    internal static void AddCellFormatsAndStyles(this Stylesheet stylesheet, IReadOnlyCollection<Models.Column> columns)
    {
        var baseCellFormat = new CellFormat();
        var baseCellFormatElements = new List<OpenXmlElement> { baseCellFormat };

        var cellStyleFormats = new CellStyleFormats { Count = 1 };
        cellStyleFormats.Append(baseCellFormatElements);

        stylesheet.AppendChild(cellStyleFormats);

        var cellFormats = new CellFormats();
        cellFormats.AppendChild(new CellFormat());

        for (int i = 0; i < columns.Count; i++)
        {
            uint numberFormatId = GetNumberFormatId(columns.ElementAt(i).Type);

            var cellFormat = new CellFormat
            {
                NumberFormatId = numberFormatId,
                ApplyNumberFormat = numberFormatId != 0
            };

            cellFormats.AppendChild(cellFormat);
        }

        cellFormats.Count = (uint)cellFormats.ChildElements.Count;
        stylesheet.AppendChild(cellFormats);

        var cellStyle = new CellStyle
        {
            Name = "Normal",
            FormatId = 0,
            BuiltinId = 0
        };

        var cellStyleElements = new List<OpenXmlElement> { cellStyle };
        var cellStyles = new CellStyles { Count = 1 };
        cellStyles.Append(cellStyleElements);

        stylesheet.AppendChild(cellStyles);
    }

    /// <summary>
    /// Gets number format by column type
    /// <para>About formats: https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.numberingformat?view=openxml-3.0.1</para>
    /// </summary>
    /// <param name="type">ColumnType</param>
    private static uint GetNumberFormatId(ColumnType type)
    {
        return type switch
        {
            ColumnType.Integer => 3,
            ColumnType.Decimal or ColumnType.Currency => 4,
            ColumnType.Percentage => 10,
            ColumnType.DateTime => 22,
            ColumnType.Date => 14,
            ColumnType.Time => 21,
            _ => 0
        };
    }
}
