using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetGen.Extensions;

internal static class WorksheetPartExtensions
{
    internal static void AddTable(this WorksheetPart worksheetPart, int rowCount, List<string> headers, bool includeTotalsColumn = false)
    {
        var tableStartIndex = includeTotalsColumn ? 2 : 1; // 1-based column index where table starts
        var tableColumnCount = includeTotalsColumn ? headers.Count - 1 : headers.Count;

        var startCell = $"{GetColumnName(tableStartIndex)}1";
        var endCell = $"{GetColumnName(tableStartIndex + tableColumnCount - 1)}{rowCount}";
        var tableRange = $"{startCell}:{endCell}";

        var tableDefPart = worksheetPart.AddNewPart<TableDefinitionPart>();
        tableDefPart.Table = new Table
        {
            Id = 1,
            Name = "Table1",
            DisplayName = "Table1",
            Reference = tableRange,
            TotalsRowShown = false,
            AutoFilter = new AutoFilter { Reference = tableRange },
            TableColumns = new TableColumns { Count = (uint)tableColumnCount }
        };

        var headerStart = includeTotalsColumn ? 1 : 0;
        for (uint i = 0; i < tableColumnCount; i++)
        {
            tableDefPart.Table.TableColumns.AppendChild(new TableColumn
            {
                Id = i + 1,
                Name = headers[headerStart + (int)i]
            });
        }

        tableDefPart.Table.TableStyleInfo = new TableStyleInfo
        {
            Name = "TableStyleMedium2",
            ShowFirstColumn = false,
            ShowLastColumn = false,
            ShowRowStripes = true,
            ShowColumnStripes = false
        };

        var tableParts = worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
        if (tableParts == null)
        {
            tableParts = new TableParts { Count = 1 };
            worksheetPart.Worksheet.AppendChild(tableParts);
        }

        var tablePart = new TablePart()
        {
            Id = worksheetPart.GetIdOfPart(tableDefPart)
        };

        tableParts.AppendChild(tablePart);
    }


    /// <summary>
    /// Gets column name by index
    /// </summary>
    /// <param name="columnIndex">1-based column index</param>
    /// <returns>
    /// The column name as a string (e.g., 1 → "A", 26 → "Z", 27 → "AA", 28 → "AB").
    /// </returns>
    private static string GetColumnName(int columnIndex)
    {
        var dividend = columnIndex;
        var sb = new StringBuilder();
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            sb.Insert(0, Convert.ToChar(65 + modulo));
            dividend = (dividend - modulo) / 26;
        }

        return sb.ToString();
    }

}

