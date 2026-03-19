using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Helpers;
using SpreadsheetGen.Models;

namespace SpreadsheetGen.Extensions;

internal static class SheetDataPartExtensions
{
    internal static void CreateHeaderRow(this SheetData sheetData, IReadOnlyCollection<Models.Column> columns, Dictionary<string, int> sharedStringDict, List<string> sharedStringList, bool includeTotalsColumn = false)
    {
        var row = new Row();
        if (includeTotalsColumn)
        {
            var leadingHeader = CellHelper.CreateSharedStringCell(string.Empty, sharedStringDict, sharedStringList);
            leadingHeader.StyleIndex = 1;
            row.AppendChild(leadingHeader);
        }

        foreach (var col in columns)
        {
            row.AppendChild(CellHelper.CreateSharedStringCell(col.Name, sharedStringDict, sharedStringList));
        }

        sheetData.AppendChild(row);
    }

    internal static void CreateDataRows(this SheetData sheetData, WorksheetData data, Dictionary<string, int> sharedStringDict, List<string> sharedStringList, bool includeTotalsColumn = false)
    {
        foreach (var rowData in data.Rows)
        {
            var row = new Row();
            if (includeTotalsColumn)
            {
                var leading = CellHelper.CreateSharedStringCell(string.Empty, sharedStringDict, sharedStringList);
                leading.StyleIndex = 1;
                row.AppendChild(leading);
            }

            for (var col = 0; col < data.Columns.Count; col++)
            {
                var value = rowData[col];
                var columnType = data.Columns.ElementAt(col).Type;
                var styleIndex = includeTotalsColumn ? (uint)(col + 2) : (uint)(col + 1);
                row.AppendChild(columnType == ColumnType.Text ?
                    CellHelper.CreateSharedStringCell(value, sharedStringDict, sharedStringList) :
                    CellHelper.CreateCell(value, columnType, styleIndex));
            }

            sheetData.AppendChild(row);
        }
    }

    internal static void CreateTotalsRow(this SheetData sheetData, WorksheetData data, Dictionary<string, int> sharedStringDict, List<string> sharedStringList)
    {
        if (data?.Columns == null || data.Columns.Count == 0)
        {
            return;
        }

        if (!data.Columns.Any(c => c.TotalType.HasValue))
        {
            return;
        }

        var headerRow = 1;
        var firstDataRow = headerRow + 1;
        var lastDataRow = headerRow + data.Rows.Count;

        var row = new Row();

        var totalLabelCell = CellHelper.CreateSharedStringCell("Total", sharedStringDict, sharedStringList);
        totalLabelCell.StyleIndex = 1;
        row.AppendChild(totalLabelCell);

        for (var col = 0; col < data.Columns.Count; col++)
        {
            var column = data.Columns.ElementAt(col);

            if (!column.TotalType.HasValue)
            {
                row.AppendChild(new Cell());
                continue;
            }

            var colLetter = GetColumnName(col + 2);
            var func = column.TotalType.Value switch
            {
                TotalType.Sum => "SUM",
                TotalType.Count => "COUNTA",
                TotalType.Average => "AVERAGE",
                _ => null
            };

            if (func == null)
            {
                row.AppendChild(new Cell());
                continue;
            }

            var formula = $"{func}({colLetter}{firstDataRow}:{colLetter}{lastDataRow})";
            var cell = new Cell { CellFormula = new CellFormula(formula) };
            row.AppendChild(cell);
        }

        sheetData.AppendChild(row);
    }

    private static string GetColumnName(int columnIndex)
    {
        var dividend = columnIndex;
        var sb = new System.Text.StringBuilder();
        while (dividend > 0)
        {
            var modulo = (dividend - 1) % 26;
            sb.Insert(0, Convert.ToChar(65 + modulo));
            dividend = (dividend - modulo) / 26;
        }

        return sb.ToString();
    }
}
