using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Helpers;
using SpreadsheetGen.Models;

namespace SpreadsheetGen.Extensions;

internal static class SheetDataPartExtensions
{
    internal static void CreateHeaderRow(this SheetData sheetData, IReadOnlyCollection<Models.Column> columns, Dictionary<string, int> sharedStringDict, List<string> sharedStringList)
    {
        var row = new Row();
        foreach (var col in columns)
        {
            row.AppendChild(CellHelper.CreateSharedStringCell(col.Name, sharedStringDict, sharedStringList));
        }

        sheetData.AppendChild(row);
    }

    internal static void CreateDataRows(this SheetData sheetData, WorksheetData data, Dictionary<string, int> sharedStringDict, List<string> sharedStringList)
    {
        foreach (var rowData in data.Rows)
        {
            var row = new Row();
            for (var col = 0; col < data.Columns.Count; col++)
            {
                var value = rowData[col];
                var columnType = data.Columns.ElementAt(col).Type;
                var styleIndex = (uint)(col + 1); // index 0 is default
                row.AppendChild(columnType == ColumnType.Text ?
                    CellHelper.CreateSharedStringCell(value, sharedStringDict, sharedStringList) :
                    CellHelper.CreateCell(value, columnType, styleIndex));
            }

            sheetData.AppendChild(row);
        }
    }
}
