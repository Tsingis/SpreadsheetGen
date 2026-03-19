using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Extensions;
using SpreadsheetGen.Models;

namespace SpreadsheetGen;

public static class SpreadsheetGenerator
{
    public static async Task<byte[]> ToByteArray(this WorksheetData data)
    {
        ArgumentNullException.ThrowIfNull(data);

        await using var memoryStream = new MemoryStream();
        using (var spreadsheet = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        {
            SetupDocument(spreadsheet, data);
        }

        return memoryStream.ToArray();
    }

    public static SpreadsheetDocument ToSpreadsheetDocument(this WorksheetData data, Stream stream)
    {
        ArgumentNullException.ThrowIfNull(data);

        SpreadsheetDocument spreadsheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        SetupDocument(spreadsheet, data);
        return spreadsheet;
    }

    private static void SetupDocument(SpreadsheetDocument spreadsheet, WorksheetData data)
    {
        var workbookPart = spreadsheet.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new List<OpenXmlElement> { new SheetData() });

        var sharedStringDict = new Dictionary<string, int>();
        var sharedStringList = new List<string>();

        workbookPart.SetupSheet(worksheetPart);

        if (data.Columns?.Count > 0)
        {
            var totalsShown = data.Columns.Any(c => c.TotalType.HasValue);

            if (totalsShown)
            {
                var effectiveColumns = new List<Models.Column> { new() { Name = string.Empty, Type = ColumnType.Text } };
                effectiveColumns.AddRange(data.Columns);

                var effectiveRows = data.Rows.Select(r =>
                {
                    var arr = new object[data.Columns.Count + 1];
                    arr[0] = string.Empty;
                    for (var i = 0; i < data.Columns.Count; i++)
                    {
                        arr[i + 1] = r[i];
                    }
                    return arr;
                }).ToList();

                workbookPart.AddStyles(effectiveColumns);
                worksheetPart.Worksheet.AutoFitColumnWidths(effectiveColumns, effectiveRows);
            }
            else
            {
                workbookPart.AddStyles(data.Columns);
                worksheetPart.Worksheet.AutoFitColumnWidths(data.Columns, data.Rows);
            }

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            sheetData.CreateHeaderRow(data.Columns, sharedStringDict, sharedStringList, totalsShown);
            sheetData.CreateDataRows(data, sharedStringDict, sharedStringList, totalsShown);

            if (totalsShown)
            {
                sheetData.CreateTotalsRow(data, sharedStringDict, sharedStringList);
            }

            var tableRowCount = data.Rows.Count + 1;
            var headers = totalsShown ? new List<string> { string.Empty } : [];

            if (totalsShown)
            {
                headers.AddRange(data.Columns.Select(x => x.Name));
            }
            else
            {
                headers = data.Columns.Select(x => x.Name).ToList();
            }

            // includeTotalsColumn = totalsShown ; totalsRowShown = false so totals row remains outside the table
            worksheetPart.AddTable(tableRowCount, headers, includeTotalsColumn: totalsShown);
        }

        if (sharedStringList.Count > 0)
        {
            var sharedStringTablePart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringTablePart.SharedStringTable = new SharedStringTable();
            foreach (var item in sharedStringList)
            {
                sharedStringTablePart.SharedStringTable.AppendChild(new SharedStringItem { Text = new Text(item) });
            }
            sharedStringTablePart.SharedStringTable.Count = (uint)sharedStringList.Count;
            sharedStringTablePart.SharedStringTable.UniqueCount = (uint)sharedStringList.Count;
            sharedStringTablePart.SharedStringTable.Save();
        }

        worksheetPart.Worksheet.Save();
        workbookPart.Workbook.Save();
    }
}
