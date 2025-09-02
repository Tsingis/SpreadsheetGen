using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
            workbookPart.AddStyles(data.Columns);
            worksheetPart.Worksheet.AutoFitColumnWidths(data.Columns, data.Rows);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            sheetData.CreateHeaderRow(data.Columns, sharedStringDict, sharedStringList);
            sheetData.CreateDataRows(data, sharedStringDict, sharedStringList);

            worksheetPart.AddTable(data.Rows.Count + 1, data.Columns.Select(x => x.Name).ToList());
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
