using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetGen.Extensions;

internal static class WorksheetExtensions
{
    internal static void AutoFitColumnWidths(this Worksheet worksheet, IReadOnlyCollection<Models.Column> columns, IReadOnlyCollection<object[]> rows)
    {
        var cols = new Columns();
        for (var i = 0; i < columns.Count; i++)
        {
            var maxLen = columns.ElementAt(i).Name.Length;
            foreach (object[] row in rows)
            {
                var value = row[i]?.ToString() ?? string.Empty;
                if (value.Length > maxLen)
                {
                    maxLen = value.Length;
                }
            }
            var width = Math.Max(8, maxLen * 0.9 + 2);
            cols.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Column { Min = (uint)(i + 1), Max = (uint)(i + 1), Width = width, CustomWidth = true });
        }
        worksheet.InsertAt(cols, 0);
    }
}
