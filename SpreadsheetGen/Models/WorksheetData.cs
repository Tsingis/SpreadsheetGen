namespace SpreadsheetGen.Models;

public class WorksheetData
{
    public IReadOnlyCollection<Column> Columns { get; set; }
    public IReadOnlyCollection<object[]> Rows { get; set; }
}
