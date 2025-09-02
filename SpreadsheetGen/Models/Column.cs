using SpreadsheetGen.Enums;

namespace SpreadsheetGen.Models;

public class Column
{
    public string Name { get; set; }
    public ColumnType Type { get; set; } = ColumnType.Text;
    public TotalType? TotalType { get; set; }
}
