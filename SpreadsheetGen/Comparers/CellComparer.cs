using DocumentFormat.OpenXml.Spreadsheet;

namespace SpreadsheetGen.Comparers;

public class CellComparer : IEqualityComparer<Cell>
{
    public bool Equals(Cell x, Cell y)
    {
        if (x == null || y == null)
        {
            return false;
        }

        return Equals(x.DataType, y.DataType) && x.CellValue?.Text == y.CellValue?.Text;
    }

    public int GetHashCode(Cell obj)
    {
        return HashCode.Combine(obj?.DataType, obj?.CellValue?.Text);
    }
}
