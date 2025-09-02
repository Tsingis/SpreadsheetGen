using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;

namespace SpreadsheetGen.Helpers;

internal static class CellHelper
{
    internal static Cell CreateCell(object value, ColumnType columnType, uint? styleIndex = null)
    {
        var cell = columnType switch
        {
            ColumnType.Integer => CreateIntegerCell(value),
            ColumnType.Decimal => CreateDecimalCell(value),
            ColumnType.Percentage => CreatePercentageCell(value),
            ColumnType.Currency => CreateCurrencyCell(value),
            ColumnType.DateTime => CreateDateTimeCell(value),
            ColumnType.Date => CreateDateCell(value),
            ColumnType.Time => CreateTimeCell(value),
            ColumnType.Boolean => CreateBooleanCell(value),
            _ => throw new NotSupportedException($"{nameof(columnType)}: '{columnType}' is not supported")
        };

        if (styleIndex.HasValue)
        {
            cell.StyleIndex = styleIndex.Value;
        }

        return cell;
    }

    internal static Cell CreateSharedStringCell(object value, Dictionary<string, int> sharedStringDict, List<string> sharedStringList)
    {
        var str = value?.ToString() ?? string.Empty;
        if (!sharedStringDict.TryGetValue(str, out int idx))
        {
            idx = sharedStringList.Count;
            sharedStringDict[str] = idx;
            sharedStringList.Add(str);
        }

        var cellValue = idx.ToString(CultureInfo.InvariantCulture);

        var cell = new Cell()
        {
            DataType = CellValues.SharedString,
            CellValue = new CellValue(cellValue)
        };

        return cell;
    }

    private static Cell CreateIntegerCell(object value)
    {
        var cellValue = value == null
            ? string.Empty
            : Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);

        var cell = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(cellValue)
        };

        return cell;
    }

    private static Cell CreateDecimalCell(object value)
    {
        var cellValue = value == null
            ? string.Empty
            : Convert.ToDecimal(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);

        var cell = new Cell()
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(cellValue)
        };

        return cell;
    }

    private static Cell CreatePercentageCell(object value)
    {
        return CreateDecimalCell(value);
    }

    private static Cell CreateCurrencyCell(object value)
    {
        throw new NotImplementedException();
    }

    /// <summary>
    /// Date values in excel are index based on OLE Automation date unless formatted
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    private static Cell CreateDateTimeCell(object value)
    {
        var cellValue = value switch
        {
            DateTime dt => dt.ToOADate().ToString(CultureInfo.InvariantCulture),
            DateOnly dOnly => dOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture),
            null => string.Empty,
            _ => Convert.ToDateTime(value, CultureInfo.InvariantCulture).ToOADate().ToString(CultureInfo.InvariantCulture)
        };

        var cell = new Cell()
        {
            DataType = null,
            CellValue = new CellValue(cellValue)
        };

        return cell;
    }

    /// <summary>
    /// Date values in excel are index based on OLE Automation date unless formatted
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    private static Cell CreateDateCell(object value)
    {
        var cellValue = value switch
        {
            DateOnly dOnly => dOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture),
            DateTime dt => dt.ToOADate().ToString(CultureInfo.InvariantCulture),
            _ => string.Empty
        };

        var cell = new Cell()
        {
            DataType = null,
            CellValue = new CellValue(cellValue)
        };

        return cell;
    }

    //TODO: Format as time
    private static Cell CreateTimeCell(object value)
    {
        var cellValue = value == null
            ? string.Empty
            : ((TimeOnly)value).ToString("HH:mm:ss", CultureInfo.InvariantCulture);

        var cell = new Cell()
        {
            DataType = CellValues.String,
            CellValue = new CellValue(cellValue)
        };

        return cell;
    }

    private static Cell CreateBooleanCell(object value)
    {
        var cellValue = value != null && Convert.ToBoolean(value, CultureInfo.InvariantCulture);

        var cell = new Cell()
        {
            DataType = CellValues.Boolean,
            CellValue = new CellValue(cellValue)
        };

        return cell;
    }
}
