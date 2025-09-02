using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetGen.Enums;

namespace SpreadsheetGen;

internal static class CellGenerator
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

    internal static Cell CreateIntegerCell(object value)
    {
        var cellValue = value == null ? string.Empty : Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
        return new Cell
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(cellValue)
        };
    }

    internal static Cell CreateDecimalCell(object value)
    {
        var cellValue = value == null ? string.Empty : Convert.ToDecimal(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
        return new Cell
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(cellValue)
        };
    }

    internal static Cell CreatePercentageCell(object value)
    {
        return CreateDecimalCell(value);
    }

    internal static Cell CreateCurrencyCell(object value)
    {
        throw new NotImplementedException();
    }

    internal static Cell CreateDateTimeCell(object value)
    {
        var cellValue = string.Empty;
        if (value != null)
        {
            if (value is DateTime dt)
            {
                cellValue = dt.ToOADate().ToString(CultureInfo.InvariantCulture);
            }
            else if (value is DateOnly dOnly)
            {
                cellValue = dOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture);
            }
            else
            {
                cellValue = Convert.ToDateTime(value, CultureInfo.InvariantCulture).ToOADate().ToString(CultureInfo.InvariantCulture);
            }
        }
        return new Cell
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(cellValue)
        };
    }

    internal static Cell CreateDateCell(object value)
    {
        if (value is DateOnly dOnly)
        {
            return new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(dOnly.ToDateTime(TimeOnly.MinValue).ToOADate().ToString(CultureInfo.InvariantCulture))
            };
        }
        if (value is DateTime dt)
        {
            return new Cell
            {
                DataType = CellValues.Number,
                CellValue = new CellValue(dt.ToOADate().ToString(CultureInfo.InvariantCulture))
            };
        }

        return new Cell
        {
            DataType = CellValues.Number,
            CellValue = new CellValue(string.Empty)
        };
    }

    internal static Cell CreateTimeCell(object value)
    {
        var cellValue = value == null ? string.Empty : ((TimeOnly)value).ToString("HH:mm:ss", CultureInfo.InvariantCulture);
        return new Cell
        {
            DataType = CellValues.String,
            CellValue = new CellValue(cellValue)
        };
    }

    internal static Cell CreateBooleanCell(object value)
    {
        return new Cell
        {
            DataType = CellValues.Boolean,
            CellValue = new CellValue((bool)value)
        };
    }

    internal static Cell CreateStringCell(object value, Dictionary<string, int> sharedStringDict, List<string> sharedStringList)
    {
        var str = value?.ToString() ?? string.Empty;
        if (!sharedStringDict.TryGetValue(str, out int idx))
        {
            idx = sharedStringList.Count;
            sharedStringDict[str] = idx;
            sharedStringList.Add(str);
        }
        return new Cell
        {
            DataType = CellValues.SharedString,
            CellValue = new CellValue(idx.ToString(CultureInfo.InvariantCulture))
        };
    }
}
