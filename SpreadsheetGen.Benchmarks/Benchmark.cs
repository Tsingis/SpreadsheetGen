using BenchmarkDotNet.Attributes;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Models;

namespace SpreadsheetGen.Benchmarks;

[MemoryDiagnoser]
[MarkdownExporterAttribute.GitHub]
public class Benchmark
{
    [ParamsSource(nameof(GetRowCounts))]
    public int RowCount;

    public static IEnumerable<int> GetRowCounts()
    {
        return Program.RowCounts;
    }

    private WorksheetData _data;

    [GlobalSetup]
    public void Setup()
    {
        List<Column> columns =
        [
            new() { Name = "Integer", Type = ColumnType.Integer },
            new() { Name = "String", Type = ColumnType.Text },
            new() { Name = "Decimal", Type = ColumnType.Decimal },
            new() { Name = "Date", Type = ColumnType.Date },
            new() { Name = "Boolean", Type = ColumnType.Boolean },
            new() { Name = "Percentage", Type = ColumnType.Percentage },
            new() { Name = "DateTime", Type = ColumnType.DateTime },
            new() { Name = "Time", Type = ColumnType.Time },
        ];

        object[] row = [
            1, "John Doe", 50_000.01m, new DateTime(2020, 1, 15, 0, 0, 0, DateTimeKind.Utc), true,
            1.2345m, new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc), new TimeOnly(12, 15, 22)
        ];

        var rows = Enumerable.Repeat(row, RowCount).ToList();

        _data = new WorksheetData
        {
            Columns = columns,
            Rows = rows
        };
    }

    [Benchmark]
    public async Task<byte[]> CreateByteArray()
    {
        return await _data.ToByteArray();
    }
}
