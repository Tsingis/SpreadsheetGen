using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Running;

namespace SpreadsheetGen.Benchmarks;

public static class Program
{
    internal static ICollection<int> RowCounts { get; } = [];
    private static readonly int[] s_defaultRowCounts = [100, 1_000, 10_000];

    public static void Main(string[] args)
    {
        if (args?.Length > 0)
        {
            var parts = args[0].Split(',', StringSplitOptions.RemoveEmptyEntries);
            foreach (var part in parts)
            {
                if (int.TryParse(part, out var count))
                {
                    RowCounts.Add(count);
                }
            }
        }

        if (RowCounts.Count == 0)
        {
            s_defaultRowCounts.ToList().ForEach(RowCounts.Add);
        }

        var config = ManualConfig
            .Create(DefaultConfig.Instance)
            .HideColumns(Column.StdDev, Column.Median, Column.Error);

        BenchmarkRunner.Run<Benchmark>(config);
    }
}
