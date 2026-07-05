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
        var benchmarkArgs = ParseArguments(args ?? []);

        var config = ManualConfig
            .Create(DefaultConfig.Instance)
            .HideColumns(Column.StdDev, Column.Median, Column.Error);

        BenchmarkSwitcher
            .FromAssembly(typeof(Program).Assembly)
            .Run(benchmarkArgs, config);
    }

    private static string[] ParseArguments(string[] args)
    {
        var benchmarkArgs = new List<string>();
        var rowCounts = new List<int>();

        for (var index = 0; index < args.Length; index++)
        {
            var arg = args[index];

            if (arg.Equals("--row-counts", StringComparison.OrdinalIgnoreCase) && index + 1 < args.Length)
            {
                rowCounts.AddRange(ParseRowCounts(args[index + 1]));
                continue;
            }

            benchmarkArgs.Add(arg);
        }

        RowCounts.Clear();

        var selectedRowCounts = rowCounts.Count == 0 ? s_defaultRowCounts.ToList() : rowCounts;

        foreach (var rowCount in selectedRowCounts)
        {
            RowCounts.Add(rowCount);
        }

        return benchmarkArgs.ToArray();
    }

    private static IEnumerable<int> ParseRowCounts(string value)
    {
        return value
            .Split(' ', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(part => int.TryParse(part, out var count) ? count : 0)
            .Where(count => count > 0);
    }
}
