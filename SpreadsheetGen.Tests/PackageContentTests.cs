using System.Collections.ObjectModel;
using System.IO.Compression;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Models;
using Xunit;

namespace SpreadsheetGen.Tests;

public class PackageContentTests
{
    private readonly string[] _expectedFiles =
        [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
        ];

    [Theory]
    [ClassData(typeof(HasContent))]
    public async Task Package_Contains_ExpectedFiles(byte[] excel)
    {
        string filePath = $"{nameof(Package_Contains_ExpectedFiles)}_{Guid.NewGuid()}.xlsx";
        await File.WriteAllBytesAsync(filePath, excel, TestContext.Current.CancellationToken);

        using (ZipArchive archive = ZipFile.OpenRead(filePath))
        {
            ReadOnlyCollection<ZipArchiveEntry> entries = archive.Entries;
            HashSet<string> files = entries.Select(e => e.FullName).ToHashSet();

            Assert.True(_expectedFiles.All(files.Contains));
        }

        File.Delete(filePath);
    }

    [Theory]
    [ClassData(typeof(HasContent))]
    public async Task Package_Contains_SharedStringData(byte[] excel)
    {
        string filePath = $"{nameof(Package_Contains_SharedStringData)}_{Guid.NewGuid()}.xlsx";
        await File.WriteAllBytesAsync(filePath, excel, TestContext.Current.CancellationToken);

        using (ZipArchive archive = ZipFile.OpenRead(filePath))
        {
            // Strings information stored in xl/sharedStrings
            ReadOnlyCollection<ZipArchiveEntry> entries = archive.Entries;
            ZipArchiveEntry sharedStringsEntry = entries.FirstOrDefault(e => e.FullName == "xl/sharedStrings.xml");
            if (sharedStringsEntry != null)
            {
                using StreamReader sr = new(sharedStringsEntry.Open());
                string sharedXml = await sr.ReadToEndAsync(TestContext.Current.CancellationToken);

                Assert.Contains("DateTime", sharedXml, StringComparison.InvariantCulture);
            }
        }

        File.Delete(filePath);
    }

    [Theory]
    [ClassData(typeof(HasContent))]
    public async Task Package_Contains_SheetData(byte[] excel)
    {
        string filePath = $"{nameof(Package_Contains_SheetData)}_{Guid.NewGuid()}.xlsx";
        await File.WriteAllBytesAsync(filePath, excel, TestContext.Current.CancellationToken);

        using (ZipArchive archive = ZipFile.OpenRead(filePath))
        {
            // Look for sheet data
            ReadOnlyCollection<ZipArchiveEntry> entries = archive.Entries;
            ZipArchiveEntry firstSheetEntry = entries.FirstOrDefault(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.InvariantCulture));
            using StreamReader sr = new(firstSheetEntry?.Open());
            string xmlContent = await sr.ReadToEndAsync(TestContext.Current.CancellationToken);

            Assert.Contains("<x:sheetData>", xmlContent, StringComparison.InvariantCulture); //Start tag of the data part
        }

        File.Delete(filePath);
    }

    [Theory]
    [ClassData(typeof(NoContent))]
    public async Task Package_Contains_NoSheetData(byte[] excel)
    {
        string filePath = $"{nameof(Package_Contains_NoSheetData)}_{Guid.NewGuid()}.xlsx";
        await File.WriteAllBytesAsync(filePath, excel, TestContext.Current.CancellationToken);

        using (ZipArchive archive = ZipFile.OpenRead(filePath))
        {
            ReadOnlyCollection<ZipArchiveEntry> entries = archive.Entries;
            HashSet<string> files = entries.Select(e => e.FullName).ToHashSet();

            Assert.True(_expectedFiles.All(files.Contains));

            // Look for sheet data
            ZipArchiveEntry firstSheetEntry = entries.FirstOrDefault(e => e.FullName.StartsWith("xl/worksheets/sheet", StringComparison.InvariantCulture));
            using StreamReader sr = new(firstSheetEntry?.Open());
            string xmlContent = await sr.ReadToEndAsync(TestContext.Current.CancellationToken);

            Assert.DoesNotContain("<x:sheetData>", xmlContent, StringComparison.InvariantCulture); //Start tag of the data part
        }

        File.Delete(filePath);
    }

    internal class HasContent : TheoryData<byte[]>
    {
        public HasContent()
        {
            Add(CreateSampleData().Result);
        }

        private static async Task<byte[]> CreateSampleData()
        {
            List<Column> columns =
            [
                new() { Name = nameof(ColumnType.Integer), Type = ColumnType.Integer },
                new() { Name = nameof(ColumnType.Text), Type = ColumnType.Text },
                new() { Name = nameof(ColumnType.Decimal), Type = ColumnType.Decimal },
                new() { Name = nameof(ColumnType.Date), Type = ColumnType.Date },
                new() { Name = nameof(ColumnType.Boolean), Type = ColumnType.Boolean },
                new() { Name = nameof(ColumnType.Percentage), Type = ColumnType.Percentage },
                new() { Name = nameof(ColumnType.DateTime), Type = ColumnType.DateTime },
                new() { Name = nameof(ColumnType.Time), Type = ColumnType.Time },
            ];

            List<object[]> rows =
            [
                [
                    1, "John Doe", 50_000.01m, new DateTime(2020, 1, 15, 0, 0, 0, DateTimeKind.Utc), true,
                    1.2345m, new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc), new TimeOnly(12, 15, 22)
                ],
                [
                    2, "Jane 'Very Long Name' Smith", 60_500.50m, new DateTime(2021, 5, 10, 0, 5, 0, DateTimeKind.Utc), false,
                    0.25m, new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc), new TimeOnly(15, 44, 0, 123)
                ],
            ];

            var data = new WorksheetData()
            {
                Columns = columns,
                Rows = rows
            };

            return await data.ToByteArray();
        }
    }

    internal class NoContent : TheoryData<byte[]>
    {
        public NoContent()
        {
            Add(CreateSampleNoData().Result);
        }

        private static async Task<byte[]> CreateSampleNoData()
        {
            var data = new WorksheetData();
            return await data.ToByteArray();
        }
    }
}
