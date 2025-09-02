using System.Collections.ObjectModel;
using System.IO.Compression;
using SpreadsheetGen.Enums;
using SpreadsheetGen.Models;
using Xunit;

namespace SpreadsheetGen.Tests;


[CollectionDefinition(nameof(PackageContentTestsCollection))]
public class PackageContentTestsCollection : ICollectionFixture<TestPackageFixture> { }

[Collection(nameof(PackageContentTestsCollection))]
public class PackageContentTests
{
    private readonly TestPackageFixture _fixture;

    public PackageContentTests(TestPackageFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Package_WithData_Contains_ExpectedFiles()
    {
        using var archive = ZipFile.OpenRead(_fixture.ContentFilePath);
        var files = archive.Entries.Select(x => x.FullName).ToHashSet();

        string[] expectedFiles =
        [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
            "xl/tables/table1.xml",
            "xl/worksheets/sheet1.xml",
            "xl/_rels/workbook.xml.rels",
        ];

        Assert.True(expectedFiles.All(files.Contains));
    }

    [Fact]
    public async Task Package_WithData_Contains_SharedStringData()
    {
        using var archive = ZipFile.OpenRead(_fixture.ContentFilePath);
        var sharedStringsEntry = archive.Entries
            .FirstOrDefault(x => x.FullName.Equals("xl/sharedStrings.xml", StringComparison.Ordinal));
        if (sharedStringsEntry != null)
        {
            using var sr = new StreamReader(sharedStringsEntry.Open());
            var sharedXml = await sr.ReadToEndAsync(TestContext.Current.CancellationToken);

            Assert.Contains("<x:si>", sharedXml, StringComparison.InvariantCulture);
        }
    }

    [Fact]
    public async Task Package_WithData_Contains_SheetData()
    {
        using var archive = ZipFile.OpenRead(_fixture.ContentFilePath);
        var firstSheetEntry = archive.Entries
            .FirstOrDefault(x => x.FullName.StartsWith("xl/worksheets/sheet", StringComparison.InvariantCulture));
        using var sr = new StreamReader(firstSheetEntry?.Open());
        var xmlContent = await sr.ReadToEndAsync(TestContext.Current.CancellationToken);

        Assert.Contains("<x:sheetData>", xmlContent, StringComparison.InvariantCulture);
    }

    [Fact]
    public void Package_NoData_Contains_ExpectedFiles()
    {
        using var archive = ZipFile.OpenRead(_fixture.NoContentFilePath);
        var files = archive.Entries.Select(x => x.FullName).ToHashSet();

        string[] expectedFiles =
        [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
        ];

        Assert.True(expectedFiles.All(files.Contains));
    }

    [Fact]
    public async Task Package_NoData_Contains_NoSheetData()
    {
        using var archive = ZipFile.OpenRead(_fixture.NoContentFilePath);
        var firstSheetEntry = archive.Entries
            .FirstOrDefault(x => x.FullName.StartsWith("xl/worksheets/sheet", StringComparison.InvariantCulture));
        using var sr = new StreamReader(firstSheetEntry?.Open());
        var xmlContent = await sr.ReadToEndAsync(TestContext.Current.CancellationToken);

        Assert.DoesNotContain("<x:sheetData>", xmlContent, StringComparison.InvariantCulture);
    }
}

public class TestPackageFixture : IDisposable
{
    public string ContentFilePath { get; }
    public string NoContentFilePath { get; }

    private bool _disposed;

    public TestPackageFixture()
    {
        ContentFilePath = Path.GetFullPath($"test_has_content_{Guid.NewGuid()}.xlsx");
        NoContentFilePath = Path.GetFullPath($"test_no_content_{Guid.NewGuid()}.xlsx");

        List<Column> columns =
        [
            new() { Name = "Text", Type = ColumnType.Text },
            new() { Name = "Integer", Type = ColumnType.Integer },
            new() { Name = "Decimal", Type = ColumnType.Decimal },
            new() { Name = "Date", Type = ColumnType.Date },
            new() { Name = "Boolean", Type = ColumnType.Boolean },
            new() { Name = "Percentage", Type = ColumnType.Percentage },
            new() { Name = "DateTime", Type = ColumnType.DateTime },
            new() { Name = "Time", Type = ColumnType.Time },
        ];

        List<object[]> rows =
        [
            [
                "John Doe",
                1,
                50_000.01m,
                new DateTime(2020, 1, 15, 0, 0, 0, DateTimeKind.Utc),
                true,
                1.2345m,
                new DateTime(2020, 12, 24, 23, 59, 59,
                DateTimeKind.Utc),
                new TimeOnly(12, 15, 22)
            ],
            [
                "Jane 'Very Long Name' Smith",
                2,
                60_500.50m,
                new DateTime(2021, 5, 10, 0, 5, 0, DateTimeKind.Utc),
                false,
                0.25m,
                new DateTime(2020, 12, 24, 23, 59, 59, DateTimeKind.Utc),
                new TimeOnly(15, 44, 0, 123)
            ],
        ];

        var contentData = new WorksheetData()
        {
            Columns = columns,
            Rows = rows
        };

        var noContentData = new WorksheetData();

        File.WriteAllBytes(ContentFilePath, contentData.ToByteArray().GetAwaiter().GetResult());
        File.WriteAllBytes(NoContentFilePath, noContentData.ToByteArray().GetAwaiter().GetResult());
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!_disposed)
        {
            if (disposing)
            {
                File.Delete(ContentFilePath);
                File.Delete(NoContentFilePath);
            }
            _disposed = true;
        }
    }

    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }
}
