using System.Text.Json;
using ThesisEngine.Core;
using Xunit;

/// <summary>
/// R6 目录格式检查。样本预期见 tests/fixtures/synthetic/toc/manifest.json。
/// </summary>
public sealed class TocFormatAnalyzerTests
{
    private static string Fixture(string name) =>
        Path.Combine(RepositoryRoot(), "tests", "fixtures", "synthetic", "toc", name);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static TocFormatResult Analyze(string name) => new TocFormatAnalyzer().Analyze(Fixture(name));

    [Fact]
    public void CompliantSampleRecognizesTitleAndThreeEntryLevels()
    {
        var result = Analyze("toc_compliant.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(4, result.CheckedParagraphs);
        Assert.Equal(1, result.TocTitleParagraph);
        Assert.Equal(3, result.Entries.Count);
        Assert.Collection(result.Entries,
            e => { Assert.Equal("章目录条目", e.Level); },
            e => { Assert.Equal("一级目录条目", e.Level); },
            e => { Assert.Equal("二级目录条目", e.Level); });
    }

    [Fact]
    public void TitleWrongReportsSixCodesWithoutLine()
    {
        var result = Analyze("toc_title_wrong.docx");

        Assert.Equal(6, result.Issues.Count);
        Assert.Equal(new[] { "C1", "C2", "C3", "C4", "C6", "C7" },
            result.Issues.Select(i => i.Code).ToArray());
        Assert.All(result.Issues, i => Assert.Equal("R6", i.RuleId));
        Assert.Contains(result.Issues, i => i.Dimension == "Font" && i.Actual == "宋体");
        Assert.Contains(result.Issues, i => i.Dimension == "Size" && i.Actual == "14磅");
    }

    [Fact]
    public void SingleEntrySizeDeviationReportsOnlyItsOwnCode()
    {
        var result = Analyze("toc_entry_size_wrong.docx");

        Assert.Single(result.Issues);
        Assert.Equal("C15", result.Issues[0].Code);
        Assert.Equal("一级目录条目", result.Issues[0].Part);
        Assert.Equal("Size", result.Issues[0].Dimension);
        Assert.Equal("12磅", result.Issues[0].Expected);
    }

    [Fact]
    public void SingleEntryIndentDeviationReportsOnlyItsOwnCode()
    {
        var result = Analyze("toc_entry_indent_wrong.docx");

        Assert.Single(result.Issues);
        Assert.Equal("C25", result.Issues[0].Code);
        Assert.Equal("二级目录条目", result.Issues[0].Part);
        Assert.Equal("Indent", result.Issues[0].Dimension);
    }

    [Fact]
    public void UnexpandedTocFieldIsUndeterminedNotReportedAsPass()
    {
        var result = Analyze("toc_field_unexpanded.docx");

        Assert.Empty(result.Issues);
        Assert.Empty(result.Entries);
        Assert.Equal(1, result.TocTitleParagraph);
        Assert.Contains(result.NotChecked, x => x.Contains("域结果未展开", StringComparison.Ordinal));
    }

    [Fact]
    public void NoTocIsReportedAsNotCheckedNotAsPass()
    {
        var result = Analyze("toc_no_toc.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(0, result.CheckedParagraphs);
        Assert.Equal(0, result.TocTitleParagraph);
        Assert.Contains(result.NotChecked, x => x.Contains("未识别到 TOC", StringComparison.Ordinal));
    }

    [Fact]
    public void AlignmentDependencyIsListedAsUndetermined()
    {
        var result = Analyze("toc_compliant.docx");

        Assert.Contains(result.NotChecked, x =>
            x.Contains("页码右对齐", StringComparison.Ordinal) && x.Contains("T6", StringComparison.Ordinal));
    }

    [Fact]
    public void JsonSerializationKeepsIssuesEntriesAndNotCheckedBuckets()
    {
        var result = Analyze("toc_title_wrong.docx");
        var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        using var document = JsonDocument.Parse(json);

        Assert.True(document.RootElement.TryGetProperty("Issues", out var issues));
        Assert.Equal(6, issues.GetArrayLength());
        Assert.True(document.RootElement.TryGetProperty("Entries", out var entries));
        Assert.True(entries.GetArrayLength() > 0);
        Assert.True(document.RootElement.TryGetProperty("NotChecked", out _));
    }
}
