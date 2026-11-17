using System.Text.Json;
using ThesisEngine.Core;
using Xunit;

/// <summary>
/// R7–R10 标题层级格式检查。样本预期见 tests/fixtures/synthetic/headings/manifest.json。
/// </summary>
public sealed class HeadingFormatAnalyzerTests
{
    private static string Fixture(string name) =>
        Path.Combine(RepositoryRoot(), "tests", "fixtures", "synthetic", "headings", name);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static HeadingFormatResult Analyze(string name) => new HeadingFormatAnalyzer().Analyze(Fixture(name));

    [Fact]
    public void CompliantSampleRecognizesFourLevelsWithoutIssues()
    {
        var result = Analyze("headings_compliant.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(4, result.CheckedHeadings);
        Assert.Equal(4, result.RecognizedHeadings.Count);
        Assert.Equal(1, result.ChapterStartParagraph);
        Assert.Equal(5, result.SectionEndParagraph);
        Assert.Collection(result.RecognizedHeadings,
            h => { Assert.Equal("章标题", h.Level); Assert.Equal(0, h.OutlineLevel); },
            h => { Assert.Equal("一级标题", h.Level); Assert.Equal(1, h.OutlineLevel); },
            h => { Assert.Equal("二级标题", h.Level); Assert.Equal(2, h.OutlineLevel); },
            h => { Assert.Equal("三级标题", h.Level); Assert.Equal(3, h.OutlineLevel); });
    }

    [Fact]
    public void AllWrongChapterReportsSixContinuousCodesH1ToH6()
    {
        var result = Analyze("headings_all_wrong.docx");

        Assert.Equal(6, result.Issues.Count);
        Assert.Equal(new[] { "H1", "H2", "H3", "H4", "H5", "H6" },
            result.Issues.Select(i => i.Code).ToArray());
        Assert.All(result.Issues, i => Assert.Equal("R7", i.RuleId));
        Assert.Contains(result.Issues, i => i.Dimension == "Font" && i.Actual == "宋体");
        Assert.Contains(result.Issues, i => i.Dimension == "Size" && i.Actual == "12磅");
        Assert.Contains(result.Issues, i => i.Dimension == "Bold" && i.Actual == "未加粗");
    }

    [Fact]
    public void SingleSizeDeviationReportsOnlyItsOwnCode()
    {
        var result = Analyze("headings_size_wrong.docx");

        Assert.Single(result.Issues);
        Assert.Equal("H8", result.Issues[0].Code);
        Assert.Equal("R8", result.Issues[0].RuleId);
        Assert.Equal("Size", result.Issues[0].Dimension);
    }

    [Fact]
    public void ModeTwoNumberingIsRecognizedAndCompliant()
    {
        var result = Analyze("headings_mode2.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(4, result.CheckedHeadings);
        Assert.All(result.RecognizedHeadings, h => Assert.Equal("模式二", h.NumberingMode));
    }

    [Fact]
    public void MissingChapterBoundaryIsReportedAsNotCheckedNotAsPass()
    {
        var result = Analyze("headings_no_boundary.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(0, result.CheckedHeadings);
        Assert.Equal(0, result.ChapterStartParagraph);
        Assert.Contains(result.NotChecked, x => x.Contains("未识别到章标题", StringComparison.Ordinal));
    }

    [Fact]
    public void ConflictDimensionsAreListedAsUndeterminedNotAsIssues()
    {
        var result = Analyze("headings_compliant.docx");

        Assert.Contains(result.NotChecked, x => x.Contains("对齐", StringComparison.Ordinal) && x.Contains("冲突", StringComparison.Ordinal));
        Assert.Contains(result.NotChecked, x => x.Contains("RFC-3", StringComparison.Ordinal));
        Assert.Contains(result.NotChecked, x => x.Contains("加粗", StringComparison.Ordinal) && x.Contains("二级", StringComparison.Ordinal));
    }

    [Fact]
    public void JsonSerializationKeepsIssuesRecognizedAndNotCheckedBuckets()
    {
        var result = Analyze("headings_all_wrong.docx");
        var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        using var document = JsonDocument.Parse(json);

        Assert.True(document.RootElement.TryGetProperty("Issues", out var issues));
        Assert.Equal(6, issues.GetArrayLength());
        Assert.True(document.RootElement.TryGetProperty("RecognizedHeadings", out var headings));
        Assert.True(headings.GetArrayLength() > 0);
        Assert.True(document.RootElement.TryGetProperty("NotChecked", out _));
    }
}
