using System.Text.Json;
using ThesisEngine.Core;
using Xunit;

/// <summary>
/// R11 正文段落格式检查（补齐两端对齐/首行缩进/固定行距/段前后）。
/// 样本预期见 tests/fixtures/synthetic/body-text/manifest.json。
/// </summary>
public sealed class BodyTextFormatAnalyzerTests
{
    private static string Fixture(string name) =>
        Path.Combine(RepositoryRoot(), "tests", "fixtures", "synthetic", "body-text", name);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static BodyTextFormatResult Analyze(string name) => new BodyTextFormatAnalyzer().Analyze(Fixture(name));

    [Fact]
    public void CompliantSampleHasNoIssuesAndRecognizesBodyBoundaries()
    {
        var result = Analyze("body_compliant.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(1, result.ChapterStartParagraph);
        Assert.Equal(5, result.SectionEndParagraph);
        Assert.Equal(3, result.CheckedParagraphs);
        Assert.Equal(3, result.BodyParagraphs.Count);
        Assert.Contains(result.NotChecked, x => x.Contains("#43", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("body_align_left.docx", "T1")]
    [InlineData("body_indent_wrong.docx", "T2")]
    [InlineData("body_line_spacing_wrong.docx", "T3")]
    [InlineData("body_before_wrong.docx", "T4")]
    [InlineData("body_after_wrong.docx", "T4")]
    public void SingleDimensionDeviationReportsOnlyItsOwnCode(string sample, string code)
    {
        var result = Analyze(sample);

        Assert.NotEmpty(result.Issues);
        Assert.All(result.Issues, issue => Assert.Equal(code, issue.Code));
        Assert.Equal(3, result.Issues.Count);
    }

    [Fact]
    public void MathParagraphExemptsLineSpacingButNotOtherDimensions()
    {
        var result = Analyze("body_math_line_exempt.docx");

        Assert.Empty(result.Issues);
        Assert.Contains(result.NotChecked, x => x.Contains("数学表达式", StringComparison.Ordinal));
    }

    [Fact]
    public void MissingChapterBoundaryIsReportedAsNotCheckedNotAsPass()
    {
        var result = Analyze("body_no_boundary.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(0, result.CheckedParagraphs);
        Assert.Equal(0, result.ChapterStartParagraph);
        Assert.Contains(result.NotChecked, x => x.Contains("未识别到章标题", StringComparison.Ordinal));
    }

    [Fact]
    public void JsonSerializationUsesPascalCaseAndKeepsThreeBuckets()
    {
        var result = Analyze("body_align_left.docx");
        var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        using var document = JsonDocument.Parse(json);

        Assert.True(document.RootElement.TryGetProperty("Issues", out var issues));
        Assert.True(issues.GetArrayLength() > 0);
        Assert.True(document.RootElement.TryGetProperty("NotChecked", out _));
        Assert.True(document.RootElement.TryGetProperty("BodyParagraphs", out _));
    }
}
