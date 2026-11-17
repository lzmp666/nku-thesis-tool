using System.Text.Json;
using ThesisEngine.Core;
using Xunit;

/// <summary>
/// R12 图题/表题格式检查。样本预期见 tests/fixtures/synthetic/captions/manifest.json。
/// </summary>
public sealed class CaptionFormatAnalyzerTests
{
    private static string Fixture(string name) =>
        Path.Combine(RepositoryRoot(), "tests", "fixtures", "synthetic", "captions", name);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static CaptionFormatResult Analyze(string name) => new CaptionFormatAnalyzer().Analyze(Fixture(name));

    [Fact]
    public void CompliantSampleRecognizesFigureAndTable()
    {
        var result = Analyze("captions_compliant.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(2, result.CheckedCaptions);
        Assert.Equal(1, result.FigureCount);
        Assert.Equal(1, result.TableCount);
        Assert.Collection(result.Captions,
            c => Assert.Equal("图题", c.CaptionType),
            c => Assert.Equal("表题", c.CaptionType));
    }

    [Fact]
    public void FigureWrongReportsSixContinuousCodesK1ToK6()
    {
        var result = Analyze("captions_figure_wrong.docx");

        Assert.Equal(6, result.Issues.Count);
        Assert.Equal(new[] { "K1", "K2", "K3", "K4", "K5", "K6" },
            result.Issues.Select(i => i.Code).ToArray());
        Assert.All(result.Issues, i => Assert.Equal("R12", i.RuleId));
        Assert.Contains(result.Issues, i => i.Dimension == "Font" && i.Actual == "黑体");
        Assert.Contains(result.Issues, i => i.Dimension == "After" && i.Expected == "12磅");
    }

    [Fact]
    public void TableAfterDeviationReportsOnlyItsOwnCode()
    {
        var result = Analyze("captions_table_after_wrong.docx");

        Assert.Single(result.Issues);
        Assert.Equal("K12", result.Issues[0].Code);
        Assert.Equal("表题", result.Issues[0].CaptionType);
        Assert.Equal("After", result.Issues[0].Dimension);
        Assert.Equal("6磅", result.Issues[0].Expected);
    }

    [Theory]
    [InlineData("captions_unrecognized.docx")]
    [InlineData("captions_none.docx")]
    public void NoRecognizedCaptionIsReportedAsNotCheckedNotAsPass(string sample)
    {
        var result = Analyze(sample);

        Assert.Empty(result.Issues);
        Assert.Equal(0, result.CheckedCaptions);
        Assert.Contains(result.NotChecked, x => x.Contains("未识别到带编号的图题", StringComparison.Ordinal));
    }

    [Fact]
    public void RenderDependentDimensionsAreListedAsUndetermined()
    {
        var result = Analyze("captions_compliant.docx");

        Assert.Contains(result.NotChecked, x => x.Contains("相对图表对象的位置", StringComparison.Ordinal));
        Assert.Contains(result.NotChecked, x => x.Contains("RFC-3", StringComparison.Ordinal));
        Assert.Contains(result.NotChecked, x => x.Contains("R13", StringComparison.Ordinal));
    }

    [Fact]
    public void JsonSerializationKeepsIssuesCaptionsAndNotCheckedBuckets()
    {
        var result = Analyze("captions_figure_wrong.docx");
        var json = JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        using var document = JsonDocument.Parse(json);

        Assert.True(document.RootElement.TryGetProperty("Issues", out var issues));
        Assert.Equal(6, issues.GetArrayLength());
        Assert.True(document.RootElement.TryGetProperty("Captions", out var captions));
        Assert.True(captions.GetArrayLength() > 0);
        Assert.True(document.RootElement.TryGetProperty("NotChecked", out _));
    }
}
