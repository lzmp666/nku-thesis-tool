using ThesisEngine.Core;
using Xunit;

/// <summary>
/// R14-R17 尾注部件格式检查。样本预期见
/// tests/fixtures/synthetic/back-matter/manifest.json。
/// </summary>
public sealed class BackMatterFormatAnalyzerTests
{
    private static string Fixture(string name) =>
        Path.Combine(RepositoryRoot(), "tests", "fixtures", "synthetic", "back-matter", name);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static BackMatterFormatResult Analyze(string name) => new BackMatterFormatAnalyzer().Analyze(Fixture(name));

    [Fact]
    public void CompliantSampleHasNoIssuesAndAllFourSectionsRecognized()
    {
        var result = Analyze("back_matter_compliant.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(4, result.Sections.Count);
        Assert.Equal(4, result.CheckedParagraphs);
        Assert.Equal(new[] { "R14", "R15", "R16", "R17" }, result.Sections.Select(s => s.RuleId).ToArray());
        Assert.Single(result.NotChecked);
        Assert.Contains("R7", result.NotChecked[0]);
    }

    [Fact]
    public void MissingSectionsAreReportedAsNotCheckedNotAsPass()
    {
        var result = Analyze("back_matter_no_sections.docx");

        Assert.Empty(result.Issues);
        Assert.Empty(result.Sections);
        Assert.Equal(0, result.CheckedParagraphs);
        Assert.Equal(4, result.MissingParts.Count);
        Assert.Equal(5, result.NotChecked.Count);
        Assert.Contains(result.NotChecked, s => s.Contains("R14") && s.Contains("未识别"));
    }

    [Theory]
    [InlineData("back_matter_reference_cn_font_wrong.docx", "B1", 2)]
    [InlineData("back_matter_reference_size_wrong.docx", "B3", 2)]
    [InlineData("back_matter_reference_spacing_wrong.docx", "B4", 2)]
    [InlineData("back_matter_appendix_size_wrong.docx", "B7", 4)]
    [InlineData("back_matter_appendix_alignment_left.docx", "B8", 4)]
    [InlineData("back_matter_appendix_indent_wrong.docx", "B9", 4)]
    [InlineData("back_matter_symbol_western_font_wrong.docx", "B12", 6)]
    [InlineData("back_matter_ack_cn_font_wrong.docx", "B15", 8)]
    [InlineData("back_matter_ack_spacing_wrong.docx", "B17", 8)]
    public void OneDimensionDeviationReportsExactlyOneExpectedCode(string file, string expectedCode, int expectedParagraph)
    {
        var result = Analyze(file);

        var issue = Assert.Single(result.Issues);
        Assert.Equal(expectedCode, issue.Code);
        Assert.Equal(expectedParagraph, issue.ParagraphIndex);
    }

    [Fact]
    public void MathParagraphIsExemptedFromLineSpacingButNotFromOtherChecks()
    {
        var result = Analyze("back_matter_appendix_math_line_exempt.docx");

        Assert.Empty(result.Issues);
        Assert.Contains(result.NotChecked, s => s.Contains("数学表达式") && s.Contains("行距"));
    }

    [Fact]
    public void SubsectionHeadingsAreNotTreatedAsBodyParagraphs()
    {
        var result = Analyze("back_matter_two_appendices.docx");

        Assert.Empty(result.Issues);
        Assert.Equal(4, result.Sections.Count);
        Assert.Equal(5, result.CheckedParagraphs);
        var appendix = Assert.Single(result.Sections, s => s.RuleId == "R15");
        Assert.Equal(2, appendix.BodyParagraphs.Count);
    }

    [Fact]
    public void FangSongIsRequiredForAcknowledgementsNotSimSun()
    {
        var result = Analyze("back_matter_ack_cn_font_wrong.docx");
        var issue = Assert.Single(result.Issues);
        Assert.Equal("仿宋", issue.Expected);
        Assert.Equal("宋体", issue.Actual);
    }

    [Fact]
    public void SourceNoteMarksDraftAndClause()
    {
        var result = Analyze("back_matter_compliant.docx");
        Assert.Contains("4.8 表4.7", result.SourceNote);
        Assert.Contains("draft", result.SourceNote);
    }
}
