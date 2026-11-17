using ThesisEngine.Core;
using Xunit;

/// <summary>
/// R4/R5 中英文摘要格式检查。样本预期见
/// tests/fixtures/synthetic/abstract/manifest.json。
/// </summary>
public sealed class AbstractFormatAnalyzerTests
{
    private static string Fixture(string name) =>
        Path.Combine(RepositoryRoot(), "tests", "fixtures", "synthetic", "abstract", name);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }

    private static AbstractFormatResult Analyze(string name) => new AbstractFormatAnalyzer().Analyze(Fixture(name));

    [Fact]
    public void CompliantSampleHasNoIssuesAndBothSectionsRecognized()
    {
        var result = Analyze("abstract_compliant.docx");

        Assert.Empty(result.Issues);
        Assert.Empty(result.NotChecked);
        Assert.Empty(result.MissingParts);
        Assert.Equal(2, result.Sections.Count);
        Assert.Equal(6, result.CheckedParagraphs);
        Assert.Contains(result.Sections, s => s.Section == "中文摘要" && s.TitleParagraph == 1);
        Assert.Contains(result.Sections, s => s.Section == "英文摘要" && s.TitleParagraph == 4);
    }

    [Fact]
    public void MissingAbstractSectionsAreReportedAsNotCheckedNotAsPass()
    {
        var result = Analyze("abstract_no_sections.docx");

        Assert.Empty(result.Issues);
        Assert.Empty(result.Sections);
        Assert.Equal(0, result.CheckedParagraphs);
        Assert.Equal(2, result.NotChecked.Count);
        Assert.Contains(result.NotChecked, s => s.Contains("中文摘要") && s.Contains("未识别"));
        Assert.Contains(result.NotChecked, s => s.Contains("英文摘要") && s.Contains("未识别"));
    }

    [Theory]
    [InlineData("abstract_cn_title_font_wrong.docx", "A1", 1)]
    [InlineData("abstract_cn_title_size_wrong.docx", "A2", 1)]
    [InlineData("abstract_cn_title_not_bold.docx", "A3", 1)]
    [InlineData("abstract_cn_title_left.docx", "A4", 1)]
    [InlineData("abstract_cn_title_spacing_wrong.docx", "A5", 1)]
    [InlineData("abstract_cn_body_font_wrong.docx", "A6", 2)]
    [InlineData("abstract_cn_body_size_wrong.docx", "A7", 2)]
    [InlineData("abstract_cn_body_spacing_wrong.docx", "A8", 2)]
    [InlineData("abstract_cn_keyword_not_bold.docx", "A9", 3)]
    [InlineData("abstract_en_title_font_wrong.docx", "A10", 4)]
    [InlineData("abstract_en_body_font_wrong.docx", "A15", 5)]
    [InlineData("abstract_en_keyword_not_bold.docx", "A18", 6)]
    public void OneDimensionDeviationReportsExactlyOneExpectedCode(string file, string expectedCode, int expectedParagraph)
    {
        var result = Analyze(file);

        var issue = Assert.Single(result.Issues);
        Assert.Equal(expectedCode, issue.Code);
        Assert.Equal(expectedParagraph, issue.ParagraphIndex);
        Assert.Empty(result.NotChecked);
    }

    [Fact]
    public void EnglishTitleFontExpectsArialNotTimesNewRoman()
    {
        var result = Analyze("abstract_en_title_font_wrong.docx");
        var issue = Assert.Single(result.Issues);
        Assert.Equal("A10", issue.Code);
        Assert.Equal("Arial", issue.Expected);
        Assert.Equal("Times New Roman", issue.Actual);
    }

    [Fact]
    public void SourceNoteMarksDraftAndClause()
    {
        var result = Analyze("abstract_compliant.docx");
        Assert.Contains("4.5 表4.4", result.SourceNote);
        Assert.Contains("draft", result.SourceNote);
    }
}
