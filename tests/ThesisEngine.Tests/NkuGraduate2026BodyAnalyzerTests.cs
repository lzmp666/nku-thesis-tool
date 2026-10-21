using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class NkuGraduate2026BodyAnalyzerTests
{
    private static readonly NkuGraduate2026BodyProfile Profile = NkuGraduate2026BodyProfile.Draft;

    [Fact]
    public void CorrectMixedScriptBodyTextProducesNoDiagnostics()
    {
        var result = Analyze(Run("中文 Latin", "宋体", "Times New Roman", "Times New Roman", 12));

        Assert.Empty(result.Diagnostics);
        Assert.Equal("nku-graduate-2026-body-v1", result.RuleVersion);
        Assert.Contains("draft", result.Coverage);
        Assert.Contains("1 个已确认正文文本片段", result.Coverage);
    }

    [Fact]
    public void IndependentlyChecksHanLatinSlotsAndSize()
    {
        var result = Analyze(Run("中文 Latin", "黑体", "Arial", "Calibri", 10.5));

        Assert.Collection(result.Diagnostics,
            d => Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdFontSize, d.RuleId),
            d => Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdChineseFont, d.RuleId),
            d => Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdLatinFont, d.RuleId));
        Assert.All(result.Diagnostics, d =>
        {
            Assert.Equal("manual", d.FixLevel);
            Assert.Contains("M26-036", d.Message);
            Assert.Contains("规则仍为草案", d.Message);
        });
    }

    [Fact]
    public void OnlyChecksFontSlotsForScriptsPresentInTheText()
    {
        var han = Analyze(Run("只有中文", "黑体", null, null, 12));
        var latin = Analyze(Run("Latin only", null, "Arial", "Arial", 12));

        Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdChineseFont, Assert.Single(han.Diagnostics).RuleId);
        Assert.Empty(han.UnsupportedScopes);
        Assert.Equal(NkuGraduate2026BodyAnalyzer.RuleIdLatinFont, Assert.Single(latin.Diagnostics).RuleId);
        Assert.Empty(latin.UnsupportedScopes);
    }

    [Fact]
    public void UnknownScopeAndEffectiveValuesAreUnsupportedNotViolations()
    {
        var paragraph = new ParagraphSnapshot(0,
            new[] { new TextRunSnapshot(0, "黑体", "Arial", 10, "正文", "Arial", "Arial") },
            ParagraphScopeClassification.Unknown, "边界不明");
        var result = new NkuGraduate2026BodyAnalyzer().Analyze(
            new DocumentSnapshot("word/document.xml", new[] { paragraph }, Array.Empty<string>()), Profile);

        Assert.Empty(result.Diagnostics);
        Assert.Contains(result.UnsupportedScopes, x => x.Contains("paragraph[0]") && x.Contains("边界不明"));
    }

    [Fact]
    public void ExcludedParagraphsDoNotAffectCoverageOrDiagnostics()
    {
        var paragraph = new ParagraphSnapshot(0,
            new[] { new TextRunSnapshot(0, "黑体", "Arial", 10, "参考文献", "Arial", "Arial") },
            ParagraphScopeClassification.Excluded, "标题或非正文部件");
        var result = new NkuGraduate2026BodyAnalyzer().Analyze(
            new DocumentSnapshot("word/document.xml", new[] { paragraph }, Array.Empty<string>()), Profile);

        Assert.Empty(result.Diagnostics);
        Assert.Empty(result.UnsupportedScopes);
        Assert.Contains("已检查 0 个", result.Coverage);
    }

    [Fact]
    public void UnresolvedStyleChainIsUnsupportedWithoutInspectingRunValues()
    {
        var result = Analyze(Run("中文 Latin", "黑体", "Arial", "Arial", 10, unresolvedStyle: true));

        Assert.Empty(result.Diagnostics);
        Assert.Contains(result.UnsupportedScopes, x => x.Contains("有效样式链无法解析"));
    }

    [Fact]
    public void RejectsApprovalUntilProfileAcceptanceTestsAreComplete()
    {
        var profile = Profile with { Status = "approved" };
        var snapshot = new DocumentSnapshot("word/document.xml", Array.Empty<ParagraphSnapshot>(), Array.Empty<string>());

        Assert.Throws<ArgumentException>(() => new NkuGraduate2026BodyAnalyzer().Analyze(snapshot, profile));
    }

    private static AnalysisResult Analyze(TextRunSnapshot run)
    {
        var paragraph = new ParagraphSnapshot(0, new[] { run }, ParagraphScopeClassification.ConfirmedBodyProse);
        var snapshot = new DocumentSnapshot("word/document.xml", new[] { paragraph }, Array.Empty<string>());
        return new NkuGraduate2026BodyAnalyzer().Analyze(snapshot, Profile);
    }

    private static TextRunSnapshot Run(string text, string? eastAsia, string? ascii, string? hAnsi, double? size,
        bool unresolvedStyle = false) =>
        new(0, eastAsia, ascii, size, text, ascii, hAnsi, unresolvedStyle);
}

