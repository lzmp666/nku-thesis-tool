using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class FontSizeAnalyzerTests
{
    private static readonly DemoRuleSet DefaultProfile = DemoRuleSet.NkuDraftSample;

    private static ParagraphSnapshot Body(int paragraphIndex, IReadOnlyList<TextRunSnapshot> runs) =>
        new(paragraphIndex, runs, ParagraphScopeClassification.ConfirmedBodyProse);

    [Fact]
    public void AllCorrectProducesNoDiagnostics()
    {
        var snapshot = new DocumentSnapshot(
            PartName: "word/document.xml",
            Paragraphs: new List<ParagraphSnapshot>
            {
                Body(0, new List<TextRunSnapshot>
                {
                    new(0, "宋体", "Times New Roman", 12.0, "正确字体字号的文本"),
                }),
            },
            UnsupportedScopes: Array.Empty<string>());

        var analyzer = new FontSizeAnalyzer();
        var result = analyzer.Analyze(snapshot, DefaultProfile);

        Assert.Empty(result.Diagnostics);
        Assert.Equal(DefaultProfile.RuleVersion, result.RuleVersion);
        Assert.Contains("1 个正文文本片段", result.Coverage);
    }

    [Fact]
    public void WrongEastAsianFontProducesDiagnostic()
    {
        var snapshot = new DocumentSnapshot(
            PartName: "word/document.xml",
            Paragraphs: new List<ParagraphSnapshot>
            {
                Body(0, new List<TextRunSnapshot>
                {
                    new(0, "黑体", "Times New Roman", 12.0, "黑体文本"),
                }),
            },
            UnsupportedScopes: Array.Empty<string>());

        var analyzer = new FontSizeAnalyzer();
        var result = analyzer.Analyze(snapshot, DefaultProfile);

        var diag = Assert.Single(result.Diagnostics);
        Assert.Equal(FontSizeAnalyzer.RuleIdEastAsiaFont, diag.RuleId);
        Assert.Equal("warning", diag.Severity);
        Assert.Equal("word/document.xml#paragraph[0]/run[0]", diag.Location);
        Assert.Equal("宋体", diag.Expected);
        Assert.Equal("黑体", diag.Actual);
        Assert.Equal("auto", diag.FixLevel);
    }

    [Fact]
    public void WrongAsciiFontProducesDiagnostic()
    {
        var snapshot = new DocumentSnapshot(
            PartName: "word/document.xml",
            Paragraphs: new List<ParagraphSnapshot>
            {
                Body(0, new List<TextRunSnapshot>
                {
                    new(0, "宋体", "Arial", 12.0, "Arial text"),
                }),
            },
            UnsupportedScopes: Array.Empty<string>());

        var analyzer = new FontSizeAnalyzer();
        var result = analyzer.Analyze(snapshot, DefaultProfile);

        var diag = Assert.Single(result.Diagnostics);
        Assert.Equal(FontSizeAnalyzer.RuleIdAsciiFont, diag.RuleId);
        Assert.Equal("Arial", diag.Actual);
        Assert.Equal("auto", diag.FixLevel);
    }

    [Fact]
    public void WrongHAnsiFontIsDetectedEvenWhenAsciiSlotIsCorrect()
    {
        var snapshot = new DocumentSnapshot(
            "word/document.xml",
            new List<ParagraphSnapshot>
            {
                Body(0, new List<TextRunSnapshot>
                {
                    new(0, "宋体", "Times New Roman", 12.0, "Latin text",
                        AsciiSlotFont: "Times New Roman", HAnsiFont: "Arial")
                })
            },
            Array.Empty<string>());

        var result = new FontSizeAnalyzer().Analyze(snapshot, DefaultProfile);

        var diagnostic = Assert.Single(result.Diagnostics);
        Assert.Equal(FontSizeAnalyzer.RuleIdAsciiFont, diagnostic.RuleId);
        Assert.Equal("ASCII=Times New Roman; HAnsi=Arial", diagnostic.Actual);
        Assert.Equal("auto", diagnostic.FixLevel);
    }

    [Fact]
    public void WrongFontSizeProducesDiagnostic()
    {
        var snapshot = new DocumentSnapshot(
            PartName: "word/document.xml",
            Paragraphs: new List<ParagraphSnapshot>
            {
                Body(0, new List<TextRunSnapshot>
                {
                    new(0, "宋体", "Times New Roman", 14.0, "14pt文本"),
                }),
            },
            UnsupportedScopes: Array.Empty<string>());

        var analyzer = new FontSizeAnalyzer();
        var result = analyzer.Analyze(snapshot, DefaultProfile);

        var diag = Assert.Single(result.Diagnostics);
        Assert.Equal(FontSizeAnalyzer.RuleIdFontSize, diag.RuleId);
        Assert.Equal("12pt", diag.Expected);
        Assert.Equal("14pt", diag.Actual);
    }

    [Fact]
    public void UnknownFontProducesUnknownDiagnostic()
    {
        var snapshot = new DocumentSnapshot(
            PartName: "word/document.xml",
            Paragraphs: new List<ParagraphSnapshot>
            {
                Body(0, new List<TextRunSnapshot>
                {
                    new(0, null, null, null, "未知格式文本"),
                }),
            },
            UnsupportedScopes: Array.Empty<string>());

        var analyzer = new FontSizeAnalyzer();
        var result = analyzer.Analyze(snapshot, DefaultProfile);

        Assert.Equal(2, result.Diagnostics.Count);
        Assert.DoesNotContain(result.Diagnostics, d => d.RuleId == FontSizeAnalyzer.RuleIdAsciiFont);
        Assert.All(result.Diagnostics, d => Assert.Equal("(unknown)", d.Actual));
    }

    [Fact]
    public void MultipleRunsProduceMultipleDiagnostics()
    {
        var snapshot = new DocumentSnapshot(
            PartName: "word/document.xml",
            Paragraphs: new List<ParagraphSnapshot>
            {
                Body(0, new List<TextRunSnapshot>
                {
                    new(0, "宋体", "Times New Roman", 12.0, "正确"),
                    new(1, "黑体", "Arial", 14.0, "全错 English"),
                }),
                Body(1, new List<TextRunSnapshot>
                {
                    new(0, "宋体", "Times New Roman", 12.0, "正确"),
                }),
            },
            UnsupportedScopes: Array.Empty<string>());

        var analyzer = new FontSizeAnalyzer();
        var result = analyzer.Analyze(snapshot, DefaultProfile);

        // 第一个段落第二个run三项全错 = 3个diagnostic
        Assert.Equal(3, result.Diagnostics.Count);
        Assert.All(result.Diagnostics, d => Assert.Contains("paragraph[0]/run[1]", d.Location));
    }

    [Fact]
    public void UnsupportedScopesArePassedThrough()
    {
        var unsupported = new[] { "headers/footers", "textboxes", "tables" };
        var snapshot = new DocumentSnapshot(
            PartName: "word/document.xml",
            Paragraphs: new List<ParagraphSnapshot>(),
            UnsupportedScopes: unsupported);

        var analyzer = new FontSizeAnalyzer();
        var result = analyzer.Analyze(snapshot, DefaultProfile);

        Assert.Empty(result.Diagnostics);
        Assert.Equal(unsupported, result.UnsupportedScopes);
    }

    [Fact]
    public void HanOnlyTextDoesNotProduceAFalseWesternFontIssue()
    {
        var snapshot = new DocumentSnapshot("word/document.xml",
            new[] { new ParagraphSnapshot(0, new[]
            {
                new TextRunSnapshot(0, "宋体", "Arial", 12, "纯中文正文"),
            }) }, Array.Empty<string>());

        var result = new FontSizeAnalyzer().Analyze(snapshot, DefaultProfile);

        Assert.Empty(result.Diagnostics);
    }

    [Fact]
    public void KnownCaptionAndTableScopesAreNotCheckedWithBodyFontRules()
    {
        var snapshot = new DocumentSnapshot("word/document.xml", new[]
        {
            new ParagraphSnapshot(0, new[] { new TextRunSnapshot(0, "黑体", "Arial", 10.5, "表1-1 English") },
                ParagraphScopeClassification.Excluded, "表格单元格；正文规则不适用"),
            new ParagraphSnapshot(1, new[] { new TextRunSnapshot(0, "黑体", "Arial", 10.5, "图1-1 系统结构") },
                ParagraphScopeClassification.Excluded, "图表题注应使用题注专属规则"),
        }, Array.Empty<string>());

        Assert.Empty(new FontSizeAnalyzer().Analyze(snapshot, DefaultProfile).Diagnostics);
    }

    [Fact]
    public void UnknownParagraphScopeIsReportedAsUnsupportedInsteadOfUsingBodyRules()
    {
        var snapshot = new DocumentSnapshot("word/document.xml", new[]
        {
            new ParagraphSnapshot(0, new[] { new TextRunSnapshot(0, "黑体", "Arial", 10.5, "边界未确认") }),
        }, Array.Empty<string>());

        var result = new FontSizeAnalyzer().Analyze(snapshot, DefaultProfile);

        Assert.Empty(result.Diagnostics);
        Assert.Contains(result.UnsupportedScopes, value => value.Contains("正文范围未确认", StringComparison.Ordinal));
        Assert.Contains("已检查 0 个正文文本片段", result.Coverage);
    }

    [Fact]
    public void NullArgumentsThrow()
    {
        var analyzer = new FontSizeAnalyzer();
        var snapshot = new DocumentSnapshot("word/document.xml", Array.Empty<ParagraphSnapshot>(), Array.Empty<string>());

        Assert.Throws<ArgumentNullException>(() => analyzer.Analyze(null!, DefaultProfile));
        Assert.Throws<ArgumentNullException>(() => analyzer.Analyze(snapshot, null!));
    }
}
