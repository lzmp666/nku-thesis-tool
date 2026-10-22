using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class DiagnosticGrouperTests
{
    [Fact]
    public void SameRuleAndExpectedValueAcrossRunsBecomesOneIssueAndKeepsEveryOccurrence()
    {
        var analyzer = new FontSizeAnalyzer();
        var snapshot = new DocumentSnapshot("word/document.xml", new[]
        {
            new ParagraphSnapshot(0, new[]
            {
                new TextRunSnapshot(0, "黑体", "Arial", 12, "第一行"),
                new TextRunSnapshot(1, "楷体", "Arial", 12, "第二行")
            }),
            new ParagraphSnapshot(1, new[]
            {
                new TextRunSnapshot(0, "黑体", "Arial", 12, "第三行")
            })
        }, Array.Empty<string>());

        var result = analyzer.Analyze(snapshot, DemoRuleSet.NkuDraftSample);
        var fontIssue = Assert.Single(result.Diagnostics, d => d.RuleId == FontSizeAnalyzer.RuleIdEastAsiaFont);

        Assert.Equal("宋体", fontIssue.Expected);
        Assert.Equal("黑体、楷体", fontIssue.Actual);
        Assert.Equal(3, fontIssue.Occurrences.Count);
        Assert.Equal(new[]
        {
            "word/document.xml#paragraph[0]/run[0]",
            "word/document.xml#paragraph[0]/run[1]",
            "word/document.xml#paragraph[1]/run[0]"
        }, fontIssue.Occurrences.Select(o => o.Location));
        Assert.Equal(new[] { "黑体", "楷体", "黑体" }, fontIssue.Occurrences.Select(o => o.Actual));
    }

    [Fact]
    public void GroupedAutoIssuePlansOnePatchForEveryOccurrence()
    {
        var first = "word/document.xml#paragraph[0]/run[0]";
        var second = "word/document.xml#paragraph[4]/run[2]";
        var grouped = DiagnosticGrouper.Group(new[]
        {
            new Diagnostic(FontSizeAnalyzer.RuleIdEastAsiaFont, "warning", first, "宋体", "黑体", "auto", "font"),
            new Diagnostic(FontSizeAnalyzer.RuleIdEastAsiaFont, "warning", second, "宋体", "楷体", "auto", "font")
        });

        var plan = PatchPlanner.PlanPatches(new AnalysisResult(grouped, "demo", "", Array.Empty<string>()));

        Assert.Equal(new[] { first, second }, plan.Patches.Select(p => p.Location));
        Assert.Equal(new[] { "黑体", "楷体" }, plan.Patches.Select(p => p.ExpectedBefore));
        Assert.Empty(plan.DeferredDiagnostics);
    }

    [Fact]
    public void DifferentExpectedValuesRemainSeparateIssues()
    {
        var grouped = DiagnosticGrouper.Group(new[]
        {
            new Diagnostic("font", "warning", "p0", "宋体", "黑体", "auto", "font"),
            new Diagnostic("font", "warning", "p1", "楷体", "黑体", "auto", "font")
        });

        Assert.Equal(2, grouped.Count);
    }
}
