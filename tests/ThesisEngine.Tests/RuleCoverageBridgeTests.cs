using ThesisEngine.Core.Analysis;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// Regression for the rule-coverage gaps that used to be marked "未实现" even though
/// the underlying analyser already checked them:
/// - R1 gutter / header-footer distance (D3/D4 in DocumentAnalyzer, SectionLayout in DocxLayoutProbe)
/// - R16 symbol table / CV (BackMatterFormatAnalyzer)
/// </summary>
public sealed class RuleCoverageBridgeTests
{
    private static string Fixture(string relative)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return Path.Combine(directory.FullName, "tests", "fixtures", "synthetic", relative);
    }

    [Fact]
    public void PageSetupGutterAndHeaderFooterAreActuallyChecked()
    {
        var result = new RuleAnalysisService().Analyze(
            Fixture(Path.Combine("document", "document_toc_auto.docx")));

        var gutter = Assert.Single(result.Coverage, c => c.RuleId == "R1/gutter_cm");
        Assert.True(gutter.Implemented);
        Assert.Equal(1, gutter.Applicable);
        Assert.Null(gutter.Reason);

        var headerFooter = Assert.Single(result.Coverage, c => c.RuleId == "R1/header_footer_cm");
        Assert.True(headerFooter.Implemented);
        Assert.Equal(1, headerFooter.Applicable);

        Assert.DoesNotContain(result.UncheckedItems,
            u => u.Reason.Contains("装订线检测未实现", StringComparison.Ordinal)
              || u.Reason.Contains("UI 对应关系待确认", StringComparison.Ordinal));
    }

    [Fact]
    public void GutterDeviationIsReportedAsFinding()
    {
        // 合成一个装订线非零的样本：在合规模板上把 w:gutter 改成 567 twips（1cm）。
        var source = Fixture(Path.Combine("document", "document_toc_auto.docx"));
        var target = Path.Combine(Path.GetTempPath(), "rule-bridge-gutter-" + Guid.NewGuid().ToString("N") + ".docx");
        try
        {
            using (var zip = System.IO.Compression.ZipFile.OpenRead(source))
            using (var copy = System.IO.Compression.ZipFile.Open(target, System.IO.Compression.ZipArchiveMode.Create))
            {
                foreach (var entry in zip.Entries)
                {
                    var outEntry = copy.CreateEntry(entry.FullName);
                    using var input = entry.Open();
                    using var output = outEntry.Open();
                    if (entry.FullName != "word/document.xml")
                    {
                        input.CopyTo(output);
                        continue;
                    }
                    using var reader = new StreamReader(input);
                    var xml = reader.ReadToEnd();
                    xml = System.Text.RegularExpressions.Regex.Replace(xml,
                        @"(<w:pgMar\b[^>]*?\bw:gutter="")[0-9]+("")", "${1}567${2}");
                    if (!xml.Contains("w:gutter=\"567\""))
                        xml = System.Text.RegularExpressions.Regex.Replace(xml,
                            @"<w:pgMar\b([^>]*?)/>", "<w:pgMar gutter=\"567\"$1/>");
                    using var writer = new StreamWriter(output);
                    writer.Write(xml);
                }
            }

            var result = new RuleAnalysisService().Analyze(target);
            Assert.Contains(result.Findings, f => f.RuleId == "R1/gutter_cm");
        }
        finally
        {
            if (File.Exists(target)) File.Delete(target);
        }
    }

    [Theory]
    [InlineData("back_matter_symbol_western_font_wrong.docx", "R16/font.western")]
    public void SymbolTableIssueIsBridgedIntoRuleCoverage(string sample, string expectedRuleId)
    {
        var result = new RuleAnalysisService().Analyze(Fixture(Path.Combine("back-matter", sample)));

        var rule = Assert.Single(result.Coverage, c => c.RuleId == expectedRuleId);
        Assert.True(rule.Implemented);
        Assert.True(rule.Failures >= 1);
        Assert.Contains(result.Findings, f => f.RuleId == expectedRuleId);

        // 合并判定与标题复用必须如实保留为未检查，不能冒充合规。
        foreach (var id in new[] { "R16/spacing.before_pt", "R16/spacing.after_pt", "R16/heading" })
        {
            var row = Assert.Single(result.Coverage, c => c.RuleId == id);
            Assert.False(row.Implemented);
        }
    }

    [Fact]
    public void AbsentSymbolTableLeavesR16Unchecked()
    {
        var result = new RuleAnalysisService().Analyze(
            Fixture(Path.Combine("back-matter", "back_matter_no_sections.docx")));

        var fontRule = Assert.Single(result.Coverage, c => c.RuleId == "R16/font.cn");
        Assert.False(fontRule.Implemented);
        Assert.Contains(result.UncheckedItems, u => u.Reason.Contains("未识别到符号注释表", StringComparison.Ordinal));
    }
}
