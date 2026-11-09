using System.Text.Json;
using ThesisEngine.Core;
using ThesisEngine.Core.Analysis;
using ThesisEngine.Core.Rules;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class RuleAnalysisTests
{
    [Fact]
    public void CompliantSampleHasZeroFindingsButNonEmptyUncheckedList()
    {
        var result = new RuleAnalysisService().Analyze(Fixture("analysis_compliant.docx"));

        Assert.Equal("draft", result.RuleSet.Status);
        Assert.Empty(result.Findings);
        Assert.Equal(0, result.FailedChecks);
        Assert.True(result.PassedChecks > 0);
        Assert.NotEmpty(result.UncheckedItems);
        Assert.Contains(result.Coverage, row => !row.Implemented);
        Assert.Contains(result.Coverage, row => row.RuleId == "R1/paper_size" && row.Applicable == 1);
    }

    [Fact]
    public void NoncompliantSampleReportsEveryDeliberateBodyDeviation()
    {
        var result = new RuleAnalysisService().Analyze(Fixture("analysis_noncompliant.docx"));

        Assert.Equal(8, result.Findings.Count);
        Assert.Equal(8, result.Findings.Sum(finding => finding.Occurrences.Count));
        var ruleIds = result.Findings.Select(finding => finding.RuleId).ToHashSet(StringComparer.Ordinal);
        Assert.Contains("R11/font.cn", ruleIds);
        Assert.Contains("R11/font.western", ruleIds);
        Assert.Contains("R11/size_pt", ruleIds);
        Assert.Contains("R11/alignment", ruleIds);
        Assert.Contains("R11/first_line_indent_chars", ruleIds);
        Assert.Contains("R11/line_spacing", ruleIds);
        Assert.Contains("R11/spacing.before_pt", ruleIds);
        Assert.Contains("R11/spacing.after_pt", ruleIds);
        Assert.NotEmpty(result.UncheckedItems);
    }

    [Fact]
    public void ThreeStatesArePresentInOneRun()
    {
        var result = new RuleAnalysisService().Analyze(Fixture("analysis_noncompliant.docx"));

        Assert.True(result.PassedChecks > 0);
        Assert.True(result.FailedChecks > 0);
        Assert.NotEmpty(result.UncheckedItems);
        Assert.True(result.CheckedParagraphCount > 0);
    }

    [Fact]
    public void FullThesisProducesARealIssueListWithTraceableFields()
    {
        var result = new RuleAnalysisService().Analyze(Fixture(Path.Combine("full-thesis", "synthetic_complete_thesis.docx")));

        Assert.True(result.Findings.Count >= 30);
        Assert.True(result.Findings.Sum(finding => finding.Occurrences.Count) > result.Findings.Count);
        Assert.Equal(116, result.Coverage.Count);
        Assert.NotEmpty(result.UncheckedItems);
        Assert.Contains(result.Findings, finding => finding.RuleId == "R11/font.cn");
        Assert.Contains(result.Findings, finding => finding.RuleId == "R11/font.western");
        Assert.Contains(result.Findings, finding => finding.RuleId == "R11/size_pt");
        Assert.Contains(result.Findings, finding => finding.RuleId == "R1/paper_size");
        Assert.Contains(result.Findings, finding => finding.RuleId == "R7/line_spacing");
        Assert.Contains(result.Findings, finding => finding.RuleId == "R14/line_spacing");
        Assert.All(result.Findings, finding =>
        {
            Assert.False(string.IsNullOrWhiteSpace(finding.RuleId));
            Assert.False(string.IsNullOrWhiteSpace(finding.Clause));
            Assert.False(string.IsNullOrWhiteSpace(finding.Location));
            Assert.StartsWith("rules.", finding.YamlPath, StringComparison.Ordinal);
            Assert.InRange(finding.Confidence, 0.01, 0.75);
            Assert.NotEmpty(finding.Occurrences);
        });
    }

    [Fact]
    public void RepeatedProblemIsGroupedButAllOccurrencesAreRetained()
    {
        var result = new RuleAnalysisService().Analyze(Fixture(Path.Combine("full-thesis", "synthetic_complete_thesis.docx")));

        var lineSpacing = result.Findings.Single(finding => finding.RuleId == "R11/line_spacing");
        Assert.True(lineSpacing.Occurrences.Count > 1);
        Assert.Equal(lineSpacing.Occurrences.Count, lineSpacing.Occurrences
            .Select(occurrence => occurrence.Location)
            .Distinct(StringComparer.Ordinal)
            .Count());
        Assert.All(lineSpacing.Occurrences, occurrence => Assert.StartsWith("paragraph[", occurrence.Location, StringComparison.Ordinal));
    }

    [Fact]
    public void RunningTwiceProducesTheSameFindingOrder()
    {
        var service = new RuleAnalysisService();
        var first = service.Analyze(Fixture(Path.Combine("full-thesis", "synthetic_complete_thesis.docx")));
        var second = service.Analyze(Fixture(Path.Combine("full-thesis", "synthetic_complete_thesis.docx")));

        Assert.Equal(Describe(first), Describe(second));
    }

    [Fact]
    public void JsonOutputContainsCoverageFindingsAndUncheckedItems()
    {
        var result = new RuleAnalysisService().Analyze(Fixture("analysis_noncompliant.docx"));
        var json = RuleAnalysisReportWriter.ToJson(result);

        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.Equal("thesis-analysis-1", root.GetProperty("schemaVersion").GetString());
        Assert.Equal("draft", root.GetProperty("ruleSet").GetProperty("status").GetString());
        Assert.Equal(8, root.GetProperty("findings").GetArrayLength());
        Assert.True(root.GetProperty("coverage").GetArrayLength() >= 100);
        Assert.True(root.GetProperty("unchecked").GetArrayLength() > 0);
        var first = root.GetProperty("findings")[0];
        Assert.False(string.IsNullOrWhiteSpace(first.GetProperty("ruleId").GetString()));
        Assert.StartsWith("D2026 ", first.GetProperty("d2026").GetString(), StringComparison.Ordinal);
        Assert.True(first.GetProperty("occurrences").GetArrayLength() > 0);
    }

    [Fact]
    public void TextOutputNeverClaimsPlaceholderOrFullCompliance()
    {
        var result = new RuleAnalysisService().Analyze(Fixture("analysis_noncompliant.docx"));
        var text = RuleAnalysisReportWriter.ToText(result);

        Assert.Contains("=== 论文格式检查报告（规则分析）===", text);
        Assert.Contains("规则覆盖表", text);
        Assert.Contains("未检查清单", text);
        Assert.Contains("draft", text);
        Assert.DoesNotContain("占位模式", text, StringComparison.Ordinal);
        Assert.DoesNotContain("全文合规", text, StringComparison.Ordinal);
    }

    [Fact]
    public void AnnotatedCopyIsCreatedFromParagraphFindings()
    {
        var result = new RuleAnalysisService().Analyze(Fixture("analysis_noncompliant.docx"));
        var output = Path.Combine(Path.GetTempPath(), $"thesis-annotated-{Guid.NewGuid():N}.docx");
        try
        {
            var written = WordAnnotationWriter.WriteAnnotatedCopy(Fixture("analysis_noncompliant.docx"), output, result.BuildDiagnostics());
            Assert.Equal(Path.GetFullPath(output), written);
            Assert.True(File.Exists(output));
            Assert.True(new FileInfo(output).Length > 0);
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void ReportSwitchSuppressesDetailsButKeepsTheCount()
    {
        var ruleFile = Path.Combine(Path.GetTempPath(), $"thesis-report-{Guid.NewGuid():N}.yaml");
        File.WriteAllText(ruleFile, """
            schema_version: "0.2.0"
            status: draft
            profile:
              profile_id: report-switch
              rule_version: report-switch-1
              standard_version: "2026"
              status: draft
              selection: explicit
              scope_policy:
                include: [confirmed_body_prose]
                exclude: []
                unknown: unsupported
              source:
                title: test
                url: "https://example.com"
                version: "2026"
                locator: "4.7"
                quote: "test"
                verified_on: "2026-01-01"
            rules:
              - id: NKU-REQ-902
                matrix_id: M26-902
                scope: [body]
                property: font_size
                report: false
                unknown_policy: unsupported
                expect: {value: 12, unit: pt}
                severity: warning
                fix: manual
                source:
                  title: test
                  url: "https://example.com"
                  version: "2026"
                  locator: "4.7"
                  quote: "test"
                  verified_on: "2026-01-01"
            """);
        try
        {
            var result = new RuleAnalysisService().Analyze(Fixture("analysis_noncompliant.docx"), "report-switch", ruleFile);
            Assert.Empty(result.Findings);
            Assert.True(result.SuppressedFindings > 0);
            Assert.Contains(result.Coverage, row => row.RuleId == "NKU-REQ-902" && row.Failures > 0);
            Assert.Contains(result.UncheckedItems, item => item.Reason.Contains("预警开关", StringComparison.Ordinal));
        }
        finally
        {
            File.Delete(ruleFile);
        }
    }

    [Fact]
    public void LayoutProbeReadsSectionsIndentAndBold()
    {
        var probe = new DocxLayoutProbe().Read(Fixture("analysis_compliant.docx"));

        var section = Assert.Single(probe.Sections);
        Assert.Equal("A4", section.PageSize);
        Assert.Equal(3.8, section.TopCm!.Value, 2);
        Assert.Equal(3.8, section.BottomCm!.Value, 2);
        Assert.Equal(3.2, section.LeftCm!.Value, 2);
        Assert.Equal(3.2, section.RightCm!.Value, 2);

        var body = probe.Paragraphs[1];
        Assert.Equal("justify", body.Alignment);
        Assert.Equal("exact", body.LineSpacingKind);
        Assert.Equal(20, body.LineSpacingPt!.Value, 2);
        Assert.Equal(24, body.FirstLinePt!.Value, 2);

        var heading = probe.Paragraphs[0];
        Assert.True(heading.ValuesKnown);
        Assert.True(heading.Runs[0].BoldKnown);
        Assert.True(heading.Runs[0].Bold);
    }

    private static string Describe(RuleAnalysisResult result) =>
        string.Join("\n", result.Findings.Select(finding =>
            $"{finding.RuleId}|{finding.Location}|{finding.Expected}|{finding.Actual}|{finding.Occurrences.Count}"));

    private static string Fixture(string name) =>
        Path.Combine(RepositoryRoot(), "tests", "fixtures", "synthetic", name);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }
}
