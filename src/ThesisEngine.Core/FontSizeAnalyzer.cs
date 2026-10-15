namespace ThesisEngine.Core;

/// <summary>
/// Demo-only 字体与字号规则诊断器。
/// 检查东亚字体、ASCII/HAnsi 字体、字号三项规则。
/// 所有规则均标记为 draft/demo，不得用于正式 NKU 合规检查。
/// </summary>
public sealed class FontSizeAnalyzer : IDemoAnalyzer
{
    // Rule IDs aligned with tests/fixtures/synthetic/demo/manifest.json (issue #27).
    public const string RuleIdEastAsiaFont = "demo/east-asian-font";
    public const string RuleIdAsciiFont = "demo/western-font";
    public const string RuleIdFontSize = "demo/font-size";

    private const string SeverityWarning = "warning";
    private const string FixLevelAuto = "auto";
    private const string FixLevelSuggest = "suggest";
    private const string FixLevelManual = "manual";

    public AnalysisResult Analyze(DocumentSnapshot snapshot, DemoRuleSet profile)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(profile);

        var diagnostics = new List<Diagnostic>();

        foreach (var paragraph in snapshot.Paragraphs)
        {
            foreach (var run in paragraph.Runs)
            {
                var location = $"{snapshot.PartName}#paragraph[{paragraph.ParagraphIndex}]/run[{run.RunIndex}]";

                CheckEastAsiaFont(run, location, profile, diagnostics);
                CheckAsciiFont(run, location, profile, diagnostics);
                CheckFontSize(run, location, profile, diagnostics);
            }
        }

        var coverage = ComputeCoverage(snapshot);

        return new AnalysisResult(
            Diagnostics: diagnostics,
            RuleVersion: profile.RuleVersion,
            Coverage: coverage,
            UnsupportedScopes: snapshot.UnsupportedScopes);
    }

    private static void CheckEastAsiaFont(
        TextRunSnapshot run,
        string location,
        DemoRuleSet profile,
        List<Diagnostic> diagnostics)
    {
        if (string.IsNullOrWhiteSpace(run.EastAsiaFont))
        {
            // 未知值不报告为合规，也不报告为不合规——标记为无法判定
            // 按 demo 规则，未知值不得当作通过，记为 warning
            diagnostics.Add(new Diagnostic(
                RuleId: RuleIdEastAsiaFont,
                Severity: SeverityWarning,
                Location: location,
                Expected: profile.EastAsiaFontExpected,
                Actual: "(unknown)",
                FixLevel: FixLevelManual,
                Message: "东亚字体未设置或无法解析，无法判定是否符合要求。"));
            return;
        }

        if (!string.Equals(run.EastAsiaFont, profile.EastAsiaFontExpected, StringComparison.Ordinal))
        {
            diagnostics.Add(new Diagnostic(
                RuleId: RuleIdEastAsiaFont,
                Severity: SeverityWarning,
                Location: location,
                Expected: profile.EastAsiaFontExpected,
                Actual: run.EastAsiaFont,
                FixLevel: FixLevelAuto,
                Message: $"东亚字体应为「{profile.EastAsiaFontExpected}」，实际为「{run.EastAsiaFont}」。"));
        }
    }

    private static void CheckAsciiFont(
        TextRunSnapshot run,
        string location,
        DemoRuleSet profile,
        List<Diagnostic> diagnostics)
    {
        if (string.IsNullOrWhiteSpace(run.AsciiFont))
        {
            diagnostics.Add(new Diagnostic(
                RuleId: RuleIdAsciiFont,
                Severity: SeverityWarning,
                Location: location,
                Expected: profile.AsciiFontExpected,
                Actual: "(unknown)",
                FixLevel: FixLevelManual,
                Message: "ASCII/HAnsi 字体未设置或无法解析，无法判定是否符合要求。"));
            return;
        }

        if (!string.Equals(run.AsciiFont, profile.AsciiFontExpected, StringComparison.Ordinal))
        {
            diagnostics.Add(new Diagnostic(
                RuleId: RuleIdAsciiFont,
                Severity: SeverityWarning,
                Location: location,
                Expected: profile.AsciiFontExpected,
                Actual: run.AsciiFont,
                FixLevel: FixLevelSuggest,
                Message: $"ASCII/HAnsi 字体应为「{profile.AsciiFontExpected}」，实际为「{run.AsciiFont}」。"));
        }
    }

    private static void CheckFontSize(
        TextRunSnapshot run,
        string location,
        DemoRuleSet profile,
        List<Diagnostic> diagnostics)
    {
        if (run.FontSizePt is null)
        {
            diagnostics.Add(new Diagnostic(
                RuleId: RuleIdFontSize,
                Severity: SeverityWarning,
                Location: location,
                Expected: $"{profile.FontSizeExpectedPt}pt",
                Actual: "(unknown)",
                FixLevel: FixLevelManual,
                Message: "字号未设置或无法解析，无法判定是否符合要求。"));
            return;
        }

        if (Math.Abs(run.FontSizePt.Value - profile.FontSizeExpectedPt) > 0.01)
        {
            diagnostics.Add(new Diagnostic(
                RuleId: RuleIdFontSize,
                Severity: SeverityWarning,
                Location: location,
                Expected: $"{profile.FontSizeExpectedPt}pt",
                Actual: $"{run.FontSizePt.Value}pt",
                FixLevel: FixLevelAuto,
                Message: $"字号应为 {profile.FontSizeExpectedPt}pt，实际为 {run.FontSizePt.Value}pt。"));
        }
    }

    private static string ComputeCoverage(DocumentSnapshot snapshot)
    {
        var totalRuns = snapshot.Paragraphs.Sum(p => p.Runs.Count);
        return $"已检查 {totalRuns} 个文本片段，涉及 {snapshot.Paragraphs.Count} 个段落";
    }
}
