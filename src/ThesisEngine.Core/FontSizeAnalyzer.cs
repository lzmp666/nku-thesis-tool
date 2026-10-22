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
    private const string FixLevelManual = "manual";

    public AnalysisResult Analyze(DocumentSnapshot snapshot, DemoRuleSet profile)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(profile);

        var diagnostics = new List<Diagnostic>();
        var unsupported = new List<string>(snapshot.UnsupportedScopes);
        var checkedRuns = 0;
        var checkedParagraphs = 0;

        foreach (var paragraph in snapshot.Paragraphs)
        {
            // Known headings, table cells, and figure/table captions have separate
            // requirements. Never apply body-prose demo values to them.
            if (paragraph.Scope == ParagraphScopeClassification.Excluded)
                continue;
            if (paragraph.Scope != ParagraphScopeClassification.ConfirmedBodyProse)
            {
                unsupported.Add("正文范围未确认：" + (paragraph.ScopeReason ?? "缺少可靠章节边界"));
                continue;
            }

            checkedParagraphs++;

            foreach (var run in paragraph.Runs)
            {
                if (string.IsNullOrWhiteSpace(run.Text)) continue;
                checkedRuns++;
                var location = $"{snapshot.PartName}#paragraph[{paragraph.ParagraphIndex}]/run[{run.RunIndex}]";

                if (ContainsHan(run.Text))
                    CheckEastAsiaFont(run, location, profile, diagnostics);
                if (ContainsLatinLetter(run.Text))
                    CheckAsciiFont(run, location, profile, diagnostics);
                CheckFontSize(run, location, profile, diagnostics);
            }
        }

        var coverage = $"已检查 {checkedRuns} 个正文文本片段，涉及 {checkedParagraphs} 个确认正文段落";

        return new AnalysisResult(
            Diagnostics: DiagnosticGrouper.Group(diagnostics),
            RuleVersion: profile.RuleVersion,
            Coverage: coverage,
            UnsupportedScopes: unsupported.Distinct(StringComparer.Ordinal).ToArray());
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
        // ASCII (0-127) and HAnsi (128-255) are separate OOXML slots. A run
        // can look correct in one slot while still carrying the wrong font in
        // the other, so compare both effective values before allowing repair.
        var ascii = run.AsciiSlotFont ?? run.AsciiFont;
        var hAnsi = run.HAnsiFont ?? run.AsciiFont;
        if (string.IsNullOrWhiteSpace(ascii) && string.IsNullOrWhiteSpace(hAnsi))
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

        var actual = DescribeWesternFont(ascii, hAnsi);
        if (!string.Equals(ascii, profile.AsciiFontExpected, StringComparison.Ordinal) ||
            !string.Equals(hAnsi, profile.AsciiFontExpected, StringComparison.Ordinal))
        {
            diagnostics.Add(new Diagnostic(
                RuleId: RuleIdAsciiFont,
                Severity: SeverityWarning,
                Location: location,
                Expected: profile.AsciiFontExpected,
                Actual: actual,
                // The Word add-in always shows a before/after preview and
                // requires explicit confirmation before applying this patch.
                FixLevel: FixLevelAuto,
                Message: $"ASCII/HAnsi 西文字体应统一为「{profile.AsciiFontExpected}」，实际为「{actual}」。"));
        }
    }

    private static string DescribeWesternFont(string? ascii, string? hAnsi)
    {
        var normalizedAscii = string.IsNullOrWhiteSpace(ascii) ? "(unknown)" : ascii;
        var normalizedHAnsi = string.IsNullOrWhiteSpace(hAnsi) ? "(unknown)" : hAnsi;
        return string.Equals(normalizedAscii, normalizedHAnsi, StringComparison.Ordinal)
            ? normalizedAscii
            : $"ASCII={normalizedAscii}; HAnsi={normalizedHAnsi}";
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

    private static bool ContainsHan(string text) => text.EnumerateRunes().Any(r =>
    {
        var value = r.Value;
        return value is >= 0x3400 and <= 0x4DBF or
               >= 0x4E00 and <= 0x9FFF or
               >= 0x20000 and <= 0x2FA1F;
    });

    private static bool ContainsLatinLetter(string text) => text.EnumerateRunes().Any(r =>
        System.Text.Rune.IsLetter(r) && (r.Value is >= 0x0041 and <= 0x024F or
                                         >= 0x1E00 and <= 0x1EFF or
                                         >= 0x2C60 and <= 0x2C7F or
                                         >= 0xA720 and <= 0xA7FF or
                                         >= 0xAB30 and <= 0xAB6F));
}
