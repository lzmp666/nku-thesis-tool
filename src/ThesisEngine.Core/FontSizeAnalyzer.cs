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
    public const string RuleIdCaptionEastAsiaFont = "demo/caption-east-asian-font";
    public const string RuleIdCaptionFontSize = "demo/caption-font-size";
    public const string RuleIdCaptionAlignment = "demo/caption-alignment";
    public const string RuleIdCaptionLineSpacing = "demo/caption-line-spacing";
    public const string RuleIdCaptionSpaceBefore = "demo/caption-space-before";
    public const string RuleIdCaptionSpaceAfter = "demo/caption-space-after";
    public const string RuleIdCaptionPlacement = "demo/caption-placement";
    public const string RuleIdCaptionNumberGap = "demo/caption-number-gap";

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
        var checkedCaptions = 0;

        foreach (var paragraph in snapshot.Paragraphs)
        {
            if (paragraph.Caption is not null)
            {
                checkedCaptions++;
                CheckCaption(paragraph, diagnostics);
                continue;
            }

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

        var coverage = $"已检查 {checkedRuns} 个正文文本片段，涉及 {checkedParagraphs} 个确认正文段落；识别 {checkedCaptions} 条图/表题注";

        return new AnalysisResult(
            Diagnostics: DiagnosticGrouper.Group(diagnostics),
            RuleVersion: profile.RuleVersion,
            Coverage: coverage,
            UnsupportedScopes: unsupported.Distinct(StringComparer.Ordinal).ToArray());
    }

    private static void CheckCaption(ParagraphSnapshot paragraph, List<Diagnostic> diagnostics)
    {
        var captionName = paragraph.Caption == CaptionKind.Figure ? "图题" : "表题";
        var location = paragraph.Runs.Count == 0
            ? null
            : $"word/document.xml#paragraph[{paragraph.ParagraphIndex}]/run[{paragraph.Runs[0].RunIndex}]";
        if (location is null) return;

        foreach (var run in paragraph.Runs.Where(run => !string.IsNullOrWhiteSpace(run.Text)))
        {
            var runLocation = $"word/document.xml#paragraph[{paragraph.ParagraphIndex}]/run[{run.RunIndex}]";
            if (ContainsHan(run.Text))
                CheckCaptionValue(run.EastAsiaFont, "宋体", RuleIdCaptionEastAsiaFont, "东亚字体",
                    captionName, run.HasUnresolvedStyle, runLocation, diagnostics);
            CheckCaptionValue(run.FontSizePt?.ToString(System.Globalization.CultureInfo.InvariantCulture) + "pt",
                "10.5pt", RuleIdCaptionFontSize, "字号", captionName,
                run.HasUnresolvedStyle || run.FontSizePt is null, runLocation, diagnostics);
        }

        var format = paragraph.ParagraphFormatting;
        if (format is null || format.HasUnresolvedStyle)
        {
            AddCaptionDiagnostic(RuleIdCaptionAlignment, location, "居中", "无法解析",
                captionName + "对齐方式无法可靠解析，需要人工检查。", "manual", diagnostics);
            AddCaptionDiagnostic(RuleIdCaptionLineSpacing, location, "单倍", "无法解析",
                captionName + "行距无法可靠解析，需要人工检查。", "manual", diagnostics);
            AddCaptionDiagnostic(RuleIdCaptionSpaceBefore, location, "6pt", "无法解析",
                captionName + "段前距无法可靠解析，需要人工检查。", "manual", diagnostics);
            AddCaptionDiagnostic(RuleIdCaptionSpaceAfter, location,
                paragraph.Caption == CaptionKind.Figure ? "12pt" : "6pt", "无法解析",
                captionName + "段后距无法可靠解析，需要人工检查。", "manual", diagnostics);
        }
        else
        {
            CheckCaptionValue(format.Alignment, "center", RuleIdCaptionAlignment, "对齐方式",
                captionName, false, location, diagnostics);
            CheckCaptionValue(format.LineSpacing, "single", RuleIdCaptionLineSpacing, "行距",
                captionName, false, location, diagnostics);
            CheckCaptionValue(format.SpaceBeforePt.ToString(System.Globalization.CultureInfo.InvariantCulture) + "pt",
                "6pt", RuleIdCaptionSpaceBefore, "段前距", captionName, false, location, diagnostics);
            var expectedAfter = paragraph.Caption == CaptionKind.Figure ? "12pt" : "6pt";
            CheckCaptionValue(format.SpaceAfterPt.ToString(System.Globalization.CultureInfo.InvariantCulture) + "pt",
                expectedAfter, RuleIdCaptionSpaceAfter, "段后距", captionName, false, location, diagnostics);
        }

        if (paragraph.CaptionPlacementCorrect != true)
        {
            var expectedPosition = paragraph.Caption == CaptionKind.Figure ? "图片下方" : "表格上方";
            var actualPosition = paragraph.CaptionPlacementCorrect is null ? "无法确认对应对象" : "位置不符合要求";
            AddCaptionDiagnostic(RuleIdCaptionPlacement, location, expectedPosition, actualPosition,
                captionName + "应置于" + expectedPosition + "；系统不会自动移动内容，请人工调整。", "manual", diagnostics);
        }

        var labelMatch = System.Text.RegularExpressions.Regex.Match(paragraph.Text.TrimStart(),
            @"^(?:图|表)\s*(?:[A-Z]|[0-9]+)(?:[.\-—][0-9]+)*");
        if (!labelMatch.Success || labelMatch.Length >= paragraph.Text.TrimStart().Length ||
            paragraph.Text.TrimStart()[labelMatch.Length] != '\u3000')
        {
            var actualGap = labelMatch.Success && labelMatch.Length < paragraph.Text.TrimStart().Length
                ? "序号后不是一个全角空格"
                : "未识别到完整题名或序号";
            AddCaptionDiagnostic(RuleIdCaptionNumberGap, location, "序号与题名间一个汉字符宽度", actualGap,
                captionName + "序号与题名之间应空一个汉字符宽度；这会改变题名文本，需人工调整。", "manual", diagnostics);
        }
    }

    private static void CheckCaptionValue(string? actual, string expected, string ruleId,
        string propertyName, string captionName, bool unresolved, string location,
        List<Diagnostic> diagnostics)
    {
        if (unresolved || string.IsNullOrWhiteSpace(actual) || string.Equals(actual, "pt", StringComparison.Ordinal))
        {
            AddCaptionDiagnostic(ruleId, location, expected, "无法解析",
                captionName + propertyName + "无法可靠解析，保留人工检查。", "manual", diagnostics);
            return;
        }

        var equal = ruleId == RuleIdCaptionFontSize
            ? PatchPlanner.TryParsePoints(actual, out var actualPoints) &&
              PatchPlanner.TryParsePoints(expected, out var expectedPoints) && Math.Abs(actualPoints - expectedPoints) <= 0.01
            : string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        if (!equal)
            AddCaptionDiagnostic(ruleId, location, expected, actual,
                captionName + propertyName + "应为「" + expected + "」，实际为「" + actual + "」。",
                "auto", diagnostics);
    }

    private static void AddCaptionDiagnostic(string ruleId, string location, string expected,
        string actual, string detail, string fixLevel, List<Diagnostic> diagnostics)
    {
        var source = "依据：南开大学《研究生学位论文写作规范（2026版）》第 4.7 节表 4.6。";
        diagnostics.Add(new Diagnostic(ruleId, "warning", location, expected, actual, fixLevel,
            source + " " + detail));
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
