using System.Globalization;
using System.Text;

namespace ThesisEngine.Core;

/// <summary>
/// An explicitly selected, draft-only profile for the first three 2026 body-text rules.
/// It deliberately has no automatic repair behavior and is not the Demo profile.
/// </summary>
public sealed record NkuGraduate2026BodyProfile(
    string ProfileId,
    string RuleVersion,
    string Status,
    string ChineseFontExpected,
    string LatinFontExpected,
    double FontSizeExpectedPt)
{
    public static NkuGraduate2026BodyProfile Draft => new(
        ProfileId: "nku-graduate-2026-body",
        RuleVersion: "nku-graduate-2026-body-v1",
        Status: "draft",
        ChineseFontExpected: "宋体",
        LatinFontExpected: "Times New Roman",
        FontSizeExpectedPt: 12.0);
}

/// <summary>
/// Draft analyzer for confirmed body prose only. Unknown scope or effective formatting is
/// reported as unsupported and never converted into a pass or a violation.
/// </summary>
public sealed class NkuGraduate2026BodyAnalyzer
{
    public const string RuleIdChineseFont = "NKU-REQ-042";
    public const string RuleIdLatinFont = "NKU-REQ-043";
    public const string RuleIdFontSize = "NKU-REQ-044";

    private const string Source = "南开大学研究生学位论文写作规范（2026版）第4.7节表4.6（M26-036）";

    public AnalysisResult Analyze(DocumentSnapshot snapshot, NkuGraduate2026BodyProfile profile)
    {
        ArgumentNullException.ThrowIfNull(snapshot);
        ArgumentNullException.ThrowIfNull(profile);
        if (!string.Equals(profile.Status, "draft", StringComparison.Ordinal))
            throw new ArgumentException("规则仍为草案，完成适用范围与样例验收前不能按已批准规则执行。", nameof(profile));
        var implemented = NkuGraduate2026BodyProfile.Draft;
        if (!string.Equals(profile.ProfileId, implemented.ProfileId, StringComparison.Ordinal) ||
            !string.Equals(profile.RuleVersion, implemented.RuleVersion, StringComparison.Ordinal) ||
            !string.Equals(profile.ChineseFontExpected, implemented.ChineseFontExpected, StringComparison.Ordinal) ||
            !string.Equals(profile.LatinFontExpected, implemented.LatinFontExpected, StringComparison.Ordinal) ||
            profile.FontSizeExpectedPt != implemented.FontSizeExpectedPt)
            throw new ArgumentException("所选配置的身份、版本或规则值与已实现的 2026 正文草案不匹配。", nameof(profile));

        var diagnostics = new List<Diagnostic>();
        var unsupported = new List<string>(snapshot.UnsupportedScopes);
        var checkedRuns = 0;

        foreach (var paragraph in snapshot.Paragraphs)
        {
            if (paragraph.Scope == ParagraphScopeClassification.Excluded)
                continue;

            if (paragraph.Scope != ParagraphScopeClassification.ConfirmedBodyProse)
            {
                unsupported.Add($"正文范围未确认：paragraph[{paragraph.ParagraphIndex}]（{paragraph.ScopeReason ?? "缺少可靠章节边界"}）");
                continue;
            }

            foreach (var run in paragraph.Runs)
            {
                if (string.IsNullOrWhiteSpace(run.Text))
                    continue;

                var location = $"{snapshot.PartName}#paragraph[{paragraph.ParagraphIndex}]/run[{run.RunIndex}]";
                if (run.HasUnresolvedStyle)
                {
                    unsupported.Add($"有效样式链无法解析：{location}");
                    continue;
                }

                checkedRuns++;
                CheckFontSize(run, location, profile, diagnostics, unsupported);
                if (ContainsHan(run.Text))
                    CheckFont(run.EastAsiaFont, profile.ChineseFontExpected, RuleIdChineseFont,
                        "中文字体", "eastAsia", location, diagnostics, unsupported);
                if (ContainsLatinLetter(run.Text))
                    CheckLatinFont(run, location, profile, diagnostics, unsupported);
            }
        }

        var distinctUnsupported = unsupported.Distinct(StringComparer.Ordinal).ToArray();
        var coverage = $"规则状态：草案（{profile.Status}）；已检查 {checkedRuns} 个已确认正文文本片段；未覆盖 {distinctUnsupported.Length} 项";
        return new AnalysisResult(DiagnosticGrouper.Group(diagnostics), profile.RuleVersion, coverage, distinctUnsupported);
    }

    private static void CheckFontSize(TextRunSnapshot run, string location, NkuGraduate2026BodyProfile profile,
        List<Diagnostic> diagnostics, List<string> unsupported)
    {
        if (run.FontSizePt is null)
        {
            unsupported.Add($"字号有效值未知：{location}");
            return;
        }

        if (Math.Abs(run.FontSizePt.Value - profile.FontSizeExpectedPt) > 0.01)
        {
            var expected = profile.FontSizeExpectedPt.ToString("0.##", CultureInfo.InvariantCulture);
            var actual = run.FontSizePt.Value.ToString("0.##", CultureInfo.InvariantCulture);
            diagnostics.Add(new Diagnostic(RuleIdFontSize, "warning", location,
                $"{expected}pt", $"{actual}pt", "manual",
                $"{Source}：正文字号应为 {expected}pt，实际为 {actual}pt。规则仍为草案，需人工确认。"));
        }
    }

    private static void CheckFont(string? actual, string expected, string ruleId, string label, string slot,
        string location, List<Diagnostic> diagnostics, List<string> unsupported)
    {
        if (string.IsNullOrWhiteSpace(actual))
        {
            unsupported.Add($"{slot} 字体槽有效值未知：{location}");
            return;
        }

        if (!string.Equals(actual, expected, StringComparison.Ordinal))
        {
            diagnostics.Add(new Diagnostic(ruleId, "warning", location, expected, actual, "manual",
                $"{Source}：{label}应为「{expected}」，实际为「{actual}」。规则仍为草案，需人工确认。"));
        }
    }

    private static void CheckLatinFont(TextRunSnapshot run, string location, NkuGraduate2026BodyProfile profile,
        List<Diagnostic> diagnostics, List<string> unsupported)
    {
        if (string.IsNullOrWhiteSpace(run.AsciiSlotFont) || string.IsNullOrWhiteSpace(run.HAnsiFont))
        {
            unsupported.Add($"ascii 与 hAnsi 字体槽未能同时解析：{location}");
            return;
        }

        if (!string.Equals(run.AsciiSlotFont, profile.LatinFontExpected, StringComparison.Ordinal) ||
            !string.Equals(run.HAnsiFont, profile.LatinFontExpected, StringComparison.Ordinal))
        {
            var actual = $"ascii={run.AsciiSlotFont}; hAnsi={run.HAnsiFont}";
            diagnostics.Add(new Diagnostic(RuleIdLatinFont, "warning", location,
                $"ascii={profile.LatinFontExpected}; hAnsi={profile.LatinFontExpected}", actual, "manual",
                $"{Source}：英文文字的 ascii 与 hAnsi 字体均应为「{profile.LatinFontExpected}」，实际为「{actual}」。规则仍为草案，需人工确认。"));
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
        Rune.IsLetter(r) && (r.Value is >= 0x0041 and <= 0x024F or
                             >= 0x1E00 and <= 0x1EFF or
                             >= 0x2C60 and <= 0x2C7F or
                             >= 0xA720 and <= 0xA7FF or
                             >= 0xAB30 and <= 0xAB6F));
}

/// <summary>
/// Explicit-profile composition for callers that want to analyze a DOCX with the 2026 draft.
/// No default profile is selected by this service.
/// </summary>
public sealed class NkuGraduate2026BodyAnalysisService
{
    private readonly WordFormattingReader _reader;
    private readonly NkuGraduate2026BodyAnalyzer _analyzer;

    public NkuGraduate2026BodyAnalysisService(WordFormattingReader? reader = null, NkuGraduate2026BodyAnalyzer? analyzer = null)
    {
        _reader = reader ?? new WordFormattingReader();
        _analyzer = analyzer ?? new NkuGraduate2026BodyAnalyzer();
    }

    public AnalysisResult Analyze(string path, NkuGraduate2026BodyProfile profile)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(path);
        ArgumentNullException.ThrowIfNull(profile);
        return _analyzer.Analyze(_reader.Read(path), profile);
    }
}

internal sealed record BodyScopeInput(
    int ParagraphIndex,
    string Text,
    bool IsHeading,
    bool HasUnsupportedInlineContent,
    bool IsTableCell = false,
    bool IsCaption = false,
    bool HasEquation = false);

internal sealed record BodyScopeDecision(ParagraphScopeClassification Classification, string? Reason);

/// <summary>
/// Confirms only ordinary paragraphs between a recognized chapter heading and a recognized
/// non-body section heading. Missing either boundary fails closed to unknown.
/// </summary>
internal static class NkuGraduate2026BodyScopeClassifier
{
    private static readonly System.Text.RegularExpressions.Regex ChapterHeading = new(
        @"^(?:第[一二三四五六七八九十百零〇0-9]+章|Chapter\s+(?:[0-9]+|[IVXLCDM]+)\b)",
        System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.IgnoreCase |
        System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    private static readonly System.Text.RegularExpressions.Regex EndHeading = new(
        @"^(参考文献|主要参考文献|References|Bibliography|附\s*录|Appendix|致\s*谢|Acknowledg(e)?ments|攻读学位期间|个人简历|在学期间成果)",
        System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    private static readonly System.Text.RegularExpressions.Regex KnownExcludedHeading = new(
        @"^(摘要|中文摘要|英文摘要|Abstract|ABSTRACT|目录|Contents|CONTENTS|图目录|表目录|封面|声明|授权书)\s*$",
        System.Text.RegularExpressions.RegexOptions.Compiled | System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant);

    public static IReadOnlyList<BodyScopeDecision> Classify(IReadOnlyList<BodyScopeInput> paragraphs)
    {
        ArgumentNullException.ThrowIfNull(paragraphs);
        var chapterStart = -1;
        var bodyEnd = -1;
        for (var i = 0; i < paragraphs.Count; i++)
        {
            if (!paragraphs[i].IsHeading) continue;
            var headingText = paragraphs[i].Text.Trim();
            if (chapterStart < 0 && ChapterHeading.IsMatch(headingText))
                chapterStart = i;
            else if (chapterStart >= 0 && EndHeading.IsMatch(headingText))
            {
                bodyEnd = i;
                break;
            }
        }

        var boundariesKnown = chapterStart >= 0 && bodyEnd > chapterStart;
        var decisions = new BodyScopeDecision[paragraphs.Count];
        for (var i = 0; i < paragraphs.Count; i++)
        {
            var paragraph = paragraphs[i];
            if (paragraph.IsTableCell)
            {
                decisions[i] = new(ParagraphScopeClassification.Excluded, "表格单元格；正文规则不适用，表格专属规则尚未启用");
            }
            else if (paragraph.IsCaption)
            {
                decisions[i] = new(ParagraphScopeClassification.Excluded, "图表题注应使用题注专属规则");
            }
            else if (string.IsNullOrWhiteSpace(paragraph.Text) ||
                paragraph.IsHeading && KnownExcludedHeading.IsMatch(paragraph.Text.Trim()) ||
                boundariesKnown && paragraph.IsHeading && (i == chapterStart || i == bodyEnd || i > chapterStart && i < bodyEnd))
            {
                decisions[i] = new(ParagraphScopeClassification.Excluded, "标题或非正文部件");
            }
            else if (!boundariesKnown)
            {
                decisions[i] = new(ParagraphScopeClassification.Unknown,
                    "未同时识别到正文起始章标题和正文结束部件标题");
            }
            else if (i <= chapterStart || i >= bodyEnd)
            {
                decisions[i] = new(ParagraphScopeClassification.Unknown, "位于已确认正文边界之外");
            }
            else if (paragraph.HasUnsupportedInlineContent)
            {
                decisions[i] = new(ParagraphScopeClassification.Unknown, "含字段、修订或其他未支持的内联结构");
            }
            else if (paragraph.IsHeading)
            {
                decisions[i] = new(ParagraphScopeClassification.Excluded, "正文中的标题段落不适用本批正文段落规则");
            }
            else if (paragraph.HasEquation)
            {
                decisions[i] = new(ParagraphScopeClassification.ConfirmedBodyProse,
                    "含公式对象；仅检查普通文字，公式对象单独标记为未覆盖");
            }
            else
            {
                decisions[i] = new(ParagraphScopeClassification.ConfirmedBodyProse, null);
            }
        }

        return decisions;
    }
}
