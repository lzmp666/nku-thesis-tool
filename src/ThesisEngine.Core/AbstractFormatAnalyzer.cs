using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Core;

/// <summary>一条摘要格式问题。ParagraphIndex 为 1 基段落序号（document.xml 里 w:p 的顺序）。</summary>
public sealed record AbstractIssue(
    string Code,
    string Title,
    int ParagraphIndex,
    string Expected,
    string Actual,
    string Message);

/// <summary>识别到的摘要段落边界，便于报告里说明查了哪几段。</summary>
public sealed record AbstractSectionSpan(
    string Section,
    int? TitleParagraph,
    IReadOnlyList<int> BodyParagraphs,
    int? KeywordParagraph);

/// <summary>
/// 摘要格式分析结果。NotChecked 单独列出「压根没查 / 无法判定」，不混进 Issues。
/// </summary>
public sealed record AbstractFormatResult(
    IReadOnlyList<AbstractIssue> Issues,
    IReadOnlyList<AbstractSectionSpan> Sections,
    IReadOnlyList<string> NotChecked,
    IReadOnlyList<string> MissingParts,
    int CheckedParagraphs,
    string RuleVersion,
    string SourceNote);

/// <summary>
/// 中英文摘要格式检查（R4 / R5，依据 general.yaml 4.5 表4.4），只查不改。
///
/// 口径：
/// - 中文摘要标题 黑体18磅加粗居中、单倍行距、段前24磅段后18磅；正文宋体12磅、
///   固定值行距20磅、段前段后0磅，「关键词」三字加粗。
/// - 英文摘要标题 Arial 18磅加粗居中、单倍行距、段前24磅段后18磅；正文 Times New Roman
///   12磅、固定值行距20磅、段前段后0磅，「Key Words」两词加粗（此处按规则原文用 Arial，
///   与页眉 Abstract 用 TNR 不是一回事）。
/// - 只认样式链能展开到的显式值；判不出就列未判定，不拿默认值凑问题数。
/// - 规则状态为 draft（S1 2025 版条文，待 2026 版全文核验），提示需人工确认。
/// - 不识别「哪一节是摘要」的语义，只按标题文字定位；定位不到就列未检查，不假装合规。
/// </summary>
public sealed class AbstractFormatAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string SourceNote =
        "南开大学研究生学位论文写作规范 4.5 表4.4（general.yaml R4/R5，来源 S1 2025 版，规则状态 draft，需人工确认）";
    private const string RuleVersion = "nku-graduate-general-r4r5-draft";

    // 字号：w:sz 是半磅。18pt=36，12pt=24。
    private const int TitleSizeHalfPoints = 36;
    private const int BodySizeHalfPoints = 24;
    // 行距：twips。单倍 = auto/240；固定值 20 磅 = exact/400。
    private const int SingleLineTwips = 240;
    private const int Exact20Twips = 400;
    // 段距：twips，1 磅 = 20 twips。24 磅=480、18 磅=360。
    private const int TitleBeforeTwips = 480;
    private const int TitleAfterTwips = 360;
    private const int ZeroTwips = 0;

    private static readonly Regex ChineseTitlePattern = new(@"^中文摘要$|^摘\s*要$", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex EnglishTitlePattern = new(@"^Abstract$", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex ChineseKeywordPattern = new(@"^关\s*键\s*词", RegexOptions.Compiled | RegexOptions.CultureInvariant);
    private static readonly Regex EnglishKeywordPattern = new(@"^Key\s*words?\b", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    // 遇到这些标题即视为摘要区结束，避免把正文误当摘要正文。
    private static readonly Regex SectionBoundaryPattern = new(
        @"^第[一二三四五六七八九十百零〇0-9]+章|^Chapter\b|^目\s*录$|^Contents$|^参考文献$|^References$|^致\s*谢$|^附\s*录\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public AbstractFormatResult Analyze(string docxPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(docxPath);

        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var document = ReadPart(archive, "word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");
        var styles = ReadPart(archive, "word/styles.xml");

        var missingParts = new List<string>();
        if (styles is null) missingParts.Add("word/styles.xml");

        // 复用页眉/页脚的样式链展开逻辑读正文：run rPr → pStyle → basedOn → docDefaults。
        var reading = PartFormatReader.Read(document, "word/document.xml", styles);
        var readings = reading.Readings;
        var texts = readings.Select(r => r.Text.Trim()).ToList();

        var issues = new List<AbstractIssue>();
        var notChecked = new List<string>(reading.Undetermined.Where(u => u.StartsWith("word/document.xml", StringComparison.Ordinal)));
        var sections = new List<AbstractSectionSpan>();

        var cnTitleIndex = FindIndex(texts, ChineseTitlePattern);
        var enTitleIndex = FindIndex(texts, EnglishTitlePattern);

        if (cnTitleIndex < 0)
            notChecked.Add("中文摘要：未识别到「摘要」标题段落，R4 相关检查未执行");
        if (enTitleIndex < 0)
            notChecked.Add("英文摘要：未识别到「Abstract」标题段落，R5 相关检查未执行");

        if (cnTitleIndex >= 0)
        {
            // 中文摘要区止于英文摘要标题（若有），否则止于下一个分节标题或文末。
            var stop = enTitleIndex > cnTitleIndex
                ? enTitleIndex
                : SectionBoundaryAtOrAfter(texts, texts.Count, cnTitleIndex + 1);
            var keyword = FindIndex(texts, ChineseKeywordPattern, cnTitleIndex + 1, stop);
            var body = BodyRange(texts, cnTitleIndex + 1, keyword >= 0 ? keyword : stop, SectionBoundaryPattern);
            sections.Add(new AbstractSectionSpan("中文摘要", cnTitleIndex + 1, body.Select(i => i + 1).ToArray(),
                keyword >= 0 ? keyword + 1 : null));
            CheckTitle(readings[cnTitleIndex], cnTitleIndex + 1, "中文摘要", "黑体", isWesternTitle: false,
                PartFormatReader.IsSimHei, issues, notChecked);
            CheckBody(readings, body, "中文摘要", isWestern: false, PartFormatReader.IsSimSun,
                issues, notChecked, out _);
            CheckKeyword(readings, keyword, "关键词", "中文摘要", issues, notChecked);
        }

        if (enTitleIndex >= 0)
        {
            var stop = SectionBoundaryAtOrAfter(texts, texts.Count, enTitleIndex + 1);
            var keyword = FindIndex(texts, EnglishKeywordPattern, enTitleIndex + 1, stop);
            var body = BodyRange(texts, enTitleIndex + 1, keyword >= 0 ? keyword : stop, SectionBoundaryPattern);
            sections.Add(new AbstractSectionSpan("英文摘要", enTitleIndex + 1, body.Select(i => i + 1).ToArray(),
                keyword >= 0 ? keyword + 1 : null));
            CheckTitle(readings[enTitleIndex], enTitleIndex + 1, "英文摘要", "Arial", isWesternTitle: true,
                PartFormatReader.IsArial, issues, notChecked);
            CheckBody(readings, body, "英文摘要", isWestern: true, PartFormatReader.IsTimesNewRoman,
                issues, notChecked, out _);
            CheckKeyword(readings, keyword, "Key Words", "英文摘要", issues, notChecked);
        }

        if (styles is null)
            notChecked.Add("缺少 word/styles.xml：只按 run 上的显式 rPr / pPr 判定，样式链不可用");

        return new AbstractFormatResult(Issues: issues, Sections: sections, NotChecked: notChecked.Distinct(StringComparer.Ordinal).ToArray(),
            MissingParts: missingParts, CheckedParagraphs: CountChecked(readings, sections), RuleVersion: RuleVersion, SourceNote: SourceNote);
    }

    private static int CountChecked(IReadOnlyList<PartFormatReading> readings, IReadOnlyList<AbstractSectionSpan> sections)
    {
        var indices = new HashSet<int>();
        foreach (var s in sections)
        {
            if (s.TitleParagraph.HasValue) indices.Add(s.TitleParagraph.Value - 1);
            foreach (var p in s.BodyParagraphs) indices.Add(p - 1);
            if (s.KeywordParagraph.HasValue) indices.Add(s.KeywordParagraph.Value - 1);
        }
        return indices.Count(i => i >= 0 && i < readings.Count);
    }

    private static void CheckTitle(PartFormatReading r, int paragraphIndex, string section, string expectedFont,
        bool isWesternTitle, Func<string?, bool> fontMatch, List<AbstractIssue> issues, List<string> notChecked)
    {
        var codePrefix = isWesternTitle ? "A" : "A";
        var offset = isWesternTitle ? 9 : 0;

        // 字体：中文摘要标题看 eastAsia，英文摘要标题看西文字体槽。
        var font = isWesternTitle ? r.WesternFont : r.EastAsiaFont;
        if (string.IsNullOrWhiteSpace(font))
        {
            notChecked.Add($"{section}标题字体：样式链未解析出字体槽，未判定（第 {paragraphIndex} 段）");
        }
        else if (!fontMatch(font))
        {
            Add(issues, codePrefix + (1 + offset), $"{section}标题字体", paragraphIndex, expectedFont, font,
                $"{section}标题应为「{expectedFont}」，实际为「{font}」。依据 {SourceNote}。");
        }

        if (r.SizeHalfPoints is null)
        {
            notChecked.Add($"{section}标题字号：样式链未解析出字号，未判定（第 {paragraphIndex} 段）");
        }
        else if (r.SizeHalfPoints.Value != TitleSizeHalfPoints)
        {
            Add(issues, codePrefix + (2 + offset), $"{section}标题字号", paragraphIndex, "18磅", $"{r.SizeHalfPoints.Value / 2.0:0.#}磅",
                $"{section}标题应为 18 磅，实际为 {r.SizeHalfPoints.Value / 2.0:0.#} 磅。依据 {SourceNote}。");
        }

        if (!r.BoldResolved)
        {
            notChecked.Add($"{section}标题加粗：样式链未解析出加粗开关，未判定（第 {paragraphIndex} 段）");
        }
        else if (r.Bold != true)
        {
            Add(issues, codePrefix + (3 + offset), $"{section}标题加粗", paragraphIndex, "加粗", "未加粗",
                $"{section}标题应加粗，实际未加粗。依据 {SourceNote}。");
        }

        if (!r.AlignmentResolved)
        {
            notChecked.Add($"{section}标题对齐：样式链未解析出对齐，未判定（第 {paragraphIndex} 段）");
        }
        else if (!string.Equals(r.Alignment, "center", StringComparison.OrdinalIgnoreCase))
        {
            Add(issues, codePrefix + (4 + offset), $"{section}标题对齐", paragraphIndex, "居中", r.Alignment,
                $"{section}标题应居中，实际为「{r.Alignment}」。依据 {SourceNote}。");
        }

        CheckSpacing(r, paragraphIndex, $"{section}标题", singleLine: true,
            expectedBefore: TitleBeforeTwips, expectedAfter: TitleAfterTwips,
            code: codePrefix + (5 + offset), title: $"{section}标题行距与段距", issues, notChecked);
    }

    private static void CheckBody(IReadOnlyList<PartFormatReading> readings, IReadOnlyList<int> bodyIndices,
        string section, bool isWestern, Func<string?, bool> fontMatch, List<AbstractIssue> issues,
        List<string> notChecked, out int checkedBody)
    {
        checkedBody = 0;
        var offset = isWestern ? 9 : 0;
        foreach (var index in bodyIndices)
        {
            var r = readings[index];
            checkedBody++;

            var font = isWestern ? r.WesternFont : r.EastAsiaFont;
            if (string.IsNullOrWhiteSpace(font))
            {
                notChecked.Add($"{section}正文第 {index + 1} 段字体：样式链未解析出字体槽，未判定");
            }
            else if (!fontMatch(font))
            {
                Add(issues, "A" + (6 + offset), $"{section}正文字体", index + 1,
                    isWestern ? "Times New Roman" : "宋体", font,
                    $"{section}正文应用 {(isWestern ? "Times New Roman" : "宋体")}，实际为「{font}」。依据 {SourceNote}。");
            }

            if (r.SizeHalfPoints is null)
            {
                notChecked.Add($"{section}正文第 {index + 1} 段字号：样式链未解析出字号，未判定");
            }
            else if (r.SizeHalfPoints.Value != BodySizeHalfPoints)
            {
                Add(issues, "A" + (7 + offset), $"{section}正文字号", index + 1, "12磅", $"{r.SizeHalfPoints.Value / 2.0:0.#}磅",
                    $"{section}正文应为 12 磅，实际为 {r.SizeHalfPoints.Value / 2.0:0.#} 磅。依据 {SourceNote}。");
            }

            CheckSpacing(r, index + 1, $"{section}正文", singleLine: false,
                expectedBefore: ZeroTwips, expectedAfter: ZeroTwips,
                code: "A" + (8 + offset), title: $"{section}正文行距与段距", issues, notChecked);
        }

        if (bodyIndices.Count == 0)
            notChecked.Add($"{section}：标题后未识别到正文段落，正文格式未检查");
    }

    private static void CheckSpacing(PartFormatReading r, int paragraphIndex, string label, bool singleLine,
        int expectedBefore, int expectedAfter, string code, string title,
        List<AbstractIssue> issues, List<string> notChecked)
    {
        if (!r.SpacingResolved)
        {
            notChecked.Add($"{label}行距/段距：样式链未解析出完整 spacing，未判定（第 {paragraphIndex} 段）");
            return;
        }

        var ruleOk = singleLine
            ? string.Equals(r.LineSpacingRule, "auto", StringComparison.OrdinalIgnoreCase)
            : string.Equals(r.LineSpacingRule, "exact", StringComparison.OrdinalIgnoreCase);
        var lineOk = singleLine
            ? r.LineSpacingValue == SingleLineTwips
            : r.LineSpacingValue == Exact20Twips;
        var beforeOk = r.SpaceBeforeTwips == expectedBefore;
        var afterOk = r.SpaceAfterTwips == expectedAfter;

        if (ruleOk && lineOk && beforeOk && afterOk) return;

        var expectedLine = singleLine ? "单倍行距" : "固定值20磅";
        var actualLine = $"{r.LineSpacingRule}/{r.LineSpacingValue}";
        var expectedSpace = $"段前{expectedBefore / 20}磅/段后{expectedAfter / 20}磅";
        var actualSpace = $"段前{r.SpaceBeforeTwips / 20}磅/段后{r.SpaceAfterTwips / 20}磅";
        Add(issues, code, title, paragraphIndex, $"{expectedLine}，{expectedSpace}", $"{actualLine}，{actualSpace}",
            $"{label}应为 {expectedLine}、{expectedSpace}，实际为 {actualLine}、{actualSpace}。依据 {SourceNote}。");
    }

    private static void CheckKeyword(IReadOnlyList<PartFormatReading> readings, int keywordIndex, string expectedLabel,
        string section, List<AbstractIssue> issues, List<string> notChecked)
    {
        if (keywordIndex < 0)
        {
            notChecked.Add($"{section}：未识别到「{expectedLabel}」行，关键词加粗未检查");
            return;
        }

        var r = readings[keywordIndex];
        if (!r.BoldResolved)
        {
            notChecked.Add($"{section}「{expectedLabel}」加粗：样式链未解析出加粗开关，未判定（第 {keywordIndex + 1} 段）");
            return;
        }
        if (r.Bold != true)
        {
            var offset = section == "英文摘要" ? 9 : 0;
            Add(issues, "A" + (9 + offset), $"{section}{expectedLabel}加粗", keywordIndex + 1, "加粗", "未加粗",
                $"{section}「{expectedLabel}」应加粗，实际未加粗。依据 {SourceNote}。");
        }
    }

    private static void Add(List<AbstractIssue> issues, string code, string title, int paragraphIndex,
        string expected, string actual, string message)
        => issues.Add(new AbstractIssue(code, title, paragraphIndex, expected, actual, message));

    private static int FindIndex(IReadOnlyList<string> texts, Regex pattern, int start = 0, int end = -1)
    {
        var upper = end < 0 ? texts.Count : Math.Min(end, texts.Count);
        for (var i = Math.Max(0, start); i < upper; i++)
            if (pattern.IsMatch(texts[i])) return i;
        return -1;
    }

    private static int SectionBoundaryAtOrAfter(IReadOnlyList<string> texts, int fallback, int start)
    {
        for (var i = start; i < texts.Count; i++)
            if (SectionBoundaryPattern.IsMatch(texts[i])) return i;
        return fallback;
    }

    private static IReadOnlyList<int> BodyRange(IReadOnlyList<string> texts, int start, int end, Regex boundary)
    {
        var result = new List<int>();
        for (var i = start; i < end && i < texts.Count; i++)
        {
            if (string.IsNullOrWhiteSpace(texts[i])) continue;
            if (boundary.IsMatch(texts[i])) break;
            result.Add(i);
        }
        return result;
    }

    private static XDocument? ReadPart(ZipArchive archive, string entry)
    {
        var item = archive.GetEntry(entry);
        if (item is null) return null;
        using var stream = item.Open();
        return XDocument.Load(stream);
    }
}
