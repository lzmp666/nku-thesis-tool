using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Core;

/// <summary>一条正文段落格式问题。ParagraphIndex 为 document.xml 里 w:p 的 1 基序号。</summary>
public sealed record BodyTextIssue(
    string Code,
    string Title,
    int ParagraphIndex,
    string Expected,
    string Actual,
    string Message);

/// <summary>正文格式分析结果。NotChecked 单列「压根没查 / 无法判定」，不混进 Issues。</summary>
public sealed record BodyTextFormatResult(
    IReadOnlyList<BodyTextIssue> Issues,
    int ChapterStartParagraph,
    int SectionEndParagraph,
    IReadOnlyList<int> BodyParagraphs,
    IReadOnlyList<string> NotChecked,
    IReadOnlyList<string> MissingParts,
    int CheckedParagraphs,
    string RuleVersion,
    string SourceNote);

/// <summary>
/// 正文段落格式检查（R11 补齐，依据 general.yaml 4.7 表4.6，来源 S2026 2026 版），只查不改。
///
/// 范围边界（与「口径边界」一致）：
/// - R11 的字体（宋体/Times New Roman 12磅）三项已由 #43 的 NkuGraduate2026BodyAnalyzer 检查，
///   本分析器**不重复**字体字号，只补齐 demo_coverage 明确列为「未检查」的四项：
///   两端对齐、首行缩进 2 汉字符、固定值行距 20 磅、段前段后 0 磅。
/// - 正文区间：识别到的第一个章标题之后，到参考文献/附录/致谢等尾注标题之前。边界任意一侧
///   识别不到就如实列未检查，不拿封面/摘要冒充正文，也不假装合规。
/// - 标题段落、图表题注、表格内段落不是正文，跳过；这些各有专属规则。
/// - 含数学表达式的段落按条文豁免行距（「可根据表达需要设置该段的行距」），但对齐/缩进/段距照查。
/// - 只认样式链能展开到的显式值；判不出列未判定，不拿默认值凑问题数。
/// - 规则状态 approved（#43 已按 2026 版核对，全文已核验）。
/// </summary>
public sealed class BodyTextFormatAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private const string SourceNote =
        "南开大学研究生学位论文写作规范 4.7 表4.6（general.yaml R11，来源 S2026 2026 版全文，#43 已核对，规则状态 approved，RFC-1/4 已定案）";
    private const string RuleVersion = "nku-graduate-body-r11-layout-draft";

    // 行距/段距单位 twips：固定值 20 磅 = 400；1 磅 = 20 twips。
    private const int Exact20Twips = 400;
    private const int ZeroTwips = 0;
    // w:firstLineChars 以 1/100 字符计，2 汉字符 = 200。
    private const int TwoCharFirstLine = 200;
    // 对齐：两端对齐在 OOXML 里是 both（Word 里显示 justify）。
    private const string JustifyAlignment = "both";

    private static readonly Regex ChapterHeading = new(
        @"^(?:第[一二三四五六七八九十百零〇0-9]+章|Chapter\s+(?:[0-9]+|[IVXLCDM]+)\b)",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex EndHeading = new(
        @"^(?:主要)?参考文献$|^(?:References|Bibliography)$|^附\s*录\b|^Appendix(?:es)?\b|^致\s*谢$|^Acknowledg(?:e)?ments$"
        + @"|^攻读学位期间|^在学期间.{0,20}(?:成果|论文)|^(?:个人)?简历$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    // 图表题注：与 FigureTableAnalyzer 的口径一致，题注不属于正文。
    private static readonly Regex CaptionPattern = new(
        @"^(?:图|表|续表|Figure|Table|Fig\.)\s*[0-9]",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public BodyTextFormatResult Analyze(string docxPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(docxPath);

        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var document = ReadPart(archive, "word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");
        var styles = ReadPart(archive, "word/styles.xml");

        var missingParts = new List<string>();
        if (styles is null) missingParts.Add("word/styles.xml");

        var reading = PartFormatReader.Read(document, "word/document.xml", styles);
        var readings = reading.Readings;
        var tableParagraphs = FindTableParagraphs(document);
        var mathParagraphs = FindMathParagraphs(document);

        var issues = new List<BodyTextIssue>();
        var notChecked = new List<string>(reading.Undetermined.Where(u => u.StartsWith("word/document.xml", StringComparison.Ordinal)));

        // 正文边界：第一个章标题与随后第一个尾注标题之间。PartFormatReader 跳过无文字段，
        // 但保留原 w:p 序号，这里用 readings 的下标定位、输出原序号。
        var chapterStart = -1;
        var sectionEnd = -1;
        for (var i = 0; i < readings.Count; i++)
        {
            var text = readings[i].Text.Trim();
            if (chapterStart < 0 && ChapterHeading.IsMatch(text)) { chapterStart = i; continue; }
            if (chapterStart >= 0 && EndHeading.IsMatch(text)) { sectionEnd = i; break; }
        }

        var bodyIndices = new List<int>();
        if (chapterStart < 0)
            notChecked.Add("未识别到章标题，正文起点未确定，R11 排版检查未执行");
        if (chapterStart >= 0 && sectionEnd < 0)
            notChecked.Add("识别到章标题但未识别到参考文献/附录/致谢等结束标题，正文终点未确定，R11 排版检查未执行");

        if (chapterStart >= 0 && sectionEnd > chapterStart)
        {
            for (var i = chapterStart + 1; i < sectionEnd; i++)
            {
                var text = readings[i].Text.Trim();
                if (string.IsNullOrWhiteSpace(text)) continue;
                // 章内小节标题（如「1.1 xxx」）与题注、表格段落不是正文。
                if (ChapterHeading.IsMatch(text) || CaptionPattern.IsMatch(text)) continue;
                if (tableParagraphs.Contains(readings[i].ParagraphIndex)) continue;
                bodyIndices.Add(i);
            }
        }

        if (chapterStart >= 0 && sectionEnd > chapterStart && bodyIndices.Count == 0)
            notChecked.Add("正文区间内未识别到正文段落，R11 排版检查未执行");

        var mathExempted = false;
        foreach (var index in bodyIndices)
        {
            var r = readings[index];
            CheckBodyParagraph(r, mathParagraphs.Contains(r.ParagraphIndex),
                ref mathExempted, issues, notChecked);
        }

        notChecked.Add("正文字体与字号（宋体/Times New Roman 12磅）由 R11 的 #43 实现检查，本命令不重复判定");
        if (styles is null)
            notChecked.Add("缺少 word/styles.xml：只按 run 上的显式 rPr / pPr 判定，样式链不可用");

        return new BodyTextFormatResult(
            Issues: issues,
            ChapterStartParagraph: chapterStart >= 0 ? readings[chapterStart].ParagraphIndex : 0,
            SectionEndParagraph: sectionEnd >= 0 ? readings[sectionEnd].ParagraphIndex : 0,
            BodyParagraphs: bodyIndices.Select(i => readings[i].ParagraphIndex).ToArray(),
            NotChecked: notChecked.Distinct(StringComparer.Ordinal).ToArray(),
            MissingParts: missingParts,
            CheckedParagraphs: bodyIndices.Count,
            RuleVersion: RuleVersion,
            SourceNote: SourceNote);
    }

    private static void CheckBodyParagraph(PartFormatReading r, bool hasMath,
        ref bool mathExempted, List<BodyTextIssue> issues, List<string> notChecked)
    {
        var paragraphIndex = r.ParagraphIndex;
        var label = $"正文第 {paragraphIndex} 段";

        // 两端对齐
        if (!r.AlignmentResolved)
        {
            notChecked.Add($"{label}对齐：样式链未解析出对齐，未判定");
        }
        else if (!string.Equals(r.Alignment, JustifyAlignment, StringComparison.OrdinalIgnoreCase))
        {
            Add(issues, "T1", "正文两端对齐", paragraphIndex, JustifyAlignment, r.Alignment,
                $"{label}应两端对齐，实际为「{r.Alignment}」。依据 {SourceNote}。");
        }

        // 首行缩进 2 汉字符
        if (!HasTwoCharFirstLine(r))
        {
            if (r.FirstLineChars is null && r.FirstLineTwips is null)
                notChecked.Add($"{label}首行缩进：样式链未解析出 w:ind，未判定");
            else
                Add(issues, "T2", "正文首行缩进", paragraphIndex, "2 个汉字符", DescribeIndent(r),
                    $"{label}应首行缩进 2 个汉字符，实际为「{DescribeIndent(r)}」。依据 {SourceNote}。");
        }

        // 段前段后 0 磅（含公式段也要查）
        if (r.SpacingResolved && (r.SpaceBeforeTwips != ZeroTwips || r.SpaceAfterTwips != ZeroTwips))
        {
            Add(issues, "T4", "正文段距", paragraphIndex, "段前0磅/段后0磅",
                $"段前{r.SpaceBeforeTwips / 20}磅/段后{r.SpaceAfterTwips / 20}磅",
                $"{label}应段前后 0 磅，实际为段前{r.SpaceBeforeTwips / 20}磅/段后{r.SpaceAfterTwips / 20}磅。依据 {SourceNote}。");
        }
        else if (!r.SpacingResolved)
        {
            notChecked.Add($"{label}段距：样式链未解析出段前后，未判定");
        }

        // 固定值行距 20 磅；含公式段按条文豁免
        if (hasMath)
        {
            if (!mathExempted)
            {
                notChecked.Add("正文：含数学表达式的段落行距按表达需要调整，未按固定 20 磅判定（条文豁免，仍需人工确认实际行距合理）");
                mathExempted = true;
            }
            return;
        }

        if (!r.SpacingResolved)
        {
            // 段距未判定时上面已记，这里只补行距维度，避免重复。
            return;
        }

        var ruleOk = string.Equals(r.LineSpacingRule, "exact", StringComparison.OrdinalIgnoreCase);
        var lineOk = r.LineSpacingValue == Exact20Twips;
        if (!ruleOk || !lineOk)
        {
            Add(issues, "T3", "正文固定行距", paragraphIndex, "固定值20磅",
                $"{r.LineSpacingRule}/{r.LineSpacingValue}",
                $"{label}应为固定值 20 磅，实际为 {r.LineSpacingRule}/{r.LineSpacingValue}。依据 {SourceNote}。");
        }
    }

    private static bool HasTwoCharFirstLine(PartFormatReading r)
    {
        if (r.FirstLineChars is int chars) return chars == TwoCharFirstLine;
        if (r.FirstLineTwips is int twips && r.SizeHalfPoints is int size) return twips == size * 20;
        return false;
    }

    private static string DescribeIndent(PartFormatReading r)
    {
        if (r.FirstLineChars is int chars) return $"{chars / 100.0:0.#} 汉字符";
        if (r.FirstLineTwips is int twips) return $"{twips / 20.0:0.#} 磅";
        return "未解析出";
    }

    private static HashSet<int> FindMathParagraphs(XDocument document)
    {
        var result = new HashSet<int>();
        var index = 0;
        foreach (var paragraph in document.Root?.Descendants(W + "p") ?? Enumerable.Empty<XElement>())
        {
            index++;
            if (paragraph.Descendants(M + "oMath").Any() || paragraph.Descendants(M + "oMathPara").Any())
                result.Add(index);
        }
        return result;
    }

    private static HashSet<int> FindTableParagraphs(XDocument document)
    {
        var result = new HashSet<int>();
        var index = 0;
        foreach (var paragraph in document.Root?.Descendants(W + "p") ?? Enumerable.Empty<XElement>())
        {
            index++;
            if (paragraph.Ancestors(W + "tbl").Any()) result.Add(index);
        }
        return result;
    }

    private static void Add(List<BodyTextIssue> issues, string code, string title, int paragraphIndex,
        string expected, string actual, string message)
        => issues.Add(new BodyTextIssue(code, title, paragraphIndex, expected, actual, message));

    private static XDocument? ReadPart(ZipArchive archive, string entry)
    {
        var item = archive.GetEntry(entry);
        if (item is null) return null;
        using var stream = item.Open();
        return XDocument.Load(stream);
    }
}
