using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Core;

/// <summary>一条尾注部件格式问题。ParagraphIndex 为 document.xml 里 w:p 的 1 基序号。</summary>
public sealed record BackMatterIssue(
    string Code,
    string Title,
    string Section,
    int ParagraphIndex,
    string Expected,
    string Actual,
    string Message);

/// <summary>识别到的尾注部件边界，便于报告里说明查了哪几段。</summary>
public sealed record BackMatterSectionSpan(
    string Section,
    string RuleId,
    int TitleParagraph,
    IReadOnlyList<int> BodyParagraphs);

/// <summary>尾注部件格式分析结果。NotChecked 单列「压根没查 / 无法判定」，不混进 Issues。</summary>
public sealed record BackMatterFormatResult(
    IReadOnlyList<BackMatterIssue> Issues,
    IReadOnlyList<BackMatterSectionSpan> Sections,
    IReadOnlyList<string> NotChecked,
    IReadOnlyList<string> MissingParts,
    int CheckedParagraphs,
    string RuleVersion,
    string SourceNote);

/// <summary>
/// 尾注部件正文格式检查（R14/R15/R16/R17，依据 general.yaml 4.8 表4.7），只查不改。
///
/// 口径：
/// - 参考文献：宋体 / Times New Roman 10.5 磅，固定值行距 16 磅，段前后 0 磅。
/// - 附录：宋体 / Times New Roman 12 磅，两端对齐，首行缩进 2 汉字符，固定值行距 20 磅，段前后 0 磅。
///   条文允许含数学表达式的段落按表达需要调整行距，因此含 OMML 公式的段落只豁免行距，其余维度照查。
/// - 符号注释表 / 个人简历及成果：宋体 / Times New Roman 10.5 磅，固定值行距 16 磅，段前后 0 磅。
/// - 勘误页 / 致谢：仿宋 12 磅，固定值行距 16 磅，段前后 0 磅。
/// - 各部件标题「同各章标题（R7）」由标题检查负责；本分析器只判正文段落，标题格式列入未检查，不冒充合规。
/// - 只认样式链能展开到的显式值；判不出列未判定，不拿默认值凑问题数。
/// - 规则状态 draft（S1 条文），提示需人工确认。
/// </summary>
public sealed class BackMatterFormatAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace M = "http://schemas.openxmlformats.org/officeDocument/2006/math";

    private const string SourceNote =
        "南开大学研究生学位论文写作规范 4.8 表4.7（general.yaml R14/R15/R16/R17，来源 S1 2025 版，规则状态 draft，需人工确认）";
    private const string RuleVersion = "nku-graduate-general-r14r17-draft";

    // w:sz 是半磅：10.5pt=21、12pt=24。行距是 twips：16pt=320、20pt=400。
    private const int SizeSmall4HalfPoints = 21;
    private const int Size4HalfPoints = 24;
    private const int Exact16Twips = 320;
    private const int Exact20Twips = 400;
    private const int ZeroTwips = 0;
    // w:firstLineChars 以 1/100 字符计，2 汉字符 = 200。
    private const int TwoCharFirstLine = 200;

    private static readonly Regex ReferencesTitle = new(
        @"^(?:主要)?参考文献$|^(?:References|Bibliography)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex AppendixTitle = new(
        @"^附\s*录(?:\s*[A-Z0-9一二三四五六七八九十]+)?(?:\s*[:：、.．\-—]\s*.+)?$"
        + @"|^附\s*录\s+[A-Z0-9一二三四五六七八九十]+\s+.+$"
        + @"|^Appendix(?:es)?(?:\s*[A-Z0-9]+)?(?:\s*[:：.．\-—]\s*.+)?$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex SymbolCvTitle = new(
        @"^(?:主要)?(?:符号|标志|缩略语|简称).{0,12}(?:注释?表|说明|表)$|^(?:个人)?简历$|^在学期间.{0,20}(?:成果|论文)|^攻读学位期间.{0,20}(?:成果|论文)|^(?:Resume|Curriculum\s*Vitae)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex AcknowledgementTitle = new(
        @"^致\s*谢$|^Acknowledg(?:e)?ments$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private sealed record SectionSpec(
        string Name,
        string RuleId,
        Regex TitlePattern,
        string? CnFont,
        Func<string?, bool>? CnFontMatch,
        string? WesternFont,
        int SizeHalfPoints,
        string? Alignment,
        bool RequireTwoCharFirstLine,
        int ExactLineTwips,
        int CodeStart,
        bool ExemptMathLineSpacing,
        IReadOnlyList<int> DimensionOffsets);

    // 下标 1..6 依次对应：中文字体、西文字体、字号、对齐、首行缩进、行距与段距。
    // 未适用的维度留 0，保证每个部件的问题码连续且不冲突。
    private static readonly int[] RefOffsets = { 0, 1, 2, 3, 0, 0, 4 };
    private static readonly int[] AppendixOffsets = { 0, 1, 2, 3, 4, 5, 6 };
    private static readonly int[] SymbolOffsets = { 0, 1, 2, 3, 0, 0, 4 };
    private static readonly int[] AckOffsets = { 0, 1, 0, 2, 0, 0, 3 };

    private static readonly IReadOnlyList<SectionSpec> Sections = new[]
    {
        new SectionSpec("参考文献", "R14", ReferencesTitle,
            "宋体", PartFormatReader.IsSimSun, "Times New Roman",
            SizeSmall4HalfPoints, null, false, Exact16Twips, 0, false, RefOffsets),
        new SectionSpec("附录", "R15", AppendixTitle,
            "宋体", PartFormatReader.IsSimSun, "Times New Roman",
            Size4HalfPoints, "both", true, Exact20Twips, 4, true, AppendixOffsets),
        new SectionSpec("符号表与个人简历", "R16", SymbolCvTitle,
            "宋体", PartFormatReader.IsSimSun, "Times New Roman",
            SizeSmall4HalfPoints, null, false, Exact16Twips, 10, false, SymbolOffsets),
        new SectionSpec("致谢", "R17", AcknowledgementTitle,
            "仿宋", IsFangSong, null,
            Size4HalfPoints, null, false, Exact16Twips, 14, false, AckOffsets),
    };

    public BackMatterFormatResult Analyze(string docxPath)
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
        var mathParagraphs = FindMathParagraphs(document);

        var issues = new List<BackMatterIssue>();
        var notChecked = new List<string>(reading.Undetermined.Where(u => u.StartsWith("word/document.xml", StringComparison.Ordinal)));
        var spans = new List<BackMatterSectionSpan>();

        // 标题位置按真实 w:p 序号记录；PartFormatReader 跳过无文字段，但保留原序号。
        var titles = new List<(int Index, SectionSpec Spec)>();
        for (var i = 0; i < readings.Count; i++)
        {
            foreach (var spec in Sections)
            {
                if (!spec.TitlePattern.IsMatch(readings[i].Text.Trim())) continue;
                if (titles.Any(t => ReferenceEquals(t.Spec, spec))) break; // 一个部件只认第一处标题
                titles.Add((i, spec));
                break;
            }
        }

        foreach (var spec in Sections)
        {
            if (!titles.Any(t => ReferenceEquals(t.Spec, spec)))
            {
                missingParts.Add(spec.Name);
                notChecked.Add($"{spec.Name}（{spec.RuleId}）：未识别到标题段落，正文格式未检查");
            }
        }

        // 部件区间：本部件标题之后，到下一个「不同部件」标题之前；中间的同部件子标题（附录 A/B）不算正文。
        for (var t = 0; t < titles.Count; t++)
        {
            var (titleIndex, spec) = titles[t];
            var nextTitleIndex = t + 1 < titles.Count ? titles[t + 1].Index : readings.Count;
            var bodyIndices = new List<int>();
            for (var i = titleIndex + 1; i < nextTitleIndex; i++)
            {
                var text = readings[i].Text.Trim();
                // 同部件子标题（附录 A/B）与其它部件标题都不算正文段落，避免拿标题格式误报正文问题。
                if (Sections.Any(candidate => candidate.TitlePattern.IsMatch(text))) continue;
                if (string.IsNullOrWhiteSpace(text)) continue;
                bodyIndices.Add(i);
            }

            spans.Add(new BackMatterSectionSpan(spec.Name, spec.RuleId,
                readings[titleIndex].ParagraphIndex,
                bodyIndices.Select(i => readings[i].ParagraphIndex).ToArray()));

            var mathExempted = false;
            foreach (var index in bodyIndices)
            {
                var r = readings[index];
                CheckBodyParagraph(r, spec, mathParagraphs.Contains(r.ParagraphIndex),
                    ref mathExempted, issues, notChecked);
            }

            if (bodyIndices.Count == 0)
                notChecked.Add($"{spec.Name}（{spec.RuleId}）：标题后未识别到正文段落，正文格式未检查");
        }

        notChecked.Add("各部件标题要求同各章标题（R7），本命令只判正文段落，标题格式未判定");
        if (styles is null)
            notChecked.Add("缺少 word/styles.xml：只按 run 上的显式 rPr / pPr 判定，样式链不可用");

        return new BackMatterFormatResult(
            Issues: issues,
            Sections: spans,
            NotChecked: notChecked.Distinct(StringComparer.Ordinal).ToArray(),
            MissingParts: missingParts,
            CheckedParagraphs: spans.Sum(s => s.BodyParagraphs.Count),
            RuleVersion: RuleVersion,
            SourceNote: SourceNote);
    }

    private static void CheckBodyParagraph(PartFormatReading r, SectionSpec spec, bool hasMath,
        ref bool mathExempted, List<BackMatterIssue> issues, List<string> notChecked)
    {
        var paragraphIndex = r.ParagraphIndex;
        var label = $"{spec.Name}（{spec.RuleId}）第 {paragraphIndex} 段";

        // 中文字体只对含中文的段落有意义；纯西文段没有 eastAsia 要求，不算漏判。
        if (!r.WesternOnly && spec.CnFont is not null && spec.CnFontMatch is not null)
        {
            if (string.IsNullOrWhiteSpace(r.EastAsiaFont))
            {
                notChecked.Add($"{label}中文字体：样式链未解析出 eastAsia 字体槽，未判定");
            }
            else if (!spec.CnFontMatch(r.EastAsiaFont))
            {
                Add(issues, Code(spec.CodeStart, spec.DimensionOffsets[1]), $"{spec.Name}正文中文字体", spec.Name, paragraphIndex,
                    spec.CnFont, r.EastAsiaFont,
                    $"{label}中文应为「{spec.CnFont}」，实际为「{r.EastAsiaFont}」。依据 {SourceNote}。");
            }
        }

        if (spec.WesternFont is not null)
        {
            if (string.IsNullOrWhiteSpace(r.WesternFont))
            {
                notChecked.Add($"{label}西文字体：样式链未解析出 ascii/hAnsi 字体槽，未判定");
            }
            else if (!PartFormatReader.IsTimesNewRoman(r.WesternFont))
            {
                Add(issues, Code(spec.CodeStart, spec.DimensionOffsets[2]), $"{spec.Name}正文西文字体", spec.Name, paragraphIndex,
                    spec.WesternFont, r.WesternFont,
                    $"{label}西文应为 Times New Roman，实际为「{r.WesternFont}」。依据 {SourceNote}。");
            }
        }

        if (r.SizeHalfPoints is null)
        {
            notChecked.Add($"{label}字号：样式链未解析出字号，未判定");
        }
        else if (r.SizeHalfPoints.Value != spec.SizeHalfPoints)
        {
            Add(issues, Code(spec.CodeStart, spec.DimensionOffsets[3]), $"{spec.Name}正文字号", spec.Name, paragraphIndex,
                $"{spec.SizeHalfPoints / 2.0:0.#}磅", $"{r.SizeHalfPoints.Value / 2.0:0.#}磅",
                $"{label}应为 {spec.SizeHalfPoints / 2.0:0.#} 磅，实际为 {r.SizeHalfPoints.Value / 2.0:0.#} 磅。依据 {SourceNote}。");
        }

        if (spec.Alignment is not null)
        {
            if (!r.AlignmentResolved)
            {
                notChecked.Add($"{label}对齐：样式链未解析出对齐，未判定");
            }
            else if (!string.Equals(r.Alignment, spec.Alignment, StringComparison.OrdinalIgnoreCase))
            {
                Add(issues, Code(spec.CodeStart, spec.DimensionOffsets[4]), $"{spec.Name}正文对齐", spec.Name, paragraphIndex,
                    spec.Alignment, r.Alignment,
                    $"{label}应两端对齐，实际为「{r.Alignment}」。依据 {SourceNote}。");
            }
        }

        if (spec.RequireTwoCharFirstLine)
        {
            if (!HasTwoCharFirstLine(r))
            {
                // 完全读不到缩进才算未判定；读到但不是 2 字就是明确的不符合。
                if (r.FirstLineChars is null && r.FirstLineTwips is null)
                {
                    notChecked.Add($"{label}首行缩进：样式链未解析出 w:ind，未判定");
                }
                else
                {
                    Add(issues, Code(spec.CodeStart, spec.DimensionOffsets[5]), $"{spec.Name}正文首行缩进", spec.Name, paragraphIndex,
                        "2 个汉字符", DescribeIndent(r),
                        $"{label}应首行缩进 2 个汉字符，实际为「{DescribeIndent(r)}」。依据 {SourceNote}。");
                }
            }
        }

        // R15 允许含数学表达式的段落按表达需要调整行距；只豁免行距，段前后仍要查。
        var spacingConcerned = !(hasMath && spec.ExemptMathLineSpacing);
        if (hasMath && spec.ExemptMathLineSpacing && !mathExempted)
        {
            notChecked.Add($"{spec.Name}（{spec.RuleId}）：含数学表达式的段落行距按表达需要调整，未按固定 {spec.ExactLineTwips / 20} 磅判定");
            mathExempted = true;
        }

        if (spacingConcerned)
        {
            CheckSpacing(r, spec, paragraphIndex, issues, notChecked);
        }
        else if (r.SpacingResolved && (r.SpaceBeforeTwips != ZeroTwips || r.SpaceAfterTwips != ZeroTwips))
        {
            Add(issues, Code(spec.CodeStart, spec.DimensionOffsets[6]), $"{spec.Name}正文段距", spec.Name, paragraphIndex,
                "段前0磅/段后0磅", $"段前{r.SpaceBeforeTwips / 20}磅/段后{r.SpaceAfterTwips / 20}磅",
                $"{label}即使含公式仍应段前后 0 磅。依据 {SourceNote}。");
        }
    }

    private static void CheckSpacing(PartFormatReading r, SectionSpec spec, int paragraphIndex,
        List<BackMatterIssue> issues, List<string> notChecked)
    {
        if (!r.SpacingResolved)
        {
            notChecked.Add($"{spec.Name}（{spec.RuleId}）第 {paragraphIndex} 段行距/段距：样式链未解析出完整 spacing，未判定");
            return;
        }

        var ruleOk = string.Equals(r.LineSpacingRule, "exact", StringComparison.OrdinalIgnoreCase);
        var lineOk = r.LineSpacingValue == spec.ExactLineTwips;
        var beforeOk = r.SpaceBeforeTwips == ZeroTwips;
        var afterOk = r.SpaceAfterTwips == ZeroTwips;
        if (ruleOk && lineOk && beforeOk && afterOk) return;

        Add(issues, Code(spec.CodeStart, spec.DimensionOffsets[6]), $"{spec.Name}正文行距与段距", spec.Name, paragraphIndex,
            $"固定值{spec.ExactLineTwips / 20}磅，段前0磅/段后0磅",
            $"{r.LineSpacingRule}/{r.LineSpacingValue}，段前{r.SpaceBeforeTwips / 20}磅/段后{r.SpaceAfterTwips / 20}磅",
            $"{spec.Name}（{spec.RuleId}）第 {paragraphIndex} 段应为固定值 {spec.ExactLineTwips / 20} 磅、段前后 0 磅，实际为 {r.LineSpacingRule}/{r.LineSpacingValue}、段前{r.SpaceBeforeTwips / 20}磅/段后{r.SpaceAfterTwips / 20}磅。依据 {SourceNote}。");
    }

    private static bool HasTwoCharFirstLine(PartFormatReading r)
    {
        if (r.FirstLineChars is int chars) return chars == TwoCharFirstLine;
        // 没有 firstLineChars 时按字号换算：2 字 × (半点/2)磅 × 20 twips = 半点 × 20。
        if (r.FirstLineTwips is int twips && r.SizeHalfPoints is int size) return twips == size * 20;
        return false;
    }

    private static string DescribeIndent(PartFormatReading r)
    {
        if (r.FirstLineChars is int chars) return $"{chars / 100.0:0.#} 汉字符";
        if (r.FirstLineTwips is int twips) return $"{twips / 20.0:0.#} 磅";
        return "未解析出";
    }

    /// <summary>仿宋在中文版 Word 里存成「仿宋」或内部名 FangSong。</summary>
    private static bool IsFangSong(string? font)
    {
        if (font is null) return false;
        var normalized = font.Replace(" ", string.Empty).ToLowerInvariant();
        return normalized is "仿宋" or "fangsong" || normalized == "仿宋_gb2312";
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

    private static string Code(int codeStart, int offset) => "B" + (codeStart + offset);

    private static void Add(List<BackMatterIssue> issues, string code, string title, string section,
        int paragraphIndex, string expected, string actual, string message)
        => issues.Add(new BackMatterIssue(code, title, section, paragraphIndex, expected, actual, message));

    private static XDocument? ReadPart(ZipArchive archive, string entry)
    {
        var item = archive.GetEntry(entry);
        if (item is null) return null;
        using var stream = item.Open();
        return XDocument.Load(stream);
    }
}
