using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Core;

/// <summary>一条标题格式问题。ParagraphIndex 为 document.xml 里 w:p 的 1 基序号。</summary>
public sealed record HeadingIssue(
    string Code,
    string RuleId,
    string Level,
    int ParagraphIndex,
    string HeadingText,
    string Dimension,
    string Expected,
    string Actual,
    string Message);

/// <summary>一个被识别到的分级标题（信息用途：编号模式只展示，不参与合规判定）。</summary>
public sealed record HeadingInfo(
    int ParagraphIndex,
    int OutlineLevel,
    string Level,
    string Text,
    string NumberingMode);

/// <summary>标题层级格式分析结果。NotChecked 单列「压根没查 / 条文冲突无法判定」。</summary>
public sealed record HeadingFormatResult(
    IReadOnlyList<HeadingIssue> Issues,
    IReadOnlyList<HeadingInfo> RecognizedHeadings,
    int ChapterStartParagraph,
    int SectionEndParagraph,
    IReadOnlyList<string> NotChecked,
    IReadOnlyList<string> MissingParts,
    int CheckedHeadings,
    string RuleVersion,
    string SourceNote);

/// <summary>
/// 标题层级格式检查（R7–R10，依据 general.yaml 4.7 表4.6，来源 S4 2025 版 / 2026 版 S1），只查不改。
///
/// 口径边界（与 docs/rule-notes/headings.md、nku-2026-coverage-matrix.md M26-035 一致）：
/// - 标题级别只看 w:outlineLvl 或 HeadingN /「标题 N」样式（沿 basedOn 链），不靠字体特征硬猜。
///   识别不到的文档如实列未检查，不拿正文或封面冒充标题。
/// - 检查范围：识别到的第一个章标题（大纲 0 级且文本像章标题）之后，到参考文献/附录/致谢等
///   结束标题之前。章/一级/二级/三级分别对应 R7/R8/R9/R10。
/// - 只判条文给了明确值且无版本冲突的维度：字体（黑体）、字号（16/14/13/12 磅）、
///   章与一级加粗、单倍行距、段前段后。二级/三级原文未写加粗，不据此判。
/// - 对齐依编号模式（模式一居左、模式二居中），且 S4 与 S5/S7 表述冲突（headings.md T2/C3）；
///   「序号与题名间空一个汉字」宽度语义未定义（RFC-3）。这两类列未判定，不伪造硬规则。
/// - 样式链读不到的维度列未判定，不拿默认值凑问题数。规则状态 draft，需人工确认。
/// - 问题码用每级 CodeStart + 维度偏移映射，H1–H22 连续，避免「前缀 + 序号」拼出 H111。
/// </summary>
public sealed class HeadingFormatAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string SourceNote =
        "南开大学研究生学位论文写作规范 4.7 表4.6 标题 R7–R10（general.yaml，来源 S4 2025 版 / S1 2026 版，规则状态 draft，对齐与编号模式待负责人确认）";
    private const string RuleVersion = "nku-graduate-headings-r7r10-draft";

    // 单倍行距 = lineRule auto + line 240（1 行 = 240 twips）。
    private const string AutoLineRule = "auto";
    private const int SingleLineValue = 240;

    // 章标题文本形态（模式一「1 绪论」/ 模式二「第一章 绪论」/ 英文 Chapter N）。
    private static readonly Regex ChapterPattern = new(
        @"^(?:第[一二三四五六七八九十百零〇0-9]+章|Chapter\s+(?:[0-9]+|[IVXLCDM]+)\b|[0-9]+(?:\s|$))",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex EndPattern = new(
        @"^(?:主要)?参考文献$|^(?:References|Bibliography)$|^附\s*录\b|^Appendix(?:es)?\b|^致\s*谢$|^Acknowledg(?:e)?ments$"
        + @"|^攻读学位期间|^在学期间.{0,20}(?:成果|论文)|^(?:个人)?简历$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex HeadingStyleName = new(
        @"^(?:heading|标题)\s*([1-9])$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    // 每级规范值与问题码映射。Font/Size/Bold/Line/Before/After 为维度偏移（相对 CodeStart）。
    // 章/一级含 Bold（6 维），二级/三级无加粗要求（5 维）；CodeStart 使 H1–H22 连续。
    private sealed record LevelSpec(
        int OutlineLevel,
        string Level,
        string RuleId,
        int CodeStart,
        int SizeHalfPoints,
        int SizePt,
        int BeforeTwips,
        int AfterTwips,
        bool CheckBold,
        IReadOnlyDictionary<string, int> Offsets);

    private static readonly IReadOnlyDictionary<string, int> SixDims = new Dictionary<string, int>
    {
        ["Font"] = 1, ["Size"] = 2, ["Bold"] = 3, ["Line"] = 4, ["Before"] = 5, ["After"] = 6,
    };
    private static readonly IReadOnlyDictionary<string, int> FiveDims = new Dictionary<string, int>
    {
        ["Font"] = 1, ["Size"] = 2, ["Line"] = 3, ["Before"] = 4, ["After"] = 5,
    };

    private static readonly LevelSpec[] Levels =
    {
        new(0, "章标题", "R7", 0, 32, 16, 480, 360, true, SixDims),
        new(1, "一级标题", "R8", 6, 28, 14, 480, 120, true, SixDims),
        new(2, "二级标题", "R9", 12, 26, 13, 240, 120, false, FiveDims),
        new(3, "三级标题", "R10", 17, 24, 12, 240, 120, false, FiveDims),
    };

    public HeadingFormatResult Analyze(string docxPath)
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
        var formats = new Dictionary<int, PartFormatReading>();
        foreach (var r in reading.Readings) formats[r.ParagraphIndex] = r;

        var styleIndex = BuildStyleIndex(styles);
        var rawHeadings = new List<(int Index, int Level, string Text)>();
        var deeperLevelCount = 0;
        var paragraphIndex = 0;
        foreach (var paragraph in document.Root?.Descendants(W + "p") ?? Enumerable.Empty<XElement>())
        {
            paragraphIndex++;
            var pPr = paragraph.Element(W + "pPr");
            var text = string.Concat(paragraph.Descendants(W + "t").Select(node => node.Value)).Trim();
            var styleId = pPr?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;
            var level = ResolveOutlineLevel(pPr, styleId, styleIndex);
            if (level is null) continue;
            if (level > 3) { deeperLevelCount++; continue; }
            if (level < 0) continue;
            rawHeadings.Add((paragraphIndex, level.Value, text));
        }

        var issues = new List<HeadingIssue>();
        var notChecked = new List<string>(reading.Undetermined.Where(u =>
            u.StartsWith("word/document.xml", StringComparison.Ordinal)));

        // 正文边界：第一个「章样式 + 像章标题」的标题，到其后第一个结束标题之前。
        var chapterPos = -1;
        for (var i = 0; i < rawHeadings.Count; i++)
        {
            if (rawHeadings[i].Level == 0 && ChapterPattern.IsMatch(rawHeadings[i].Text))
            { chapterPos = i; break; }
        }

        var recognized = new List<HeadingInfo>();
        var checkedHeadings = 0;
        var chapterParagraph = 0;
        var sectionEndParagraph = 0;

        if (chapterPos < 0)
        {
            notChecked.Add("未识别到章标题（大纲 0 级且文本像章标题），R7–R10 标题层级检查未执行");
        }
        else
        {
            chapterParagraph = rawHeadings[chapterPos].Index;
            var endParagraph = int.MaxValue;
            for (var i = chapterPos + 1; i < rawHeadings.Count; i++)
            {
                if (EndPattern.IsMatch(rawHeadings[i].Text)) { endParagraph = rawHeadings[i].Index; break; }
            }
            if (endParagraph != int.MaxValue) sectionEndParagraph = endParagraph;

            for (var i = chapterPos; i < rawHeadings.Count; i++)
            {
                var (index, level, text) = rawHeadings[i];
                if (index >= endParagraph) break;
                var spec = Levels.First(s => s.OutlineLevel == level);
                recognized.Add(new HeadingInfo(index, level, spec.Level, text, NumberingModeOf(text)));
                if (!formats.TryGetValue(index, out var r))
                {
                    notChecked.Add($"{spec.Level}第 {index} 段「{Short(text)}」：样式链解析跳过了该段，格式未判定");
                    continue;
                }
                checkedHeadings++;
                CheckHeading(spec, r, issues, notChecked);
            }
        }

        if (deeperLevelCount > 0)
            notChecked.Add($"另有 {deeperLevelCount} 个大纲 4 级及以上的标题，R7–R10 只覆盖章至三级，未检查");
        if (styles is null)
            notChecked.Add("缺少 word/styles.xml：大纲级别只能按段落直接 w:outlineLvl 识别，样式链不可用");

        // 条文冲突 / 未定义维度：统一列未判定，不伪造硬规则。
        notChecked.Add("对齐：模式一居左、模式二居中，但 S4 与 S5/S7 表述冲突（headings.md T2/C3），本次不据以判定");
        notChecked.Add("序号与题名间空一个汉字：宽度语义未定义（headings.md RFC-3），未按字符计数判定");
        notChecked.Add("二级/三级标题加粗：原文未写加粗（general.yaml 明确不得默认为加粗），不据此判定");

        return new HeadingFormatResult(
            Issues: issues,
            RecognizedHeadings: recognized,
            ChapterStartParagraph: chapterParagraph,
            SectionEndParagraph: sectionEndParagraph,
            NotChecked: notChecked.Distinct(StringComparer.Ordinal).ToArray(),
            MissingParts: missingParts,
            CheckedHeadings: checkedHeadings,
            RuleVersion: RuleVersion,
            SourceNote: SourceNote);
    }

    private static void CheckHeading(LevelSpec spec, PartFormatReading r,
        List<HeadingIssue> issues, List<string> notChecked)
    {
        var label = $"{spec.Level}第 {r.ParagraphIndex} 段「{Short(r.Text)}」";

        void Add(string dim, string expected, string actual)
        {
            var code = $"H{spec.CodeStart + spec.Offsets[dim]}";
            issues.Add(new HeadingIssue(
                code, spec.RuleId, spec.Level, r.ParagraphIndex, r.Text, dim, expected, actual,
                $"{label}：{dim}期望「{expected}」，实际「{actual}」。依据 {SourceNote}。"));
        }

        // 字体：黑体（仅 eastAsia；中英混排的西文字体槽位未规定，不判）。
        if (!r.FontResolved)
            notChecked.Add($"{label}字体：样式链未解析出字体，未判定");
        else if (!PartFormatReader.IsSimHei(r.EastAsiaFont))
            Add("Font", "黑体", r.EastAsiaFont ?? "未解析出");

        // 字号：章 16 / 一级 14 / 二级 13 / 三级 12 磅。
        if (!r.SizeResolved)
            notChecked.Add($"{label}字号：样式链未解析出字号，未判定");
        else if (r.SizeHalfPoints != spec.SizeHalfPoints)
            Add("Size", $"{spec.SizePt}磅", $"{r.SizeHalfPoints / 2.0:0.#}磅");

        // 加粗：仅章/一级要求加粗。
        if (spec.CheckBold)
        {
            if (!r.BoldResolved)
                notChecked.Add($"{label}加粗：样式链未解析出加粗开关，未判定");
            else if (r.Bold != true)
                Add("Bold", "加粗", "未加粗");
        }

        // 单倍行距：lineRule auto + line 240。读不到列未判定，不拿默认凑数。
        if (r.LineSpacingRule is null || r.LineSpacingValue is null)
        {
            notChecked.Add($"{label}行距：样式链未解析出行距，未判定");
        }
        else if (!(string.Equals(r.LineSpacingRule, AutoLineRule, StringComparison.OrdinalIgnoreCase)
                   && r.LineSpacingValue == SingleLineValue))
        {
            Add("Line", "单倍行距", $"{r.LineSpacingRule}/{r.LineSpacingValue}");
        }

        // 段前段后。
        if (r.SpaceBeforeTwips is int before && before != spec.BeforeTwips)
            Add("Before", $"{spec.BeforeTwips / 20}磅", $"{before / 20.0:0.#}磅");
        else if (r.SpaceBeforeTwips is null)
            notChecked.Add($"{label}段前：样式链未解析出段前，未判定");

        if (r.SpaceAfterTwips is int after && after != spec.AfterTwips)
            Add("After", $"{spec.AfterTwips / 20}磅", $"{after / 20.0:0.#}磅");
        else if (r.SpaceAfterTwips is null)
            notChecked.Add($"{label}段后：样式链未解析出段后，未判定");
    }

    private static int? ResolveOutlineLevel(
        XElement? pPr,
        string? styleId,
        IReadOnlyDictionary<string, (string? Name, int? Outline, string? BasedOn)> styles)
    {
        var direct = pPr?.Element(W + "outlineLvl")?.Attribute(W + "val")?.Value;
        if (int.TryParse(direct, out var d)) return d;

        var visited = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        while (current is not null && visited.Add(current) && styles.TryGetValue(current, out var style))
        {
            if (style.Outline is int outline) return outline;
            if (style.Name is string name)
            {
                var match = HeadingStyleName.Match(name);
                if (match.Success && int.TryParse(match.Groups[1].Value, out var n)) return n - 1;
            }
            current = style.BasedOn;
        }
        return null;
    }

    private static IReadOnlyDictionary<string, (string? Name, int? Outline, string? BasedOn)> BuildStyleIndex(
        XDocument? styles)
    {
        var result = new Dictionary<string, (string?, int?, string?)>(StringComparer.Ordinal);
        if (styles?.Root is null) return result;
        foreach (var style in styles.Root.Elements(W + "style"))
        {
            var id = style.Attribute(W + "styleId")?.Value;
            if (id is null) continue;
            var name = style.Element(W + "name")?.Attribute(W + "val")?.Value;
            var outlineText = style.Element(W + "pPr")?.Element(W + "outlineLvl")?.Attribute(W + "val")?.Value;
            var outline = int.TryParse(outlineText, out var o) ? o : (int?)null;
            var basedOn = style.Element(W + "basedOn")?.Attribute(W + "val")?.Value;
            result[id] = (name, outline, basedOn);
        }
        return result;
    }

    private static string NumberingModeOf(string text)
    {
        if (Regex.IsMatch(text, @"^[0-9]+(?:\.[0-9]+){0,3}(?:\s|$)")) return "模式一";
        if (Regex.IsMatch(text,
                @"^(?:第[一二三四五六七八九十百零〇0-9]+[章节]|[一二三四五六七八九十]+、|[（(][一二三四五六七八九十0-9]+[）)])"))
            return "模式二";
        return "未识别";
    }

    private static string Short(string text) => text.Length <= 24 ? text : text[..24] + "…";

    private static XDocument? ReadPart(ZipArchive archive, string entry)
    {
        var item = archive.GetEntry(entry);
        if (item is null) return null;
        using var stream = item.Open();
        return XDocument.Load(stream);
    }
}
