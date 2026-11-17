using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Core;

/// <summary>一条目录格式问题。ParagraphIndex 为 document.xml 里 w:p 的 1 基序号。</summary>
public sealed record TocIssue(
    string Code,
    string RuleId,
    string Part,
    int ParagraphIndex,
    string Text,
    string Dimension,
    string Expected,
    string Actual,
    string Message);

/// <summary>一个被识别到的目录条目（信息用途）。</summary>
public sealed record TocEntry(
    int ParagraphIndex,
    string Level,
    string Text);

/// <summary>目录格式分析结果。NotChecked 单列「压根没查 / 检测对象未决」。</summary>
public sealed record TocFormatResult(
    IReadOnlyList<TocIssue> Issues,
    IReadOnlyList<TocEntry> Entries,
    int TocTitleParagraph,
    IReadOnlyList<string> NotChecked,
    IReadOnlyList<string> MissingParts,
    int CheckedParagraphs,
    string RuleVersion,
    string SourceNote);

/// <summary>
/// 目录格式检查（R6，依据 general.yaml 4.6 表4.5，来源 S1 2025 版），只查不改。
///
/// 口径边界（与 general.yaml toc.open_question=T6、RFC-7 一致）：
/// - 条目级别只按 TOC1/TOC2/TOC3 样式（styleId 或样式名，大小写/中英文不敏感）识别，
///   分别对应章/一级/二级目录条目；不靠文本缩进硬猜。目录标题用文本「目录/Contents」作锚点。
/// - 检测对象是已展开的目录段落（域结果或静态文本），不是 TOC 域指令本身：若文档只有 TOC 域
///   指令、域结果未展开（无 TOC 样式条目），如实列未判定，不假装检查过（T6 未决点）。
/// - 只判条文给了明确值且无渲染依赖的维度：字体（宋体）、字号（14/12/10.5 磅）、单倍行距、
///   段前后（6/0 磅）、左缩进（0/1/2 汉字）；目录标题判黑体 16 磅加粗居中、单倍、段前后 24/18。
/// - 条目「两端对齐 + 页码右对齐」走制表位，依赖渲染结果（T6/RFC-7），列未判定，不据段落 jc 硬判。
/// - 判不出的维度列未判定，不拿默认值凑问题数。规则状态 draft，需人工确认。
/// - 问题码用每级 CodeStart + 维度偏移映射，C1–C25 连续。
/// </summary>
public sealed class TocFormatAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string SourceNote =
        "南开大学研究生学位论文写作规范 4.6 表4.5 目录 R6（general.yaml，来源 S1 2025 版，规则状态 draft，检测对象 T6/RFC-7 待负责人确认）";
    private const string RuleVersion = "nku-graduate-toc-r6-draft";

    private const string AutoLineRule = "auto";
    private const int SingleLineValue = 240;

    private static readonly Regex TocTitlePattern = new(
        @"^(?:目\s*录|Table\s+of\s+Contents|Contents)$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex TocStylePattern = new(
        @"^toc\s*([123])$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    private static readonly Regex TocFieldPattern = new(
        @"\bTOC\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    // 每级规范值与问题码映射。缩进用 LeftChars（1 汉字=100），无 leftChars 时按字号换算 left twips。
    private sealed record TocSpec(
        string Level,
        string RuleId,
        int CodeStart,
        int SizeHalfPoints,
        string SizeLabel,
        int BeforeTwips,
        int AfterTwips,
        int IndentChars,
        IReadOnlyDictionary<string, int> Offsets);

    private static readonly IReadOnlyDictionary<string, int> TitleDims = new Dictionary<string, int>
    {
        ["Font"] = 1, ["Size"] = 2, ["Bold"] = 3, ["Align"] = 4,
        ["Line"] = 5, ["Before"] = 6, ["After"] = 7,
    };
    private static readonly IReadOnlyDictionary<string, int> EntryDims = new Dictionary<string, int>
    {
        ["Font"] = 1, ["Size"] = 2, ["Line"] = 3, ["Before"] = 4, ["After"] = 5, ["Indent"] = 6,
    };

    private static readonly TocSpec TitleSpec = new(
        "目录标题", "R6", 0, 32, "16磅", 480, 360, 0, TitleDims);
    private static readonly TocSpec[] EntrySpecs =
    {
        new("章目录条目", "R6", 7, 28, "14磅", 120, 0, 0, EntryDims),
        new("一级目录条目", "R6", 13, 24, "12磅", 120, 0, 1, EntryDims),
        new("二级目录条目", "R6", 19, 21, "10.5磅", 120, 0, 2, EntryDims),
    };

    public TocFormatResult Analyze(string docxPath)
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

        var styleNames = BuildStyleNames(styles);
        var rawEntries = new List<(int Index, int Level, string Text)>();
        var tocTitleParagraph = 0;
        var hasTocField = false;
        var paragraphIndex = 0;

        foreach (var paragraph in document.Root?.Descendants(W + "p") ?? Enumerable.Empty<XElement>())
        {
            paragraphIndex++;
            var pPr = paragraph.Element(W + "pPr");
            var text = string.Concat(paragraph.Descendants(W + "t").Select(node => node.Value)).Trim();
            var styleId = pPr?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;

            // TOC 域指令：fldChar/instrText/fldSimple 里出现 TOC。
            if (!hasTocField &&
                (paragraph.Descendants(W + "instrText").Any(n => TocFieldPattern.IsMatch(n.Value ?? ""))
                 || paragraph.Descendants(W + "fldSimple").Any(n => TocFieldPattern.IsMatch(n.Attribute(W + "instr")?.Value ?? ""))))
                hasTocField = true;

            if (tocTitleParagraph == 0 && TocTitlePattern.IsMatch(text)) tocTitleParagraph = paragraphIndex;

            var level = TocStyleLevel(styleId, styleNames);
            if (level is null) continue;
            // TOC 域未展开时，段落带 TOC 样式但域结果（文字）为空，不是可检查的条目。
            if (string.IsNullOrWhiteSpace(text)) continue;
            rawEntries.Add((paragraphIndex, level.Value, text));
        }

        var issues = new List<TocIssue>();
        var notChecked = new List<string>(reading.Undetermined.Where(u =>
            u.StartsWith("word/document.xml", StringComparison.Ordinal)));
        var recognized = new List<TocEntry>();
        var checkedParagraphs = 0;

        // 目录标题：文本锚点定位，检查黑体 16 磅加粗居中、单倍、段前后 24/18。
        if (tocTitleParagraph > 0)
        {
            if (formats.TryGetValue(tocTitleParagraph, out var titleReading))
            {
                checkedParagraphs++;
                CheckTitle(titleReading, issues, notChecked);
            }
            else
            {
                notChecked.Add($"目录标题第 {tocTitleParagraph} 段「目录」：样式链解析跳过，格式未判定");
            }
        }

        // 目录条目：按 TOC 样式分级检查。无目录标题但有条目时仍检查条目，标题列未判定。
        if (rawEntries.Count == 0)
        {
            if (hasTocField)
                notChecked.Add("检测到 TOC 域指令但域结果未展开（无 TOC1/TOC2/TOC3 样式条目），条目格式未检查（T6 检测对象未决，需 Word 更新域后复核）");
            else
                notChecked.Add("未识别到 TOC1/TOC2/TOC3 样式的目录条目，条目格式未检查（不靠缩进/文本硬猜）");
        }
        else if (tocTitleParagraph == 0)
        {
            notChecked.Add("识别到目录条目但未识别到「目录」标题，目录标题格式未判定");
        }

        foreach (var (index, level, text) in rawEntries)
        {
            var spec = EntrySpecs[level - 1];
            recognized.Add(new TocEntry(index, spec.Level, text));
            if (!formats.TryGetValue(index, out var r))
            {
                notChecked.Add($"{spec.Level}第 {index} 段「{Short(text)}」：样式链解析跳过，格式未判定");
                continue;
            }
            checkedParagraphs++;
            CheckEntry(spec, r, issues, notChecked);
        }

        notChecked.Add("目录条目的两端对齐与页码右对齐走制表位、依赖渲染结果（T6/RFC-7 未决），本次不据段落 jc 判定");
        if (styles is null)
            notChecked.Add("缺少 word/styles.xml：TOC 样式名无法解析，条目无法分级，样式链不可用");

        return new TocFormatResult(
            Issues: issues,
            Entries: recognized,
            TocTitleParagraph: tocTitleParagraph,
            NotChecked: notChecked.Distinct(StringComparer.Ordinal).ToArray(),
            MissingParts: missingParts,
            CheckedParagraphs: checkedParagraphs,
            RuleVersion: RuleVersion,
            SourceNote: SourceNote);
    }

    private static void CheckTitle(PartFormatReading r, List<TocIssue> issues, List<string> notChecked)
    {
        var label = $"目录标题第 {r.ParagraphIndex} 段「目录」";

        void Add(string dim, string expected, string actual)
        {
            issues.Add(new TocIssue(
                $"C{TitleSpec.CodeStart + TitleDims[dim]}", TitleSpec.RuleId, TitleSpec.Level,
                r.ParagraphIndex, r.Text, dim, expected, actual,
                $"{label}：{dim}期望「{expected}」，实际「{actual}」。依据 {SourceNote}。"));
        }

        if (!r.FontResolved) notChecked.Add($"{label}字体：未判定");
        else if (!PartFormatReader.IsSimHei(r.EastAsiaFont)) Add("Font", "黑体", r.EastAsiaFont ?? "未解析出");

        if (!r.SizeResolved) notChecked.Add($"{label}字号：未判定");
        else if (r.SizeHalfPoints != TitleSpec.SizeHalfPoints) Add("Size", "16磅", $"{r.SizeHalfPoints / 2.0:0.#}磅");

        if (!r.BoldResolved) notChecked.Add($"{label}加粗：未判定");
        else if (r.Bold != true) Add("Bold", "加粗", "未加粗");

        if (!r.AlignmentResolved) notChecked.Add($"{label}对齐：未判定");
        else if (!string.Equals(r.Alignment, "center", StringComparison.OrdinalIgnoreCase)) Add("Align", "居中", r.Alignment);

        CheckLineAndSpacing(r, TitleSpec, label, Add, notChecked, checkIndent: false);
    }

    private static void CheckEntry(TocSpec spec, PartFormatReading r, List<TocIssue> issues, List<string> notChecked)
    {
        var label = $"{spec.Level}第 {r.ParagraphIndex} 段「{Short(r.Text)}」";

        void Add(string dim, string expected, string actual)
        {
            issues.Add(new TocIssue(
                $"C{spec.CodeStart + EntryDims[dim]}", spec.RuleId, spec.Level,
                r.ParagraphIndex, r.Text, dim, expected, actual,
                $"{label}：{dim}期望「{expected}」，实际「{actual}」。依据 {SourceNote}。"));
        }

        if (!r.FontResolved) notChecked.Add($"{label}字体：未判定");
        else if (!PartFormatReader.IsSimSun(r.EastAsiaFont)) Add("Font", "宋体", r.EastAsiaFont ?? "未解析出");

        if (!r.SizeResolved) notChecked.Add($"{label}字号：未判定");
        else if (r.SizeHalfPoints != spec.SizeHalfPoints) Add("Size", spec.SizeLabel, $"{r.SizeHalfPoints / 2.0:0.#}磅");

        CheckLineAndSpacing(r, spec, label, Add, notChecked, checkIndent: true);
    }

    private static void CheckLineAndSpacing(
        PartFormatReading r, TocSpec spec, string label,
        Action<string, string, string> add, List<string> notChecked, bool checkIndent)
    {
        if (r.LineSpacingRule is null || r.LineSpacingValue is null)
            notChecked.Add($"{label}行距：未判定");
        else if (!(string.Equals(r.LineSpacingRule, AutoLineRule, StringComparison.OrdinalIgnoreCase)
                   && r.LineSpacingValue == SingleLineValue))
            add("Line", "单倍行距", $"{r.LineSpacingRule}/{r.LineSpacingValue}");

        if (r.SpaceBeforeTwips is int before && before != spec.BeforeTwips)
            add("Before", $"{spec.BeforeTwips / 20}磅", $"{before / 20.0:0.#}磅");
        else if (r.SpaceBeforeTwips is null)
            notChecked.Add($"{label}段前：未判定");

        if (r.SpaceAfterTwips is int after && after != spec.AfterTwips)
            add("After", $"{spec.AfterTwips / 20}磅", $"{after / 20.0:0.#}磅");
        else if (r.SpaceAfterTwips is null)
            notChecked.Add($"{label}段后：未判定");

        if (!checkIndent) return;

        var (ok, undetermined) = IndentOk(r, spec.IndentChars);
        if (undetermined)
            notChecked.Add($"{label}左缩进：样式链未解析出 leftChars/left，未判定");
        else if (!ok)
            add("Indent", $"左缩进{spec.IndentChars}汉字", DescribeIndent(r));
    }

    /// <summary>1 汉字 ≈ 字号磅数；leftChars 以 1/100 字符计（100=1 汉字）；left twips ≈ 字号磅数×20/汉字。</summary>
    private static (bool Ok, bool Undetermined) IndentOk(PartFormatReading r, int expectedChars)
    {
        if (r.LeftChars is int chars) return (chars == expectedChars * 100, false);
        if (r.LeftTwips is int twips)
        {
            if (expectedChars == 0) return (twips == 0, false);
            if (r.SizeHalfPoints is int size && size > 0)
                return (Math.Abs(twips - expectedChars * size * 10) <= 20, false);
            return (false, true);
        }
        // 章条目要求 0 缩进：无 w:ind 即合规；其余级别缺值判不出。
        return expectedChars == 0 ? (true, false) : (false, true);
    }

    private static string DescribeIndent(PartFormatReading r)
    {
        if (r.LeftChars is int chars) return $"{chars / 100.0:0.#}汉字";
        if (r.LeftTwips is int twips) return $"{twips}twips";
        return "未解析出";
    }

    private static int? TocStyleLevel(string? styleId, IReadOnlyDictionary<string, string> styleNames)
    {
        if (styleId is null) return null;
        var match = TocStylePattern.Match(styleId);
        if (!match.Success && styleNames.TryGetValue(styleId, out var name))
            match = TocStylePattern.Match(name);
        return match.Success && int.TryParse(match.Groups[1].Value, out var level) ? level : null;
    }

    private static IReadOnlyDictionary<string, string> BuildStyleNames(XDocument? styles)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (styles?.Root is null) return result;
        foreach (var style in styles.Root.Elements(W + "style"))
        {
            var id = style.Attribute(W + "styleId")?.Value;
            var name = style.Element(W + "name")?.Attribute(W + "val")?.Value;
            if (id is not null && name is not null) result[id] = name;
        }
        return result;
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
