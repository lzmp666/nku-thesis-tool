using System.IO.Compression;
using System.Xml.Linq;
using ThesisEngine.Core.Document;

namespace ThesisEngine.Core;

/// <summary>一条题注格式问题。ParagraphIndex 为 document.xml 里 w:p 的 1 基序号。</summary>
public sealed record CaptionIssue(
    string Code,
    string RuleId,
    string CaptionType,
    int ParagraphIndex,
    string Text,
    string Dimension,
    string Expected,
    string Actual,
    string Message);

/// <summary>一个被识别到的题注（信息用途）。</summary>
public sealed record CaptionInfo(
    int ParagraphIndex,
    string CaptionType,
    string Text);

/// <summary>题注格式分析结果。NotChecked 单列「压根没查 / 渲染依赖 / 未决」。</summary>
public sealed record CaptionFormatResult(
    IReadOnlyList<CaptionIssue> Issues,
    IReadOnlyList<CaptionInfo> Captions,
    IReadOnlyList<string> NotChecked,
    IReadOnlyList<string> MissingParts,
    int CheckedCaptions,
    int FigureCount,
    int TableCount,
    string RuleVersion,
    string SourceNote);

/// <summary>
/// 图题/表题格式检查（R12，依据 general.yaml 4.7 表4.6，来源 S1 2025 版），只查不改。
///
/// 口径边界（与 general.yaml figures_and_tables.v1_check:false、覆盖矩阵 M26-037 一致）：
/// - 题注用 FigureTableParagraphs.Classify 的文本锚点识别（「图 1-1 / 表 2-3 / 续表 / Figure / Table」
///   带编号），不靠图片/表格对象的视觉识别。这样即便标了 v1_check:false（对象识别未决），
///   题注段落自身的纯文本格式仍可判，不假装检查过对象级维度。
/// - 只判条文给明确值且不依赖渲染的维度：宋体、10.5 磅、居中、单倍行距、段前后
///   （图题段前 6 / 段后 12 磅，表题段前 6 / 段后 6 磅）。
/// - 序题间空一个汉字（RFC-3 宽度语义未定义）、题注相对图表对象的位置（图下/表上，渲染依赖）、
///   R13 表达式编号（依赖公式对象/渲染）一律列未判定，不据此判合规。
/// - 判不出的维度列未判定，不拿默认值凑问题数。规则状态 draft，需人工确认。
/// - 问题码用每类 CodeStart + 维度偏移映射，K1–K12 连续。
/// </summary>
public sealed class CaptionFormatAnalyzer
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private const string SourceNote =
        "南开大学研究生学位论文写作规范 4.7 表4.6 图题/表题 R12（general.yaml，来源 S1 2025 版，规则状态 draft；对象识别 v1_check:false、RFC-3 待负责人确认）";
    private const string RuleVersion = "nku-graduate-caption-r12-draft";

    private const string AutoLineRule = "auto";
    private const int SingleLineValue = 240;

    private sealed record CaptionSpec(
        string CaptionType,
        string RuleId,
        int CodeStart,
        int AfterTwips,
        IReadOnlyDictionary<string, int> Offsets);

    private static readonly IReadOnlyDictionary<string, int> Dims = new Dictionary<string, int>
    {
        ["Font"] = 1, ["Size"] = 2, ["Align"] = 3, ["Line"] = 4, ["Before"] = 5, ["After"] = 6,
    };
    private const int BeforeTwips = 120; // 段前 6 磅
    private const int SizeHalfPoints = 21; // 10.5 磅（五号）

    private static readonly CaptionSpec FigureSpec = new("图题", "R12", 0, 240, Dims); // 段后 12 磅
    private static readonly CaptionSpec TableSpec = new("表题", "R12", 6, 120, Dims);  // 段后 6 磅

    public CaptionFormatResult Analyze(string docxPath)
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

        var issues = new List<CaptionIssue>();
        var notChecked = new List<string>(reading.Undetermined.Where(u =>
            u.StartsWith("word/document.xml", StringComparison.Ordinal)));
        var captions = new List<CaptionInfo>();
        var checkedCaptions = 0;
        var figureCount = 0;
        var tableCount = 0;

        var body = document.Root?.Element(W + "body");
        var paragraphIndex = 0;
        foreach (var paragraph in body is null
                     ? Enumerable.Empty<XElement>()
                     : FigureTableParagraphs.Of(body))
        {
            paragraphIndex++;
            var text = FigureTableParagraphs.Text(paragraph);
            var kind = FigureTableParagraphs.Classify(text);
            if (kind is null) continue;

            var isFigure = kind == CaptionKind.Figure;
            if (isFigure) figureCount++; else tableCount++;
            var spec = isFigure ? FigureSpec : TableSpec;
            captions.Add(new CaptionInfo(paragraphIndex, spec.CaptionType, text));

            if (!formats.TryGetValue(paragraphIndex, out var r))
            {
                notChecked.Add($"{spec.CaptionType}第 {paragraphIndex} 段「{Short(text)}」：样式链解析跳过，格式未判定");
                continue;
            }
            checkedCaptions++;
            CheckCaption(spec, r, issues, notChecked);
        }

        if (captions.Count == 0)
            notChecked.Add("未识别到带编号的图题/表题段落（图/表/续表/Figure/Table + 编号），题注格式未检查（不靠图片对象猜测）");
        else
            notChecked.Add("题注相对图表对象的位置（图题在图下、表题在表上）依赖渲染与对象识别（v1_check:false），本次未判位置对错");

        notChecked.Add("序题间空一个汉字的宽度语义未定义（RFC-3），未按字符计数判定");
        notChecked.Add("R13 表达式编号（Times New Roman 10.5 磅、右对齐）依赖公式对象/渲染，本次未判，列未判定");
        if (styles is null)
            notChecked.Add("缺少 word/styles.xml：样式链不可用，只按显式 rPr / pPr 判定");

        return new CaptionFormatResult(
            Issues: issues,
            Captions: captions,
            NotChecked: notChecked.Distinct(StringComparer.Ordinal).ToArray(),
            MissingParts: missingParts,
            CheckedCaptions: checkedCaptions,
            FigureCount: figureCount,
            TableCount: tableCount,
            RuleVersion: RuleVersion,
            SourceNote: SourceNote);
    }

    private static void CheckCaption(CaptionSpec spec, PartFormatReading r,
        List<CaptionIssue> issues, List<string> notChecked)
    {
        var label = $"{spec.CaptionType}第 {r.ParagraphIndex} 段「{Short(r.Text)}」";

        void Add(string dim, string expected, string actual)
        {
            issues.Add(new CaptionIssue(
                $"K{spec.CodeStart + spec.Offsets[dim]}", spec.RuleId, spec.CaptionType,
                r.ParagraphIndex, r.Text, dim, expected, actual,
                $"{label}：{dim}期望「{expected}」，实际「{actual}」。依据 {SourceNote}。"));
        }

        if (!r.FontResolved) notChecked.Add($"{label}字体：未判定");
        else if (!PartFormatReader.IsSimSun(r.EastAsiaFont)) Add("Font", "宋体", r.EastAsiaFont ?? "未解析出");

        if (!r.SizeResolved) notChecked.Add($"{label}字号：未判定");
        else if (r.SizeHalfPoints != SizeHalfPoints) Add("Size", "10.5磅", $"{r.SizeHalfPoints / 2.0:0.#}磅");

        if (!r.AlignmentResolved) notChecked.Add($"{label}对齐：未判定");
        else if (!string.Equals(r.Alignment, "center", StringComparison.OrdinalIgnoreCase)) Add("Align", "居中", r.Alignment);

        if (r.LineSpacingRule is null || r.LineSpacingValue is null)
            notChecked.Add($"{label}行距：未判定");
        else if (!(string.Equals(r.LineSpacingRule, AutoLineRule, StringComparison.OrdinalIgnoreCase)
                   && r.LineSpacingValue == SingleLineValue))
            Add("Line", "单倍行距", $"{r.LineSpacingRule}/{r.LineSpacingValue}");

        if (r.SpaceBeforeTwips is int before && before != BeforeTwips)
            Add("Before", "6磅", $"{before / 20.0:0.#}磅");
        else if (r.SpaceBeforeTwips is null)
            notChecked.Add($"{label}段前：未判定");

        if (r.SpaceAfterTwips is int after && after != spec.AfterTwips)
            Add("After", $"{spec.AfterTwips / 20}磅", $"{after / 20.0:0.#}磅");
        else if (r.SpaceAfterTwips is null)
            notChecked.Add($"{label}段后：未判定");
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
