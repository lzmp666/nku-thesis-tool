using System.Xml.Linq;

namespace ThesisEngine.Core.Document;

/// <summary>
/// 页眉 / 页脚部件里，文字实际呈现出来的字体、字号、对齐（#112 第二档）。
///
/// 判定的前提是能展开样式链：run 的 rPr → 段落样式 → basedOn → docDefaults。
/// 链断了、或者包里根本没有 styles.xml，就按属性分别记「未判定」，不拿左对齐 /
/// 12 磅之类的默认值去凑问题数。这一条和仓库「查了没问题」要跟「压根没查」分开
/// 的口径一致。run 上显式写了值照样判，不因为缺 styles.xml 就把能判的也丢掉。
///
/// 只认显式能读到的值：字体取 rFonts/@eastAsia 与 @ascii(或 @hAnsi)，
/// 字号取 sz/@val（半点），对齐取 jc/@val。
/// </summary>
public static class PartFormatReader
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>R2 / R3 都写 10.5 磅，XML 里按半点存 21。</summary>
    public const int RequiredSizeHalfPoints = 21;

    public static PartFormatResult Read(XDocument part, string partName, XDocument? styles)
    {
        var readings = new List<PartFormatReading>();
        var undetermined = new List<string>();

        if (part.Root is null)
        {
            undetermined.Add($"{partName}：部件为空，格式未判定");
            return new PartFormatResult(readings, undetermined);
        }

        if (styles is null)
        {
            // 真 Word 文档一定有 styles.xml。缺了它的多半是手工拼的骨架包，
            // 这时链上什么都没有：显式写的值照判，没写的不能按「默认左对齐」硬报。
            undetermined.Add($"{partName}：文档没有 styles.xml，样式链不可用，只按显式 rPr / jc 判定");
        }

        var styleIndex = BuildStyleIndex(styles);
        var defaultStyleId = FindDefaultParagraphStyleId(styles);
        var defaults = styles?.Root?.Element(W + "docDefaults");
        var defaultRunProperties = defaults?.Element(W + "rPrDefault")?.Element(W + "rPr");
        var defaultParagraphProperties = defaults?.Element(W + "pPrDefault")?.Element(W + "pPr");

        var paragraphIndex = 0;
        foreach (var paragraph in part.Root.Descendants(W + "p"))
        {
            paragraphIndex++;

            var runs = paragraph.Elements(W + "r")
                .Where(run => run.Elements(W + "t").Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                .ToList();
            if (runs.Count == 0) continue;

            var text = string.Concat(paragraph.Descendants(W + "t").Select(node => node.Value)).Trim();
            var westernOnly = text.All(character => character <= 0x7F);

            var pPr = paragraph.Element(W + "pPr");
            var styleId = pPr?.Element(W + "pStyle")?.Attribute(W + "val")?.Value ?? defaultStyleId;
            var chain = StyleChain(styleId, styleIndex).ToList();
            var alignment = ResolveAlignment(pPr, chain, defaultParagraphProperties, styles is not null);

            var resolved = runs
                .Select(run => ResolveRun(run.Element(W + "rPr"), chain, defaultRunProperties))
                .ToList();

            // 同一段落里多个 run 格式不一致时，取谁都不代表这一段。宁可标未判定，
            // 也不能挑一个值报出去让人照着改。
            var sizeValues = resolved.Select(item => item.SizeHalfPoints).Distinct().ToList();
            var eastAsiaValues = resolved.Select(item => item.EastAsiaFont).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var westernValues = resolved.Select(item => item.WesternFont).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            if (sizeValues.Count > 1 || eastAsiaValues.Count > 1 || westernValues.Count > 1)
            {
                undetermined.Add(
                    $"{partName} 第 {paragraphIndex} 段：段内多个文字 run 的格式不一致，未判定（需人工确认）");
                continue;
            }

            var first = resolved[0];
            readings.Add(new PartFormatReading(
                PartName: partName,
                ParagraphIndex: paragraphIndex,
                Text: text,
                WesternOnly: westernOnly,
                StylesPartAvailable: styles is not null,
                EastAsiaFont: first.EastAsiaFont,
                WesternFont: first.WesternFont,
                SizeHalfPoints: first.SizeHalfPoints,
                Alignment: alignment.Value,
                AlignmentResolved: alignment.Resolved,
                FontResolved: first.FontResolved,
                SizeResolved: first.SizeResolved));
        }

        return new PartFormatResult(readings, undetermined);
    }

    /// <summary>宋体 / SimSun 在中英文版 Word 里都可能存成不同写法。</summary>
    public static bool IsSimSun(string? font) => MatchesAny(font, "宋体", "simsun");

    /// <summary>Times New Roman 的空格在不同版本里时有时无。</summary>
    public static bool IsTimesNewRoman(string? font) => MatchesAny(font, "timesnewroman");

    private static bool MatchesAny(string? font, params string[] accepted)
    {
        if (font is null) return false;
        var normalized = font.Replace(" ", string.Empty).ToLowerInvariant();
        return accepted.Any(candidate => normalized == candidate.ToLowerInvariant());
    }

    private static RunFormat ResolveRun(
        XElement? direct,
        IReadOnlyList<XElement> chain,
        XElement? defaults)
    {
        var eastAsia = ResolveRunValue(direct, chain, defaults, "rFonts", "eastAsia");
        var western = ResolveRunValue(direct, chain, defaults, "rFonts", "ascii")
            ?? ResolveRunValue(direct, chain, defaults, "rFonts", "hAnsi");
        var size = ResolveRunValue(direct, chain, defaults, "sz", "val");

        return new RunFormat(
            EastAsiaFont: eastAsia,
            WesternFont: western,
            SizeHalfPoints: int.TryParse(size, out var halfPoints) ? halfPoints : null,
            FontResolved: eastAsia is not null || western is not null,
            SizeResolved: size is not null);
    }

    private static string? ResolveRunValue(
        XElement? direct,
        IReadOnlyList<XElement> chain,
        XElement? defaults,
        string childName,
        string attributeName)
    {
        var fromRun = direct?.Element(W + childName)?.Attribute(W + attributeName)?.Value;
        if (fromRun is not null) return fromRun;

        foreach (var style in chain)
        {
            var fromStyle = style.Element(W + "rPr")?.Element(W + childName)?.Attribute(W + attributeName)?.Value;
            if (fromStyle is not null) return fromStyle;
        }

        return defaults?.Element(W + childName)?.Attribute(W + attributeName)?.Value;
    }

    private static (string Value, bool Resolved) ResolveAlignment(
        XElement? pPr,
        IReadOnlyList<XElement> chain,
        XElement? defaults,
        bool stylesAvailable)
    {
        var fromParagraph = pPr?.Element(W + "jc")?.Attribute(W + "val")?.Value;
        if (fromParagraph is not null) return (fromParagraph, true);

        foreach (var style in chain)
        {
            var fromStyle = style.Element(W + "pPr")?.Element(W + "jc")?.Attribute(W + "val")?.Value;
            if (fromStyle is not null) return (fromStyle, true);
        }

        var fromDefaults = defaults?.Element(W + "jc")?.Attribute(W + "val")?.Value;
        if (fromDefaults is not null) return (fromDefaults, true);

        // 样式链完整时，链上全都没有 jc 就等同于 Word 默认的左对齐，这是确定事实。
        // 链不可用时看不到「页眉」样式，缺省值未必是真的，只能标未判定。
        return stylesAvailable ? ("left", true) : ("unknown", false);
    }

    private static Dictionary<string, XElement> BuildStyleIndex(XDocument? styles)
    {
        var result = new Dictionary<string, XElement>(StringComparer.Ordinal);
        if (styles?.Root is null) return result;

        foreach (var style in styles.Root.Elements(W + "style"))
        {
            var id = style.Attribute(W + "styleId")?.Value;
            if (id is not null) result[id] = style;
        }

        return result;
    }

    private static string? FindDefaultParagraphStyleId(XDocument? styles)
    {
        if (styles?.Root is null) return null;

        var defaultStyle = styles.Root.Elements(W + "style")
            .FirstOrDefault(style =>
                string.Equals(style.Attribute(W + "type")?.Value, "paragraph", StringComparison.Ordinal) &&
                IsOn(style.Attribute(W + "default")?.Value));
        return defaultStyle?.Attribute(W + "styleId")?.Value;
    }

    private static bool IsOn(string? value) => value is "1" or "true" or "on";

    /// <summary>沿 basedOn 往上走。带环防护，坏文档不该把检查器挂死。</summary>
    private static IEnumerable<XElement> StyleChain(
        string? styleId,
        IReadOnlyDictionary<string, XElement> index)
    {
        var visited = new HashSet<string>(StringComparer.Ordinal);
        var current = styleId;
        while (current is not null && visited.Add(current) && index.TryGetValue(current, out var style))
        {
            yield return style;
            current = style.Element(W + "basedOn")?.Attribute(W + "val")?.Value;
        }
    }

    private sealed record RunFormat(
        string? EastAsiaFont,
        string? WesternFont,
        int? SizeHalfPoints,
        bool FontResolved,
        bool SizeResolved);
}

public sealed record PartFormatResult(
    IReadOnlyList<PartFormatReading> Readings,
    IReadOnlyList<string> Undetermined);

/// <summary>一段有文字的页眉 / 页脚段落，展开样式链之后的实际格式。</summary>
public sealed record PartFormatReading(
    string PartName,
    int ParagraphIndex,
    string Text,
    bool WesternOnly,
    bool StylesPartAvailable,
    string? EastAsiaFont,
    string? WesternFont,
    int? SizeHalfPoints,
    string Alignment,
    bool AlignmentResolved,
    bool FontResolved,
    bool SizeResolved)
{
    /// <summary>三个维度都能读到值，才算这一段完整查过。</summary>
    public bool FullyResolved => FontResolved && SizeResolved && AlignmentResolved;
}
