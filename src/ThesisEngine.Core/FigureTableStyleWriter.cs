using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// 图表相关的 Word 样式定义：
/// <list type="bullet">
///   <item><c>NKUThreeLine</c>：三线表（顶/底 1.5 磅、表头线 0.5 磅、无竖线与其他横线）。</item>
///   <item><c>NKUFigure</c>：图片所在段落（居中）。</item>
/// </list>
/// 排版通过“套用样式”完成，而不是逐个表格/图片写直接格式。
/// </summary>
public static class FigureTableStyleWriter
{
    public const string ThreeLineTableStyleId = "NKUThreeLine";
    public const string FigureParagraphStyleId = "NKUFigure";
    public const string FigureCaptionStyleId = "NKUFigureCaption";
    public const string TableCaptionStyleId = "NKUTableCaption";

    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    /// <summary>确保 styles.xml 中存在两个图表样式；已存在则保持原样（幂等）。</summary>
    public static void EnsureStyles(XDocument stylesDocument)
    {
        var root = stylesDocument.Root ?? throw new InvalidDataException("styles.xml 缺少根节点");
        if (!HasStyle(root, FigureParagraphStyleId))
        {
            root.Add(new XElement(W + "style",
                new XAttribute(W + "type", "paragraph"),
                new XAttribute(W + "styleId", FigureParagraphStyleId),
                new XElement(W + "name", new XAttribute(W + "val", "NKU Figure")),
                new XElement(W + "qFormat"),
                new XElement(W + "pPr", new XElement(W + "jc", new XAttribute(W + "val", "center")))));
        }

        if (!HasStyle(root, ThreeLineTableStyleId))
        {
            root.Add(new XElement(W + "style",
                new XAttribute(W + "type", "table"),
                new XAttribute(W + "styleId", ThreeLineTableStyleId),
                new XElement(W + "name", new XAttribute(W + "val", "NKU Three-Line Table")),
                new XElement(W + "uiPriority", new XAttribute(W + "val", "59")),
                new XElement(W + "qFormat"),
                new XElement(W + "tblPr", ThreeLineBorders()),
                new XElement(W + "tblStylePr",
                    new XAttribute(W + "type", "firstRow"),
                    new XElement(W + "tcPr", new XElement(W + "tcBorders", Border("bottom", "single", "4"))))));
        }

        // 图题：段前 6 磅、段后 12 磅、单倍行距、宋体 10.5 磅居中。
        if (!HasStyle(root, FigureCaptionStyleId))
        {
            root.Add(CaptionStyle(FigureCaptionStyleId, "NKU Figure Caption", "240"));
        }

        // 表题：段前 6 磅、段后 6 磅，其余同图题。
        if (!HasStyle(root, TableCaptionStyleId))
        {
            root.Add(CaptionStyle(TableCaptionStyleId, "NKU Table Caption", "120"));
        }
    }

    private static XElement CaptionStyle(string styleId, string name, string afterTwips) => new(
        W + "style",
        new XAttribute(W + "type", "paragraph"),
        new XAttribute(W + "styleId", styleId),
        new XElement(W + "name", new XAttribute(W + "val", name)),
        new XElement(W + "qFormat"),
        new XElement(W + "pPr",
            new XElement(W + "spacing",
                new XAttribute(W + "before", "120"),
                new XAttribute(W + "after", afterTwips),
                new XAttribute(W + "line", "240"),
                new XAttribute(W + "lineRule", "auto")),
            new XElement(W + "jc", new XAttribute(W + "val", "center"))),
        new XElement(W + "rPr",
            new XElement(W + "rFonts",
                new XAttribute(W + "ascii", "Times New Roman"),
                new XAttribute(W + "hAnsi", "Times New Roman"),
                new XAttribute(W + "eastAsia", "宋体")),
            new XElement(W + "sz", new XAttribute(W + "val", "21")),
            new XElement(W + "szCs", new XAttribute(W + "val", "21"))));

    /// <summary>三线表的表级边框定义，供样式与降级路径共用。</summary>
    public static XElement ThreeLineBorders() => new(W + "tblBorders",
        Border("top", "single", "12"),
        Border("left", "nil", null),
        Border("bottom", "single", "12"),
        Border("right", "nil", null),
        Border("insideH", "nil", null),
        Border("insideV", "nil", null));

    public static XElement Border(string side, string value, string? size)
    {
        var border = new XElement(W + side, new XAttribute(W + "val", value));
        if (size != null)
        {
            border.SetAttributeValue(W + "sz", size);
            border.SetAttributeValue(W + "space", "0");
            border.SetAttributeValue(W + "color", "000000");
        }

        return border;
    }

    /// <summary>读取样式表中某表格样式定义的三线边框（不跟随 basedOn 链，仅本样式）。</summary>
    public static XElement? FindTableBorders(XDocument? stylesDocument, string? styleId)
    {
        var root = stylesDocument?.Root;
        if (root == null || string.IsNullOrWhiteSpace(styleId)) return null;
        var style = root.Elements(W + "style").FirstOrDefault(s =>
            (string?)s.Attribute(W + "styleId") == styleId &&
            (string?)s.Attribute(W + "type") == "table");
        return style?.Element(W + "tblPr")?.Element(W + "tblBorders");
    }

    private static bool HasStyle(XElement root, string styleId) =>
        root.Elements(W + "style").Any(s => (string?)s.Attribute(W + "styleId") == styleId);
}
