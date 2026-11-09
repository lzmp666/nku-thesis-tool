using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// 图表排版分析器：检测图片/表格排版问题与可疑的跨页设置。
/// 这是“能否改成样式化排版”的初步体检，不代替 Word 实机渲染检查。
/// </summary>
public sealed class FigureTableAnalyzer
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    public FigureTableAnalysisResult Analyze(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var documentPart = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");

        XDocument doc;
        using (var reader = new StreamReader(documentPart.Open()))
        {
            doc = XDocument.Load(reader);
        }

        var body = doc.Root?.Element(W + "body")
            ?? throw new InvalidDataException("文档中没有 body");

        var images = body.Descendants().Where(e => e.Name.LocalName == "blip").ToList();
        var tables = body.Elements(W + "tbl").ToList();

        return new FigureTableAnalysisResult(
            ImagesCount: images.Count,
            TablesCount: tables.Count,
            PageBreakIssues: FindPageBreakIssues(body),
            TableIssues: FindTableIssues(tables));
    }

    private static List<string> FindTableIssues(List<XElement> tables)
    {
        var issues = new List<string>();
        for (var index = 0; index < tables.Count; index++)
        {
            var number = index + 1;
            var borders = tables[index].Element(W + "tblPr")?.Element(W + "tblBorders");
            if (borders == null)
            {
                // 边框可能继承自表格样式，仅凭 document.xml 无法确认它是不是三线表。
                issues.Add($"表格 {number} 未显式设置三线表边框（边框可能继承自表格样式），请确认或套用三线表样式");
                continue;
            }

            if (IsBorderVisible(borders.Element(W + "insideH")) || IsBorderVisible(borders.Element(W + "insideV")))
            {
                issues.Add($"表格 {number} 存在内部横线或竖线，不是三线表");
                continue;
            }

            if (IsBorderVisible(borders.Element(W + "left")) || IsBorderVisible(borders.Element(W + "right")))
            {
                issues.Add($"表格 {number} 存在左右外框线；三线表只保留顶线、表头下线和底线");
            }
        }

        return issues;
    }

    /// <summary>
    /// 边框是否可见。OOXML 中 <c>w:val="nil"</c> / <c>"none"</c> 表示无边框；
    /// 缺少 val 时 Word 按 single 处理；元素缺失表示继承样式，这里保守地按“可见”判断。
    /// </summary>
    private static bool IsBorderVisible(XElement? border)
    {
        if (border == null) return true;
        var value = (string?)border.Attribute(W + "val") ?? "single";
        return !value.Equals("nil", StringComparison.OrdinalIgnoreCase)
            && !value.Equals("none", StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> FindPageBreakIssues(XElement body)
    {
        var issues = new List<string>();
        foreach (var paragraph in body.Elements(W + "p"))
        {
            var properties = paragraph.Element(W + "pPr");
            var text = string.Concat(paragraph.Descendants(W + "t").Select(t => t.Value)).Trim();

            if (IsOn(properties?.Element(W + "pageBreakBefore")))
            {
                issues.Add($"段落“{Shorten(text)}”设置了段前分页，可能产生大段空白页");
                continue;
            }

            var isTableCaption = text.StartsWith("表", StringComparison.Ordinal)
                || text.StartsWith("Table", StringComparison.OrdinalIgnoreCase);
            if (isTableCaption && text.Length <= 40 && !IsOn(properties?.Element(W + "keepNext")))
            {
                issues.Add($"表题“{Shorten(text)}”未设置与下段同页，可能与表格分离");
                continue;
            }

            if (text.Length > 50 && !IsOn(properties?.Element(W + "widowControl")))
            {
                issues.Add($"段落“{Shorten(text)}”未开启孤行控制，可能出现跨页孤行");
            }
        }

        return issues;
    }

    /// <summary>读取开关型段落属性：元素存在即视为开启，val=0/false/off 视为关闭。</summary>
    private static bool IsOn(XElement? element)
    {
        if (element == null) return false;
        var value = (string?)element.Attribute(W + "val");
        if (string.IsNullOrEmpty(value)) return true;
        return !value.Equals("0", StringComparison.Ordinal)
            && !value.Equals("false", StringComparison.OrdinalIgnoreCase)
            && !value.Equals("off", StringComparison.OrdinalIgnoreCase);
    }

    private static string Shorten(string text) => text.Length <= 18 ? text : text[..18] + "…";
}

public sealed record FigureTableAnalysisResult(
    int ImagesCount,
    int TablesCount,
    List<string> PageBreakIssues,
    List<string> TableIssues);
