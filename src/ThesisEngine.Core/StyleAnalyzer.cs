using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// Word 样式体系分析器。
/// 检测文档中哪些段落套用了样式，哪些是手动打的格式（含表格内段落），
/// 并报告引用了不存在样式的段落。
/// </summary>
public sealed class StyleAnalyzer
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    public StyleAnalysisResult Analyze(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var stylesPart = TryReadXml(archive, "word/styles.xml");
        var documentPart = TryReadXml(archive, "word/document.xml");

        var allStyles = ExtractAllStyles(stylesPart);
        var paragraphs = ExtractParagraphs(documentPart);

        var styledParagraphs = new List<StyledParagraphInfo>();
        var unstyledParagraphs = new List<UnstyledParagraphInfo>();
        var invalidStyleParagraphs = new List<StyledParagraphInfo>();
        var overriddenParagraphs = new List<DirectFormatOverrideInfo>();

        foreach (var para in paragraphs)
        {
            var pStyle = para.Element(W + "pPr")?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;
            var text = para.Descendants(W + "t").Select(t => t.Value).Aggregate("", (a, b) => a + b);

            if (string.IsNullOrWhiteSpace(text)) continue;

            var inTable = para.Ancestors(W + "tbl").Any();
            var preview = text.Length > 50 ? text.Substring(0, 50) + "..." : text;

            if (string.IsNullOrEmpty(pStyle))
            {
                // 没有套用样式的段落
                var directFormat = ExtractDirectFormatting(para);
                unstyledParagraphs.Add(new UnstyledParagraphInfo(preview, directFormat, inTable));
            }
            else if (!allStyles.ContainsKey(pStyle))
            {
                // 引用了样式表中不存在的样式（无效样式引用）
                invalidStyleParagraphs.Add(new StyledParagraphInfo(pStyle, pStyle, preview, inTable));
                styledParagraphs.Add(new StyledParagraphInfo(pStyle, pStyle, preview, inTable));
            }
            else
            {
                // 套用了样式的段落
                allStyles.TryGetValue(pStyle, out var styleName);
                styledParagraphs.Add(new StyledParagraphInfo(pStyle, styleName ?? pStyle, preview, inTable));

                // 套了样式不代表格式生效：run 上的直接格式优先级高于样式，
                // 只改 styles.xml 对这类段落无效。以前这里算「已归属」直接放过，
                // 于是「改样式→全文自动刷新」在真实文档上是假的，报告还是 0 条。
                var overrides = DirectFormatCleaner.DetectOverrides(para);
                if (overrides.Count > 0)
                {
                    overriddenParagraphs.Add(new DirectFormatOverrideInfo(
                        pStyle,
                        styleName ?? pStyle,
                        preview,
                        string.Join("、", overrides),
                        inTable));
                }
            }
        }

        return new StyleAnalysisResult(
            allStyles,
            styledParagraphs,
            unstyledParagraphs,
            invalidStyleParagraphs,
            overriddenParagraphs);
    }

    private static Dictionary<string, string> ExtractAllStyles(XDocument? stylesDoc)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (stylesDoc?.Root == null) return result;

        foreach (var style in stylesDoc.Root.Elements(W + "style"))
        {
            var styleId = style.Attribute(W + "styleId")?.Value;
            var name = style.Element(W + "name")?.Attribute(W + "val")?.Value;
            if (!string.IsNullOrEmpty(styleId))
            {
                result[styleId] = name ?? styleId;
            }
        }
        return result;
    }

    /// <summary>枚举正文所有段落，包括表格（含嵌套表格）内的段落。</summary>
    private static List<XElement> ExtractParagraphs(XDocument? documentDoc)
    {
        var body = documentDoc?.Root?.Element(W + "body");
        if (body == null) return new List<XElement>();

        return body.Descendants(W + "p").ToList();
    }

    private static string ExtractDirectFormatting(XElement para)
    {
        var pPr = para.Element(W + "pPr");
        var info = new List<string>();

        if (pPr?.Element(W + "jc") != null) info.Add("手动对齐");
        if (pPr?.Element(W + "spacing") != null) info.Add("手动行距");
        if (pPr?.Element(W + "ind") != null) info.Add("手动缩进");

        var runProp = para.Descendants(W + "rPr").FirstOrDefault();
        if (runProp?.Element(W + "rFonts") != null) info.Add("手动字体");
        if (runProp?.Element(W + "sz") != null) info.Add("手动字号");

        return info.Count > 0 ? string.Join("、", info) : "无格式";
    }

    private static XDocument? TryReadXml(ZipArchive archive, string entryName)
    {
        var entry = archive.GetEntry(entryName);
        if (entry == null) return null;
        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }
}

public sealed record StyleAnalysisResult(
    Dictionary<string, string> AllStyles,
    List<StyledParagraphInfo> StyledParagraphs,
    List<UnstyledParagraphInfo> UnstyledParagraphs,
    List<StyledParagraphInfo> InvalidStyleParagraphs,
    List<DirectFormatOverrideInfo> DirectFormatOverrides);

public sealed record StyledParagraphInfo(string StyleId, string StyleName, string Preview, bool InTable);

public sealed record UnstyledParagraphInfo(string Preview, string DirectFormat, bool InTable);

/// <summary>套了样式、但仍有直接格式压在上面的段落。</summary>
public sealed record DirectFormatOverrideInfo(
    string StyleId,
    string StyleName,
    string Preview,
    string Overrides,
    bool InTable);
