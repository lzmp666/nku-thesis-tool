using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// 自动给无样式段落套用对应样式（结构识别 + 样式归属）。
/// 学长要求：不建议文中出现没有样式归属的内容。
///
/// 结构识别规则（按优先级）：
/// - "第X章 / Chapter N" 或结构性标题（摘要、目录、参考文献等） → Heading1
/// - "X.Y" 编号标题 → Heading2
/// - "X.Y.Z" 编号标题 → Heading3
/// - 短行 + 加粗 → Heading1
/// - "[N]" 开头的参考文献条目 → Bibliography
/// - 摘要段落 → Abstract
/// - 其他 → Normal
/// 所有候选样式都必须真实存在于文档样式表中，否则回退 Normal；
/// 表格内的无样式段落同样会被识别并套样式。
/// </summary>
public sealed class StyleAutoAssigner
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    private const int MaxHeadingLength = 50;

    private static readonly Regex ChapterHeadingRegex = new(
        @"^\s*(第[一二三四五六七八九十百0-9]+[章节]|Chapter\s+\d+)",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex Level2HeadingRegex = new(
        @"^\s*\d+\.\d+(?!\.\d)(\s|　|\p{Zs})?\s*\S",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex Level3HeadingRegex = new(
        @"^\s*\d+\.\d+\.\d+(\s|　|\p{Zs})?\s*\S",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex StructuralHeadingRegex = new(
        @"^\s*(摘\s*要|Abstract|ABSTRACT|目\s*录|参考文献|结\s*论|致\s*谢|引\s*言|绪\s*论|关键词|Key\s?words)\s*$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly Regex BibliographyEntryRegex = new(
        @"^\s*\[\d+\]",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private readonly TextWriter _output;

    public StyleAutoAssigner(TextWriter? output = null)
    {
        _output = output ?? Console.Out;
    }

    /// <summary>一次自动套用的统计结果。</summary>
    public sealed record StyleAssignmentResult(
        int TotalAssigned,
        IReadOnlyDictionary<string, int> AssignedCounts);

    /// <summary>
    /// 自动识别文档结构并给无样式段落套用样式（含表格内段落）。
    /// </summary>
    public StyleAssignmentResult AutoAssignStyles(string inputPath, string outputPath)
    {
        var doc = ReadDocumentXml(inputPath);
        var availableStyles = ReadAvailableStyleIds(inputPath);
        var body = doc.Root?.Element(W + "body")
            ?? throw new InvalidDataException("文档中没有 body");

        var assignedCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        foreach (var para in body.Descendants(W + "p"))
        {
            var pPr = para.Element(W + "pPr");
            var existingStyle = pPr?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;

            // 已经有样式了就跳过
            if (!string.IsNullOrEmpty(existingStyle)) continue;

            var text = GetParagraphText(para);
            if (string.IsNullOrWhiteSpace(text)) continue;

            var candidate = Classify(text, para, availableStyles);
            if (candidate == null) continue;

            if (pPr == null)
            {
                pPr = new XElement(W + "pPr");
                para.AddFirst(pPr);
            }

            var pStyle = pPr.Element(W + "pStyle");
            if (pStyle == null)
            {
                pStyle = new XElement(W + "pStyle");
                pPr.AddFirst(pStyle);
            }
            pStyle.SetAttributeValue(W + "val", candidate);

            assignedCounts[candidate] = assignedCounts.GetValueOrDefault(candidate) + 1;
        }

        WriteDocumentXml(inputPath, outputPath, doc);

        var result = new StyleAssignmentResult(
            assignedCounts.Values.Sum(),
            assignedCounts);

        _output.WriteLine($"已自动给 {result.TotalAssigned} 个无样式段落套用样式" +
            (result.AssignedCounts.Count > 0
                ? "：" + string.Join("、", result.AssignedCounts.Select(kv => $"{kv.Key}×{kv.Value}"))
                : string.Empty));

        return result;
    }

    /// <summary>
    /// 根据文本内容识别段落应归属的样式。
    /// 候选样式不存在于样式表时回退 Normal；Normal 也不存在时返回 null（跳过）。
    /// </summary>
    private static string? Classify(string text, XElement para, IReadOnlySet<string> availableStyles)
    {
        var trimmed = text.Trim();
        var isBold = para.Descendants(W + "b").Any();

        string? candidate;
        if (trimmed.Length <= MaxHeadingLength && Level3HeadingRegex.IsMatch(trimmed))
            candidate = "Heading3";
        else if (trimmed.Length <= MaxHeadingLength && Level2HeadingRegex.IsMatch(trimmed))
            candidate = "Heading2";
        else if (trimmed.Length <= MaxHeadingLength && ChapterHeadingRegex.IsMatch(trimmed))
            candidate = "Heading1";
        else if (trimmed.Length <= MaxHeadingLength && StructuralHeadingRegex.IsMatch(trimmed))
            candidate = "Heading1";
        else if (BibliographyEntryRegex.IsMatch(trimmed))
            candidate = "Bibliography";
        else if (IsAbstractParagraph(trimmed))
            candidate = "Abstract";
        else if (trimmed.Length < 30 && isBold)
            candidate = "Heading1";
        else
            candidate = "Normal";

        return Resolve(candidate, availableStyles);
    }

    private static bool IsAbstractParagraph(string trimmed) =>
        trimmed.StartsWith("摘要", StringComparison.Ordinal) ||
        trimmed.StartsWith("Abstract", StringComparison.OrdinalIgnoreCase) ||
        trimmed.StartsWith("ABSTRACT", StringComparison.Ordinal);

    /// <summary>候选样式存在就返回；否则回退 Normal；Normal 也缺失时跳过。</summary>
    private static string? Resolve(string candidate, IReadOnlySet<string> availableStyles)
    {
        if (availableStyles.Contains(candidate)) return candidate;
        if (availableStyles.Contains("Normal")) return "Normal";
        return null;
    }

    private static string GetParagraphText(XElement para) =>
        para.Descendants(W + "t").Select(t => t.Value).Aggregate("", (a, b) => a + b);

    /// <summary>读取样式表中实际存在的样式 ID（含默认样式）。</summary>
    private static IReadOnlySet<string> ReadAvailableStyleIds(string docxPath)
    {
        var styles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = archive.GetEntry("word/styles.xml");
        if (entry == null) return styles;

        using var reader = new StreamReader(entry.Open());
        var stylesDoc = XDocument.Load(reader);
        foreach (var style in stylesDoc.Root?.Elements(W + "style") ?? Enumerable.Empty<XElement>())
        {
            var styleId = style.Attribute(W + "styleId")?.Value;
            if (!string.IsNullOrEmpty(styleId))
                styles.Add(styleId);
        }
        return styles;
    }

    private static XDocument ReadDocumentXml(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");

        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }

    private static void WriteDocumentXml(string inputPath, string outputPath, XDocument doc)
    {
        if (!string.Equals(Path.GetFullPath(inputPath), Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(inputPath, outputPath, overwrite: true);
        }

        using var stream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update);

        archive.GetEntry("word/document.xml")?.Delete();

        var newEntry = archive.CreateEntry("word/document.xml");
        using var writer = new StreamWriter(newEntry.Open());
        doc.Save(writer);
    }
}
