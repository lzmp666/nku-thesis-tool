using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// 清除压过样式的直接格式（Issue #118）。
///
/// 背景：Word 里 run 的直接格式优先级高于段落样式。真实论文（尤其是手改过的）
/// 绝大多数段落上都有直接格式，只改 styles.xml 对它们完全无效——用户在 Word 里
/// 看到的还是原来的字号，排版一条没变，检测端也报不出来。
///
/// 口径（保守优先，只删确实该由样式管的属性）：
/// - run 级：字体 rFonts、字号 sz/szCs、加粗 b/bCs、斜体 i/iCs、下划线 u、颜色 color。
/// - 段落级：缩进 ind 整块删除；spacing 只删行距 line/lineRule，段前段后保留。
///
/// 刻意不动的三类，以及为什么：
/// - 段落 jc（对齐）：真实文档里居中的图题、居中的「摘要」标题都是直接格式设的，
///   它们所在的标题/题注样式并不带居中。删掉 jc 会让这些段落掉回左对齐，
///   是肉眼可见的破坏。所以只检测、报告，不自动删。
/// - vertAlign（上下标）、lang（语言标记）、rStyle：不是「格式与样式打架」，
///   删了会破坏引用标注、拼写检查和字符样式。
/// - 段前段后间距：标题多半靠它撑开版面，样式里没有对应定义，删了版面会挤在一起。
/// </summary>
public sealed class DirectFormatCleaner
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    /// <summary>run 级里该由样式管的属性。</summary>
    private static readonly string[] RunProperties =
    {
        "rFonts", "sz", "szCs", "b", "bCs", "i", "iCs", "u", "color"
    };

    /// <summary>段落级里该由样式管的属性（整块删除）。</summary>
    private static readonly string[] ParagraphProperties = { "ind" };

    public sealed record CleanResult(
        int ParagraphsScanned,
        int ParagraphsTouched,
        int RunsTouched,
        int ParagraphsSkippedNoStyle,
        IReadOnlyDictionary<string, int> RemovedCounts);

    /// <summary>
    /// 清除 inputPath 里压过样式的直接格式，写到 outputPath。
    /// 只处理已经有有效段落样式的段落——没有样式的段落，直接格式就是它唯一的格式，
    /// 删掉等于把版面抹平，交给 StyleAutoAssigner 先套样式再来。
    /// </summary>
    public CleanResult Clean(string inputPath, string outputPath)
    {
        var doc = ReadXml(inputPath, "word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");
        var styleIds = ReadAvailableStyleIds(inputPath);
        var body = doc.Root?.Element(W + "body")
            ?? throw new InvalidDataException("文档中没有 body");

        var removed = new Dictionary<string, int>(StringComparer.Ordinal);
        int scanned = 0, touched = 0, runsTouched = 0, skippedNoStyle = 0;

        void Bump(string kind) => removed[kind] = removed.GetValueOrDefault(kind) + 1;

        foreach (var para in body.Descendants(W + "p"))
        {
            // 文本框里的内容常见「整块手动排版」，不属于本工具管辖范围，跳过。
            if (para.Ancestors(W + "txbxContent").Any()) continue;

            // 空段落没有可读的内容，可能承载分页符或书签，一律不动。
            if (string.IsNullOrWhiteSpace(ParagraphText(para))) continue;

            scanned++;

            var pPr = para.Element(W + "pPr");
            var styleId = pPr?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;
            if (string.IsNullOrEmpty(styleId) || !styleIds.Contains(styleId))
            {
                skippedNoStyle++;
                continue;
            }

            var paragraphChanged = false;

            // ---- 段落级 ----
            foreach (var name in ParagraphProperties)
            {
                var element = pPr?.Element(W + name);
                if (element == null) continue;

                element.Remove();
                Bump("段落" + Label(name));
                paragraphChanged = true;
            }

            var spacing = pPr?.Element(W + "spacing");
            if (spacing != null)
            {
                // 只清行距，段前段后保留：标题靠它撑版面，样式里通常没有对应定义。
                var line = spacing.Attribute(W + "line");
                var lineRule = spacing.Attribute(W + "lineRule");
                if (line != null || lineRule != null)
                {
                    line?.Remove();
                    lineRule?.Remove();
                    Bump("段落手动行距");
                    paragraphChanged = true;

                    if (!spacing.Attributes().Any())
                    {
                        spacing.Remove();
                    }
                }
            }

            // ---- run 级 ----
            foreach (var run in para.Descendants(W + "r"))
            {
                // 域代码（fldChar / instrText）本身不带可见文字，动它没有意义还容易踩坑。
                if (!run.Elements(W + "t").Any()) continue;

                var rPr = run.Element(W + "rPr");
                if (rPr == null) continue;

                var runChanged = false;
                foreach (var name in RunProperties)
                {
                    var element = rPr.Element(W + name);
                    if (element == null) continue;

                    element.Remove();
                    Bump(Label(name));
                    runChanged = true;
                }

                if (runChanged)
                {
                    runsTouched++;
                    paragraphChanged = true;

                    if (!rPr.HasElements && !rPr.HasAttributes)
                    {
                        rPr.Remove(); // 空 rPr 会被 Word 当成噪声，顺手清掉
                    }
                }
            }

            if (paragraphChanged) touched++;
        }

        WriteXml(inputPath, outputPath, doc);

        return new CleanResult(scanned, touched, runsTouched, skippedNoStyle, removed);
    }

    /// <summary>
    /// 列出某个段落上「压过样式」的直接格式，供 analyze-styles 报告。
    /// 段落没有样式时不报——那种情况归「没有套用样式的段落」管。
    /// </summary>
    public static IReadOnlyList<string> DetectOverrides(XElement para)
    {
        var pPr = para.Element(W + "pPr");
        if (pPr == null) return Array.Empty<string>();

        var found = new List<string>();

        if (pPr.Element(W + "ind") != null) found.Add("手动缩进");

        var spacing = pPr.Element(W + "spacing");
        if (spacing?.Attribute(W + "line") != null || spacing?.Attribute(W + "lineRule") != null)
        {
            found.Add("手动行距");
        }

        // 对齐方式只是提示，不自动改：居中的题注/标题多半就是这么设的。
        if (pPr.Element(W + "jc") != null) found.Add("手动对齐（不自动清除）");

        // 只统计 run 级的 rPr，且只统计真正带文字的 run，和 Clean 实际清理的范围保持一致——
        // 否则 analyze 会报出 full-format 修不了的项，用户来回试。
        // 段落标记的格式（pPr/rPr）不算：它作用于回车符，不影响正文字形，
        // 动它还可能改变编号的外观。批注引用 run（无 w:t）也不算：清掉会破坏批注标记。
        var kinds = para.Descendants(W + "r")
            .Where(r => r.Elements(W + "t").Any())
            .Select(r => r.Element(W + "rPr"))
            .Where(rPr => rPr != null)
            .SelectMany(rPr => RunProperties.Where(name => rPr!.Element(W + name) != null))
            .Select(Label)
            .Distinct(StringComparer.Ordinal)
            .Where(label => label.Length > 0);

        found.AddRange(kinds);
        return found;
    }

    private static string Label(string propertyName) => propertyName switch
    {
        "rFonts" => "手动字体",
        "sz" or "szCs" => "手动字号",
        "b" or "bCs" => "手动加粗",
        "i" or "iCs" => "手动斜体",
        "u" => "手动下划线",
        "color" => "手动颜色",
        "ind" => "缩进",
        _ => string.Empty
    };

    private static string ParagraphText(XElement para) =>
        para.Descendants(W + "t").Select(t => t.Value).Aggregate("", (a, b) => a + b);

    private static IReadOnlySet<string> ReadAvailableStyleIds(string docxPath)
    {
        var styles = ReadXml(docxPath, "word/styles.xml");
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (styles?.Root == null) return ids;

        foreach (var style in styles.Root.Elements(W + "style"))
        {
            var id = style.Attribute(W + "styleId")?.Value;
            if (!string.IsNullOrEmpty(id)) ids.Add(id);
        }
        return ids;
    }

    private static XDocument? ReadXml(string docxPath, string entryName)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var entry = archive.GetEntry(entryName);
        if (entry == null) return null;

        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }

    private static void WriteXml(string inputPath, string outputPath, XDocument doc)
    {
        if (!string.Equals(
                Path.GetFullPath(inputPath),
                Path.GetFullPath(outputPath),
                StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(inputPath, outputPath, overwrite: true);
        }

        using var stream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update);

        archive.GetEntry("word/document.xml")?.Delete();

        var entry = archive.CreateEntry("word/document.xml");
        using var writer = new StreamWriter(entry.Open());
        doc.Save(writer);
    }
}
