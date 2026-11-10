using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace ThesisEngine.Core.Document;

/// <summary>
/// 文档级自动修复（#112 第二档）：只写 sectPr 下的 pgSz / pgMar 属性。
///
/// 边界比功能重要。issue 把修复分成两类，这里只做「可安全自动改」那一类：
/// 不新增 XML 元素、不删除元素、不动 w:t、不动 header/footer 部件、不动 settings.xml。
/// 目录、页码、空白页、页眉页脚字体字号一律只报告不改，误删内容或改坏域的风险担不起。
///
/// 规范值不在这里另写一份：直接复用 DocumentAnalyzer 的常量，检查和修复必须
/// 认同同一批数字，否则会出现「检查说合规、修复又改一遍」这种来回打架。
/// </summary>
public sealed class DocumentFormatter
{
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    public DocumentFormatResult Apply(string inputPath, string outputPath)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(inputPath);
        ArgumentException.ThrowIfNullOrWhiteSpace(outputPath);

        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException($"找不到输入文件：{inputPath}", inputPath);
        }

        if (string.Equals(
                Path.GetFullPath(inputPath),
                Path.GetFullPath(outputPath),
                OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal))
        {
            throw new ArgumentException("输出路径不能与输入路径相同；为保护原文件，未执行写入。");
        }

        using var input = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var source = new ZipArchive(input, ZipArchiveMode.Read);

        var documentEntry = source.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");

        XDocument document;
        using (var reader = new StreamReader(documentEntry.Open(), Encoding.UTF8))
        {
            document = XDocument.Load(reader);
        }

        var body = document.Root?.Element(W + "body")
            ?? throw new InvalidDataException("文档中没有 body");

        // 复用检查器的节收集逻辑。没有 body 级 sectPr 时它会用最后一节兜底，
        // 那个元素在列表里出现两次，去重避免对同一节点改两遍。
        var sections = DocumentAnalyzer.CollectSectionProperties(body).Distinct().ToList();

        var fixes = new List<DocumentFormatFix>();
        var skips = new List<string>();

        for (var i = 0; i < sections.Count; i++)
        {
            FixSection(sections[i], i + 1, fixes, skips);
        }

        WritePackage(outputPath, source, document);

        return new DocumentFormatResult(sections.Count, fixes, skips);
    }

    private static void FixSection(
        XElement sectPr,
        int index,
        List<DocumentFormatFix> fixes,
        List<string> skips)
    {
        var pgSz = sectPr.Element(W + "pgSz");
        if (pgSz is null)
        {
            skips.Add($"第 {index} 节：sectPr 里没有 pgSz，不自动新增元素（sectPr 子元素有 schema 顺序，" +
                "乱插会让 Word 拒绝打开），纸张需人工设置");
        }
        else
        {
            SetTwips(pgSz, "w", DocumentAnalyzer.A4WidthCm, "D1", "纸张宽度", index, fixes);
            SetTwips(pgSz, "h", DocumentAnalyzer.A4HeightCm, "D1", "纸张高度", index, fixes);
        }

        var pgMar = sectPr.Element(W + "pgMar");
        if (pgMar is null)
        {
            skips.Add($"第 {index} 节：sectPr 里没有 pgMar，不自动新增元素，页边距需人工设置");
            return;
        }

        SetTwips(pgMar, "top", DocumentAnalyzer.MarginTopBottomCm, "D2", "上页边距", index, fixes);
        SetTwips(pgMar, "bottom", DocumentAnalyzer.MarginTopBottomCm, "D2", "下页边距", index, fixes);
        SetTwips(pgMar, "left", DocumentAnalyzer.MarginLeftRightCm, "D2", "左页边距", index, fixes);
        SetTwips(pgMar, "right", DocumentAnalyzer.MarginLeftRightCm, "D2", "右页边距", index, fixes);
        SetTwips(pgMar, "header", DocumentAnalyzer.HeaderFooterCm, "D3", "页眉距边界", index, fixes);
        SetTwips(pgMar, "footer", DocumentAnalyzer.HeaderFooterCm, "D3", "页脚距边界", index, fixes);
        SetTwips(pgMar, "gutter", DocumentAnalyzer.GutterCm, "D4", "装订线", index, fixes);
    }

    /// <summary>
    /// 先按检查器的口径判一次：已经在 0.1cm 容差内的值不动。
    ///
    /// 这一步不能省。实测 Word 16 存 3.8cm 是 2154 twips，而按 cm 换算是 2155；
    /// 两种写法都合规。若修复器无条件写自己的换算值，会把本来就合规的文档
    /// 全改一遍，还和检查器「2155 没问题」的说法打架。检查和修复必须认同
    /// 同一批数字、同一个容差。
    /// </summary>
    private static void SetTwips(
        XElement element,
        string attributeName,
        double expectedCm,
        string code,
        string label,
        int sectionIndex,
        List<DocumentFormatFix> fixes)
    {
        var current = element.Attribute(W + attributeName)?.Value;
        if (current is not null &&
            int.TryParse(current, NumberStyles.Integer, CultureInfo.InvariantCulture, out var currentTwips) &&
            Math.Abs(currentTwips / DocumentAnalyzer.TwipsPerCm - expectedCm) <= DocumentAnalyzer.ToleranceCm)
        {
            return;
        }

        var expected = CmToTwips(expectedCm);
        var expectedText = expected.ToString(CultureInfo.InvariantCulture);
        element.SetAttributeValue(W + attributeName, expectedText);
        fixes.Add(new DocumentFormatFix(
            sectionIndex,
            code,
            $"{label} {current ?? "未设置"} → {expectedText} twips"));
    }

    /// <summary>
    /// 规范只给了 cm，twips 是唯一可推导的换算：1cm = 1440/2.54 twips。
    /// Word 16 自己存 A4 是 11907×16839，这里算出 11906×16838，差 1 twip 属取整方式不同，
    /// 检查器的 0.1cm 容差覆盖得住。不直接抄 Word 的值，是因为它没有条款依据。
    /// </summary>
    private static int CmToTwips(double cm) =>
        (int)Math.Round(cm * DocumentAnalyzer.TwipsPerCm, MidpointRounding.AwayFromZero);

    /// <summary>
    /// 按原顺序重建包：除 document.xml 外所有条目原样复制字节，条目顺序也不变。
    /// 用 ZipArchiveMode.Update 会把替换后的 document.xml 挪到包末尾，虽然不影响打开，
    /// 但「只动 sectPr」这句话就没那么硬了。
    /// </summary>
    private static void WritePackage(string outputPath, ZipArchive source, XDocument document)
    {
        using var output = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
        using var target = new ZipArchive(output, ZipArchiveMode.Create);

        foreach (var entry in source.Entries)
        {
            var newEntry = target.CreateEntry(entry.FullName, CompressionLevel.Optimal);
            // 不带上原时间戳，第二次跑同一份文件就会因为 ZIP 元数据不同而逐字节不同，
            // 幂等性验证也就没法做。
            newEntry.LastWriteTime = entry.LastWriteTime;
            using var destination = newEntry.Open();

            if (string.Equals(entry.FullName, "word/document.xml", StringComparison.Ordinal))
            {
                using var writer = new StreamWriter(destination, new UTF8Encoding(false), leaveOpen: true);
                // DisableFormatting：默认 Save 会重排整份文档、插入一堆空白文本节点，
                // 正文虽没变但 diff 没法看，也可能被别的工具当成内容变化。
                document.Save(writer, SaveOptions.DisableFormatting);
                writer.Flush();
            }
            else
            {
                using var origin = entry.Open();
                origin.CopyTo(destination);
            }
        }
    }
}

public sealed record DocumentFormatResult(
    int SectionCount,
    IReadOnlyList<DocumentFormatFix> Fixes,
    IReadOnlyList<string> Skips);

public sealed record DocumentFormatFix(
    int SectionIndex,
    string Code,
    string Description);
