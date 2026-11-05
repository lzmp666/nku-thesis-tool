using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// 自动给无样式段落套用对应样式。
/// 学长要求：不建议文中出现没有样式归属的内容。
/// </summary>
public sealed class StyleAutoAssigner
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    /// <summary>
    /// 自动给无样式段落套用样式。
    /// 规则：
    /// - 短行 + 加粗 → 当成标题
    /// - 其他 → 套用 Normal 样式
    /// </summary>
    public void AutoAssignStyles(string inputPath, string outputPath)
    {
        var doc = ReadDocumentXml(inputPath);
        var body = doc.Root?.Element(W + "body");
        if (body == null) return;

        var paragraphs = body.Elements(W + "p").ToList();
        int assignedCount = 0;

        foreach (var para in paragraphs)
        {
            var pPr = para.Element(W + "pPr");
            var existingStyle = pPr?.Element(W + "pStyle")?.Attribute(W + "val")?.Value;

            // 已经有样式了就跳过
            if (!string.IsNullOrEmpty(existingStyle)) continue;

            var text = para.Descendants(W + "t").Select(t => t.Value).Aggregate("", (a, b) => a + b);
            if (string.IsNullOrWhiteSpace(text)) continue;

            // 判断是不是标题：短文本（<30字） + 加粗
            var isBold = para.Descendants(W + "b").Any();
            var isHeading = text.Length < 30 && isBold;

            // 创建 pPr
            if (pPr == null)
            {
                pPr = new XElement(W + "pPr");
                para.AddFirst(pPr);
            }

            // 套用样式
            var pStyle = pPr.Element(W + "pStyle");
            if (pStyle == null)
            {
                pStyle = new XElement(W + "pStyle");
                pPr.AddFirst(pStyle);
            }

            if (isHeading)
            {
                pStyle.SetAttributeValue(W + "val", "Heading1");
            }
            else
            {
                pStyle.SetAttributeValue(W + "val", "Normal");
            }

            assignedCount++;
        }

        WriteDocumentXml(inputPath, outputPath, doc);
        Console.WriteLine($"已自动给 {assignedCount} 个无样式段落套用样式");
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
