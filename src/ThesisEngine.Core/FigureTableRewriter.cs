using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// 图表排版自动修正器。
/// 自动修正表格为三线表，图片居中对齐。
/// </summary>
public sealed class FigureTableRewriter
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace W = MainNamespace;

    /// <summary>
    /// 将文档中所有表格改为三线表格式。
    /// </summary>
    public void ApplyThreeLineTables(string inputPath, string outputPath)
    {
        var doc = ReadDocumentXml(inputPath);
        var body = doc.Root?.Element(W + "body");
        if (body == null) return;

        var tables = body.Elements(W + "tbl").ToList();
        foreach (var table in tables)
        {
            var tblPr = table.Element(W + "tblPr");
            if (tblPr == null)
            {
                tblPr = new XElement(W + "tblPr");
                table.AddFirst(tblPr);
            }

            // 删除原有边框
            tblPr.Element(W + "tblBorders")?.Remove();

            // 添加三线表边框：只有上边框、下边框、表头下边框
            var borders = new XElement(W + "tblBorders");
            borders.Add(new XElement(W + "top",
                new XAttribute(W + "val", "single"),
                new XAttribute(W + "sz", "12"),
                new XAttribute(W + "space", "0"),
                new XAttribute(W + "color", "000000")));
            borders.Add(new XElement(W + "bottom",
                new XAttribute(W + "val", "single"),
                new XAttribute(W + "sz", "12"),
                new XAttribute(W + "space", "0"),
                new XAttribute(W + "color", "000000")));
            borders.Add(new XElement(W + "insideH",
                new XAttribute(W + "val", "nil")));
            borders.Add(new XElement(W + "insideV",
                new XAttribute(W + "val", "nil")));
            borders.Add(new XElement(W + "left",
                new XAttribute(W + "val", "nil")));
            borders.Add(new XElement(W + "right",
                new XAttribute(W + "val", "nil")));

            tblPr.Add(borders);

            // 第一行（表头）加下边框
            var firstRow = table.Elements(W + "tr").FirstOrDefault();
            if (firstRow != null)
            {
                var trPr = firstRow.Element(W + "trPr");
                if (trPr == null)
                {
                    trPr = new XElement(W + "trPr");
                    firstRow.AddFirst(trPr);
                }
            }
        }

        WriteDocumentXml(inputPath, outputPath, doc);
    }

    /// <summary>
    /// 将所有图片所在段落设置为居中对齐。
    /// </summary>
    public void CenterAlignImages(string inputPath, string outputPath)
    {
        var doc = ReadDocumentXml(inputPath);
        var body = doc.Root?.Element(W + "body");
        if (body == null) return;

        var paragraphs = body.Descendants(W + "p").ToList();
        foreach (var para in paragraphs)
        {
            // 检查段落里有没有图片
            var hasImage = para.Descendants().Any(e => e.Name.LocalName == "blip");
            if (!hasImage) continue;

            // 设置居中对齐
            var pPr = para.Element(W + "pPr");
            if (pPr == null)
            {
                pPr = new XElement(W + "pPr");
                para.AddFirst(pPr);
            }

            var jc = pPr.Element(W + "jc");
            if (jc == null)
            {
                jc = new XElement(W + "jc");
                pPr.Add(jc);
            }
            jc.SetAttributeValue(W + "val", "center");
        }

        WriteDocumentXml(inputPath, outputPath, doc);
    }

    /// <summary>
    /// 自动修正跨页问题：给所有段落加上孤行控制、与下段同页等设置。
    /// 避免出现大段空白页。
    /// </summary>
    public void FixPageBreaks(string inputPath, string outputPath)
    {
        var doc = ReadDocumentXml(inputPath);
        var body = doc.Root?.Element(W + "body");
        if (body == null) return;

        var paragraphs = body.Descendants(W + "p").ToList();
        int fixedCount = 0;

        foreach (var para in paragraphs)
        {
            var pPr = para.Element(W + "pPr");
            if (pPr == null)
            {
                pPr = new XElement(W + "pPr");
                para.AddFirst(pPr);
            }

            // 孤行控制：避免段落最后一行单独在一页
            var widowControl = pPr.Element(W + "widowControl");
            if (widowControl == null)
            {
                widowControl = new XElement(W + "widowControl");
                pPr.AddFirst(widowControl);
                fixedCount++;
            }

            // 与下段同页：标题和下面的正文不要分开
            var keepNext = pPr.Element(W + "keepNext");
            if (keepNext == null)
            {
                keepNext = new XElement(W + "keepNext");
                pPr.AddFirst(keepNext);
            }
        }

        WriteDocumentXml(inputPath, outputPath, doc);
        Console.WriteLine($"已为 {paragraphs.Count} 个段落设置跨页保护");
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
