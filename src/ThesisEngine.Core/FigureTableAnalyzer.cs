using System.IO.Compression;
using System.Xml.Linq;

namespace ThesisEngine.Core;

/// <summary>
/// 图表排版分析器。
/// 检测文档中的图片、表格排版问题，以及跨页设置。
/// </summary>
public sealed class FigureTableAnalyzer
{
    private const string MainNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
    private static readonly XNamespace W = MainNamespace;
    private static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";

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

        // 统计图片数量
        var images = body.Descendants().Where(e => e.Name.LocalName == "blip").ToList();

        // 统计表格数量
        var tables = body.Elements(W + "tbl").ToList();

        // 检测跨页设置问题
        var pageBreakIssues = new List<string>();
        foreach (var para in body.Elements(W + "p"))
        {
            var pPr = para.Element(W + "pPr");
            if (pPr == null) continue;

            // 检测孤行控制
            if (pPr.Element(W + "widowControl") == null)
            {
                var text = para.Descendants(W + "t").Select(t => t.Value).Aggregate("", (a, b) => a + b);
                if (!string.IsNullOrWhiteSpace(text) && text.Length > 50)
                {
                    pageBreakIssues.Add("段落未开启孤行控制，可能出现跨页孤行");
                }
            }
        }

        // 检测表格边框问题（三线表应该只有上下边框和表头下边框）
        var tableIssues = new List<string>();
        for (int i = 0; i < tables.Count; i++)
        {
            var tblPr = tables[i].Element(W + "tblPr");
            var borders = tblPr?.Element(W + "tblBorders");
            if (borders != null)
            {
                // 有全边框的表格可能不是三线表
                if (borders.Element(W + "insideH") != null || borders.Element(W + "insideV") != null)
                {
                    tableIssues.Add($"表格 {i+1} 有完整内边框，建议改成三线表");
                }
            }
        }

        return new FigureTableAnalysisResult(
            ImagesCount: images.Count,
            TablesCount: tables.Count,
            PageBreakIssues: pageBreakIssues,
            TableIssues: tableIssues);
    }
}

public sealed record FigureTableAnalysisResult(
    int ImagesCount,
    int TablesCount,
    List<string> PageBreakIssues,
    List<string> TableIssues);
