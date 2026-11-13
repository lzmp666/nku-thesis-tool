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
        var stylesDocument = TryReadStylesXml(inputPath);
        if (stylesDocument != null) FigureTableStyleWriter.EnsureStyles(stylesDocument);
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

            // 正常情况下走“三线表样式”：只套样式、不写直接格式。
            if (stylesDocument != null)
            {
                tblPr.Element(W + "tblBorders")?.Remove();

                var tableStyle = tblPr.Element(W + "tblStyle");
                if (tableStyle == null)
                {
                    tableStyle = new XElement(W + "tblStyle");
                    InsertOrdered(tblPr, tableStyle, TablePropertySequence);
                }

                tableStyle.SetAttributeValue(W + "val", FigureTableStyleWriter.ThreeLineTableStyleId);

                // 首行条件格式承载表头分隔线，需要打开 tblLook.firstRow。
                var look = tblPr.Element(W + "tblLook");
                if (look == null)
                {
                    look = new XElement(W + "tblLook");
                    InsertOrdered(tblPr, look, TablePropertySequence);
                }

                look.SetAttributeValue(W + "firstRow", "1");

                // 单元格级边框优先于表级，必须清除才能让样式生效。
                foreach (var cell in table.Descendants(W + "tc"))
                {
                    cell.Element(W + "tcPr")?.Element(W + "tcBorders")?.Remove();
                }

                continue;
            }

            // 删除原有边框
            tblPr.Element(W + "tblBorders")?.Remove();

            // 三线表：顶线、表头下线、底线；其余边线显式为 nil。
            // 子元素顺序必须遵循 OOXML（top,left,bottom,right,insideH,insideV），
            // 否则 Word 会判定文档损坏。
            var borders = new XElement(W + "tblBorders",
                Border("top", "single", "12"),
                Border("left", "nil", null),
                Border("bottom", "single", "12"),
                Border("right", "nil", null),
                Border("insideH", "nil", null),
                Border("insideV", "nil", null));
            InsertOrdered(tblPr, borders, TablePropertySequence);

            // 单元格级边框优先级高于表级：必须全部清除，否则 Word 打开仍是全框线。
            foreach (var cell in table.Descendants(W + "tc"))
            {
                cell.Element(W + "tcPr")?.Element(W + "tcBorders")?.Remove();
            }

            // 表头下边框：0.5 磅细线（sz 以 1/8 磅为单位，0.5pt = 4）。
            var firstRow = table.Elements(W + "tr").FirstOrDefault();
            if (firstRow != null)
            {
                foreach (var cell in firstRow.Elements(W + "tc"))
                {
                    var cellProperties = cell.Element(W + "tcPr");
                    if (cellProperties == null)
                    {
                        cellProperties = new XElement(W + "tcPr");
                        cell.AddFirst(cellProperties);
                    }

                    var cellBorders = cellProperties.Element(W + "tcBorders");
                    if (cellBorders == null)
                    {
                        cellBorders = new XElement(W + "tcBorders");
                        InsertOrdered(cellProperties, cellBorders, TableCellPropertySequence);
                    }

                    cellBorders.Element(W + "bottom")?.Remove();
                    InsertOrdered(cellBorders, Border("bottom", "single", "4"), TableBorderSequence);
                }
            }
        }

        WriteDocumentXml(inputPath, outputPath, doc);
        if (stylesDocument != null) WriteStylesXml(inputPath, outputPath, stylesDocument);
    }

    /// <summary>
    /// 题注排版：图题套 <c>NKUFigureCaption</c>、表题套 <c>NKUTableCaption</c>。
    /// 两个样式都定义段前 6 磅、单倍行距、宋体 10.5 磅居中，段后分别为 12/6 磅；
    /// “续表”按表题处理。
    /// </summary>
    public void ApplyCaptionStyles(string inputPath, string outputPath)
    {
        var doc = ReadDocumentXml(inputPath);
        var stylesDocument = TryReadStylesXml(inputPath);
        if (stylesDocument != null) FigureTableStyleWriter.EnsureStyles(stylesDocument);

        var body = doc.Root?.Element(W + "body");
        if (body == null) return;

        foreach (var paragraph in FigureTableParagraphs.Of(body))
        {
            var kind = FigureTableParagraphs.Classify(FigureTableParagraphs.Text(paragraph));
            if (kind == null) continue;

            var pPr = paragraph.Element(W + "pPr");
            if (pPr == null)
            {
                pPr = new XElement(W + "pPr");
                paragraph.AddFirst(pPr);
            }

            if (stylesDocument != null)
            {
                var paragraphStyle = pPr.Element(W + "pStyle");
                if (paragraphStyle == null)
                {
                    paragraphStyle = new XElement(W + "pStyle");
                    InsertOrdered(pPr, paragraphStyle, ParagraphPropertySequence);
                }

                paragraphStyle.SetAttributeValue(W + "val", kind == CaptionKind.Table
                    ? FigureTableStyleWriter.TableCaptionStyleId
                    : FigureTableStyleWriter.FigureCaptionStyleId);
                continue;
            }

            // styles.xml 缺失时的降级路径：直接写段落属性。
            var spacing = pPr.Element(W + "spacing");
            if (spacing == null)
            {
                spacing = new XElement(W + "spacing");
                InsertOrdered(pPr, spacing, ParagraphPropertySequence);
            }

            spacing.SetAttributeValue(W + "before", "120");
            spacing.SetAttributeValue(W + "after", kind == CaptionKind.Table ? "120" : "240");
            spacing.SetAttributeValue(W + "line", "240");
            spacing.SetAttributeValue(W + "lineRule", "auto");

            var jc = pPr.Element(W + "jc");
            if (jc == null)
            {
                jc = new XElement(W + "jc");
                InsertOrdered(pPr, jc, ParagraphPropertySequence);
            }

            jc.SetAttributeValue(W + "val", "center");
        }

        WriteDocumentXml(inputPath, outputPath, doc);
        if (stylesDocument != null) WriteStylesXml(inputPath, outputPath, stylesDocument);
    }

    /// <summary>
    /// 将所有图片所在段落设置为居中对齐。
    /// </summary>
    public void CenterAlignImages(string inputPath, string outputPath)
    {
        var doc = ReadDocumentXml(inputPath);
        var stylesDocument = TryReadStylesXml(inputPath);
        if (stylesDocument != null) FigureTableStyleWriter.EnsureStyles(stylesDocument);

        var body = doc.Root?.Element(W + "body");
        if (body == null) return;

        var paragraphs = FigureTableParagraphs.Of(body).ToList();
        foreach (var para in paragraphs)
        {
            // 检查段落里有没有图片
            var hasImage = para.Descendants().Any(e => e.Name.LocalName == "blip");
            if (!hasImage) continue;

            var pPr = para.Element(W + "pPr");
            if (pPr == null)
            {
                pPr = new XElement(W + "pPr");
                para.AddFirst(pPr);
            }

            if (stylesDocument != null)
            {
                // 套用图片段落样式（居中），不再写直接对齐格式。
                pPr.Element(W + "jc")?.Remove();
                var paragraphStyle = pPr.Element(W + "pStyle");
                if (paragraphStyle == null)
                {
                    paragraphStyle = new XElement(W + "pStyle");
                    InsertOrdered(pPr, paragraphStyle, ParagraphPropertySequence);
                }

                paragraphStyle.SetAttributeValue(W + "val", FigureTableStyleWriter.FigureParagraphStyleId);
                continue;
            }

            var jc = pPr.Element(W + "jc");
            if (jc == null)
            {
                jc = new XElement(W + "jc");
                InsertOrdered(pPr, jc, ParagraphPropertySequence);
            }
            jc.SetAttributeValue(W + "val", "center");
        }

        WriteDocumentXml(inputPath, outputPath, doc);
        if (stylesDocument != null) WriteStylesXml(inputPath, outputPath, stylesDocument);
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

        var paragraphs = FigureTableParagraphs.Of(body).ToList();
        int widowCount = 0;
        int captionCount = 0;

        foreach (var para in paragraphs)
        {
            var pPr = para.Element(W + "pPr");
            if (pPr == null)
            {
                pPr = new XElement(W + "pPr");
                para.AddFirst(pPr);
            }

            // 孤行控制：Word 默认开启，这里只是把继承来的设置显式化，
            // 避免不同样式来源导致正文出现跨页孤行。
            if (pPr.Element(W + "widowControl") == null)
            {
                EnsureOrderedProperty(pPr, "widowControl");
                widowCount++;
            }

            // 只对题注做“保持同页”处理。绝不能给所有段落加 keepNext：
            // 那会让整篇文档被串成一块，Word 反而插入大面积空白。
            var text = string.Concat(para.Descendants(W + "t").Select(t => t.Value)).Trim();
            if (text.Length == 0 || text.Length > 40) continue;

            var isTableCaption = text.StartsWith("表", StringComparison.Ordinal)
                || text.StartsWith("Table", StringComparison.OrdinalIgnoreCase);
            var isFigureCaption = text.StartsWith("图", StringComparison.Ordinal)
                || text.StartsWith("Figure", StringComparison.OrdinalIgnoreCase)
                || text.StartsWith("Fig.", StringComparison.OrdinalIgnoreCase);
            if (!isTableCaption && !isFigureCaption) continue;

            // 题注本身不要被拆到两页
            EnsureOrderedProperty(pPr, "keepLines");
            if (isTableCaption)
            {
                // 表题在表格上方，必须与表格同页
                EnsureOrderedProperty(pPr, "keepNext");
            }

            captionCount++;
        }

        WriteDocumentXml(inputPath, outputPath, doc);
        Console.WriteLine($"已为 {paragraphs.Count} 个段落设置孤行控制（新增 {widowCount}）,并处理 {captionCount} 条图表题注的跨页保护");
    }

    /// <summary>
    /// 按 OOXML 规定的 <c>w:pPr</c> 子元素顺序插入开关属性。
    /// 顺序错误会让 Word 判定文档损坏。
    /// </summary>
    private static XElement EnsureOrderedProperty(XElement pPr, string name)
    {
        var existing = pPr.Element(W + name);
        if (existing != null) return existing;

        var child = new XElement(W + name, new XAttribute(W + "val", "1"));
        InsertOrdered(pPr, child, ParagraphPropertySequence);
        return child;
    }

    private static XElement Border(string side, string value, string? size)
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

    /// <summary>按 OOXML 序列把元素插入到正确位置，未知元素放到末尾。</summary>
    private static void InsertOrdered(XElement parent, XElement child, IReadOnlyList<string> sequence)
    {
        var targetOrder = OrderOf(sequence, child.Name.LocalName);
        var later = parent.Elements().FirstOrDefault(element => OrderOf(sequence, element.Name.LocalName) > targetOrder);
        if (later != null)
        {
            later.AddBeforeSelf(child);
        }
        else
        {
            parent.Add(child);
        }
    }

    private static int OrderOf(IReadOnlyList<string> sequence, string name)
    {
        for (var index = 0; index < sequence.Count; index++)
        {
            if (string.Equals(sequence[index], name, StringComparison.Ordinal)) return index;
        }

        return int.MaxValue;
    }

    private static readonly string[] TablePropertySequence =
    {
        "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
        "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar",
        "tblLook", "tblCaption", "tblDescription"
    };

    private static readonly string[] TableBorderSequence =
    {
        "top", "left", "bottom", "right", "insideH", "insideV", "tl2br", "tr2bl"
    };

    private static readonly string[] TableCellPropertySequence =
    {
        "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
        "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel",
        "cellMerge", "tcPrChange"
    };

    private static readonly string[] ParagraphPropertySequence =
    {
        "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
        "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
        "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
        "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap",
        "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId",
        "cnfStyle", "rPr", "sectPr", "pPrChange"
    };

    private static XDocument ReadDocumentXml(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);

        var entry = archive.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("文档中没有 document.xml");

        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }

    private static XDocument? TryReadStylesXml(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = archive.GetEntry("word/styles.xml");
        if (entry == null) return null;
        using var reader = new StreamReader(entry.Open());
        return XDocument.Load(reader);
    }

    /// <summary>写入 styles.xml；调用前 output 必须已存在（由 WriteDocumentXml 创建）。</summary>
    private static void WriteStylesXml(string inputPath, string outputPath, XDocument doc)
    {
        using var stream = new FileStream(outputPath, FileMode.Open, FileAccess.ReadWrite);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Update);
        archive.GetEntry("word/styles.xml")?.Delete();
        var entry = archive.CreateEntry("word/styles.xml");
        using var writer = new StreamWriter(entry.Open());
        doc.Save(writer);
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
