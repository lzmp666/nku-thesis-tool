using System.IO.Compression;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

public sealed class FigureTableAnalyzerTests
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void ThreeLineTableWithNilInnerBordersIsAccepted()
    {
        var path = CreateDocx(Table("""
            <w:tblBorders>
              <w:top w:val="single" w:sz="12"/>
              <w:bottom w:val="single" w:sz="12"/>
              <w:insideH w:val="nil"/>
              <w:insideV w:val="nil"/>
              <w:left w:val="nil"/>
              <w:right w:val="nil"/>
            </w:tblBorders>
            """));
        try
        {
            var result = new FigureTableAnalyzer().Analyze(path);

            Assert.Equal(1, result.TablesCount);
            Assert.Empty(result.TableIssues);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void GridTableWithVisibleInnerBordersIsReported()
    {
        var path = CreateDocx(Table("""
            <w:tblBorders>
              <w:top w:val="single" w:sz="12"/>
              <w:bottom w:val="single" w:sz="12"/>
              <w:insideH w:val="single" w:sz="4"/>
              <w:insideV w:val="single" w:sz="4"/>
            </w:tblBorders>
            """));
        try
        {
            var result = new FigureTableAnalyzer().Analyze(path);

            var issue = Assert.Single(result.TableIssues);
            Assert.Contains("内部横线或竖线", issue);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void TableWithoutExplicitBordersIsReported()
    {
        var path = CreateDocx(Table(string.Empty));
        try
        {
            var result = new FigureTableAnalyzer().Analyze(path);

            var issue = Assert.Single(result.TableIssues);
            Assert.Contains("未显式设置三线表边框", issue);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void VisibleCellBordersAreReportedEvenWhenTableBordersLookThreeLine()
    {
        var path = CreateDocx("""
            <w:tbl>
              <w:tblPr>
                <w:tblBorders>
                  <w:top w:val="single" w:sz="12"/>
                  <w:left w:val="nil"/>
                  <w:bottom w:val="single" w:sz="12"/>
                  <w:right w:val="nil"/>
                  <w:insideH w:val="nil"/>
                  <w:insideV w:val="nil"/>
                </w:tblBorders>
              </w:tblPr>
              <w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/></w:tcBorders></w:tcPr>
                <w:p><w:r><w:t>项目</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            """);
        try
        {
            var issue = Assert.Single(new FigureTableAnalyzer().Analyze(path).TableIssues);
            Assert.Contains("单元格", issue);
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void PageBreakBeforeAndTableCaptionWithoutKeepNextAreReported()
    {
        var path = CreateDocx("""
            <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>设置了段前分页的段落</w:t></w:r></w:p>
            <w:p><w:r><w:t>表1 合成表格</w:t></w:r></w:p>
            """);
        try
        {
            var result = new FigureTableAnalyzer().Analyze(path);

            Assert.Contains(result.PageBreakIssues, issue => issue.Contains("段前分页"));
            Assert.Contains(result.PageBreakIssues, issue => issue.Contains("与下段同页"));
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 关键回归：三线表重写写入的是 <c>insideH/insideV val="nil"</c>，
    /// 分析器必须把它们视为“无内部边框”，否则修复后问题数反而增加。
    /// </summary>
    [Fact]
    public void ThreeLineRewriteMakesTableIssuesDisappear()
    {
        var root = FindRepositoryRoot();
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_object_layout_errors.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var analyzer = new FigureTableAnalyzer();
            var before = analyzer.Analyze(input);
            Assert.NotEmpty(before.TableIssues);

            new FigureTableRewriter().ApplyThreeLineTables(input, output);
            new FigureTableRewriter().CenterAlignImages(output, output);
            new FigureTableRewriter().FixPageBreaks(output, output);

            var after = analyzer.Analyze(output);
            Assert.Equal(before.TablesCount, after.TablesCount);
            Assert.Empty(after.TableIssues);
            Assert.Empty(after.PageBreakIssues);

            // 排版通过样式实现：表格套用 NKUThreeLine、图片段落套用 NKUFigure，
            // 不写直接边框/直接对齐；表头线 0.5 磅写在样式的首行条件格式里。
            using (var check = ZipFile.OpenRead(output))
            {
                XDocument xml;
                using (var reader = new StreamReader(check.GetEntry("word/document.xml")!.Open()))
                {
                    xml = XDocument.Load(reader);
                }

                XNamespace w = W;
                var tableProperties = xml.Descendants(w + "tbl").First().Element(w + "tblPr");
                Assert.Equal("NKUThreeLine", (string?)tableProperties?.Element(w + "tblStyle")?.Attribute(w + "val"));
                Assert.Null(tableProperties?.Element(w + "tblBorders"));
                Assert.DoesNotContain(xml.Descendants(w + "tcPr"), cell => cell.Element(w + "tcBorders") != null);

                var figureParagraphs = xml.Descendants(w + "p")
                    .Where(paragraph => paragraph.Descendants().Any(e => e.Name.LocalName == "blip"))
                    .ToList();
                Assert.NotEmpty(figureParagraphs);
                foreach (var paragraph in figureParagraphs)
                {
                    var properties = paragraph.Element(w + "pPr");
                    Assert.Equal("NKUFigure", (string?)properties?.Element(w + "pStyle")?.Attribute(w + "val"));
                    Assert.Null(properties?.Element(w + "jc"));
                }

                XDocument styles;
                using (var reader = new StreamReader(check.GetEntry("word/styles.xml")!.Open()))
                {
                    styles = XDocument.Load(reader);
                }

                var tableStyle = styles.Descendants(w + "style")
                    .First(style => (string?)style.Attribute(w + "styleId") == "NKUThreeLine");
                var headerBottom = tableStyle.Descendants(w + "tblStylePr").Descendants(w + "bottom").First();
                Assert.Equal("4", (string?)headerBottom.Attribute(w + "sz"));
                Assert.NotNull(styles.Descendants(w + "style")
                    .FirstOrDefault(style => (string?)style.Attribute(w + "styleId") == "NKUFigure"));
            }

            // 在真实文档包上校验：新增的段落属性必须符合 OOXML 结构，否则 Word 会报文件损坏。
            using var package = WordprocessingDocument.Open(output, false);
            var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019)
                .Validate(package)
                .Select(error => error.Description)
                .ToArray();
            Assert.Empty(errors);
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    /// <summary>
    /// 跨页修复只应保护题注；给所有段落加 keepNext 会把整篇文档串成一块，
    /// Word 反而插入大段空白。
    /// </summary>
    [Fact]
    public void PageBreakFixEnablesWidowControlButOnlyProtectsCaptions()
    {
        var path = CreateDocx("""
            <w:p><w:r><w:t>这是一段很长的正文段落，用来验证孤行控制会被显式开启，长度需要超过五十个字符才能触发分析器的检查逻辑，因此这里补充一些文字。</w:t></w:r></w:p>
            <w:p><w:r><w:t>表1 合成表格</w:t></w:r></w:p>
            <w:p><w:r><w:t>普通段落</w:t></w:r></w:p>
            """);
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            new FigureTableRewriter().FixPageBreaks(path, output);

            Assert.Empty(new FigureTableAnalyzer().Analyze(output).PageBreakIssues);

            using var archive = ZipFile.OpenRead(output);
            using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
            var xml = XDocument.Load(reader);
            XNamespace w = W;
            var paragraphs = xml.Descendants(w + "p").ToList();
            Assert.NotNull(paragraphs[1].Element(w + "pPr")?.Element(w + "keepNext"));
            Assert.Null(paragraphs[2].Element(w + "pPr")?.Element(w + "keepNext"));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void TableCellParagraphsAndContinuedCaptionsAreRecognized()
    {
        var path = CreateDocx("""
            <w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>续表4.1 合成记录</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
            """);
        try
        {
            using var archive = ZipFile.OpenRead(path);
            using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
            var document = XDocument.Load(reader);
            XNamespace w = W;
            var body = document.Root!.Element(w + "body")!;

            var paragraph = Assert.Single(FigureTableParagraphs.Of(body));
            var text = FigureTableParagraphs.Text(paragraph);
            Assert.Equal(CaptionKind.Table, FigureTableParagraphs.Classify(text));
            Assert.True(FigureTableParagraphs.IsContinuedTable(text));
        }
        finally { File.Delete(path); }
    }

    [Fact]
    public void CaptionStylesCarryR12SpacingAndAreAppliedToCaptions()
    {
        var root = FindRepositoryRoot();
        var input = Path.Combine(root, "tests", "fixtures", "synthetic", "full-thesis", "synthetic_object_layout_errors.docx");
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            new FigureTableRewriter().ApplyCaptionStyles(input, output);

            using var archive = ZipFile.OpenRead(output);
            XDocument document;
            using (var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open()))
            {
                document = XDocument.Load(reader);
            }

            XDocument styles;
            using (var reader = new StreamReader(archive.GetEntry("word/styles.xml")!.Open()))
            {
                styles = XDocument.Load(reader);
            }

            XNamespace w = W;
            Assert.Contains(document.Descendants(w + "pStyle"),
                style => (string?)style.Attribute(w + "val") == "NKUFigureCaption");
            Assert.Contains(document.Descendants(w + "pStyle"),
                style => (string?)style.Attribute(w + "val") == "NKUTableCaption");

            var figureStyle = styles.Descendants(w + "style")
                .First(style => (string?)style.Attribute(w + "styleId") == "NKUFigureCaption");
            var spacing = figureStyle.Descendants(w + "spacing").First();
            Assert.Equal("120", (string?)spacing.Attribute(w + "before"));
            Assert.Equal("240", (string?)spacing.Attribute(w + "after"));
            Assert.Equal("240", (string?)spacing.Attribute(w + "line"));
            Assert.Equal("auto", (string?)spacing.Attribute(w + "lineRule"));
            Assert.Equal("21", (string?)figureStyle.Descendants(w + "sz").First().Attribute(w + "val"));

            var tableStyle = styles.Descendants(w + "style")
                .First(style => (string?)style.Attribute(w + "styleId") == "NKUTableCaption");
            Assert.Equal("120", (string?)tableStyle.Descendants(w + "spacing").First().Attribute(w + "after"));
        }
        finally
        {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void FigureParagraphWithoutKeepNextIsReportedAndFixed()
    {
        var path = CreateDocx("""
            <w:p><w:r><w:blip/></w:r></w:p>
            <w:p><w:r><w:t>图1-1 合成图片</w:t></w:r></w:p>
            """);
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var before = new FigureTableAnalyzer().Analyze(path).PageBreakIssues;
            Assert.Contains(before, issue => issue.Contains("图片段落未设置"));

            new FigureTableRewriter().FixPageBreaks(path, output);

            Assert.DoesNotContain(new FigureTableAnalyzer().Analyze(output).PageBreakIssues,
                issue => issue.Contains("图片段落未设置"));

            using var archive = ZipFile.OpenRead(output);
            using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
            var xml = XDocument.Load(reader);
            XNamespace w = W;
            var figure = xml.Descendants(w + "p").First(p => p.Descendants(w + "blip").Any());
            Assert.NotNull(figure.Element(w + "pPr")?.Element(w + "keepNext"));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public void TableRowsWithoutCantSplitOrRepeatingHeaderAreReportedAndFixed()
    {
        var path = CreateDocx("""
            <w:tbl><w:tblPr/>
              <w:tr><w:tc><w:p><w:r><w:t>表头</w:t></w:r></w:p></w:tc></w:tr>
              <w:tr><w:tc><w:p><w:r><w:t>数据</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            """);
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            var before = new FigureTableAnalyzer().Analyze(path).PageBreakIssues;
            Assert.Contains(before, issue => issue.Contains("未禁止跨页断行"));
            Assert.Contains(before, issue => issue.Contains("跨页重复表头"));

            new FigureTableRewriter().FixPageBreaks(path, output);

            var after = new FigureTableAnalyzer().Analyze(output).PageBreakIssues;
            Assert.DoesNotContain(after, issue => issue.Contains("未禁止跨页断行"));
            Assert.DoesNotContain(after, issue => issue.Contains("跨页重复表头"));

            using var archive = ZipFile.OpenRead(output);
            using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
            var xml = XDocument.Load(reader);
            XNamespace w = W;
            var rows = xml.Descendants(w + "tr").ToList();
            Assert.All(rows, row => Assert.NotNull(row.Element(w + "trPr")?.Element(w + "cantSplit")));
            Assert.NotNull(rows[0].Element(w + "trPr")?.Element(w + "tblHeader"));
            Assert.Null(rows[1].Element(w + "trPr")?.Element(w + "tblHeader"));

            // 回归：cantSplit / tblHeader 不能带 w:val="1"，Office2019 校验会判非法值。
            Assert.All(rows, row =>
            {
                Assert.Null(row.Element(w + "trPr")?.Element(w + "cantSplit")?.Attribute(w + "val"));
                Assert.Null(row.Element(w + "trPr")?.Element(w + "tblHeader")?.Attribute(w + "val"));
            });
            // 完整 OOXML 合法性由 ThreeLineRewriteMakesTableIssuesDisappear 在
            // synthetic_object_layout_errors.docx 上用 OpenXmlValidator 覆盖；
            // 这里的 CreateDocx 极简包没有 MainDocumentPart 关系，不做整包校验。
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    /// <summary>
    /// 题注不能被拆到两页，缺 keepLines 要单独报出来。
    /// 图题不要求 keepNext（图在图下方），只有表题要求。
    /// </summary>
    [Fact]
    public void CaptionWithoutKeepLinesIsReported()
    {
        var path = CreateDocx("""
            <w:p><w:r><w:t>图1 合成图</w:t></w:r></w:p>
            """);
        try
        {
            var issues = new FigureTableAnalyzer().Analyze(path).PageBreakIssues;

            Assert.Contains(issues, issue => issue.Contains("段中不分页"));
            Assert.DoesNotContain(issues, issue => issue.Contains("与下段同页"));
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 续表的第一行常把题注放进表格单元格。只取 body 的直接子节点会整段漏掉，
    /// 输出 0 条与「全都合规」无法区分。
    /// </summary>
    [Fact]
    public void CaptionInsideTableCellIsDetected()
    {
        var path = CreateDocx("""
            <w:tbl>
              <w:tr><w:tc>
                <w:p><w:r><w:t>表2（续）合成表格</w:t></w:r></w:p>
              </w:tc></w:tr>
            </w:tbl>
            """);
        try
        {
            var issues = new FigureTableAnalyzer().Analyze(path).PageBreakIssues;

            Assert.Contains(issues, issue => issue.Contains("表题"));
            Assert.Contains(issues, issue => issue.Contains("段中不分页"));
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 同一段落可能同时命中多项。旧实现命中第一条就 continue，
    /// 会让「修完还剩几条」对不上。
    /// </summary>
    [Fact]
    public void ParagraphWithSeveralPageBreakProblemsReportsAll()
    {
        var path = CreateDocx("""
            <w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>图2 段前分页的题注</w:t></w:r></w:p>
            """);
        try
        {
            var issues = new FigureTableAnalyzer().Analyze(path).PageBreakIssues;

            Assert.Contains(issues, issue => issue.Contains("段前分页"));
            Assert.Contains(issues, issue => issue.Contains("段中不分页"));
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 单元格里的正文段落不判孤行控制：Word 不对表格行做孤行拆分，
    /// 硬判会让每个表格都报一串问题。
    /// </summary>
    [Fact]
    public void LongTextInsideTableCellIsNotReportedForWidowControl()
    {
        var path = CreateDocx("""
            <w:tbl>
              <w:tr><w:tc>
                <w:p><w:r><w:t>这是一段放在表格单元格里的很长正文，用来确认单元格内的段落不会被要求开启孤行控制，因为表格行本身不按孤行拆分，硬判只会制造误报。</w:t></w:r></w:p>
              </w:tc></w:tr>
            </w:tbl>
            """);
        try
        {
            var issues = new FigureTableAnalyzer().Analyze(path).PageBreakIssues;

            Assert.DoesNotContain(issues, issue => issue.Contains("孤行控制"));
        }
        finally { File.Delete(path); }
    }

    /// <summary>
    /// 修复要和分析对称：表内题注同样加 keepLines/keepNext，
    /// 否则会出现「检测报、修复不改」的假修复。
    /// </summary>
    [Fact]
    public void PageBreakFixAlsoProtectsCaptionInsideTableCell()
    {
        var path = CreateDocx("""
            <w:tbl>
              <w:tr><w:tc>
                <w:p><w:r><w:t>表3（续）表内题注</w:t></w:r></w:p>
              </w:tc></w:tr>
            </w:tbl>
            """);
        var output = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        try
        {
            new FigureTableRewriter().FixPageBreaks(path, output);

            Assert.Empty(new FigureTableAnalyzer().Analyze(output).PageBreakIssues);

            using var archive = ZipFile.OpenRead(output);
            using var reader = new StreamReader(archive.GetEntry("word/document.xml")!.Open());
            var xml = XDocument.Load(reader);
            XNamespace w = W;
            var caption = xml.Descendants(w + "p")
                .First(paragraph => string.Concat(paragraph.Descendants(w + "t").Select(t => t.Value)).StartsWith("表3"));
            var properties = caption.Element(w + "pPr");
            Assert.NotNull(properties?.Element(w + "keepLines"));
            Assert.NotNull(properties?.Element(w + "keepNext"));
        }
        finally
        {
            File.Delete(path);
            if (File.Exists(output)) File.Delete(output);
        }
    }

    /// <summary>#92 第 10 条：形似题注但编号解析不出来，要报一条，不能静默放过。</summary>
    [Fact]
    public void UnrecognizedCaptionNumberIsReported()
    {
        var path = CreateDocx("""
            <w:p><w:r><w:t>图一 系统架构</w:t></w:r></w:p>
            """);
        try
        {
            var result = new FigureTableAnalyzer().Analyze(path);
            Assert.Contains(result.TableIssues, issue => issue.Contains("编号无法识别"));
        }
        finally { File.Delete(path); }
    }

    /// <summary>反例：正常题注和普通正文都不该报无法识别，防止规则太宽误伤。</summary>
    [Fact]
    public void NormalCaptionsAndBodyTextAreNotFlaggedAsUnrecognized()
    {
        var path = CreateDocx("""
            <w:p><w:r><w:t>图 1-1 系统架构</w:t></w:r></w:p>
            <w:p><w:r><w:t>如表所示，结果已在前面说明。</w:t></w:r></w:p>
            <w:p><w:r><w:t>续表 2-3 参数对照（编号形态正确）</w:t></w:r></w:p>
            """);
        try
        {
            var result = new FigureTableAnalyzer().Analyze(path);
            Assert.DoesNotContain(result.TableIssues, issue => issue.Contains("编号无法识别"));
        }
        finally { File.Delete(path); }
    }

    private static string Table(string borderXml) =>
        $"<w:tbl><w:tblPr>{borderXml}</w:tblPr>"
        + "<w:tr><w:tc><w:p><w:r><w:t>项目</w:t></w:r></w:p></w:tc>"
        + "<w:tc><w:p><w:r><w:t>数值</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";

    private static string CreateDocx(string bodyXml)
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".docx");
        using var archive = ZipFile.Open(path, ZipArchiveMode.Create);
        Add(archive, "[Content_Types].xml",
            """<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>""");
        Add(archive, "_rels/.rels",
            """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>""");
        Add(archive, "word/document.xml",
            $"<w:document xmlns:w=\"{W}\"><w:body>{bodyXml}</w:body></w:document>");
        return path;
    }

    private static void Add(ZipArchive archive, string name, string content)
    {
        using var writer = new StreamWriter(archive.CreateEntry(name).Open());
        writer.Write(content);
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "NkuThesisTool.sln")))
            directory = directory.Parent;
        Assert.NotNull(directory);
        return directory!.FullName;
    }
}
