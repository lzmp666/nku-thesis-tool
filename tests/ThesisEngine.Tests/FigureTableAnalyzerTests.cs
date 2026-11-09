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
            new FigureTableRewriter().FixPageBreaks(output, output);

            var after = analyzer.Analyze(output);
            Assert.Equal(before.TablesCount, after.TablesCount);
            Assert.Empty(after.TableIssues);
            Assert.Empty(after.PageBreakIssues);

            // 表头线 0.5 磅（sz=4），且表头之外不得残留单元格边框。
            using (var check = ZipFile.OpenRead(output))
            {
                using var reader = new StreamReader(check.GetEntry("word/document.xml")!.Open());
                var xml = XDocument.Load(reader);
                XNamespace w = W;
                var rows = xml.Descendants(w + "tbl").First().Elements(w + "tr").ToList();
                var headerBorder = rows[0].Elements(w + "tc").First()
                    .Element(w + "tcPr")?.Element(w + "tcBorders")?.Element(w + "bottom");
                Assert.NotNull(headerBorder);
                Assert.Equal("4", (string?)headerBorder!.Attribute(w + "sz"));
                foreach (var row in rows.Skip(1))
                {
                    foreach (var cell in row.Elements(w + "tc"))
                    {
                        Assert.Null(cell.Element(w + "tcPr")?.Element(w + "tcBorders"));
                    }
                }
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
