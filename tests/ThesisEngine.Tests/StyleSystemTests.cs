using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using ThesisEngine.Core;
using Xunit;

namespace ThesisEngine.Tests;

/// <summary>
/// Issue #90：Word 样式体系读取与批量重写 —— 样式归属检测与结构识别测试。
/// 重点覆盖此前缺口：表格内段落漏检、无效样式引用、结构识别（多级标题/摘要/参考文献）。
/// </summary>
public sealed class StyleSystemTests
{
    // ---------- StyleAnalyzer ----------

    [Fact]
    public void AnalyzerDetectsParagraphsInsideTables()
    {
        var path = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"),
                P("普通正文段落", style: "Normal"),
                Table(
                    Row(Cell(P("表格里的无样式段落"))),
                    Row(Cell(P("表格里有样式的段落", style: "Normal")))));

            CreateDocx(path, body, ("Normal", "Normal"));

            var result = new StyleAnalyzer().Analyze(path);

            var unstyled = Assert.Single(result.UnstyledParagraphs);
            Assert.Contains("表格里", unstyled.Preview);
            Assert.True(unstyled.InTable);

            var styledInTable = Assert.Single(result.StyledParagraphs, p => p.InTable);
            Assert.Equal("Normal", styledInTable.StyleId);
        }
        finally { TryDelete(path); }
    }

    [Fact]
    public void AnalyzerReportsInvalidStyleReferences()
    {
        var path = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"),
                P("引用了不存在样式的段落", style: "GhostStyle"),
                P("正常段落", style: "Normal"));

            CreateDocx(path, body, ("Normal", "Normal"));

            var result = new StyleAnalyzer().Analyze(path);

            var invalid = Assert.Single(result.InvalidStyleParagraphs);
            Assert.Equal("GhostStyle", invalid.StyleId);
            // 有效样式段落不受影响
            Assert.Equal(2, result.StyledParagraphs.Count);
            Assert.Empty(result.UnstyledParagraphs);
        }
        finally { TryDelete(path); }
    }

    [Fact]
    public void AnalyzerListsAllStyles()
    {
        var path = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"), P("正文", style: "Normal"));
            CreateDocx(path, body,
                ("Normal", "Normal"), ("Heading1", "heading 1"), ("Caption", "caption"));

            var result = new StyleAnalyzer().Analyze(path);

            Assert.Equal(3, result.AllStyles.Count);
            Assert.Equal("heading 1", result.AllStyles["Heading1"]);
        }
        finally { TryDelete(path); }
    }

    // ---------- StyleAutoAssigner ----------

    [Fact]
    public void AssignerRecognizesHeadingLevels()
    {
        var path = NewTempDocxPath();
        var output = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"),
                P("第一章 绪论"),
                P("本章介绍研究背景。"),
                P("1.1 研究背景"),
                P("1.1.1 国内研究现状"),
                P("以下是正文内容。"));

            CreateDocx(path, body, ("Normal", "Normal"), ("Heading1", "heading 1"),
                ("Heading2", "heading 2"), ("Heading3", "heading 3"));

            var result = new StyleAutoAssigner(new StringWriter()).AutoAssignStyles(path, output);
            var assigned = ReadAssignedStyles(output);

            Assert.Equal("Heading1", assigned["第一章 绪论"]);
            Assert.Equal("Heading2", assigned["1.1 研究背景"]);
            Assert.Equal("Heading3", assigned["1.1.1 国内研究现状"]);
            Assert.Equal("Normal", assigned["本章介绍研究背景。"]);
            Assert.Equal(5, result.TotalAssigned);
        }
        finally { TryDelete(path); TryDelete(output); }
    }

    [Fact]
    public void AssignerRecognizesStructuralHeadingsAbstractAndBibliography()
    {
        var path = NewTempDocxPath();
        var output = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"),
                P("摘  要"),
                P("摘要：本文研究了样式体系。"),
                P("参考文献"),
                P("[1] 张三. 论文写作[M]. 北京: 出版社, 2026."),
                P("[2] 李四. 又一篇论文[J]. 学报, 2025."));

            CreateDocx(path, body, ("Normal", "Normal"), ("Heading1", "heading 1"),
                ("Abstract", "Abstract"), ("Bibliography", "Bibliography"));

            new StyleAutoAssigner(new StringWriter()).AutoAssignStyles(path, output);
            var assigned = ReadAssignedStyles(output);

            Assert.Equal("Heading1", assigned["摘  要"]);
            Assert.Equal("Abstract", assigned["摘要：本文研究了样式体系。"]);
            Assert.Equal("Heading1", assigned["参考文献"]);
            Assert.Equal("Bibliography", assigned["[1] 张三. 论文写作[M]. 北京: 出版社, 2026."]);
            Assert.Equal("Bibliography", assigned["[2] 李四. 又一篇论文[J]. 学报, 2025."]);
        }
        finally { TryDelete(path); TryDelete(output); }
    }

    [Fact]
    public void AssignerFallsBackToNormalWhenCandidateStyleMissing()
    {
        var path = NewTempDocxPath();
        var output = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"),
                P("第一章 绪论"),
                P("[1] 某条参考文献。"));

            // 样式表里没有 Heading1 / Bibliography
            CreateDocx(path, body, ("Normal", "Normal"));

            new StyleAutoAssigner(new StringWriter()).AutoAssignStyles(path, output);
            var assigned = ReadAssignedStyles(output);

            Assert.Equal("Normal", assigned["第一章 绪论"]);
            Assert.Equal("Normal", assigned["[1] 某条参考文献。"]);
        }
        finally { TryDelete(path); TryDelete(output); }
    }

    [Fact]
    public void AssignerCoversParagraphsInsideTablesAndKeepsExistingStyles()
    {
        var path = NewTempDocxPath();
        var output = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"),
                Table(
                    Row(Cell(P("表格里的无样式段落"))),
                    Row(Cell(P("表格里已有样式的段落", style: "Normal")))));

            CreateDocx(path, body, ("Normal", "Normal"));

            var result = new StyleAutoAssigner(new StringWriter()).AutoAssignStyles(path, output);
            var assigned = ReadAssignedStyles(output);

            // 表格内无样式段落被套上 Normal；已有样式段落保持不动
            Assert.Equal("Normal", assigned["表格里的无样式段落"]);
            Assert.Equal("Normal", assigned["表格里已有样式的段落"]);
            Assert.Equal(1, result.TotalAssigned);
        }
        finally { TryDelete(path); TryDelete(output); }
    }

    // ---------- CLI 报告 ----------

    [Fact]
    public void CliAnalyzeStylesReportsTableAndInvalidReferences()
    {
        var path = NewTempDocxPath();
        try
        {
            var body = new XElement(W("body"),
                Table(Row(Cell(P("表格里的无样式段落")))),
                P("幽灵样式段落", style: "GhostStyle"));
            CreateDocx(path, body, ("Normal", "Normal"));

            var output = new StringWriter();
            var error = new StringWriter();
            var command = NewStyleAnalyzeCommand(CliAssembly, output, error);

            var exitCode = Execute(command, new[] { path });

            Assert.Equal(0, exitCode);
            Assert.Contains("GhostStyle", output.ToString());
            Assert.Contains("不存在", output.ToString());
            Assert.Contains("表格内", output.ToString());
            Assert.Contains("Normal", output.ToString());
        }
        finally { TryDelete(path); }
    }

    // ---------- 测试辅助 ----------

    private static readonly System.Reflection.Assembly CliAssembly = LoadCliAssembly();

    private static object NewStyleAnalyzeCommand(System.Reflection.Assembly cli, TextWriter output, TextWriter error)
    {
        var type = cli.GetType("ThesisEngine.Cli.StyleAnalyzeCommand")!;
        return System.Activator.CreateInstance(type, null, output, error)!;
    }

    private static int Execute(object command, string[] args)
    {
        var execute = command.GetType().GetMethod("Execute", new[] { typeof(string[]) })!;
        return (int)execute.Invoke(command, new object[] { args })!;
    }

    private static System.Reflection.Assembly LoadCliAssembly()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "NkuThesisTool.sln")))
            root = root.Parent;
        Assert.NotNull(root);

        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
        var cliBin = Path.Combine(root!.FullName, "src", "ThesisEngine.Cli", "bin", configuration);
        var dll = Directory
            .EnumerateFiles(cliBin, "ThesisEngine.Cli.dll", SearchOption.AllDirectories)
            .FirstOrDefault();
        Assert.True(dll is not null, $"CLI was not built under: {cliBin}");
        return System.Reflection.Assembly.LoadFrom(dll);
    }

    private static string NewTempDocxPath() =>
        Path.Combine(Path.GetTempPath(), "style-tests-" + Guid.NewGuid().ToString("N") + ".docx");

    private static void TryDelete(string path)
    {
        try { File.Delete(path); }
        catch (IOException) { }
    }

    private static Dictionary<string, string> ReadAssignedStyles(string docxPath)
    {
        using var stream = new FileStream(docxPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        var entry = archive.GetEntry("word/document.xml")!;
        using var reader = new StreamReader(entry.Open());
        var doc = XDocument.Load(reader);

        var result = new Dictionary<string, string>();
        foreach (var para in doc.Root!.Element(W("body"))!.Descendants(W("p")))
        {
            var text = para.Descendants(W("t")).Select(t => t.Value).Aggregate("", (a, b) => a + b).Trim();
            if (text.Length == 0) continue;
            var style = para.Element(W("pPr"))?.Element(W("pStyle"))?.Attribute(W("val"))?.Value;
            result[text] = style ?? string.Empty;
        }
        return result;
    }

    private static XElement P(string text, string? style = null, bool inTable = false)
    {
        var p = new XElement(W("p"));
        if (style != null)
        {
            p.Add(new XElement(W("pPr"),
                new XElement(W("pStyle"), new XAttribute(W("val"), style))));
        }
        p.Add(new XElement(W("r"), new XElement(W("t"), text)));
        return p;
    }

    private static XElement Table(params XElement[] rows) =>
        new(W("tbl"), rows);

    private static XElement Row(params XElement[] cells) =>
        new(W("tr"), cells);

    private static XElement Cell(params XElement[] paragraphs) =>
        new(W("tc"), paragraphs);

    private static XName W(string localName) =>
        (XNamespace)"http://schemas.openxmlformats.org/wordprocessingml/2006/main" + localName;

    private static void CreateDocx(string path, XElement body, params (string Id, string Name)[] styles)
    {
        var stylesXml = new XDocument(new XElement(W("styles"),
            styles.Select(s => new XElement(W("style"),
                new XAttribute(W("type"), "paragraph"),
                new XAttribute(W("styleId"), s.Id),
                new XElement(W("name"), new XAttribute(W("val"), s.Name))))));

        var document = new XDocument(new XElement(W("document"), body));

        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Create);
        WriteEntry(zip, "[Content_Types].xml",
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
            "</Types>");
        WriteEntry(zip, "_rels/.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "</Relationships>");
        WriteEntry(zip, "word/_rels/document.xml.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>" +
            "</Relationships>");
        WriteEntry(zip, "word/document.xml", document.ToString(SaveOptions.DisableFormatting));
        WriteEntry(zip, "word/styles.xml", stylesXml.ToString(SaveOptions.DisableFormatting));
    }

    private static void WriteEntry(ZipArchive zip, string name, string content)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content);
    }
}
